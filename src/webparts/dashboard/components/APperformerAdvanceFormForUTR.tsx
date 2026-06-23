import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

import logo from "../assets/sona-comstarlogo.png";

import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { useEffect, useState } from "react";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import Swal from "sweetalert2";

interface IProps {
  context: any;
  itemId: number;
  formData: any;
  onClose: () => void;
}
interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

const APperformerAdvanceFormForUTR: React.FC<IProps> = ({
  context,
  itemId,
  formData,
  onClose,
}) => {
  const sp = spfi().using(SPFx(context));
  const [attachments, setAttachments] = useState<any[]>([]);
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [advanceHistory, setAdvanceHistory] = useState<any[]>([]);
  const [itemData, setItemData] = useState<any>(null);
  const [approverRemarks, setApproverRemarks] = useState("");
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");
  const [UTRRemarks, setUTRRemarks] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };

  const poAmountBasic = itemData?.POAmountBasic != null ? String(itemData.POAmountBasic) : "0";
  const poAmountGST = itemData?.POAmountGST != null ? String(itemData.POAmountGST) : "0";
  const poAmountOther = itemData?.POAmountOther != null ? String(itemData.POAmountOther) : "0";
  const mrnAmountBasic = itemData?.MRNAmountBasic != null ? String(itemData.MRNAmountBasic) : "0";
  const mrnAmountGST = itemData?.MRNAmountGST != null ? String(itemData.MRNAmountGST) : "0";
  const mrnAmountOther = itemData?.MRNAmountOther != null ? String(itemData.MRNAmountOther) : "0";

  const poAmountTotal = (
    Number(poAmountBasic) + Number(poAmountGST) + Number(poAmountOther)
  ).toFixed(2);

  const mrnAmountTotal = (
    Number(mrnAmountBasic) + Number(mrnAmountGST) + Number(mrnAmountOther)
  ).toFixed(2);

  const norm = (s: string) => (s || "").toLowerCase().trim();

  const isPaid = norm(itemData?.Status) === "paid";
  const isPendingUTR = norm(itemData?.Status) === "pending for utr update";
  const isSentBack = norm(itemData?.Status) === "send back";
  const isSaveAsDraft = norm(itemData?.Status) === "save as draft";
  const isWithRequester = isSentBack || isSaveAsDraft;

  const currentApproverId = Number(itemData?.CurrentApproverId);

  const currentApproverMatrixIndex = approvalMatrix.findIndex(
    (a: any) => Number(a.Id) === currentApproverId,
  );

  const initiatorClass: string = (() => {
    if (isPaid) return "approved";
    if (isWithRequester) return "current";
    if (currentApproverMatrixIndex >= 0) return "approved";
    return "pending";
  })();

  const getApproverRibbonClass = (approver: any, index: number): string => {
    if (isPaid) return "approved";
    if (isWithRequester) return "pending";

    const approverHistory = workflowHistory.find(
      (x: any) =>
        norm(x.ActionBy || x.CurrentApprover || "") === norm(approver.Name),
    );

    if (
      norm(approverHistory?.ActionTaken) === "reject" ||
      norm(approverHistory?.ActionTaken) === "rejected" ||
      norm(approverHistory?.Action) === "reject" ||
      norm(approverHistory?.Action) === "rejected"
    ) {
      return "rejected";
    }

    if (
      norm(approverHistory?.ActionTaken) === "approved" ||
      norm(approverHistory?.Action) === "approved"
    ) {
      return "approved";
    }

    if (isPendingUTR) {
      const performerIndex = approvalMatrix.findIndex(
        (x: any) => x.Role && x.Role.toLowerCase().trim() === "performer",
      );
      if (index === performerIndex) return "current";
      if (index < performerIndex) return "approved";
    }

    if (Number(approver.Id) === currentApproverId) {
      return "current";
    }

    if (index < currentApproverMatrixIndex) {
      return "approved";
    }

    return "pending";
  };

  const getAttachments = async (capexId: string) => {
    try {
      const safe = capexId.replace(/\//g, "_");
      const path = `/sites/SonaFinance/InstallationCommision/${safe}`;
      const files = await sp.web.getFolderByServerRelativePath(path).files();
      setAttachments(files);
    } catch {
      setAttachments([]);
    }
  };

  const getPastMRNDetails = async (selectedPONumber: string) => {
    if (!selectedPONumber) { setPreviousAdvances([]); return; }
    try {
      const result = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.select(
          "PONumber", "PODate", "POAmount", "MRNNumber", "MRNDtae",
          "MRNAmountwithGST", "RequestedAmountforPayment",
          "VoucherDate", "VoucherNumber", "Status"
        )
        .filter(`PONumber eq '${selectedPONumber}' and Status eq 'Paid'`)
        .orderBy("Created", false)();
      setPreviousAdvances(result);
    } catch (error) {
      console.error("Error fetching Past MRN Details:", error);
      setPreviousAdvances([]);
    }
  };

  const getAdvanceHistory = async (selectedPONumber: string) => {
    if (!selectedPONumber) { setAdvanceHistory([]); return; }
    try {
      const result = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.select(
          "PONumber", "RequestAdvanceAmount", "Created",
          "VoucherDate", "PaidAmount", "VouchingNumber", "Status"
        )
        .filter(`PONumber eq '${selectedPONumber}'`)
        .orderBy("Created", false)();
      setAdvanceHistory(result);
    } catch (error) {
      console.error("Error fetching Advance History:", error);
      setAdvanceHistory([]);
    }
  };

  const getVendors = async () => {
    try {
      const data = await sp.web.lists
        .getByTitle("VendorMaster")
        .items.select("Id", "VendorCode", "VendorName")();
      setVendors(data);
    } catch (error) {
      console.error("Vendor fetch error:", error);
    }
  };

  const getItemById = async () => {
    try {
      const item = await sp.web.lists
        .getByTitle("Installation")
        .items.getById(itemId)
        .select("*")
        .expand("")();

      setItemData(item);
      setApproverRemarks(item.ApproverRemarks || "");

      const matchedVendor = vendors.find(
        (v) => String(v.VendorCode).trim() === String(item.VendorCode).trim(),
      );
      setSelectedVendorId(matchedVendor?.Id || null);
      setSelectedVendorName(item.VendorName || "");
      setSelectedVendorCode(item.VendorCode || "");

      if (item.CapexID) {
        await getAttachments(item.CapexID);
      } else if (item.PaymentId) {
        await getAttachments(item.PaymentId);
      }

      if (item.PONumber) {
        await getPastMRNDetails(item.PONumber);
        await getAdvanceHistory(item.PONumber);
      }

      if (item.ApprovalMatrix) {
        try {
          const parsed =
            typeof item.ApprovalMatrix === "string"
              ? JSON.parse(item.ApprovalMatrix)
              : item.ApprovalMatrix;
          setApprovalMatrix(Array.isArray(parsed) ? parsed : []);
        } catch (e) {
          console.error("ApprovalMatrix parse error", e);
          setApprovalMatrix([]);
        }
      } else {
        setApprovalMatrix([]);
      }

      if (item.WorkFlowHistory) {
        try {
          const parsed =
            typeof item.WorkFlowHistory === "string"
              ? JSON.parse(item.WorkFlowHistory)
              : item.WorkFlowHistory;
          setWorkflowHistory(Array.isArray(parsed) ? parsed : []);
        } catch (e) {
          console.error("WorkFlowHistory parse error", e);
          setWorkflowHistory([]);
        }
      } else {
        setWorkflowHistory([]);
      }
    } catch (error) {
      console.error("Fetch error:", error);
    }
  };

  useEffect(() => {
    if (!context || !itemId) return;
    const loadData = async () => {
      await getVendors();
      await getItemById();
    };
    void loadData();
  }, [context, itemId]);

  useEffect(() => {
    if (!itemData || vendors.length === 0) return;
    const vendor = vendors.find(
      (v) =>
        String(v.VendorCode).trim().toLowerCase() ===
        String(itemData.VendorCode).trim().toLowerCase(),
    );
    if (vendor) {
      setSelectedVendorId(vendor.Id);
      setSelectedVendorName(vendor.VendorName);
      setSelectedVendorCode(vendor.VendorCode);
    } else {
      setSelectedVendorId(null);
      setSelectedVendorName(itemData.VendorName || "");
      setSelectedVendorCode(itemData.VendorCode || "");
    }
  }, [itemData, vendors]);

  const handleApprove = async () => {
    try {
      if (!UTRDate || !UTRNumber || !UTRRemarks) {
        await Swal.fire({
          title: "Validation",
          text: "UTR Date, UTR Number and UTR Remarks are required.",
          icon: "warning",
        });
        return;
      }

      const flow = itemData.ApprovalMatrix ? JSON.parse(itemData.ApprovalMatrix) : [];
      const currentUserId = context.pageContext.legacyPageContext.userId;
      const currentIndex = flow.findIndex((a: any) => a.Id === currentUserId);
      if (currentIndex !== -1) {
        flow[currentIndex].Status = "Approved";
      }

      const history = itemData.WorkFlowHistory ? JSON.parse(itemData.WorkFlowHistory) : [];
      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Paid",
        Comment: UTRRemarks,
        Date: new Date().toISOString(),
      });

      await sp.web.lists.getByTitle("Installation").items.getById(itemId).update({
        ApproverRemarks: approverRemarks,
        UTRDate: new Date(UTRDate),
        UTRNumber: UTRNumber,
        Status: "Paid",
        ApprovalMatrix: JSON.stringify(flow),
        WorkFlowHistory: JSON.stringify(history),
        CurrentApproverId: null,
      });

      await Swal.fire({ title: "Success", text: "Paid Successfully", icon: "success" });
      onClose();
    } catch (error) {
      console.error(error);
      await Swal.fire({ title: "Error", text: "Something went wrong.", icon: "error" });
    }
  };

  const handleSendBack = async () => {
    try {
      if (!UTRRemarks) {
        await Swal.fire({ title: "Validation", text: "Please enter UTR Remarks.", icon: "warning" });
        return;
      }

      const flow = itemData.ApprovalMatrix ? JSON.parse(itemData.ApprovalMatrix) : [];
      const currentUserId = context.pageContext.legacyPageContext.userId;
      const currentIndex = flow.findIndex((a: any) => a.Id === currentUserId);
      if (currentIndex !== -1) {
        flow[currentIndex].Status = "Send Back";
      }

      let previousApproverId = null;
      if (flow[currentIndex - 1]) {
        flow[currentIndex - 1].Status = "In Progress";
        previousApproverId = flow[currentIndex - 1].Id;
      }

      const history = itemData.WorkFlowHistory ? JSON.parse(itemData.WorkFlowHistory) : [];
      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Send Back",
        Comment: UTRRemarks,
        Date: new Date().toISOString(),
      });

      await sp.web.lists.getByTitle("Installation").items.getById(itemId).update({
        ApproverRemarks: approverRemarks,
        Status: "Send Back",
        ApprovalMatrix: JSON.stringify(flow),
        WorkFlowHistory: JSON.stringify(history),
        CurrentApproverId: previousApproverId,
      });

      await Swal.fire({ title: "Success", text: "Request Sent Back Successfully", icon: "success" });
      onClose();
    } catch (error) {
      console.error(error);
    }
  };

  const handleReject = async () => {
    try {
      if (!UTRRemarks) {
        await Swal.fire({ title: "Validation", text: "Please enter UTR Remarks.", icon: "warning" });
        return;
      }

      const flow = itemData.ApprovalMatrix ? JSON.parse(itemData.ApprovalMatrix) : [];
      const currentUserId = context.pageContext.legacyPageContext.userId;
      const currentIndex = flow.findIndex((a: any) => a.Id === currentUserId);
      if (currentIndex !== -1) {
        flow[currentIndex].Status = "Reject";
      }

      const history = itemData.WorkFlowHistory ? JSON.parse(itemData.WorkFlowHistory) : [];
      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Reject",
        Comment: UTRRemarks,
        Date: new Date().toISOString(),
      });

      await sp.web.lists.getByTitle("Installation").items.getById(itemId).update({
        ApproverRemarks: approverRemarks,
        Status: "Reject",
        ApprovalMatrix: JSON.stringify(flow),
        WorkFlowHistory: JSON.stringify(history),
        CurrentApproverId: null,
      });

      await Swal.fire({ title: "Success", text: "Request Rejected Successfully", icon: "success" });
      onClose();
    } catch (error) {
      console.error(error);
    }
  };

  const handleExit = () => {
    onClose();
  };

  if (!itemData) return <div>Loading...</div>;

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            <div className="bordered">
              <img src={logo} />
              <h1> Installation Commisioning Request(AP Performer - UTR) </h1>
            </div>
            <div className="approval-ribbon">
              <div className={`ribbon-step ${initiatorClass}`}>
                {itemData.EmployeeName}
              </div>
              {approvalMatrix.map((approver: any, index: number) => (
                <div
                  key={index}
                  className={`ribbon-step ${getApproverRibbonClass(approver, index)}`}
                >
                  {approver.Name}
                  <br />
                  <small>{approver.Role}</small>
                </div>
              ))}
            </div>
            <div className="borderedbox">
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Requestor Information</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Employee Code" className="font">Employee Code</label> : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Name" className="font">Employee Name </label> : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Email" className="font">Employee Email </label> : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Email}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Contact No" className="font">Contact No</label> : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Status" className="font">Employee Status</label> : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.EmployeeStatus}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Division" className="font">Division</label> : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Location" className="font">Location</label> : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="RM" className="font">RM</label> : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.ReportingManager}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="HOD" className="font">HOD</label> : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.HOD}</label>
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Vendor &amp; PO Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Vendor Code</label>
                    <input type="text" value={itemData?.VendorCode || selectedVendorCode || ""} className="form-control readonly" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Name</label>
                    <input value={itemData.VendorName || ""} className="form-control readonly" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number</label>
                    <input value={itemData.PONumber || ""} className="form-control readonly" readOnly />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date</label>
                    <input type="date" value={itemData.POdate ? new Date(itemData.POdate).toISOString().split("T")[0] : ""} className="form-control readonly" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Payment Terms</label>
                    <input value={itemData.POPaymentTerms || ""} className="form-control readonly" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount (Incl. GST)</label>
                    <input value={itemData.POAmount || ""} className="form-control readonly" readOnly />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Basic Amount</label>
                    <input value={poAmountBasic} className="form-control readonly computed-field" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO GST Amount</label>
                    <input value={poAmountGST} className="form-control readonly computed-field" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Other Amount</label>
                    <input value={poAmountOther} className="form-control readonly computed-field" readOnly />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">MRN Basic Amount</label>
                    <input value={mrnAmountBasic} className="form-control readonly computed-field" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN GST Amount</label>
                    <input value={mrnAmountGST} className="form-control readonly computed-field" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Other Amount</label>
                    <input value={mrnAmountOther} className="form-control readonly computed-field" readOnly />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Amount Total</label>
                    <input value={poAmountTotal} className="form-control readonly computed-field" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Amount Total</label>
                    <input value={mrnAmountTotal} className="form-control readonly computed-field" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Total Amount to be Capitalized</label>
                    <input value={itemData.TotalamounttobeCapitalized || ""} className="form-control readonly" readOnly />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Whether GST to be Capitalized</label>
                    <input value={itemData.GSTToBeCapitalized || ""} className="form-control readonly" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Asset Code(s)</label>
                    <input value={itemData.AssetCodes || ""} className="form-control readonly" readOnly />
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Past MRN Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    <div style={{ overflowX: "auto" }}>
                      <table className="custom-table">
                        <thead>
                          <tr>
                            <th className="px-4 py-2">PO Number</th>
                            <th className="px-4 py-2">PO Date</th>
                            <th className="px-4 py-2">PO Amount</th>
                            <th className="px-4 py-2">MRN No</th>
                            <th className="px-4 py-2">MRN Date</th>
                            <th className="px-4 py-2">MRN Amount</th>
                            <th className="px-4 py-2">Advance Adjustment</th>
                            <th className="px-4 py-2">Paid Amount</th>
                            <th className="px-4 py-2">Remarks</th>
                          </tr>
                        </thead>
                        <tbody>
                          {previousAdvances.length === 0 ? (
                            <tr><td colSpan={9} style={{ textAlign: "center" }}>No Data</td></tr>
                          ) : (
                            previousAdvances.map((item: any, index: number) => (
                              <tr key={index}>
                                <td className="px-4 py-2">{item.PONumber}</td>
                                <td className="px-4 py-2">{item.PODate ? new Date(item.PODate).toLocaleDateString() : ""}</td>
                                <td className="px-4 py-2">{item.POAmount}</td>
                                <td className="px-4 py-2">{item.MRNNumber}</td>
                                <td className="px-4 py-2">{item.MRNDtae ? new Date(item.MRNDtae).toLocaleDateString() : ""}</td>
                                <td className="px-4 py-2">{item.MRNAmountwithGST}</td>
                                <td className="px-4 py-2"></td>
                                <td className="px-4 py-2">{item.RequestedAmountforPayment}</td>
                                <td className="px-4 py-2"></td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Advance History (to be PO Specific)</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    <div style={{ overflowX: "auto" }}>
                      <table className="custom-table">
                        <thead>
                          <tr>
                            <th className="px-4 py-2">PO Number</th>
                            <th className="px-4 py-2">Previous Advance</th>
                            <th className="px-4 py-2">Amount Requested Date</th>
                            <th className="px-4 py-2">Amount Paid Date</th>
                            <th className="px-4 py-2">MRN No</th>
                            <th className="px-4 py-2">Settled Amount</th>
                            <th className="px-4 py-2">Pending Advance</th>
                          </tr>
                        </thead>
                        <tbody>
                          {advanceHistory.length === 0 ? (
                            <tr><td colSpan={7} style={{ textAlign: "center" }}>No Data</td></tr>
                          ) : (
                            advanceHistory.map((item: any, index: number) => {
                              const pending = Math.max(0, Number(item.RequestAdvanceAmount || 0) - Number(item.PaidAmount || 0));
                              return (
                                <tr key={index}>
                                  <td className="px-4 py-2">{item.PONumber}</td>
                                  <td className="px-4 py-2">{item.RequestAdvanceAmount}</td>
                                  <td className="px-4 py-2">{item.Created ? new Date(item.Created).toLocaleDateString() : ""}</td>
                                  <td className="px-4 py-2">{item.VoucherDate ? new Date(item.VoucherDate).toLocaleDateString() : ""}</td>
                                  <td className="px-4 py-2"></td>
                                  <td className="px-4 py-2">{item.PaidAmount}</td>
                                  <td className="px-4 py-2">{pending}</td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Voucher Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Voucher Date</label>
                    <input className="form-control readonly" readOnly value={itemData.VoucherDate ? new Date(itemData.VoucherDate).toISOString().split("T")[0] : ""} type="date" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Voucher Number</label>
                    <input className="form-control readonly" readOnly value={itemData.VoucherNumber || ""} />
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Attachments</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attachments</label>
                    {attachments.length === 0 ? (
                      <p>No attachments</p>
                    ) : (
                      <ul>
                        {attachments.map((file: any, index: number) => (
                          <li key={index}>
                            <a href={file.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">{file.Name}</a>
                          </li>
                        ))}
                      </ul>
                    )}
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Workflow History</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    {workflowHistory.length === 0 ? (
                      <p>No history available</p>
                    ) : (
                      <div className="workflow-history">
                        <table>
                          <thead>
                            <tr>
                              <th>Action By</th>
                              <th>Action Taken</th>
                              <th>Date</th>
                              <th>Comment</th>
                            </tr>
                          </thead>
                          <tbody>
                            {workflowHistory.map((h: any, index: number) => (
                              <tr key={index}>
                                <td>{h.CurrentApprover}</td>
                                <td>{h.ActionTaken}</td>
                                <td>{h.Date ? new Date(h.Date).toLocaleString() : ""}</td>
                                <td>{h.Comment}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>UTR Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">UTR Date</label>
                    <input type="date" value={UTRDate} className="form-control" onChange={(e) => setUTRDate(e.target.value)} />
                  </div>
                  <div className="col-md-4">
                    <label className="font">UTR Number</label>
                    <input value={UTRNumber} className="form-control" onChange={(e) => setUTRNumber(e.target.value)} />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-12">
                    <label className="font">UTR Remarks</label>
                    <textarea className="form-control" rows={4} value={UTRRemarks} onChange={(e) => setUTRRemarks(e.target.value)} />
                  </div>
                </div>
              </div>

              <div style={{ display: "flex", justifyContent: "center", gap: "5px", marginBottom: "1rem", marginTop: "1rem" }}>
                <a onClick={handleApprove} className="submit-btn">Paid</a>
                <a onClick={handleSendBack} className="Rework-btn">Sent Back</a>
                <a onClick={handleReject} className="Reject-btn">Reject</a>
                <a onClick={handleExit} className="reset-btn">Exit</a>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default APperformerAdvanceFormForUTR;   