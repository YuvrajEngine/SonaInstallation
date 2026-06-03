import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import Swal from "sweetalert2";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
interface IProps {
  context: any;
  itemId: number;
  formData: any;
  onClose: () => void;
}

import logo from "../assets/sona-comstarlogo.png";

interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}
const APperformerAdvanceform: React.FC<IProps> = ({
  context,
  itemId,
  formData,
  onClose,
}) => {
  const sp = spfi().using(SPFx(context));
  const [attachments, setAttachments] = useState<any[]>([]);
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [itemData, setItemData] = useState<any>(null);
  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [voucherNumber, setVoucherNumber] = useState("");
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [gstAdjustment, setGstAdjustment] = useState<number>(0);
  const [otherAdjustment, setOtherAdjustment] = useState<number>(0);
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const totalCapitalizedAmount =
    Number(itemData?.TotalPaymentofProject || 0) +
    Number(gstAdjustment || 0) +
    Number(otherAdjustment || 0);

  const norm = (s: string) => (s || "").toLowerCase().trim();

  const isPaid = norm(itemData?.Status) === "paid";
  const isPendingUTR = norm(itemData?.Status) === "pending for utr update";
  const isSentBack = norm(itemData?.Status) === "send back";
  const isSaveAsDraft = norm(itemData?.Status) === "save as draft";
  const isWithRequester = isSentBack || isSaveAsDraft;

  const currentApproverId = Number(itemData?.CurrentApproverId);

  const currentApproverMatrixIndex = approvalMatrix.findIndex(
    (a: any) => Number(a.Id) === currentApproverId
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
        norm(x.ActionBy || x.CurrentApprover || "") === norm(approver.Name)
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
        (x: any) => x.Role && x.Role.toLowerCase().trim() === "performer"
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

  const getPreviousAdvances = async (vendorId: number) => {
    try {
      debugger;
      console.log("Fetching for Vendor:", vendorId);

      const data = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.select(
          "PONumber",
          "TotalPaymentofProject",
          "Created",
          "VoucherDate",
          "PaidAmount",
          "Status",
          "VendorCode",
          "VoucherNumber",
        )
        .expand("VendorCode")
        .filter(`VendorCode/Id eq ${vendorId} and Status eq 'Paid'`)
        .orderBy("Created", false)();

      console.log("DATA:", data);

      void setPreviousAdvances(data);
    } catch (error) {
      console.error("Error fetching previous advances:", error);
      void setPreviousAdvances([]);
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
  const getAttachments = async (capexId: string) => {
    try {
      const safe = capexId.replace(/\//g, "_");
      const path = `/sites/SonaFinance/InstallationCommision/${safe}`;

      const files = await sp.web.getFolderByServerRelativePath(path).files();

      void setAttachments(files);
    } catch {
      void setAttachments([]);
    }
  };
  const getItemById = async () => {
    try {
      const item = await sp.web.lists
        .getByTitle("Installation")
        .items.getById(itemId)();
      setItemData(item);
      setApproverRemarks("");
      const matchedVendor = vendors.find(
        (v) => v.VendorCode === item.VendorCode,
      );

      setSelectedVendorId(matchedVendor?.Id || null);
      setSelectedVendorName(item.VendorName || "");

      setSelectedVendorName(item.VendorName);
      setGstAdjustment(Number(item.GSTAdjustmentifAny || 0));
      setOtherAdjustment(Number(item.OtherAdjustmentifany || 0));

      if (item.CapexID) {
        await getAttachments(item.CapexID);
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

    console.log("Saved VendorCode:", itemData.VendorCode);

    const vendor = vendors.find(
      (v) => String(v.VendorCode).trim() === String(itemData.VendorCode).trim(),
    );
    if (vendor) {
      setSelectedVendorId(vendor.Id);
      setSelectedVendorName(vendor.VendorName);
      setSelectedVendorCode(vendor.VendorCode);
      void getPreviousAdvances(vendor.Id);
    } else {
      setSelectedVendorId(null);
      setSelectedVendorName(itemData.VendorName || "");
      setSelectedVendorCode(itemData.VendorCode || "");
    }
  }, [itemData, vendors]);

  const handleApprove = async () => {
    try {
      if (!voucherDate || !voucherNumber || !approverRemarks?.trim()) {
        await Swal.fire({
          title: "Validation",
          text: "Voucher Date, Voucher Number and Remarks are required.",
          icon: "warning",
        });
        return;
      }
      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentIndex = flow.findIndex((a: any) => a.Id === currentUserId);

      if (currentIndex !== -1) {
        flow[currentIndex].Status = "Approved";
      }

      const history = itemData.WorkFlowHistory
        ? JSON.parse(itemData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Vouched",
        Comment: approverRemarks,
        Date: new Date().toISOString(),
      });

      await sp.web.lists
        .getByTitle("Installation")
        .items.getById(itemId)
        .update({
          ApproverRemarks: "",
          VoucherDate: new Date(voucherDate),
          VoucherNumber: voucherNumber,
          GSTAdjustmentifAny: gstAdjustment.toString(),
          OtherAdjustmentifany: otherAdjustment.toString(),
          TotalamounttobeCapitalized: totalCapitalizedAmount.toString(),

          Status: "Pending for UTR Update",
          ApprovalMatrix: JSON.stringify(flow),
          WorkFlowHistory: JSON.stringify(history),
        });

      await Swal.fire({
        title: "Success",
        text: "Approved Successfully",
        icon: "success",
      });
      onClose();
    } catch (error) {
      console.error(error);
    }
  };

  const handleSendBack = async () => {
    try {
      if (!voucherDate || !voucherNumber || !approverRemarks?.trim()) {
        await Swal.fire({
          title: "Validation",
          text: "Voucher Date, Voucher Number and Remarks are required.",
          icon: "warning",
        });
        return;
      }

      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

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

      const history = itemData.WorkFlowHistory
        ? JSON.parse(itemData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Send Back",
        Comment: approverRemarks,
        Date: new Date().toISOString(),
      });

      await sp.web.lists
        .getByTitle("Installation")
        .items.getById(itemId)
        .update({
          ApproverRemarks: "",
          VoucherDate: new Date(voucherDate),
          VoucherNumber: voucherNumber,
          GSTAdjustmentifAny: gstAdjustment.toString(),
          OtherAdjustmentifany: otherAdjustment.toString(),
          TotalamounttobeCapitalized: totalCapitalizedAmount.toString(),

          Status: "Send Back",

          ApprovalMatrix: JSON.stringify(flow),
          WorkFlowHistory: JSON.stringify(history),
          CurrentApproverId: null,
        });

      await Swal.fire({
        title: "Success",
        text: "Request Sent Back Successfully",
        icon: "success",
      });
      onClose();
    } catch (error) {
      console.error(error);
    }
  };

  const handleReject = async () => {
    try {
      if (!voucherDate || !voucherNumber || !approverRemarks?.trim()) {
        await Swal.fire({
          title: "Validation",
          text: "Voucher Date, Voucher Number and Remarks are required.",
          icon: "warning",
        });
        return;
      }

      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentIndex = flow.findIndex((a: any) => a.Id === currentUserId);

      if (currentIndex !== -1) {
        flow[currentIndex].Status = "Reject";
      }

      const history = itemData.WorkFlowHistory
        ? JSON.parse(itemData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Reject",
        Comment: approverRemarks,
        Date: new Date().toISOString(),
      });

      await sp.web.lists
        .getByTitle("Installation")
        .items.getById(itemId)
        .update({
          ApproverRemarks: "",
          VoucherDate: new Date(voucherDate),
          VoucherNumber: voucherNumber,
          GSTAdjustmentifAny: gstAdjustment.toString(),
          OtherAdjustmentifany: otherAdjustment.toString(),
          TotalamounttobeCapitalized: totalCapitalizedAmount.toString(),

          Status: "Reject",

          ApprovalMatrix: JSON.stringify(flow),
          WorkFlowHistory: JSON.stringify(history),
          CurrentApproverId: null,
        });

      await Swal.fire({
        title: "Success",
        text: "Request Rejected Successfully",
        icon: "success",
      });
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
              <h1> Installation Commisioning Request(AP Performer) </h1>
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
                    <label htmlFor="Employee Code" className="font">
                      Employee Code
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Name" className="font">
                      Employee Name{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Email" className="font">
                      Employee Email{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Email}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Contact No" className="font">
                      Contact No
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Status" className="font">
                      Employee Status
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {itemData.EmployeeStatus}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Division" className="font">
                      Division
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Location" className="font">
                      Location
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="RM" className="font">
                      RM
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {itemData.ReportingManager}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="HOD" className="font">
                      HOD
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.HOD}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Vendor & PO Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Vendor Code</label>
                    <input
                      type="text"
                      value={itemData?.VendorCode || selectedVendorCode || ""}
                      className="form-control readonly"
                      readOnly
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Name</label>
                    <input
                      value={itemData.VendorName || ""}
                      className="form-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number</label>
                    <input
                      value={itemData.PONumber || ""}
                      className="form-control readonly"
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date</label>
                    <input
                      type="date"
                      value={
                        itemData.POdate
                          ? new Date(itemData.POdate).toLocaleDateString(
                              "en-GB",
                            )
                          : ""
                      }
                      className="form-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Payment Terms</label>
                    <input
                      value={itemData.POPaymentTerms || ""}
                      className="form-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount </label>
                    <input
                      value={itemData.POAmount || ""}
                      className="form-control readonly"
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">
                      Total Payment for the Project
                    </label>
                    <input
                      value={itemData.TotalPaymentofProject || ""}
                      className="form-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Gst Adjustment(Any)</label>
                    <input
                      type="number"
                      className="form-control"
                      value={gstAdjustment}
                      onChange={(e) => setGstAdjustment(Number(e.target.value))}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Other Adjustment(Any)</label>
                    <input
                      type="number"
                      className="form-control"
                      value={otherAdjustment}
                      onChange={(e) =>
                        setOtherAdjustment(Number(e.target.value))
                      }
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">
                      Total Project Amount to be Capitalized
                    </label>
                    <input
                      value={totalCapitalizedAmount.toFixed(2)}
                      className="form-control readonly"
                    />
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
                            <th className="px-4 py-2">Po Amount</th>
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
                            <tr>
                              <td colSpan={7} style={{ textAlign: "center" }}>
                                No Data
                              </td>
                            </tr>
                          ) : (
                            previousAdvances.map((item: any, index: number) => {
                              const pending = Math.max(
                                0,
                                Number(item.RequestAdvanceAmount || 0) -
                                  Number(item.PaidAmount || 0),
                              );

                              return (
                                <tr key={index}>
                                  <td className="px-4 py-2">{item.PONumber}</td>
                                  <td className="px-4 py-2">
                                    {item.RequestAdvanceAmount}
                                  </td>

                                  <td className="px-4 py-2">
                                    {item.Created
                                      ? new Date(
                                          item.Created,
                                        ).toLocaleDateString()
                                      : ""}
                                  </td>

                                  <td className="px-4 py-2">
                                    {item.VoucherDate
                                      ? new Date(
                                          item.VoucherDate,
                                        ).toLocaleDateString()
                                      : ""}
                                  </td>

                                  <td className="px-4 py-2">
                                    {item.VoucherNumber}
                                  </td>
                                  <td className="px-4 py-2">
                                    {item.PaidAmount}
                                  </td>
                                  <td className="px-4 py-2">{pending}</td>
                                  <td className="px-4 py-2">
                                    {item.PaidAmount}
                                  </td>
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
                <label>Advance History(to be PO Specific)</label>
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
                          {previousAdvances.length === 0 ? (
                            <tr>
                              <td colSpan={7} style={{ textAlign: "center" }}>
                                No Data
                              </td>
                            </tr>
                          ) : (
                            previousAdvances.map((item: any, index: number) => {
                              const pending = Math.max(
                                0,
                                Number(item.RequestAdvanceAmount || 0) -
                                  Number(item.PaidAmount || 0),
                              );

                              return (
                                <tr key={index}>
                                  <td className="px-4 py-2">{item.PONumber}</td>
                                  <td className="px-4 py-2">
                                    {item.RequestAdvanceAmount}
                                  </td>

                                  <td className="px-4 py-2">
                                    {item.Created
                                      ? new Date(
                                          item.Created,
                                        ).toLocaleDateString()
                                      : ""}
                                  </td>

                                  <td className="px-4 py-2">
                                    {item.VoucherDate
                                      ? new Date(
                                          item.VoucherDate,
                                        ).toLocaleDateString()
                                      : ""}
                                  </td>

                                  <td className="px-4 py-2">
                                    {item.VoucherNumber}
                                  </td>
                                  <td className="px-4 py-2">
                                    {item.PaidAmount}
                                  </td>
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

              <div
                className="main-formcontainer"
                style={{ marginTop: "10px !important" }}
              >
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Voucher Date</label>
                    <input
                      type="date"
                      className="form-control"
                      value={voucherDate}
                      onChange={(e) => setVoucherDate(e.target.value)}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Voucher Number</label>
                    <input
                      value={voucherNumber}
                      className="form-control"
                      onChange={(e) => setVoucherNumber(e.target.value)}
                    />
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Attachments</label>
              </div>
              <div className="main-formcontainer'">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attachments</label>
                    {attachments.length === 0 ? (
                      <p>No attachments</p>
                    ) : (
                      <ul>
                        {attachments.map((file: any, index: number) => (
                          <li key={index}>
                            <a
                              href={file.ServerRelativeUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                            >
                              {file.Name}
                            </a>
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
                        {workflowHistory.map((h, index) => (
                          <div key={index} className="history-item">
                            <div>
                              {h.ActionTaken === "Submitted" && "📩 "}
                              {h.ActionTaken === "Approved" && "✅ "}
                              {h.ActionTaken === "Rejected" && "❌ "}
                              {h.ActionTaken === "Send Back" && "↩ "}
                              {h.ActionTaken === "Vouched" && "💰 "}
                              {h.ActionTaken === "Paid" && "💸 "}
                              {h.ActionTaken}
                            </div>

                            <div>
                              <b>{h.CurrentApprover}</b>
                            </div>
                            <div>{h.Comment}</div>
                            <div>
                              {h.Date ? new Date(h.Date).toLocaleString() : ""}
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Approver Action</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    <label className="font">Approver Remarks</label>
                    <textarea
                      value={approverRemarks}
                      onChange={(e) => setApproverRemarks(e.target.value)}
                      className="form-control"
                      rows={4}
                    />{" "}
                  </div>
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  gap: "5px",
                  marginBottom: "1rem",
                  marginTop: "1rem",
                }}
              >
                <a onClick={handleApprove} className="submit-btn">
                  Approve
                </a>
                <a onClick={handleSendBack} className="Rework-btn">
                  Sent Back
                </a>
                <a onClick={handleReject} className="Reject-btn">
                  Reject
                </a>
                <a onClick={handleExit} className="reset-btn">
                  Exit
                </a>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default APperformerAdvanceform;
