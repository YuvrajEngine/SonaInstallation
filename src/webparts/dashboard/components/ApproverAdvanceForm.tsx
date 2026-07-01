import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import Swal from "sweetalert2";
require("sweetalert2/dist/sweetalert2.min.css");
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

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

const ApproverAdvanceForm: React.FC<IProps> = ({
  context,
  itemId,
  formData,
  onClose,
}) => {
  const sp = spfi().using(SPFx(context));
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [advanceHistory, setAdvanceHistory] = useState<any[]>([]);
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [itemData, setItemData] = useState<any>(null);
  const tenantUrl = context.pageContext.site.absoluteUrl.split("/sites/")[0];
  const vendorSp = spfi(`${tenantUrl}/sites/RLY_AccountsPayable_UAT`).using(SPFx(context));
  const [approverRemarks, setApproverRemarks] = useState("");
  const [attachments, setAttachments] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const [gstAdjustment, setGstAdjustment] = useState<number>(0);
  const [otherAdjustment, setOtherAdjustment] = useState<number>(0);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
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

  const totalCapitalizedAmount =
    Number(itemData?.TotalPaymentofProject || 0) +
    Number(gstAdjustment || 0) +
    Number(otherAdjustment || 0);

  const norm = (s: string) => (s || "").toLowerCase().trim();

  const isPaid = norm(itemData?.Status) === "paid";
  const isSentBack = norm(itemData?.Status) === "send back";
  const isSaveAsDraft = norm(itemData?.Status) === "save as draft";
  const isPendingVouching = norm(itemData?.Status) === "pending for vouching update";
  const isPendingUTR = norm(itemData?.Status) === "pending for utr update";
  const isWithRequester = isSentBack || isSaveAsDraft;

  const currentApproverId = Number(itemData?.CurrentApproverId);

  const currentApproverMatrixIndex = approvalMatrix.findIndex(
    (a: any) => Number(a.Id) === currentApproverId,
  );

  const initiatorClass: string = (() => {
    if (isPaid) return "approved";
    if (isPendingVouching || isPendingUTR) return "approved";
    if (isWithRequester) return "current";
    if (currentApproverMatrixIndex >= 0) return "approved";
    return "pending";
  })();

  const getApproverRibbonClass = (approver: any, index: number): string => {
    if (isPaid) return "approved";
    if (isPendingVouching || isPendingUTR) return "approved";
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
      if (!capexId) return;
      const safeCapexId = capexId.replace(/\//g, "_");
      const folderPath = `InstallationCommision/${safeCapexId}`;
      const files = await sp.web.getFolderByServerRelativePath(folderPath).files();
      setAttachments(files || []);
    } catch (error) {
      console.log("Attachment fetch error:", error);
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
          "VoucherDate", "VoucherNumber", "Status",
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
          "VoucherDate", "PaidAmount", "VouchingNumber", "Status",
        )
        .filter(`PONumber eq '${selectedPONumber}'`)
        .orderBy("Created", false)();
      setAdvanceHistory(result);
    } catch (error) {
      console.error("Error fetching Advance History:", error);
      setAdvanceHistory([]);
    }
  };

  const uploadAttachments = async (capexId: string) => {
    try {
      if (attachments.length === 0) return;
      const safeCapexId = capexId.replace(/\//g, "_");
      const folderPath = `CapexAdvanceDocs/${safeCapexId}`;
      await sp.web.folders.addUsingPath(folderPath);
      for (const file of attachments) {
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, file, { Overwrite: true });
      }
    } catch (error) {
      console.error("Upload error:", error);
    }
  };

  const getVendors = async () => {
    try {
      const data = await vendorSp.web.lists
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
        .items.getById(itemId)();

      setItemData(item);
      setApproverRemarks("");

      const matchedVendor = vendors.find((v) => v.VendorCode === item.VendorCode);
      setSelectedVendorId(matchedVendor?.Id || null);
      setSelectedVendorName(item.VendorName || "");
      setGstAdjustment(Number(item.GSTAdjustmentifAny || 0));
      setOtherAdjustment(Number(item.OtherAdjustmentifany || 0));

      if (item.PaymentId) {
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
    if (!itemData) return;
    const vendor = vendors.find(
      (v) =>
        String(v.VendorCode || "").trim().toLowerCase() ===
        String(itemData.VendorCode || "").trim().toLowerCase(),
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

  useEffect(() => {
    if (!context || !itemId) return;
    const loadData = async () => {
      try {
        await getVendors();
        await getItemById();
      } catch (err) {
        console.error("Load Error:", err);
        await Swal.fire({ title: "Error", text: "Failed to load data", icon: "error" });
      }
    };
    void loadData();
  }, [context, itemId]);

  const handleApprove = async () => {
    try {
      if (!approverRemarks?.trim()) {
        await Swal.fire({ title: "Validation", text: "Please enter Remarks", icon: "warning" });
        return;
      }
      const currentUserId = Number(context.pageContext.legacyPageContext.userId);
      if (Number(itemData.CurrentApproverId) !== currentUserId) {
        await Swal.fire({ title: "Access Denied", text: "You are not the current approver.", icon: "warning" });
        return;
      }
      const flow = itemData.ApprovalMatrix ? JSON.parse(itemData.ApprovalMatrix) : [];
      const currentIndex = flow.findIndex((a: any) => Number(a.Id) === currentUserId);
      if (currentIndex === -1) {
        await Swal.fire({ title: "Error", text: "User not found in approval matrix", icon: "error" });
        return;
      }
      flow[currentIndex].Status = "Approved";
      const performerIndex = flow.findIndex(
        (x: any) => x.Role && x.Role.toLowerCase().trim() === "performer",
      );
      const beforePerformerIndex = performerIndex > 0 ? performerIndex - 1 : -1;
      let nextApproverId: number | null = null;
      let finalStatus = "Pending for Approval";
      if (performerIndex !== -1 && currentIndex === beforePerformerIndex) {
        flow[performerIndex].Status = "In Progress";
        nextApproverId = Number(flow[performerIndex].Id);
        finalStatus = "Pending for Vouching Update";
      } else if (flow[currentIndex + 1]) {
        flow[currentIndex + 1].Status = "In Progress";
        nextApproverId = Number(flow[currentIndex + 1].Id);
        finalStatus = "Pending for Approval";
      } else {
        nextApproverId = null;
        finalStatus = "Approved";
      }
      const history = itemData.WorkFlowHistory ? JSON.parse(itemData.WorkFlowHistory) : [];
      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Approved",
        Comment: approverRemarks,
        Date: new Date().toISOString(),
      });
      await sp.web.lists.getByTitle("Installation").items.getById(itemId).update({
        ApproverRemarks: approverRemarks,
        GSTAdjustmentifAny: gstAdjustment.toString(),
        OtherAdjustmentifany: otherAdjustment.toString(),
        TotalamounttobeCapitalized: totalCapitalizedAmount.toString(),
        Status: finalStatus,
        ApprovalMatrix: JSON.stringify(flow),
        WorkFlowHistory: JSON.stringify(history),
        CurrentApproverId: nextApproverId,
      });
      await Swal.fire({ title: "Success", text: "Approved Successfully", icon: "success" });
      onClose();
    } catch (error) {
      console.error("Approval Error:", error);
      await Swal.fire({ title: "Error", text: "Something went wrong while approving.", icon: "error" });
    }
  };

  const handleSendBack = async () => {
    try {
      if (!approverRemarks) {
        await Swal.fire({ title: "Validation", text: "Please enter Remarks", icon: "warning" });
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
        Comment: approverRemarks,
        Date: new Date().toISOString(),
      });
      await sp.web.lists.getByTitle("Installation").items.getById(itemId).update({
        ApproverRemarks: approverRemarks,
        Status: "Send Back",
        ApprovalMatrix: JSON.stringify(flow),
        WorkFlowHistory: JSON.stringify(history),
        CurrentApproverId: previousApproverId,
      });
      await Swal.fire({ title: "Success", text: "Request sent back successfully", icon: "success" });
      onClose();
    } catch (error) {
      console.error(error);
    }
  };

  const buildApprovalFlow = async (employeeData: any) => {
    try {
      const flow: any[] = [];
      if (employeeData.ReportingManagerId) {
        flow.push({ Id: employeeData.ReportingManagerId, Name: employeeData.ReportingManager, Role: "RM", Status: "Pending" });
      }
      if (employeeData.HODId) {
        flow.push({ Id: employeeData.HODId, Name: employeeData.HOD, Role: "HOD", Status: "" });
      }
      const matrix = await sp.web.lists
        .getByTitle("InstallationCommisionApprovalMatrix")
        .items.select("Approver/Id", "Approver/Title", "Role/RoleName", "Level/Level")
        .expand("Approver", "Role", "Level")
        .filter("Status eq 'Active'")
        .orderBy("Level", true)();
      matrix.forEach((m: any) => {
        flow.push({ Id: m.Approver.Id, Name: m.Approver.Title, Role: m.Role.RoleName, Status: "" });
      });
      if (flow.length > 0) flow[0].Status = "In Progress";
      return flow;
    } catch (e) {
      console.log(e);
      return [];
    }
  };

  const handleReject = async () => {
    try {
      if (!approverRemarks?.trim()) {
        await Swal.fire({ title: "Validation", text: "Please enter Remarks", icon: "warning" });
        return;
      }
      const flow = itemData.ApprovalMatrix ? JSON.parse(itemData.ApprovalMatrix) : [];
      const currentUserId = context.pageContext.legacyPageContext.userId;
      const currentIndex = flow.findIndex((a: any) => a.Id === currentUserId);
      if (currentIndex !== -1) flow[currentIndex].Status = "Reject";
      const history = itemData.WorkFlowHistory ? JSON.parse(itemData.WorkFlowHistory) : [];
      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Reject",
        Comment: approverRemarks,
        Date: new Date().toISOString(),
      });
      await sp.web.lists.getByTitle("Installation").items.getById(itemId).update({
        ApproverRemarks: approverRemarks,
        Status: "Reject",
        ApprovalMatrix: JSON.stringify(flow),
        WorkFlowHistory: JSON.stringify(history),
        CurrentApproverId: null,
      });
      await Swal.fire({ title: "Success", text: "Request rejected successfully", icon: "success" });
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
              <h1> Installation Commisioning Request(Approver) </h1>
            </div>

            <div className="approval-ribbon">
              <div className={`ribbon-step ${initiatorClass}`}>
                {itemData.EmployeeName}
              </div>
              {approvalMatrix.map((approver: any, index: number) => (
                <div key={index} className={`ribbon-step ${getApproverRibbonClass(approver, index)}`}>
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
                    <label htmlFor="Employee Code" className="font">Employee Code</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Name" className="font">Employee Name </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Email" className="font">Employee Email </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Email}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Contact No" className="font">Contact No</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Status" className="font">Employee Status</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.EmployeeStatus}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Division" className="font">Division</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Location" className="font">Location</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="RM" className="font">RM</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.ReportingManager}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="HOD" className="font">HOD</label>{" "}
                    : &nbsp;&nbsp;
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
                    <input
                      type="text"
                      value={itemData?.VendorCode || selectedVendorCode || ""}
                      className="form-control readonly"
                      readOnly
                    />
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
                    <input
                      type="date"
                      value={itemData.POdate ? new Date(itemData.POdate).toISOString().split("T")[0] : ""}
                      className="form-control readonly"
                      readOnly
                    />
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
                            {/* <th className="px-4 py-2">MRN No</th> */}
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
                                {/* <td className="px-4 py-2">{item.MRNNumber}</td> */}
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
                            {/* <th className="px-4 py-2">MRN No</th>
                            <th className="px-4 py-2">Settled Amount</th>
                            <th className="px-4 py-2">Pending Advance</th> */}
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
                                  {/* <td className="px-4 py-2"></td>
                                  <td className="px-4 py-2">{item.PaidAmount}</td>
                                  <td className="px-4 py-2">{pending}</td> */}
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
                <label>Approver Action</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    <label className="font">Approver Remarks</label>
                    <textarea
                      value={approverRemarks}
                      onChange={(e) => setApproverRemarks(e.target.value)}
                    />
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Upload Document</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attachments</label>
                    {attachments.length === 0 ? (
                      <p>No attachments found</p>
                    ) : (
                      <ul>
                        {attachments.map((file: any, index: number) => (
                          <li key={index}>
                            <a href={file.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">
                              {file.Name}
                            </a>
                          </li>
                        ))}
                      </ul>
                    )}
                  </div>
                </div>
              </div>

              <div style={{ display: "flex", justifyContent: "center", gap: "5px", marginBottom: "1rem", marginTop: "1rem" }}>
                <a onClick={handleApprove} className="submit-btn">Approve</a>
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

export default ApproverAdvanceForm;