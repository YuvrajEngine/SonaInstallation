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
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [itemData, setItemData] = useState<any>(null);
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

  const totalCapitalizedAmount =
    Number(itemData?.TotalPaymentofProject || 0) +
    Number(gstAdjustment || 0) +
    Number(otherAdjustment || 0);

  const norm = (s: string) => (s || "").toLowerCase().trim();

  const status = norm(itemData?.Status);

  const isPaid = status === "paid";
  const isSaveAsDraft = status === "save as draft";
  const isSentBack =
    status === "send back" ||
    workflowHistory.some(
      (x: any) =>
        norm(x.ActionTaken) === "send back" || norm(x.Action) === "send back"
    );
  const isPendingVouching = status === "pending for vouching update";
  const isPendingUTR = status === "pending for utr update";

  const rejectedEntry = workflowHistory.find(
    (w: any) => norm(w.ActionTaken) === "reject" || norm(w.Action) === "reject"
  );
  const rejectedByName = norm(
    rejectedEntry?.CurrentApprover || rejectedEntry?.ActionBy || ""
  );
  const isRejected = rejectedByName !== "";

  const currentApproverId = Number(itemData?.CurrentApproverId);

  const currentApproverIndex = approvalMatrix.findIndex(
    (a: any) => Number(a.Id) === currentApproverId
  );

  const initiatorRibbonClass: string = (() => {
    if (isPaid) return "approved";
    if (isSaveAsDraft || isSentBack) return "current";
    if (isPendingVouching || isPendingUTR) return "approved";
    if (isRejected && currentApproverIndex === -1) return "current";
    if (currentApproverIndex >= 0) return "approved";
    return "current";
  })();

  const getApproverRibbonClass = (approver: any, index: number): string => {
    if (isPaid) return "approved";
    if (isSaveAsDraft || isSentBack) return "pending";
    if (isPendingVouching || isPendingUTR) return "approved";
    if (norm(approver.Name) === rejectedByName) return "rejected";
    if (isRejected && index > currentApproverIndex) return "pending";
    if (Number(approver.Id) === currentApproverId) return "current";
    if (index < currentApproverIndex) return "approved";
    return "pending";
  };

  const getAttachments = async (capexId: string) => {
    debugger;
    try {
      if (!capexId) return;
      const safeCapexId = capexId.replace(/\//g, "_");
      const folderPath = `InstallationCommision/${safeCapexId}`;
      console.log("Folder Path:", folderPath);
      const files = await sp.web.getFolderByServerRelativePath(folderPath).files();
      console.log("Files:", files);
      setAttachments(files || []);
    } catch (error) {
      console.log("Attachment fetch error:", error);
      setAttachments([]);
    }
  };

  const getPreviousAdvances = async (vendorId: number) => {
    try {
      const vendor = vendors.find((v) => v.Id === vendorId);
      if (!vendor) return;
      const data = await sp.web.lists
        .getByTitle("Installation")
        .items.select(
          "PONumber",
          "TotalPaymentofProject",
          "Created",
          "VoucherDate",
          "PaidAmount",
          "Status",
          "VendorCode",
          "VoucherNumber"
        )
        .filter(`Status eq 'Paid' and VendorCode eq '${vendor.VendorCode}'`)
        .orderBy("Created", false)();
      console.log("DATA:", data);
      setPreviousAdvances(data);
    } catch (error) {
      console.error("Error fetching previous advances:", error);
      setPreviousAdvances([]);
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
    console.log("Installation VendorCode:", itemData.VendorCode);
    console.log("Vendor Master:", vendors);
    const vendor = vendors.find(
      (v) =>
        String(v.VendorCode || "").trim().toLowerCase() ===
        String(itemData.VendorCode || "").trim().toLowerCase()
    );
    console.log("Matched Vendor:", vendor);
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
        (x: any) => x.Role && x.Role.toLowerCase().trim() === "performer"
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
              <div className={`ribbon-step ${initiatorRibbonClass}`}>
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
                <label>Vendor & PO Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Vendor Code</label>
                    <input type="text" value={itemData?.VendorCode || selectedVendorCode || ""} className="form-control readonly" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Name</label>
                    <input value={itemData.VendorName || ""} className="form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number</label>
                    <input value={itemData.PONumber || ""} className="form-control readonly" />
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
                    <input value={itemData.POPaymentTerms || ""} className="form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount (GST)</label>
                    <input value={itemData.POAmount || ""} className="form-control readonly" />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Total Payment for the Project</label>
                    <input value={itemData.TotalPaymentofProject || ""} className="form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Gst Adjustment(Any)</label>
                    <input type="number" value={gstAdjustment} onChange={(e) => setGstAdjustment(Number(e.target.value))} className="form-control" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Other Adjustment</label>
                    <input type="number" value={otherAdjustment} onChange={(e) => setOtherAdjustment(Number(e.target.value))} className="form-control" />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font" style={{ color: "red" }}>Total Project Amount to be Capitalized</label>
                    <input value={totalCapitalizedAmount.toFixed(2)} className="form-control readonly" />
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
                            <th>PO Number</th>
                            <th>PO Date</th>
                            <th>Po Amount</th>
                            <th>MRN No</th>
                            <th>MRN Date</th>
                            <th>MRN Amount</th>
                            <th>Advance Adjustment</th>
                            <th>Paid Amount</th>
                            <th>Remarks</th>
                          </tr>
                        </thead>
                        <tbody>
                          {previousAdvances.length === 0 ? (
                            <tr><td colSpan={9} style={{ textAlign: "center" }}>No Data</td></tr>
                          ) : (
                            previousAdvances.map((item: any, index: number) => {
                              const pending = Math.max(0, Number(item.TotalPaymentofProject || 0) - Number(item.PaidAmount || 0));
                              return (
                                <tr key={index}>
                                  <td className="px-4 py-2">{item.PONumber}</td>
                                  <td className="px-4 py-2">{item.TotalPaymentofProject}</td>
                                  <td className="px-4 py-2">{item.Created ? new Date(item.Created).toLocaleDateString() : ""}</td>
                                  <td className="px-4 py-2">{item.VoucherDate ? new Date(item.VoucherDate).toLocaleDateString() : ""}</td>
                                  <td className="px-4 py-2">{item.VoucherNumber}</td>
                                  <td className="px-4 py-2">{item.PaidAmount}</td>
                                  <td className="px-4 py-2">{pending}</td>
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
                <label>Advance History(to be PO Specific)</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    <div style={{ overflowX: "auto" }}>
                      <table className="custom-table">
                        <thead>
                          <tr>
                            <th>PO Number</th>
                            <th>Previous Advance</th>
                            <th>Amount Requested Date</th>
                            <th>Amount Paid Date</th>
                            <th>MRN No</th>
                            <th>Settled Amount</th>
                            <th>Pending Advance</th>
                          </tr>
                        </thead>
                        <tbody>
                          {previousAdvances.length === 0 ? (
                            <tr><td colSpan={7} style={{ textAlign: "center" }}>No Data</td></tr>
                          ) : (
                            previousAdvances.map((item: any, index: number) => {
                              const pending = Math.max(0, Number(item.TotalPaymentofProject || 0) - Number(item.PaidAmount || 0));
                              return (
                                <tr key={index}>
                                  <td className="px-4 py-2">{item.PONumber}</td>
                                  <td className="px-4 py-2">{item.TotalPaymentofProject}</td>
                                  <td className="px-4 py-2">{item.Created ? new Date(item.Created).toLocaleDateString() : ""}</td>
                                  <td className="px-4 py-2">{item.VoucherDate ? new Date(item.VoucherDate).toLocaleDateString() : ""}</td>
                                  <td className="px-4 py-2">{item.VoucherNumber}</td>
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
                            <div><b>{h.CurrentApprover}</b></div>
                            <div>{h.Comment}</div>
                            <div>{h.Date ? new Date(h.Date).toLocaleString() : ""}</div>
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
                    <textarea value={approverRemarks} onChange={(e) => setApproverRemarks(e.target.value)} />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Upoad Document</label>
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
                            <a href={file.ServerRelativeUrl} target="_blank">{file.Name}</a>
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