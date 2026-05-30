import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
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
  onClose
}) => {
  const sp = spfi().using(SPFx(context));
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [itemData, setItemData] = useState<any>(null);
  const [approverRemarks, setApproverRemarks] = useState("");
  const [attachments, setAttachments] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [vendors, setVendors] = useState<IVendor[]>([]);
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

  const getAttachments = async (capexId: string) => {
    debugger;
    try {
      if (!capexId) return;

      const safeCapexId = capexId.replace(/\//g, "_");

      // ✅ FIXED PATH (MOST IMPORTANT)
      const folderPath = `InstallationCommision/${safeCapexId}`;

      console.log("Folder Path:", folderPath);

      const files = await sp.web
        .getFolderByServerRelativePath(folderPath)
        .files();

      console.log("Files:", files);

      setAttachments(files || []);
    } catch (error) {
      console.log("Attachment fetch error:", error);
      setAttachments([]);
    }
  };
  const getPreviousAdvances = async (vendorId: number) => {
    try {
      debugger;
      console.log("Fetching for Vendor:", vendorId);

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
        )
        .filter(`Status eq 'Paid'`)
        .orderBy("Created", false)();
      console.log("DATA:", data);

      void setPreviousAdvances(data);
    } catch (error) {
      console.error("Error fetching previous advances:", error);
      void setPreviousAdvances([]);
    }
  };

  const uploadAttachments = async (capexId: string) => {
    try {
      if (attachments.length === 0) return;

      const safeCapexId = capexId.replace(/\//g, "_"); // ✅ move here

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

  // ✅ Fetch Item by ID
  const getItemById = async () => {
    try {
      const item = await sp.web.lists
        .getByTitle("Installation")
        .items.getById(itemId)();
      // .select("*", "VendorCode/Id", "VendorCode/VendorCode")
      //  .expand( "VendorCode")();
      // 👈 ADD

      setItemData(item);
      setApproverRemarks(item.ApproverRemarks || "");

      // ✅ FIX: Set VendorId + Name
      const matchedVendor = vendors.find(
        (v) => v.VendorCode === item.VendorCode,
      );

      setSelectedVendorId(matchedVendor?.Id || null);
      setSelectedVendorName(item.VendorName || "");

      // 🔥 IMPORTANT
      setSelectedVendorName(item.VendorName); // optional
      setGstAdjustment(Number(item.GSTAdjustmentifAny || 0));
      setOtherAdjustment(Number(item.OtherAdjustmentifany || 0));


      // ✅ FETCH ATTACHMENTS
      if (item.PaymentId) {
        await getAttachments(item.PaymentId);
      }
      // ✅ Approval Matrix
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

      // ✅ Workflow History
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
      await getItemById(); // 👈 FIRST load item to get VendorCode
      await getVendors(); // 👈 FIRST load vendors
      //  await getItemById(); // 👈 THEN item
    };

    void loadData();
  }, [context, itemId]);

  // ✅ Approve
  const handleApprove = async () => {
    try {
      if (!approverRemarks) {
        alert("Please enter Remarks");
        return;
      }

      // 🔥 GET FLOW
      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentIndex = flow.findIndex(
        (a: any) => a.Id === currentUserId
      );

      if (currentIndex === -1) {
        alert("You are not current approver");
        return;
      }

      // ✅ UPDATE CURRENT
      flow[currentIndex].Status = "Approved";

      let nextApproverId = null;

      if (flow[currentIndex + 1]) {
        flow[currentIndex + 1].Status = "In Progress";
        nextApproverId = flow[currentIndex + 1].Id;
      }

      // 🔥 ROLE BASED STATUS
      const currentRole = flow[currentIndex]?.Role;

      let finalStatus = itemData.Status;

      if (currentRole === "RM") {
        finalStatus = "Pending for Approver";
      } else if (currentRole === "HOD") {
        finalStatus = "Pending for PF Approver";
      } else {
        finalStatus = nextApproverId ? "Pending" : "Approved";
      }

      // 🔥 HISTORY
      const history = itemData.WorkFlowHistory
        ? JSON.parse(itemData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Approved",
        Comment: approverRemarks,
        Date: new Date().toISOString()
      });

      await sp.web.lists
        .getByTitle("Installation")
        .items.getById(itemId)
        .update({
          ApproverRemarks: approverRemarks,
          GSTAdjustmentifAny: gstAdjustment.toString(),
          OtherAdjustmentifany: otherAdjustment.toString(),
          TotalamounttobeCapitalized: totalCapitalizedAmount.toString(),

          Status: finalStatus,

          ApprovalMatrix: JSON.stringify(flow),
          WorkFlowHistory: JSON.stringify(history),
          CurrentApproverId: nextApproverId
        });

      alert("Approved successfully ✅");
      onClose();
    } catch (error) {
      console.error(error);
    }
  };

  // ✅ Sent Back
  const handleSendBack = async () => {
    try {
      if (!approverRemarks) {
        alert("Please enter Remarks");
        return;
      }

      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentIndex = flow.findIndex(
        (a: any) => a.Id === currentUserId
      );

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
        Date: new Date().toISOString()
      });

      await sp.web.lists
        .getByTitle("Installation")
        .items.getById(itemId)
        .update({
          ApproverRemarks: approverRemarks,
          Status: "Send Back",

          ApprovalMatrix: JSON.stringify(flow),
          WorkFlowHistory: JSON.stringify(history),
          CurrentApproverId: previousApproverId
        });

      alert("Send Back ✅");

      onClose();
    } catch (error) {
      console.error(error);
    }
  };

  // ✅ Reject
  const handleReject = async () => {
    try {
      if (!approverRemarks) {
        alert("Please enter Remarks");
        return;
      }

      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentIndex = flow.findIndex(
        (a: any) => a.Id === currentUserId
      );

      if (currentIndex !== -1) {
        flow[currentIndex].Status = "Rejected";
      }

      const history = itemData.WorkFlowHistory
        ? JSON.parse(itemData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Rejected",
        Comment: approverRemarks,
        Date: new Date().toISOString()
      });

      await sp.web.lists
        .getByTitle("Installation")
        .items.getById(itemId)
        .update({
          ApproverRemarks: approverRemarks,
          Status: "Rejected",

          ApprovalMatrix: JSON.stringify(flow),
          WorkFlowHistory: JSON.stringify(history),
          CurrentApproverId: null
        });

      alert("Rejected ❌");

      onClose();

    } catch (error) {
      console.error(error);
    }
  };

  const handleExit = () => {
  onClose();
  };

  // ⛔ Wait until data loads
  if (!itemData) return <div>Loading...</div>;

  return (
    <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
      <div className='row'>
        <div className='col-md-12'>
          <div className='Main-Boxpoup'>
            {/* 🔹 Header */}
            <div className="bordered">
              <img src={logo} />
              <h1> Installation Commisioning Request(Approver) </h1>
            </div>
            {approvalMatrix.length === 0 ? (
              <p>No approval data</p>
            ) : (
              <div className="displayWF">
                <ul className="approval-flow">
                  {approvalMatrix.map((a, index) => (
                    <li
                      key={index}
                      className={`approval-step ${a.Status === "In Progress"
                        ? "active"
                        : a.Status === "Approved"
                          ? "approved"
                          : a.Status === "Rejected"
                            ? "rejected"
                            : a.Status === "Send Back"
                              ? "sendback"
                              : ""
                        }`}
                    >
                      {a.Role} - {a.Name}
                    </li>
                  ))}
                </ul>
              </div>
            )}
            <div className='borderedbox'>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Requestor Information</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Code" className='font'>Employee Code</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.EmployeeCode}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Name" className='font'>Employee Name </label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.EmployeeName}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Email" className='font'>Employee Email </label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.Email}</label>
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Contact No" className='font'>Contact No</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.ContactNo}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Status" className='font'>Employee Status</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.EmployeeStatus}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Division" className='font'>Division</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.Division}</label>
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Location" className='font'>Location</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.Location}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="RM" className='font'>RM</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.ReportingManager}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="HOD" className='font'>HOD</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {itemData.HOD}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Vendor & PO Details</label>
              </div>
              <div className='main-formcontainer'>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className='font'>Vendor Code</label>
                    <select
                      value={selectedVendorId ?? ""}
                      disabled={true} className="formtext-control readonly"
                      onChange={(e) => {
                        const id = Number(e.target.value);
                        const vendor = vendors.find((v) => v.Id === id);
                        setSelectedVendorId(id);
                        setSelectedVendorName(vendor?.VendorName || "");
                      }}
                    >
                      <option value="">Select Vendor</option>
                      {vendors.map((v) => (
                        <option key={v.Id} value={v.Id}>
                          {v.VendorCode}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="col-md-4">
                    <label className='font'>Vendor Name</label>
                    <input value={itemData.VendorName || ""} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>PO Number</label>
                    <input value={itemData.PONumber || ""} className='form-control readonly' />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className='font'>PO Date</label>
                    <input type="date" value={itemData.POdate ? new Date(itemData.POdate).toLocaleDateString("en-GB") : ""} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>PO Payment Terms</label>
                    <input value={itemData.POPaymentTerms || ""} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>PO Amount (GST)</label>
                    <input value={itemData.POAmount || ""} className='form-control readonly' />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Total Payment for the Project</label>
                    <input value={itemData.TotalPaymentofProject || ""} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Gst Adjustment(Any)</label>
                    <input type="number" value={gstAdjustment} onChange={(e) => setGstAdjustment(Number(e.target.value))} className='form-control' />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Other Adjustment</label>
                    <input type="number" value={otherAdjustment} onChange={(e) => setOtherAdjustment(Number(e.target.value))} className='form-control' />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font" style={{ color: "red" }}>Total Project Amount to be Capitalized</label>
                    <input value={totalCapitalizedAmount.toFixed(2)} className='form-control readonly' />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Past MRN Details</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-12'>
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
                            <th> Advance Adjustment</th>
                            <th>Paid Amount</th>
                            <th>Remarks</th>
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
                                Number(item.TotalPaymentofProject || 0) -
                                Number(item.PaidAmount || 0),
                              );

                              return (
                                <tr key={index}>
                                  <td className="px-4 py-2">{item.PONumber}</td>
                                  <td className="px-4 py-2">{item.TotalPaymentofProject}</td>

                                  <td className="px-4 py-2">
                                    {item.Created
                                      ? new Date(item.Created).toLocaleDateString()
                                      : ""}
                                  </td>

                                  <td className="px-4 py-2">
                                    {item.VoucherDate
                                      ? new Date(item.VoucherDate).toLocaleDateString()
                                      : ""}
                                  </td>

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
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-12'>
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
                            <tr>
                              <td colSpan={7} style={{ textAlign: "center" }}>
                                No Data
                              </td>
                            </tr>
                          ) : (
                            previousAdvances.map((item: any, index: number) => {
                              const pending = Math.max(
                                0,
                                Number(item.TotalPaymentofProject || 0) -
                                Number(item.PaidAmount || 0),
                              );

                              return (
                                <tr key={index}>
                                  <td className="px-4 py-2">{item.PONumber}</td>
                                  <td className="px-4 py-2">{item.TotalPaymentofProject}</td>

                                  <td className="px-4 py-2">
                                    {item.Created
                                      ? new Date(item.Created).toLocaleDateString()
                                      : ""}
                                  </td>

                                  <td className="px-4 py-2">
                                    {item.VoucherDate
                                      ? new Date(item.VoucherDate).toLocaleDateString()
                                      : ""}
                                  </td>

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
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-12'>
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
                            <div>
                              {h.Date
                                ? new Date(h.Date).toLocaleString()
                                : ""}
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
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-12'>
                    <label className="font">Approver Remarks</label>
                    <textarea value={approverRemarks} onChange={(e) => setApproverRemarks(e.target.value)} />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Upoad Document</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label className="font">Attachments</label>
                    {attachments.length === 0 ? (
                      <p>No attachments found</p>
                    ) : (
                      <ul>
                        {attachments.map((file: any, index: number) => (
                          <li key={index}>
                            <a href={file.ServerRelativeUrl} target="_blank">
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

export default ApproverAdvanceForm;
