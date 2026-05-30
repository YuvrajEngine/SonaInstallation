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
interface IProps {
  context: any;
  itemId: number; // 👈 IMPORTANT,
  formData: any; // 👈 IMPORTANT
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
}) => {
  const sp = spfi().using(SPFx(context));
  const [attachments, setAttachments] = useState<any[]>([]);
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [itemData, setItemData] = useState<any>(null);
  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [voucherNumber, setVoucherNumber] = useState("");
  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");
  const [UTRRemarks, setUTRRemarks] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);
  // ✅ Fetch Item by ID
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
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
  const getPreviousAdvances = async (vendorId: number) => {
    try {
      debugger;
      console.log("Fetching for Vendor:", vendorId);

      const data = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.select(
          "PONumber",
          "RequestAdvanceAmount",
          "Created",
          "VoucherDate",

          "PaidAmount",
          "Status",
          "VendorCode/Id",
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
  // ✅ Fetch Item by ID
  const getItemById = async () => {
    try {
      const item = await sp.web.lists
        .getByTitle("Installation")
        .items.getById(itemId)
        .select("*")
        .expand("")();
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
      // setGstAdjustment(Number(item.GSTAdjustmentifAny || 0));
      //setOtherAdjustment(Number(item.OtherAdjustmentifany || 0));

      // ✅ FETCH ATTACHMENTS
      if (item.PaymentId) {
        await getAttachments(item.PaymentId);
      }
      // ✅ Approval Matrix (SAFE PARSE)
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

      // ✅ Workflow History (SAFE PARSE)
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
      //await getItemById(); // 👈 FIRST load item to get VendorCode
      await getVendors(); // 👈 FIRST load vendors
      await getItemById(); // 👈 THEN item
    };

    void loadData();
  }, [context, itemId]);


  // ✅ Approve
  const handleApprove = async () => {
    try {
      if (!UTRDate || !UTRNumber || !UTRRemarks) {
        alert("All fields required");
        return;
      }

      // 🔥 FLOW
      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentIndex = flow.findIndex(
        (a: any) => a.Id === currentUserId
      );

      if (currentIndex !== -1) {
        flow[currentIndex].Status = "Approved";
      }

      // 🔥 HISTORY
      const history = itemData.WorkFlowHistory
        ? JSON.parse(itemData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Paid",
        Comment: UTRRemarks,
        Date: new Date().toISOString()
      });

      await sp.web.lists
        .getByTitle("Installation")
        .items.getById(itemId)
        .update({
          ApproverRemarks: approverRemarks,
          UTRDate: new Date(UTRDate),
          UTRNumber: UTRNumber,
          //UTRRemarks: UTRRemarks,

          Status: "Paid",

          ApprovalMatrix: JSON.stringify(flow),
          WorkFlowHistory: JSON.stringify(history),

          CurrentApproverId: null // ✅ FINAL STAGE
        });

      alert("Paid successfully ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/Commision.aspx?page=Performer";

    } catch (error) {
      console.error(error);
      alert("Error ❌");
    }
  };

  // ✅ Sent Back
  const handleSendBack = async () => {
    try {
      if (!UTRRemarks) {
        alert("Please enter remarks");
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

      // 🔥 MOVE BACK
      let previousApproverId = null;
      if (flow[currentIndex - 1]) {
        flow[currentIndex - 1].Status = "In Progress";
        previousApproverId = flow[currentIndex - 1].Id;
      }

      // 🔥 HISTORY
      const history = itemData.WorkFlowHistory
        ? JSON.parse(itemData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Send Back",
        Comment: UTRRemarks,
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

          CurrentApproverId: null
        });

      alert("Send Back ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/Commision.aspx?page=Performer";

    } catch (error) {
      console.error(error);
    }
  };

  // ✅ Reject
  const handleReject = async () => {
    try {
      if (!UTRRemarks) {
        alert("Please enter remarks");
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
        Comment: UTRRemarks,
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

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/Commision.aspx?page=Performer";

    } catch (error) {
      console.error(error);
    }
  };

  const handleExit = () => {
    window.location.href = `https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/Commision.aspx?page=Performer`;
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
              <h1> Advance Payment (Approver) </h1>
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
                      disabled={true}
                      className='formtext-control'
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
                    <input value={itemData.VendorName} className='form-control readonly' />
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
                    <label className='font'>PO Terms</label>
                    <input value={itemData.POAdvanceTerms || ""} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>PO Amount </label>
                    <input value={itemData.POAmtGST || ""} className='form-control readonly' />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Capatalized Amount</label>
                    <input value={itemData.RequestAdvanceAmount || ""} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Paid Amount</label>
                    <input value={itemData.PaidAmount || ""} className='form-control readonly' />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Advance History(to be PO Specific)</label>
              </div>
              <div className='main-formcontainer'>
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
                                  <td className="px-4 py-2">{item.RequestAdvanceAmount}</td>
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
              <div className='main-formcontainer'>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Voucher Date</label>
                    <input className="form-control readonly" value={itemData.VoucherDate ? new Date(itemData.VoucherDate).toLocaleDateString("en-GB") : ""} />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Voucher Number</label>
                    <input className="form-control readonly" value={itemData.VouchingNumber || ""} />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-12">
                    <label className="font">Approver Remarks</label>
                    <label className="fonttext textbox readonly" style={{ height: "auto", width: "100%" }}>{approverRemarks}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Attachments</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
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
                  <div className='col-md-4'>
                    <label className="font">UTR Date</label>
                    <input type="date" value={UTRDate} className="form-control" onChange={(e) => setUTRDate(e.target.value)} />
                  </div>
                  <div className='col-md-4'>
                    <label className="font">UTR Number</label>
                    <input value={UTRNumber} className="form-control" onChange={(e) => setUTRNumber(e.target.value)} />
                  </div>

                </div>
                <div className='row mb-20'>
                  <div className='col-md-12'>
                    <label className="font">UTR Remarks</label>
                    <input className="form-control" onChange={(e) => setUTRRemarks(e.target.value)} />
                  </div>
                </div>
              </div>
              <div style={{ display: "flex", justifyContent: "center", gap: "5px", marginBottom: "1rem", marginTop: "1rem" }}>
                <a onClick={handleApprove} className="submit-btn">
                  Paid
                </a>
                <a onClick={handleSendBack} className="Rework-btn">
                  Sent Back
                </a>
                <a onClick={handleReject} className="Reject-btn">
                  Reject
                </a>
                <a href="#" onClick={handleExit} className="reset-btn">
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

export default APperformerAdvanceFormForUTR;
