import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import logo from "../assets/sona-comstarlogo.png";

interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

const ViewAdvanceForm = ({ context, formData, onClose }: any) => {
  const sp = spfi().using(SPFx(context));

  const [attachments, setAttachments] = useState<any[]>([]);
  const [employee, setEmployee] = useState<any>({});
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [PONumber, setPONumber] = useState("");
  const [TotalPaymentofProject, setTotalPaymentofProject] = useState("");
  const [GSTAdjustmentifAny, setGSTAdjustmentifAny] = useState("");
  const [OtherAdjustmentifany, setOtherAdjustmentifany] = useState("");
  const [TotalamounttobeCapitalized, setTotalamounttobeCapitalized] = useState("");
  const [POdate, setPOdate] = useState("");
  const [POPaymentTerms, setPOPaymentTerms] = useState("");
  const [POAmount, setPOAmount] = useState("");
  const [vendorName, setVendorName] = useState("");
  const [approverDetails, setApproverDetails] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);
  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [VouchingNumber, setVouchingNumber] = useState("");
  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");

  const norm = (s: string) => (s || "").toLowerCase().trim();

  const status = norm(formData?.Status);

  const isPaid = status === "paid";
  const isSaveAsDraft = status === "save as draft";
  const isSentBack =
    status === "send back" ||
    workflowHistory.some(
      (x: any) =>
        norm(x.Action) === "sent back" || norm(x.ActionTaken) === "sent back"
    );
  const isPendingVouching = status === "pending for vouching update";
  const isPendingUTR = status === "pending for utr update";

  const rejectedEntry = workflowHistory.find(
    (w: any) =>
      norm(w.Action) === "rejected" || norm(w.ActionTaken) === "rejected"
  );
  const rejectedByName = norm(
    rejectedEntry?.ActionBy || rejectedEntry?.CurrentApprover || ""
  );
  const isRejected = rejectedByName !== "";

  const approvedCount = workflowHistory.filter(
    (x: any) =>
      norm(x.Action) === "approved" || norm(x.ActionTaken) === "approved"
  ).length;

  const initiatorRibbonClass: string = (() => {
    if (isPaid) return "approved";
    if (isSaveAsDraft || isSentBack) return "current";
    if (isPendingVouching || isPendingUTR) return "approved";
    if (isRejected && approvedCount === 0) return "current";
    if (approvedCount > 0) return "approved";
    if (status === "pending for approval" && approverDetails.length > 0) return "approved";
    return "current";
  })();

  const getApproverRibbonClass = (approver: any, index: number): string => {
    if (isPaid) return "approved";
    if (isSaveAsDraft || isSentBack) return "pending";
    if (isPendingVouching || isPendingUTR) return "approved";
    if (norm(approver.Name) === rejectedByName) return "rejected";
    if (isRejected && index >= approvedCount) return "pending";
    if (index < approvedCount) return "approved";
    if (index === approvedCount) return "current";
    return "pending";
  };

  const getAttachments = async (PaymentId: string) => {
    try {
      if (!PaymentId) return;
      const safePaymentId = PaymentId.replace(/\//g, "_");
      const folderPath = `/sites/SonaFinance/InstallationCommision/${safePaymentId}`;
      const files = await sp.web.getFolderByServerRelativePath(folderPath).files();
      setAttachments(files || []);
    } catch (error) {
      console.log("Attachment fetch error:", error);
      setAttachments([]);
    }
  };

  const getVendors = async () => {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName")();
    setVendors(data);
  };

  const getLoggedInUser = async () => {
    try {
      const currentUser = await sp.web.currentUser();
      const email = currentUser.Email;
      const user = await sp.web.lists
        .getByTitle("EmployeeMaster")
        .items.select(
          "EmployeeCode",
          "EmployeeName",
          "Division",
          "Location",
          "EmployeeEmail",
          "ReportingManager/Title",
          "HOD/Title",
          "ContactNo",
          "EmployeeStatus",
          "CostCenter"
        )
        .expand("ReportingManager", "HOD")
        .filter(`EmployeeEmail eq '${email}'`)
        .top(1)();
      if (user.length > 0) setEmployee(user[0]);
    } catch (error) {
      console.log("Error fetching user:", error);
    }
  };

  useEffect(() => {
    void getLoggedInUser();
    void getVendors();
  }, []);

  useEffect(() => {
    if (!formData) return;
    setPONumber(formData.PONumber || "");
    setPOdate(formData.POdate?.split("T")[0] || "");
    setPOPaymentTerms(formData.POPaymentTerms || "");
    setPOAmount(formData.POAmount || "");
    setTotalPaymentofProject(formData.TotalPaymentofProject || "");
    setGSTAdjustmentifAny(formData.GSTAdjustmentifAny || "");
    setOtherAdjustmentifany(formData.OtherAdjustmentifany || "");
    setTotalamounttobeCapitalized(formData.TotalamounttobeCapitalized || "");
    setVendorName(formData.VendorName || "");
    setApproverRemarks(formData.ApproverRemarks || "");
    setVoucherDate(formData.VoucherDate?.split("T")[0] || "");
    setVouchingNumber(formData.VouchingNumber || "");
    setUTRDate(formData.UTRDate?.split("T")[0] || "");
    setUTRNumber(formData.UTRNumber || "");

    const vendor = vendors.find(
      (v) => String(v.VendorCode).trim() === String(formData.VendorCode).trim()
    );
    if (vendor) {
      setSelectedVendorId(vendor.Id);
      setSelectedVendorName(vendor.VendorName);
      setSelectedVendorCode(vendor.VendorCode);
    } else {
      setSelectedVendorId(null);
      setSelectedVendorName(formData.VendorName || "");
      setSelectedVendorCode(formData.VendorCode || "");
    }

    if (formData.PaymentId) void getAttachments(formData.PaymentId);
  }, [formData, vendors]);

  useEffect(() => {
    if (!formData) return;
    try {
      const approvalMatrix =
        typeof formData.ApprovalMatrix === "string"
          ? JSON.parse(formData.ApprovalMatrix)
          : formData.ApprovalMatrix || [];
      setApproverDetails(approvalMatrix);
    } catch (error) {
      console.log("Approval Matrix Parse Error", error);
      setApproverDetails([]);
    }
    try {
      const history =
        typeof formData.WorkFlowHistory === "string"
          ? JSON.parse(formData.WorkFlowHistory)
          : formData.WorkFlowHistory || [];
      setWorkflowHistory(history);
    } catch (error) {
      console.log("Workflow Parse Error", error);
      setWorkflowHistory([]);
    }
  }, [formData]);

  const handleExit = () => {
    if (onClose) onClose();
    else window.location.reload();
  };

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">

            <div className="bordered">
              <img src={logo} alt="Sona Logo" />
              <h1>Advance Payment (View)</h1>
            </div>

            <div className="approval-ribbon">
              <div className={`ribbon-step ${initiatorRibbonClass}`}>
                {employee.EmployeeName || "Initiator"}
              </div>
              {approverDetails.map((approver: any, index: number) => (
                <div
                  key={index}
                  className={`ribbon-step ${getApproverRibbonClass(approver, index)}`}
                >
                  {approver.Name}
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
                    <label className="font">Employee Code</label>
                    &nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{employee.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Name</label>
                    &nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{employee.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Email</label>
                    &nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{employee.EmployeeEmail}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Contact No</label>
                    &nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{employee.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Status</label>
                    &nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{employee.EmployeeStatus}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Division</label>
                    &nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{employee.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Location</label>
                    &nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{employee.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">RM</label>
                    &nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{employee.ReportingManager?.Title}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">HOD</label>
                    &nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{employee.HOD?.Title}</label>
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
                    <select
                      value={selectedVendorId ?? ""}
                      disabled={true}
                      className="formtext-control readonly"
                      onChange={(e) => {
                        const id = Number(e.target.value);
                        const vendor = vendors.find((v) => v.Id === id);
                        setSelectedVendorId(id);
                        setSelectedVendorName(vendor?.VendorName || "");
                      }}
                    >
                      <option value="">Select Vendor</option>
                      {vendors.map((v: IVendor) => (
                        <option key={v.Id} value={v.Id}>
                          {v.VendorCode}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Name</label>
                    <input value={vendorName} readOnly className="form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number</label>
                    <input value={PONumber} readOnly className="form-control readonly" />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date</label>
                    <input type="date" value={POdate} readOnly className="form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Payment Terms</label>
                    <input value={POPaymentTerms} readOnly className="form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount (GST)</label>
                    <input value={POAmount} readOnly className="form-control readonly" />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Total Payment for the Project</label>
                    <input value={TotalPaymentofProject} readOnly className="form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">GST Adjustment (if Any)</label>
                    <input value={GSTAdjustmentifAny} readOnly className="form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Other Adjustment</label>
                    <input value={OtherAdjustmentifany} readOnly className="form-control readonly" />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font" style={{ color: "red" }}>
                      Total Project Amount to be Capitalized
                    </label>
                    <input value={TotalamounttobeCapitalized} readOnly className="form-control readonly" />
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Approver Details</label>
              </div>

              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Voucher Date</label>
                    <input type="date" value={voucherDate} readOnly className="form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Voucher Number</label>
                    <input value={VouchingNumber} readOnly className="form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">UTR Date</label>
                    <input type="date" value={UTRDate} readOnly className="form-control readonly" />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">UTR Number</label>
                    <input value={UTRNumber} readOnly className="form-control readonly" />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-12">
                    <label className="font">Approver Remarks</label>
                    <label className="fonttext textbox readonly" style={{ width: "100%", height: "auto" }}>
                      {approverRemarks}
                    </label>
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
                    {attachments.length > 0 ? (
                      <ul>
                        {attachments.map((file: any, index: number) => (
                          <li key={index}>
                            <a href={file.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">
                              {file.Name}
                            </a>
                          </li>
                        ))}
                      </ul>
                    ) : (
                      <span className="fonttext">No Attachments</span>
                    )}
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
                <a href="#" onClick={handleExit} className="reset-btn">
                  Back
                </a>
              </div>

            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default ViewAdvanceForm;