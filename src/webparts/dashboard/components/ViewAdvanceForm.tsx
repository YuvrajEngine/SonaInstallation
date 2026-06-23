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
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [approverDetails, setApproverDetails] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);

  const [PONumber, setPONumber] = useState("");
  const [POdate, setPOdate] = useState("");
  const [POPaymentTerms, setPOPaymentTerms] = useState("");
  const [POAmount, setPOAmount] = useState("");
  const [vendorName, setVendorName] = useState("");

  const [poAmountBasic, setPoAmountBasic] = useState("");
  const [poAmountGST, setPoAmountGST] = useState("");
  const [poAmountOther, setPoAmountOther] = useState("");
  const [mrnAmountBasic, setMrnAmountBasic] = useState("");
  const [mrnAmountGST, setMrnAmountGST] = useState("");
  const [mrnAmountOther, setMrnAmountOther] = useState("");

  const [assetCodes, setAssetCodes] = useState("");
  const [gstToBeCapitalized, setGstToBeCapitalized] = useState("");

  const [TotalamounttobeCapitalized, setTotalamounttobeCapitalized] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [VoucherNumber, setVoucherNumber] = useState("");
  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");

  const poAmountTotal = (
    Number(poAmountBasic || 0) +
    Number(poAmountGST || 0) +
    Number(poAmountOther || 0)
  ).toFixed(2);

  const mrnAmountTotal = (
    Number(mrnAmountBasic || 0) +
    Number(mrnAmountGST || 0) +
    Number(mrnAmountOther || 0)
  ).toFixed(2);

  const norm = (s: string) => (s || "").toLowerCase().trim();
  const status = norm(formData?.Status);

  const isPaid = status === "paid";
  const isSaveAsDraft = status === "save as draft";
  const isSentBack =
    status === "send back" ||
    workflowHistory.some(
      (x: any) =>
        norm(x.Action) === "sent back" || norm(x.ActionTaken) === "sent back",
    );
  const isPendingVouching = status === "pending for vouching update";
  const isPendingUTR = status === "pending for utr update";

  const rejectedEntry = workflowHistory.find(
    (w: any) =>
      norm(w.Action) === "reject" ||
      norm(w.ActionTaken) === "reject" ||
      norm(w.Action) === "rejected" ||
      norm(w.ActionTaken) === "rejected",
  );
  const rejectedByName = norm(rejectedEntry?.ActionBy || rejectedEntry?.CurrentApprover || "");
  const isRejected = rejectedByName !== "";

  const approvedCount = workflowHistory.filter(
    (x: any) =>
      norm(x.Action) === "approved" || norm(x.ActionTaken) === "approved",
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
    if (norm(approver.Name) === rejectedByName) return "rejected";
    if (isRejected && index >= approvedCount) return "pending";
    if (isPendingUTR) {
      const performerIndex = approverDetails.findIndex(
        (x: any) => x.Role && x.Role.toLowerCase().trim() === "performer",
      );
      if (performerIndex !== -1) {
        if (index === performerIndex) return "current";
        if (index < performerIndex) return "approved";
        return "pending";
      }
    }
    if (isPendingVouching) return "approved";
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

  useEffect(() => {
    void getVendors();
  }, []);

  useEffect(() => {
    if (!formData) return;
    setPONumber(formData.PONumber || "");
    setPOdate(formData.POdate?.split("T")[0] || "");
    setPOPaymentTerms(formData.POPaymentTerms || "");
    setPOAmount(formData.POAmount || "");
    setTotalamounttobeCapitalized(formData.TotalamounttobeCapitalized || "");
    setVendorName(formData.VendorName || "");
    setVoucherDate(formData.VoucherDate?.split("T")[0] || "");
    setVoucherNumber(formData.VoucherNumber || "");
    setUTRDate(formData.UTRDate?.split("T")[0] || "");
    setUTRNumber(formData.UTRNumber || "");

    setPoAmountBasic(formData.POAmountBasic != null ? String(formData.POAmountBasic) : "0");
    setPoAmountGST(formData.POAmountGST != null ? String(formData.POAmountGST) : "0");
    setPoAmountOther(formData.POAmountOther != null ? String(formData.POAmountOther) : "0");
    setMrnAmountBasic(formData.MRNAmountBasic != null ? String(formData.MRNAmountBasic) : "0");
    setMrnAmountGST(formData.MRNAmountGST != null ? String(formData.MRNAmountGST) : "0");
    setMrnAmountOther(formData.MRNAmountOther != null ? String(formData.MRNAmountOther) : "0");
    setGstToBeCapitalized(formData.GSTToBeCapitalized || "");
    setAssetCodes(formData.AssetCodes || "");

    const vendor = vendors.find(
      (v) => String(v.VendorCode).trim() === String(formData.VendorCode).trim(),
    );
    setSelectedVendorCode(vendor ? vendor.VendorCode : formData.VendorCode || "");

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
    } catch {
      setApproverDetails([]);
    }
    try {
      const history =
        typeof formData.WorkFlowHistory === "string"
          ? JSON.parse(formData.WorkFlowHistory)
          : formData.WorkFlowHistory || [];
      setWorkflowHistory(history);
    } catch {
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
              <h1>Installation Commissioning Request (View)</h1>
            </div>

            <div className="approval-ribbon">
              <div className={`ribbon-step ${initiatorRibbonClass}`}>
                {formData?.EmployeeName || "Initiator"}
              </div>
              {approverDetails.map((approver: any, index: number) => (
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
                    <label className="font">Employee Code</label>&nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{formData?.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Name</label>&nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{formData?.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Email</label>&nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{formData?.Email}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Contact No</label>&nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{formData?.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Employee Status</label>&nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{formData?.EmployeeStatus}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Division</label>&nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{formData?.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Location</label>&nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{formData?.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">RM</label>&nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{formData?.ReportingManager}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">HOD</label>&nbsp;:&nbsp;&nbsp;
                    <label className="fonttext">{formData?.HOD}</label>
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
                    <input value={selectedVendorCode} readOnly className="form-control readonly" />
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
                    <label className="font">PO Amount (Incl. GST)</label>
                    <input value={POAmount} readOnly className="form-control readonly" />
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Basic Amount</label>
                    <input value={poAmountBasic} readOnly className="form-control readonly computed-field" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO GST Amount</label>
                    <input value={poAmountGST} readOnly className="form-control readonly computed-field" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Other Amount</label>
                    <input value={poAmountOther} readOnly className="form-control readonly computed-field" />
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">MRN Basic Amount</label>
                    <input value={mrnAmountBasic} readOnly className="form-control readonly computed-field" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN GST Amount</label>
                    <input value={mrnAmountGST} readOnly className="form-control readonly computed-field" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Other Amount</label>
                    <input value={mrnAmountOther} readOnly className="form-control readonly computed-field" />
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Amount Total</label>
                    <input value={poAmountTotal} readOnly className="form-control readonly computed-field" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Amount Total</label>
                    <input value={mrnAmountTotal} readOnly className="form-control readonly computed-field" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Total Amount to be Capitalized</label>
                    <input value={TotalamounttobeCapitalized} readOnly className="form-control readonly" />
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Whether GST to be Capitalized</label>
                    <input value={gstToBeCapitalized} readOnly className="form-control readonly" />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Asset Code(s)</label>
                    <input value={assetCodes} readOnly className="form-control readonly" />
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
                    <input value={VoucherNumber} readOnly className="form-control readonly" />
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

              <div style={{ display: "flex", justifyContent: "center", gap: "5px", marginBottom: "1rem", marginTop: "1rem" }}>
                <a href="#" onClick={handleExit} className="reset-btn">Back</a>
              </div>

            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default ViewAdvanceForm;