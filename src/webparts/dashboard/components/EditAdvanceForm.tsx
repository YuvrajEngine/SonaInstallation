import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";

import logo from "../assets/sona-comstarlogo.png";

interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

const EditAdvanceForm = ({ context, formData, onClose }: any) => {
  const sp = spfi().using(SPFx(context));

  // =========================
  // STATES
  // =========================
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [employee, setEmployee] = useState<any>({});
  const [attachments, setAttachments] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [selectedUser, setSelectedUser] = useState<any[]>([]);
  const [PONumber, setPONumber] = useState("");
  const [TotalPaymentofProject, setTotalPaymentofProject] = useState("");
  const [GSTAdjustmentifAny, setGSTAdjustmentifAny] = useState("");
  const [OtherAdjustmentifany, setOtherAdjustmentifany] = useState("");
  const [TotalamounttobeCapitalized, setTotalamounttobeCapitalized] =
    useState("");
  const [POdate, setPOdate] = useState("");
  const [POPaymentTerms, setPOPaymentTerms] = useState("");
  const [POAmount, setPOAmount] = useState("");
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");

  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [poAmount, setPoAmount] = useState("");
  const [advanceAmount, setAdvanceAmount] = useState("");
  const [paidAmount, setPaidAmount] = useState("");
  const [expectedDate, setExpectedDate] = useState("");
  const [vendorName, setVendorName] = useState("");
  const [glCode, setGlCode] = useState("");
  const [costCenter, setCostCenter] = useState("");
  const [remarks, setRemarks] = useState("");
  const [projectDesc, setProjectDesc] = useState("");

  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [voucherNumber, setVoucherNumber] = useState("");
  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };
  const totalCapitalized =
    Number(TotalPaymentofProject || 0) +
    Number(GSTAdjustmentifAny || 0) +
    Number(OtherAdjustmentifany || 0);

  // =========================
  // LOAD DATA
  // =========================
  const handleNumberChange = (value: string, setter: any) => {
    // Allow only numbers and decimal (max one dot)
    const regex = /^\d*\.?\d*$/;

    if (regex.test(value)) {
      void setter(value);
    }
  };
  const getVendors = async () => {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName")();
    void setVendors(data);
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
          "CostCenter",
        )
        .expand("ReportingManager", "HOD")
        .filter(`EmployeeEmail eq '${email}'`)
        .top(1)();

      if (user.length > 0) {
        setEmployee(user[0]);
      }
    } catch (error) {
      console.log("Error fetching user:", error);
    }
  };

  const getAttachments = async (PaymentId: string) => {
    try {
      const safe = PaymentId.replace(/\//g, "_");
      const path = `/sites/SonaFinance/InstallationCommision/${safe}`;

      const files = await sp.web.getFolderByServerRelativePath(path).files();

      void setAttachments(files);
    } catch {
      void setAttachments([]);
    }
  };

  // =========================
  // UPLOAD FILES
  // =========================

  // =========================
  // VALIDATION
  // =========================

  const validateForm = () => {
    const errors: string[] = [];

    if (!selectedVendorId) {
      errors.push("Please select the Vendor code");
    }

    if (!poNumber) {
      errors.push("Please update PO Number");
    }

    if (!paidAmount || Number(paidAmount) <= 0) {
      errors.push("Please update Total Payment ");
    }

    // 🔥 NEW VALIDATION

    if (
      (!attachments || attachments.length === 0) &&
      (!selectedFiles || selectedFiles.length === 0)
    ) {
      errors.push("Please upload at least one attachment");
    }

    return errors;
  };
  const validateForm1 = () => {
    const errors: string[] = [];

    if (!selectedVendorId) {
      errors.push("Please select the Vendor code");
    }

    if (!poNumber) {
      errors.push("Please update PO Number");
    }

    if (!poDate) {
      errors.push("Please update PO date");
    }

    if (!poTerms) {
      errors.push("Please update PO Terms");
    }

    if (!poAmount) {
      errors.push("Please update PO Amount");
    }

    if (!advanceAmount || Number(advanceAmount) <= 0) {
      errors.push("Please update Capatalized Amount");
    }

    if (!paidAmount || Number(paidAmount) <= 0) {
      errors.push("Please update Paid Amount");
    }

    // 🔥 NEW VALIDATION
    if (
      advanceAmount &&
      paidAmount &&
      Number(paidAmount) > Number(advanceAmount)
    ) {
      errors.push("Paid Amount cannot be greater than Capatalized Amount");
    }

    if (!expectedDate) {
      errors.push("Please update Settlement Date");
    }

    if (!selectedUser || selectedUser.length === 0) {
      errors.push("Please select PIC Name");
    }

    if (!remarks) {
      errors.push("Please enter Remarks");
    }

    if (
      (!attachments || attachments.length === 0) &&
      (!selectedFiles || selectedFiles.length === 0)
    ) {
      errors.push("Please upload at least one attachment");
    }

    return errors;
  };

  const handleExit = () => {
    if (onClose) {
      onClose();
    } else {
      window.location.reload();
    }
  };
  const ensureFolder = async (folderPath: string) => {
    try {
      await sp.web.getFolderByServerRelativePath(folderPath)();
    } catch {
      // create folder if not exists
      const parentPath = folderPath.substring(0, folderPath.lastIndexOf("/"));
      const folderName = folderPath.substring(folderPath.lastIndexOf("/") + 1);

      await sp.web
        .getFolderByServerRelativePath(parentPath)
        .folders.addUsingPath(folderName);
    }
  };
  const uploadFiles = async () => {
    try {
      if (!formData?.PaymentId || selectedFiles.length === 0) return;

      const safe = formData.PaymentId.replace(/\//g, "_");
      const folderPath = `/sites/SonaFinance/InstallationCommision/${safe}`;

      // ✅ Ensure folder exists
      await ensureFolder(folderPath);

      for (const file of selectedFiles) {
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, file, { Overwrite: true });
      }

      setSelectedFiles([]);
      await getAttachments(formData.PaymentId);
    } catch (error) {
      console.error("Upload error:", error);
      alert("File upload failed ❌");
    }
  };

  // =========================
  // UPDATE
  // =========================

  const handleSubmit = async () => {
    try {
      const errors = validateForm();
      if (errors.length > 0) {
        alert(errors.join("\n"));
        return;
      }

      const ensuredUser = await sp.web.ensureUser(
        selectedUser[0]?.secondaryText,
      );

      await sp.web.lists
        .getByTitle("Installation")
        .items.getById(formData.ID)
        .update({
          Title: formData.PaymentId,
          PaymentId: formData.PaymentId,

          EmployeeCode: employee.EmployeeCode || "",
          EmployeeName: employee.EmployeeName || "",
          Division: employee.Division || "",
          Location: employee.Location || "",
          Email: employee.EmployeeEmail || "",

          ReportingManager: employee.ReportingManager?.Title || "",
          HOD: employee.HOD?.Title || "",
          ContactNo: employee.ContactNo || "",
          EmployeeStatus: employee.EmployeeStatus || "",

          VendorCode: selectedVendorId ? selectedVendorId.toString() : "",
          VendorName: selectedVendorName || "",

          PONumber: PONumber || "",
          POdate: POdate ? new Date(POdate) : null,
          POPaymentTerms: POPaymentTerms || "",
          POAmount: POAmount || "",

          TotalPaymentofProject: TotalPaymentofProject || "0",
          GSTAdjustmentifAny: GSTAdjustmentifAny || "0",
          OtherAdjustmentifany: OtherAdjustmentifany || "0",
          TotalamounttobeCapitalized: totalCapitalized.toString(),

          Status: "Pending for Approver",
        });


      if (selectedFiles.length > 0) {
        await uploadFiles();
      }

      alert("Updated successfully ✅");
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/Commision.aspx?page=User";
      void handleExit();
    } catch (e) {
      console.error(e);
      alert("Error ❌");
    }
  };

  // =========================
  // BIND DATA
  // =========================
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
    setSelectedVendorId(formData.VendorCodeId || null); // ✅ ADD THIS
    setSelectedVendorName(formData.VendorName || ""); // ✅ ADD THIS

    setApproverRemarks(formData.ApproverRemarks || "");
    setVoucherDate(formData.VoucherDate?.split("T")[0] || "");
    //setVouchingNumber(formData.VouchingNumber || "");
    setUTRDate(formData.UTRDate?.split("T")[0] || "");
    setUTRNumber(formData.UTRNumber || "");

    // ✅ PIC FIX
    if (formData?.PICName?.Title) {
      setSelectedUser([
        {
          text: formData.PICName.Title,
          secondaryText: formData.PICName.EMail,
        },
      ]);
    }

    if (formData.PaymentId) {
      void getAttachments(formData.PaymentId);
    }
  }, [formData]);

  useEffect(() => {
    void getLoggedInUser();
    void getVendors();
  }, []);

  // =========================
  // UI
  // =========================
  return (
    <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
      <div className='row'>
        <div className='col-md-12'>
          <div className='Main-Boxpoup'>
            {/* 🔹 Header */}
            <div className="bordered">
              <img src={logo} />
              <h1> Edit Advance Payment </h1>
            </div>
            <div className='borderedbox'>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Requestor Information</label>
              </div>
              <div className='main-formcontainer'>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Code" className='font'>Employee Code</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.EmployeeCode}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Name" className='font'>Employee Name </label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.EmployeeName}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Email" className='font'>Employee Email </label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.EmployeeEmail}</label>
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Contact No" className='font'>Contact No</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.ContactNo}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Employee Status" className='font'>Employee Status</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.EmployeeStatus}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="Division" className='font'>Division</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.Division}</label>
                  </div>
                </div>
                <div className='row mb-20'>
                  <div className='col-md-4'>
                    <label htmlFor="Location" className='font'>Location</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.Location}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="RM" className='font'>RM</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.ReportingManager?.Title}</label>
                  </div>
                  <div className='col-md-4'>
                    <label htmlFor="HOD" className='font'>HOD</label> : &nbsp;&nbsp;
                    <label className='fonttext'>  {employee.HOD?.Title}</label>
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
                      value={selectedVendorId || ""}
                      className="formtext-control"
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
                    <input value={selectedVendorName || vendorName} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>PO Number</label>
                    <input value={PONumber} className='form-control' onChange={(e) => setPONumber(e.target.value)} />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className='font'>PO Date</label>
                    <input type="date" value={POdate} className='form-control' onChange={(e) => setPOdate(e.target.value)} />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>PO Payment Terms</label>
                    <input value={POPaymentTerms} className='form-control' onChange={(e) => setPOPaymentTerms(e.target.value)} />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>PO Amount </label>
                    <input value={POAmount} className='form-control' onChange={(e) => setPOAmount(e.target.value)} />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Total Payment for the Project</label>
                    <input value={TotalPaymentofProject} className='form-control' onChange={(e) => setTotalPaymentofProject(e.target.value)} />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Gst Adjustment(Any)</label>
                    <input value={GSTAdjustmentifAny} className='form-control' onChange={(e) => setGSTAdjustmentifAny(e.target.value)} />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Other Adjustment</label>
                    <input value={OtherAdjustmentifany} className='form-control' onChange={(e) => setOtherAdjustmentifany(e.target.value)} />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font" style={{ color: "red" }}>Total Project Amount to be Capitalized</label>
                    <input value={totalCapitalized} className='form-control readonly' />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Previous Advances</label>
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
                            <th className="px-4 py-2">Requested Date</th>
                            <th className="px-4 py-2">Paid Date</th>
                            <th className="px-4 py-2">MRN No</th>
                            <th className="px-4 py-2">Settled Amount</th>
                            <th className="px-4 py-2">Pending Advance</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td className="px-4 py-2" colSpan={7} style={{ textAlign: "center" }}>
                              No Data
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Upload Docuement</label>
              </div>
              <div className='main-formcontainer'>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attachments</label>
                    {attachments.length > 0 && (
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
                    <input type="file" className="form-control" multiple onChange={(e) => { if (e.target.files) { setSelectedFiles(Array.from(e.target.files)); } }} />
                  </div>
                </div>
              </div>
              <div style={{ display: "flex", justifyContent: "center", gap: "5px", marginBottom: "1rem", marginTop: "1rem" }}>
                <a onClick={handleSubmit} className="submit-btn">
                  Update
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
export default EditAdvanceForm;
