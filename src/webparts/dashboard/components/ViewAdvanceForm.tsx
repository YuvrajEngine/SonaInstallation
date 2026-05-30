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
interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

import logo from "../assets/sona-comstarlogo.png";

const ViewAdvanceForm = ({ context, formData, onClose }: any) => {
  const [attachments, setAttachments] = useState<any[]>([]);
  const sp = spfi().using(SPFx(context));
  const [employee, setEmployee] = useState<any>({});
  // 🔹 Employee//TotalPaymentofProject
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [selectedUser, setSelectedUser] = useState<any[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");

  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [employeeName, setEmployeeName] = useState("");
  const [employeeEmail, setEmployeeEmail] = useState("");
  const [VendorCode, setVendorCode] = useState<number | null>(null);
  // 🔹 Form fields//TotalPaymentofProject//PONumber
  const [PONumber, setPONumber] = useState("");
  const [TotalPaymentofProject, setTotalPaymentofProject] = useState("");
  const [GSTAdjustmentifAny, setGSTAdjustmentifAny] = useState("");
  const [OtherAdjustmentifany, setOtherAdjustmentifany] = useState("");
  const [TotalamounttobeCapitalized, setTotalamounttobeCapitalized] = useState("");
  const [POdate, setPOdate] = useState("");
  const [POPaymentTerms, setPOPaymentTerms] = useState("");
  const [POAmount, setPOAmount] = useState("");
  const [advanceAmount, setAdvanceAmount] = useState("");
  // const [TotalamounttobeCapitalized, setPaidAmount] = useState("");
  const [expectedDate, setExpectedDate] = useState("");
  const [vendorName, setVendorName] = useState("");
  const [glCode, setGlCode] = useState("");
  const [costCenter, setCostCenter] = useState("");
  const [remarks, setRemarks] = useState("");
  const [projectDesc, setProjectDesc] = useState("");

  // 🔹 Extra fields
  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [VouchingNumber, setVouchingNumber] = useState("");
  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };
  const getAttachments = async (PaymentId: string) => {
    try {
      if (!PaymentId) return;

      const safePaymentId = PaymentId.replace(/\//g, "_");

      const folderPath = `/sites/SonaFinance/InstallationCommision/${safePaymentId}`;

      console.log("Folder Path:", folderPath); // ✅ DEBUG

      const files = await sp.web
        .getFolderByServerRelativePath(folderPath)
        .files();

      console.log("Files:", files); // ✅ DEBUG

      setAttachments(files || []);
    } catch (error) {
      console.log("Attachment fetch error:", error);
      setAttachments([]);
    }
  };

  const uploadFiles = async () => {
    if (!formData?.PaymentId || selectedFiles.length === 0) return;

    const safe = formData.PaymentId.replace(/\//g, "_");
    const path = `/sites/SonaFinance/CapexAdvanceDocs/${safe}`;

    for (const file of selectedFiles) {
      await sp.web
        .getFolderByServerRelativePath(path)
        .files.addUsingPath(file.name, file, { Overwrite: true });
    }

    void setSelectedFiles([]);
    void getAttachments(formData.PaymentId);
  };

  // ✅ Fetch Item by ID
  const getVendors = async () => {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName")();
    void setVendors(data);
  };

  // ✅ Bind SharePoint Data
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
    setVouchingNumber(formData.VouchingNumber || "");
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


  const handleExit = () => {
    if (onClose) onClose();
    else window.location.reload();
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
  useEffect(() => {
    void getLoggedInUser();
    void getVendors();
  }, []);
  return (
    <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
      <div className='row'>
        <div className='col-md-12'>
          <div className='Main-Boxpoup'>
            {/* 🔹 Header */}
            <div className="bordered">
              <img src={logo} />
              <h1> Advance Payment (View) </h1>
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
                    <input value={vendorName} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>PO Number</label>
                    <input value={PONumber} className='form-control readonly' />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className='font'>PO Date</label>
                    <input type="date" value={POdate} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>PO Payment Terms</label>
                    <input value={POPaymentTerms} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>PO Amount (GST)</label>
                    <input value={POAmount} className='form-control readonly' />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Total Payment for the Project</label>
                    <input value={TotalPaymentofProject} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Gst Adjustment(Any)</label>
                    <input value={GSTAdjustmentifAny} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Other Adjustment</label>
                    <input value={OtherAdjustmentifany} className='form-control readonly' />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font" style={{ color: "red" }}>Total Project Amount to be Capitalized</label>
                    <input value={TotalamounttobeCapitalized} className='form-control readonly' />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Approver Details</label>
              </div>
              <div className='main-formcontainer'>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className='font'>Voucher Date</label>
                    <input type="date" value={voucherDate} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>Voucher Number</label>
                    <input value={VouchingNumber} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>UTR Date</label>
                    <input type="date" value={UTRDate} className='form-control readonly' />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className='font'>UTR Number</label>
                    <input value={UTRNumber} className='form-control readonly' />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-12">
                    <label className='font'>Approver Remarks</label>
                    <label className="fonttext textbox readonly" style={{ width: "100%", height: "auto" }}>{approverRemarks}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Upload Docuement</label>
              </div>
              <div className='main-formcontainer'>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className='font'>Attachments</label>
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
                  </div>
                </div>
              </div>
              <div style={{ display: "flex", justifyContent: "center", gap: "5px", marginBottom: "1rem", marginTop: "1rem" }}>
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
