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
import Swal from "sweetalert2";

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
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
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
  //Approval Flow
  const [approverDetails, setApproverDetails] = useState<any[]>([]);
  const [approvers, setApprovers] = useState<number[]>([]);

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
          "ReportingManager/Id",
          "ReportingManager/Title",
          "HOD/Id",
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

  const buildApprovalFlow = async () => {
    try {
      const baseApprovers: any[] = [];

      // RM
      if (employee.ReportingManager?.Id) {
        baseApprovers.push({
          Id: employee.ReportingManager.Id,
          Name: employee.ReportingManager.Title,
          Role: "RM",
          Level: 1,
          status: "Pending",
        });
      }

      // HOD
      if (employee.HOD?.Id) {
        baseApprovers.push({
          Id: employee.HOD.Id,
          Name: employee.HOD.Title,
          Role: "HOD",
          Level: 2,
          status: "",
        });
      }

      const matrixData = await sp.web.lists
        .getByTitle("InstallationCommisionApprovalMatrix")
        .items.select(
          "Role/RoleName",
          "Approver/Id",
          "Approver/Title",
          "Level/Level",
        )
        .expand("Role", "Approver", "Level")
        .filter("Status eq 'Active'")
        .orderBy("Level", true)();

      const matrixApprovers = matrixData.map((item: any, index: number) => ({
        Id: item.Approver?.Id,
        Name: item.Approver?.Title,
        Role: item.Role?.RoleName,
        Level: baseApprovers.length + index + 1,
        status: "",
      }));

      const fullFlow = [...baseApprovers, ...matrixApprovers];

      if (fullFlow.length > 0) {
        fullFlow[0].status = "Pending";
      }

      setApproverDetails(fullFlow);
      setApprovers(fullFlow.map((a) => a.Id));

      return fullFlow;
    } catch (error) {
      console.error("Approval Flow Error:", error);
      return [];
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

  const deleteAttachment = async (fileName: string) => {
    const result = await Swal.fire({
      title: "Delete Attachment?",
      text: `Are you sure you want to delete "${fileName}"?`,
      icon: "warning",
      showCancelButton: true,
      confirmButtonText: "Yes, Delete",
      cancelButtonText: "Cancel",
      confirmButtonColor: "#d33",
    });

    if (!result.isConfirmed) return;

    try {
      const safe = formData.PaymentId.replace(/\//g, "_");

      const filePath = `/sites/SonaFinance/InstallationCommision/${safe}/${fileName}`;

      await sp.web.getFileByServerRelativePath(filePath).delete();

      await Swal.fire({
        title: "Deleted",
        text: "Attachment deleted successfully.",
        icon: "success",
        timer: 1500,
        showConfirmButton: false,
      });

      await getAttachments(formData.PaymentId);
    } catch (error) {
      console.error("Delete Error", error);

      await Swal.fire({
        title: "Error",
        text: "Unable to delete attachment.",
        icon: "error",
      });
    }
  };
  // UPLOAD FILES

  // VALIDATION

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
    onClose();
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

const handleSave = async (
  status: "Save as Draft" | "Pending for Approval",
) => {
  try {
    const confirmation =
      status === "Save as Draft"
        ? await Swal.fire({
            title: "Save as Draft?",
            text: "Do you want to save this request as Draft?",
            icon: "question",
            showCancelButton: true,
            confirmButtonText: "Save Draft",
            cancelButtonText: "Cancel",
          })
        : await Swal.fire({
            title: "Submit Request?",
            text: "Do you want to submit this request for approval?",
            icon: "question",
            showCancelButton: true,
            confirmButtonText: "Submit",
            cancelButtonText: "Cancel",
          });

    if (!confirmation.isConfirmed) return;

    const flow = await buildApprovalFlow();

    const currentApproverId =
      status === "Pending for Approval" && flow.length > 0
        ? flow[0].Id
        : null;

    const existingHistoryRaw = formData?.WorkFlowHistory || "[]";
    let existingHistory: any[] = [];
    try {
      existingHistory = JSON.parse(existingHistoryRaw);
      if (!Array.isArray(existingHistory)) existingHistory = [];
    } catch {
      existingHistory = [];
    }

    existingHistory.push({
      CurrentApprover: employee.EmployeeName || "",
      ActionTaken: status === "Save as Draft" ? "Saved as Draft" : "Submitted",
      Comment: status === "Save as Draft" ? "Saved as draft" : "Request submitted",
      Date: new Date().toISOString(),
    });

    await sp.web.lists
      .getByTitle("Installation")
      .items.getById(formData.ID)
      .update({
        Status: status,
        ApprovalMatrix: JSON.stringify(flow),
        CurrentApproverId: currentApproverId,
        WorkFlowHistory: JSON.stringify(existingHistory), 
      });

    if (selectedFiles.length > 0) {
      await uploadFiles();
    }

    await Swal.fire({
      title: "Success",
      text:
        status === "Save as Draft"
          ? "Draft Saved Successfully"
          : "Submitted Successfully",
      icon: "success",
    });

    onClose();
  } catch (error) {
    console.error(error);

    await Swal.fire({
      title: "Error",
      text: "Something went wrong.",
      icon: "error",
    });
  }
};

  // BIND DATA
  useEffect(() => {
    if (!formData || vendors.length === 0) return;

    setPONumber(formData.PONumber || "");
    setPOdate(formData.POdate?.split("T")[0] || "");
    setPOPaymentTerms(formData.POPaymentTerms || "");
    setPOAmount(formData.POAmount || "");
    setTotalPaymentofProject(formData.TotalPaymentofProject || "");
    setGSTAdjustmentifAny(formData.GSTAdjustmentifAny || "");
    setOtherAdjustmentifany(formData.OtherAdjustmentifany || "");
    setTotalamounttobeCapitalized(formData.TotalamounttobeCapitalized || "");

    setVendorName(formData.VendorName || "");

    console.log("Saved VendorCode:", formData.VendorCode);
    console.log("Vendor Master:", vendors);

    // Match saved VendorCode with VendorMaster
    const vendor = vendors.find(
      (v) => String(v.VendorCode).trim() === String(formData.VendorCode).trim(),
    );

    console.log("Matched Vendor:", vendor);

    if (vendor) {
      setSelectedVendorId(vendor.Id);
      setSelectedVendorName(vendor.VendorName);
      setSelectedVendorCode(vendor.VendorCode);
    } else {
      setSelectedVendorId(null);
      setSelectedVendorName(formData.VendorName || "");
      setSelectedVendorCode(formData.VendorCode || "");
    }

    setApproverRemarks(formData.ApproverRemarks || "");
    setVoucherDate(formData.VoucherDate?.split("T")[0] || "");
    setUTRDate(formData.UTRDate?.split("T")[0] || "");
    setUTRNumber(formData.UTRNumber || "");

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
  }, [formData, vendors]);

  useEffect(() => {
    void getLoggedInUser();
    void getVendors();
  }, []);

  useEffect(() => {
    if (employee?.EmployeeCode) {
      void buildApprovalFlow();
    }
  }, [employee]);

  // UI
  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            {/* 🔹 Header */}
            <div className="bordered">
              <img src={logo} />
              <h1> Edit Advance Payment </h1>
            </div>
            {/* Approval Ribbon */}
            <div className="approval-ribbon">
              <div className="ribbon-step current">
                {employee.EmployeeName || "Initiator"}
              </div>

              {approverDetails.map((approver, index) => (
                <div
                  key={index}
                  className={`ribbon-step ${
                    approver.status === "Pending" ? "pending" : "pending"
                  }`}
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
                    <label htmlFor="Employee Code" className="font">
                      Employee Code
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Name" className="font">
                      Employee Name{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Email" className="font">
                      Employee Email{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {employee.EmployeeEmail}
                    </label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Contact No" className="font">
                      Contact No
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Status" className="font">
                      Employee Status
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {employee.EmployeeStatus}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Division" className="font">
                      Division
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Location" className="font">
                      Location
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="RM" className="font">
                      RM
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {employee.ReportingManager?.Title}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="HOD" className="font">
                      HOD
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.HOD?.Title}</label>
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
                    <select
                      value={selectedVendorId || ""}
                      className="formtext-control"
                      onChange={(e) => {
                        const id = Number(e.target.value);

                        const vendor = vendors.find((v) => v.Id === id);

                        setSelectedVendorId(id);
                        setSelectedVendorName(vendor?.VendorName || "");
                        setSelectedVendorCode(vendor?.VendorCode || "");
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
                    <label className="font">Vendor Name</label>
                    <input
                      value={selectedVendorName || vendorName}
                      className="form-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number</label>
                    <input
                      value={PONumber}
                      className="form-control"
                      onChange={(e) => setPONumber(e.target.value)}
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date</label>
                    <input
                      type="date"
                      value={POdate}
                      className="form-control"
                      onChange={(e) => setPOdate(e.target.value)}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Payment Terms</label>
                    <input
                      value={POPaymentTerms}
                      className="form-control"
                      onChange={(e) => setPOPaymentTerms(e.target.value)}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount </label>
                    <input
                      value={POAmount}
                      className="form-control"
                      onChange={(e) => setPOAmount(e.target.value)}
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">
                      Total Payment for the Project
                    </label>
                    <input
                      value={TotalPaymentofProject}
                      className="form-control"
                      onChange={(e) => setTotalPaymentofProject(e.target.value)}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Gst Adjustment(Any)</label>
                    <input
                      value={GSTAdjustmentifAny}
                      className="form-control"
                      onChange={(e) => setGSTAdjustmentifAny(e.target.value)}
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Other Adjustment</label>
                    <input
                      value={OtherAdjustmentifany}
                      className="form-control"
                      onChange={(e) => setOtherAdjustmentifany(e.target.value)}
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font" style={{ color: "red" }}>
                      Total Project Amount to be Capitalized
                    </label>
                    <input
                      value={totalCapitalized}
                      className="form-control readonly"
                    />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Previous Advances</label>
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
                            <th className="px-4 py-2">Requested Date</th>
                            <th className="px-4 py-2">Paid Date</th>
                            <th className="px-4 py-2">MRN No</th>
                            <th className="px-4 py-2">Settled Amount</th>
                            <th className="px-4 py-2">Pending Advance</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td
                              className="px-4 py-2"
                              colSpan={7}
                              style={{ textAlign: "center" }}
                            >
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
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attachments</label>
                    {attachments.length > 0 && (
                      <ul>
                        {attachments.map((file: any, index: number) => (
                          <li
                            key={index}
                            style={{
                              display: "flex",
                              alignItems: "center",
                              gap: "10px",
                            }}
                          >
                            <a
                              href={file.ServerRelativeUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                            >
                              {file.Name}
                            </a>

                            <span
                              style={{
                                color: "red",
                                cursor: "pointer",
                                fontWeight: "bold",
                                fontSize: "16px",
                              }}
                              onClick={() => deleteAttachment(file.Name)}
                            >
                              ✖
                            </span>
                          </li>
                        ))}
                      </ul>
                    )}
                    <input
                      type="file"
                      className="form-control"
                      multiple
                      onChange={(e) => {
                        if (e.target.files) {
                          setSelectedFiles(Array.from(e.target.files));
                        }
                      }}
                    />
                  </div>
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  gap: "10px",
                  marginBottom: "1rem",
                  marginTop: "1rem",
                }}
              >
                <button
                  type="button"
                  className="draft-btn"
                  onClick={() => handleSave("Save as Draft")}
                >
                  Save as Draft
                </button>

                <button
                  type="button"
                  className="submit-btn"
                  onClick={() => handleSave("Pending for Approval")}
                >
                  Submit
                </button>

                <button
                  type="button"
                  className="reset-btn"
                  onClick={handleExit}
                >
                  Exit
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
export default EditAdvanceForm;
