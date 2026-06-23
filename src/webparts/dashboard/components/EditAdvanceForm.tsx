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

interface IPOData {
  Id: number;
  PONumber: string;
  PODate: string;
  POPaymentTerms: string;
  POAmount: string;
  POBasicAmount: string;
  POGSTAmount: string;
  POOtherAmount: string;
  MRNBasicAmount: string;
  MRNGSTAmount: string;
  MRNOtherAmount: string;
  MRNNumber: string;
  MRNDtae: string;
  MRNAmountwithGST: string;
  RequestedAmountforPayment: string;
  VoucherDate: string;
  VoucherNumber: string;
}

const EditAdvanceForm = ({ context, formData, onClose }: any) => {
  const sp = spfi().using(SPFx(context));

  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [employee, setEmployee] = useState<any>({});
  const [attachments, setAttachments] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [selectedUser, setSelectedUser] = useState<any[]>([]);

  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const [selectedVendorName, setSelectedVendorName] = useState("");

  const [poList, setPoList] = useState<IPOData[]>([]);
  const [poLoading, setPoLoading] = useState(false);

  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [poAmount, setPoAmount] = useState("");

  const [poBasicAmount, setPoBasicAmount] = useState("");
  const [poGSTAmount, setPoGSTAmount] = useState("");
  const [poOtherAmount, setPoOtherAmount] = useState("");

  const [mrnBasicAmount, setMrnBasicAmount] = useState("");
  const [mrnGSTAmount, setMrnGSTAmount] = useState("");
  const [mrnOtherAmount, setMrnOtherAmount] = useState("");

  const [advanceAmount, setAdvanceAmount] = useState("");

  const [gstToBeCapitalized, setGstToBeCapitalized] = useState(false);
  const [showGstTooltip, setShowGstTooltip] = useState(false);

  const [assetCodes, setAssetCodes] = useState<string[]>([""]);

  const [remarks, setRemarks] = useState("");

  const [approverDetails, setApproverDetails] = useState<any[]>([]);
  const [approvers, setApprovers] = useState<number[]>([]);

  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [advanceHistory, setAdvanceHistory] = useState<any[]>([]);

  const poAmountTotal = (
    Number(poBasicAmount || 0) +
    Number(poGSTAmount || 0) +
    Number(poOtherAmount || 0)
  ).toFixed(2);

  const mrnAmountTotal = (
    Number(mrnBasicAmount || 0) +
    Number(mrnGSTAmount || 0) +
    Number(mrnOtherAmount || 0)
  ).toFixed(2);

  const paidAmount = Number(advanceAmount || 0).toFixed(2);

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };

  const handleNumberChange = (value: string, setter: any) => {
    if (/^\d*\.?\d*$/.test(value)) setter(value);
  };

  const handleAssetCodeChange = (index: number, value: string) => {
    const updated = [...assetCodes];
    updated[index] = value;
    setAssetCodes(updated);
  };

  const addAssetCode = () => setAssetCodes([...assetCodes, ""]);

  const removeAssetCode = (index: number) => {
    if (assetCodes.length === 1) return;
    setAssetCodes(assetCodes.filter((_, i) => i !== index));
  };

  const getAssetCodesForSave = (): string => {
    return assetCodes
      .map((c) => c.trim())
      .filter((c) => c !== "")
      .join(", ");
  };

  const clearPOAndMRNFields = () => {
    setPoDate("");
    setPoTerms("");
    setPoAmount("");
    setPoBasicAmount("");
    setPoGSTAmount("");
    setPoOtherAmount("");
    setMrnBasicAmount("");
    setMrnGSTAmount("");
    setMrnOtherAmount("");
  };

  const getPaidPOs = async () => {
    setPoLoading(true);
    try {
      const result = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.select(
          "Id", "PONumber", "PODate", "POPaymentTerms", "POAmount",
          "POBasicAmount", "POGSTAmount", "POOtherAmount",
          "MRNBasicAmount", "MRNGSTAmount", "MRNOtherAmount",
          "MRNNumber", "MRNDtae", "MRNAmountwithGST",
          "RequestedAmountforPayment", "VoucherDate", "VoucherNumber", "Status",
        )
        .filter(`Status eq 'Paid'`)
        .orderBy("Created", false)
        .top(500)();

      const seen = new Set<string>();
      const uniquePOs: IPOData[] = [];

      for (const item of result) {
        if (!item.PONumber || seen.has(item.PONumber)) continue;
        seen.add(item.PONumber);
        uniquePOs.push({
          Id: item.Id,
          PONumber: item.PONumber,
          PODate: item.PODate || "",
          POPaymentTerms: item.POPaymentTerms || "",
          POAmount: item.POAmount || "",
          POBasicAmount: item.POBasicAmount != null ? String(item.POBasicAmount) : "0",
          POGSTAmount: item.POGSTAmount != null ? String(item.POGSTAmount) : "0",
          POOtherAmount: item.POOtherAmount != null ? String(item.POOtherAmount) : "0",
          MRNBasicAmount: item.MRNBasicAmount != null ? String(item.MRNBasicAmount) : "0",
          MRNGSTAmount: item.MRNGSTAmount != null ? String(item.MRNGSTAmount) : "0",
          MRNOtherAmount: item.MRNOtherAmount != null ? String(item.MRNOtherAmount) : "0",
          MRNNumber: item.MRNNumber || "",
          MRNDtae: item.MRNDtae || "",
          MRNAmountwithGST: item.MRNAmountwithGST || "",
          RequestedAmountforPayment: item.RequestedAmountforPayment || "",
          VoucherDate: item.VoucherDate || "",
          VoucherNumber: item.VoucherNumber || "",
        });
      }

      setPoList(uniquePOs);
    } catch (error) {
      console.error("Error fetching PO list from CapexPayment:", error);
      setPoList([]);
    } finally {
      setPoLoading(false);
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

  const getVendors = async () => {
    try {
      const data = await sp.web.lists
        .getByTitle("VendorMaster")
        .items.select("Id", "VendorCode", "VendorName")();
      setVendors(data);
    } catch (error) {
      console.error("Error fetching vendors:", error);
    }
  };

  const getLoggedInUser = async () => {
    try {
      const currentUser = await sp.web.currentUser();
      const email = currentUser.Email;

      const user = await sp.web.lists
        .getByTitle("EmployeeMaster")
        .items.select(
          "EmployeeCode", "EmployeeName", "Division", "Location",
          "EmployeeEmail", "ReportingManager/Id", "ReportingManager/Title",
          "HOD/Id", "HOD/Title", "ContactNo", "EmployeeStatus", "CostCenter",
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

      if (employee.ReportingManager?.Id) {
        baseApprovers.push({
          Id: employee.ReportingManager.Id,
          Name: employee.ReportingManager.Title,
          Role: "RM",
          Level: 1,
          status: "Pending",
        });
      }

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
        .items.select("Role/RoleName", "Approver/Id", "Approver/Title", "Level/Level")
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

  const validateForm = () => {
    const errors: string[] = [];

    if (!selectedVendorId) errors.push("Vendor Name is required");
    if (!poNumber) errors.push("PO Number is required");
    if (!poDate) errors.push("PO Date is required");
    if (!advanceAmount || Number(advanceAmount) <= 0)
      errors.push("Total Amount to be Capitalized must be greater than zero");
    if (getAssetCodesForSave() === "")
      errors.push("At least one Asset Code is required");
    if (
      (!attachments || attachments.length === 0) &&
      (!selectedFiles || selectedFiles.length === 0)
    )
      errors.push("Please upload at least one attachment");

    return errors;
  };

  const handleExit = () => {
    onClose();
  };

  const ensureFolder = async (folderPath: string) => {
    try {
      await sp.web.getFolderByServerRelativePath(folderPath)();
    } catch {
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

  const handleSave = async (status: "Save as Draft" | "Pending for Approval") => {
    try {
      if (status === "Pending for Approval") {
        const errors = validateForm();
        if (errors.length > 0) {
          await Swal.fire({
            title: "Validation",
            html: errors.map((e) => `• ${e}`).join("<br/>"),
            icon: "error",
          });
          return;
        }
      }

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
        status === "Pending for Approval" && flow.length > 0 ? flow[0].Id : null;

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

      const selectedVendor = vendors.find((v) => v.Id === selectedVendorId);

      await sp.web.lists
        .getByTitle("Installation")
        .items.getById(formData.ID)
        .update({
          Status: status,
          ApprovalMatrix: JSON.stringify(flow),
          CurrentApproverId: currentApproverId,
          WorkFlowHistory: JSON.stringify(existingHistory),
          VendorCode: selectedVendor?.VendorCode || "",
          VendorName: selectedVendor?.VendorName || "",
          PONumber: poNumber || "",
          POdate: poDate ? new Date(poDate) : null,
          POPaymentTerms: poTerms || "",
          POAmount: poAmount || "0",
          POAmountBasic: poBasicAmount || "0",
          POAmountGST: poGSTAmount || "0",
          POAmountOther: poOtherAmount || "0",
          POAmountTotal: poAmountTotal || "0",
          MRNAmountBasic: mrnBasicAmount || "0",
          MRNAmountGST: mrnGSTAmount || "0",
          MRNAmountOther: mrnOtherAmount || "0",
          MRNAmountTotal: mrnAmountTotal || "0",
          TotalamounttobeCapitalized: paidAmount || "0",
          GSTToBeCapitalized: gstToBeCapitalized ? "Yes" : "No",
          AssetCodes: getAssetCodesForSave(),
        });

      if (selectedFiles.length > 0) {
        await uploadFiles();
      }

      await Swal.fire({
        title: "Success",
        text: status === "Save as Draft" ? "Draft Saved Successfully" : "Submitted Successfully",
        icon: "success",
      });

      onClose();
    } catch (error) {
      console.error(error);
      await Swal.fire({ title: "Error", text: "Something went wrong.", icon: "error" });
    }
  };

  useEffect(() => {
    if (!formData || vendors.length === 0) return;

    setPoNumber(formData.PONumber || "");
    setPoDate(formData.POdate?.split("T")[0] || "");
    setPoTerms(formData.POPaymentTerms || "");
    setPoAmount(formData.POAmount || "");

    setPoBasicAmount(formData.POAmountBasic != null ? String(formData.POAmountBasic) : "0");
    setPoGSTAmount(formData.POAmountGST != null ? String(formData.POAmountGST) : "0");
    setPoOtherAmount(formData.POAmountOther != null ? String(formData.POAmountOther) : "0");
    setMrnBasicAmount(formData.MRNAmountBasic != null ? String(formData.MRNAmountBasic) : "0");
    setMrnGSTAmount(formData.MRNAmountGST != null ? String(formData.MRNAmountGST) : "0");
    setMrnOtherAmount(formData.MRNAmountOther != null ? String(formData.MRNAmountOther) : "0");
    setAdvanceAmount(formData.TotalamounttobeCapitalized || "");
    setGstToBeCapitalized(formData.GSTToBeCapitalized === "Yes");

    if (formData.AssetCodes) {
      const codes = String(formData.AssetCodes)
        .split(",")
        .map((c: string) => c.trim())
        .filter((c: string) => c !== "");
      setAssetCodes(codes.length > 0 ? codes : [""]);
    } else {
      setAssetCodes([""]);
    }

    const vendor = vendors.find(
      (v) => String(v.VendorCode).trim() === String(formData.VendorCode).trim(),
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

    if (formData.PaymentId) {
      void getAttachments(formData.PaymentId);
    }

    if (formData.PONumber) {
      void getPastMRNDetails(formData.PONumber);
      void getAdvanceHistory(formData.PONumber);
    }
  }, [formData, vendors]);

  useEffect(() => {
    void getLoggedInUser();
    void getVendors();
    void getPaidPOs();
  }, []);

  useEffect(() => {
    if (employee?.EmployeeCode) {
      void buildApprovalFlow();
    }
  }, [employee]);

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            <div className="bordered">
              <img src={logo} />
              <h1> Edit Advance Payment </h1>
            </div>

            <div className="approval-ribbon">
              <div className="ribbon-step current">
                {employee.EmployeeName || "Initiator"}
              </div>
              {approverDetails.map((approver, index) => (
                <div key={index} className={`ribbon-step ${approver.status === "Pending" ? "pending" : "pending"}`}>
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
                    <label htmlFor="Employee Code" className="font">Employee Code</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Name" className="font">Employee Name</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Email" className="font">Employee Email</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.EmployeeEmail}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Contact No" className="font">Contact No</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Status" className="font">Employee Status</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.EmployeeStatus}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Division" className="font">Division</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Location" className="font">Location</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="RM" className="font">RM</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.ReportingManager?.Title}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="HOD" className="font">HOD</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.HOD?.Title}</label>
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Vendor &amp; PO Details</label>
              </div>
              <div className="main-formcontainer">

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Vendor Name</label>
                    <select
                      value={selectedVendorId || ""}
                      className="formtext-control"
                      onChange={(e) => {
                        const id = Number(e.target.value);
                        const vendor = vendors.find((v) => v.Id === id);
                        setSelectedVendorId(id || null);
                        setSelectedVendorCode(vendor?.VendorCode || "");
                        setSelectedVendorName(vendor?.VendorName || "");
                      }}
                    >
                      <option value="">Select Vendor</option>
                      {vendors.map((v) => (
                        <option key={v.Id} value={v.Id}>{v.VendorName}</option>
                      ))}
                    </select>
                  </div>

                  <div className="col-md-4">
                    <label className="font">Vendor Code</label>
                    <input value={selectedVendorCode} className="form-control readonly" readOnly />
                  </div>

                  <div className="col-md-4">
                    <label className="font">PO Number</label>
                    <select
                      value={poNumber}
                      className="formtext-control"
                      disabled={poLoading}
                      onChange={(e) => {
                        const val = e.target.value;
                        const selectedPO = poList.find((item) => item.PONumber === val);
                        setPoNumber(val);
                        if (selectedPO) {
                          setPoDate(selectedPO.PODate ? new Date(selectedPO.PODate).toISOString().split("T")[0] : "");
                          setPoTerms(selectedPO.POPaymentTerms || "");
                          setPoAmount(selectedPO.POAmount || "");
                          setPoBasicAmount(selectedPO.POBasicAmount || "0");
                          setPoGSTAmount(selectedPO.POGSTAmount || "0");
                          setPoOtherAmount(selectedPO.POOtherAmount || "0");
                          setMrnBasicAmount(selectedPO.MRNBasicAmount || "0");
                          setMrnGSTAmount(selectedPO.MRNGSTAmount || "0");
                          setMrnOtherAmount(selectedPO.MRNOtherAmount || "0");
                        } else {
                          clearPOAndMRNFields();
                        }
                        void getPastMRNDetails(val);
                        void getAdvanceHistory(val);
                      }}
                    >
                      <option value="">
                        {poLoading
                          ? "Loading PO Numbers..."
                          : poList.length === 0
                          ? "No Paid POs found"
                          : "Select PO Number"}
                      </option>
                      {poList.map((item) => (
                        <option key={item.Id} value={item.PONumber}>{item.PONumber}</option>
                      ))}
                    </select>
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date</label>
                    <input type="date" value={poDate} className="form-control readonly" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Payment Terms</label>
                    <input value={poTerms} className="form-control readonly" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount (Incl. GST)</label>
                    <input value={poAmount} className="form-control readonly" readOnly />
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Basic Amount</label>
                    <input value={poBasicAmount} className="form-control readonly computed-field" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO GST Amount</label>
                    <input value={poGSTAmount} className="form-control readonly computed-field" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Other Amount</label>
                    <input value={poOtherAmount} className="form-control readonly computed-field" readOnly />
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">MRN Basic Amount</label>
                    <input value={mrnBasicAmount} className="form-control readonly computed-field" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN GST Amount</label>
                    <input value={mrnGSTAmount} className="form-control readonly computed-field" readOnly />
                  </div>
                  <div className="col-md-4">
                    <label className="font">MRN Other Amount</label>
                    <input value={mrnOtherAmount} className="form-control readonly computed-field" readOnly />
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
                    <label className="font" style={{ color: "#000000" }}>Total Amount to be Capitalized</label>
                    <input
                      value={advanceAmount}
                      className="form-control"
                      placeholder="0.00"
                      onChange={(e) => handleNumberChange(e.target.value, setAdvanceAmount)}
                    />
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4" style={{ display: "flex", alignItems: "flex-end", paddingBottom: "4px" }}>
                    <div className="gst-capitalized-row">
                      <input
                        type="checkbox"
                        id="gstCapitalized"
                        checked={gstToBeCapitalized}
                        onChange={(e) => setGstToBeCapitalized(e.target.checked)}
                        className="gst-checkbox"
                      />
                      <label htmlFor="gstCapitalized" className="gst-checkbox-label font">
                        Whether GST to be Capitalized
                      </label>
                      <span
                        className="info-icon"
                        onMouseEnter={() => setShowGstTooltip(true)}
                        onMouseLeave={() => setShowGstTooltip(false)}
                      >
                        &#9432;
                        {showGstTooltip && (
                          <span className="info-tooltip">Info not added yet!</span>
                        )}
                      </span>
                    </div>
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-12">
                    <label className="font">
                      Asset Code(s) <span style={{ color: "red" }}>*</span>
                    </label>
                    <div className="asset-codes-container">
                      {assetCodes.map((code, index) => (
                        <div key={index} className="asset-code-row">
                          <input
                            value={code}
                            className={`form-control asset-code-input${code.trim() === "" ? " input-error" : ""}`}
                            placeholder={`Asset Code ${index + 1}`}
                            onChange={(e) => handleAssetCodeChange(index, e.target.value)}
                          />
                          <button
                            type="button"
                            className="asset-code-remove-btn"
                            onClick={() => removeAssetCode(index)}
                            disabled={assetCodes.length === 1}
                            title="Remove"
                          >
                            &times;
                          </button>
                          {index === assetCodes.length - 1 && (
                            <button
                              type="button"
                              className="asset-code-add-btn"
                              onClick={addAssetCode}
                              title="Add another asset code"
                            >
                              + Add
                            </button>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}><label>Past MRN Details</label></div>
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
                            <tr><td colSpan={9} style={{ textAlign: "center" }}>No Data</td></tr>
                          ) : (
                            previousAdvances.map((item: any, index: number) => (
                              <tr key={index}>
                                <td className="px-4 py-2">{item.PONumber}</td>
                                <td className="px-4 py-2">{item.PODate ? new Date(item.PODate).toLocaleDateString() : ""}</td>
                                <td className="px-4 py-2">{item.POAmount}</td>
                                <td className="px-4 py-2">{item.MRNNumber}</td>
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

              <div className="heading1" style={{ marginTop: "10px" }}><label>Advance History (to be PO Specific)</label></div>
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
                                  <td className="px-4 py-2"></td>
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
                <label>Upload Docuement</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attachments</label>
                    {attachments.length > 0 && (
                      <ul>
                        {attachments.map((file: any, index: number) => (
                          <li key={index} style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                            <a href={file.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">
                              {file.Name}
                            </a>
                            <span
                              style={{ color: "red", cursor: "pointer", fontWeight: "bold", fontSize: "16px" }}
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

              <div style={{ display: "flex", justifyContent: "center", gap: "10px", marginBottom: "1rem", marginTop: "1rem" }}>
                <button type="button" className="draft-btn" onClick={() => handleSave("Save as Draft")}>
                  Save as Draft
                </button>
                <button type="button" className="submit-btn" onClick={() => handleSave("Pending for Approval")}>
                  Submit
                </button>
                <button type="button" className="reset-btn" onClick={handleExit}>
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