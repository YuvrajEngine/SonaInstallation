import * as React from "react";
import "./advanced.scss";
import Swal from "sweetalert2";
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

interface IPOData {
  Id: number;
  PONumber: string;
  PODate: string;
  POPaymentTerms: string;
  POAmount: string;
  MRNNumber: string;
  MRNDtae: string;
  MRNAmountwithGST: string;
  RequestedAmountforPayment: string;
  VoucherDate: string;
  VoucherNumber: string;
}

const NewAdvanceform = ({ context, onClose }: any) => {
  const sp = spfi().using(SPFx(context));
  const [employee, setEmployee] = React.useState<any>({});
  const [attachments, setAttachments] = useState<File[]>([]);
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [advanceHistory, setAdvanceHistory] = useState<any[]>([]);
  const [data, setData] = React.useState<any[]>([]);
  const [employeeName, setEmployeeName] = React.useState("");
  const [pickerKey, setPickerKey] = React.useState<number>(0);
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
  const [selectedUser, setSelectedUser] = useState<any[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [poList, setPoList] = useState<IPOData[]>([]);
  const [poAmount, setPoAmount] = useState("");

  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");

  // New PO Amount fields
  const [poAmountBasic, setPoAmountBasic] = useState("");
  const [poAmountGST, setPoAmountGST] = useState("");
  const [poAmountOther, setPoAmountOther] = useState("");

  // MRN Amount Basic
  const [mrnAmountBasic, setMrnAmountBasic] = useState("");

  // Asset codes (multiple)
  const [assetCodes, setAssetCodes] = useState<string[]>([""]);

  // GST capitalization tickbox
  const [gstToBeCapitalized, setGstToBeCapitalized] = useState(false);
  const [showGstTooltip, setShowGstTooltip] = useState(false);

  const [expectedDate, setExpectedDate] = useState("");
  const [glCode, setGlCode] = useState("FIN-001");
  const [costCenter, setCostCenter] = useState("");
  const [remarks, setRemarks] = useState("");
  const [projectDesc, setProjectDesc] = useState("");
  const [advanceAmount, setAdvanceAmount] = useState("");
  const [approverDetails, setApproverDetails] = useState<any[]>([]);
  const [approvers, setApprovers] = useState<number[]>([]);

  // PO Amount Total computed
  const poAmountTotal = (
    Number(poAmountBasic || 0) +
    Number(poAmountGST || 0) +
    Number(poAmountOther || 0)
  ).toFixed(2);

  const paidAmount = Number(advanceAmount || 0).toFixed(2);

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };

  const handleNumberChange = (value: string, setter: any) => {
    const regex = /^\d*\.?\d*$/;
    if (regex.test(value)) {
      void setter(value);
    }
  };

  // Asset code handlers
  const handleAssetCodeChange = (index: number, value: string) => {
    const updated = [...assetCodes];
    updated[index] = value;
    setAssetCodes(updated);
  };

  const addAssetCode = () => {
    setAssetCodes([...assetCodes, ""]);
  };

  const removeAssetCode = (index: number) => {
    if (assetCodes.length === 1) return;
    const updated = assetCodes.filter((_, i) => i !== index);
    setAssetCodes(updated);
  };

  // Fetch PO list from CapexPayment where Status = 'Paid'
  const getPaidPOs = async () => {
    try {
      const result = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.select(
          "Id",
          "PONumber",
          "PODate",
          "POPaymentTerms",
          "POAmount",
          "MRNNumber",
          "MRNDtae",
          "MRNAmountwithGST",
          "RequestedAmountforPayment",
          "VoucherDate",
          "VoucherNumber",
          "Status",
        )
        .filter(`Status eq 'Paid'`)
        .orderBy("Created", false)
        .top(500)();

      const uniquePOs = result.filter(
        (item: any, index: number, self: any[]) =>
          item.PONumber &&
          index === self.findIndex((x: any) => x.PONumber === item.PONumber),
      );

      setPoList(uniquePOs);
    } catch (error) {
      console.log("Error fetching PO list from CapexPayment:", error);
      setPoList([]);
    }
  };

  // Fetch Past MRN Details from CapexPayment based on PONumber
  const getPastMRNDetails = async (selectedPONumber: string) => {
    try {
      if (!selectedPONumber) {
        setPreviousAdvances([]);
        return;
      }
      const result = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.select(
          "PONumber",
          "PODate",
          "POAmount",
          "MRNNumber",
          "MRNDtae",
          "MRNAmountwithGST",
          "RequestedAmountforPayment",
          "VoucherDate",
          "VoucherNumber",
          "Status",
        )
        .filter(`PONumber eq '${selectedPONumber}' and Status eq 'Paid'`)
        .orderBy("Created", false)();

      setPreviousAdvances(result);
    } catch (error) {
      console.error("Error fetching Past MRN Details:", error);
      setPreviousAdvances([]);
    }
  };

  // Fetch Advance History from CapexAdvance based on PONumber
  const getAdvanceHistory = async (selectedPONumber: string) => {
    try {
      if (!selectedPONumber) {
        setAdvanceHistory([]);
        return;
      }
      const result = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.select(
          "PONumber",
          "RequestAdvanceAmount",
          "Created",
          "VoucherDate",
          "PaidAmount",
          "VouchingNumber",
          "Status",
        )
        .filter(`PONumber eq '${selectedPONumber}'`)
        .orderBy("Created", false)();

      setAdvanceHistory(result);
    } catch (error) {
      console.error("Error fetching Advance History:", error);
      setAdvanceHistory([]);
    }
  };

  //  GET LIST DATA
  const getCapexData = async () => {
    debugger;
    try {
      const items = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.select(
          "ID",
          "Title",
          "Created",
          "EmployeeName",
          "VendorName",
          "VendorCode/Id",
          "VendorCode/VendorCode",
          "PONumber",
          "RequestAdvanceAmount",
          "Status",
        )
        .expand("VendorCode")
        .orderBy("ID", false)();

      const formatted = items.map((item: any) => ({
        ID: item.ID,
        id: item.Title,
        date: item.Created
          ? new Date(item.Created).toLocaleDateString("en-GB")
          : "",
        EmployeeName: item.EmployeeName,
        vendor: item.VendorName || "",
        vendorCode: item.VendorCode?.VendorCode || "",
        po: item.PONumber || "",
        amount: item.RequestAdvanceAmount || 0,
        status: item.Status || "",
      }));

      setData(formatted);
    } catch (error) {
      console.error("Data error:", error);
    }
  };

  const handleExit = async () => {
    const result = await Swal.fire({
      title: "Exit Form?",
      text: "Any unsaved changes will be lost.",
      icon: "warning",
      showCancelButton: true,
      confirmButtonText: "Yes, Exit",
      cancelButtonText: "Cancel",
    });

    if (result.isConfirmed) {
      onClose();
    }
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
        void setEmployee(user[0]);
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

  const getVendors = async () => {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName")();
    void setVendors(data);
  };

  const getFinancialYear = () => {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;
    if (month >= 4) {
      return (year + 1).toString().slice(-2);
    }
    return year.toString().slice(-2);
  };

  const generatePaymentId = async (): Promise<string> => {
    try {
      const fy = getFinancialYear();

      const incrementalData = await sp.web.lists
        .getByTitle("InstallationIncrementalID")
        .items.select("Id", "IncrementalID")
        .orderBy("Id", false)
        .top(1)();

      let nextNumber = 1;

      if (incrementalData.length > 0 && incrementalData[0].IncrementalID) {
        nextNumber = Number(incrementalData[0].IncrementalID) + 1;
      }

      const paddedNumber = nextNumber.toString().padStart(5, "0");

      await sp.web.lists.getByTitle("InstallationIncrementalID").items.add({
        Title: `INT-${nextNumber}`,
        IncrementalID: nextNumber.toString(),
      });

      return `INT/${fy}/${paddedNumber}`;
    } catch (error) {
      console.error("Generate Payment ID Error:", error);
      return `INT/${getFinancialYear()}/00001`;
    }
  };

  const uploadAttachments = async (PaymentId: string) => {
    try {
      if (!attachments || attachments.length === 0) return;

      const safePaymentId = PaymentId.replace(/\//g, "_");
      const libraryName = "InstallationCommision";
      const webUrl = context.pageContext.web.serverRelativeUrl;
      const folderPath = `${webUrl}/${libraryName}/${safePaymentId}`;

      await sp.web.folders.addUsingPath(`${libraryName}/${safePaymentId}`);

      for (const file of attachments) {
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, file, { Overwrite: true });
      }

      console.log("Files uploaded successfully");
    } catch (error) {
      console.error("Upload error:", error);
    }
  };

  const validateForm = (): string[] => {
    const errors: string[] = [];

    if (!selectedVendorId) {
      errors.push("Vendor Code is required");
    }

    if (!selectedVendorName) {
      errors.push("Vendor Name is required");
    }

    if (!poNumber) {
      errors.push("PO Number is required");
    }

    if (!poDate) {
      errors.push("PO Date is required");
    }

    if (!advanceAmount || Number(advanceAmount) <= 0) {
      errors.push("Total Amount to be Capitalized must be greater than zero");
    }

    // Asset code: at least one non-empty value required
    const filledAssetCodes = assetCodes.filter((c) => c.trim() !== "");
    if (filledAssetCodes.length === 0) {
      errors.push("At least one Asset Code is required");
    }

    if (!attachments || attachments.length === 0) {
      errors.push("Please upload at least one attachment");
    }

    return errors;
  };

  const buildItemPayload = (
    PaymentId: string,
    flow: any[],
    status: string,
    actionTaken: string,
  ) => {
    const history = [
      {
        CurrentApprover: employee.EmployeeName,
        ActionTaken: actionTaken,
        Comment: remarks || actionTaken,
        Date: new Date().toISOString(),
      },
    ];

    const currentApproverId = flow.length > 0 ? flow[0].Id : null;
    const selectedVendor = vendors.find((v) => v.Id === selectedVendorId);
    const filledAssetCodes = assetCodes.filter((c) => c.trim() !== "");

    return {
      Title: PaymentId,
      PaymentId: PaymentId,
      EmployeeCode: employee.EmployeeCode || "",
      EmployeeName: employee.EmployeeName || "",
      Division: employee.Division || "",
      Location: employee.Location || "",
      Email: employee.EmployeeEmail || "",
      ReportingManager: employee.ReportingManager?.Title || "",
      HOD: employee.HOD?.Title || "",
      ContactNo: employee.ContactNo || "",
      EmployeeStatus: employee.EmployeeStatus || "",
      VendorCode: selectedVendor?.VendorCode || "",
      VendorName: selectedVendor?.VendorName || "",
      PONumber: poNumber || "",
      POdate: poDate ? new Date(poDate) : null,
      POPaymentTerms: poTerms || "",
      POAmount: poAmountTotal || "0",
      // New fields — add these columns to the Installation list (see notes below)
      MRNAmountBasic: mrnAmountBasic || "0",
      POAmountBasic: poAmountBasic || "0",
      POAmountGST: poAmountGST || "0",
      POAmountOther: poAmountOther || "0",
      POAmountTotal: poAmountTotal || "0",
      AssetCodes: filledAssetCodes.join(", "),
      GSTToBeCapitalized: gstToBeCapitalized ? "Yes" : "No",
      TotalPaymentofProject: advanceAmount || "0",
      GSTAdjustmentifAny: "0",
      OtherAdjustmentifany: "0",
      TotalamounttobeCapitalized: paidAmount || "0",
      Status: status,
      ApprovalMatrix: JSON.stringify(flow),
      CurrentApproverId: currentApproverId,
      WorkFlowHistory: JSON.stringify(history),
    };
  };

  const handleSubmit = async () => {
    try {
      const errors = validateForm();

      if (errors.length > 0) {
        await Swal.fire({
          title: "Validation",
          html: errors.map((e) => `• ${e}`).join("<br/>"),
          icon: "error",
        });
        return;
      }

      const confirmSubmit = await Swal.fire({
        title: "Submit Request?",
        text: "Are you sure you want to submit this Installation Commissioning Request?",
        icon: "question",
        showCancelButton: true,
        confirmButtonText: "Submit",
        cancelButtonText: "Cancel",
      });

      if (!confirmSubmit.isConfirmed) return;

      Swal.fire({
        title: "Submitting...",
        text: "Please wait",
        allowOutsideClick: false,
        didOpen: () => {
          Swal.showLoading();
        },
      });

      const PaymentId = await generatePaymentId();
      const flow = await buildApprovalFlow();
      const payload = buildItemPayload(
        PaymentId,
        flow,
        "Pending for Approval",
        "Submitted",
      );

      await sp.web.lists.getByTitle("Installation").items.add(payload);
      await uploadAttachments(PaymentId);

      await Swal.fire({
        title: "Success",
        text: "Request submitted successfully.",
        icon: "success",
        confirmButtonText: "OK",
      });

      onClose();
    } catch (error: any) {
      console.error(error);
      await Swal.fire({
        title: "Submission Failed",
        text: error?.message || "Something went wrong",
        icon: "error",
      });
    }
  };

  const handledraft = async () => {
    try {
      const confirmDraft = await Swal.fire({
        title: "Save as Draft?",
        text: "Do you want to save this request as Draft?",
        icon: "question",
        showCancelButton: true,
        confirmButtonText: "Save Draft",
        cancelButtonText: "Cancel",
      });

      if (!confirmDraft.isConfirmed) return;

      Swal.fire({
        title: "Saving Draft...",
        text: "Please wait",
        allowOutsideClick: false,
        didOpen: () => {
          Swal.showLoading();
        },
      });

      const PaymentId = await generatePaymentId();
      const flow = await buildApprovalFlow();
      const payload = buildItemPayload(
        PaymentId,
        flow,
        "Save as Draft",
        "Saved as draft",
      );

      await sp.web.lists.getByTitle("Installation").items.add(payload);

      if (attachments.length > 0) {
        await uploadAttachments(PaymentId);
      }

      await Swal.fire({
        title: "Draft Saved",
        text: "Request saved successfully as Draft.",
        icon: "success",
        confirmButtonText: "OK",
      });

      onClose();
    } catch (error: any) {
      console.error(error);
      await Swal.fire({
        title: "Draft Save Failed",
        text: error?.message || "Something went wrong",
        icon: "error",
      });
    }
  };

  React.useEffect(() => {
    debugger;
    if (!context) return;
    debugger;
    void getLoggedInUser();
    void getVendors();
    void getPaidPOs();
  }, [context]);

  React.useEffect(() => {
    if (employee?.EmployeeCode) {
      void buildApprovalFlow();
    }
  }, [employee]);

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            {/* 🔹 Header */}
            <div className="bordered">
              <img src={logo} />
              <h1> Installation Commisioning Request </h1>
            </div>
            {/* Approval Ribbon */}
            <div className="approval-ribbon">
              {/* Initiator */}
              <div className="ribbon-step current">
                {employee.EmployeeName || "Initiator"}
              </div>

              {/* Dynamic Approvers */}
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
                    <label className="fonttext">{employee.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Name" className="font">
                      Employee Name{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">{employee.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Email" className="font">
                      Employee Email{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">{employee.EmployeeEmail}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Contact No" className="font">
                      Contact No
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">{employee.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Status" className="font">
                      Employee Status
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {employee.EmployeeStatus}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Division" className="font">
                      Division
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">{employee.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Location" className="font">
                      Location
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">{employee.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="RM" className="font">
                      RM
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {employee.ReportingManager?.Title}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="HOD" className="font">
                      HOD
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">{employee.HOD?.Title}</label>
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Vendor &amp; PO Details</label>
              </div>
              <div className="main-formcontainer">
                {/* Row 1: Vendor Code, Vendor Name, PO Number */}
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Vendor Code</label>
                    <select
                      value={selectedVendorId || ""}
                      onChange={(e) => {
                        const id = Number(e.target.value);
                        const vendor = vendors.find((v) => v.Id === id);
                        setSelectedVendorId(id);
                        setSelectedVendorCode(vendor?.VendorCode || "");
                        setSelectedVendorName(vendor?.VendorName || "");
                      }}
                      className="formtext-control"
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
                      value={selectedVendorName}
                      className="form-control readonly"
                      readOnly
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number</label>
                    <select
                      value={poNumber}
                      className="formtext-control"
                      onChange={(e) => {
                        const val = e.target.value;
                        const selectedPO = poList.find(
                          (item) => item.PONumber === val,
                        );
                        setPoNumber(val);
                        if (selectedPO) {
                          setPoDate(
                            selectedPO.PODate
                              ? new Date(selectedPO.PODate)
                                  .toISOString()
                                  .split("T")[0]
                              : "",
                          );
                          setPoTerms(selectedPO.POPaymentTerms || "");
                          setPoAmount(selectedPO.POAmount || "");
                        } else {
                          setPoDate("");
                          setPoTerms("");
                          setPoAmount("");
                        }
                        void getPastMRNDetails(val);
                        void getAdvanceHistory(val);
                      }}
                    >
                      <option value="">Select PO Number</option>
                      {poList.map((item) => (
                        <option key={item.Id} value={item.PONumber}>
                          {item.PONumber}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date</label>
                    <input
                      type="date"
                      value={poDate}
                      className="form-control readonly"
                      readOnly
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Payment Terms</label>
                    <input
                      value={poTerms}
                      className="form-control readonly"
                      readOnly
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount (GST)</label>
                    <input
                      value={poAmount}
                      className="form-control readonly"
                      readOnly
                    />
                  </div>
                </div>

                {/* Row 3: MRN Amount Basic, PO Amount Basic, PO Amount GST */}
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">MRN Amount Basic</label>
                    <input
                      value={mrnAmountBasic}
                      className="form-control"
                      placeholder="0.00"
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setMrnAmountBasic)
                      }
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount Basic</label>
                    <input
                      value={poAmountBasic}
                      className="form-control"
                      placeholder="0.00"
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setPoAmountBasic)
                      }
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount GST</label>
                    <input
                      value={poAmountGST}
                      className="form-control"
                      placeholder="0.00"
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setPoAmountGST)
                      }
                    />
                  </div>
                </div>

                {/* Row 4: PO Amount Other, PO Amount Total */}
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Amount Other</label>
                    <input
                      value={poAmountOther}
                      className="form-control"
                      placeholder="0.00"
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setPoAmountOther)
                      }
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font" style={{ color: "#000000" }}>
                      PO Amount Total (Basic + GST + Other)
                    </label>
                    <input
                      value={poAmountTotal}
                      className="form-control readonly computed-field"
                      readOnly
                    />
                  </div>
                </div>

                {/* Row 5: Total Amount to be Capitalized + GST tickbox */}
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font" style={{ color: "#000000" }}>
                      Total Amount to be Capitalized
                    </label>
                    <input
                      value={advanceAmount}
                      className="form-control"
                      placeholder="0.00"
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setAdvanceAmount)
                      }
                    />
                  </div>
                  <div
                    className="col-md-4"
                    style={{
                      display: "flex",
                      alignItems: "flex-end",
                      paddingBottom: "4px",
                    }}
                  >
                    <div className="gst-capitalized-row">
                      <input
                        type="checkbox"
                        id="gstCapitalized"
                        checked={gstToBeCapitalized}
                        onChange={(e) =>
                          setGstToBeCapitalized(e.target.checked)
                        }
                        className="gst-checkbox"
                      />
                      <label
                        htmlFor="gstCapitalized"
                        className="gst-checkbox-label font"
                      >
                        Whether GST to be Capitalized
                      </label>
                      <span
                        className="info-icon"
                        onMouseEnter={() => setShowGstTooltip(true)}
                        onMouseLeave={() => setShowGstTooltip(false)}
                      >
                        &#9432;
                        {showGstTooltip && (
                          <span className="info-tooltip">
                            Info not added yet!
                          </span>
                        )}
                      </span>
                    </div>
                  </div>
                </div>

                {/* Asset Codes — multiple entry, no empty save */}
                <div className="row mb-20">
                  <div className="col-md-12">
                    <label className="font">
                      Asset Code(s){" "}
                      <span style={{ color: "red" }}>*</span>
                    </label>
                    <div className="asset-codes-container">
                      {assetCodes.map((code, index) => (
                        <div key={index} className="asset-code-row">
                          <input
                            value={code}
                            className={`form-control asset-code-input${code.trim() === "" ? " input-error" : ""}`}
                            placeholder={`Asset Code ${index + 1}`}
                            onChange={(e) =>
                              handleAssetCodeChange(index, e.target.value)
                            }
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

              {/* Past MRN Details — original columns kept, data from CapexPayment */}
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
                            <th className="px-4 py-2">Po Amount</th>
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
                            <tr>
                              <td colSpan={9} style={{ textAlign: "center" }}>
                                No Data
                              </td>
                            </tr>
                          ) : (
                            previousAdvances.map((item: any, index: number) => (
                              <tr key={index}>
                                {/* PO Number — CapexPayment: PONumber */}
                                <td className="px-4 py-2">{item.PONumber}</td>
                                {/* PO Date — CapexPayment: PODate */}
                                <td className="px-4 py-2">
                                  {item.PODate
                                    ? new Date(item.PODate).toLocaleDateString()
                                    : ""}
                                </td>
                                {/* PO Amount — CapexPayment: POAmount */}
                                <td className="px-4 py-2">{item.POAmount}</td>
                                {/* MRN No — CapexPayment: MRNNumber */}
                                <td className="px-4 py-2">{item.MRNNumber}</td>
                                {/* MRN Date — CapexPayment: MRNDtae */}
                                <td className="px-4 py-2">
                                  {item.MRNDtae
                                    ? new Date(
                                        item.MRNDtae,
                                      ).toLocaleDateString()
                                    : ""}
                                </td>
                                {/* MRN Amount — CapexPayment: MRNAmountwithGST */}
                                <td className="px-4 py-2">
                                  {item.MRNAmountwithGST}
                                </td>
                                {/* Advance Adjustment — not in list, show empty */}
                                <td className="px-4 py-2"></td>
                                {/* Paid Amount — CapexPayment: RequestedAmountforPayment */}
                                <td className="px-4 py-2">
                                  {item.RequestedAmountforPayment}
                                </td>
                                {/* Remarks — not in list, show empty */}
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

              {/* Advance History — original columns kept, data from CapexAdvance */}
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
                            <tr>
                              <td colSpan={7} style={{ textAlign: "center" }}>
                                No Data
                              </td>
                            </tr>
                          ) : (
                            advanceHistory.map((item: any, index: number) => {
                              const pending = Math.max(
                                0,
                                Number(item.RequestAdvanceAmount || 0) -
                                  Number(item.PaidAmount || 0),
                              );
                              return (
                                <tr key={index}>
                                  {/* PO Number — CapexAdvance: PONumber */}
                                  <td className="px-4 py-2">{item.PONumber}</td>
                                  {/* Previous Advance — CapexAdvance: RequestAdvanceAmount */}
                                  <td className="px-4 py-2">
                                    {item.RequestAdvanceAmount}
                                  </td>
                                  {/* Amount Requested Date — CapexAdvance: Created */}
                                  <td className="px-4 py-2">
                                    {item.Created
                                      ? new Date(
                                          item.Created,
                                        ).toLocaleDateString()
                                      : ""}
                                  </td>
                                  {/* Amount Paid Date — CapexAdvance: VoucherDate */}
                                  <td className="px-4 py-2">
                                    {item.VoucherDate
                                      ? new Date(
                                          item.VoucherDate,
                                        ).toLocaleDateString()
                                      : ""}
                                  </td>
                                  {/* MRN No — not in CapexAdvance, show empty */}
                                  <td className="px-4 py-2"></td>
                                  {/* Settled Amount — CapexAdvance: PaidAmount */}
                                  <td className="px-4 py-2">
                                    {item.PaidAmount}
                                  </td>
                                  {/* Pending Advance — computed */}
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
                    <label className="font">Attachment</label>
                    <input
                      type="file"
                      className="form-control"
                      multiple
                      onChange={(e) => {
                        if (e.target.files) {
                          setAttachments(Array.from(e.target.files));
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
                  gap: "5px",
                  marginBottom: "1rem",
                  marginTop: "1rem",
                }}
              >
                <button
                  type="button"
                  onClick={() => void handleSubmit()}
                  className="submit-btn"
                >
                  Submit
                </button>

                <button
                  type="button"
                  onClick={() => void handledraft()}
                  className="Rework-btn"
                >
                  Save as Draft
                </button>

                <button
                  type="button"
                  onClick={() => void handleExit()}
                  className="reset-btn"
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

export default NewAdvanceform;