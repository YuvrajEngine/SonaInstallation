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
// interface IVendor {
//     VendorCode: string;
//     VendorName: string;
// }


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
  POAdvanceTerms: string;
  POAmtGST: string;
}

const NewAdvanceform = ({ context }: any) => {
  const sp = spfi().using(SPFx(context));
  const [employee, setEmployee] = React.useState<any>({});
  //const [selectedUser, setSelectedUser] = useState<any>(null);
  const [attachments, setAttachments] = useState<File[]>([]);
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [data, setData] = React.useState<any[]>([]);
  const [employeeName, setEmployeeName] = React.useState("");
  const [pickerKey, setPickerKey] = React.useState<number>(0);
  const [vendors, setVendors] = useState<IVendor[]>([]);

  const [selectedUser, setSelectedUser] = useState<any[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");
  const [poList, setPoList] = useState<IPOData[]>([]);
  const [poAmount, setPoAmount] = useState("");

  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");

  const [expectedDate, setExpectedDate] = useState("");

  const [glCode, setGlCode] = useState("FIN-001");

  const [costCenter, setCostCenter] = useState("");

  const [remarks, setRemarks] = useState("");
  const [projectDesc, setProjectDesc] = useState("");
  const [advanceAmount, setAdvanceAmount] = useState("");
  const [gstAdjustment, setGstAdjustment] = useState("");
  const [otherAdjustment, setOtherAdjustment] = useState("");

  // Auto calculate Total Project Amount to be Capitalized
  const paidAmount = (
    Number(advanceAmount || 0) +
    Number(gstAdjustment || 0) +
    Number(otherAdjustment || 0)
  ).toFixed(2);
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };

  const handleNumberChange = (value: string, setter: any) => {
    // Allow only numbers and decimal (max one dot)
    const regex = /^\d*\.?\d*$/;

    if (regex.test(value)) {
      void setter(value);
    }
  };
  const getPaidPOs = async () => {
    try {
      debugger;

      const data = await sp.web.lists
        .getByTitle("CapexAdvance")
        .items.select(
          "Id",
          "PONumber",
          "PODate",
          "POAdvanceTerms",
          "POAmtGST",
          "Status",
        )
        .filter(`Status eq 'Paid'`)
        .orderBy("Created", false)
        .top(500)();

      const uniquePOs = data.filter(
        (item, index, self) =>
          item.PONumber &&
          index === self.findIndex((x) => x.PONumber === item.PONumber),
      );

      setPoList(uniquePOs);
    } catch (error) {
      console.log("Error fetching PO list:", error);
      setPoList([]);
    }
  };

  // ✅ GET LIST DATA
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
          "VendorCode/VendorCode", // 👈 IMPORTANT
          "PONumber",
          "RequestAdvanceAmount",
          "Status",
        )
        .expand("VendorCode") // 👈 MUST
        .orderBy("ID", false)();

      const formatted = items.map((item: any) => ({
        ID: item.ID,
        id: item.Title,
        date: item.Created
          ? new Date(item.Created).toLocaleDateString("en-GB")
          : "",
        EmployeeName: item.EmployeeName,

        vendor: item.VendorName || "",
        vendorCode: item.VendorCode?.VendorCode || "", // 👈 FIX

        po: item.PONumber || "",
        amount: item.RequestAdvanceAmount || 0,
        status: item.Status || "",
      }));

      setData(formatted);
    } catch (error) {
      console.error("Data error:", error);
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

  const handleExit = () => {
    window.location.href =
      "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/Commision.aspx?page=User";
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
        void setEmployee(user[0]);
      }
    } catch (error) {
      console.log("Error fetching user:", error);
    }
  };
  const buildApprovalFlow = async () => {
    try {
      const flow: any[] = [];

      // =========================
      // 🔥 1. ADD RM
      // =========================
      if (employee.ReportingManager?.Title) {
        const rmUser = await sp.web.ensureUser(
          employee.ReportingManager.Title
        );

        flow.push({
          Id: rmUser.Id,
          Name: employee.ReportingManager.Title,
          Role: "RM",
          Level: 1,
          Status: "In Progress"
        });
      }

      // =========================
      // 🔥 2. ADD HOD
      // =========================
      if (employee.HOD?.Title) {
        const hodUser = await sp.web.ensureUser(
          employee.HOD.Title
        );

        flow.push({
          Id: hodUser.Id,
          Name: employee.HOD.Title,
          Role: "HOD",
          Level: 2,
          Status: "Pending"
        });
      }

      // =========================
      // 🔥 3. GET MATRIX APPROVERS
      // =========================
      const matrixData = await sp.web.lists
        .getByTitle("InstallationCommisionApprovalMatrix")
        .items
        .select("Role/RoleName,Approver/Id,Approver/Title,Level/Level")
        .expand("Role", "Approver", "Level")
        .filter("Status eq 'Active'")
        .orderBy("Level", true)();

      const matrixApprovers = matrixData.map(
        (item: any, index: number) => ({
          Id: item.Approver?.Id,
          Name: item.Approver?.Title,
          Role: item.Role?.RoleName,
          Level: item.length + index + 1,
          Status: "Pending"
        })
      );

      const fullFlow = [...flow, ...matrixApprovers];

      return fullFlow;

    } catch (error) {
      console.error("Approval Flow Error:", error);
      return [];
    }
  };
  const getVendors = async () => {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName")(); // ✅ Id required

    void setVendors(data);
  };

  const getFinancialYear = () => {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;

    if (month >= 4) {
      // April to March
      return `${year.toString().slice(-2)}-${(year + 1).toString().slice(-2)}`;
    } else {
      return `${(year - 1).toString().slice(-2)}-${year.toString().slice(-2)}`;
    }
  };
  const generatePaymentId = async () => {
    try {
      const fy = getFinancialYear(); // 25-26

      const items = await sp.web.lists
        .getByTitle("Installation")
        .items.select("PaymentId", "ID")
        .filter(`startswith(PaymentId,'INT/${fy}/')`)
        .orderBy("ID", false)
        .top(1)();

      let nextNumber = 1;

      if (items.length > 0 && items[0].PaymentId) {
        const lastId = items[0].PaymentId;

        const parts = lastId.split("/");

        if (parts.length === 3) {
          const lastNumber = parseInt(parts[2]);

          if (!isNaN(lastNumber)) {
            nextNumber = lastNumber + 1;
          }
        }
      }

      const formattedNumber = nextNumber.toString().padStart(5, "0");

      return `INT/${fy}/${formattedNumber}`;
    } catch (error) {
      alert(error);
      console.error("Capex ID Error:", error);
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

      // ✅ Ensure folder
      await sp.web.folders.addUsingPath(`${libraryName}/${safePaymentId}`);

      // ✅ Upload files properly
      for (const file of attachments) {
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, file, { Overwrite: true });
      }

      console.log("✅ Files uploaded successfully");
    } catch (error) {
      console.error("❌ Upload error:", error);
    }
  };

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

    if (!attachments || attachments.length === 0) {
      errors.push("Please upload at least one attachment");
    }

    return errors;
  };

  const handleSubmit = async () => {
    try {
      debugger;
      const errors = validateForm();

      if (errors.length > 0) {
        alert(errors.join("\n")); // 👈 shows exactly like your screenshot
        return;
      }

      const PaymentId = await generatePaymentId();
      const fields = await sp.web.lists
        .getByTitle("Installation")
        .fields.filter("InternalName eq 'PaymentId'")();

      console.log("Fields:", fields);
      //   alert("Fields: " + JSON.stringify(fields));
      const flow = await buildApprovalFlow();
      const currentApproverId = flow.length > 0 ? flow[0].Id : null;
      const history = [
        {
          CurrentApprover: employee.EmployeeName,
          ActionTaken: "Submitted",
          Comment: remarks || "Request submitted",
          Date: new Date().toISOString()
        }
      ];

      await sp.web.lists.getByTitle("Installation").items.add({
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

        VendorCode: selectedVendorId ? selectedVendorId.toString() : "",
        VendorName: selectedVendorName || "",

        PONumber: poNumber || "",
        POdate: poDate ? new Date(poDate) : null,
        POPaymentTerms: poTerms || "",
        POAmount: poAmount || "",

        TotalPaymentofProject: advanceAmount || "0",
        GSTAdjustmentifAny: gstAdjustment || "0",
        OtherAdjustmentifany: otherAdjustment || "0",
        TotalamounttobeCapitalized: paidAmount || "0",

        //VoucherDate: voucherDate ? new Date(voucherDate) : null,
        //UTRDate: UTRDate ? new Date(UTRDate) : null,

        Status: "Pending for Approver",
        ApprovalMatrix: JSON.stringify(flow),
        CurrentApproverId: currentApproverId,
        WorkFlowHistory: JSON.stringify(history)
      });

      debugger;
      await uploadAttachments(PaymentId); // 🔥 FIXED

      console.log("Attachments:", attachments);
      alert("Submitted successfully ✅");

      // 🔥 REDIRECT
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/Commision.aspx?page=User";
    } catch (error) {
      console.error("ERROR:", error);
      alert(error);
    }
  };

  const handledraft = async () => {
    try {
      debugger;
      const PaymentId = await generatePaymentId();

      // let ensuredUserId: number | null = null;

      // // ✅ Only process if user selected
      // if (selectedUser && selectedUser.length > 0) {
      //   const userEmail = selectedUser[0]?.secondaryText;

      //   if (userEmail) {
      //     const ensuredUser = await sp.web.ensureUser(userEmail);
      //     ensuredUserId = ensuredUser.Id;
      //   }
      // }
      const flow = await buildApprovalFlow();
      const currentApproverId = flow.length > 0 ? flow[0].Id : null;
      const history = [
        {
          CurrentApprover: employee.EmployeeName,
          ActionTaken: "Draft",
          Comment: "Saved as draft",
          Date: new Date().toISOString()
        }
      ];
      await sp.web.lists.getByTitle("Installation").items.add({
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

        VendorCode: selectedVendorId ? selectedVendorId.toString() : "",
        VendorName: selectedVendorName || "",

        PONumber: poNumber ? poNumber.toString() : "",
        POdate: poDate ? new Date(poDate) : null,
        POPaymentTerms: poTerms || "",
        POAmount: poAmount || "",

        TotalPaymentofProject: advanceAmount || "0",
        GSTAdjustmentifAny: gstAdjustment || "0",
        OtherAdjustmentifany: otherAdjustment || "0",
        TotalamounttobeCapitalized: paidAmount || "0",

        Status: "Draft",
        ApprovalMatrix: JSON.stringify(flow),
        CurrentApproverId: currentApproverId,
        WorkFlowHistory: JSON.stringify(history)
      });

      await uploadAttachments(PaymentId);


      alert("Draft saved successfully ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/Commision.aspx?page=User";
    } catch (error) {
      console.error("ERROR:", error);
      alert("ERROR");
    }
  };

  React.useEffect(() => {
    debugger;
    if (!context) return;
    debugger;
    void getLoggedInUser();
    void getVendors(); // 👈 ADD THIS
    void getPaidPOs();
  }, [context]);

  return (
    <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
      <div className='row'>
        <div className='col-md-12'>
          <div className='Main-Boxpoup'>
            {/* 🔹 Header */}
            <div className="bordered">
              <img src={logo} />
              <h1> Installation Commisioning Request </h1>
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
                    <select value={selectedVendorId || ""} onChange={(e) => {
                      const id = Number(e.target.value);
                      const vendor = vendors.find((v) => v.Id === id);
                      setSelectedVendorId(id); setSelectedVendorName(vendor?.VendorName || "");
                      if (id) { void getPreviousAdvances(id); }
                    }} className='formtext-control'>
                      <option value="">Select Vendor</option>
                      {vendors.map(
                        (v,) => (
                          <option key={v.Id} value={v.Id} >
                            {v.VendorCode}
                          </option>
                        ),
                      )}
                    </select>
                  </div>
                  <div className="col-md-4">
                    <label className='font'>Vendor Name</label>
                    <input value={selectedVendorName} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>PO Number</label>
                    <select
                      value={poNumber} className='formtext-control'
                      onChange={(e) => {
                        const selectedPO = poList.find(
                          (item) => item.PONumber === e.target.value,
                        );

                        setPoNumber(e.target.value);

                        if (selectedPO) {
                          setPoDate(
                            selectedPO.PODate
                              ? new Date(selectedPO.PODate).toISOString().split("T")[0]
                              : "",
                          );

                          setPoTerms(selectedPO.POAdvanceTerms || "");
                          setPoAmount(selectedPO.POAmtGST || "");
                        }
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
                    <label className='font'>PO Date</label>
                    <input type="date" value={poDate} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>PO Payment Terms</label>
                    <input value={poTerms} className='form-control readonly' />
                  </div>
                  <div className="col-md-4">
                    <label className='font'>PO Amount (GST)</label>
                    <input value={poAmount} className='form-control readonly' />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Total Payment for the Project</label>
                    <input value={advanceAmount} className='form-control' onChange={(e) => handleNumberChange(e.target.value, setAdvanceAmount)} />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Gst Adjustment(Any)</label>
                    <input value={gstAdjustment} className='form-control' onChange={(e) => handleNumberChange(e.target.value, setGstAdjustment)} />
                  </div>
                  <div className="col-md-4">
                    <label className="font">Other Adjustment</label>
                    <input value={otherAdjustment} className='form-control' onChange={(e) => handleNumberChange(e.target.value, setOtherAdjustment)} />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font" style={{ color: "red" }}>Total Project Amount to be Capitalized</label>
                    <input value={paidAmount} className='form-control readonly' />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Past MRN Details</label>
              </div>
              <div className='main-formcontainer'>
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
                            <th className="px-4 py-2"> Advance Adjustment</th>
                            <th className="px-4 py-2">Paid Amount</th>
                            <th className="px-4 py-2">Remarks</th>
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
                <label>Upload Docuement</label>
              </div>
              <div className='main-formcontainer'>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attachment</label>
                    <input type="file" className="form-control" multiple onChange={(e) => { if (e.target.files) { setAttachments(Array.from(e.target.files)); } }} />
                  </div>
                </div>
              </div>
              <div style={{ display: "flex", justifyContent: "center", gap: "5px", marginBottom: "1rem", marginTop: "1rem" }}>
                <a onClick={handleSubmit} className="submit-btn">
                  Submit
                </a>
                <a onClick={handledraft} className="Rework-btn">
                  Save as Draft
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

export default NewAdvanceform;
