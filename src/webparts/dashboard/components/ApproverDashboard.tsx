import * as React from "react";
import "./userDashboardsc.scss";
import NewAdvanceform from "./NewAdvanceform";
import ViewAdvanceForm from "./ViewAdvanceForm";

import { useState } from "react";

// import sonalogo from "../assets/SonaPNGLogo.png";
// import userlogo from "../assets/userlogo.png";
// import "../assets/bootstrap/css/bootstrap.css";

import logo from "../assets/SonaPNGLogo.png";
import Edit from "../assets/Pencil.png";
import User from "../assets/Userlogo.png";

import ApproverAdvanceForm from "./ApproverAdvanceForm";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
interface IProps {
  context: any;
  itemId: number;
  formData: any;
}

interface UserDashboardProps {
  context: any;
}

const ApproverDashboard: React.FC<UserDashboardProps> = ({ context }) => {
  const sp = spfi().using(SPFx(context));
  //const [formType, setFormType] = useState<"new" | "view" | null>(null);
  const [formType, setFormType] = useState<"new" | "view" | "approve" | null>(
    null,
  );

  const [activeMenu, setActiveMenu] = React.useState("My Request");
  const [searchText, setSearchText] = React.useState("");
  const [statusFilter, setStatusFilter] = React.useState("");
  const [showForm, setShowForm] = React.useState(false);
  const [data, setData] = React.useState<any[]>([]);
  const [currentUserName, setCurrentUserName] = React.useState("");
  const [selectedItem, setSelectedItem] = React.useState<any>(null);

  // ✅ GET CURRENT USER
  const getLoggedInUser = async () => {
    try {
      const user = await sp.web.currentUser();
      setCurrentUserName(user.Title);
    } catch (error) {
      console.error("User error:", error);
    }
  };

  const handleApproveClick = async (item: any) => {
    try {
      const fullItem = await sp.web.lists
        .getByTitle("Installation")
        .items.getById(item.ID)
        .select("*")();

      console.log("Full Item:", fullItem);

      setSelectedItem(fullItem);
      setFormType("approve");
      setShowForm(true);
    } catch (error) {
      console.error("Approve error:", error);
    }
  };

  // ✅ GET LIST DATA
  const getCapexData = async () => {
    try {
      debugger;
      const items = await sp.web.lists
        .getByTitle("Installation")
        .items.select(
          "ID",
          "Title",
          "Created",
          "EmployeeName",
          "VendorName",
          "VendorCode",
          "PONumber",
          "TotalamounttobeCapitalized",
          "Status",
        )
        // .expand("VendorCode") // ✅ ADD THIS LINE
        .filter(`Status eq 'Pending for Approver'`)
        .orderBy("ID", false)();

      const formatted = items.map((item: any) => ({
        ID: item.ID,
        PaymentId: item.Title,
        date: item.Created
          ? new Date(item.Created).toLocaleDateString("en-GB")
          : "",
        EmployeeName: item.EmployeeName || "",
        VendorName: item.VendorName || "",
        VendorCode: item.VendorCode || "",
        PONumber: item.PONumber || "",
        TotalamounttobeCapitalized: item.TotalamounttobeCapitalized || "",
        status: item.Status || "",
      }));

      setData(formatted);
    } catch (error) {
      console.error("Data error:", error);
    }
  };

  // ✅ VIEW CLICK

  // ✅ FILTER
  const filteredData = data.filter((item) => {
    const text = searchText.toLowerCase();
    const status = statusFilter.toLowerCase();

    let menuFilter = true;

    if (activeMenu === "Paid") {
      menuFilter = item.status?.toLowerCase() === "paid";
    } else if (activeMenu === "Rejected") {
      menuFilter = item.status?.toLowerCase() === "rejected";
    } else if (activeMenu === "My Request") {
      menuFilter = true;
    }

    return (
      menuFilter &&
      (item.PaymentId?.toLowerCase().includes(text) ||
        item.EmployeeName?.toLowerCase().includes(text) ||
        item.VendorName?.toLowerCase().includes(text) ||
        item.VendorCode?.toLowerCase().includes(text) ||
        item.PONumber?.toLowerCase().includes(text)) &&
      (!status || item.status?.toLowerCase().includes(status))
    );
  });

  // ✅ LOAD DATA
  React.useEffect(() => {
    if (!context) return;
    debugger;
    void getLoggedInUser();
    void getCapexData();
  }, [context]);

  // ✅ OPEN VIEW PAGE
  if (showForm) {
    if (formType === "approve") {
      return (
        <ApproverAdvanceForm
          context={context}
          formData={selectedItem}
          itemId={selectedItem?.ID}
          onClose={() => {
            setShowForm(false);
            setSelectedItem(null);
            void getCapexData();
          }}
        />
      );
    }
  }

  return (
    <>
      <div style={{ display: "flex", width: "100%" }}>
        <div className="sidebar">
          <div className="sidehead">
            <div className="logo">
              <img src={logo} width="25px" height="25px" />
            </div>
            <div className="sidehead-right">SONA COMSTAR</div>
          </div>

          <div className="sidehead-user">
            <img
              src={User}
              style={{ margin: "10px 20px" }}
              width={20}
              height={20}
            />
            {currentUserName}
          </div>

          <ul className="nav">
            <li className="nav-item">
              <a
                className={
                  activeMenu === "My Request" ? " nav-link active" : "nav-link"
                }
                onClick={() => setActiveMenu("My Request")}
                style={{ cursor: "pointer" }}
              >
                My Request
              </a>
            </li>
            <li className="nav-item">
              <a
                className={
                  activeMenu === "Paid" ? " nav-link  active" : "nav-link"
                }
                onClick={() => setActiveMenu("Paid")}
                style={{ cursor: "pointer" }}
              >
                Paid
              </a>
            </li>
            <li className="nav-item">
              <a
                className={
                  activeMenu === "Rejected" ? "nav-link  active" : "nav-link"
                }
                onClick={() => setActiveMenu("Rejected")}
                style={{ cursor: "pointer" }}
              >
                Rejected
              </a>
            </li>
          </ul>
        </div>
        <div
          className="main"
          style={{ width: "calc(100% - 250px)", transition: "width 0.3s" }}
        >
          <div className="header">
            <div className="left-banner">
              <div className="logo-text">
                <h2> Installation Approver Dashbaord </h2>
              </div>
            </div>
          </div>
          <div className="col-md-12 mainsecond">
            <div>
              <input
                placeholder="Search"
                value={searchText}
                className="form-control"
                style={{ width: "250px;" }}
                onChange={(e) => setSearchText(e.target.value)}
              />
            </div>
            <div>
              <select
                value={statusFilter}
                className="formtext-control"
                onChange={(e) => setStatusFilter(e.target.value)}
              >
                <option value="">All</option>
                <option value="Submitted">Submitted</option>
                <option value="Approved">Approved</option>
                <option value="Rejected">Rejected</option>
                <option value="Draft">Draft</option>
              </select>
            </div>
          </div>
          <main className="Main-Dash mx-2">
            <div style={{ overflowX: "auto" }}>
              <div className="table-vert-scroll">
                <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                  <thead
                    className="text-white"
                    style={{ backgroundColor: "rgb(60, 62, 69)" }}
                  >
                    <tr>
                      <th className="px-4 py-2">Action</th>
                      <th className="px-4 py-2">Payment ID</th>
                      <th className="px-4 py-2">Requestor Name</th>
                      <th className="px-4 py-2">Requestor Date</th>
                      <th className="px-4 py-2">Requestor Type</th>
                      <th className="px-4 py-2">Vendor Code</th>
                      <th className="px-4 py-2">Vendor Name</th>
                      <th className="px-4 py-2">PO Number</th>
                      <th className="px-4 py-2">Capatalized Amount</th>
                      <th className="px-4 py-2">Pending With</th>
                      <th className="px-4 py-2">Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredData.length === 0 ? (
                      <tr>
                        <td colSpan={7} style={{ textAlign: "center" }}>
                          No Data
                        </td>
                      </tr>
                    ) : (
                      filteredData.map((item, i) => (
                        <tr key={i}>
                          <td className="px-4 py-2">
                            <span
                              onClick={() => handleApproveClick(item)}
                              style={{ cursor: "pointer" }}
                            >
                              <img src={Edit} width={15} alt="View" />
                            </span>
                          </td>
                          <td className="px-4 py-2">{item.PaymentId}</td>
                          <td className="px-4 py-2">{item.EmployeeName}</td>
                          <td className="px-4 py-2">{item.date}</td>
                          <td className="px-4 py-2">Install commission</td>
                          <td className="px-4 py-2">{item.VendorCode}</td>
                          <td className="px-4 py-2">{item.VendorName}</td>
                          <td className="px-4 py-2">{item.PONumber}</td>
                          <td className="px-4 py-2">
                            ₹ {item.TotalamounttobeCapitalized}
                          </td>
                          <td className="px-4 py-2">Approver</td>
                          <td className="px-4 py-2">{item.status}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </main>
        </div>
      </div>
    </>
  );
};

export default ApproverDashboard;
