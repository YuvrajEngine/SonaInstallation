import * as React from "react";
import "./userDashboardsc.scss";
import NewAdvanceform from "./NewAdvanceform";
import ViewAdvanceForm from "./ViewAdvanceForm";
import EditAdvanceForm from "./EditAdvanceForm";

import { useState } from "react";

import "bootstrap/dist/css/bootstrap.min.css";

import logo from "../assets/SonaPNGLogo.png";
import Edit from "../assets/Pencil.png";
import View from "../assets/Eye.png";
import User from "../assets/Userlogo.png";

import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

interface UserDashboardProps {
  context: any;
}

const UserDashboard: React.FC<UserDashboardProps> = ({ context }) => {
  const sp = spfi().using(SPFx(context));
  const [formType, setFormType] = useState<"new" | "view" | "Edit" | null>(
    null,
  );

  const [activeMenu, setActiveMenu] = React.useState("My Request");
  const [searchText, setSearchText] = React.useState("");
  const [statusFilter, setStatusFilter] = React.useState("");
  const [showForm, setShowForm] = React.useState(false);
  const [data, setData] = React.useState<any[]>([]);
  const [currentUserName, setCurrentUserName] = React.useState("");
  const [currentUserEmail, setCurrentUserEmail] = React.useState("");
  const [selectedItem, setSelectedItem] = React.useState<any>(null);

  const getCapexData = async (userEmail?: string) => {
    const filterEmail = userEmail || currentUserEmail;
    if (!filterEmail) return;

    try {
      const items = await sp.web.lists
        .getByTitle("Installation")
        .items.filter(`Email eq '${filterEmail}'`)
        .select(
          "ID",
          "Title",
          "Created",
          "EmployeeName",
          "Email",
          "VendorName",
          "VendorCode",
          "PONumber",
          "TotalamounttobeCapitalized",
          "Status",
          "CurrentApproverId",
          "CurrentApprover/Title",
          "CurrentApprover/EMail",
        )
        .expand("CurrentApprover")
        .orderBy("ID", false)();

      const formatted = items.map((item: any) => ({
        ID: item.ID,
        PaymentId: item.Title,
        date: item.Created
          ? new Date(item.Created).toLocaleDateString("en-GB")
          : "",
        EmployeeName: item.EmployeeName || "",
        Email: item.Email || "",
        VendorName: item.VendorName || "",
        VendorCode: item.VendorCode || "",
        PONumber: item.PONumber || "",
        TotalamounttobeCapitalized: item.TotalamounttobeCapitalized || "",
        status: item.Status || "",
        CurrentApprover: item.CurrentApprover?.Title || "",
      }));

      setData(formatted);
    } catch (error) {
      console.error("Data error:", error);
    }
  };

  const handleViewClick = async (item: any) => {
    try {
      const fullItem = await sp.web.lists
        .getByTitle("Installation")
        .items.getById(item.ID)
        .select("*")();

      console.log(fullItem);

      setSelectedItem(fullItem);
      setFormType("view");
      setShowForm(true);
    } catch (error) {
      console.error("View error:", error);
    }
  };

  const handleEditClick = async (item: any) => {
    try {
      const fullItem = await sp.web.lists
        .getByTitle("Installation")
        .items.getById(item.ID)
        .select("*")();

      console.log(fullItem);

      setSelectedItem(fullItem);
      setFormType("Edit");
      setShowForm(true);
    } catch (error) {
      console.error("Edit error:", error);
    }
  };

  const filteredData = data.filter((item) => {
    const isOwner =
      item.Email?.toLowerCase() === currentUserEmail?.toLowerCase();
    if (!isOwner) return false;

    const text = searchText.toLowerCase();
    const status = statusFilter.toLowerCase();

    let menuFilter = true;

    if (activeMenu === "Paid") {
      menuFilter = item.status?.toLowerCase() === "paid";
    } else if (activeMenu === "Reject") {
      menuFilter = item.status?.toLowerCase() === "reject";
    } else if (activeMenu === "My Request") {
      const status = item.status?.toLowerCase();
      menuFilter = status !== "paid" && status !== "reject";
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

  React.useEffect(() => {
    if (!context) return;

    const init = async () => {
      try {
        const user = await sp.web.currentUser();
        setCurrentUserName(user.Title);
        setCurrentUserEmail(user.Email);
        await getCapexData(user.Email);
      } catch (error) {
        console.error("Init error:", error);
      }
    };

    void init();
  }, [context]);

  if (showForm) {
    if (formType === "view") {
      return (
        <ViewAdvanceForm
          context={context}
          formData={selectedItem}
          onClose={() => {
            setShowForm(false);
            setFormType(null);
            void getCapexData();
          }}
        />
      );
    }

    if (formType === "new") {
      return (
        <NewAdvanceform
          context={context}
          onClose={() => {
            setShowForm(false);
            setFormType(null);
            void getCapexData();
          }}
        />
      );
    }

    if (formType === "Edit") {
      return (
        <EditAdvanceForm
          context={context}
          formData={selectedItem}
          onClose={() => {
            setShowForm(false);
            setFormType(null);
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
                  activeMenu === "Reject" ? "nav-link  active" : "nav-link"
                }
                onClick={() => setActiveMenu("Reject")}
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
                <h2> Installation User Dashboard </h2>
              </div>
            </div>
          </div>
          <div className="mainsecondapprove">
            <div className="mainsecondsmall">
              <div>
                <input
                  placeholder="Search"
                  value={searchText}
                  className="form-control"
                  style={{ width: "250px", padding: "0px 10px" }}
                  onChange={(e) => setSearchText(e.target.value)}
                />
              </div>
              <div>
                <select
                  value={statusFilter}
                  className="formtext-control"
                  style={{ width: "250px" }}
                  onChange={(e) => setStatusFilter(e.target.value)}
                >
                  <option value="">All</option>
                  <option value="Pending for Approval">
                    Pending for Approval
                  </option>
                  <option value="Reject">Reject</option>
                  <option value="Paid">Paid</option>
                  <option value="Save as Draft">Save as Draft</option>
                  <option value="Send Back">Send Back</option>
                </select>
              </div>
            </div>
            <div>
              <a
                onClick={() => {
                  setSelectedItem(null);
                  setFormType("new");
                  setShowForm(true);
                }}
                className="create-button"
              >
                New Request
              </a>
            </div>
          </div>
          <main className="Main-Dash mx-2">
            <div style={{ overflowX: "auto" }}>
              <div className="table-vert-scroll">
                <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                  <thead>
                    <tr>
                      <th className="px-4 py-2">Action</th>
                      <th className="px-4 py-2">Payment ID</th>
                      <th className="px-4 py-2">Requestor Name</th>
                      <th className="px-4 py-2">Requestor Date</th>
                      <th className="px-4 py-2">Request Type</th>
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
                          <td
                            className="px-4 py-2"
                            style={{
                              display: "flex",
                              gap: "10px",
                              cursor: "pointer",
                            }}
                          >
                            <span
                              onClick={() => handleViewClick(item)}
                              style={{ cursor: "pointer" }}
                            >
                              <img src={View} width={15} alt="View" />
                            </span>

                            {(item.status?.toLowerCase() === "save as draft" ||
                              item.status?.toLowerCase() === "send back") && (
                              <span
                                onClick={() => handleEditClick(item)}
                                style={{ cursor: "pointer" }}
                              >
                                <img src={Edit} width={15} alt="Edit" />
                              </span>
                            )}
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
                          <td className="px-4 py-2">
                            {item.CurrentApprover || "-"}
                          </td>
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

export default UserDashboard;
