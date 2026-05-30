import * as React from 'react';
import ApproverDashboard from './ApproverDashboard';
import "./usersite.scss"
import UserDashboard from "./UserDashboard";
import APperformerDashboard from './APperformerDashboard';
import { IDashboardProps } from './IDashboardProps';

import ApLogo from '../assets/ApDashboard.png';
import UserLogo from '../assets/UserDashboard.png';
import ApproverLogo from '../assets/ApproverDashboard.png';

export default function CapexDasboard(props: IDashboardProps) {

  const [page, setPage] = React.useState<string>("home");

  // ✅ Navigation function (updates URL + state)
  const navigate = (pageName: string) => {
    const url = `${props.context.pageContext.web.absoluteUrl}/SitePages/Commision.aspx?page=${pageName}`;
    window.history.pushState({}, "", url);
    setPage(pageName);
  };

  // ✅ On page load → read URL
  React.useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const pageParam = params.get("page");

    if (pageParam) {
      setPage(pageParam);
    }
  }, []);

  // ✅ Handle browser back/forward
  React.useEffect(() => {
    const handlePopState = () => {
      const params = new URLSearchParams(window.location.search);
      const pageParam = params.get("page") || "home";
      setPage(pageParam);
    };

    window.addEventListener("popstate", handlePopState);

    return () => {
      window.removeEventListener("popstate", handlePopState);
    };
  }, []);

  return (
    <div>

      {/* ✅ HOME PAGE */}
      {page === "home" && (
        <>

          <div className="main-container">
            <div className="headSheet"><h2>Installation Commission</h2></div>
            <section className='hero'>
              <div className="overlay"></div>
              <div className="hero-content">
                <div className='card-container'>
                  <div className="infoCard" onClick={() => navigate("User")}>
                    <div className="cardContent">
                      <div className="cardalin">
                        <span className="boximage">
                          <img src={UserLogo} width="25" height="25" />
                        </span>
                        <h4>User Dasboard</h4>
                      </div>
                    </div>
                  </div>
                  <div className="infoCard" onClick={() => navigate("Approver")}>
                    <div className="cardContent">
                      <div className="cardalin">
                        <span className="boximage">
                          <img src={ApproverLogo} width="25" height="25" />
                        </span>
                        <h4>Approver Dasboard</h4>
                      </div>
                    </div>
                  </div>
                  <div className="infoCard" onClick={() => navigate("Performer")}>
                    <div className="cardContent">
                      <div className="cardalin">
                        <span className="boximage">
                          <img src={ApLogo} width="25" height="25" />
                        </span>
                        <h4>Ap Performer Dasboard</h4>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </section>
          </div>
        </>
      )}

      {/* ✅ USER DASHBOARD */}
      {page === "User" && (
        <UserDashboard context={props.context} />
      )}


      {page === "Approver" && (
        <ApproverDashboard context={props.context} />
      )}


      {page === "Performer" && (
        <APperformerDashboard context={props.context} />
      )}



    </div>
  );
}