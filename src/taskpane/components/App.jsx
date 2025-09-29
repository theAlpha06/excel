import React, { useState, useEffect } from "react";
import SalesforceAuth from "./SalesforceAuth";
import Welcome from "./Welcome";
import PushData from "./PushData";

const App = () => {
  const [isOfficeInitialized, setIsOfficeInitialized] = useState(false);
  const [activeTab, setActiveTab] = useState("auth");
  const [isConnected, setIsConnected] = useState(false);
  const [orgType, setOrgType] = useState("sandbox");

  useEffect(() => {
    Office.onReady(() => {
      setIsOfficeInitialized(true);
    });
  }, []);

  if (!isOfficeInitialized) {
    return (
      <div className="ms-welcome">
        <div className="ms-welcome__main">
          <h2>Loading...</h2>
        </div>
      </div>
    );
  }

  const tabStyle = {
    display: "flex",
    marginBottom: "20px",
    borderBottom: "1px solid #ccc",
  };

  const container = {
    height: "fit-content",
    padding: "1rem",
  };

  const tabItemStyle = (isActive) => ({
    padding: "10px 20px",
    cursor: "pointer",
    backgroundColor: isActive ? "#f0f0f0" : "transparent",
    borderBottom: isActive ? "2px solid #0078d4" : "none",
    fontWeight: isActive ? "bold" : "normal",
  });

  return (
    <div style={container}>
      {isConnected ? (
        <div style={{ paddingBottom: "60px" }}>
          <div style={tabStyle}>
            <div style={tabItemStyle(activeTab === "auth")} onClick={() => setActiveTab("auth")}>
              Profile
            </div>
            {/* <div style={tabItemStyle(activeTab === "data")} onClick={() => setActiveTab("data")}>
              Pull Data
            </div> */}
            <div style={tabItemStyle(activeTab === "insert")} onClick={() => setActiveTab("insert")}>
              Push Data
            </div>
          </div>

          {activeTab === "auth" && (
            <SalesforceAuth
              setIsConnected={setIsConnected}
              isConnected={isConnected}
              orgType={orgType}
            />
          )}
          {/* {activeTab === "data" && <SalesforceData isAuthenticated={isConnected} />} */}
          {activeTab === "insert" && <PushData />}
        </div>
      ) : (
        <>
          <Welcome orgType={orgType} setOrgType={setOrgType} />
          <div style={{ marginTop: "20px" }}>
            <SalesforceAuth
              setIsConnected={setIsConnected}
              isConnected={isConnected}
              orgType={orgType}
            />
          </div>
        </>
      )}
    </div>
  );
};

export default App;
