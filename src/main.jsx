import React, { useState } from "react";
import { createRoot } from "react-dom/client";
import EntityApp from "./EntityApp";
import LoginScreen from "./LoginScreen";
import "./styles.css";
import "./ui.css";

function Root() {
  const [token, setToken] = useState(() => sessionStorage.getItem("emplus_token"));
  const [clientId, setClientId] = useState(() => sessionStorage.getItem("emplus_clientId"));
  const [sessionExpired, setSessionExpired] = useState(false);
  const [emulation, setEmulation] = useState(() => {
    if (!sessionStorage.getItem("emplus_token") || !sessionStorage.getItem("emplus_admin_token")) {
      return null;
    }
    try {
      const raw = sessionStorage.getItem("emplus_emulation");
      return raw ? JSON.parse(raw) : null;
    } catch {
      return null;
    }
  });

  const handleLogin = (newToken, newClientId) => {
    sessionStorage.setItem("emplus_token", newToken);
    sessionStorage.setItem("emplus_clientId", newClientId || "");
    setToken(newToken);
    setClientId(newClientId || "");
    setSessionExpired(false);
  };

  const handleSignOut = (expired = false) => {
    sessionStorage.removeItem("emplus_token");
    sessionStorage.removeItem("emplus_clientId");
    sessionStorage.removeItem("emplus_admin_token");
    sessionStorage.removeItem("emplus_emulation");
    setToken(null);
    setClientId(null);
    setEmulation(null);
    setSessionExpired(!!expired);
  };

  if (!token) return <LoginScreen onLogin={handleLogin} sessionExpired={sessionExpired} />;
  const handleStartEmulation = (emulatedToken, emulatedClientId, identity) => {
    if (emulation) throw new Error("An emulation session is already active.");
    sessionStorage.setItem("emplus_admin_token", token);
    sessionStorage.setItem("emplus_token", emulatedToken);
    sessionStorage.setItem("emplus_clientId", emulatedClientId || "");
    sessionStorage.setItem("emplus_emulation", JSON.stringify(identity));
    setEmulation(identity);
    setToken(emulatedToken);
    setClientId(emulatedClientId || "");
  };
  const handleStopEmulation = () => {
    const adminToken = sessionStorage.getItem("emplus_admin_token");
    if (!adminToken) {
      handleSignOut();
      return;
    }
    sessionStorage.setItem("emplus_token", adminToken);
    sessionStorage.removeItem("emplus_admin_token");
    sessionStorage.removeItem("emplus_emulation");
    setToken(adminToken);
    setEmulation(null);
  };

  return (
    <>
      {emulation && (
        <div className="emulation-banner" style={{
          display: "flex", alignItems: "center", justifyContent: "center", gap: 12,
          padding: "8px 16px", background: "#fff7ed", borderBottom: "1px solid #fdba74",
          color: "#9a3412", fontSize: 13, fontWeight: 600,
        }}>
          <span>Emulating {emulation.loginId} as {emulation.adminLoginId}. Changes are saved to their account.</span>
          <button
            type="button"
            onClick={handleStopEmulation}
            style={{
              padding: "4px 10px", border: "1px solid #c2410c", borderRadius: 5,
              background: "#fff", color: "#9a3412", fontWeight: 600, cursor: "pointer",
            }}
          >
            Return to admin
          </button>
        </div>
      )}
      <EntityApp
        key={token}
        token={token}
        clientId={clientId}
        onSignOut={handleSignOut}
        onStartEmulation={handleStartEmulation}
        isEmulating={!!emulation}
      />
    </>
  );
}

const el = document.getElementById("root");
createRoot(el).render(
  <React.StrictMode>
    <Root />
  </React.StrictMode>
);
