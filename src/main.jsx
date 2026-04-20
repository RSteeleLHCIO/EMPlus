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
    setToken(null);
    setClientId(null);
    setSessionExpired(!!expired);
  };

  if (!token) return <LoginScreen onLogin={handleLogin} sessionExpired={sessionExpired} />;
  return <EntityApp token={token} clientId={clientId} onSignOut={handleSignOut} />;
}

const el = document.getElementById("root");
createRoot(el).render(
  <React.StrictMode>
    <Root />
  </React.StrictMode>
);
