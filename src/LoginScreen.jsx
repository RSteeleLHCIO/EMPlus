import React, { useState } from "react";

const API_BASE = import.meta.env.VITE_API_URL || "";

export default function LoginScreen({ onLogin, sessionExpired }) {
  const [loginId, setLoginId] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);
    try {
      const res = await fetch(`${API_BASE}/api/auth/login`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ loginId: loginId.trim().toLowerCase(), password }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Login failed");
      onLogin(data.token, data.clientId, data.personId || null);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{
      display: "flex", alignItems: "center", justifyContent: "center",
      minHeight: "100vh", background: "#f5f5f7",
    }}>
      <div style={{
        background: "#fff", borderRadius: 12, padding: "40px 48px",
        boxShadow: "0 2px 16px rgba(0,0,0,0.1)", width: 360,
      }}>
        <div style={{ textAlign: "center", marginBottom: '16px' }}>
          <img
            src="/emplus-logo.png"
            alt="EMPlus"
            style={{ width: 200, height: "auto" }}
          />
        </div>
        <h2 style={{ margin: "0 0 24px", fontSize: 22, fontWeight: 700, color: "#1a1a2e" }}>
          Sign In
        </h2>
        {sessionExpired && (
          <div style={{
            marginBottom: 16, padding: "10px 12px", background: "#fefce8",
            border: "1px solid #fde68a", borderRadius: 8, color: "#92400e", fontSize: 13,
          }}>
            Your session has expired. Please sign in again.
          </div>
        )}
        <form onSubmit={handleSubmit}>
          <div style={{ marginBottom: 16 }}>
            <label style={{ display: "block", marginBottom: 6, fontSize: 13, fontWeight: 500, color: "#444" }}>
              Login ID
            </label>
            <input
              type="text"
              value={loginId}
              onChange={(e) => setLoginId(e.target.value)}
              required
              autoFocus
              style={{
                width: "100%", padding: "10px 12px", border: "1px solid #ddd",
                borderRadius: 8, fontSize: 14, boxSizing: "border-box",
              }}
            />
          </div>
          <div style={{ marginBottom: 24 }}>
            <label style={{ display: "block", marginBottom: 6, fontSize: 13, fontWeight: 500, color: "#444" }}>
              Password
            </label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              required
              style={{
                width: "100%", padding: "10px 12px", border: "1px solid #ddd",
                borderRadius: 8, fontSize: 14, boxSizing: "border-box",
              }}
            />
          </div>
          {error && (
            <div style={{
              marginBottom: 16, padding: "10px 12px", background: "#fef2f2",
              border: "1px solid #fecaca", borderRadius: 8, color: "#b91c1c", fontSize: 13,
            }}>
              {error}
            </div>
          )}
          <button
            type="submit"
            disabled={loading}
            style={{
              width: "100%", padding: "10px 12px", background: "#1a1a2e",
              color: "#fff", border: "none", borderRadius: 8, fontSize: 15,
              fontWeight: 600, cursor: loading ? "wait" : "pointer",
            }}
          >
            {loading ? "Signing in…" : "Sign In"}
          </button>
        </form>
      </div>
    </div>
  );
}
