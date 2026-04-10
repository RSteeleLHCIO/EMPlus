import React from "react";
import { createRoot } from "react-dom/client";
import EntityApp from "./EntityApp";
import "./styles.css";
import "./ui.css";

const el = document.getElementById("root");
createRoot(el).render(
    <React.StrictMode>
        <EntityApp />
    </React.StrictMode>
);
