import React, { useEffect, useRef } from "react";
import "../../ui.css";

// Very small dialog shim. When `open` is falsy, render nothing.
export function Dialog({ open, onOpenChange, children }) {
    const overlayMouseDown = useRef(false);

    useEffect(() => {
        if (open) {
            const prev = document.body.style.overflow;
            document.body.style.overflow = "hidden";
            return () => { document.body.style.overflow = prev; };
        }
    }, [open]);

    if (!open) return null;
    // Only close when both mousedown AND mouseup happened on the overlay backdrop
    // (not when the user drags text from inside the dialog out to the overlay).
    return (
        <div
            className="dialog-overlay"
            onMouseDown={(e) => { overlayMouseDown.current = e.target === e.currentTarget; }}
            onClick={(e) => { if (overlayMouseDown.current && e.target === e.currentTarget && onOpenChange) onOpenChange(null); }}
        >
            {children}
        </div>
    );
}

export function DialogContent({ children, className = "", style = {} }) {
    return (
        <div
            className={`dialog-content ${className}`.trim()}
            style={style}
            onClick={(e) => e.stopPropagation()}
        >
            {children}
        </div>
    );
}

export function DialogHeader({ children, className = "", style = {} }) {
    const baseStyle = { marginBottom: 8, ...style };
    return <div className={className} style={baseStyle}>{children}</div>;
}

export function DialogTitle({ children }) {
    return <h3 style={{ margin: 0 }}>{children}</h3>;
}

export function DialogFooter({ children, className = "", style = {} }) {
    const baseStyle = { marginTop: 12, display: "flex", justifyContent: "flex-end", gap: 8, ...style };
    return <div className={className} style={baseStyle}>{children}</div>;
}
