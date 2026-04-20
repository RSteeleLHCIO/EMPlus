import React from "react";

// Small button shim used by the preview. Accepts `variant` (primary|secondary|ghost)
export const Button = React.forwardRef(function Button({ children, variant, className = "", type = "button", ...props }, ref) {
    const base = "btn";
    const vclass =
        variant === "ghost"
            ? "btn-ghost"
            : variant === "secondary"
                ? "btn-secondary"
                : variant === "outline"
                    ? "btn-outline"
                    : "btn-primary";
    return (
        <button ref={ref} type={type} className={`${base} ${vclass} ${className}`.trim()} {...props}>
            {children}
        </button>
    );
});
