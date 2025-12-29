import React from 'react';

interface ModeSwitchProps {
    mode: "ASK" | "REDLINE" | "DRAFT";
    setMode: (mode: "ASK" | "REDLINE" | "DRAFT") => void;
}

const ModeSwitch: React.FC<ModeSwitchProps> = ({ mode, setMode }) => {
    const btn = (active: boolean) => ({
        flex: 1,
        padding: "8px",
        border: active ? "1px solid #000" : "1px solid #e5e5e5",
        background: active ? "#000" : "#fff",
        color: active ? "#fff" : "#666",
        cursor: "pointer",
        fontSize: "12px",
        fontWeight: "500",
        transition: "0.2s",
        fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
        borderRadius: "6px",
    });

    return (
        <div style={{ display: "flex", padding: "15px 16px 0", gap: "8px" }}>
            <button style={btn(mode === "ASK")} onClick={() => setMode("ASK")}>
                Ask
            </button>
            <button style={btn(mode === "REDLINE")} onClick={() => setMode("REDLINE")}>
                Redline
            </button>
        </div>
    );
};

export default ModeSwitch;
