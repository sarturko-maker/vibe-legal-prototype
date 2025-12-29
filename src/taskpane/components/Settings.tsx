import React, { useState, useEffect } from 'react';

interface SettingsProps {
    apiKey: string;
    setApiKey: (key: string) => void;
    selectedModel: string;
    setSelectedModel: (model: string) => void;
    onBack: () => void;
}

const Settings: React.FC<SettingsProps> = ({ apiKey, setApiKey, selectedModel, setSelectedModel, onBack }) => {
    const [status, setStatus] = useState("");
    const [models, setModels] = useState<any[]>([]);

    useEffect(() => {
        if (apiKey) handleSync();
    }, [apiKey]);

    const handleSync = async () => {
        setStatus("Loading...");
        try {
            if (!apiKey) throw new Error("No API Key");
            const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
            const response = await fetch(url);
            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error?.message || "Failed to fetch models");
            }
            const data = await response.json();
            const list = (data.models || [])
                .filter((m: any) => m.supportedGenerationMethods?.includes("generateContent"))
                .sort((a: any, b: any) => b.name.localeCompare(a.name));

            setModels(list);
            setStatus(list.length + " models found");

            const currentExists = list.find((m: any) => m.name.replace("models/", "") === selectedModel);
            if (!currentExists && list.length > 0) {
                setSelectedModel(list[0].name.replace("models/", ""));
            }
        } catch (e) {
            setStatus("Error: Check Key");
            setModels([]);
        }
    };

    return (
        <div style={{ padding: "20px", background: "#f9f9f9", height: "100%", fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif', fontSize: "13px" }}>
            <h3 style={{ marginTop: 0, fontSize: "14px", marginBottom: "20px" }}>Settings</h3>
            <div style={{ marginBottom: "20px" }}>
                <label style={{ display: "block", fontWeight: "600", marginBottom: "6px", color: "#444" }}>API Key</label>
                <input
                    type="password"
                    value={apiKey}
                    onChange={(e) => setApiKey(e.target.value)}
                    style={{ width: "100%", padding: "10px", border: "1px solid #ddd", borderRadius: "6px", fontSize: "13px" }}
                    placeholder="Enter Gemini API Key"
                />
            </div>
            <div style={{ marginBottom: "20px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "6px" }}>
                    <label style={{ fontWeight: "600", color: "#444" }}>Model Selection</label>
                    <button
                        onClick={handleSync}
                        style={{
                            fontSize: "12px",
                            background: "#f0f0f0",
                            border: "1px solid #ccc",
                            borderRadius: "4px",
                            padding: "4px 10px",
                            color: "#000",
                            cursor: "pointer",
                        }}
                    >
                        ↻ Refresh Models
                    </button>
                </div>
                <select
                    value={selectedModel}
                    onChange={(e) => setSelectedModel(e.target.value)}
                    style={{
                        width: "100%",
                        padding: "10px",
                        border: "1px solid #ddd",
                        borderRadius: "6px",
                        fontSize: "13px",
                        background: "#fff",
                    }}
                >
                    {models.length > 0
                        ? models.map((m) => (
                            <option key={m.name} value={m.name.replace("models/", "")}>
                                {m.displayName || m.name}
                            </option>
                        ))
                        : <option value={selectedModel}>{selectedModel || "Default"}</option>
                    }
                </select>
                <div style={{ fontSize: "11px", color: "#666", marginTop: "6px" }}>{status}</div>
            </div>
            <button
                onClick={() => {
                    localStorage.setItem("vibe_api_key", apiKey);
                    localStorage.setItem("vibe_model", selectedModel);
                    onBack();
                }}
                style={{
                    width: "100%",
                    background: "#000",
                    color: "#fff",
                    padding: "12px",
                    border: "none",
                    borderRadius: "6px",
                    cursor: "pointer",
                    fontWeight: "600",
                    fontSize: "13px",
                }}
            >
                Done
            </button>
        </div>
    );
};

export default Settings;

