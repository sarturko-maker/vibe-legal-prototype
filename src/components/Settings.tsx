/**
 * Settings Component for Vibe Legal
 * Side panel for API key and model selection
 * Uses panel pattern like Definitions for consistent UX
 */

import React, { useState, useEffect, useRef } from 'react';
import { useSettings } from '../state/SettingsContext';
import './Settings.css';

interface SettingsProps {
    onSettingsSaved?: (apiKey: string) => void;
    isOpen?: boolean;
    onClose?: () => void;
}

const Settings: React.FC<SettingsProps> = ({ onSettingsSaved, isOpen: externalIsOpen, onClose }) => {
    const {
        geminiApiKey,
        setGeminiApiKey,
        geminiModel,
        setGeminiModel
    } = useSettings();

    const [status, setStatus] = useState("");
    const [models, setModels] = useState<any[]>([]);
    const panelRef = useRef<HTMLDivElement>(null);

    // Use external control if provided
    const isOpen = externalIsOpen !== undefined ? externalIsOpen : false;
    const closePanel = onClose || (() => { });

    // Local state for editing (so Cancel can discard)
    const [localApiKey, setLocalApiKey] = useState(geminiApiKey);
    const [localModel, setLocalModel] = useState(geminiModel);

    // Reset local state when panel opens
    useEffect(() => {
        if (isOpen) {
            setLocalApiKey(geminiApiKey);
            setLocalModel(geminiModel);
        }
    }, [isOpen, geminiApiKey, geminiModel]);

    // Load models when api key changes
    useEffect(() => {
        if (localApiKey && localApiKey.length > 10) {
            handleSync(localApiKey);
        }
    }, [localApiKey]);

    const handleSync = async (apiKey?: string) => {
        const keyToUse = apiKey || localApiKey;
        setStatus("Loading...");
        try {
            if (!keyToUse) throw new Error("No API Key");
            const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${keyToUse}`;
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

            // Auto-select first model if current doesn't exist
            const currentExists = list.find((m: any) => m.name.replace("models/", "") === localModel);
            if (!currentExists && list.length > 0) {
                setLocalModel(list[0].name.replace("models/", ""));
            }
        } catch (e) {
            setStatus("Error: Check Key");
            setModels([]);
        }
    };

    // Done handler - saves and triggers callback
    const handleDone = () => {
        setGeminiApiKey(localApiKey);
        setGeminiModel(localModel);

        if (onSettingsSaved) {
            onSettingsSaved(localApiKey);
        }

        closePanel();
    };

    if (!isOpen) return null;

    return (
        <div className="panel-overlay">
            <div className="panel-overlay__backdrop" onClick={closePanel} />
            <div className="panel" ref={panelRef}>
                <div className="panel__header">
                    <div className="panel__header-row">
                        <h2 className="panel__title">Settings</h2>
                        <button className="panel__close" onClick={closePanel}>
                            Close
                        </button>
                    </div>
                </div>

                <div className="panel__content">
                    <div className="settings__section">
                        <label className="settings__label">API Key</label>
                        <input
                            type="password"
                            className="settings__input"
                            value={localApiKey}
                            onChange={(e) => setLocalApiKey(e.target.value)}
                            placeholder="Enter Gemini API Key"
                        />
                    </div>

                    <div className="settings__section">
                        <div className="settings__label-row">
                            <label className="settings__label">Model</label>
                            <button onClick={() => handleSync()} className="settings__refresh">
                                Refresh
                            </button>
                        </div>
                        <div className="settings__select-wrapper">
                            <select
                                className="settings__select"
                                value={localModel}
                                onChange={(e) => setLocalModel(e.target.value)}
                            >
                                {models.length > 0
                                    ? models.map((m) => (
                                        <option key={m.name} value={m.name.replace("models/", "")}>
                                            {m.displayName || m.name}
                                        </option>
                                    ))
                                    : <option value={localModel}>{localModel || "Default"}</option>
                                }
                            </select>
                            <span className="settings__select-arrow">▼</span>
                        </div>
                    </div>

                    {status && (
                        <div className="settings__status">
                            <p className="settings__status-text">{status}</p>
                        </div>
                    )}
                </div>

                <div className="settings__footer">
                    <button
                        className="settings__save-btn"
                        onClick={handleDone}
                        disabled={!localApiKey}
                    >
                        Save Changes
                    </button>
                </div>
            </div>
        </div>
    );
};

export { Settings };
