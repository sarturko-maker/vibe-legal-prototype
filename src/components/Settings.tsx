/**
 * Settings Component for Vibe Legal
 * Side panel for API key and model selection
 * Uses panel pattern like Definitions for consistent UX
 */

import React, { useState, useEffect, useRef } from 'react';
import { useSettings } from '../state/SettingsContext';
import { fetchOpenAICompatibleModels } from '../services/groq';
import { AIProvider } from '../types';
import { DeployLocallyModal } from './DeployLocallyModal';
import './Settings.css';

interface SettingsProps {
    onSettingsSaved?: (apiKey: string) => void;
    isOpen?: boolean;
    onClose?: () => void;
    previewMode?: boolean;
    onPreviewModeChange?: (enabled: boolean) => void;
}

const Settings: React.FC<SettingsProps> = ({ onSettingsSaved, isOpen: externalIsOpen, onClose, previewMode, onPreviewModeChange }) => {
    const {
        provider,
        setProvider,
        geminiApiKey,
        setGeminiApiKey,
        groqApiKey,
        setGroqApiKey,
        mistralApiKey,
        setMistralApiKey,
        geminiModel,
        setGeminiModel,
        groqModelFast,
        setGroqModelFast,
        mistralModelFast,
        setMistralModelFast
    } = useSettings();

    const [status, setStatus] = useState("");
    const [geminiModels, setGeminiModels] = useState<any[]>([]);
    const [openaiModels, setOpenaiModels] = useState<string[]>([]);  // For Groq/Mistral
    const [deployLocallyOpen, setDeployLocallyOpen] = useState(false);
    const panelRef = useRef<HTMLDivElement>(null);

    // Use external control if provided
    const isOpen = externalIsOpen !== undefined ? externalIsOpen : false;
    const closePanel = onClose || (() => { });

    // Get API key for current provider
    const getApiKeyForProvider = (p: AIProvider) => {
        switch (p) {
            case 'groq': return groqApiKey;
            case 'mistral': return mistralApiKey;
            default: return geminiApiKey;
        }
    };

    // Get model for current provider
    const getModelForProvider = (p: AIProvider) => {
        switch (p) {
            case 'groq': return groqModelFast;
            case 'mistral': return mistralModelFast;
            default: return geminiModel;
        }
    };

    // Local state for editing (so Cancel can discard)
    const [localProvider, setLocalProvider] = useState<AIProvider>(provider);
    const [localApiKey, setLocalApiKey] = useState(getApiKeyForProvider(provider));
    const [localModel, setLocalModel] = useState(getModelForProvider(provider));

    // Reset local state when panel opens
    useEffect(() => {
        if (isOpen) {
            setLocalProvider(provider);
            setLocalApiKey(getApiKeyForProvider(provider));
            setLocalModel(getModelForProvider(provider));
            setOpenaiModels([]);  // Reset model list
        }
    }, [isOpen, provider]);

    // Update local API key when provider changes
    useEffect(() => {
        setLocalApiKey(getApiKeyForProvider(localProvider));
        setLocalModel(getModelForProvider(localProvider));
        setOpenaiModels([]);  // Reset model list when provider changes
    }, [localProvider]);

    // Load models when api key changes (Gemini only - auto-sync)
    useEffect(() => {
        if (localProvider === 'gemini' && localApiKey && localApiKey.length > 10) {
            handleSyncGemini(localApiKey);
        }
    }, [localApiKey, localProvider]);

    const handleSyncGemini = async (apiKey?: string) => {
        const keyToUse = apiKey || localApiKey;
        setStatus("Loading Gemini models...");
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

            setGeminiModels(list);
            setStatus(list.length + " models found");

            // Auto-select first model if current doesn't exist
            const currentExists = list.find((m: any) => m.name.replace("models/", "") === localModel);
            if (!currentExists && list.length > 0) {
                setLocalModel(list[0].name.replace("models/", ""));
            }
        } catch (e) {
            setStatus("Error: Check Key");
            setGeminiModels([]);
        }
    };

    const handleTestOpenAIConnection = async () => {
        const providerName = localProvider === 'groq' ? 'Groq' : 'Mistral';
        setStatus(`Testing ${providerName} connection...`);
        try {
            if (!localApiKey) throw new Error("No API Key");
            const modelList = await fetchOpenAICompatibleModels(localProvider, localApiKey);
            setOpenaiModels(modelList);
            setStatus(`✓ Connected! ${modelList.length} models available`);

            // Set default model if current doesn't exist
            if (!modelList.includes(localModel) && modelList.length > 0) {
                setLocalModel(modelList[0]);
            }
        } catch (e: any) {
            setStatus(`✗ ${e.message || 'Connection failed'}`);
            setOpenaiModels([]);
        }
    };

    // Done handler - saves and triggers callback
    const handleDone = () => {
        // Save provider
        setProvider(localProvider);

        // Save API key and model to correct fields
        switch (localProvider) {
            case 'groq':
                setGroqApiKey(localApiKey);
                setGroqModelFast(localModel);
                break;
            case 'mistral':
                setMistralApiKey(localApiKey);
                setMistralModelFast(localModel);
                break;
            default:
                setGeminiApiKey(localApiKey);
                setGeminiModel(localModel);
                break;
        }

        if (onSettingsSaved) {
            onSettingsSaved(localApiKey);
        }

        closePanel();
    };

    if (!isOpen) return null;

    const isOpenAICompatible = localProvider === 'groq' || localProvider === 'mistral';

    // Get label for current provider
    const getProviderLabel = () => {
        switch (localProvider) {
            case 'groq': return 'Groq';
            case 'mistral': return 'Mistral';
            default: return 'Gemini';
        }
    };

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
                    {/* Provider Selection */}
                    <div className="settings__section">
                        <label className="settings__label">AI Provider</label>
                        <div className="settings__select-wrapper">
                            <select
                                className="settings__select"
                                value={localProvider}
                                onChange={(e) => setLocalProvider(e.target.value as AIProvider)}
                            >
                                <option value="gemini">Google Gemini</option>
                                <option value="groq">Groq (Llama)</option>
                                <option value="mistral">Mistral</option>
                            </select>
                            <span className="settings__select-arrow">▼</span>
                        </div>
                    </div>

                    {/* API Key */}
                    <div className="settings__section">
                        <div className="settings__label-row">
                            <label className="settings__label">
                                {getProviderLabel()} API Key
                            </label>
                            {isOpenAICompatible && (
                                <button
                                    onClick={handleTestOpenAIConnection}
                                    className="settings__refresh"
                                    disabled={!localApiKey}
                                >
                                    Test Connection
                                </button>
                            )}
                        </div>
                        <input
                            type="password"
                            className="settings__input"
                            value={localApiKey}
                            onChange={(e) => setLocalApiKey(e.target.value)}
                            placeholder={`Enter ${getProviderLabel()} API Key`}
                        />
                    </div>

                    {/* Model Selection */}
                    <div className="settings__section">
                        <div className="settings__label-row">
                            <label className="settings__label">Model</label>
                            {localProvider === 'gemini' && (
                                <button onClick={() => handleSyncGemini()} className="settings__refresh">
                                    Refresh
                                </button>
                            )}
                        </div>
                        <div className="settings__select-wrapper">
                            <select
                                className="settings__select"
                                value={localModel}
                                onChange={(e) => setLocalModel(e.target.value)}
                            >
                                {isOpenAICompatible ? (
                                    // Groq/Mistral models
                                    openaiModels.length > 0
                                        ? openaiModels.map((modelId) => (
                                            <option key={modelId} value={modelId}>
                                                {modelId}
                                            </option>
                                        ))
                                        : <option value={localModel}>{localModel || "Test connection to load models"}</option>
                                ) : (
                                    // Gemini models
                                    geminiModels.length > 0
                                        ? geminiModels.map((m) => (
                                            <option key={m.name} value={m.name.replace("models/", "")}>
                                                {m.displayName || m.name}
                                            </option>
                                        ))
                                        : <option value={localModel}>{localModel || "Default"}</option>
                                )}
                            </select>
                            <span className="settings__select-arrow">▼</span>
                        </div>
                    </div>

                    {status && (
                        <div className="settings__status">
                            <p className="settings__status-text">{status}</p>
                        </div>
                    )}

                    {/* Deploy Locally - only visible in Preview Mode */}
                    {previewMode && (
                        <div className="settings__section">
                            <button
                                className="settings__deploy-btn"
                                onClick={() => setDeployLocallyOpen(true)}
                            >
                                Deploy Locally
                            </button>
                        </div>
                    )}

                    {/* Preview Mode Toggle */}
                    <div className={`settings__section settings__section--preview ${previewMode ? 'settings__section--active' : ''}`}>
                        <div className="settings__preview-row">
                            <div className="settings__preview-info">
                                <label className="settings__label">Preview Mode</label>
                                <p className="settings__preview-desc">
                                    Explore upcoming features that are still in development. These are demos only and not yet functional.
                                </p>
                            </div>
                            <button
                                className="settings__preview-toggle"
                                onClick={() => onPreviewModeChange?.(!previewMode)}
                            >
                                <div className={`toggle-switch__track ${previewMode ? 'toggle-switch__track--preview' : ''}`}>
                                    <div className="toggle-switch__thumb" />
                                </div>
                            </button>
                        </div>
                    </div>
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

            {/* Deploy Locally Modal */}
            <DeployLocallyModal
                isOpen={deployLocallyOpen}
                onClose={() => setDeployLocallyOpen(false)}
            />
        </div>
    );
};

export { Settings };
