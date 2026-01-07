/**
 * Settings Context for Vibe Legal
 * Manages API keys, provider selection, author settings
 */

import React, { createContext, useContext, useState, ReactNode } from 'react';
import { SettingsState, AIProvider, AuthorMode } from '../types';

const defaultSettings: SettingsState = {
    provider: 'gemini',
    geminiApiKey: '',
    claudeApiKey: '',
    groqApiKey: '',
    mistralApiKey: '',
    geminiModel: 'gemini-2.0-flash-exp',
    claudeModel: 'claude-sonnet-4-20250514',
    groqModelFast: 'llama-3.3-70b-versatile',
    groqModelSlow: 'llama-3.3-70b-versatile',
    mistralModelFast: 'mistral-small-latest',
    mistralModelSlow: 'mistral-large-latest',
    authorMode: 'auto',
    customAuthor: '',
    detectedAuthor: null
};

interface SettingsContextType extends SettingsState {
    setProvider: (provider: AIProvider) => void;
    setGeminiApiKey: (key: string) => void;
    setClaudeApiKey: (key: string) => void;
    setGroqApiKey: (key: string) => void;
    setMistralApiKey: (key: string) => void;
    setGeminiModel: (model: string) => void;
    setClaudeModel: (model: string) => void;
    setGroqModelFast: (model: string) => void;
    setGroqModelSlow: (model: string) => void;
    setMistralModelFast: (model: string) => void;
    setMistralModelSlow: (model: string) => void;
    setAuthorMode: (mode: AuthorMode) => void;
    setCustomAuthor: (name: string) => void;
    setDetectedAuthor: (author: string | null) => void;
    getCurrentApiKey: () => string;
    getCurrentModel: () => string;
    getAuthorName: () => string;
}

const SettingsContext = createContext<SettingsContextType | null>(null);

export function SettingsProvider({ children }: { children: ReactNode }) {
    const [settings, setSettings] = useState<SettingsState>(defaultSettings);

    const getCurrentApiKey = () => {
        switch (settings.provider) {
            case 'gemini':
                return settings.geminiApiKey;
            case 'groq':
                return settings.groqApiKey;
            case 'mistral':
                return settings.mistralApiKey;
            case 'claude':
            default:
                return settings.claudeApiKey;
        }
    };

    const getCurrentModel = () => {
        switch (settings.provider) {
            case 'gemini':
                return settings.geminiModel;
            case 'groq':
                return settings.groqModelFast;
            case 'mistral':
                return settings.mistralModelFast;
            case 'claude':
            default:
                return settings.claudeModel;
        }
    };

    const getAuthorName = () => {
        switch (settings.authorMode) {
            case 'auto':
                return settings.detectedAuthor || 'Vibe AI';
            case 'vibe':
                return 'Vibe AI';
            case 'custom':
                return settings.customAuthor.trim() || 'Vibe AI';
            default:
                return 'Vibe AI';
        }
    };

    const value: SettingsContextType = {
        ...settings,
        setProvider: (provider) => setSettings(s => ({ ...s, provider })),
        setGeminiApiKey: (geminiApiKey) => setSettings(s => ({ ...s, geminiApiKey })),
        setClaudeApiKey: (claudeApiKey) => setSettings(s => ({ ...s, claudeApiKey })),
        setGroqApiKey: (groqApiKey) => setSettings(s => ({ ...s, groqApiKey })),
        setMistralApiKey: (mistralApiKey) => setSettings(s => ({ ...s, mistralApiKey })),
        setGeminiModel: (geminiModel) => setSettings(s => ({ ...s, geminiModel })),
        setClaudeModel: (claudeModel) => setSettings(s => ({ ...s, claudeModel })),
        setGroqModelFast: (groqModelFast) => setSettings(s => ({ ...s, groqModelFast })),
        setGroqModelSlow: (groqModelSlow) => setSettings(s => ({ ...s, groqModelSlow })),
        setMistralModelFast: (mistralModelFast) => setSettings(s => ({ ...s, mistralModelFast })),
        setMistralModelSlow: (mistralModelSlow) => setSettings(s => ({ ...s, mistralModelSlow })),
        setAuthorMode: (authorMode) => setSettings(s => ({ ...s, authorMode })),
        setCustomAuthor: (customAuthor) => setSettings(s => ({ ...s, customAuthor })),
        setDetectedAuthor: (detectedAuthor) => setSettings(s => ({ ...s, detectedAuthor })),
        getCurrentApiKey,
        getCurrentModel,
        getAuthorName
    };

    return (
        <SettingsContext.Provider value={value}>
            {children}
        </SettingsContext.Provider>
    );
}

export function useSettings() {
    const context = useContext(SettingsContext);
    if (!context) {
        throw new Error('useSettings must be used within SettingsProvider');
    }
    return context;
}
