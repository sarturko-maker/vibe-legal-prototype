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
    geminiModel: 'gemini-2.0-flash-exp',
    claudeModel: 'claude-sonnet-4-20250514',
    authorMode: 'auto',
    customAuthor: '',
    detectedAuthor: null
};

interface SettingsContextType extends SettingsState {
    setProvider: (provider: AIProvider) => void;
    setGeminiApiKey: (key: string) => void;
    setClaudeApiKey: (key: string) => void;
    setGeminiModel: (model: string) => void;
    setClaudeModel: (model: string) => void;
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
        return settings.provider === 'gemini'
            ? settings.geminiApiKey
            : settings.claudeApiKey;
    };

    const getCurrentModel = () => {
        return settings.provider === 'gemini'
            ? settings.geminiModel
            : settings.claudeModel;
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
        setGeminiModel: (geminiModel) => setSettings(s => ({ ...s, geminiModel })),
        setClaudeModel: (claudeModel) => setSettings(s => ({ ...s, claudeModel })),
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
