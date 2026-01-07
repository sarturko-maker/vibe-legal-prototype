/**
 * Application State Types for Vibe Legal
 * Extracted from vibe-legal-beta-0.2.yaml + 0.3 Pro Mode
 */

import { ContractMap, FocusedClause, StyleInfo } from './document';
import { Operation } from './operations';

// Provider types
export type AIProvider = 'gemini' | 'claude' | 'groq' | 'mistral';

// Provider configuration
export interface ProviderConfig {
    provider: AIProvider;
    geminiApiKey: string;
    claudeApiKey: string;
    groqApiKey: string;
    mistralApiKey: string;
    geminiModel: string;
    claudeModel: string;
    groqModelFast: string;
    groqModelSlow: string;
    mistralModelFast: string;
    mistralModelSlow: string;
}

// Author mode for track changes
export type AuthorMode = 'auto' | 'vibe' | 'custom';

export interface AuthorSettings {
    mode: AuthorMode;
    customName: string;
    detectedAuthor: string | null;
}

// Chat message
export interface Message {
    role: 'user' | 'bot';
    content: string;
}

// App mode - simplified
export type AppMode = 'chat' | 'draft';
export type ProSubMode = 'ask' | 'draft' | 'auto';  // Kept for backwards compatibility

// Amendment preview for Draft mode
export interface AmendmentPreview {
    id: string;
    operation: Operation;
    clauseNumber?: string;
    clauseTitle?: string;
    description: string;
    diffHtml: string;
    originalText: string;
    newText: string;
}

// Amendment for Auto mode queue
export interface Amendment {
    id: string;
    operation: Operation;
    clauseNumber?: string;
    clauseTitle?: string;
    description: string;
    previewHtml: string;
    status: 'pending' | 'accepted' | 'rejected' | 'revised';
}

// Settings Context State
export interface SettingsState {
    provider: AIProvider;
    geminiApiKey: string;
    claudeApiKey: string;
    groqApiKey: string;
    mistralApiKey: string;
    geminiModel: string;
    claudeModel: string;
    groqModelFast: string;
    groqModelSlow: string;
    mistralModelFast: string;
    mistralModelSlow: string;
    authorMode: AuthorMode;
    customAuthor: string;
    detectedAuthor: string | null;
}

// Document Context State
export interface DocumentState {
    contractMap: ContractMap | null;
    styleCache: Map<string, StyleInfo>;
    documentHash: string | null;
    isAnalyzing: boolean;
    focusedClause: FocusedClause | null;
}

// Chat Context State
export interface ChatState {
    messages: Message[];
    isProcessing: boolean;
    processingStage: string;
    appMode: AppMode;
    proSubMode: ProSubMode;  // Kept for backwards compatibility
    pendingPreview: AmendmentPreview | null;
    pendingOperations: Operation[];  // Operations waiting to be applied in Draft mode
    currentOpIndex: number;  // Current operation index for one-at-a-time review
    previewOriginalTexts: { [key: number]: string };  // Original texts for diff display
    amendmentQueue: Amendment[];
    currentAmendmentIndex: number;
}

// Full App State (for reference)
export interface AppState extends SettingsState, DocumentState, ChatState { }

// Risk Tolerance for negotiation calibration (ADR-012)
export interface RiskTolerance {
    enabled: boolean;
    level: number;              // 0-100 (0=conservative, 100=aggressive)
    bestPosition: string;       // Examples of most favorable terms
    fallbackPosition: string;   // Examples of acceptable terms to close deal
}

export type RiskLabel = 'Conservative' | 'Moderate' | 'Aggressive';

export function getRiskLabel(level: number): RiskLabel {
    if (level <= 33) return 'Conservative';
    if (level <= 66) return 'Moderate';
    return 'Aggressive';
}

export function getRiskSublabel(level: number): string {
    if (level <= 10) return 'Very Conservative';
    if (level <= 25) return 'Conservative';
    if (level <= 40) return 'Conservative-Moderate';
    if (level <= 60) return 'Moderate';
    if (level <= 75) return 'Moderate-Aggressive';
    if (level <= 90) return 'Aggressive';
    return 'Very Aggressive';
}

export const initialRiskTolerance: RiskTolerance = {
    enabled: false,
    level: 50,
    bestPosition: '',
    fallbackPosition: ''
};
