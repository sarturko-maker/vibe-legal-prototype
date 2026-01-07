/**
 * InputArea Component
 * Text input with "Preview changes" toggle and Settings button
 * Simple UI: checkbox controls whether changes are previewed first
 */

import React, { useState, useEffect, KeyboardEvent } from 'react';
import { useChat } from '../state/ChatContext';
import { useChatHistory } from '../state/ChatHistoryContext';
import { useSettings } from '../state/SettingsContext';
import { useToast } from '../state/ToastContext';
import { handleAction } from '../services/handleAction';
import { SideState, DetectedParties } from './SideSelector';
import { DealContextState } from '../prompts/systemPrompt';
import { RiskTolerance } from '../types/state';
import { getTitleFromFirstMessage } from '../services/chatTitleGeneration';
import './InputArea.css';

// Store last message for re-execution on Accept
let lastMessage = '';

interface InputAreaProps {
    selectedSide?: SideState | null;
    detectedParties?: DetectedParties | null;
    dealContext?: DealContextState | null;
    riskTolerance?: RiskTolerance | null;  // ADR-012
    onOpenSettings?: () => void;  // Callback to open Settings
    previewMode?: boolean;  // Preview Mode for demo features
}

export function InputArea({ selectedSide, detectedParties, dealContext, riskTolerance, onOpenSettings, previewMode }: InputAreaProps) {
    const [input, setInput] = useState('');
    const [previewEnabled, setPreviewEnabled] = useState(false);
    const [webSearchEnabled, setWebSearchEnabled] = useState(false);
    const { showToast } = useToast();
    const {
        messages,            // Get chat history for context
        isProcessing,
        addMessage,
        setIsProcessing,
        setProcessingStage,
        setPendingPreview,
        setPendingOperations,
        setPreviewOriginalTexts
    } = useChat();
    const {
        addMessageToCurrentChat,
        updateCurrentChatTitle,
        currentChat
    } = useChatHistory();
    const { getCurrentApiKey, getCurrentModel, getAuthorName, provider } = useSettings();

    // Auto-generate title when chat reaches 3 messages
    useEffect(() => {
        if (currentChat &&
            currentChat.title === 'New Chat' &&
            currentChat.messages.length >= 3) {
            const title = getTitleFromFirstMessage(currentChat.messages);
            if (title && title !== 'New Chat') {
                updateCurrentChatTitle(title);
            }
        }
    }, [currentChat?.messages.length, currentChat?.title]);

    // Helper to add message to both contexts
    const addMessageToBoth = (msg: { role: 'user' | 'bot'; content: string }) => {
        addMessage(msg);
        addMessageToCurrentChat(msg);
    };

    const hasApiKey = !!getCurrentApiKey();

    const handleSend = async () => {
        const message = input.trim();
        if (!message || isProcessing) return;

        const apiKey = getCurrentApiKey();
        if (!apiKey) {
            addMessageToBoth({ role: 'bot', content: '⚠️ Please add your API key in Settings.' });
            return;
        }

        // Store for potential re-execution
        lastMessage = message;

        setInput('');
        addMessageToBoth({ role: 'user', content: message });
        setIsProcessing(true);

        try {
            const model = getCurrentModel();
            const author = getAuthorName();

            // Use previewEnabled toggle to determine mode
            const mode = previewEnabled ? 'draft' : 'chat';

            // Log side selection and chat history for debugging
            console.log('[InputArea] Sending with side:', selectedSide?.selected || 'neutral');
            console.log('[InputArea] Detected parties:', detectedParties ?
                `${detectedParties.partyA.shortName} vs ${detectedParties.partyB.shortName}` : 'none');
            console.log('[InputArea] Risk tolerance:', riskTolerance?.enabled ? `${riskTolerance.level}%` : 'none');
            console.log('[InputArea] Chat history:', messages.length, 'messages');

            const result = await handleAction(
                message,
                apiKey,
                model,
                author,
                mode,
                (stage) => setProcessingStage(stage),
                selectedSide,        // Pass side selection
                detectedParties,     // Pass detected parties
                messages,            // Pass chat history for context
                dealContext,         // Pass deal context
                riskTolerance,       // Pass risk tolerance (ADR-012)
                provider             // Pass AI provider
            );

            if (result.isDraft && result.preview) {
                // PREVIEW MODE: Store preview, don't apply yet
                // Note: No chat message here - PreviewPanel header shows pending count
                setPendingOperations(result.preview.operations);
                setPreviewOriginalTexts(result.preview.originalTexts || {});
                setPendingPreview({
                    id: Date.now().toString(),
                    operation: result.preview.operations[0],
                    description: result.preview.message,
                    diffHtml: '',
                    originalText: '',
                    newText: ''
                });
            } else if (result.success) {
                // IMMEDIATE MODE: Changes already applied with track changes
                addMessageToBoth({ role: 'bot', content: result.answer || 'Done.' });
            } else {
                addMessageToBoth({ role: 'bot', content: `❌ ${result.error}` });
            }

        } catch (error: any) {
            addMessageToBoth({ role: 'bot', content: `❌ Error: ${error.message}` });
        } finally {
            setIsProcessing(false);
            setProcessingStage('');
        }
    };

    const handleKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleSend();
        }
    };

    return (
        <div className="input-area">
            <div className="input-row">
                <textarea
                    className="input-textarea"
                    value={input}
                    onChange={(e) => setInput(e.target.value)}
                    onKeyDown={handleKeyDown}
                    placeholder="Ask about this document..."
                    disabled={isProcessing}
                    rows={1}
                />
                <button
                    className="input-send-btn"
                    onClick={handleSend}
                    disabled={!input.trim() || isProcessing}
                >
                    Send
                </button>
            </div>

            <div className="input-footer">
                <div className="input-footer__left">
                    <button
                        className="toggle-switch"
                        onClick={() => setPreviewEnabled(!previewEnabled)}
                    >
                        <div className={`toggle-switch__track ${previewEnabled ? 'toggle-switch__track--active' : ''}`}>
                            <div className="toggle-switch__thumb" />
                        </div>
                        <span className="toggle-switch__label">Preview changes</span>
                    </button>

                    {previewMode && (
                        <button
                            className="toggle-switch toggle-switch--preview"
                            onClick={() => {
                                setWebSearchEnabled(!webSearchEnabled);
                                showToast('Coming soon');
                            }}
                        >
                            <div className={`toggle-switch__track ${webSearchEnabled ? 'toggle-switch__track--active' : ''}`}>
                                <div className="toggle-switch__thumb" />
                            </div>
                            <span className="toggle-switch__label">Web Search</span>
                        </button>
                    )}
                </div>

                <div className="input-footer__right">
                    {previewMode && (
                        <span className="settings-link preview-mode-indicator">
                            Preview Mode
                            <span className="settings-link__dot settings-link__dot--preview" />
                        </span>
                    )}

                    <button
                        className="settings-link"
                        onClick={onOpenSettings}
                    >
                        Settings
                        {!hasApiKey && <span className="settings-link__dot settings-link__dot--warning" />}
                    </button>
                </div>
            </div>
        </div>
    );
}

// Export for use by PreviewPanel Accept handler
export function getLastMessage() {
    return lastMessage;
}
