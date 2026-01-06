/**
 * Chat Context for Vibe Legal
 * Manages messages, mode, processing state, amendment queue
 */

import React, { createContext, useContext, useState, ReactNode } from 'react';
import { ChatState, Message, AppMode, ProSubMode, AmendmentPreview, Amendment, Operation } from '../types';

const defaultChatState: ChatState = {
    messages: [],
    isProcessing: false,
    processingStage: '',
    appMode: 'chat',
    proSubMode: 'auto',  // Kept for backwards compatibility
    pendingPreview: null,
    pendingOperations: [],
    currentOpIndex: 0,
    previewOriginalTexts: {},
    amendmentQueue: [],
    currentAmendmentIndex: 0
};

interface ChatContextType extends ChatState {
    addMessage: (message: Message) => void;
    clearMessages: () => void;
    setMessages: (messages: Message[]) => void;
    setIsProcessing: (processing: boolean) => void;
    setProcessingStage: (stage: string) => void;
    setAppMode: (mode: AppMode) => void;
    setProSubMode: (mode: ProSubMode) => void;
    setPendingPreview: (preview: AmendmentPreview | null) => void;
    setPendingOperations: (operations: Operation[]) => void;
    setCurrentOpIndex: (index: number) => void;
    setPreviewOriginalTexts: (texts: { [key: number]: string }) => void;
    clearPendingOperations: () => void;
    addToQueue: (amendment: Amendment) => void;
    clearQueue: () => void;
    nextInQueue: () => void;
    updateAmendmentStatus: (id: string, status: Amendment['status']) => void;
}

const ChatContext = createContext<ChatContextType | null>(null);

export function ChatProvider({ children }: { children: ReactNode }) {
    const [state, setState] = useState<ChatState>(defaultChatState);

    const value: ChatContextType = {
        ...state,

        addMessage: (message) => setState(s => ({
            ...s,
            messages: [...s.messages, message]
        })),

        clearMessages: () => setState(s => ({
            ...s,
            messages: []
        })),

        setMessages: (messages) => setState(s => ({
            ...s,
            messages
        })),

        setIsProcessing: (isProcessing) => setState(s => ({ ...s, isProcessing })),

        setProcessingStage: (processingStage) => setState(s => ({ ...s, processingStage })),

        setAppMode: (appMode) => setState(s => ({ ...s, appMode })),

        setProSubMode: (proSubMode) => setState(s => ({ ...s, proSubMode })),

        setPendingPreview: (pendingPreview) => setState(s => ({ ...s, pendingPreview })),

        setPendingOperations: (pendingOperations) => setState(s => ({ ...s, pendingOperations, currentOpIndex: 0 })),

        setCurrentOpIndex: (currentOpIndex) => setState(s => ({ ...s, currentOpIndex })),

        setPreviewOriginalTexts: (previewOriginalTexts) => setState(s => ({ ...s, previewOriginalTexts })),

        clearPendingOperations: () => setState(s => ({ ...s, pendingOperations: [], currentOpIndex: 0, previewOriginalTexts: {} })),

        addToQueue: (amendment) => setState(s => ({
            ...s,
            amendmentQueue: [...s.amendmentQueue, amendment]
        })),

        clearQueue: () => setState(s => ({
            ...s,
            amendmentQueue: [],
            currentAmendmentIndex: 0
        })),

        nextInQueue: () => setState(s => ({
            ...s,
            currentAmendmentIndex: Math.min(s.currentAmendmentIndex + 1, s.amendmentQueue.length - 1)
        })),

        updateAmendmentStatus: (id, status) => setState(s => ({
            ...s,
            amendmentQueue: s.amendmentQueue.map(a =>
                a.id === id ? { ...a, status } : a
            )
        }))
    };

    return (
        <ChatContext.Provider value={value}>
            {children}
        </ChatContext.Provider>
    );
}

export function useChat() {
    const context = useContext(ChatContext);
    if (!context) {
        throw new Error('useChat must be used within ChatProvider');
    }
    return context;
}
