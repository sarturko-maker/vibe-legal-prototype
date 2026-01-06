/**
 * Chat History Context
 * Manages persistent chat conversations with localStorage
 */

import React, { createContext, useContext, useState, useEffect, useCallback, ReactNode } from 'react';
import {
    Chat,
    ChatHistoryState,
    CHAT_HISTORY_CONFIG,
    createNewChat,
    initialChatHistoryState
} from '../types/chatHistory';
import { Message } from '../types/state';

interface ChatHistoryContextType extends ChatHistoryState {
    currentChat: Chat | null;
    startNewChat: () => void;
    switchToChat: (chatId: string) => void;
    deleteChat: (chatId: string) => void;
    addMessageToCurrentChat: (message: Message) => void;
    updateCurrentChatTitle: (title: string) => void;
    clearAllChats: () => void;
}

const ChatHistoryContext = createContext<ChatHistoryContextType | null>(null);

/**
 * Load chat history from localStorage
 */
function loadChatHistory(): ChatHistoryState {
    try {
        const stored = localStorage.getItem(CHAT_HISTORY_CONFIG.storageKey);
        if (stored) {
            const parsed = JSON.parse(stored);
            console.log('[ChatHistory] Loaded', parsed.chats?.length || 0, 'chats');
            return parsed;
        }
    } catch (e) {
        console.error('[ChatHistory] Failed to load:', e);
    }
    return initialChatHistoryState;
}

/**
 * Save chat history to localStorage
 */
function saveChatHistory(state: ChatHistoryState): void {
    try {
        localStorage.setItem(CHAT_HISTORY_CONFIG.storageKey, JSON.stringify(state));
        console.log('[ChatHistory] Saved', state.chats.length, 'chats');
    } catch (e) {
        console.error('[ChatHistory] Failed to save:', e);
        // If quota exceeded, try to clean up old chats
        if ((e as any).name === 'QuotaExceededError') {
            const cleaned = {
                ...state,
                chats: state.chats.slice(0, Math.max(1, state.chats.length - 10))
            };
            try {
                localStorage.setItem(CHAT_HISTORY_CONFIG.storageKey, JSON.stringify(cleaned));
            } catch {
                console.error('[ChatHistory] Unable to save even after cleanup');
            }
        }
    }
}

export function ChatHistoryProvider({ children }: { children: ReactNode }) {
    const [state, setState] = useState<ChatHistoryState>(loadChatHistory);

    // Save to localStorage whenever state changes
    useEffect(() => {
        saveChatHistory(state);
    }, [state]);

    // Get current chat
    const currentChat = state.currentChatId
        ? state.chats.find(c => c.id === state.currentChatId) || null
        : null;

    // Start a new chat
    const startNewChat = useCallback(() => {
        setState(prev => {
            // If current chat is empty (no messages), don't create another empty one
            const currentChat = prev.currentChatId
                ? prev.chats.find(c => c.id === prev.currentChatId)
                : null;

            if (currentChat && currentChat.messages.length === 0) {
                console.log('[ChatHistory] Current chat is empty, not creating new one');
                return prev; // Keep the existing empty chat
            }

            // Create new chat
            const newChat = createNewChat();
            let chats = [newChat, ...prev.chats];

            // Enforce max chats limit
            if (chats.length > CHAT_HISTORY_CONFIG.maxChats) {
                chats = chats.slice(0, CHAT_HISTORY_CONFIG.maxChats);
            }

            console.log('[ChatHistory] Started new chat:', newChat.id);
            return {
                currentChatId: newChat.id,
                chats
            };
        });
    }, []);

    // Switch to existing chat
    const switchToChat = useCallback((chatId: string) => {
        setState(prev => {
            if (prev.chats.some(c => c.id === chatId)) {
                console.log('[ChatHistory] Switched to chat:', chatId);
                return { ...prev, currentChatId: chatId };
            }
            return prev;
        });
    }, []);

    // Delete a chat
    const deleteChat = useCallback((chatId: string) => {
        setState(prev => {
            const filtered = prev.chats.filter(c => c.id !== chatId);
            const newCurrentId = prev.currentChatId === chatId
                ? (filtered[0]?.id || null)
                : prev.currentChatId;
            console.log('[ChatHistory] Deleted chat:', chatId);
            return {
                currentChatId: newCurrentId,
                chats: filtered
            };
        });
    }, []);

    // Add message to current chat
    const addMessageToCurrentChat = useCallback((message: Message) => {
        setState(prev => {
            if (!prev.currentChatId) {
                // Create new chat if none exists
                const newChat = createNewChat();
                newChat.messages = [message];
                newChat.updatedAt = new Date().toISOString();
                return {
                    currentChatId: newChat.id,
                    chats: [newChat, ...prev.chats].slice(0, CHAT_HISTORY_CONFIG.maxChats)
                };
            }

            return {
                ...prev,
                chats: prev.chats.map(chat => {
                    if (chat.id === prev.currentChatId) {
                        const messages = [...chat.messages, message]
                            .slice(-CHAT_HISTORY_CONFIG.maxMessagesPerChat);
                        return {
                            ...chat,
                            messages,
                            updatedAt: new Date().toISOString()
                        };
                    }
                    return chat;
                })
            };
        });
    }, []);

    // Update current chat title
    const updateCurrentChatTitle = useCallback((title: string) => {
        setState(prev => {
            if (!prev.currentChatId) return prev;
            return {
                ...prev,
                chats: prev.chats.map(chat =>
                    chat.id === prev.currentChatId
                        ? { ...chat, title }
                        : chat
                )
            };
        });
    }, []);

    // Clear all chats
    const clearAllChats = useCallback(() => {
        setState(initialChatHistoryState);
        console.log('[ChatHistory] Cleared all chats');
    }, []);

    const value: ChatHistoryContextType = {
        ...state,
        currentChat,
        startNewChat,
        switchToChat,
        deleteChat,
        addMessageToCurrentChat,
        updateCurrentChatTitle,
        clearAllChats
    };

    return (
        <ChatHistoryContext.Provider value={value}>
            {children}
        </ChatHistoryContext.Provider>
    );
}

export function useChatHistory() {
    const context = useContext(ChatHistoryContext);
    if (!context) {
        throw new Error('useChatHistory must be used within ChatHistoryProvider');
    }
    return context;
}
