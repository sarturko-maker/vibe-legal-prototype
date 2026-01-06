/**
 * Chat History Types
 * Data structures for persistent chat management
 */

import { Message } from './state';

/**
 * A single chat conversation
 */
export interface Chat {
    id: string;
    title: string;
    messages: Message[];
    createdAt: string;  // ISO string for JSON serialization
    updatedAt: string;  // ISO string for JSON serialization
}

/**
 * State for chat history management
 */
export interface ChatHistoryState {
    currentChatId: string | null;
    chats: Chat[];
}

/**
 * Storage configuration
 */
export const CHAT_HISTORY_CONFIG = {
    storageKey: 'vibelegal_chat_history',
    maxChats: 50,
    maxMessagesPerChat: 100
};

/**
 * Create a new empty chat
 */
export function createNewChat(): Chat {
    const now = new Date().toISOString();
    return {
        id: generateChatId(),
        title: 'New Chat',
        messages: [],
        createdAt: now,
        updatedAt: now
    };
}

/**
 * Generate a unique chat ID
 */
function generateChatId(): string {
    return `chat_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * Format timestamp for display (e.g., "5m", "2h", "1d")
 */
export function formatTimeAgo(isoString: string): string {
    const date = new Date(isoString);
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();

    const minutes = Math.floor(diffMs / 60000);
    if (minutes < 60) return `${minutes}m`;

    const hours = Math.floor(minutes / 60);
    if (hours < 24) return `${hours}h`;

    const days = Math.floor(hours / 24);
    if (days < 30) return `${days}d`;

    const months = Math.floor(days / 30);
    return `${months}mo`;
}

/**
 * Default initial state
 */
export const initialChatHistoryState: ChatHistoryState = {
    currentChatId: null,
    chats: []
};
