/**
 * Chat History Service for Vibe Legal
 * Manages chat sessions in localStorage
 */

export interface ChatSession {
    id: string;
    name: string;
    messages: Array<{ role: string; content: string }>;
    createdAt: string;
    updatedAt: string;
}

const MAX_SAVED_CHATS = 20;
const STORAGE_KEY = 'vibe-legal-chat-history';

/**
 * Generate a unique chat ID
 */
export function generateChatId(): string {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

/**
 * Save a chat session to localStorage
 */
export function saveChat(chat: ChatSession): void {
    const history = getHistory();
    const idx = history.findIndex(c => c.id === chat.id);

    if (idx >= 0) {
        // Update existing chat
        history[idx] = { ...chat, updatedAt: new Date().toISOString() };
    } else {
        // Add new chat at beginning
        history.unshift(chat);
    }

    // Keep only MAX_SAVED_CHATS
    localStorage.setItem(STORAGE_KEY, JSON.stringify(history.slice(0, MAX_SAVED_CHATS)));
}

/**
 * Get all chat sessions from localStorage
 */
export function getHistory(): ChatSession[] {
    try {
        return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    } catch {
        return [];
    }
}

/**
 * Load a specific chat by ID
 */
export function loadChat(id: string): ChatSession | null {
    return getHistory().find(c => c.id === id) || null;
}

/**
 * Delete a chat by ID
 */
export function deleteChat(id: string): void {
    const history = getHistory().filter(c => c.id !== id);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(history));
}

/**
 * Generate a chat name from the first user message
 */
export function generateChatName(messages: Array<{ role: string; content: string }>): string {
    const firstUserMsg = messages.find(m => m.role === 'user');
    if (!firstUserMsg) return 'New Chat';

    // Take first 30 chars, truncate at word boundary
    const text = firstUserMsg.content.substring(0, 40);
    const lastSpace = text.lastIndexOf(' ');
    const truncated = lastSpace > 20 ? text.substring(0, lastSpace) : text;

    return truncated + (firstUserMsg.content.length > 40 ? '...' : '');
}

/**
 * Create a new empty chat session
 */
export function createNewChat(): ChatSession {
    return {
        id: generateChatId(),
        name: 'New Chat',
        messages: [],
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
    };
}
