/**
 * Chat Title Generation Service
 * AI-powered generation of chat titles from conversation context
 */

import { Message } from '../types/state';
import { callGeminiApi } from './gemini/client';

/**
 * Generate a descriptive title for a chat based on its messages
 */
export async function generateChatTitle(
    messages: Message[],
    apiKey: string,
    model: string = 'gemini-2.0-flash'
): Promise<string> {
    // Only use first 3 messages for context
    const contextMessages = messages.slice(0, 3);

    if (contextMessages.length === 0) {
        return getDefaultTitle();
    }

    const conversationContext = contextMessages
        .map(m => `${m.role === 'user' ? 'User' : 'Assistant'}: ${m.content.slice(0, 200)}`)
        .join('\n\n');

    const systemPrompt = 'You are a helpful assistant that generates short, descriptive titles.';

    const userPrompt = `Based on this legal contract conversation, generate a brief, descriptive title (max 5 words) that captures the main topic or action. Return ONLY the title, no quotes, no explanation.

Conversation:
${conversationContext}

Title:`;

    try {
        console.log('[generateChatTitle] Generating title for', contextMessages.length, 'messages');

        const response = await callGeminiApi(apiKey, model, systemPrompt, userPrompt);

        // Clean up the response
        const title = response
            .replace(/['"]/g, '')       // Remove quotes
            .replace(/^Title:\s*/i, '') // Remove "Title:" prefix if present
            .trim()
            .slice(0, 50);              // Max 50 chars

        if (title && title.length >= 3) {
            console.log('[generateChatTitle] Generated:', title);
            return title;
        }

        return getDefaultTitle();
    } catch (error) {
        console.error('[generateChatTitle] Error:', error);
        return getDefaultTitle();
    }
}

/**
 * Generate a fallback title from first user message
 */
export function getTitleFromFirstMessage(messages: Message[]): string {
    const firstUserMessage = messages.find(m => m.role === 'user');
    if (!firstUserMessage) {
        return getDefaultTitle();
    }

    // Take first ~5 words
    const words = firstUserMessage.content.trim().split(/\s+/);
    const title = words.slice(0, 5).join(' ');

    if (title.length > 40) {
        return title.slice(0, 40) + '...';
    }

    return title + (words.length > 5 ? '...' : '');
}

/**
 * Default title with timestamp
 */
function getDefaultTitle(): string {
    const now = new Date();
    return `Chat - ${now.toLocaleDateString()}`;
}
