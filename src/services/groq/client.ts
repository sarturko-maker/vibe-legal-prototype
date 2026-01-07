/**
 * OpenAI-Compatible API Client for Vibe Legal
 * Handles API calls to Groq and Mistral (both use OpenAI-compatible format)
 */

import { logError } from '../../utils/logger';
import { AIProvider } from '../../types';

// API endpoints
const API_BASES: Record<string, string> = {
    groq: 'https://api.groq.com/openai/v1',
    mistral: 'https://api.mistral.ai/v1'
};

/**
 * Get the base URL for a provider
 */
function getBaseUrl(provider: AIProvider): string {
    return API_BASES[provider] || API_BASES.groq;
}

/**
 * Message format for OpenAI-compatible APIs
 */
export interface OpenAIMessage {
    role: 'system' | 'user' | 'assistant';
    content: string;
}

// Keep legacy export name for backwards compatibility
export type GroqMessage = OpenAIMessage;

/**
 * Call OpenAI-compatible API with messages and get response.
 * Works with Groq and Mistral.
 */
export async function callOpenAICompatibleApi(
    provider: AIProvider,
    apiKey: string,
    model: string,
    messages: OpenAIMessage[]
): Promise<string> {
    const baseUrl = getBaseUrl(provider);
    const url = `${baseUrl}/chat/completions`;

    const payload = {
        model: model,
        messages: messages,
        temperature: 0.1,
        max_tokens: 8000
    };

    try {
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${apiKey}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            const errorData = await response.json();
            const errorMessage = errorData.error?.message || `HTTP ${response.status}`;
            throw new Error(`${provider} API error: ${errorMessage}`);
        }

        const data = await response.json();

        // Extract text from OpenAI format response
        const text = data.choices?.[0]?.message?.content;
        if (!text) {
            throw new Error(`No content in ${provider} response`);
        }

        return text;
    } catch (error: any) {
        logError(`${provider} API call failed:`, error.message);
        throw error;
    }
}

// Legacy function name for backwards compatibility
export async function callGroqApi(
    apiKey: string,
    model: string,
    messages: OpenAIMessage[]
): Promise<string> {
    return callOpenAICompatibleApi('groq', apiKey, model, messages);
}

/**
 * Call OpenAI-compatible API and parse JSON response.
 * Handles markdown code blocks and cleanup.
 */
export async function callOpenAICompatibleForJSON(
    provider: AIProvider,
    apiKey: string,
    model: string,
    prompt: string
): Promise<any> {
    // Add JSON instruction to prompt
    const jsonPrompt = prompt + '\n\nRespond with valid JSON only. No markdown, no explanation.';

    const messages: OpenAIMessage[] = [
        { role: 'user', content: jsonPrompt }
    ];

    const text = await callOpenAICompatibleApi(provider, apiKey, model, messages);

    // Clean up response - strip markdown code blocks if present
    let cleanJson = text.trim();
    if (cleanJson.startsWith('```json')) {
        cleanJson = cleanJson.slice(7);
    }
    if (cleanJson.startsWith('```')) {
        cleanJson = cleanJson.slice(3);
    }
    if (cleanJson.endsWith('```')) {
        cleanJson = cleanJson.slice(0, -3);
    }

    try {
        return JSON.parse(cleanJson.trim());
    } catch (error: any) {
        logError(`Failed to parse ${provider} JSON response:`, error.message);
        logError('Raw response:', text.substring(0, 500));
        throw new Error(`Failed to parse JSON from ${provider} response`);
    }
}

// Legacy function name
export async function callGroqForJSON(
    apiKey: string,
    model: string,
    prompt: string
): Promise<any> {
    return callOpenAICompatibleForJSON('groq', apiKey, model, prompt);
}

/**
 * Fetch available models from OpenAI-compatible API.
 * Works with Groq and Mistral.
 */
export async function fetchOpenAICompatibleModels(
    provider: AIProvider,
    apiKey: string
): Promise<string[]> {
    const baseUrl = getBaseUrl(provider);
    const url = `${baseUrl}/models`;

    try {
        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${apiKey}`
            }
        });

        if (!response.ok) {
            const errorData = await response.json();
            const errorMessage = errorData.error?.message || `HTTP ${response.status}`;
            throw new Error(`Failed to fetch models: ${errorMessage}`);
        }

        const data = await response.json();

        // Extract model IDs from response
        const models = (data.data || [])
            .map((m: any) => m.id)
            .filter((id: string) => !!id)
            .sort();

        return models;
    } catch (error: any) {
        logError(`Failed to fetch ${provider} models:`, error.message);
        throw error;
    }
}

// Legacy function name  
export async function fetchGroqModels(apiKey: string): Promise<string[]> {
    return fetchOpenAICompatibleModels('groq', apiKey);
}

/**
 * Build OpenAI-format messages from system prompt, user prompt, and chat history.
 */
export function buildGroqMessages(
    systemPrompt: string,
    userPrompt: string,
    chatHistory?: { role: 'user' | 'bot'; content: string }[]
): OpenAIMessage[] {
    const messages: OpenAIMessage[] = [];

    // System message
    if (systemPrompt) {
        messages.push({ role: 'system', content: systemPrompt });
    }

    // Chat history
    if (chatHistory && chatHistory.length > 0) {
        for (const msg of chatHistory) {
            messages.push({
                role: msg.role === 'user' ? 'user' : 'assistant',
                content: msg.content
            });
        }
    }

    // Current user message
    messages.push({ role: 'user', content: userPrompt });

    return messages;
}

// Alias for clarity
export const buildOpenAIMessages = buildGroqMessages;
