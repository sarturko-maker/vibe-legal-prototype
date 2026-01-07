/**
 * Gemini API Client for Vibe Legal
 * Handles API calls to Google Generative AI
 * Also provides unified routing to Groq when provider is set
 */

import { AIRouterResponse, Operation, AIProvider } from '../../types';
import { logError } from '../../utils/logger';
import { callGroqApi, buildGroqMessages } from '../groq';

// API endpoint
const GEMINI_API_BASE = 'https://generativelanguage.googleapis.com/v1beta/models';

/**
 * Call Gemini API with a prompt and get response.
 */
export async function callGeminiApi(
    apiKey: string,
    model: string,
    systemPrompt: string,
    userPrompt: string
): Promise<string> {
    const url = `${GEMINI_API_BASE}/${model}:generateContent?key=${apiKey}`;

    const payload = {
        contents: [
            {
                role: 'user',
                parts: [{ text: systemPrompt + '\n\n' + userPrompt }]
            }
        ],
        generationConfig: {
            temperature: 0.2,
            maxOutputTokens: 8192  // Increased for complex multi-operation responses
        }
    };

    try {
        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Gemini API error ${response.status}: ${errorText}`);
        }

        const data = await response.json();

        // Extract text from response
        const text = data.candidates?.[0]?.content?.parts?.[0]?.text;
        if (!text) {
            throw new Error('No text in Gemini response');
        }

        return text;
    } catch (error: any) {
        logError('Gemini API call failed:', error.message);
        throw error;
    }
}

/**
 * Call Gemini router and parse operations.
 */
export async function callGeminiRouter(
    apiKey: string,
    model: string,
    systemPrompt: string,
    userPrompt: string
): Promise<AIRouterResponse> {
    const text = await callGeminiApi(apiKey, model, systemPrompt, userPrompt);
    return parseRouterResponse(text);
}

/**
 * Unified AI Router - routes to Gemini, Groq, or Mistral based on provider.
 * This is the main entry point for AI calls that need provider awareness.
 */
export async function callAIRouter(
    provider: AIProvider,
    apiKey: string,
    model: string,
    systemPrompt: string,
    userPrompt: string
): Promise<AIRouterResponse> {
    if (provider === 'groq' || provider === 'mistral') {
        // Both Groq and Mistral use OpenAI-compatible format
        const messages = buildGroqMessages(systemPrompt, userPrompt);
        const { callOpenAICompatibleApi } = await import('../groq');
        const text = await callOpenAICompatibleApi(provider, apiKey, model, messages);
        return parseRouterResponse(text);
    }

    // Default to Gemini (handles 'gemini' and 'claude' for now)
    return callGeminiRouter(apiKey, model, systemPrompt, userPrompt);
}

/**
 * Unified AI call that returns raw text - routes to Gemini, Groq, or Mistral.
 * Use this for document analysis, party detection, and other non-router calls.
 */
export async function callAIForText(
    provider: AIProvider,
    apiKey: string,
    model: string,
    systemPrompt: string,
    userPrompt: string
): Promise<string> {
    if (provider === 'groq' || provider === 'mistral') {
        // Both Groq and Mistral use OpenAI-compatible format
        const messages = buildGroqMessages(systemPrompt, userPrompt);
        const { callOpenAICompatibleApi } = await import('../groq');
        return callOpenAICompatibleApi(provider, apiKey, model, messages);
    }

    // Default to Gemini
    return callGeminiApi(apiKey, model, systemPrompt, userPrompt);
}

/**
 * Normalize operation field names to handle AI variations.
 * Fixes: AMENDSIMPLE → AMEND_SIMPLE, targetid → target_id, etc.
 */
function normalizeOperation(op: any): Operation {
    const normalized: any = {};

    // Normalize type field
    if (op.type) {
        let type = op.type.toUpperCase();
        // Fix missing underscores
        if (type === 'AMENDSIMPLE') type = 'AMEND_SIMPLE';
        if (type === 'INSERTAFTER') type = 'INSERT';
        normalized.type = type;
    }

    // Copy and normalize field names
    for (const key of Object.keys(op)) {
        const lowerKey = key.toLowerCase();

        // Map variations to canonical names
        if (lowerKey === 'targetid' || lowerKey === 'target_id') {
            normalized.target_id = op[key];
        } else if (lowerKey === 'insertafter' || lowerKey === 'insert_after') {
            normalized.insert_after = op[key];
        } else if (lowerKey === 'amendedtext' || lowerKey === 'amended_text') {
            normalized.amended_text = op[key];
        } else if (lowerKey === 'findtext' || lowerKey === 'find_text') {
            normalized.find_text = op[key];
        } else if (lowerKey === 'replacetext' || lowerKey === 'replace_text') {
            normalized.replace_text = op[key];
        } else if (key !== 'type') {
            // Copy other fields as-is
            normalized[key] = op[key];
        }
    }

    return normalized as Operation;
}

/**
 * Scan text to find the first complete JSON object.
 * Handles nested braces and ignores braces inside strings.
 */
function extractJsonFromString(text: string): string | null {
    let braceCount = 0;
    let jsonStart = -1;
    let inString = false;
    let escape = false;

    for (let i = 0; i < text.length; i++) {
        const char = text[i];

        if (inString) {
            if (escape) {
                escape = false;
            } else if (char === '\\') {
                escape = true;
            } else if (char === '"') {
                inString = false;
            }
            continue;
        }

        if (char === '"') {
            inString = true;
            continue;
        }

        if (char === '{') {
            if (braceCount === 0) jsonStart = i;
            braceCount++;
        } else if (char === '}') {
            braceCount--;
            if (braceCount === 0 && jsonStart !== -1) {
                return text.substring(jsonStart, i + 1);
            }
        }
    }
    return null;
}

/**
 * Sanitize JSON string to fix common LLM issues like unescaped newlines in strings.
 * Replaces literal newlines/tabs inside JSON string values with proper escape sequences.
 */
function sanitizeJsonString(text: string): string {
    // Replace literal newlines inside strings with \n escape
    // This regex finds content between quotes and escapes newlines/tabs within
    let result = '';
    let inString = false;
    let escaped = false;

    for (let i = 0; i < text.length; i++) {
        const char = text[i];

        if (escaped) {
            result += char;
            escaped = false;
            continue;
        }

        if (char === '\\') {
            result += char;
            escaped = true;
            continue;
        }

        if (char === '"') {
            inString = !inString;
            result += char;
            continue;
        }

        if (inString) {
            // Replace control characters with escape sequences
            if (char === '\n') {
                result += '\\n';
            } else if (char === '\r') {
                result += '\\r';
            } else if (char === '\t') {
                result += '\\t';
            } else {
                result += char;
            }
        } else {
            result += char;
        }
    }

    return result;
}

/**
 * Parse router response JSON.
 */
export function parseRouterResponse(text: string): AIRouterResponse {
    let jsonText = text.trim();

    // 0. Handle case where AI returns the JSON as a quoted string literal
    // Check if text starts with a quote and contains JSON
    if ((jsonText.startsWith('"') || jsonText.startsWith("'")) && jsonText.includes('"intent"')) {
        try {
            // Try to parse as a string literal first
            const unwrapped = JSON.parse(jsonText);
            if (typeof unwrapped === 'string') {
                jsonText = unwrapped;
            }
        } catch {
            // Not a valid string literal, continue with normal parsing
        }
    }

    // 1. Try to extract from markdown code block
    const codeBlockMatch = jsonText.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (codeBlockMatch) {
        jsonText = codeBlockMatch[1].trim();
    } else {
        // 2. If no code block, try to extract raw JSON object
        // This handles cases where AI adds text before/after the JSON
        const extracted = extractJsonFromString(jsonText);
        if (extracted) {
            jsonText = extracted;
        }
    }

    // 3. Sanitize JSON to fix unescaped control characters (common with Mistral)
    jsonText = sanitizeJsonString(jsonText);

    try {
        const parsed = JSON.parse(jsonText);

        // Normalize operations array
        let operations: Operation[] = [];
        if (Array.isArray(parsed.operations)) {
            operations = parsed.operations.map(normalizeOperation);
        } else if (parsed.operation) {
            operations = [normalizeOperation(parsed.operation)];
        }

        return {
            intent: parsed.intent || (operations.length > 0 ? 'MODIFY' : 'ANSWER'),
            operations,
            answer: parsed.answer,
            explanation: parsed.explanation || parsed.description
        };
    } catch (error: any) {
        logError('Failed to parse router response:', error.message);
        logError('Raw response (first 500 chars):', text.substring(0, 500));

        // Final fallback: check if we can scrape a "MODIFY" intent from the raw text
        // even if JSON parsing failed completely
        if (text.includes('"intent": "MODIFY"') || text.includes('"intent": "modify"')) {
            // We can't safely proceed with operations if we can't parse them, but we can log it
            logError('Found intent: MODIFY in raw text, but JSON parse failed.');
        }

        return {
            intent: 'ANSWER',
            operations: [],
            answer: text, // Fallback: treat raw text as answer
            error: 'Failed to parse AI response: ' + error.message
        };
    }
}
