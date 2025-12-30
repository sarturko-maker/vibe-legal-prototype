/**
 * Gemini API Client for Vibe Legal
 * Handles API calls to Google Generative AI
 */

import { AIRouterResponse, Operation } from '../../types';
import { logError } from '../../utils/logger';

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
 * Parse router response JSON.
 */
export function parseRouterResponse(text: string): AIRouterResponse {
    let jsonText = text;

    // 1. Try to extract from markdown code block
    const codeBlockMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (codeBlockMatch) {
        jsonText = codeBlockMatch[1].trim();
    } else {
        // 2. If no code block, try to extract raw JSON object
        // This handles cases where AI adds text before/after the JSON
        const extracted = extractJsonFromString(text);
        if (extracted) {
            jsonText = extracted;
        }
    }

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
