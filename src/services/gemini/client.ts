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
            maxOutputTokens: 4096
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
 * Parse router response JSON.
 */
export function parseRouterResponse(text: string): AIRouterResponse {
    // Try to extract JSON from response
    let jsonText = text;

    // Handle markdown code blocks
    const codeBlockMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (codeBlockMatch) {
        jsonText = codeBlockMatch[1].trim();
    }

    try {
        const parsed = JSON.parse(jsonText);

        // Normalize operations array
        let operations: Operation[] = [];
        if (Array.isArray(parsed.operations)) {
            operations = parsed.operations;
        } else if (parsed.operation) {
            operations = [parsed.operation];
        }

        return {
            intent: parsed.intent || (operations.length > 0 ? 'MODIFY' : 'ANSWER'),
            operations,
            answer: parsed.answer,
            explanation: parsed.explanation || parsed.description
        };
    } catch (error: any) {
        logError('Failed to parse router response:', error.message);
        return {
            intent: 'ANSWER',
            operations: [],
            answer: text, // Fallback: treat raw text as answer
            error: 'Failed to parse AI response: ' + error.message
        };
    }
}
