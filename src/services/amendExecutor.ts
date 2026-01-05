/**
 * Amend Executor Service
 * Two-step AI execution: uses original text as read-only reference.
 */

import { callGeminiApi } from './gemini/client';

interface ExecuteChangeParams {
    originalMarkdown: string;
    changeDescription: string;
    userInstruction: string;
    apiKey: string;
    model: string;
}

/**
 * Execute an amend change using strict two-step AI approach.
 * The AI receives the original text as a read-only reference and must only make the specified change.
 */
export async function executeAmendChange(params: ExecuteChangeParams): Promise<string> {
    const { originalMarkdown, changeDescription, userInstruction, apiKey, model } = params;

    const prompt = `You are a precise text editor. Your job is to make ONLY the exact change requested.

ORIGINAL TEXT (REFERENCE - preserve character-for-character):
"""
${originalMarkdown}
"""

USER'S INSTRUCTION:
${userInstruction}

WHAT TO CHANGE:
${changeDescription}

CRITICAL INSTRUCTIONS:
1. The ORIGINAL TEXT above is your REFERENCE
2. Make ONLY the change specified
3. Preserve ALL other text character-for-character from ORIGINAL TEXT
4. Do NOT:
   - Fix grammar in unchanged portions
   - Adjust punctuation around the change
   - Add or remove conjunctions ("or", "and", "the")
   - Rephrase for clarity
   - Make any editorial improvements

5. Return ONLY the complete amended text, nothing else
6. No preamble, no explanation, no markdown formatting - just the text`;

    console.log('[amendExecutor] Executing change:', changeDescription.substring(0, 100));

    const response = await callGeminiApi(apiKey, model, '', prompt);
    const result = response.trim();

    console.log('[amendExecutor] Result length:', result.length);

    return result;
}
