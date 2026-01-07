/**
 * Party Detection Service
 * Uses AI to identify the two main parties in a contract
 */

import { callAIForText } from './gemini/client';
import { AIProvider } from '../types';

export interface DetectedParties {
    partyA: { shortName: string; fullName: string; role: string };
    partyB: { shortName: string; fullName: string; role: string };
}

/**
 * Detect the two main parties in a contract document
 */
export async function detectParties(
    documentText: string,
    apiKey: string,
    model: string = 'gemini-2.0-flash',
    provider: AIProvider = 'gemini'
): Promise<DetectedParties | null> {
    console.log('[detectParties] Starting...');
    console.log('[detectParties] API key present:', !!apiKey);
    console.log('[detectParties] Document length:', documentText?.length);

    if (!apiKey) {
        console.error('[detectParties] No API key provided');
        return null;
    }

    if (!documentText || documentText.length < 50) {
        console.error('[detectParties] Document text too short or missing');
        return null;
    }

    const systemPrompt = `You are a legal document parser. Extract party information from contracts. Return valid JSON only.`;

    const userPrompt = `Analyze this contract and identify the two main parties.

Return JSON only, no markdown code blocks, no explanation:
{
  "partyA": {
    "shortName": "ABC",
    "fullName": "ABC Holdings Limited",
    "role": "Buyer"
  },
  "partyB": {
    "shortName": "XYZ",
    "fullName": "XYZ International B.V.",
    "role": "Seller"
  }
}

Rules for shortName:
- If company name exists, use abbreviated version (max 15 chars)
- If no company name, use the role (e.g., "Buyer", "Seller")
- Max 15 characters

Contract text (first 3000 chars):
${documentText.substring(0, 3000)}`;

    try {
        console.log('[detectParties] Calling AI with provider:', provider, 'model:', model);
        const response = await callAIForText(provider, apiKey, model, systemPrompt, userPrompt);
        console.log('[detectParties] Raw response length:', response?.length);
        console.log('[detectParties] Raw response:', response?.substring(0, 500));

        // Extract JSON from response (handle potential markdown code blocks)
        let jsonStr = response;

        // Remove markdown code blocks if present
        const codeBlockMatch = response.match(/```(?:json)?\s*([\s\S]*?)```/);
        if (codeBlockMatch) {
            jsonStr = codeBlockMatch[1];
        }

        // Extract JSON object
        const jsonMatch = jsonStr.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
            const parsed = JSON.parse(jsonMatch[0]);

            // Validate structure
            if (parsed.partyA?.shortName && parsed.partyB?.shortName) {
                console.log('[detectParties] SUCCESS - Detected:', parsed.partyA.shortName, 'vs', parsed.partyB.shortName);
                return parsed as DetectedParties;
            } else {
                console.warn('[detectParties] Invalid structure:', parsed);
            }
        } else {
            console.warn('[detectParties] No JSON found in response');
        }

        return null;
    } catch (e) {
        console.error('[detectParties] Error:', e);
        return null;
    }
}
