/**
 * AI Self-Validator Service
 * Shows the AI the calculated diff and asks if it's correct.
 * AI can self-correct if it made unintended changes.
 */

import { callGeminiApi } from './gemini/client';
import { formatDiffForAI, TextChange } from '../utils/diffFormatter';

interface ValidationParams {
    originalText: string;
    userInstruction: string;
    changeDescription: string;
    amendedText: string;
    calculatedDiff: TextChange[];
    apiKey: string;
    model: string;
}

export interface ValidationResult {
    approved: boolean;
    feedback?: string;
    corrected_text?: string;
}

/**
 * Validate changes with AI self-review.
 * Shows the AI what track changes the lawyer will see and asks if that's correct.
 */
export async function validateChangesWithAI(params: ValidationParams): Promise<ValidationResult> {
    const {
        originalText,
        userInstruction,
        changeDescription,
        amendedText,
        calculatedDiff,
        apiKey,
        model
    } = params;

    const diffPreview = formatDiffForAI(calculatedDiff);

    const validationPrompt = `You previously made changes to a legal document paragraph.
Now you need to validate that the track changes are correct.

ORIGINAL TEXT:
"""
${originalText}
"""

USER'S INSTRUCTION:
${userInstruction}

WHAT YOU SAID YOU'D CHANGE:
${changeDescription}

YOUR AMENDED TEXT:
"""
${amendedText}
"""

HERE'S WHAT WILL APPEAR AS TRACK CHANGES TO THE LAWYER:
${diffPreview}

VALIDATION QUESTION:
Look at the track changes above. Did you intend to make ALL of those changes?

Common mistakes to check for:
- Changed punctuation unintentionally (commas, semicolons, colons)
- Added/removed small words unintentionally (be, to, or, and, the)
- Adjusted grammar around your main change when you shouldn't have

Respond ONLY with a JSON object:

If everything is correct:
{"approved": true}

If you made unintended changes:
{
  "approved": false,
  "feedback": "I accidentally changed X to Y which wasn't requested",
  "corrected_text": "[full corrected paragraph text here]"
}

RESPOND ONLY WITH THE JSON OBJECT, NO OTHER TEXT.`;

    console.log('[aiSelfValidator] Validating changes...');
    console.log('[aiSelfValidator] Diff preview:', diffPreview);

    const response = await callGeminiApi(apiKey, model, '', validationPrompt);

    try {
        // Clean response (handle markdown code blocks)
        let jsonText = response.trim();
        const codeBlockMatch = jsonText.match(/```(?:json)?\s*([\s\S]*?)```/);
        if (codeBlockMatch) {
            jsonText = codeBlockMatch[1].trim();
        }

        const result: ValidationResult = JSON.parse(jsonText);

        if (result.approved) {
            console.log('[aiSelfValidator] ✓ AI approved all changes');
        } else {
            console.log('[aiSelfValidator] ✗ AI self-correction triggered');
            console.log('[aiSelfValidator] Feedback:', result.feedback);
        }

        return result;
    } catch (error: any) {
        console.warn('[aiSelfValidator] Failed to parse validation response:', error.message);
        console.warn('[aiSelfValidator] Raw response:', response.substring(0, 200));
        // Default to approved if we can't parse - don't block the operation
        return { approved: true };
    }
}
