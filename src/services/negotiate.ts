/**
 * Negotiate Service
 * AI-powered debate generation for contract negotiations
 */

import { DebateMessage } from '../types/negotiate';
import { DealContextState } from '../prompts/systemPrompt';
import { DetectedParties } from './documentAnalysis';

export interface NegotiateContext {
    position: string;
    dealContext?: DealContextState | null;
    detectedParties?: DetectedParties | null;
}

function buildNegotiatePrompt(
    ctx: NegotiateContext,
    history: DebateMessage[],
    nextSide: 'for' | 'against'
): string {

    let prompt = `You are a skilled commercial lawyer in a contract negotiation.

=== NEGOTIATION POINT ===
"${ctx.position}"
`;

    // Add deal context if available
    if (ctx.dealContext?.description?.trim()) {
        prompt += `
=== DEAL CONTEXT ===
${ctx.dealContext.description}
`;
    }

    // Add party names if available
    if (ctx.detectedParties) {
        prompt += `
=== PARTIES ===
Party A: ${ctx.detectedParties.partyA.shortName} (${ctx.detectedParties.partyA.role})
Party B: ${ctx.detectedParties.partyB.shortName} (${ctx.detectedParties.partyB.role})
`;
    }

    // Add debate history
    if (history.length > 0) {
        prompt += `
=== DEBATE SO FAR ===
${history.map(m => `${m.side.toUpperCase()}: "${m.headline}" - ${m.explanation}`).join('\n\n')}
`;
    }

    prompt += `
=== YOUR TASK ===
Generate the next argument ${nextSide === 'for' ? 'IN FAVOUR OF' : 'AGAINST'} the negotiation point.

${history.length > 0 ? 'Your argument must DIRECTLY RESPOND to the previous argument.' : 'Start with a strong opening argument.'}

Be:
- Persuasive and specific
- Realistic (what a real lawyer would say)
- Concise but substantive

Respond with JSON ONLY (no markdown):
{
  "headline": "Short punchy headline (5-10 words)",
  "explanation": "Two sentences max explaining the argument."
}`;

    return prompt;
}

export async function generateDebateArgument(
    apiKey: string,
    ctx: NegotiateContext,
    history: DebateMessage[],
    nextSide: 'for' | 'against'
): Promise<{ headline: string; explanation: string }> {

    const prompt = buildNegotiatePrompt(ctx, history, nextSide);

    console.log('[generateDebateArgument] Side:', nextSide, 'History:', history.length);

    const response = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent?key=${apiKey}`,
        {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                contents: [{ parts: [{ text: prompt }] }],
                generationConfig: { temperature: 0.7 }
            })
        }
    );

    const data = await response.json();
    const text = data.candidates?.[0]?.content?.parts?.[0]?.text || '';

    // Parse JSON (strip markdown if present)
    const jsonStr = text.replace(/```json\n?|\n?```/g, '').trim();
    const result = JSON.parse(jsonStr);

    console.log('[generateDebateArgument] Result:', result.headline);

    return {
        headline: result.headline,
        explanation: result.explanation
    };
}

export async function runAutoDebate(
    apiKey: string,
    ctx: NegotiateContext,
    maxRounds: number,
    onMessage: (msg: DebateMessage) => void,
    shouldStop: () => boolean
): Promise<void> {

    const history: DebateMessage[] = [];
    let currentSide: 'for' | 'against' = 'for';

    for (let round = 0; round < maxRounds * 2; round++) {
        // Check if we should stop
        if (shouldStop()) {
            console.log('[runAutoDebate] Stopped by user');
            break;
        }

        try {
            const result = await generateDebateArgument(apiKey, ctx, history, currentSide);

            const msg: DebateMessage = {
                id: history.length + 1,
                side: currentSide,
                author: 'ai',
                headline: result.headline,
                explanation: result.explanation,
                timestamp: new Date()
            };

            history.push(msg);
            onMessage(msg);

            // Delay before next argument
            await new Promise(resolve => setTimeout(resolve, 800));

            // Switch sides
            currentSide = currentSide === 'for' ? 'against' : 'for';

        } catch (error) {
            console.error('[runAutoDebate] Error:', error);
            break;
        }
    }

    console.log('[runAutoDebate] Complete:', history.length, 'messages');
}
