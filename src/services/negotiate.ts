/**
 * Negotiate Service
 * AI-powered debate generation for contract negotiations
 */

import { DebateMessage } from '../types/negotiate';
import { DealContextState } from '../prompts/systemPrompt';
import { DetectedParties } from './documentAnalysis';
import { RiskTolerance } from '../types/state';
import { SideState } from '../components/SideSelector';
import { callAIForText } from './gemini/client';
import { AIProvider } from '../types';

export interface NegotiateContext {
    position: string;
    dealContext?: DealContextState | null;
    detectedParties?: DetectedParties | null;
    riskTolerance?: RiskTolerance | null;
    userSide: 'for' | 'against';
    selectedSide?: SideState;
}

function buildNegotiatePrompt(
    ctx: NegotiateContext,
    history: DebateMessage[],
    nextSide: 'for' | 'against'
): string {

    // Determine identity
    const pA = ctx.detectedParties?.partyA;
    const pB = ctx.detectedParties?.partyB;

    // User's identity (from global side selector)
    const userIsPartyA = ctx.selectedSide?.selected === 'partyA';
    const userName = userIsPartyA ? pA?.shortName : pB?.shortName;
    const userRole = userIsPartyA ? 'Party A (Buyer/Client)' : 'Party B (Seller/Provider)';
    const userDesc = userIsPartyA ? pA?.role : pB?.role;

    // Opponent's identity
    const oppName = userIsPartyA ? pB?.shortName : pA?.shortName;
    const oppRole = userIsPartyA ? 'Party B (Seller/Provider)' : 'Party A (Buyer/Client)';
    const oppDesc = userIsPartyA ? pB?.role : pA?.role;

    // Are we arguing FOR the user (User's Side)?
    const isArguingForUser = nextSide === ctx.userSide;

    // Current Speaker identity
    const speakerName = isArguingForUser ? (userName || 'Client') : (oppName || 'Counterparty');
    const speakerRoleLabel = isArguingForUser ? (userDesc || userRole) : (oppDesc || oppRole);

    let prompt = `You are a skilled commercial lawyer representing ${speakerName} (${speakerRoleLabel}).
You are debating a contract negotiation point.

=== PROPOSAL BY ${userName || 'User'} ===
"${ctx.position}"
`;

    if (isArguingForUser) {
        // === USER'S SIDE (Full Context) ===
        prompt += `
=== YOUR CONTEXT (PRIVILEGED) ===
You are representing the USER (${speakerRoleLabel}). You have access to the deal background and client preferences.

`;
        if (ctx.dealContext?.description?.trim()) {
            prompt += `DEAL BACKGROUND:\n${ctx.dealContext.description}\n\n`;
        }

        if (ctx.detectedParties) {
            prompt += `PARTIES:\nParty A: ${pA?.shortName} (${pA?.role})\nParty B: ${pB?.shortName} (${pB?.role})\n\n`;
        }

        if (ctx.riskTolerance?.enabled) {
            prompt += `CLIENT PREFERENCES:
Best Outcome: ${ctx.riskTolerance.bestPosition || 'Maximum protection'}
Fallback: ${ctx.riskTolerance.fallbackPosition || 'Market-standard'}
Risk Tolerance: ${ctx.riskTolerance.level}%
`;
        }

        prompt += `
INSTRUCTIONS:
- Argue IN FAVOR of the negotiation point (or justify why it's reasonable).
- Use your privileged context to explain WHY this matters to your client.
- Be persuasive but commercially reasonable.
`;

    } else {
        // === ADVERSARY SIDE (Market Practice Only) ===
        prompt += `
=== YOUR CONTEXT (OPPOSING COUNSEL) ===
You are representing the COUNTERPARTY (${speakerRoleLabel}). 
You DO NOT know the other side's private context. You operate based on MARKET PRACTICE and protecting your client's interests.

INSTRUCTIONS:
- PUSH BACK against the negotiation point.
- Argue that it is NOT market practice or is unreasonable.
- Protect your client (${speakerRoleLabel}) from risk.
- Do NOT concede easily.
`;
    }

    // Add history
    if (history.length > 0) {
        prompt += `
=== DEBATE SO FAR ===
${history.map(m => {
            // Map side to Party Name
            const isUserSide = m.side === ctx.userSide;
            const name = isUserSide ? (userName || 'side ' + m.side) : (oppName || 'side ' + m.side);
            return `${name?.toUpperCase() || 'SIDE ' + m.side.toUpperCase()}: "${m.headline}" - ${m.explanation}`;
        }).join('\n\n')}
`;
    }

    prompt += `
=== YOUR TASK ===
Generate the next argument for ${speakerName}.
${history.length > 0 ? 'Respond directly to the previous argument.' : 'Open the debate with a strong justification.'}

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
    nextSide: 'for' | 'against',
    model: string = 'gemini-2.0-flash',
    provider: AIProvider = 'gemini'
): Promise<{ headline: string; explanation: string }> {

    const prompt = buildNegotiatePrompt(ctx, history, nextSide);

    console.log('[generateDebateArgument] Side:', nextSide, 'History:', history.length, 'Provider:', provider);

    const systemPrompt = 'You are a skilled commercial lawyer generating debate arguments. Return valid JSON only.';
    const text = await callAIForText(provider, apiKey, model, systemPrompt, prompt);

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
    shouldStop: () => boolean,
    model: string = 'gemini-2.0-flash',
    provider: AIProvider = 'gemini'
): Promise<void> {

    const history: DebateMessage[] = [];
    let currentSide: 'for' | 'against' = ctx.userSide; // ALWAYS start with User's side

    // Limit to 2 rounds (4 turns total)
    const effectiveMaxRounds = 2;

    for (let round = 0; round < effectiveMaxRounds * 2; round++) {
        // Check if we should stop
        if (shouldStop()) {
            console.log('[runAutoDebate] Stopped by user');
            break;
        }

        try {
            const result = await generateDebateArgument(apiKey, ctx, history, currentSide, model, provider);

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
