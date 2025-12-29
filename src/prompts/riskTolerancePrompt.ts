/**
 * Risk Tolerance Prompt Builder
 * Builds AI prompt section for risk calibration with extrapolation framework.
 * ADR-012: Risk Tolerance Feature
 * 
 * IMPORTANT: AI applies risk tolerance SILENTLY - no percentages or paragraph IDs in responses.
 */

import { RiskTolerance, getRiskSublabel } from '../types/state';
import { SideState } from '../components/SideSelector';

/**
 * Build risk tolerance instruction for AI prompts.
 * Includes extrapolation framework for applying risk appetite to ALL commercial terms.
 */
export function buildRiskToleranceInstruction(
    side: SideState,
    riskTolerance: RiskTolerance
): string {
    if (!riskTolerance.enabled || side.selected === 'neutral') {
        return '';
    }

    const riskSublabel = getRiskSublabel(riskTolerance.level);

    return `
=== NEGOTIATION CALIBRATION (INTERNAL ONLY) ===

Risk Level: ${riskSublabel} (${riskTolerance.level}%)

BEST POSITION (client's ideal outcome):
${riskTolerance.bestPosition || '[Not specified - use maximum protection stance]'}

FALLBACK POSITION (acceptable to close deal):
${riskTolerance.fallbackPosition || '[Not specified - use market-standard terms]'}

---

## EXTRAPOLATION FRAMEWORK

The positions above are EXAMPLES. You must EXTRAPOLATE the user's risk appetite to ALL commercial terms.

### Calibration Table (Internal Reference)

| Term | Conservative (0-33%) | Moderate (34-66%) | Aggressive (67-100%) |
|------|---------------------|-------------------|----------------------|
| Liability caps | Low (10-50%) | Market (100%) | High (150-200%+) |
| Consequential damages | Excluded | Capped | Accepted with limits |
| Warranties | Extensive, long survival | Standard | Minimal |
| Indemnities | Narrow, capped | Balanced | Broad |
| Termination | Our right only | Mutual | Their right OK |
| Notice periods | Long (60-90 days) | Standard (30 days) | Short (14 days) OK |
| Governing law | Our jurisdiction | Neutral | Their jurisdiction OK |
| Assignment | Their consent required | Not unreasonably withheld | Free assignment OK |

### Current Calibration

${getRiskGuidance(riskTolerance.level)}

---

## RESPONSE STYLE (CRITICAL)

Apply the risk calibration SILENTLY. Your responses must:

1. **NEVER mention percentages** — Don't say "at your 35% risk level" or "given your conservative stance" or "applying your risk tolerance..."

2. **NEVER show paragraph IDs** — Refer to clauses by name ("the limitation of liability clause", "clause 5", "the warranty provisions"). NEVER use "[P24]", "paragraph 24", or any numeric paragraph references in your response to the user.

3. **Draft naturally** — Write as an experienced solicitor briefing a client. The calibration must be invisible to the reader.

4. **Be direct** — State your recommendation confidently. The calibration is already baked in.

5. **Don't explain extrapolation** — When applying to unmentioned terms, just do it. Don't say "extrapolating your risk appetite..."

GOOD RESPONSE EXAMPLE:
"The current liability cap at 100% of the Purchase Price exposes your client to significant risk. I recommend reducing this to 15% of the Purchase Price, with carve-outs limited to fraud and personal injury."

BAD RESPONSE EXAMPLE:
"At your conservative-moderate risk level (35%), the Limitation of Liability clause [P24] needs amendment. Applying your risk tolerance to this term, I recommend..."

---

## APPLICATION

Internally calibrate at ${riskTolerance.level}% on all terms:
- For terms in user's examples: follow their guidance
- For unmentioned terms: use the calibration table at ${riskTolerance.level}%
- Interpolate: ${riskTolerance.level}% of the way from Conservative to Aggressive

But in your response: write naturally, as if this is simply your professional recommendation.
`;
}

function getRiskGuidance(level: number): string {
    if (level <= 10) {
        return `At this level: Maximum protection. Draft at best position. Make no concessions.`;
    }
    if (level <= 25) {
        return `At this level: Strong protection. Stay close to best position. Minimal flexibility.`;
    }
    if (level <= 40) {
        return `At this level: Lean toward best position. Small concessions acceptable where reasonable.`;
    }
    if (level <= 60) {
        return `At this level: Balanced approach. Fair, market-standard terms for both parties.`;
    }
    if (level <= 75) {
        return `At this level: Lean toward fallback. Show flexibility. Accept market terms.`;
    }
    if (level <= 90) {
        return `At this level: Accept fallback readily. Deal completion is priority.`;
    }
    return `At this level: Accept fallback without resistance. Close the deal.`;
}
