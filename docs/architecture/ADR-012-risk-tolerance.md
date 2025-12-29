# ADR-012: Risk Tolerance Feature

**Status:** Accepted  
**Date:** 2025-12-28  
**Authors:** Artur (Vibe Legal)  
**Depends On:** Sides Feature (must be enabled first)

---

## Context

Legal negotiations operate within a **mandate range** — lawyers have instructions on ideal terms (best position) and acceptable terms (fallback position). Users need to calibrate AI suggestions along this spectrum.

The key insight is that users should provide **EXAMPLES** of their risk appetite, and the AI should **EXTRAPOLATE** that appetite to ALL commercial terms — whether mentioned or not.

---

## Decision

Implement Risk Tolerance with:

1. **Slider (0-100%)** from Conservative to Aggressive
2. **Two text fields** for example best/fallback positions
3. **Dependency on Sides** — cannot enable without a side selected
4. **Extrapolation framework** for AI to apply risk appetite universally

---

## Key Files

| File | Purpose |
|------|---------|
| [state.ts](file:///home/sarturko/vibe-legal-react/src/types/state.ts) | RiskTolerance interface and helpers |
| [RiskTolerancePanel.tsx](file:///home/sarturko/vibe-legal-react/src/components/RiskTolerancePanel.tsx) | UI component (full-pane modal) |
| [riskTolerancePrompt.ts](file:///home/sarturko/vibe-legal-react/src/prompts/riskTolerancePrompt.ts) | Prompt builder with extrapolation framework |
| [handleAction.ts](file:///home/sarturko/vibe-legal-react/src/services/handleAction.ts) | Integration point (lines 123-130) |

---

## Key Principle: Extrapolation

Users provide EXAMPLES, not exhaustive lists. The AI must:

1. **Infer overall risk appetite** from the examples
2. **Apply that appetite to ALL commercial terms**
3. **Use a standard calibration framework** (table mapping risk levels to term positions)
4. **Apply calibration SILENTLY** — no percentages or paragraph IDs in responses

---

## Risk Calibration Table

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

---

## User Interface

```
┌─────────────────────────────────────────────────┐
│  RISK TOLERANCE                     [Clear][Close]│
├─────────────────────────────────────────────────┤
│  Acting for: Buyer                              │
│                                                 │
│  Conservative ●━━━━━━━━━━━━━━━━ Aggressive      │
│                     35%                         │
│               [Conservative-Moderate]           │
│                                                 │
│  ┌─────────────────┐  ┌─────────────────┐      │
│  │ BEST POSITION   │  │ FALLBACK        │      │
│  │ (examples)      │  │ (examples)      │      │
│  │ [textarea]      │  │ [textarea]      │      │
│  └─────────────────┘  └─────────────────┘      │
│                                                 │
│  💡 AI will extrapolate to ALL terms            │
├─────────────────────────────────────────────────┤
│                     [Cancel]  [Apply]           │
└─────────────────────────────────────────────────┘
```

---

## Example Scenario: Limitation of Liability

**Best Position (Buyer):**
> Liability capped at 10% of contract value. No consequential damages. Remedies limited to repair, replace, or refund.

**Fallback Position (Buyer):**
> Liability up to 150% of contract value. Consequentials capped at 50%. Carve-outs for fraud, willful misconduct, IP infringement.

**User instruction:** "Amend the limitation of liability clause"

| Risk Level | AI Drafts |
|------------|-----------|
| 15% | "...liability shall not exceed 10% of the Purchase Price. In no event shall either party be liable for any indirect, consequential, special, incidental, or punitive damages..." |
| 50% | "...aggregate liability shall not exceed 100% of the Purchase Price. Neither party shall be liable for indirect or consequential damages, except in cases of fraud, willful misconduct, or breach of confidentiality..." |
| 85% | "...aggregate liability shall not exceed 150% of the Purchase Price. Consequential damages shall be capped at 50% of the Purchase Price. This limitation shall not apply to liability arising from (i) fraud or willful misconduct, (ii) breach of confidentiality, or (iii) infringement of intellectual property rights..." |

---

## Consequences

### Positive

- Mirrors real legal instructions (best/fallback mandate)
- AI responses calibrated to actual client appetite
- Reduces iteration on positioning
- Works universally across all commercial terms

### Negative

- Requires upfront thought from user
- Extrapolation may occasionally miss nuance for specific terms

### Mitigations

- Clear UI guidance with example placeholders
- User can adjust mid-session
- AI explains its recommendations (without mentioning percentages)
- Terms can be refined in follow-up messages

---

## Implementation Checklist

- [x] Create `RiskTolerancePanel` component
- [x] Add Risk button to toolbar (disabled until Sides enabled)
- [x] Implement state management for RiskTolerance (`state.ts`)
- [x] Add dependency logic (disable if Sides disabled)
- [x] Build risk tolerance prompt section (`riskTolerancePrompt.ts`)
- [x] Integrate into handleAction.ts with other enhancements
- [x] Add CSS styling matching modal pattern
- [x] Silent application (no percentages in AI output)

---

## Future Enhancements

1. **Per-clause risk levels** — Different tolerance for liability vs indemnity vs IP
2. **Position presets** — Template positions for common clause types
3. **Position suggestions** — AI suggests best/fallback based on document analysis
4. **Risk indicators in chat** — Show which position a suggestion aligns with

---

## References

- Sides Feature (required dependency)
- [ADR-007: Deal Context Feature](file:///home/sarturko/vibe-legal-react/docs/architecture/ADR-007-context-feature.md)
