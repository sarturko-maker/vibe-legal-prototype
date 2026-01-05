# Architecture Decision Records

This directory contains Architecture Decision Records (ADRs) documenting key technical decisions in Vibe Legal.

## What is an ADR?

An ADR captures a significant architectural decision along with its context and consequences. They help future contributors understand *why* the codebase is structured the way it is.

## ADR Index

| ADR | Title | Status |
|-----|-------|--------|
| [ADR-001](ADR-001-serverless-byok.md) | Serverless BYOK Architecture | Accepted |
| [ADR-002](ADR-002-track-changes-strategy.md) | Track Changes Strategy | Accepted |
| [ADR-003](ADR-003-track-changes-implementation.md) | Track Changes Implementation | Accepted |
| [ADR-004](ADR-004-paragraph-identification.md) | Paragraph Identification | Accepted |
| [ADR-005](ADR-005-style-preservation.md) | Style Preservation | Accepted |
| [ADR-006](ADR-006-sides-feature.md) | Sides Feature | Accepted |
| [ADR-007](ADR-007-context-feature.md) | Deal Context Feature | Accepted |
| [ADR-008](ADR-008-definitions-feature.md) | Definitions Feature | Accepted |
| [ADR-009](ADR-009-negotiate-feature.md) | Negotiate Feature | Accepted |
| [ADR-010](ADR-010-mindmap-feature.md) | Mind Map Feature | Accepted |
| [ADR-011](ADR-011-edge-case-handling.md) | Edge Case Handling | Accepted |
| [ADR-012](ADR-012-risk-tolerance.md) | Risk Tolerance | Accepted |

## Creating New ADRs

1. Copy `ADR-TEMPLATE.md`
2. Rename to `ADR-XXX-short-title.md` (use next available number)
3. Fill in all sections
4. Update this README index
5. Reference ADR number in related code comments
