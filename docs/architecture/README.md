# Architecture Decision Records

This directory contains Architecture Decision Records (ADRs) documenting key technical decisions in Vibe Legal.

## What is an ADR?

An ADR captures a significant architectural decision along with its context and consequences. They help future contributors understand *why* the codebase is structured the way it is.

## ADR Index

| ADR | Title | Status |
|-----|-------|--------|
| ADR-001 | Serverless BYOK Architecture | Reserved |
| ADR-002 | OOXML Manipulation Strategy | Reserved |
| ADR-003 | Track Changes Implementation | Reserved |
| ADR-004 | Paragraph Identification | Reserved |
| ADR-005 | Style Preservation | Reserved |
| ADR-006 | Sides Feature | Reserved |
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
