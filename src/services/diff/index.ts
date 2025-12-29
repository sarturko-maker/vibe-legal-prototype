// src/services/diff/index.ts
// Diff service exports

export {
    computeSurgicalDiff,
    convertDmpToStructuredDiffs,
    optimizeStructuredDiffs,
    hasMeaningfulChanges,
    getDiffStats
} from './surgicalDiff';

export type { StructuredDiff, DiffType } from './surgicalDiff';

export {
    generateDiffPreviewHtml,
    generateChangeSummary
} from './previewGenerator';
