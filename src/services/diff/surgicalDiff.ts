/**
 * Surgical Diff Engine for Vibe Legal
 * Integrates diff_match_patch for precise text comparison
 * Source: vibe-legal-beta-0.2.yaml L2545-2612
 */

// Declare diff_match_patch (loaded via CDN)
declare var diff_match_patch: any;

// Structured diff types
export type DiffType = 'insert' | 'delete' | 'replace' | 'equal';

export interface StructuredDiff {
    type: DiffType;
    start: number;
    end?: number;
    text?: string;
}

/**
 * Compute diff between original and new text.
 * Returns array of structured diffs suitable for OOXML application.
 */
export function computeSurgicalDiff(originalText: string, newText: string): StructuredDiff[] {
    const dmp = new diff_match_patch();
    const dmpDiffs = dmp.diff_main(originalText, newText);
    dmp.diff_cleanupSemantic(dmpDiffs);

    const structuredDiffs = convertDmpToStructuredDiffs(dmpDiffs);
    return optimizeStructuredDiffs(structuredDiffs);
}

/**
 * Convert diff_match_patch output to StructuredDiff array.
 * DMP format: [[-1, "deleted"], [0, "equal"], [1, "inserted"]]
 */
export function convertDmpToStructuredDiffs(dmpDiffs: [number, string][]): StructuredDiff[] {
    const structuredDiffs: StructuredDiff[] = [];
    let position = 0;

    for (const diff of dmpDiffs) {
        const type = diff[0];
        const text = diff[1];

        if (type === 0) {
            // Equal - just advance position
            position += text.length;
        } else if (type === -1) {
            // Delete
            structuredDiffs.push({
                type: 'delete',
                start: position,
                end: position + text.length
            });
            position += text.length;
        } else if (type === 1) {
            // Insert - don't advance position (inserted text isn't in original)
            structuredDiffs.push({
                type: 'insert',
                start: position,
                text: text
            });
        }
    }

    return structuredDiffs;
}

/**
 * Combine adjacent delete+insert at same position into replace.
 * This reduces the number of track changes in the document.
 */
export function optimizeStructuredDiffs(diffs: StructuredDiff[]): StructuredDiff[] {
    const optimized: StructuredDiff[] = [];
    let i = 0;

    while (i < diffs.length) {
        const current = diffs[i];

        // Check if this delete is followed by an insert at same position
        if (current.type === 'delete' && i + 1 < diffs.length) {
            const next = diffs[i + 1];
            if (next.type === 'insert' && next.start === current.start) {
                // Combine into replace
                optimized.push({
                    type: 'replace',
                    start: current.start,
                    end: current.end,
                    text: next.text
                });
                i += 2;
                continue;
            }
        }

        optimized.push(current);
        i++;
    }

    return optimized;
}

/**
 * Check if diffs represent a meaningful change (not just whitespace).
 */
export function hasMeaningfulChanges(diffs: StructuredDiff[]): boolean {
    return diffs.some(d => {
        if (d.type === 'equal') return false;
        const text = d.text || '';
        return text.trim().length > 0;
    });
}

/**
 * Get change statistics from diffs.
 */
export function getDiffStats(diffs: StructuredDiff[]): { insertions: number; deletions: number; replacements: number } {
    return diffs.reduce((acc, diff) => {
        switch (diff.type) {
            case 'insert': acc.insertions++; break;
            case 'delete': acc.deletions++; break;
            case 'replace': acc.replacements++; break;
        }
        return acc;
    }, { insertions: 0, deletions: 0, replacements: 0 });
}
