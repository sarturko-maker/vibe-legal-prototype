/**
 * Text Diff Utility - Deterministic minimal change detection
 * 
 * Uses diff-match-patch for character-level diffs with semantic cleanup.
 * This ensures "4.1" and "Word." are treated correctly as units,
 * avoiding the punctuation-splitting issues of word-level diffs.
 */

import { diff_match_patch } from 'diff-match-patch';

export interface MinimalChange {
    find_text: string;
    replace_text: string;
}

/**
 * Normalize spacing to fix artifacts like "losses ." -> "losses."
 * and "4.2 A ll" -> "4.2 All" (split words)
 */
function normalizeSpacing(text: string): string {
    return text
        // Remove spaces BEFORE punctuation
        .replace(/ \./g, '.')
        .replace(/ ,/g, ',')
        .replace(/ ;/g, ';')
        .replace(/ :/g, ':')
        .replace(/ \?/g, '?')
        .replace(/ !/g, '!')
        .replace(/ \)/g, ')')
        .replace(/\( /g, '(')
        // Fix split words like "A ll" -> "All" (single letter + space + lowercase)
        .replace(/\b([A-Za-z]) ([a-z]+)\b/g, '$1$2')
        // Collapse multiple spaces
        .replace(/  +/g, ' ')
        .trim();
}

/**
 * Find ALL minimal changes between original and amended text.
 * Uses Google's diff-match-patch to find semantic changes.
 */
export function findMinimalChanges(original: string, amended: string): MinimalChange[] {
    // Normalize whitespace but preserve characters
    // No need to protect decimals/punctuation as we use char-diff
    const normOriginal = original.trim().replace(/\s+/g, ' ');
    const normAmended = amended.trim().replace(/\s+/g, ' ');

    if (normOriginal === normAmended) {
        console.log('[textDiff] Texts are identical');
        return [];
    }

    const dmp = new diff_match_patch();

    // Get character-level diff
    const diffs = dmp.diff_main(normOriginal, normAmended);

    // Semantic cleanup handles "Buyer." vs "Buyer ." grouping naturally
    dmp.diff_cleanupSemantic(diffs);

    const changes: MinimalChange[] = [];
    let position = 0; // Cursor in original text
    let i = 0;

    while (i < diffs.length) {
        const [op, text] = diffs[i];

        if (op === 0) {
            // Equal - advance position
            position += text.length;
            i++;
        } else {
            // Found a change (Insertion, Deletion, or consecutive mix)
            let deleted = '';
            let inserted = '';

            // Collect consecutive edits
            // -1 = Delete, 1 = Insert, 0 = Equal
            while (i < diffs.length && diffs[i][0] !== 0) {
                const [currentOp, currentText] = diffs[i];
                if (currentOp === -1) {
                    deleted += currentText;
                } else if (currentOp === 1) {
                    inserted += currentText;
                }
                i++;
            }

            // Get Context
            // We need 1 word before and 1 word after to anchor the find/replace

            // Context Before
            // Look back safely from current position
            const lookbackSize = 30; // Scan back 30 chars to find a word boundary
            const contextBeforeStub = normOriginal.substring(Math.max(0, position - lookbackSize), position);
            const wordsBefore = contextBeforeStub.trim().split(/\s+/);
            const contextBefore = wordsBefore.length > 0 ? wordsBefore[wordsBefore.length - 1] : '';

            // Context After
            // Look forward after the deleted segment
            const lookforwardSize = 30;
            const contextAfterStub = normOriginal.substring(position + deleted.length, position + deleted.length + lookforwardSize);
            const wordsAfter = contextAfterStub.trim().split(/\s+/);
            const contextAfter = wordsAfter.length > 0 ? wordsAfter[0] : '';

            // Build find/replace strings with context
            // Add space around context if it acts as a separate word
            let find_text = deleted;
            let replace_text = inserted;

            if (contextBefore) {
                find_text = contextBefore + ' ' + find_text;
                replace_text = contextBefore + ' ' + replace_text;
            }

            if (contextAfter) {
                find_text = find_text + ' ' + contextAfter;
                replace_text = replace_text + ' ' + contextAfter;
            }

            // Cleanup whitespace
            find_text = find_text.replace(/\s+/g, ' ').trim();
            replace_text = replace_text.replace(/\s+/g, ' ').trim();

            // Add to changes if there's an actual difference
            if (find_text !== replace_text) {
                changes.push({ find_text, replace_text });
            }

            // Advance position by what was consumed from original
            position += deleted.length;
        }
    }

    console.log('[textDiff] Found', changes.length, 'changes (diff-match-patch)');

    // Final pass: Normalize punctuation and spacing
    const normalizedChanges = changes.map((c, idx) => {
        const find = normalizeSpacing(c.find_text);
        const replace = normalizeSpacing(c.replace_text);

        if (c.find_text !== find) {
            console.log(`[textDiff]   [${idx}] Normalized find: "${c.find_text}" → "${find}"`);
        }
        if (c.replace_text !== replace) {
            console.log(`[textDiff]   [${idx}] Normalized replace: "${c.replace_text}" → "${replace}"`);
        }

        console.log(`[textDiff]   [${idx}] Final: "${find}" → "${replace}"`);
        return { find_text: find, replace_text: replace };
    });

    return normalizedChanges;
}

/**
 * Legacy support
 */
export function findMinimalChange(original: string, amended: string): MinimalChange | null {
    const changes = findMinimalChanges(original, amended);
    return changes.length > 0 ? changes[0] : null;
}
