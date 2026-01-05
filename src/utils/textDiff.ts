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
    // For pure insertions: instead of find/replace, we insert_text after insert_after
    insert_after?: string;
    insert_text?: string;
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
        // NOTE: Removed the "split word" fixer (.replace(/\b([A-Za-z]) ([a-z]+)\b/g, '$1$2'))
        // It was too aggressive and would join legitimate separate words like "A when" → "Awhen"
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

            // Get Context - use 1 word before and after for anchoring
            // IMPORTANT: Extract EXACT text from original, don't reconstruct from split words
            // (splitting on whitespace breaks possessives like "Seller's" → "Seller 's")

            // Context Before - find last word before the change
            const lookbackSize = 30;
            const beforeStart = Math.max(0, position - lookbackSize);
            const contextBeforeStub = normOriginal.substring(beforeStart, position);
            // Find the last complete word
            const lastSpaceBefore = contextBeforeStub.lastIndexOf(' ');
            const contextBefore = lastSpaceBefore >= 0
                ? contextBeforeStub.substring(lastSpaceBefore + 1)
                : contextBeforeStub;

            // Context After - find first word after the change
            const lookforwardSize = 30;
            const contextAfterStub = normOriginal.substring(position + deleted.length, position + deleted.length + lookforwardSize);
            // Find the first complete word
            const firstSpaceAfter = contextAfterStub.indexOf(' ');
            const contextAfter = firstSpaceAfter >= 0
                ? contextAfterStub.substring(0, firstSpaceAfter)
                : contextAfterStub;

            // SPECIAL CASE: Pure insertion (nothing deleted, just inserting new text)
            // Use insert_after/insert_text to avoid showing context as deleted
            if (deleted.length === 0 && inserted.length > 0 && contextBefore) {
                // For pure insertion, find the anchor text and insert after it
                // Anchor should include nearby punctuation for uniqueness
                let anchor = contextBefore;
                // If there's punctuation immediately after the insertion point, include it
                const charAfter = normOriginal.charAt(position);
                if (charAfter && /[;:,.\)]/.test(charAfter)) {
                    anchor += charAfter;
                }

                changes.push({
                    find_text: anchor,
                    replace_text: anchor + inserted.trimStart(), // Insert text after anchor
                    insert_after: anchor,
                    insert_text: inserted.trimStart()
                });

                position += deleted.length;
                continue; // Skip the normal processing
            }

            // Build find/replace strings with context
            // CRITICAL: Check if we need a space between context and content
            // Don't add space before apostrophe/punctuation or if text naturally joins
            let find_text = deleted;
            let replace_text = inserted;


            if (contextBefore) {
                // Check if deleted starts with punctuation (like apostrophe) - no space needed
                const needsSpaceBeforeDeleted = deleted.length > 0 && /^[a-zA-Z0-9]/.test(deleted.charAt(0));
                const needsSpaceBeforeInserted = inserted.length > 0 && /^[a-zA-Z0-9]/.test(inserted.charAt(0));

                if (needsSpaceBeforeDeleted) {
                    find_text = contextBefore + ' ' + find_text;
                } else {
                    find_text = contextBefore + find_text;
                }

                if (needsSpaceBeforeInserted) {
                    replace_text = contextBefore + ' ' + replace_text;
                } else {
                    replace_text = contextBefore + replace_text;
                }
            }

            if (contextAfter) {
                // Check if content ends with punctuation - no space needed after
                const findEndsWithPunct = find_text.length > 0 && /[^a-zA-Z0-9]$/.test(find_text);
                const replaceEndsWithPunct = replace_text.length > 0 && /[^a-zA-Z0-9]$/.test(replace_text);

                if (findEndsWithPunct || /^[^a-zA-Z0-9]/.test(contextAfter)) {
                    find_text = find_text + contextAfter;
                } else {
                    find_text = find_text + ' ' + contextAfter;
                }

                if (replaceEndsWithPunct || /^[^a-zA-Z0-9]/.test(contextAfter)) {
                    replace_text = replace_text + contextAfter;
                } else {
                    replace_text = replace_text + ' ' + contextAfter;
                }
            }

            // Cleanup whitespace - normalize multiple spaces and trim
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

    // ==========================================================================
    // PASS 1: Filter BOUNDARY OVERLAPS where one change ends with text another begins with
    // Instead of merging (which caused corruption), we SKIP the shorter overlapping change
    // Example: "...remedy shall" and "shall be," both reference the same "shall"
    // We keep the longer one which contains more context
    // ==========================================================================
    const filteredChanges: MinimalChange[] = [];
    const skippedIndices = new Set<number>();

    for (let i = 0; i < normalizedChanges.length; i++) {
        if (skippedIndices.has(i)) continue;

        const current = normalizedChanges[i];
        let shouldSkip = false;

        // Check if this change overlaps with any other change
        for (let j = 0; j < normalizedChanges.length; j++) {
            if (j === i || skippedIndices.has(j)) continue;
            const other = normalizedChanges[j];

            // Check if current's find_text ends with text that other's find_text starts with
            // or vice versa
            let overlapLen = 0;
            const maxOverlap = Math.min(current.find_text.length, other.find_text.length);

            for (let len = 3; len <= maxOverlap; len++) {
                if (current.find_text.slice(-len) === other.find_text.slice(0, len)) {
                    overlapLen = len;
                }
                if (other.find_text.slice(-len) === current.find_text.slice(0, len)) {
                    overlapLen = len;
                }
            }

            // If significant overlap found, skip the shorter change
            if (overlapLen >= 3) {
                const overlap = current.find_text.slice(-overlapLen);
                console.log(`[textDiff] Boundary overlap found: "${overlap}" (${overlapLen} chars)`);

                if (current.find_text.length < other.find_text.length) {
                    console.log(`[textDiff] Skipping shorter change: "${current.find_text.slice(0, 40)}..."`);
                    shouldSkip = true;
                    break;
                } else {
                    console.log(`[textDiff] Skipping other change: "${other.find_text.slice(0, 40)}..."`);
                    skippedIndices.add(j);
                }
            }
        }

        if (!shouldSkip) {
            filteredChanges.push(current);
        }
        skippedIndices.add(i);
    }

    if (filteredChanges.length !== normalizedChanges.length) {
        console.log(`[textDiff] After overlap filter: ${normalizedChanges.length} → ${filteredChanges.length} changes`);
    }

    // ==========================================================================
    // PASS 2: Handle CONTAINED overlaps where one find string contains another
    // CONSERVATIVE APPROACH: Only merge when inner's find_text appears in BOTH
    // the outer's find_text AND replace_text (true substring edit case)
    // ==========================================================================
    const mergedChanges: MinimalChange[] = [];
    const usedIndices = new Set<number>();

    for (let i = 0; i < filteredChanges.length; i++) {
        if (usedIndices.has(i)) continue;

        let change = { ...filteredChanges[i] };
        let wasMerged = false;

        // Check if any other change's find_text is a substring of this one
        for (let j = 0; j < filteredChanges.length; j++) {
            if (j === i || usedIndices.has(j)) continue;
            const inner = filteredChanges[j];

            // Only merge if:
            // 1. Inner's find_text is inside outer's find_text
            // 2. Inner's find_text is ALSO in outer's replace_text (meaning we can apply the replacement)
            // 3. They're different changes
            const innerInFind = change.find_text.includes(inner.find_text);
            const innerInReplace = change.replace_text.includes(inner.find_text);

            if (innerInFind && innerInReplace && change.find_text !== inner.find_text) {
                // Safe to merge - apply inner's replacement to outer's replacement
                change.replace_text = change.replace_text.replace(inner.find_text, inner.replace_text);
                usedIndices.add(j);
                wasMerged = true;
                console.log(`[textDiff] Merged contained: "${inner.find_text}" → "${inner.replace_text}" into "${change.find_text}"`);
            } else if (innerInFind && !innerInReplace && change.find_text !== inner.find_text) {
                // Inner's find_text is in outer's find but NOT in replace - this is NOT a merge case
                // These are INDEPENDENT changes that happen to overlap, skip the inner one
                console.log(`[textDiff] Skipping overlap (not mergeable): "${inner.find_text}" not in replace text`);
            }
        }

        if (wasMerged) {
            console.log(`[textDiff] Result: "${change.find_text}" → "${change.replace_text}"`);
        }

        mergedChanges.push(change);
        usedIndices.add(i);
    }

    if (mergedChanges.length !== filteredChanges.length) {
        console.log(`[textDiff] After contained merge: ${filteredChanges.length} → ${mergedChanges.length} changes`);
    }

    return mergedChanges;
}

/**
 * Legacy support
 */
export function findMinimalChange(original: string, amended: string): MinimalChange | null {
    const changes = findMinimalChanges(original, amended);
    return changes.length > 0 ? changes[0] : null;
}
