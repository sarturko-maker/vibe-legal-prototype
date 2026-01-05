/**
 * Text Diff Utility - Deterministic minimal change detection
 * 
 * Uses diff-match-patch for character-level diffs with semantic cleanup.
 * This ensures "4.1" and "Word." are treated correctly as units,
 * avoiding the punctuation-splitting issues of word-level diffs.
 */

import { diff_match_patch } from 'diff-match-patch';
import { TextChange } from '../types/operations';

export interface MinimalChange {
    find_text: string;
    replace_text: string;
    // For pure insertions: instead of find/replace, we insert_text after insert_after
    insert_after?: string;
    insert_text?: string;
}

/**
 * Extract last N characters from text, preferring word boundaries.
 * Used to create multi-word anchors for insertions/deletions.
 */
function extractAnchor(text: string, charCount: number): string {
    if (!text || text.length === 0) {
        return '';
    }

    // Take last N chars
    let anchor = text.slice(-charCount).trim();

    // If we cut mid-word, try to find a word boundary
    if (anchor.length > 0 && text.length > charCount) {
        const firstSpace = anchor.indexOf(' ');
        if (firstSpace > 0 && firstSpace < anchor.length - 1) {
            // Cut off the partial word at the start
            anchor = anchor.substring(firstSpace + 1);
        }
    }

    return anchor;
}

/**
 * Expand replacement context to make find/replace strings unique.
 * Uses direct substring extraction to preserve exact spacing.
 * 
 * @param beforeContext - Text that comes BEFORE the change (EQUAL segment)
 * @param afterContext - Text that comes AFTER the change (EQUAL segment)
 * @param findText - Original text being replaced
 * @param replaceText - New text to replace with
 * @returns Expanded find/replace with symmetric context
 */
function expandReplacementContext(
    beforeContext: string,
    afterContext: string,
    findText: string,
    replaceText: string
): { find: string; replace: string } {
    // Use character-based extraction to preserve exact spacing
    const contextChars = 25;

    // Extract last N chars from before context (preserve spaces!)
    let leftContext = '';
    if (beforeContext.length > 0) {
        leftContext = beforeContext.slice(-contextChars);
        // Try to start at a word boundary (find first space and cut there)
        const spaceIdx = leftContext.indexOf(' ');
        if (spaceIdx > 0 && spaceIdx < leftContext.length - 1) {
            leftContext = leftContext.substring(spaceIdx); // Keep the space!
        }
    }

    // Extract first N chars from after context (preserve spaces!)
    let rightContext = '';
    if (afterContext.length > 0) {
        rightContext = afterContext.slice(0, contextChars);
        // Try to end at a word boundary (find last space and cut there)
        const lastSpaceIdx = rightContext.lastIndexOf(' ');
        if (lastSpaceIdx > 0) {
            rightContext = rightContext.substring(0, lastSpaceIdx); // Cut at space
        }
    }

    // Build expanded strings - DON'T TRIM, preserve spacing!
    const expandedFind = leftContext + findText + rightContext;
    const expandedReplace = leftContext + replaceText + rightContext;

    console.log(`[textDiff] Expanded replacement: "${findText}" → "${replaceText}"`);
    console.log(`[textDiff]   With context: "${expandedFind}" → "${expandedReplace}"`);

    return { find: expandedFind, replace: expandedReplace };
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
 * Trims common prefixes and suffixes from DELETE+INSERT pairs.
 * This makes track changes show only actual differences.
 * 
 * Example: DELETE "documentation;" + INSERT "documentation;or"
 * Becomes: EQUAL "documentation;" + INSERT "or"
 */
function cleanupDiffEdges(diffs: Array<[number, string]>): Array<[number, string]> {
    const cleaned: Array<[number, string]> = [];

    for (let i = 0; i < diffs.length; i++) {
        const [op, text] = diffs[i];

        // Look for DELETE (-1) + INSERT (1) pairs (replacements)
        if (op === -1 && i + 1 < diffs.length && diffs[i + 1][0] === 1) {
            const oldText = text;
            const newText = diffs[i + 1][1];

            // Find common prefix
            let prefixLen = 0;
            const minLen = Math.min(oldText.length, newText.length);
            while (prefixLen < minLen && oldText[prefixLen] === newText[prefixLen]) {
                prefixLen++;
            }

            // Find common suffix (after accounting for prefix)
            let suffixLen = 0;
            const remainingOld = oldText.length - prefixLen;
            const remainingNew = newText.length - prefixLen;
            const minRemaining = Math.min(remainingOld, remainingNew);

            while (suffixLen < minRemaining &&
                oldText[oldText.length - 1 - suffixLen] === newText[newText.length - 1 - suffixLen]) {
                suffixLen++;
            }

            // Extract parts
            const commonPrefix = oldText.substring(0, prefixLen);
            const commonSuffix = suffixLen > 0 ? oldText.substring(oldText.length - suffixLen) : '';

            const oldCore = oldText.substring(prefixLen, oldText.length - suffixLen);
            const newCore = newText.substring(prefixLen, newText.length - suffixLen);

            // Add cleaned operations
            if (commonPrefix.length > 0) {
                cleaned.push([0, commonPrefix]); // EQUAL
            }

            if (oldCore.length > 0) {
                cleaned.push([-1, oldCore]); // DELETE
            }

            if (newCore.length > 0) {
                cleaned.push([1, newCore]); // INSERT
            }

            if (commonSuffix.length > 0) {
                cleaned.push([0, commonSuffix]); // EQUAL
            }

            i++; // Skip next since we processed the INSERT
            continue;
        }

        // Keep all other diffs unchanged
        cleaned.push([op, text]);
    }

    return cleaned;
}

/**
 * Convert diffs to TextChange array.
 * Uses insert_after for pure insertions, delete_after for pure deletions,
 * and replace for replacements. This produces cleaner track changes.
 */
function convertDiffsToTextChanges(diffs: Array<[number, string]>): TextChange[] {
    const changes: TextChange[] = [];

    for (let i = 0; i < diffs.length; i++) {
        const [op, text] = diffs[i];

        if (op === 0) {
            continue; // Skip unchanged text
        }

        if (op === 1) {
            // INSERT - check if this is part of a replacement or pure insertion
            const prevDiff = i > 0 ? diffs[i - 1] : null;
            const prevPrevDiff = i > 1 ? diffs[i - 2] : null;
            const prevIsDelete = prevPrevDiff && prevPrevDiff[0] === -1;

            // If previous operation was DELETE (before any EQUAL), this is part of a replacement
            // The DELETE handler already processed it, so skip
            if (prevIsDelete && prevDiff && prevDiff[0] === 0) {
                // This insert follows EQUAL after DELETE - it's a standalone pure insertion
            } else if (prevPrevDiff && prevPrevDiff[0] === -1) {
                // Skip - already handled by delete
                continue;
            }

            // PURE INSERTION - use insert_after
            const beforeText = prevDiff && prevDiff[0] === 0 ? prevDiff[1] : '';
            const anchor = extractAnchor(beforeText, 40);

            if (!anchor) {
                console.warn('[textDiff] Pure insertion with no anchor context - using fallback');
                // Fallback: look for text after insertion
                const nextDiff = i + 1 < diffs.length ? diffs[i + 1] : null;
                const afterContext = nextDiff && nextDiff[0] === 0 ? nextDiff[1].slice(0, 20) : '';

                if (afterContext) {
                    changes.push({
                        type: 'replace',
                        find: afterContext,
                        replace: text + afterContext
                    });
                } else {
                    console.error('[textDiff] Pure insertion with no context at all');
                }
            } else {
                changes.push({
                    type: 'insert_after',
                    anchor: anchor,
                    text: text
                });
            }
            continue;
        }

        if (op === -1) {
            // DELETE - check if this is part of a replacement
            const nextDiff = i + 1 < diffs.length ? diffs[i + 1] : null;

            if (nextDiff && nextDiff[0] === 1) {
                // REPLACEMENT - DELETE followed by INSERT
                // Expand context to make find/replace unique in paragraph
                const prevDiff = i > 0 ? diffs[i - 1] : null;
                const afterInsertDiff = i + 2 < diffs.length ? diffs[i + 2] : null;

                const beforeContext = prevDiff && prevDiff[0] === 0 ? prevDiff[1] : '';
                const afterContext = afterInsertDiff && afterInsertDiff[0] === 0 ? afterInsertDiff[1] : '';

                const { find, replace } = expandReplacementContext(
                    beforeContext,
                    afterContext,
                    text,           // Original (being deleted)
                    nextDiff[1]     // New (being inserted)
                );

                changes.push({
                    type: 'replace',
                    find: find,
                    replace: replace
                });
                i++; // Skip next since we processed it
                continue;
            } else {
                // PURE DELETION - use delete_after
                const prevDiff = i > 0 ? diffs[i - 1] : null;
                const beforeText = prevDiff && prevDiff[0] === 0 ? prevDiff[1] : '';
                // Use longer anchor (50 chars) to push further back from any nearby insertions
                const anchor = extractAnchor(beforeText, 50);

                if (!anchor) {
                    console.warn('[textDiff] Pure deletion with no anchor context - using fallback');
                    // Fallback: use replace with empty string
                    changes.push({
                        type: 'replace',
                        find: text,
                        replace: ''
                    });
                } else {
                    changes.push({
                        type: 'delete_after',
                        anchor: anchor,
                        textToDelete: text
                    });
                }
            }
        }
    }

    return changes;
}

/**
 * Convert TextChange array to MinimalChange array for backward compatibility.
 */
function textChangesToMinimalChanges(changes: TextChange[]): MinimalChange[] {
    return changes.map(change => {
        if (change.type === 'replace') {
            return {
                find_text: change.find,
                replace_text: change.replace
            };
        } else if (change.type === 'insert_after') {
            return {
                find_text: change.anchor,
                replace_text: change.anchor + change.text,
                insert_after: change.anchor,
                insert_text: change.text
            };
        } else {
            // delete_after
            return {
                find_text: change.anchor + change.textToDelete,
                replace_text: change.anchor
            };
        }
    });
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
    let diffs = dmp.diff_main(normOriginal, normAmended);

    // Semantic cleanup handles "Buyer." vs "Buyer ." grouping naturally
    dmp.diff_cleanupSemantic(diffs);

    // NEW: Clean up edges - trim common prefix/suffix from DELETE+INSERT pairs
    // This ensures track changes show only actual differences
    diffs = cleanupDiffEdges(diffs);

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

/**
 * Find text changes with proper insert_after/delete_after/replace types.
 * Returns TextChange array for cleaner track changes.
 */
export function findTextChanges(original: string, amended: string): TextChange[] {
    const normOriginal = original.trim().replace(/\s+/g, ' ');
    const normAmended = amended.trim().replace(/\s+/g, ' ');

    if (normOriginal === normAmended) {
        console.log('[textDiff] Texts are identical');
        return [];
    }

    const dmp = new diff_match_patch();
    let diffs = dmp.diff_main(normOriginal, normAmended);
    dmp.diff_cleanupSemantic(diffs);
    diffs = cleanupDiffEdges(diffs);

    const changes = convertDiffsToTextChanges(diffs);

    console.log(`[textDiff] Found ${changes.length} text changes`);
    changes.forEach((change, idx) => {
        if (change.type === 'replace') {
            console.log(`[textDiff]   [${idx}] Replace: "${change.find.substring(0, 30)}..." → "${change.replace.substring(0, 30)}..."`);
        } else if (change.type === 'insert_after') {
            console.log(`[textDiff]   [${idx}] Insert "${change.text}" after anchor "${change.anchor}"`);
        } else if (change.type === 'delete_after') {
            console.log(`[textDiff]   [${idx}] Delete "${change.textToDelete.substring(0, 30)}..." after anchor "${change.anchor}"`);
        }
    });

    return changes;
}

// Re-export TextChange type for convenience
export type { TextChange } from '../types/operations';
