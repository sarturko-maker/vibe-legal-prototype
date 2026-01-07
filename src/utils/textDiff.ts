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
 * Count occurrences of a substring in text.
 */
function countOccurrences(text: string, searchString: string): number {
    if (!searchString || searchString.length === 0) return 0;

    let count = 0;
    let position = 0;

    while ((position = text.toLowerCase().indexOf(searchString.toLowerCase(), position)) !== -1) {
        count++;
        position += 1; // Move forward to find overlapping matches too
    }

    return count;
}

/**
 * Extract a UNIQUE anchor from text before a given position.
 * Starts with minimum words and expands until the anchor is unique.
 * 
 * CRITICAL: Returns null if no unique anchor can be found!
 * Caller must handle null and skip the operation.
 * 
 * @param fullText - The entire paragraph text (to test uniqueness)
 * @param textBeforeChange - Text before the change position
 * @param minWords - Minimum number of words to start with (default 5)
 * @returns A unique anchor string, or NULL if none found
 */
function extractUniqueAnchor(
    fullText: string,
    textBeforeChange: string,
    minWords: number = 5
): string | null {
    if (!textBeforeChange || textBeforeChange.length === 0) {
        // At start of paragraph - return empty string (valid case)
        return '';
    }

    // Split into words (preserving content)
    const words = textBeforeChange.trim().split(/\s+/).filter(w => w.length > 0);

    if (words.length === 0) {
        return '';
    }

    // Start with minimum words, expand until unique
    // Try up to all available words (max 50)
    const maxWords = Math.min(words.length, 50);

    for (let wordCount = Math.min(minWords, words.length); wordCount <= maxWords; wordCount += 2) {
        const selectedWords = words.slice(-wordCount);
        const anchor = selectedWords.join(' ');

        // CRITICAL: Reject anchors that are too short (< 5 chars trimmed)
        // Single-char anchors like "(" cause corruption in sequential operations
        const trimmedLength = anchor.replace(/\s/g, '').length;
        if (trimmedLength < 5) {
            console.warn(`[ANCHOR DEBUG] Anchor too short (${trimmedLength} chars), expanding...`);
            continue;  // Try with more words
        }

        // Test uniqueness
        const occurrences = countOccurrences(fullText, anchor);

        console.log(`[ANCHOR DEBUG] Testing ${wordCount} words (${trimmedLength} chars): "${anchor.substring(0, 50)}..." = ${occurrences} occurrences`);

        if (occurrences === 1) {
            console.log(`[ANCHOR DEBUG] ✓ Found unique anchor with ${wordCount} words`);
            return anchor;
        }
    }

    // Could NOT find unique anchor even with all available words!
    // Return NULL - caller must skip this operation
    console.error('[ANCHOR ERROR] ✗ Could not find unique anchor!');
    console.error(`[ANCHOR ERROR] Text before change: "${textBeforeChange.substring(0, 100)}..."`);
    console.error(`[ANCHOR ERROR] Full paragraph: "${fullText.substring(0, 200)}..."`);

    return null;  // CRITICAL: Return null, not a non-unique fallback!
}

/**
 * Legacy extractAnchor - kept for backward compatibility, now calls extractUniqueAnchor
 * @deprecated Use extractUniqueAnchor instead
 */
function extractAnchor(text: string, charCount: number, fullText?: string): string {
    // If fullText provided, use uniqueness-tested approach
    if (fullText) {
        return extractUniqueAnchor(fullText, text);
    }

    // Fallback to old behavior (just last N chars)
    if (!text || text.length === 0) {
        return '';
    }

    let anchor = text.slice(-charCount).trim();

    if (anchor.length > 0 && text.length > charCount) {
        const firstSpace = anchor.indexOf(' ');
        if (firstSpace > 0 && firstSpace < anchor.length - 1) {
            anchor = anchor.substring(firstSpace + 1);
        }
    }

    return anchor;
}

// Note: expandReplacementContext function removed - REPLACE type eliminated
// All replacements are now decomposed into DELETE_AFTER + INSERT_AFTER pairs

/**
 * Ensure proper spacing between anchor and inserted text.
 * Prevents text from running together like "madeDAP" or "followingthe".
 */
function ensureProperSpacing(anchor: string, textToInsert: string): string {
    // If anchor is empty (start of paragraph), don't add space
    if (!anchor) return textToInsert;

    // If anchor ends with whitespace, no space needed
    const anchorEndsWithSpace = /\s$/.test(anchor);
    if (anchorEndsWithSpace) return textToInsert;

    // If text starts with whitespace, no space needed
    const textStartsWithSpace = /^\s/.test(textToInsert);
    if (textStartsWithSpace) return textToInsert;

    // If text starts with punctuation, no space needed
    const textStartsWithPunctuation = /^[.,;:!?)}\]'"']/.test(textToInsert);
    if (textStartsWithPunctuation) return textToInsert;

    // Add leading space
    console.log(`[textDiff] Adding space before insert: "${textToInsert.substring(0, 20)}..."`);
    return ' ' + textToInsert;
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
 * Merge adjacent diff operations that are close together.
 * This converts many micro-operations into fewer phrase-level operations.
 * 
 * Rules:
 * - If two changes are separated by less than maxGapChars of EQUAL text, merge them
 * - This creates larger, more anchorable chunks
 * - Prevents issues like "Delete 't'" + "Insert 'f'" becoming separate operations
 */
function mergeAdjacentDiffs(diffs: Array<[number, string]>, maxGapChars: number = 20): Array<[number, string]> {
    if (diffs.length <= 1) return diffs;

    const merged: Array<[number, string]> = [];
    let pendingDelete = '';
    let pendingInsert = '';
    let pendingEqual = '';

    for (let i = 0; i < diffs.length; i++) {
        const [op, text] = diffs[i];

        if (op === 0) { // EQUAL
            // If we have pending changes and this EQUAL is short, absorb it
            if ((pendingDelete || pendingInsert) && text.length <= maxGapChars) {
                // Check if there's another change coming after this EQUAL
                const nextOp = i + 1 < diffs.length ? diffs[i + 1][0] : 0;
                if (nextOp !== 0) {
                    // More changes coming - absorb this EQUAL into pending
                    pendingDelete += text;
                    pendingInsert += text;
                    continue;
                }
            }

            // Flush pending changes before this EQUAL
            if (pendingDelete || pendingInsert) {
                if (pendingDelete === pendingInsert) {
                    // Actually no change - just EQUAL
                    merged.push([0, pendingDelete]);
                } else if (pendingDelete && pendingInsert) {
                    merged.push([-1, pendingDelete]);
                    merged.push([1, pendingInsert]);
                } else if (pendingDelete) {
                    merged.push([-1, pendingDelete]);
                } else if (pendingInsert) {
                    merged.push([1, pendingInsert]);
                }
                pendingDelete = '';
                pendingInsert = '';
            }

            merged.push([0, text]);

        } else if (op === -1) { // DELETE
            pendingDelete += text;

        } else if (op === 1) { // INSERT
            pendingInsert += text;
        }
    }

    // Flush any remaining pending changes
    if (pendingDelete || pendingInsert) {
        if (pendingDelete === pendingInsert) {
            merged.push([0, pendingDelete]);
        } else if (pendingDelete && pendingInsert) {
            merged.push([-1, pendingDelete]);
            merged.push([1, pendingInsert]);
        } else if (pendingDelete) {
            merged.push([-1, pendingDelete]);
        } else if (pendingInsert) {
            merged.push([1, pendingInsert]);
        }
    }

    console.log(`[textDiff] Merged ${diffs.length} diffs into ${merged.length} diffs`);
    return merged;
}

/**
 * Convert diffs to TextChange array.
 * Uses insert_after for pure insertions, delete_after for pure deletions.
 * All replacements are decomposed into DELETE + INSERT pairs.
 * 
 * @param diffs - Array of diff tuples from diff_match_patch
 * @param originalText - Full original text for anchor uniqueness testing
 */
/**
 * Extract a UNIQUE anchor before a specific position in the full text.
 * Expands word count until uniqueness is achieved or limit reached.
 * 
 * @param fullText - The entire text
 * @param position - The insertion/change position (0-indexed)
 * @param minWords - Minimum words to start with
 */
function extractUniqueAnchorAtPosition(
    fullText: string,
    position: number,
    minWords: number = 5
): string | null {
    if (position <= 0) return '';
    if (position > fullText.length) position = fullText.length;

    const textBefore = fullText.substring(0, position);
    const words = textBefore.trim().split(/\s+/).filter(w => w.length > 0);

    if (words.length === 0) return '';

    const maxWords = Math.min(words.length, 50);

    // CLAUSE NUMBER PATTERN: Accept short anchors that look like clause numbers (e.g., "6.2", "10.1.3")
    // These are typically unique within a paragraph even if short
    const clausePattern = /^\d+(\.\d+)*$/;

    // Expand anchor size until unique
    for (let wordCount = Math.min(minWords, words.length); wordCount <= maxWords; wordCount += 2) {
        const selectedWords = words.slice(-wordCount);
        const anchor = selectedWords.join(' ');

        // Check anchor length requirements
        const trimmedLength = anchor.replace(/\s/g, '').length;

        // Accept shorter anchors if they contain a clause number pattern
        const hasClauseNumber = selectedWords.some(w => clausePattern.test(w));
        const minLength = hasClauseNumber ? 2 : 5;  // Clause numbers can be shorter

        if (trimmedLength < minLength) {
            if (wordCount < 10) {
                continue;
            }
        }

        const occurrences = countOccurrences(fullText, anchor);

        console.log(`[ANCHOR DEBUG] Testing ${wordCount} words: "${anchor.substring(0, 30)}..." = ${occurrences} occurrences`);

        if (occurrences === 1) {
            return anchor;
        }
    }

    // FALLBACK: For early-paragraph changes (position < 30 chars), 
    // try using the full textBefore as anchor - it may still be unique for paragraph-local search
    if (position < 30 && textBefore.trim().length > 0) {
        const fullAnchor = textBefore.trim();
        const occurrences = countOccurrences(fullText, fullAnchor);

        console.log(`[ANCHOR DEBUG] Early-paragraph fallback: "${fullAnchor}" = ${occurrences} occurrences`);

        if (occurrences === 1) {
            console.log(`[ANCHOR DEBUG] ✓ Using early-paragraph anchor: "${fullAnchor}"`);
            return fullAnchor;
        }

        // Even if not unique in full text, return it for paragraph-local matching
        // The Word search API searches within the paragraph, so it may still work
        console.log(`[ANCHOR DEBUG] ✓ Using early-paragraph anchor (paragraph-local): "${fullAnchor}"`);
        return fullAnchor;
    }

    console.error(`[ANCHOR ERROR] Could not find unique anchor ending at pos ${position} inside full text len ${fullText.length}`);
    return null;
}

/**
 * Convert diffs to TextChange array.
 * Uses insert_after for pure insertions, delete_after for pure deletions.
 * All replacements are decomposed into DELETE + INSERT pairs.
 * 
 * @param diffs - Array of diff tuples from diff_match_patch
 * @param originalText - Full original text for anchor uniqueness testing
 */
function convertDiffsToTextChanges(diffs: Array<[number, string]>, originalText: string): TextChange[] {
    const changes: TextChange[] = [];
    let currentPos = 0; // Track position in originalText

    for (let i = 0; i < diffs.length; i++) {
        const [op, text] = diffs[i];

        if (op === 0) {
            // EQUAL: Advance position
            currentPos += text.length;
            continue;
        }

        if (op === 1) {
            // INSERT
            // Check context for replacement logic
            // (If previous op was DELETE, it's a replacement, handled in the DELETE block)
            const prevDiff = i > 0 ? diffs[i - 1] : null;
            const prevPrevDiff = i > 1 ? diffs[i - 2] : null;

            // If this insert immediately follows a DELETE (diffs[i-1] == -1), it was already handled.
            // BUT: standard dmp output puts DELETE then INSERT. 
            // My loop handles DELETE case below by looking ahead.
            // So if we encounter INSERT here, we must check if it was already handled by the PREVIOUS delete.

            if (prevDiff && prevDiff[0] === -1) {
                // Was handled by DELETE block as replacement
                continue;
            }

            // PURE INSERTION
            const anchor = extractUniqueAnchorAtPosition(originalText, currentPos);

            if (anchor === null) {
                console.warn(`[textDiff] SKIPPING INSERT - no unique anchor found at pos ${currentPos}`);
                continue;
            }

            changes.push({
                type: 'insert_after',
                anchor: anchor,
                text: ensureProperSpacing(anchor, text)
            });
            // Position does NOT advance for insertion (it inserts at currentPos)
            continue;
        }

        if (op === -1) {
            // DELETE
            // Check if NEXT is INSERT (Replacement)
            const nextDiff = i + 1 < diffs.length ? diffs[i + 1] : null;

            if (nextDiff && nextDiff[0] === 1) {
                // REPLACEMENT: DELETE + INSERT
                const anchor = extractUniqueAnchorAtPosition(originalText, currentPos);
                const oldText = text;
                const newText = nextDiff[1];

                if (anchor === null) {
                    console.warn(`[textDiff] SKIPPING REPLACEMENT - no unique anchor found at pos ${currentPos}`);
                    currentPos += oldText.length; // Skip this text
                    i++; // Skip next insert
                    continue;
                }

                // Delete old
                changes.push({
                    type: 'delete_after',
                    anchor: anchor,
                    textToDelete: oldText
                });

                // Insert new (anchored to same spot)
                changes.push({
                    type: 'insert_after',
                    anchor: anchor,
                    text: ensureProperSpacing(anchor, newText)
                });

                currentPos += oldText.length;
                i++; // Consumed next diff
                continue;
            } else {
                // PURE DELETE
                const anchor = extractUniqueAnchorAtPosition(originalText, currentPos);

                if (anchor === null) {
                    console.warn(`[textDiff] SKIPPING DELETE - no unique anchor found at pos ${currentPos}`);
                    currentPos += text.length;
                    continue;
                }

                changes.push({
                    type: 'delete_after',
                    anchor: anchor,
                    textToDelete: text
                });

                currentPos += text.length;
            }
        }
    }

    return changes;
}

/**
 * Convert TextChange array to MinimalChange array for backward compatibility.
 * Note: REPLACE type removed - only insert_after and delete_after remain
 */
function textChangesToMinimalChanges(changes: TextChange[]): MinimalChange[] {
    return changes.map(change => {
        if (change.type === 'insert_after') {
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

    // Use word-level tokenization to prevent character-level micro-operations
    // This groups changes by word boundaries, making them easier to anchor
    const wordSplitRegex = /([^\s]+|\s+)/g;
    const words1 = normOriginal.match(wordSplitRegex) || [];
    const words2 = normAmended.match(wordSplitRegex) || [];

    // Create a map of words to characters (hashing)
    const wordToChar = new Map<string, string>();
    const charToWord: string[] = [];
    let charCode = 20000; // Start high to avoid control characters

    function getCharForWord(word: string): string {
        if (!wordToChar.has(word)) {
            const char = String.fromCharCode(charCode++);
            wordToChar.set(word, char);
            charToWord.push(word);
            return char;
        }
        return wordToChar.get(word)!;
    }

    // Convert word arrays to character strings
    const chars1 = words1.map(getCharForWord).join('');
    const chars2 = words2.map(getCharForWord).join('');

    // Perform the diff on the character strings (which represent words)
    // false = Do not check for semantic cleanup yet, we want exact word matches first
    const diffsChars = dmp.diff_main(chars1, chars2, false);

    // Convert back from characters to words
    let diffs: Array<[number, string]> = diffsChars.map((diff: [number, string]) => {
        const op = diff[0];
        const text = diff[1].split('').map(char => {
            const index = char.charCodeAt(0) - 20000;
            return charToWord[index] || '';
        }).join('');
        return [op, text];
    });

    // Now apply semantic cleanup to merge remaining tiny changes
    dmp.diff_cleanupSemantic(diffs);
    diffs = cleanupDiffEdges(diffs);

    // CRITICAL: Merge adjacent changes separated by small gaps
    // This reduces 11 micro-operations into 2-3 phrase-level operations
    diffs = mergeAdjacentDiffs(diffs, 30);  // Merge changes within 30 chars of each other

    console.log(`[textDiff] After word-level diff + merge: ${diffs.length} raw diffs`);

    const changes = convertDiffsToTextChanges(diffs, normOriginal);

    console.log(`[textDiff] Found ${changes.length} text changes`);
    changes.forEach((change, idx) => {
        if (change.type === 'insert_after') {
            console.log(`[textDiff]   [${idx}] INSERT_AFTER: anchor="${change.anchor}" text="${change.text}"`);
        } else if (change.type === 'delete_after') {
            console.log(`[textDiff]   [${idx}] DELETE_AFTER: anchor="${change.anchor}" delete="${change.textToDelete.substring(0, 30)}..."`);
        }
    });

    return changes;
}

// Re-export TextChange type for convenience
export type { TextChange } from '../types/operations';
