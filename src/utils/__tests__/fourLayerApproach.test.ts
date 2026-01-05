/**
 * Four-Layer Approach Tests
 * Tests for markdown normalization and diff formatting.
 */

import { toMarkdown, debugText } from '../markdownNormalizer';
import { formatDiffForAI, TextChange, convertDmpToTextChanges } from '../diffFormatter';

describe('Markdown Normalization', () => {
    it('should normalize non-breaking spaces', () => {
        const input = 'option,\u00A0to:';
        const expected = 'option, to:';
        expect(toMarkdown(input)).toBe(expected);
    });

    it('should normalize smart quotes', () => {
        const input = '\u201Csmart\u201D quotes';
        const expected = '"smart" quotes';
        expect(toMarkdown(input)).toBe(expected);
    });

    it('should normalize smart apostrophes', () => {
        const input = 'Buyer\u2019s obligation';
        const expected = "Buyer's obligation";
        expect(toMarkdown(input)).toBe(expected);
    });

    it('should collapse multiple spaces', () => {
        const input = 'shall  be,   at';
        const expected = 'shall be, at';
        expect(toMarkdown(input)).toBe(expected);
    });

    it('should remove zero-width characters', () => {
        const input = 'text\u200Bwith\u200Czero\u200Dwidth';
        const expected = 'textwithzerowidth';
        expect(toMarkdown(input)).toBe(expected);
    });

    it('should normalize en-dash and em-dash to hyphen', () => {
        const input = 'section\u2013one\u2014two';
        const expected = 'section-one-two';
        expect(toMarkdown(input)).toBe(expected);
    });

    it('should normalize ellipsis', () => {
        const input = 'wait\u2026';
        const expected = 'wait...';
        expect(toMarkdown(input)).toBe(expected);
    });

    it('should trim lines and remove excessive blank lines', () => {
        const input = '  line one  \n\n\n\n  line two  ';
        const expected = 'line one\n\nline two';
        expect(toMarkdown(input)).toBe(expected);
    });
});

describe('Diff Formatting', () => {
    it('should format deletion', () => {
        const changes: TextChange[] = [
            { type: 'delete', text: 'sole remedy' }
        ];
        expect(formatDiffForAI(changes)).toBe('- Deleted: "sole remedy"');
    });

    it('should format insertion', () => {
        const changes: TextChange[] = [
            { type: 'insert', text: 'new clause' }
        ];
        expect(formatDiffForAI(changes)).toBe('- Inserted: "new clause"');
    });

    it('should format replacement (delete + insert)', () => {
        const changes: TextChange[] = [
            { type: 'delete', text: 'thirty' },
            { type: 'insert', text: 'sixty' }
        ];
        expect(formatDiffForAI(changes)).toBe('- Changed: "thirty" → "sixty"');
    });

    it('should skip equal changes', () => {
        const changes: TextChange[] = [
            { type: 'equal', text: 'unchanged text' },
            { type: 'delete', text: 'old' },
            { type: 'insert', text: 'new' }
        ];
        expect(formatDiffForAI(changes)).toBe('- Changed: "old" → "new"');
    });

    it('should return message when no changes', () => {
        const changes: TextChange[] = [
            { type: 'equal', text: 'unchanged text' }
        ];
        expect(formatDiffForAI(changes)).toBe('(No changes detected)');
    });

    it('should convert diff-match-patch format', () => {
        const dmpDiffs: Array<[number, string]> = [
            [0, 'unchanged '],
            [-1, 'old'],
            [1, 'new']
        ];
        const result = convertDmpToTextChanges(dmpDiffs);
        expect(result).toEqual([
            { type: 'equal', text: 'unchanged ' },
            { type: 'delete', text: 'old' },
            { type: 'insert', text: 'new' }
        ]);
    });
});

describe('Diff Edge Cleanup', () => {
    // Import findMinimalChanges for integration testing
    const { findMinimalChanges } = require('../textDiff');

    it('should produce clean insertion when text is added after punctuation', () => {
        // Original: "instructions or documentation; (d) normal wear"
        // Amended:  "instructions or documentation; or (d) normal wear"
        // Expected: Only " or" shows as inserted
        const original = 'instructions or documentation; (d) normal wear';
        const amended = 'instructions or documentation; or (d) normal wear';

        const changes = findMinimalChanges(original, amended);

        // Should have exactly one change that adds " or"
        expect(changes.length).toBe(1);
        // The change should NOT delete "documentation;" - it should only add " or"
        expect(changes[0].find_text).not.toContain('documentation;');
    });

    it('should produce clean deletion when suffix is removed', () => {
        // Original: "tear; or (e) Products."
        // Amended:  "tear."
        // Expected: Only "; or (e) Products" is deleted, "tear" and "." unchanged
        const original = 'normal wear and tear; or (e) Products.';
        const amended = 'normal wear and tear.';

        const changes = findMinimalChanges(original, amended);

        // Should have exactly one change
        expect(changes.length).toBe(1);
        // The replacement should contain "tear" in both find and replace to preserve it
        // But the actual deletion is just "; or (e) Products"
    });

    it('should handle common prefix trimming', () => {
        // Original: "documentation;"
        // Amended:  "documentation;or"
        // Expected: Just "or" inserted after "documentation;"
        const original = 'documentation;';
        const amended = 'documentation;or';

        const changes = findMinimalChanges(original, amended);

        // Should show pure insertion, not replacement
        expect(changes.length).toBe(1);
        expect(changes[0].replace_text).toContain('or');
    });
});
