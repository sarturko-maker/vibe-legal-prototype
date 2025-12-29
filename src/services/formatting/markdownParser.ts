/**
 * Markdown Parser for Vibe Legal
 * Extracts formatting from **bold**, *italic*, __underline__ markers
 * Source: vibe-legal-beta-0.2.yaml L7388-7496
 */

import { Segment } from '../../types';

// Internal type for format ranges
interface FormatRange {
    start: number;
    end: number;
    text: string;
    formatting: { bold?: boolean; italic?: boolean; underline?: boolean; strikethrough?: boolean };
    originalLength: number;
}

/**
 * Parse markdown-style formatting markers from text.
 * 
 * @example
 * parseFormattingMarkers("The **Provider** shall notify")
 * // Returns: [
 * //   { text: "The ", formatting: {} },
 * //   { text: "Provider", formatting: { bold: true } },
 * //   { text: " shall notify", formatting: {} }
 * // ]
 */
export function parseFormattingMarkers(text: string): Segment[] {
    if (!text || text.length === 0) {
        return [];
    }

    const segments: Segment[] = [];

    // Regex patterns for markdown markers
    // Order matters: check longer patterns first (** before *)
    const patterns = [
        { regex: /\*\*([^*]+?)\*\*/g, format: { bold: true, italic: false, underline: false } },
        { regex: /\*([^*]+?)\*/g, format: { bold: false, italic: true, underline: false } },
        { regex: /__([^_]+?)__/g, format: { bold: false, italic: false, underline: true } },
    ];

    const ranges: FormatRange[] = [];

    for (const pattern of patterns) {
        let match;
        pattern.regex.lastIndex = 0;

        while ((match = pattern.regex.exec(text)) !== null) {
            if (match[1].trim().length === 0) {
                continue;
            }

            ranges.push({
                start: match.index,
                end: match.index + match[0].length,
                text: match[1],
                formatting: pattern.format,
                originalLength: match[0].length
            });
        }
    }

    // Sort by start position
    ranges.sort((a, b) => a.start - b.start);

    // Remove overlapping ranges (keep first occurrence)
    const nonOverlapping: FormatRange[] = [];
    for (const range of ranges) {
        const overlaps = nonOverlapping.some(r =>
            (range.start >= r.start && range.start < r.end) ||
            (range.end > r.start && range.end <= r.end) ||
            (range.start <= r.start && range.end >= r.end)
        );
        if (!overlaps) {
            nonOverlapping.push(range);
        }
    }

    // Build segments from ranges
    let currentPos = 0;
    const defaultFormatting = { bold: false, italic: false, underline: false };

    for (const range of nonOverlapping) {
        if (range.start > currentPos) {
            const plainText = text.substring(currentPos, range.start);
            if (plainText.length > 0) {
                segments.push({ text: plainText, formatting: { ...defaultFormatting } });
            }
        }

        segments.push({
            text: range.text,
            formatting: { ...defaultFormatting, ...range.formatting }
        });

        currentPos = range.end;
    }

    // Add remaining plain text
    if (currentPos < text.length) {
        const remainingText = text.substring(currentPos);
        if (remainingText.length > 0) {
            segments.push({ text: remainingText, formatting: { ...defaultFormatting } });
        }
    }

    // If no formatting was found, return single segment
    if (segments.length === 0) {
        segments.push({ text: text, formatting: { ...defaultFormatting } });
    }

    return segments;
}

/**
 * Check if text contains any markdown formatting markers.
 */
export function hasFormattingMarkers(text: string): boolean {
    return /\*\*[^*]+\*\*|\*[^*]+\*|__[^_]+__/.test(text);
}

/**
 * Strip all markdown formatting markers, return plain text.
 */
export function stripFormattingMarkers(text: string): string {
    return text
        .replace(/\*\*([^*]+?)\*\*/g, '$1')
        .replace(/\*([^*]+?)\*/g, '$1')
        .replace(/__([^_]+?)__/g, '$1');
}
