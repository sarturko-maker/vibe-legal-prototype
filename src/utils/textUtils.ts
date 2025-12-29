/**
 * Text utilities for Vibe Legal
 * Pure helper functions for string manipulation
 */

/**
 * Escape special XML characters
 */
export function escapeXml(text: string): string {
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * Unescape XML entities back to characters
 */
export function unescapeXml(text: string): string {
    return text
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, '&');
}

/**
 * Normalize whitespace: collapse multiple spaces, trim
 */
export function normalizeWhitespace(text: string): string {
    return text.replace(/\s+/g, ' ').trim();
}

/**
 * Strip markdown formatting markers
 */
export function stripMarkdownMarkers(text: string): string {
    return text
        .replace(/\*\*(.+?)\*\*/g, '$1')       // **bold**
        .replace(/(?<!\*)\*([^*]+?)\*(?!\*)/g, '$1')  // *italic*
        .replace(/__(.+?)__/g, '$1');          // __underline__
}

/**
 * Extract clause number from text (e.g., "1.2.3 Title" -> "1.2.3")
 */
export function extractClauseNumber(text: string): string | null {
    const match = text.match(/^(\d+(?:\.\d+)*)/);
    return match ? match[1] : null;
}

/**
 * Extract clause title from text (e.g., "1.2.3 Title Here" -> "Title Here")
 */
export function extractClauseTitle(text: string): string | null {
    const match = text.match(/^\d+(?:\.\d+)*\s+(.+?)(?:\.|$)/);
    return match ? match[1].trim() : null;
}

/**
 * Truncate text with ellipsis
 */
export function truncate(text: string, maxLength: number): string {
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength - 3) + '...';
}

/**
 * Check if text looks like a heading/clause title
 */
export function isLikelyHeading(text: string): boolean {
    const trimmed = text.trim();
    // Check for numbered clause pattern or all caps
    return /^\d+\.?\s/.test(trimmed) ||
        (trimmed.length < 100 && trimmed === trimmed.toUpperCase());
}
