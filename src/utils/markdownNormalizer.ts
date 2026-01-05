/**
 * Markdown Normalizer
 * Converts Word paragraph text to clean format, eliminating formatting artifacts.
 */

/**
 * Normalize text to clean markdown format.
 * Eliminates invisible characters that cause false diffs.
 */
export function toMarkdown(wordText: string): string {
    let md = wordText;

    // Step 1: Replace non-breaking spaces with regular spaces
    md = md.replace(/\u00A0/g, ' ');

    // Step 2: Remove soft hyphens
    md = md.replace(/\u00AD/g, '');

    // Step 3: Remove zero-width characters
    md = md.replace(/[\u200B-\u200D\uFEFF]/g, '');

    // Step 4: Normalize smart quotes to straight quotes
    md = md.replace(/[\u2018\u2019]/g, "'");  // Single quotes
    md = md.replace(/[\u201C\u201D]/g, '"');  // Double quotes

    // Step 5: Normalize dashes (en-dash, em-dash → hyphen)
    md = md.replace(/[\u2013\u2014]/g, '-');

    // Step 6: Normalize ellipsis
    md = md.replace(/\u2026/g, '...');

    // Step 7: Collapse multiple spaces/tabs into single space
    md = md.replace(/[ \t]+/g, ' ');

    // Step 8: Normalize line endings
    md = md.replace(/\r\n/g, '\n');
    md = md.replace(/\r/g, '\n');

    // Step 9: Trim each line
    md = md.split('\n')
        .map(line => line.trim())
        .join('\n');

    // Step 10: Remove multiple consecutive blank lines
    md = md.replace(/\n{3,}/g, '\n\n');

    // Step 11: Final trim
    md = md.trim();

    return md;
}

/**
 * Debug helper: shows invisible characters in text
 */
export function debugText(text: string, label: string = 'Text'): void {
    console.log(`[${label}] Length: ${text.length}`);
    console.log(`[${label}] Hex:`, [...text].map(c => {
        const code = c.charCodeAt(0);
        if (code > 127 || code < 32) {
            return `\\u${code.toString(16).padStart(4, '0')}`;
        }
        return c;
    }).join(''));
}
