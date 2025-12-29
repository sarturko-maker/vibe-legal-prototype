/**
 * Preview Generator for Vibe Legal
 * Creates HTML previews for Draft mode with colored diff display
 */

import { StructuredDiff } from './surgicalDiff';

/**
 * Generate HTML preview with colored diff.
 * Deletions: red strikethrough
 * Insertions: green underline
 */
export function generateDiffPreviewHtml(
    originalText: string,
    newText: string,
    diffs: StructuredDiff[]
): string {
    const htmlParts: string[] = [];
    let position = 0;

    for (const diff of diffs) {
        // Add any unchanged text before this diff
        if (diff.start > position) {
            const unchanged = escapeHtml(originalText.substring(position, diff.start));
            htmlParts.push(`<span class="preview-unchanged">${unchanged}</span>`);
        }

        switch (diff.type) {
            case 'delete':
                const deleted = escapeHtml(originalText.substring(diff.start, diff.end));
                htmlParts.push(`<span class="preview-deletion">${deleted}</span>`);
                position = diff.end || diff.start;
                break;

            case 'insert':
                const inserted = escapeHtml(diff.text || '');
                htmlParts.push(`<span class="preview-insertion">${inserted}</span>`);
                // Don't advance position - insertion doesn't consume original text
                break;

            case 'replace':
                const deletedText = escapeHtml(originalText.substring(diff.start, diff.end));
                const replacementText = escapeHtml(diff.text || '');
                htmlParts.push(`<span class="preview-deletion">${deletedText}</span>`);
                htmlParts.push(`<span class="preview-insertion">${replacementText}</span>`);
                position = diff.end || diff.start;
                break;
        }
    }

    // Add any remaining text
    if (position < originalText.length) {
        const remaining = escapeHtml(originalText.substring(position));
        htmlParts.push(`<span class="preview-unchanged">${remaining}</span>`);
    }

    return htmlParts.join('');
}

/**
 * Generate a compact summary of changes.
 */
export function generateChangeSummary(diffs: StructuredDiff[]): string {
    let insertCount = 0;
    let deleteCount = 0;
    let replaceCount = 0;

    for (const diff of diffs) {
        switch (diff.type) {
            case 'insert': insertCount++; break;
            case 'delete': deleteCount++; break;
            case 'replace': replaceCount++; break;
        }
    }

    const parts: string[] = [];
    if (replaceCount > 0) parts.push(`${replaceCount} replacement${replaceCount > 1 ? 's' : ''}`);
    if (insertCount > 0) parts.push(`${insertCount} addition${insertCount > 1 ? 's' : ''}`);
    if (deleteCount > 0) parts.push(`${deleteCount} deletion${deleteCount > 1 ? 's' : ''}`);

    return parts.join(', ') || 'No changes';
}

/**
 * Escape HTML special characters.
 */
function escapeHtml(text: string): string {
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}
