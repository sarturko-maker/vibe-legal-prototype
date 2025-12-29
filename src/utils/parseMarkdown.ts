/**
 * Simple markdown parser for AI responses
 * Handles: **bold**, *italic*, line breaks
 */
export function parseMarkdown(text: string): string {
    if (!text) return '';

    let html = text
        // Escape HTML first
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        // Bold: **text** or __text__
        .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
        .replace(/__(.+?)__/g, '<strong>$1</strong>')
        // Italic: *text* or _text_ (simple version)
        .replace(/\*([^*]+)\*/g, '<em>$1</em>')
        .replace(/_([^_]+)_/g, '<em>$1</em>')
        // Line breaks
        .replace(/\n\n/g, '</p><p>')
        .replace(/\n/g, '<br />');

    // Wrap in paragraph
    html = '<p>' + html + '</p>';

    return html;
}
