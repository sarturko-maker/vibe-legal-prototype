/**
 * Markdown Rendering Utility
 * Converts basic markdown syntax to HTML for chat display
 */

/**
 * Render markdown text to HTML
 * Supports: headers, bold, italic, code, blockquotes, line breaks
 */
export function renderMarkdown(text: string): string {
    if (!text) return '';

    return text
        // Escape HTML to prevent XSS (but allow our generated HTML)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        // Headers (process before bold to avoid conflicts)
        .replace(/^#### (.*$)/gm, '<h5 class="md-h4">$1</h5>')
        .replace(/^### (.*$)/gm, '<h4 class="md-h3">$1</h4>')
        .replace(/^## (.*$)/gm, '<h3 class="md-h2">$1</h3>')
        .replace(/^# (.*$)/gm, '<h2 class="md-h1">$1</h2>')
        // Bold (must come before italic)
        .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
        .replace(/__([^_]+)__/g, '<strong>$1</strong>')
        // Italic
        .replace(/\*([^*]+)\*/g, '<em>$1</em>')
        .replace(/_([^_]+)_/g, '<em>$1</em>')
        // Inline code
        .replace(/`([^`]+)`/g, '<code>$1</code>')
        // Bullet lists (simple)
        .replace(/^- (.*$)/gm, '<li>$1</li>')
        .replace(/^• (.*$)/gm, '<li>$1</li>')
        // Numbered lists (simple)
        .replace(/^\d+\. (.*$)/gm, '<li>$1</li>')
        // Blockquotes
        .replace(/^&gt; (.*$)/gm, '<blockquote>$1</blockquote>')
        // Horizontal rule
        .replace(/^---$/gm, '<hr/>')
        // Line breaks - double newline = paragraph break
        .replace(/\n\n/g, '</p><p>')
        // Single newline = line break
        .replace(/\n/g, '<br/>');
}

/**
 * Sanitize user input to prevent HTML injection
 */
export function sanitizeHtml(text: string): string {
    if (!text) return '';
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}
