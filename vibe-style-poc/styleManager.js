/**
 * Style Manager
 * Handles extraction of styles from the Word document.
 */

const StyleManager = {
    /**
     * Gets a list of used paragraph styles in the document, plus common defaults.
     * @param {Word.RequestContext} context 
     * @returns {Promise<Array<{id: string, name: string}>>}
     */
    async getParagraphStyles(context) {
        // 1. Scan used styles in the body
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("style, styleId");
        await context.sync();

        const styleMap = new Map();

        // Add standard defaults first (so they appear even if not used)
        styleMap.set("Normal", "Normal");
        styleMap.set("Heading1", "Heading 1");
        styleMap.set("Heading2", "Heading 2");
        styleMap.set("Heading3", "Heading 3");
        styleMap.set("ListParagraph", "List Paragraph");

        // Add used styles
        paragraphs.items.forEach(p => {
            // styleId is the internal ID (e.g., "Heading1"), style is the display name (e.g., "Heading 1")
            if (p.styleId && p.style) {
                styleMap.set(p.styleId, p.style);
            }
        });

        // Convert to array
        return Array.from(styleMap.entries()).map(([id, name]) => ({
            id: id,
            name: name
        }));
    },

    /**
     * Gets a list of character styles (placeholder for now, as scanning runs is expensive).
     * @param {Word.RequestContext} context 
     * @returns {Promise<Array<{id: string, name: string}>>}
     */
    async getCharacterStyles(context) {
        // Scanning every run in the document is too slow for a synchronous-like feel.
        // For this POC, we'll return a static list of common character styles.
        return [
            { id: "Strong", name: "Strong" }, // Bold
            { id: "Emphasis", name: "Emphasis" }, // Italic
            { id: "SubtleEmphasis", name: "Subtle Emphasis" }
        ];
    }
};

// Export for usage in other files if using modules, or just global if vanilla script
if (typeof module !== 'undefined' && module.exports) {
    module.exports = StyleManager;
}
