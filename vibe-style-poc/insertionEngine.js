/**
 * Insertion Engine
 * Handles inserting text and applying styles using Word's native API.
 */

const InsertionEngine = {
    /**
     * Inserts a clause (text + style) after the specified range.
     * @param {Word.RequestContext} context 
     * @param {Word.Range} range The range to insert after (e.g., current selection)
     * @param {Object} clauseData { text: string, styleId: string }
     * @returns {Promise<Word.Paragraph>} The inserted paragraph
     */
    async insertClause(context, range, clauseData) {
        const { text, styleId } = clauseData;

        // 1. Insert the paragraph
        // We use "After" to insert it as a new block following the selection
        const paragraph = range.insertParagraph(text, "After");

        // 2. Apply the style
        // This is the key: we let Word handle the OXML complexity of applying the style
        if (styleId) {
            // Check if it's a standard style or custom
            // Word API expects the "Style ID" (internal name) or "Style Name" (local name)
            // It's generally safer to use the Style ID if known, but the property is .style
            paragraph.style = styleId;
        } else {
            paragraph.style = "Normal"; // Fallback
        }

        // 3. Handle inline formatting (Bold/Italic) if the AI returned markdown
        // For this POC, we'll do a simple pass for **bold** and *italic*
        // Note: This is a basic implementation. A full parser would be better.
        // We need to re-read the text to find ranges, but since we just inserted it,
        // we can assume the text content matches.

        // However, insertParagraph strips formatting usually. 
        // If we want to support bold/italic, we need to parse the text and apply formatting to ranges.
        // For this POC, let's keep it simple: Plain text insertion with Paragraph Style.
        // If the user wants bold, we can add a TODO.

        return paragraph;
    },

    /**
     * Inserts multiple blocks.
     * @param {Word.RequestContext} context 
     * @param {Word.Range} insertionPoint 
     * @param {Array} blocks Array of { text, styleId }
     */
    async insertBlocks(context, insertionPoint, blocks) {
        let currentRange = insertionPoint;

        for (const block of blocks) {
            const p = await this.insertClause(context, currentRange, block);
            // Update insertion point to be the new paragraph, so next one comes after it
            currentRange = p.getRange("End");
            // We need to sync occasionally if many blocks, but for a few it's fine to wait
        }
    }
};

if (typeof module !== 'undefined' && module.exports) {
    module.exports = InsertionEngine;
}
