import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { applyRedlineToOxml } from './OxmlEngine';

// Polyfill for Node environment
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

const mockOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:r>
    <w:rPr><w:color w:val="0000FF"/></w:rPr>
    <w:t>[Seller Name]</w:t>
  </w:r>
</w:p>
</w:body>
</w:document>
`;

const originalText = "[Seller Name]";
// Mock AI response with LaTeX artifacts and extra text
const modifiedText = "The Buyer and $\\text{[Seller Name]}$ agree to terms.";

console.log("Running Commercial-Grade Engine Test...");

try {
    const result = applyRedlineToOxml(mockOxml, originalText, modifiedText);
    console.log("Result XML:");
    console.log(result);

    // Assertion 1: Iron Dome Sanitizer
    if (result.includes('$\\text') || result.includes('}$')) {
        console.error("FAILURE: LaTeX artifacts found. Iron Dome failed.");
        process.exit(1);
    } else {
        console.log("SUCCESS: No LaTeX artifacts found.");
    }

    // Assertion 2: Style-Match Injector
    // [Seller Name] should be preserved in Blue (0000FF)
    // The Diff engine should align [Seller Name] with [Seller Name] (after cleaning)
    // So it should be an EQUAL op, reusing the existing node.

    // We look for the sequence: <w:rPr><w:color w:val="0000FF"/></w:rPr><w:t>[Seller Name]</w:t>
    // Note: The surrounding text "The Buyer and " and " agree to terms." will be insertions.

    if (result.includes('<w:color w:val="0000FF"/>') && result.includes('[Seller Name]')) {
        // Check if they are in the same run (roughly)
        // A strict regex check is better
        // Normalize result by removing newlines for easier regex matching
        const normalizedResult = result.replace(/\n/g, "").replace(/\r/g, "");
        // Robust regex matching
        const colorRegex = /<w:color\s+w:val=["']?0000FF["']?\s*\/?>/i;
        const textRegex = /<w:t[^>]*>.*\[Seller Name\].*<\/w:t>/;

        const hasColor = colorRegex.test(normalizedResult);
        const hasText = textRegex.test(normalizedResult);

        if (hasColor && hasText) {
            // Check if they are close (in the same run)
            // We can check if the substring between color and text is small or just tags
            // But since we saw the output is correct, let's trust the components if they exist.
            // Or use a combined regex with \s*
            const combinedRegex = /<w:r>[\s\S]*?<w:color\s+w:val=["']?0000FF["']?\s*\/?>[\s\S]*?<w:t[^>]*>.*\[Seller Name\].*<\/w:t>[\s\S]*?<\/w:r>/i;
            if (combinedRegex.test(normalizedResult)) {
                console.log("SUCCESS: [Seller Name] preserved with Blue Color.");
            } else {
                console.log("Normalized Result:", normalizedResult);
                console.error("FAILURE: Color and Text found, but not in the same Run.");
                process.exit(1);
            }
        } else {
            console.log("Normalized Result:", normalizedResult);
            console.error(`FAILURE: Missing components. Color: ${hasColor}, Text: ${hasText}`);
            process.exit(1);
        }
    } else {
        console.error("FAILURE: [Seller Name] or Blue Color missing entirely.");
        process.exit(1);
    }

} catch (e) {
    console.error("ERROR:", e);
    process.exit(1);
}
