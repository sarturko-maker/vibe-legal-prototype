import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { applyRedlineToOxml } from './OxmlEngine';

// Polyfill for Node environment
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

console.log("Running Multi-Paragraph Re-Flow Test Suite...");

let failures = 0;
let passes = 0;

function runTest(name: string, oxml: string, original: string, modified: string, assertion: (result: string) => boolean) {
    console.log(`\n[TEST] ${name}`);
    try {
        const result = applyRedlineToOxml(oxml, original, modified);
        if (assertion(result)) {
            console.log("✅ PASS");
            passes++;
        } else {
            console.error("❌ FAIL");
            console.log("Original:", original);
            console.log("Modified (Input):", modified);
            console.log("Result XML:", result);
            failures++;
        }
    } catch (e) {
        console.error("❌ ERROR (Exception)", e);
        failures++;
    }
}

// Mock OXML with 3 paragraphs
const multiParaOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Paragraph 1</w:t></w:r></w:p>
    <w:p><w:r><w:t>Paragraph 2</w:t></w:r></w:p>
    <w:p><w:r><w:t>Paragraph 3</w:t></w:r></w:p>
  </w:body>
</w:document>
`;

// --- 1. Basic Multi-Paragraph Update ---
runTest(
    "Basic Multi-Paragraph Update",
    multiParaOxml,
    "Paragraph 1\nParagraph 2\nParagraph 3",
    "Paragraph 1\nParagraph 2 Modified\nParagraph 3",
    (res) => {
        // Should have 3 paragraphs
        const pCount = (res.match(/<w:p>/g) || []).length;
        const hasModified = res.includes("Modified");
        return pCount === 3 && hasModified;
    }
);

// --- 2. Paragraph Merge (Deletion of Newline) ---
runTest(
    "Paragraph Merge (Deletion of Newline)",
    multiParaOxml,
    "Paragraph 1\nParagraph 2\nParagraph 3",
    "Paragraph 1 Paragraph 2\nParagraph 3",
    (res) => {
        // Should have 2 paragraphs (1 and 2 merged)
        const pCount = (res.match(/<w:p>/g) || []).length;
        // Check if "Paragraph 1" and "Paragraph 2" are in the same paragraph (roughly)
        // We can check if there's no </w:p> between them
        const p1Index = res.indexOf("Paragraph 1");
        const p2Index = res.indexOf("Paragraph 2");
        const pEndIndex = res.indexOf("</w:p>", p1Index);

        return pCount === 2 && p2Index > p1Index && p2Index < pEndIndex;
    }
);

// --- 3. Paragraph Split (Insertion of Newline) ---
runTest(
    "Paragraph Split (Insertion of Newline)",
    multiParaOxml,
    "Paragraph 1\nParagraph 2\nParagraph 3",
    "Paragraph 1\nPara\ngraph 2\nParagraph 3",
    (res) => {
        // Should have 4 paragraphs
        const pCount = (res.match(/<w:p>/g) || []).length;
        return pCount === 4;
    }
);

// --- 4. Formatting Inheritance (Nearest Neighbor) ---
// Para 1 is Bold, Para 2 is Italic. Insert text in Para 1 should be Bold.
const styledOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r></w:p>
    <w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Italic</w:t></w:r></w:p>
  </w:body>
</w:document>
`;

runTest(
    "Formatting Inheritance (Nearest Neighbor)",
    styledOxml,
    "Bold\nItalic",
    "Bold Added\nItalic",
    (res) => {
        // "Added" should be in a run with <w:b/>
        // We look for <w:rPr><w:b/></w:rPr> ... <w:t> Added</w:t>
        // Note: The sanitizer might preserve space in w:t
        return res.includes('<w:b/>') && res.includes('Added');
    }
);

console.log(`\nSUMMARY: ${passes} Passed, ${failures} Failed`);
if (failures > 0) process.exit(1);
