import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { applyRedlineToOxml } from './OxmlEngine';

// Polyfill for Node environment
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

console.log("Running Sentinel Protections Test Suite...");

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

// --- 1. Drawing Immunity ---
const drawingOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Text Before</w:t></w:r>
      <w:r>
        <w:drawing>
          <wp:inline><wp:docPr id="1" name="Picture 1"/></wp:inline>
        </w:drawing>
      </w:r>
      <w:r><w:t>Text After</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

// Try to delete everything. The drawing should survive.
// Note: The engine currently doesn't see the drawing in "originalText" unless we tokenize it.
// If we don't tokenize it, it might be lost if the paragraph is reconstructed from scratch.
runTest(
    "Drawing Immunity: Delete Surrounding Text",
    drawingOxml,
    "Text Before\uFFFCText After", // Assuming \uFFFC is injected for drawing
    "", // User deletes everything
    (res) => {
        return res.includes("<w:drawing>") && res.includes('name="Picture 1"');
    }
);

// --- 2. Field Preservation ---
const fieldOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText>DATE</w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:body>
</w:document>
`;

runTest(
    "Field Preservation",
    fieldOxml,
    "\uFFFC\uFFFC\uFFFC", // 3 sentinels
    "Modified", // Try to replace field with text
    (res) => {
        // Should preserve the field structure
        return res.includes('<w:fldChar w:fldCharType="begin"/>') &&
            res.includes('<w:instrText>DATE</w:instrText>') &&
            res.includes('<w:fldChar w:fldCharType="end"/>');
    }
);

// --- 3. Section Break Safety ---
const sectionOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
      </w:pPr>
      <w:r><w:t>Section 1</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

runTest(
    "Section Break Safety",
    sectionOxml,
    "\uFFFCSection 1",
    "Modified Section",
    (res) => {
        return res.includes("<w:sectPr>") && res.includes('<w:pgSz w:w="12240" w:h="15840"/>');
    }
);

console.log(`\nSUMMARY: ${passes} Passed, ${failures} Failed`);
if (failures > 0) process.exit(1);
