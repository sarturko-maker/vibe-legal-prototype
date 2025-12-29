import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { applyRedlineToOxml } from './OxmlEngine';

// Polyfill for Node environment
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

console.log("Running High-Fidelity Paragraph Reconstructor Test Suite...");

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

// --- 1. Property Preservation (Indentation & Spacing) ---
const indentedOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:ind w:left="720"/>
        <w:spacing w:after="200"/>
      </w:pPr>
      <w:r><w:t>Indented Text</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

runTest(
    "Preserve Indentation & Spacing",
    indentedOxml,
    "Indented Text",
    "Indented Text Modified",
    (res) => {
        return res.includes('<w:ind w:left="720"/>') && res.includes('<w:spacing w:after="200"/>');
    }
);

// --- 2. Property Preservation (Borders & Shading) ---
const shadedOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:shd w:val="clear" w:color="auto" w:fill="E6E6E6"/>
        <w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>
      </w:pPr>
      <w:r><w:t>Shaded Text</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

runTest(
    "Preserve Borders & Shading",
    shadedOxml,
    "Shaded Text",
    "Shaded Text Updated",
    (res) => {
        return res.includes('<w:shd w:val="clear" w:color="auto" w:fill="E6E6E6"/>') &&
            res.includes('<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>');
    }
);

// --- 3. Numbering Inheritance on Split ---
const numberedOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="1"/>
        </w:numPr>
      </w:pPr>
      <w:r><w:t>Item 1</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

runTest(
    "Numbering Inheritance on Split",
    numberedOxml,
    "Item 1",
    "Item 1\nItem 2",
    (res) => {
        // Should have 2 paragraphs
        const pCount = (res.match(/<w:p>/g) || []).length;
        // Both should have numPr
        const numPrCount = (res.match(/<w:numPr>/g) || []).length;
        return pCount === 2 && numPrCount === 2;
    }
);

// --- 4. Mixed Styles Preservation ---
const mixedOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r><w:t>Title</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="both"/></w:pPr>
      <w:r><w:t>Body</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

runTest(
    "Mixed Styles Preservation",
    mixedOxml,
    "Title\nBody",
    "Title\nBody Modified",
    (res) => {
        // First para should be center, second should be both
        const firstP = res.substring(res.indexOf("<w:p>"), res.indexOf("</w:p>"));
        const secondP = res.substring(res.lastIndexOf("<w:p>"));

        return firstP.includes('<w:jc w:val="center"/>') && secondP.includes('<w:jc w:val="both"/>');
    }
);

console.log(`\nSUMMARY: ${passes} Passed, ${failures} Failed`);
if (failures > 0) process.exit(1);
