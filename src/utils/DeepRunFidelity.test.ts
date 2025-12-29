import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { applyRedlineToOxml } from './OxmlEngine';

// Polyfill for Node environment
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

console.log("Running Deep Run Fidelity Test Suite...");

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

// --- 1. Deep Child Preservation ---
const deepOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:highlight w:val="yellow"/>
          <w:shd w:val="clear" w:color="auto" w:fill="00FF00"/>
          <w:color w:val="FF0000"/>
          <w:rFonts w:ascii="Arial"/>
          <w:vertAlign w:val="superscript"/>
          <w:lang w:val="en-US"/>
        </w:rPr>
        <w:t>Rich Text</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
`;

runTest(
    "Preserve All rPr Children",
    deepOxml,
    "Rich Text",
    "Rich Text Modified",
    (res) => {
        return res.includes('<w:highlight w:val="yellow"/>') &&
            res.includes('<w:shd w:val="clear" w:color="auto" w:fill="00FF00"/>') &&
            res.includes('<w:color w:val="FF0000"/>') &&
            res.includes('<w:rFonts w:ascii="Arial"/>') &&
            res.includes('<w:vertAlign w:val="superscript"/>') &&
            res.includes('<w:lang w:val="en-US"/>');
    }
);

// --- 2. Boundary Inheritance (Insertion between styles) ---
const boundaryOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>Bold</w:t>
      </w:r>
      <w:r>
        <w:rPr><w:i/></w:rPr>
        <w:t>Italic</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
`;

// Insert "New" between "Bold" and "Italic".
// "Bold" ends at index 4. "Italic" starts at index 4.
// Insertion is at index 4.
// Should inherit from "Bold" (Left Neighbor, index 3).
runTest(
    "Boundary Inheritance: Insert Between Styles",
    boundaryOxml,
    "BoldItalic",
    "BoldNewItalic",
    (res) => {
        // "New" should be Bold, not Italic.
        // We check if "New" is inside a run with <w:b/>
        // Simplistic check: find the run containing "New" and check its properties.

        // Note: The XML structure might be complex. 
        // We expect: <w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>
        //            <w:ins ...><w:r><w:rPr><w:b/></w:rPr><w:t>New</w:t></w:r></w:ins>
        //            <w:r><w:rPr><w:i/></w:rPr><w:t>Italic</w:t></w:r>

        const insIndex = res.indexOf('<w:ins');
        const newIndex = res.indexOf('New', insIndex);
        const rPrEndIndex = res.lastIndexOf('</w:rPr>', newIndex);
        const rPrStartIndex = res.lastIndexOf('<w:rPr>', rPrEndIndex);

        const rPrContent = res.substring(rPrStartIndex, rPrEndIndex);

        return rPrContent.includes('<w:b/>') && !rPrContent.includes('<w:i/>');
    }
);

// --- 3. Boundary Inheritance (Start of Paragraph) ---
const startOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>Bold</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
`;

// Insert "New" at start.
// Should inherit from "Bold" (Right Neighbor, index 0).
runTest(
    "Boundary Inheritance: Start of Paragraph",
    startOxml,
    "Bold",
    "NewBold",
    (res) => {
        // "New" should be Bold.
        const insIndex = res.indexOf('<w:ins');
        const newIndex = res.indexOf('New', insIndex);
        const rPrEndIndex = res.lastIndexOf('</w:rPr>', newIndex);
        const rPrStartIndex = res.lastIndexOf('<w:rPr>', rPrEndIndex);

        const rPrContent = res.substring(rPrStartIndex, rPrEndIndex);

        return rPrContent.includes('<w:b/>');
    }
);

console.log(`\nSUMMARY: ${passes} Passed, ${failures} Failed`);
if (failures > 0) process.exit(1);
