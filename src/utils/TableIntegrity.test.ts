import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { applyRedlineToOxml } from './OxmlEngine';

// Polyfill for Node environment
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

console.log("Running Table Integrity Protocol Test Suite...");

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

// --- 1. Single Cell Preservation ---
const tableOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p><w:r><w:t>Cell 1</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>Cell 2</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
`;

runTest(
  "Single Cell Preservation",
  tableOxml,
  "Cell 1\nCell 2",
  "Cell 1 Modified\nCell 2",
  (res) => {
    // Check if structure is preserved:
    // Should still have w:tbl, w:tr, and 2 w:tc
    const hasTable = res.includes("<w:tbl>");
    const tcCount = (res.match(/<w:tc>/g) || []).length;

    // Check if "Cell 1 Modified" is inside a w:tc
    // This is a bit loose, but if the engine flattens it, it might put it directly in body
    // or remove the table entirely if it just appends ps to body.

    // Check if "Cell 1" and "Modified" are present (they might be split)
    return hasTable && tcCount === 2 && res.includes("Cell 1") && res.includes("Modified");
  }
);

// --- 2. Multi-Cell Redline ---
runTest(
  "Multi-Cell Redline",
  tableOxml,
  "Cell 1\nCell 2",
  "Cell 1 Mod\nCell 2 Mod",
  (res) => {
    const tcCount = (res.match(/<w:tc>/g) || []).length;
    return tcCount === 2 && res.includes("Cell 1") && res.includes(" Mod") && res.includes("Cell 2");
  }
);

// --- 3. Topology Lock (No New Cells) ---
runTest(
  "Topology Lock: No New Cells",
  tableOxml,
  "Cell 1\nCell 2",
  "Cell 1\nNew Para\nCell 2",
  (res) => {
    // We inserted a paragraph. It should NOT create a new cell.
    // It should be appended to the first cell (or wherever the split happened).
    // Ideally, "New Para" is in the first cell or second cell, but definitely NOT in a new w:tc.
    const tcCount = (res.match(/<w:tc>/g) || []).length;
    return tcCount === 2;
  }
);

console.log(`\nSUMMARY: ${passes} Passed, ${failures} Failed`);
if (failures > 0) process.exit(1);
