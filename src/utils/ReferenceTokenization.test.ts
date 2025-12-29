import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { applyRedlineToOxml } from './OxmlEngine';

// Polyfill for Node environment
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

console.log("Running Reference Tokenization Test Suite...");

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

// --- 1. Footnote Preservation ---
const footnoteOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Text with footnote</w:t></w:r>
      <w:r>
        <w:footnoteReference w:id="2"/>
      </w:r>
    </w:p>
  </w:body>
</w:document>
`;

// Note: We expect the engine to tokenize the footnote reference.
// If implemented correctly, the input text should contain the token.
// But here we are testing the END-TO-END behavior.
// The user provides "Text with footnote" (and implicitly the token if we expose it, but usually the user sees clean text).
// Wait, if we tokenize, does the user SEE the token?
// The prompt says: "Tokenize: Replace tags with unique tokens... Reconstruct: Swap tokens back".
// This usually implies the tokens are internal or exposed to the diff engine.
// If they are exposed to the diff engine, they must be in `originalFullText`.
// And the `modified` text must also contain them if we want to preserve them.
// Let's assume the user (or the UI) handles the token preservation in the modified text, 
// OR the user just edits the text around it.
// For this test, we will assume the "Modified" text includes the token if we want to keep it.
// If the engine tokenizes as `{{__FN_2__}}`, we should include that in our modified string.

// Helper to strip tags for text content verification (excluding deletions)
function getTextContent(xml: string): string {
  // Remove deletions first
  const withoutDeletions = xml.replace(/<w:del[^>]*>.*?<\/w:del>/g, '');
  return withoutDeletions.replace(/<[^>]+>/g, '');
}

runTest(
  "Footnote Preservation",
  footnoteOxml,
  "Text with footnote{{__FN_2__}}",
  "Modified text{{__FN_2__}}",
  (res) => {
    const hasRef = res.includes('<w:footnoteReference w:id="2"/>');
    const hasText = getTextContent(res).includes("Modified text");
    return hasRef && hasText;
  }
);

// --- 2. Endnote Preservation ---
const endnoteOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Endnote here</w:t></w:r>
      <w:r>
        <w:endnoteReference w:id="5"/>
      </w:r>
    </w:p>
  </w:body>
</w:document>
`;

runTest(
  "Endnote Preservation",
  endnoteOxml,
  "Endnote here{{__EN_5__}}",
  "Endnote moved{{__EN_5__}}",
  (res) => {
    const hasRef = res.includes('<w:endnoteReference w:id="5"/>');
    const hasText = getTextContent(res).includes("Endnote moved");
    return hasRef && hasText;
  }
);

// --- 3. Reference Deletion (Audit) ---
runTest(
  "Reference Deletion",
  footnoteOxml,
  "Text with footnote{{__FN_2__}}",
  "Text without footnote",
  (res) => {
    const hasRef = res.includes('<w:footnoteReference w:id="2"/>');
    const hasText = getTextContent(res).includes("Text without footnote");
    return !hasRef && hasText;
  }
);

console.log(`\nSUMMARY: ${passes} Passed, ${failures} Failed`);
if (failures > 0) process.exit(1);
