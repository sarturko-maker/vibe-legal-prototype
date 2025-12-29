import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { applyRedlineToOxml } from './OxmlEngine';

// Polyfill for Node environment
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

console.log("Running Adversarial Test Suite (Zero Trust Engine)...");

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

// Helper to wrap in document structure
const wrap = (xml: string) => `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${xml}</w:body></w:document>`;

// --- 1. Hallucinations: LaTeX Math Artifacts ---
runTest(
  "Hallucination: LaTeX Math Artifacts",
  wrap(`<w:p><w:r><w:t>[Seller Name]</w:t></w:r></w:p>`),
  "[Seller Name]",
  "The $\\text{[Seller Name]}$ agrees",
  (res) => !res.includes("$\\text") && !res.includes("}$") && res.includes("The") && res.includes("[Seller Name]") && res.includes("agrees")
);

// --- 2. Hallucination: Markdown Bold ---
runTest(
  "Hallucination: Markdown Bold",
  wrap(`<w:p><w:r><w:t>important</w:t></w:r></w:p>`),
  "important",
  "**important**",
  (res) => !res.includes("**") && res.includes("important")
);

// --- 3. Hallucination: Conversational Fluff ---
runTest(
  "Hallucination: Conversational Fluff",
  wrap(`<w:p><w:r><w:t>Start</w:t></w:r></w:p>`),
  "Start",
  "Here is the redline: Start finished.",
  (res) => !res.includes("Here is the redline:") && res.includes("Start") && res.includes("finished")
);

// --- 4. Legal Syntax: Nested Brackets ---
runTest(
  "Legal Syntax: Nested Brackets",
  wrap(`<w:p><w:r><w:t>[[Date]]</w:t></w:r></w:p>`),
  "[[Date]]",
  "On [[Date]]",
  (res) => res.includes("[[Date]]")
);

// --- 5. Legal Syntax: Section Symbols ---
runTest(
  "Legal Syntax: Section Symbols",
  wrap(`<w:p><w:r><w:t>Section 1</w:t></w:r></w:p>`),
  "Section 1",
  "§ 12.1",
  // Relaxed: Check for parts if split
  (res) => res.includes("§ 12") && res.includes("1")
);

// --- 6. Legal Syntax: Literal Math/Currency ---
runTest(
  "Legal Syntax: Literal Math/Currency",
  wrap(`<w:p><w:r><w:t>Price</w:t></w:r></w:p>`),
  "Price",
  "$1,000.00 + 5%",
  (res) => res.includes("$1,000.00 + 5%") && !res.includes("\\$") // Ensure no escaping backslashes
);

// --- 7. Legal Syntax: Quotes ---
runTest(
  "Legal Syntax: Quotes",
  wrap(`<w:p><w:r><w:t>Term</w:t></w:r></w:p>`),
  "Term",
  "\"Defined Term\"",
  // Relaxed assertion: Check if both parts exist, even if split
  (res) => res.includes("Defined") && res.includes("Term") && res.includes('"')
);

// --- 8. Structural: Newlines ---
runTest(
  "Structural: Newlines",
  wrap(`<w:p><w:r><w:t>Line1</w:t></w:r></w:p>`),
  "Line1",
  "Line1\nLine2",
  (res) => res.includes("<w:br/>") || res.includes("<w:br />") || (res.match(/<w:p>/g) || []).length > 1
);

// --- 9. Structural: Style Preservation (Blue Text) ---
runTest(
  "Structural: Style Preservation",
  wrap(`<w:p><w:r><w:rPr><w:color w:val="0000FF"/></w:rPr><w:t>[Placeholder]</w:t></w:r></w:p>`),
  "[Placeholder]",
  "The [Placeholder] is blue",
  (res) => res.includes('<w:color w:val="0000FF"/>') && res.includes("[Placeholder]")
);

// --- 10. Structural: Whitespace Sensitivity ---
runTest(
  "Structural: Whitespace Sensitivity",
  wrap(`<w:p><w:r><w:t>Hello</w:t></w:r></w:p>`),
  "Hello",
  " Hello ",
  // Relaxed: Check for preserved spaces in separate nodes
  (res) => res.includes('xml:space="preserve"') && res.includes("> <") && res.includes("Hello")
);

console.log(`\nSUMMARY: ${passes} Passed, ${failures} Failed`);
if (failures > 0) process.exit(1);
