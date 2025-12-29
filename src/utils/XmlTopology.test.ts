
import { applyRedlineToOxml } from './OxmlEngine';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

// Mock Browser Environment for OxmlEngine
const globalAny: any = global;
globalAny.DOMParser = DOMParser;
globalAny.XMLSerializer = XMLSerializer;

function runTest(testName: string, originalXml: string, originalText: string, modifiedText: string, validator: (resultXml: string) => boolean) {
    console.log(`\n[TEST] ${testName}`);
    try {
        const result = applyRedlineToOxml(originalXml, originalText, modifiedText);
        if (validator(result)) {
            console.log("✅ PASS");
            return true;
        } else {
            console.log("❌ FAIL");
            console.log(`Original: ${originalText}`);
            console.log(`Modified (Input): ${modifiedText}`);
            console.log(`Result XML: \n${result}`);
            return false;
        }
    } catch (e) {
        console.log(`❌ ERROR: ${e}`);
        return false;
    }
}

// Helper to strip tags for text content verification (excluding deletions)
function getTextContent(xml: string): string {
    const withoutDeletions = xml.replace(/<w:del[^>]*>.*?<\/w:del>/g, '');
    return withoutDeletions.replace(/<[^>]+>/g, '');
}

let passes = 0;
let failures = 0;

console.log("Running XML Topology Test Suite...");

// --- 1. The 'Deep Nest' Test (Nested Tables) ---
const deepNestOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p><w:r><w:t>Outer Cell</w:t></w:r></w:p>
          <w:tbl>
            <w:tr>
              <w:tc>
                <w:p><w:r><w:t>Deep Text</w:t></w:r></w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
`;

// Note: The engine extracts text linearly. 
// "Outer Cell" + "\n" + "Deep Text" (if tables add newlines? OxmlEngine adds newlines between paragraphs)
// Let's assume the input text reflects the structure: "Outer Cell\nDeep Text"
if (runTest(
    "Deep Nest Test (Nested Tables)",
    deepNestOxml,
    "Outer Cell\nDeep Text",
    "Outer Cell\nModified Deep Text",
    (res) => {
        // Check for nested structure: w:tbl inside w:tc inside w:tbl
        const hasNestedTable = res.match(/<w:tbl>[\s\S]*?<w:tc>[\s\S]*?<w:tbl>/);
        const hasText = getTextContent(res).includes("Modified Deep Text");
        return !!hasNestedTable && hasText;
    }
)) passes++; else failures++;


// --- 2. The 'Math' Test (MathML Preservation) ---
const mathOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p>
      <w:r><w:t>Formula: </w:t></w:r>
      <m:oMath>
        <m:r><m:t>E=mc^2</m:t></m:r>
      </m:oMath>
    </w:p>
  </w:body>
</w:document>
`;

// Sentinel for Math is \uFFFC
if (runTest(
    "Math Test (MathML Preservation)",
    mathOxml,
    "Formula: \uFFFC",
    "Formula (Einstein): \uFFFC",
    (res) => {
        const hasMath = res.includes('<m:oMath>');
        const hasText = getTextContent(res).includes("Formula (Einstein):");
        return hasMath && hasText;
    }
)) passes++; else failures++;


// --- 3. The 'Text Box' Test (Floating Object Editing) ---
const textBoxOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <w:p>
      <w:r>
        <w:pict>
          <v:shape>
            <v:textbox>
              <w:txbxContent>
                <w:p><w:r><w:t>Original Content</w:t></w:r></w:p>
              </w:txbxContent>
            </v:textbox>
          </v:shape>
        </w:pict>
      </w:r>
    </w:p>
  </w:body>
</w:document>
`;

// Text Box is treated as Sentinel \uFFFC, but content is extracted?
// Wait, in Phase 6 we implemented logic to extract text box content if it's a w:pict/w:txbxContent.
// So original text should be "\uFFFC\nOriginal Content" (Sentinel for the pict + content)
// Actually, the Sentinel logic for Text Box adds the Sentinel AND extracts the content?
// Let's check OxmlEngine.ts logic.
// Logic: "Text Box Sentinel... sentinelMap.push... originalFullText += \uFFFC... propertyMap.push..."
// It does NOT extract the content into originalFullText linearly *at that point*.
// BUT, the recursive traversal `paragraphs.forEach` goes through ALL paragraphs in the body.
// If the text box paragraphs are children of `w:body` (via `getElementsByTagName("w:p")`), they will be processed.
// `getElementsByTagName("w:p")` returns ALL paragraphs, including those nested in text boxes.
// So the order depends on document order.
// The `w:pict` is inside a `w:p`. The `w:txbxContent` contains another `w:p`.
// So we will see the outer `w:p` (containing the Sentinel) AND the inner `w:p` (containing "Original Content").
// The Sentinel represents the *structure*. The inner paragraph represents the *content*.
// This might be tricky. The inner paragraph is physically inside the outer paragraph in XML?
// No, `w:txbxContent` is inside `w:pict` inside `w:r` inside `w:p`.
// So `getElementsByTagName("w:p")` will find the outer `w:p` first, then the inner `w:p`?
// Actually `getElementsByTagName` is depth-first.
// If we are iterating `paragraphs`, we hit the outer `w:p`.
// Inside that, we find `w:pict` (Sentinel). We add `\uFFFC`.
// Then we continue.
// Later in the `paragraphs` array, we will find the inner `w:p`.
// So the text will be "\uFFFC" (from outer) + "\n" + "Original Content" (from inner).
// Let's verify this assumption with the test.

if (runTest(
    "Text Box Test (Floating Object Editing)",
    textBoxOxml,
    "\uFFFC\nOriginal Content",
    "\uFFFC\nNew Content",
    (res) => {
        const hasPict = res.includes('<w:pict>');
        const hasNewContent = getTextContent(res).includes("New Content");
        // Check if New Content is actually inside w:txbxContent
        // Allow for tags between "New" and "Content" (e.g. </w:t>...<w:t>)
        const inTextBox = res.match(/<w:txbxContent>[\s\S]*?New[\s\S]*?Content[\s\S]*?<\/w:txbxContent>/);
        return hasPict && hasNewContent && !!inTextBox;
    }
)) passes++; else failures++;


console.log(`\nSUMMARY: ${passes} Passed, ${failures} Failed`);
if (failures > 0) process.exit(1);
