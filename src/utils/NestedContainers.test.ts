import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { applyRedlineToOxml } from './OxmlEngine';

// Polyfill for Node environment
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

console.log("Running Nested Containers Test Suite...");

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

// --- 1. Hyperlink Preservation ---
const hyperlinkOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId1">
        <w:r><w:t>Click Here</w:t></w:r>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>
`;

// Helper to strip tags for text content verification (excluding deletions)
function getTextContent(xml: string): string {
  // Remove deletions first
  const withoutDeletions = xml.replace(/<w:del[^>]*>.*?<\/w:del>/g, '');
  return withoutDeletions.replace(/<[^>]+>/g, '');
}

runTest(
  "Hyperlink Preservation",
  hyperlinkOxml,
  "Click Here",
  "Click Me",
  (res) => {
    // Should preserve w:hyperlink and r:id
    const hasHyperlink = res.includes('<w:hyperlink r:id="rId1">');
    const hasText = getTextContent(res).includes("Click Me");
    return hasHyperlink && hasText;
  }
);

// --- 2. Bookmark Preservation ---
const bookmarkOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="0" w:name="MyBookmark"/>
      <w:r><w:t>Bookmarked Text</w:t></w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
  </w:body>
</w:document>
`;

// Note: Bookmarks are currently NOT extracted as text, so they are invisible to the diff engine unless we tokenize them.
// If we tokenize them, we can preserve them.
runTest(
  "Bookmark Preservation",
  bookmarkOxml,
  "\uFFFCBookmarked Text\uFFFC", // Assuming tokenization
  "\uFFFCModified Text\uFFFC",
  (res) => {
    const hasStart = res.includes('<w:bookmarkStart w:id="0" w:name="MyBookmark"/>');
    const hasEnd = res.includes('<w:bookmarkEnd w:id="0"/>');
    const hasText = getTextContent(res).includes("Modified Text");
    return hasStart && hasEnd && hasText;
  }
);

// --- 3. Text Box Editing ---
const textBoxOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word">
  <w:body>
    <w:p>
      <w:r>
        <w:pict>
          <v:shape>
            <v:textbox>
              <w:txbxContent>
                <w:p><w:r><w:t>Box Content</w:t></w:r></w:p>
              </w:txbxContent>
            </v:textbox>
          </v:shape>
        </w:pict>
      </w:r>
    </w:p>
  </w:body>
</w:document>
`;

// Text Box content should be accessible.
// Currently, w:pict is treated as a Sentinel (Drawing Immunity), so the content is hidden/immutable.
// We want to allow editing INSIDE the box.
runTest(
  "Text Box Editing",
  textBoxOxml,
  "Box Content", // Should be extracted if recursion works
  "\uFFFC\nNew Box Content", // Preserve Sentinel and Newline to keep text inside box
  (res) => {
    // Check if text is present
    const hasText = getTextContent(res).includes("New Box Content");

    // Check if it is INSIDE the text box
    // We look for <w:txbxContent> ... New ... Box Content ... </w:txbxContent>
    // allowing for tags in between
    const match = res.match(/<w:txbxContent>[\s\S]*?New[\s\S]*?Box Content[\s\S]*?<\/w:txbxContent>/);

    return hasText && !!match;
  }
);

console.log(`\nSUMMARY: ${passes} Passed, ${failures} Failed`);
if (failures > 0) process.exit(1);
