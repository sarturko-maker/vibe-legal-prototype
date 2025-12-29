
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

console.log("Running Relational Integrity Test Suite...");

// --- 1. The 'Bookmark Span' Test ---
const bookmarkOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="1" w:name="MyBookmark"/>
      <w:r><w:t>Bookmarked Text</w:t></w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>
  </w:body>
</w:document>
`;

// Bookmarks are Sentinels (\uFFFC)
// Original Text: \uFFFC + "Bookmarked Text" + \uFFFC
if (runTest(
    "Bookmark Span Test",
    bookmarkOxml,
    "\uFFFCBookmarked Text\uFFFC",
    "\uFFFCEdited Text\uFFFC",
    (res) => {
        const hasStart = res.includes('<w:bookmarkStart w:id="1" w:name="MyBookmark"/>');
        const hasEnd = res.includes('<w:bookmarkEnd w:id="1"/>');
        // const hasText = getTextContent(res).includes("Edited Text");

        // Verify order: Start -> Text -> End
        // Text might be split or contain deletions.
        // We just check if "Edit" and "ed Text" (or parts of it) appear between Start and End.

        const startPos = res.indexOf('w:bookmarkStart');
        const endPos = res.indexOf('w:bookmarkEnd');

        // Remove tags to check text content *between* markers
        const contentBetween = getTextContent(res.substring(startPos, endPos));
        const hasEditedText = contentBetween.includes("Edited Text") || (contentBetween.includes("Edit") && contentBetween.includes("ed Text"));

        return hasStart && hasEnd && hasEditedText;
    }
)) passes++; else failures++;


// --- 2. The 'Scale' Test (50 Footnotes) ---
let scaleOxmlBody = "";
let scaleOriginalText = "";
let scaleModifiedText = "";

for (let i = 1; i <= 50; i++) {
    scaleOxmlBody += `<w:p><w:r><w:t>Footnote ${i}</w:t></w:r><w:r><w:footnoteReference w:id="${i}"/></w:r></w:p>`;
    // Footnotes are tokenized as {{__FN_i__}} mapped to atomic char
    // But applyRedlineToOxml expects the *original text* to contain the atomic chars?
    // No, applyRedlineToOxml takes `originalText` which is usually extracted from OXML.
    // Wait, `applyRedlineToOxml` *extracts* the text from the OXML itself inside the function.
    // The `originalText` argument is actually IGNORED by `applyRedlineToOxml` in my current implementation?
    // Let's check OxmlEngine.ts.
    // Line 11: export function applyRedlineToOxml(oxml: string, originalText: string, modifiedText: string): string
    // Line 47: let originalFullText = "";
    // ... extraction loop ...
    // The `originalText` param is NOT used for extraction. It is used? 
    // Actually, looking at OxmlEngine.ts, `originalFullText` is built from the OXML.
    // The `originalText` passed in might be used for validation or logging, or it might be completely unused.
    // Let's check if it's used.
    // It seems `originalText` param is NOT used in the logic I wrote!
    // I build `originalFullText` from the OXML.
    // So I don't need to construct `scaleOriginalText` perfectly matching the internal representation.
    // BUT, `modifiedText` MUST match the internal representation (including tokens) if I want to preserve them.
    // Wait, `modifiedText` comes from the AI. The AI sees `{{__FN_1__}}`.
    // So `scaleModifiedText` should contain `{{__FN_i__}}`.
}

const scaleOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${scaleOxmlBody}
  </w:body>
</w:document>
`;

// Construct modified text with tokens
for (let i = 1; i <= 50; i++) {
    scaleModifiedText += `Footnote ${i} Modified{{__FN_${i}__}}\n`;
}

if (runTest(
    "Scale Test (50 Footnotes)",
    scaleOxml,
    "IGNORED", // The engine extracts text from OXML
    scaleModifiedText.trim(),
    (res) => {
        // Check if all 50 references are present
        let allFound = true;
        for (let i = 1; i <= 50; i++) {
            if (!res.includes(`<w:footnoteReference w:id="${i}"/>`)) {
                console.log(`Missing Footnote ${i}`);
                allFound = false;
                break;
            }
        }
        return allFound;
    }
)) passes++; else failures++;


// --- 3. The 'Language' Test (w:lang Preservation) ---
const langOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:lang w:val="en-US"/></w:rPr>
        <w:t>Hello</w:t>
      </w:r>
      <w:r><w:t> </w:t></w:r>
      <w:r>
        <w:rPr><w:lang w:val="de-DE"/></w:rPr>
        <w:t>Welt</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
`;

if (runTest(
    "Language Test (w:lang Preservation)",
    langOxml,
    "Hello Welt",
    "Hi Welt", // Keep "Welt" (German) same, change "Hello" (English) to "Hi"
    (res) => {
        // We expect "Hi" to have en-US and "Welt" to have de-DE
        // The engine might merge them if it's not careful, or if diff is coarse.
        // "Hi" replaces "Hello". "Welt" is Equal.
        // "Hi" should inherit from "Hello" (en-US).
        // "Welt" should preserve "Welt" (de-DE).

        const hasEn = res.includes('w:val="en-US"');
        const hasDe = res.includes('w:val="de-DE"');
        const hasHi = getTextContent(res).includes("Hi");

        return hasEn && hasDe && hasHi;
    }
)) passes++; else failures++;


console.log(`\nSUMMARY: ${passes} Passed, ${failures} Failed`);
if (failures > 0) process.exit(1);
