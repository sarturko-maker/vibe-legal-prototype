
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

console.log("Running Layout Fragility Test Suite...");

// --- 1. The 'Signature' Test (Tabs & Leaders) ---
const signatureOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:tabs>
          <w:tab w:val="right" w:leader="dot" w:pos="9350"/>
        </w:tabs>
      </w:pPr>
      <w:r><w:t>Signature:</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

if (runTest(
    "Signature Test (Tabs Preservation)",
    signatureOxml,
    "Signature:",
    "Signed by:",
    (res) => {
        const hasTabs = res.includes('<w:tab w:val="right" w:leader="dot" w:pos="9350"/>');
        const hasText = getTextContent(res).includes("Signed by:");
        return hasTabs && hasText;
    }
)) passes++; else failures++;


// --- 2. The 'Orphan' Test (KeepNext) ---
const orphanOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:keepNext/>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r><w:t>Section 1</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

if (runTest(
    "Orphan Test (KeepNext Preservation)",
    orphanOxml,
    "Section 1",
    "Section 1: Introduction",
    (res) => {
        const hasKeepNext = res.includes('<w:keepNext/>');
        const hasText = getTextContent(res).includes("Section 1: Introduction");
        return hasKeepNext && hasText;
    }
)) passes++; else failures++;


// --- 3. The 'Split List' Test (Numbering Inheritance) ---
const listOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="1"/>
        </w:numPr>
      </w:pPr>
      <w:r><w:t>Item 3.1</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
`;

if (runTest(
    "Split List Test (Numbering Inheritance)",
    listOxml,
    "Item 3.1",
    "Item 3.1\nItem 3.2",
    (res) => {
        // We expect TWO paragraphs, both with numPr
        const paragraphs = res.match(/<w:p>[\s\S]*?<\/w:p>/g) || [];
        if (paragraphs.length < 2) return false;

        const p1 = paragraphs[0];
        const p2 = paragraphs[1]; // The new one

        if (!p1 || !p2) return false;

        const p1HasNum = p1.includes('<w:numId w:val="1"/>');
        const p2HasNum = p2.includes('<w:numId w:val="1"/>');
        const p2HasText = getTextContent(p2).includes("Item 3.2");

        return p1HasNum && p2HasNum && p2HasText;
    }
)) passes++; else failures++;


console.log(`\nSUMMARY: ${passes} Passed, ${failures} Failed`);
if (failures > 0) process.exit(1);
