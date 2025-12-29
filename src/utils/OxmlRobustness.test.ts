
import { mergeIdenticalRuns, stripRsidAttributes } from './OxmlPreProcessor';
import { correctDoubleEncodedEntities, cloneBulletNumbering } from './OxmlPostProcessor';

function assert(condition: boolean, message: string) {
    if (condition) {
        console.log(`✅ PASS: ${message}`);
    } else {
        console.error(`❌ FAIL: ${message}`);
        process.exit(1);
    }
}

console.log("Running Oxml Robustness Test Suite...");

// --- Pre-Processor Tests ---

// 1. Split Runs (A1)
console.log("\n[TEST] Split Runs (A1)");
const splitRunXml = `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>He</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>llo</w:t></w:r></w:p>`;
const mergedXml = mergeIdenticalRuns(splitRunXml);
// Expect single run with "Hello"
assert(mergedXml.includes('<w:t>Hello</w:t>'), 'Merged text content');
assert((mergedXml.match(/<w:r>/g) || []).length === 1, 'Reduced to single run');

// 2. RSID Drift (B1)
console.log("\n[TEST] RSID Drift (B1)");
const rsidXml = `<w:p w:rsidR="00123456" w:rsidRPr="00ABCDEF"><w:r><w:t>Text</w:t></w:r></w:p>`;
const strippedXml = stripRsidAttributes(rsidXml);
assert(!strippedXml.includes('w:rsidR='), 'Stripped rsidR');
assert(!strippedXml.includes('w:rsidRPr='), 'Stripped rsidRPr');
assert(strippedXml.includes('<w:p>'), 'Preserved element');

// --- Post-Processor Tests ---

// 3. Entity Corruption (C1)
console.log("\n[TEST] Entity Corruption (C1)");
const corruptedXml = `<w:t>A &amp;amp; B &amp;lt; C</w:t>`;
const correctedXml = correctDoubleEncodedEntities(corruptedXml);
assert(correctedXml.includes('A &amp; B'), 'Corrected &amp;');
assert(correctedXml.includes('B &lt; C'), 'Corrected &lt;');

// 4. Bullet Cloning (E2)
console.log("\n[TEST] Bullet Cloning (E2)");
// Input: Para 1 has numPr. Para 2 (new) has no numPr.
const bulletXml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
    <w:p>
        <w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>
        <w:r><w:t>Item 1</w:t></w:r>
    </w:p>
    <w:p>
        <w:r><w:t>Item 2 (New)</w:t></w:r>
    </w:p>
</w:body>
</w:document>
`;
const clonedXml = cloneBulletNumbering(bulletXml);
// Check if second paragraph now has numPr with numId=1
// We need to be careful about matching the *second* paragraph.
// The output should have two numPr elements.
const numPrCount = (clonedXml.match(/<w:numPr>/g) || []).length;
assert(numPrCount === 2, 'Cloned numPr to second paragraph');
assert(clonedXml.includes('<w:t>Item 2 (New)</w:t>'), 'Preserved text');

console.log("\nSUMMARY: All Robustness Tests Passed.");
