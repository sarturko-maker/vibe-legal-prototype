/**
 * Detailed Table Structure Diagnostic
 * 
 * This test outputs the actual OXML structure to verify table integrity.
 */

import { applyRedlineToOxml } from './OxmlEngine';
import { DOMParser, XMLSerializer } from 'xmldom';

// Polyfill for Node.js
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

function formatOxml(oxml: string): string {
    // Simple formatting for readability
    return oxml
        .replace(/></g, '>\n<')
        .split('\n')
        .map((line, i) => {
            const indent = (line.match(/<\//) ? -1 : 0) + (line.match(/<[^/]/) ? 1 : 0);
            return '  '.repeat(Math.max(0, i % 5)) + line.trim();
        })
        .join('\n');
}

console.log("🔬 Table Structure Diagnostic\n");
console.log("=".repeat(60));

// ==========================================
// TEST 1: 2x2 Table - Change text in all cells
// ==========================================
console.log("\n📊 TEST 1: 2x2 Table with edits in all cells\n");

const table2x2 = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
    <w:tbl>
        <w:tblGrid>
            <w:gridCol w:w="4500"/>
            <w:gridCol w:w="4500"/>
        </w:tblGrid>
        <w:tr>
            <w:tc><w:p><w:r><w:t>R1C1</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>R1C2</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
            <w:tc><w:p><w:r><w:t>R2C1</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>R2C2</w:t></w:r></w:p></w:tc>
        </w:tr>
    </w:tbl>
</w:body>
</w:document>`;

const original2x2 = "R1C1\nR1C2\nR2C1\nR2C2";
const modified2x2 = "Row1Col1\nRow1Col2\nRow2Col1\nRow2Col2";

const result1 = applyRedlineToOxml(table2x2, original2x2, modified2x2);

console.log("INPUT OXML (abbreviated):");
console.log("  2 rows, 2 columns, 4 cells");

console.log("\nOUTPUT STRUCTURE:");
const doc1 = new DOMParser().parseFromString(result1.oxml, "text/xml");
const tables1 = doc1.getElementsByTagName("w:tbl");
const rows1 = doc1.getElementsByTagName("w:tr");
const cells1 = doc1.getElementsByTagName("w:tc");
const paragraphs1 = doc1.getElementsByTagName("w:p");

console.log(`  Tables: ${tables1.length} (expected: 1)`);
console.log(`  Rows: ${rows1.length} (expected: 2)`);
console.log(`  Cells: ${cells1.length} (expected: 4)`);
console.log(`  Paragraphs: ${paragraphs1.length} (expected: 4)`);

// Check cell contents
console.log("\nCELL CONTENTS:");
for (let i = 0; i < cells1.length; i++) {
    const cell = cells1[i];
    const paras = cell.getElementsByTagName("w:p").length;
    const text = cell.textContent?.trim() || "(empty)";
    console.log(`  Cell ${i + 1}: ${paras} paragraph(s), text="${text.substring(0, 50)}"`);
}

const test1Pass = tables1.length === 1 && rows1.length === 2 && cells1.length === 4;
console.log(`\n${test1Pass ? '✅' : '❌'} TEST 1: ${test1Pass ? 'PASS' : 'FAIL'}`);

// ==========================================
// TEST 2: Table with mixed body content
// ==========================================
console.log("\n" + "=".repeat(60));
console.log("\n📊 TEST 2: Table between body paragraphs\n");

const mixedContent = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
    <w:p><w:r><w:t>Before Table</w:t></w:r></w:p>
    <w:tbl>
        <w:tr>
            <w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc>
        </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>After Table</w:t></w:r></w:p>
</w:body>
</w:document>`;

const originalMixed = "Before Table\nCell\nAfter Table";
const modifiedMixed = "Before Modified\nCell Modified\nAfter Modified";

const result2 = applyRedlineToOxml(mixedContent, originalMixed, modifiedMixed);

const doc2 = new DOMParser().parseFromString(result2.oxml, "text/xml");
const body2 = doc2.getElementsByTagName("w:body")[0];
const tables2 = doc2.getElementsByTagName("w:tbl");
const cells2 = doc2.getElementsByTagName("w:tc");

// Count direct children of body
let bodyParagraphs2 = 0;
let bodyTables2 = 0;
if (body2) {
    Array.from(body2.childNodes).forEach(child => {
        if (child.nodeName === "w:p") bodyParagraphs2++;
        if (child.nodeName === "w:tbl") bodyTables2++;
    });
}

console.log("OUTPUT STRUCTURE:");
console.log(`  Body direct paragraphs: ${bodyParagraphs2} (expected: 2)`);
console.log(`  Body direct tables: ${bodyTables2} (expected: 1)`);
console.log(`  Table cells: ${cells2.length} (expected: 1)`);

const test2Pass = bodyParagraphs2 === 2 && bodyTables2 === 1 && cells2.length === 1;
console.log(`\n${test2Pass ? '✅' : '❌'} TEST 2: ${test2Pass ? 'PASS' : 'FAIL'}`);

// ==========================================
// TEST 3: The exact Word selection scenario
// ==========================================
console.log("\n" + "=".repeat(60));
console.log("\n📊 TEST 3: Simulating Word table selection (Buyer/Seller)\n");

const wordTable = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
    <w:tbl>
        <w:tblPr>
            <w:tblW w:w="9000" w:type="dxa"/>
        </w:tblPr>
        <w:tblGrid>
            <w:gridCol w:w="4500"/>
            <w:gridCol w:w="4500"/>
        </w:tblGrid>
        <w:tr>
            <w:tc>
                <w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>
                <w:p>
                    <w:pPr><w:pStyle w:val="TableText"/></w:pPr>
                    <w:r>
                        <w:rPr><w:b/></w:rPr>
                        <w:t>Buyer</w:t>
                    </w:r>
                </w:p>
            </w:tc>
            <w:tc>
                <w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>
                <w:p>
                    <w:pPr><w:pStyle w:val="TableText"/></w:pPr>
                    <w:r>
                        <w:rPr><w:b/></w:rPr>
                        <w:t>Seller</w:t>
                    </w:r>
                </w:p>
            </w:tc>
        </w:tr>
        <w:tr>
            <w:tc>
                <w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>
                <w:p><w:r><w:t>John Smith</w:t></w:r></w:p>
            </w:tc>
            <w:tc>
                <w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>
                <w:p><w:r><w:t>Jane Doe</w:t></w:r></w:p>
            </w:tc>
        </w:tr>
    </w:tbl>
</w:body>
</w:document>`;

const originalWord = "Buyer\nSeller\nJohn Smith\nJane Doe";
const modifiedWord = "Purchaser\nSeller\nJohn Smith\nJane Doe";  // Only change Buyer to Purchaser

const result3 = applyRedlineToOxml(wordTable, originalWord, modifiedWord);

const doc3 = new DOMParser().parseFromString(result3.oxml, "text/xml");
const tables3 = doc3.getElementsByTagName("w:tbl");
const rows3 = doc3.getElementsByTagName("w:tr");
const cells3 = doc3.getElementsByTagName("w:tc");

console.log("OUTPUT STRUCTURE:");
console.log(`  Tables: ${tables3.length} (expected: 1)`);
console.log(`  Rows: ${rows3.length} (expected: 2)`);
console.log(`  Cells: ${cells3.length} (expected: 4)`);

console.log("\nCELL CONTENTS:");
for (let i = 0; i < cells3.length; i++) {
    const cell = cells3[i];
    const text = cell.textContent?.replace(/\s+/g, ' ').trim() || "(empty)";
    const hasBold = cell.getElementsByTagName("w:b").length > 0;
    console.log(`  Cell ${i + 1}: "${text}" ${hasBold ? '[BOLD]' : ''}`);
}

// Check for track changes
const inserts = doc3.getElementsByTagName("w:ins");
const deletes = doc3.getElementsByTagName("w:del");
console.log(`\nTRACK CHANGES: ${inserts.length} insertions, ${deletes.length} deletions`);

const test3Pass = tables3.length === 1 && rows3.length === 2 && cells3.length === 4;
console.log(`\n${test3Pass ? '✅' : '❌'} TEST 3: ${test3Pass ? 'PASS' : 'FAIL'}`);

// ==========================================
// SUMMARY
// ==========================================
console.log("\n" + "=".repeat(60));
const allPass = test1Pass && test2Pass && test3Pass;
console.log(`\n${allPass ? '✅' : '❌'} OVERALL: ${allPass ? 'ALL TESTS PASS' : 'SOME TESTS FAILED'}`);

if (!allPass) {
    process.exit(1);
}
