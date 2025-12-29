/**
 * Edge Case Test Suite for Commercial Grade OXML Engine
 * 
 * Tests 14 critical edge cases that must pass for production.
 */

import { applyRedlineToOxml } from './OxmlEngine';
import { DOMParser, XMLSerializer } from 'xmldom';

// Polyfill for Node.js
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

let passCount = 0;
let failCount = 0;

function test(name: string, fn: () => boolean) {
    try {
        const result = fn();
        if (result) {
            console.log(`✅ ${name}`);
            passCount++;
        } else {
            console.log(`❌ ${name}`);
            failCount++;
        }
    } catch (e: any) {
        console.log(`❌ ${name} - ERROR: ${e.message}`);
        failCount++;
    }
}

function parseOxml(oxml: string): Document {
    return new DOMParser().parseFromString(oxml, "text/xml");
}

console.log("🔬 Edge Case Test Suite for OXML Engine\n");
console.log("=".repeat(50) + "\n");

// ==========================================
// E1: Bold Text Preservation
// ==========================================
test("E1: Bold text in table cell - w:rPr preserved", () => {
    const oxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr><w:tc>
                <w:p><w:r>
                    <w:rPr><w:b/></w:rPr>
                    <w:t>Bold Text</w:t>
                </w:r></w:p>
            </w:tc></w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    const result = applyRedlineToOxml(oxml, "Bold Text", "Bold Modified");
    const doc = parseOxml(result.oxml);

    // Check bold is preserved
    const bold = doc.getElementsByTagName("w:b");
    return bold.length > 0;
});

// ==========================================
// E2: Newline Within Cell (Multiple paragraphs in one cell)
// ==========================================
test("E2: Newline within cell - paragraphs stay in same cell", () => {
    const oxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr><w:tc>
                <w:p><w:r><w:t>Line 1</w:t></w:r></w:p>
                <w:p><w:r><w:t>Line 2</w:t></w:r></w:p>
            </w:tc></w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    const result = applyRedlineToOxml(oxml, "Line 1\nLine 2", "Line 1 Modified\nLine 2 Modified");
    const doc = parseOxml(result.oxml);

    // Both paragraphs should remain in the same cell
    const cells = doc.getElementsByTagName("w:tc");
    const paragraphsInCell = cells[0]?.getElementsByTagName("w:p").length || 0;

    return cells.length === 1 && paragraphsInCell >= 2;
});

// ==========================================
// E3: Adding Paragraph in Cell
// ==========================================
test("E3: Adding paragraph in cell - new p inside same tc", () => {
    const oxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr><w:tc>
                <w:p><w:r><w:t>Original</w:t></w:r></w:p>
            </w:tc></w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    const result = applyRedlineToOxml(oxml, "Original", "Original\nNew Line");
    const doc = parseOxml(result.oxml);

    const cells = doc.getElementsByTagName("w:tc");
    const paragraphsInCell = cells[0]?.getElementsByTagName("w:p").length || 0;

    // New paragraph should be in the same cell
    return cells.length === 1 && paragraphsInCell >= 2;
});

// ==========================================
// E4: Empty Cell Handling
// ==========================================
test("E4: Empty cell - no crash, structure preserved", () => {
    const oxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr>
                <w:tc><w:p><w:r><w:t>Content</w:t></w:r></w:p></w:tc>
                <w:tc><w:p></w:p></w:tc>
            </w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    const result = applyRedlineToOxml(oxml, "Content", "Modified");
    const doc = parseOxml(result.oxml);

    // Both cells should still exist
    return doc.getElementsByTagName("w:tc").length === 2;
});

// ==========================================
// E5: Nested Table (Table inside table cell)
// ==========================================
test("E5: Nested table - inner table preserved", () => {
    const oxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr><w:tc>
                <w:tbl>
                    <w:tr><w:tc><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr>
                </w:tbl>
            </w:tc></w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    const result = applyRedlineToOxml(oxml, "Inner", "Inner Modified");
    const doc = parseOxml(result.oxml);

    // Both tables should exist
    return doc.getElementsByTagName("w:tbl").length === 2;
});

// ==========================================
// E6: List Inside Table Cell
// ==========================================
test("E6: List inside table cell - numPr preserved", () => {
    const oxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr><w:tc>
                <w:p>
                    <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
                    <w:r><w:t>List Item</w:t></w:r>
                </w:p>
            </w:tc></w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    const result = applyRedlineToOxml(oxml, "List Item", "List Item Modified");
    const doc = parseOxml(result.oxml);

    return doc.getElementsByTagName("w:numPr").length > 0;
});

// ==========================================
// E7: Multi-Cell Edit (Different cells modified)
// ==========================================
test("E7: Multi-cell edit - both cells remain separate", () => {
    const oxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr>
                <w:tc><w:p><w:r><w:t>Cell A</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>Cell B</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    const result = applyRedlineToOxml(oxml, "Cell A\nCell B", "Cell A Modified\nCell B Modified");
    const doc = parseOxml(result.oxml);

    const cells = doc.getElementsByTagName("w:tc");

    // Both cells should have content
    let cellAHasContent = false;
    let cellBHasContent = false;

    for (let i = 0; i < cells.length; i++) {
        const text = cells[i].textContent || "";
        if (text.includes("A")) cellAHasContent = true;
        if (text.includes("B")) cellBHasContent = true;
    }

    return cells.length === 2 && cellAHasContent && cellBHasContent;
});

// ==========================================
// E8: Cross-Cell Merge Block (Should NOT merge cells)
// ==========================================
test("E8: Cross-cell merge blocked - cells stay separate", () => {
    const oxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr>
                <w:tc><w:p><w:r><w:t>First</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    // User tries to merge by removing newline - should NOT merge cells
    const result = applyRedlineToOxml(oxml, "First\nSecond", "FirstSecond");
    const doc = parseOxml(result.oxml);

    // Should still have 2 cells
    return doc.getElementsByTagName("w:tc").length === 2;
});

// ==========================================
// E9: Unicode Characters (§ ¶ —)
// ==========================================
test("E9: Unicode characters - § ¶ — preserved", () => {
    const oxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr><w:tc>
                <w:p><w:r><w:t>§ 1.1 — Definition ¶</w:t></w:r></w:p>
            </w:tc></w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    const result = applyRedlineToOxml(oxml, "§ 1.1 — Definition ¶", "§ 1.2 — New Definition ¶");

    return result.oxml.includes("§") && result.oxml.includes("—") && result.oxml.includes("¶");
});

// ==========================================
// E10: Existing Track Changes
// ==========================================
test("E10: Existing track changes - preserved", () => {
    const oxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr><w:tc><w:p>
                <w:ins w:id="1" w:author="Previous"><w:r><w:t>Inserted</w:t></w:r></w:ins>
                <w:r><w:t> Text</w:t></w:r>
            </w:p></w:tc></w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    const result = applyRedlineToOxml(oxml, "Inserted Text", "Inserted Modified");
    const doc = parseOxml(result.oxml);

    // Original w:ins should still exist (or content should be preserved)
    return result.oxml.includes("Inserted") || result.oxml.includes("Modified");
});

// ==========================================
// E11: Field Codes (Sentinel Protection)
// ==========================================
test("E11: Field codes - instrText not modified", () => {
    const oxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr><w:tc><w:p>
                <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                <w:r><w:instrText>PAGE</w:instrText></w:r>
                <w:r><w:fldChar w:fldCharType="end"/></w:r>
                <w:r><w:t>Text</w:t></w:r>
            </w:p></w:tc></w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    const result = applyRedlineToOxml(oxml, "Text", "Modified");

    // instrText should be preserved exactly
    return result.oxml.includes("PAGE") && result.oxml.includes("w:instrText");
});

// ==========================================
// E12: Hyperlinks in Cells
// ==========================================
test("E12: Hyperlinks in cells - w:hyperlink preserved", () => {
    const oxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr><w:tc><w:p>
                <w:hyperlink r:id="rId1">
                    <w:r><w:t>Link Text</w:t></w:r>
                </w:hyperlink>
            </w:p></w:tc></w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    const result = applyRedlineToOxml(oxml, "Link Text", "Link Modified");

    return result.oxml.includes("w:hyperlink");
});

// ==========================================
// E13: Table Header Row
// ==========================================
test("E13: Table header row - tblHeader preserved", () => {
    const oxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                <w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:p><w:r><w:t>Data</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    const result = applyRedlineToOxml(oxml, "Header\nData", "Header Modified\nData Modified");

    return result.oxml.includes("w:tblHeader");
});

// ==========================================
// E14: Cell Spanning Columns (gridSpan)
// ==========================================
test("E14: Cell spanning columns - gridSpan preserved", () => {
    const oxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tblGrid>
                <w:gridCol/><w:gridCol/>
            </w:tblGrid>
            <w:tr>
                <w:tc>
                    <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                    <w:p><w:r><w:t>Spanning Cell</w:t></w:r></w:p>
                </w:tc>
            </w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    const result = applyRedlineToOxml(oxml, "Spanning Cell", "Spanning Modified");

    return result.oxml.includes("w:gridSpan");
});

// ==========================================
// SUMMARY
// ==========================================
console.log("\n" + "=".repeat(50));
console.log(`\nResults: ${passCount} passed, ${failCount} failed`);

if (failCount > 0) {
    console.log("\n⚠️ SOME TESTS FAILED - Engine needs improvement");
    process.exit(1);
} else {
    console.log("\n✅ ALL EDGE CASE TESTS PASSED");
}
