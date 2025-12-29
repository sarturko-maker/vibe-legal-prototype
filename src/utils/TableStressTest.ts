
import { DOMParser, XMLSerializer } from 'xmldom';

// Polyfills
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

// --- MOCK RECONSTRUCTION (Simplified) ---
// We use a simplified version because we are testing the SNIPER wiring, not the reconstruction engine itself.
function applyReconstructionMode(xmlDoc: any, originalText: string, modifiedText: string, serializer: any) {
    const body = xmlDoc.getElementsByTagName("w:body")[0];
    while (body.firstChild) body.removeChild(body.firstChild);

    const p = xmlDoc.createElement("w:p");
    const r = xmlDoc.createElement("w:r");
    const t = xmlDoc.createElement("w:t");
    t.textContent = modifiedText;
    r.appendChild(t);
    p.appendChild(r);
    body.appendChild(p);

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

// --- SNIPER LOGIC (Copied from vibe-legal-v5.yaml) ---
function applyTableSniperMode(xmlDoc: any, edits: any[], serializer: any) {
    let hasChanges = false;
    const tables = xmlDoc.getElementsByTagName("w:tbl");

    const findCell = (tableIndex: number, rowIndex: number, colIndex: number) => {
        if (tableIndex >= tables.length) return null;
        const table = tables[tableIndex] as any;
        const rows = Array.from(table.getElementsByTagName("w:tr"));
        if (rowIndex >= rows.length) return null;
        const row = rows[rowIndex] as any;
        const cells = Array.from(row.getElementsByTagName("w:tc"));
        if (colIndex >= cells.length) return null;
        return cells[colIndex] as any;
    };

    for (const edit of edits) {
        if (edit.type === "update_cell") {
            const cell = findCell(edit.location.table || 0, edit.location.row, edit.location.col);
            if (cell) {
                const currentText = cell.textContent || "";
                const newText = edit.content;

                let cellContentXml = "";
                Array.from(cell.childNodes).forEach((child: any) => {
                    cellContentXml += serializer.serializeToString(child);
                });

                const fakeDocStr = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${cellContentXml}</w:body></w:document>`;

                const result = applyReconstructionMode(
                    new (DOMParser as any)().parseFromString(fakeDocStr, "text/xml"),
                    currentText,
                    newText,
                    serializer
                );

                if (result.hasChanges) {
                    const resultDoc = new (DOMParser as any)().parseFromString(result.oxml, "text/xml");
                    const resultBody = resultDoc.getElementsByTagName("w:body")[0];

                    if (resultBody) {
                        while (cell.firstChild) cell.removeChild(cell.firstChild);
                        Array.from(resultBody.childNodes).forEach((child: any) => {
                            const importedNode = xmlDoc.importNode(child, true);
                            cell.appendChild(importedNode);
                        });
                        hasChanges = true;
                    }
                }
            }
        }
    }

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges };
}

// --- STRESS TESTS ---

console.log("🔥 Stress Testing Table Sniper Mode...\n");

const parser = new (DOMParser as any)();
const serializer = new (XMLSerializer as any)();

// TEST 1: Merged Cells (Grid Span)
// Row 1: [A] [B] [C]
// Row 2: [D (span 2)] [E]
console.log("TEST 1: Merged Cells (Grid Span)");
const mergedTableXml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
    <w:tbl>
        <w:tr>
            <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
            <w:tc>
                <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                <w:p><w:r><w:t>D</w:t></w:r></w:p>
            </w:tc>
            <w:tc><w:p><w:r><w:t>E</w:t></w:r></w:p></w:tc>
        </w:tr>
    </w:tbl>
</w:body>
</w:document>`;

// Target "E". In Row 2, it is the 2nd cell (index 1).
// Even though visually it might be in col 3, in OXML structure it is the 2nd <w:tc>.
const mergedEdits = [
    { type: "update_cell", location: { row: 1, col: 1 }, content: "UPDATED_E" }
];

const mergedDoc = parser.parseFromString(mergedTableXml, "text/xml");
const mergedResult = applyTableSniperMode(mergedDoc, mergedEdits, serializer);
const mergedResultDoc = parser.parseFromString(mergedResult.oxml, "text/xml");
const mergedRows = mergedResultDoc.getElementsByTagName("w:tr");
const row2Cells = mergedRows[1].getElementsByTagName("w:tc");
const cellE = row2Cells[1].textContent;

console.log(`  Target Cell Content: "${cellE}"`);
if (cellE === "UPDATED_E") console.log("  ✅ Merged Cell Access Passed");
else console.log("  ❌ Merged Cell Access Failed");


// TEST 2: Nested Table
// Cell contains a table.
console.log("\nTEST 2: Nested Table");
const nestedTableXml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
    <w:tbl>
        <w:tr>
            <w:tc>
                <w:p><w:r><w:t>Outer Cell</w:t></w:r></w:p>
                <w:tbl>
                    <w:tr><w:tc><w:p><w:r><w:t>Inner Cell</w:t></w:r></w:p></w:tc></w:tr>
                </w:tbl>
            </w:tc>
        </w:tr>
    </w:tbl>
</w:body>
</w:document>`;

// Target the Outer Cell.
// WARNING: This will likely wipe the inner table if the LLM doesn't reproduce it (which it won't in this mock).
// This test confirms that we CAN target the cell, even if the result is destructive to the nested table (expected behavior for now).
const nestedEdits = [
    { type: "update_cell", location: { row: 0, col: 0 }, content: "UPDATED_OUTER" }
];

const nestedDoc = parser.parseFromString(nestedTableXml, "text/xml");
const nestedResult = applyTableSniperMode(nestedDoc, nestedEdits, serializer);
const nestedResultDoc = parser.parseFromString(nestedResult.oxml, "text/xml");
const nestedCell = nestedResultDoc.getElementsByTagName("w:tc")[0];
const nestedContent = nestedCell.textContent;
const hasInnerTable = nestedCell.getElementsByTagName("w:tbl").length > 0;

console.log(`  Cell Content: "${nestedContent}"`);
console.log(`  Has Inner Table: ${hasInnerTable}`);

if (nestedContent === "UPDATED_OUTER" && !hasInnerTable) {
    console.log("  ✅ Nested Table Overwritten (Expected behavior for simple text replacement)");
} else {
    console.log("  ❓ Unexpected behavior");
}

console.log("\nDone.");
