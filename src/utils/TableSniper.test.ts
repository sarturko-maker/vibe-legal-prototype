
import { DOMParser, XMLSerializer } from 'xmldom';
import { diff_match_patch } from 'diff-match-patch';

// Polyfills
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

// Mock diff_match_patch if not available, but let's try to use the real one if we can.
// If the import fails, we might need to install it or mock it. 
// For now, let's assume it's available or we can mock the basic functionality needed.

// --- COPIED LOGIC FROM vibe-legal-v5.yaml (Simplified/Adapted for Test) ---

function applyReconstructionMode(xmlDoc: any, originalText: string, modifiedText: string, serializer: any) {
    // Simplified version of applyReconstructionMode for testing purposes
    // In the real app, this is the full logic. Here we just want to verify the sniper wiring.
    // We'll just do a simple text replacement in the cell for this test.

    // NOTE: In the real implementation, this function does complex diffing and reconstruction.
    // For this test, we'll simulate a successful reconstruction by just replacing the text content.

    const body = xmlDoc.getElementsByTagName("w:body")[0];
    // Clear existing content
    while (body.firstChild) body.removeChild(body.firstChild);

    // Create new paragraph with new text
    const p = xmlDoc.createElement("w:p");
    const r = xmlDoc.createElement("w:r");
    const t = xmlDoc.createElement("w:t");
    t.textContent = modifiedText;
    r.appendChild(t);
    p.appendChild(r);
    body.appendChild(p);

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

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

// --- TESTS ---

console.log("🧪 Testing Table Sniper Mode...");

const tableXml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
    <w:tbl>
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

const edits = [
    { type: "update_cell", location: { row: 0, col: 1 }, content: "UPDATED_R1C2" },
    { type: "update_cell", location: { row: 1, col: 0 }, content: "UPDATED_R2C1" }
];

const parser = new (DOMParser as any)();
const serializer = new (XMLSerializer as any)();
const doc = parser.parseFromString(tableXml, "text/xml");

const result = applyTableSniperMode(doc, edits, serializer);

console.log("Has Changes:", result.hasChanges);

const resultDoc = parser.parseFromString(result.oxml, "text/xml");
const cells = resultDoc.getElementsByTagName("w:tc");

const getCellText = (index: number) => cells[index].textContent?.trim();

console.log("Cell 0 (R1C1):", getCellText(0));
console.log("Cell 1 (R1C2):", getCellText(1));
console.log("Cell 2 (R2C1):", getCellText(2));
console.log("Cell 3 (R2C2):", getCellText(3));

const passed =
    getCellText(0) === "R1C1" &&
    getCellText(1) === "UPDATED_R1C2" &&
    getCellText(2) === "UPDATED_R2C1" &&
    getCellText(3) === "R2C2";

if (passed) {
    console.log("✅ TEST PASSED");
} else {
    console.log("❌ TEST FAILED");
    process.exit(1);
}
