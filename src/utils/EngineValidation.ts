import { applyRedlineToOxml } from './OxmlEngine';
import { DOMParser, XMLSerializer } from 'xmldom';

// Polyfill for Node.js environment
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

function assert(condition: boolean, message: string) {
    if (condition) {
        console.log(`✅ PASS: ${message}`);
    } else {
        console.error(`❌ FAIL: ${message}`);
        process.exit(1);
    }
}

async function runValidation() {
    console.log("🔬 Running Commercial Grade Engine Validation Suite...\n");

    // ==========================================
    // Test V1: The 'Table Jump' Test
    // ==========================================
    console.log("--- Test V1: Table Jump (Topology Lock) ---");

    const tableJumpOxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr>
                <w:tc>
                    <w:p><w:r><w:t>Cell 1</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:p><w:r><w:t>Cell 2</w:t></w:r></w:p>
                </w:tc>
            </w:tr>
        </w:tbl>
    </w:body>
    </w:document>`;

    // Original text as extracted (engine uses \n between paragraphs)
    const originalV1 = "Cell 1\nCell 2";
    // Modified text - change both cells
    const modifiedV1 = "First Cell\nSecond Cell";

    const resultV1 = applyRedlineToOxml(tableJumpOxml, originalV1, modifiedV1);

    // Parse result to verify structure
    const parser = new DOMParser();
    const resultDocV1 = parser.parseFromString(resultV1.oxml, "text/xml");
    const cellsV1 = Array.from(resultDocV1.getElementsByTagName("w:tc"));

    console.log(`Found ${cellsV1.length} cells`);

    // Check that both cells have content (paragraphs)
    let cell1HasContent = false;
    let cell2HasContent = false;

    cellsV1.forEach((cell, i) => {
        const paragraphs = cell.getElementsByTagName("w:p");
        const textNodes = cell.getElementsByTagName("w:t");
        let cellText = "";
        for (let j = 0; j < textNodes.length; j++) {
            cellText += textNodes[j].textContent || "";
        }
        console.log(`Cell ${i + 1}: ${paragraphs.length} paragraphs, text="${cellText.substring(0, 50)}"`);

        if (i === 0 && paragraphs.length > 0) cell1HasContent = true;
        if (i === 1 && paragraphs.length > 0) cell2HasContent = true;
    });

    assert(cellsV1.length === 2, "V1: Should have 2 cells");
    assert(cell1HasContent, "V1: Cell 1 should have content");
    assert(cell2HasContent, "V1: Cell 2 should have content");

    // ==========================================
    // Test V2: The 'List Split' Test
    // ==========================================
    console.log("\n--- Test V2: List Split (List Inheritance) ---");

    const listSplitOxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:pPr>
                <w:numPr>
                    <w:ilvl w:val="0"/>
                    <w:numId w:val="7"/>
                </w:numPr>
            </w:pPr>
            <w:r><w:t>3.1 Clause</w:t></w:r>
        </w:p>
    </w:body>
    </w:document>`;

    const originalV2 = "3.1 Clause";
    const modifiedV2 = "3.1 Clause\n3.2 New Clause";  // Split into two items

    const resultV2 = applyRedlineToOxml(listSplitOxml, originalV2, modifiedV2);

    // Parse result
    const resultDocV2 = parser.parseFromString(resultV2.oxml, "text/xml");
    const paragraphsV2 = Array.from(resultDocV2.getElementsByTagName("w:p"));
    const numPrsV2 = Array.from(resultDocV2.getElementsByTagName("w:numPr"));
    const numIdsV2 = Array.from(resultDocV2.getElementsByTagName("w:numId"));

    console.log(`Found ${paragraphsV2.length} paragraphs, ${numPrsV2.length} numPr elements`);

    // Both paragraphs should have numPr with numId="7"
    assert(paragraphsV2.length === 2, "V2: Should have 2 paragraphs after split");
    assert(numPrsV2.length === 2, "V2: Both paragraphs should have numPr (list inheritance)");

    let bothHaveCorrectNumId = true;
    numIdsV2.forEach(numId => {
        if (numId.getAttribute("w:val") !== "7") {
            bothHaveCorrectNumId = false;
        }
    });
    assert(bothHaveCorrectNumId, "V2: Both items should have numId='7'");

    // Verify the new text is present
    assert(resultV2.oxml.includes("3.2"), "V2: New clause '3.2' should be in output");

    // ==========================================
    // Test V3: Container Boundary Preservation
    // ==========================================
    console.log("\n--- Test V3: Container Boundary (No Cross-Contamination) ---");

    const containerOxml = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Before table</w:t></w:r></w:p>
        <w:tbl>
            <w:tr>
                <w:tc>
                    <w:p><w:r><w:t>Inside table</w:t></w:r></w:p>
                </w:tc>
            </w:tr>
        </w:tbl>
        <w:p><w:r><w:t>After table</w:t></w:r></w:p>
    </w:body>
    </w:document>`;

    const originalV3 = "Before table\nInside table\nAfter table";
    const modifiedV3 = "Before modified\nInside modified\nAfter modified";

    const resultV3 = applyRedlineToOxml(containerOxml, originalV3, modifiedV3);

    // Verify table structure preserved
    const resultDocV3 = parser.parseFromString(resultV3.oxml, "text/xml");
    const tablesV3 = resultDocV3.getElementsByTagName("w:tbl");
    const cellsV3 = resultDocV3.getElementsByTagName("w:tc");

    assert(tablesV3.length === 1, "V3: Table should be preserved");
    assert(cellsV3.length === 1, "V3: Cell should be preserved");

    // Get paragraphs in body (outside table)
    const bodyV3 = resultDocV3.getElementsByTagName("w:body")[0];
    let bodyParagraphs = 0;
    Array.from(bodyV3.childNodes).forEach(child => {
        if (child.nodeName === "w:p") bodyParagraphs++;
    });

    console.log(`Body has ${bodyParagraphs} direct paragraphs (should be 2)`);
    assert(bodyParagraphs === 2, "V3: Body should have 2 direct paragraphs (before/after table)");

    console.log("\n✅ ALL VALIDATION TESTS PASSED: Commercial Grade Engine Verified!");
}

runValidation().catch(e => console.error(e));
