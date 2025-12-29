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

async function runTests() {
    console.log("🚀 Running Commercial Hardening Stress Tests...\n");

    // ==========================================
    // Test L1: Preserve List Structure (Edit Middle Item)
    // ==========================================
    console.log("--- Test L1: Preserve List Structure ---");
    const listOxml = `
    <w:body>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item 1</w:t></w:r></w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item 2</w:t></w:r></w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item 3</w:t></w:r></w:p>
    </w:body>`;
    const originalTextL1 = "Item 1\nItem 2\nItem 3";
    const modifiedTextL1 = "Item 1\nItem Modified\nItem 3";

    const resultL1 = applyRedlineToOxml(listOxml, originalTextL1, modifiedTextL1);

    // Check for 3 paragraphs with numPr
    const numPrCountL1 = (resultL1.oxml.match(/<w:numPr>/g) || []).length;
    assert(numPrCountL1 === 3, `L1: Expected 3 list items, found ${numPrCountL1}`);

    // Check for the word "Modified" (may be in separate element due to diff)
    assert(resultL1.oxml.includes("Modified"), "L1: Word 'Modified' should be in output");

    // Verify track changes are present
    assert(resultL1.oxml.includes("<w:del"), "L1: Should have deletion markup");
    assert(resultL1.oxml.includes("<w:ins"), "L1: Should have insertion markup");

    // ==========================================
    // Test L2: Split List Item (1 bullet → 2 bullets)
    // ==========================================
    console.log("\n--- Test L2: Split List Item ---");
    const singleBulletOxml = `
    <w:body>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Clause 1</w:t></w:r></w:p>
    </w:body>`;
    const originalTextL2 = "Clause 1";
    const modifiedTextL2 = "Clause 1\nClause 2";

    const resultL2 = applyRedlineToOxml(singleBulletOxml, originalTextL2, modifiedTextL2);
    const numPrCountL2 = (resultL2.oxml.match(/<w:numPr>/g) || []).length;
    assert(numPrCountL2 === 2, `L2: Expected 2 list items after split, found ${numPrCountL2}`);
    assert(resultL2.oxml.includes("Clause 2"), "L2: New item 'Clause 2' should be added");

    // ==========================================
    // Test L3: No-Op Detection (Identical Text)
    // ==========================================
    console.log("\n--- Test L3: No-Op Detection ---");
    const inputOxmlL3 = `<w:body><w:p><w:r><w:t>Clause 1</w:t></w:r></w:p></w:body>`;
    const textL3 = "Clause 1";

    const resultL3 = applyRedlineToOxml(inputOxmlL3, textL3, textL3);
    assert(resultL3.hasChanges === false, "L3: Should flag No-Op when text is identical");

    // ==========================================
    // Test L4: Merge List Items (2 bullets → 1 bullet)
    // ==========================================
    console.log("\n--- Test L4: Merge List Items ---");
    const twoBulletsOxml = `
    <w:body>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>First part</w:t></w:r></w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Second part</w:t></w:r></w:p>
    </w:body>`;
    const originalTextL4 = "First part\nSecond part";
    const modifiedTextL4 = "First part Second part"; // Merged (newline deleted)

    const resultL4 = applyRedlineToOxml(twoBulletsOxml, originalTextL4, modifiedTextL4);

    // Count paragraphs (w:p)
    const pCountL4 = (resultL4.oxml.match(/<w:p[ >]/g) || []).length;
    assert(pCountL4 === 1, `L4: Expected 1 paragraph after merge, found ${pCountL4}`);

    // Content should be merged
    assert(resultL4.oxml.includes("First part"), "L4: 'First part' should exist");
    assert(resultL4.oxml.includes("Second part"), "L4: 'Second part' should exist");

    // ==========================================
    // Test L5: Legal Character Fidelity
    // ==========================================
    console.log("\n--- Test L5: Legal Character Fidelity ---");
    const legalCharsOxml = `<w:body><w:p><w:r><w:t>§ 1.1 — Definition (i)</w:t></w:r></w:p></w:body>`;
    const legalText = "§ 1.1 — Definition (i)";
    const modifiedLegalText = "§ 1.2 — Definition (ii)";

    const resultL5 = applyRedlineToOxml(legalCharsOxml, legalText, modifiedLegalText);

    // Verify legal characters survived
    assert(resultL5.oxml.includes("§"), "L5: Section symbol (§) should be preserved");
    assert(resultL5.oxml.includes("—"), "L5: Em-dash (—) should be preserved");

    // Verify the changes were tracked
    assert(resultL5.oxml.includes("1.2") || resultL5.oxml.includes("<w:ins"), "L5: Changes should be tracked");

    // ==========================================
    // Test L6: Nested List (Sub-clause editing)
    // ==========================================
    console.log("\n--- Test L6: Nested List Structure ---");
    const nestedListOxml = `
    <w:body>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>(a) Main clause</w:t></w:r></w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>(i) Sub-clause one</w:t></w:r></w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>(ii) Sub-clause two</w:t></w:r></w:p>
    </w:body>`;
    const originalTextL6 = "(a) Main clause\n(i) Sub-clause one\n(ii) Sub-clause two";
    const modifiedTextL6 = "(a) Main clause\n(i) Sub-clause one\n(ii) Sub-clause modified";

    const resultL6 = applyRedlineToOxml(nestedListOxml, originalTextL6, modifiedTextL6);

    // Verify all 3 paragraphs preserved
    const pCountL6 = (resultL6.oxml.match(/<w:p[ >]/g) || []).length;
    assert(pCountL6 === 3, `L6: Expected 3 paragraphs, found ${pCountL6}`);

    // Verify nested indent levels preserved
    const ilvl0CountL6 = (resultL6.oxml.match(/w:ilvl[^>]*w:val="0"/g) || []).length;
    const ilvl1CountL6 = (resultL6.oxml.match(/w:ilvl[^>]*w:val="1"/g) || []).length;
    assert(ilvl0CountL6 === 1, `L6: Expected 1 level-0 item, found ${ilvl0CountL6}`);
    assert(ilvl1CountL6 === 2, `L6: Expected 2 level-1 items, found ${ilvl1CountL6}`);

    console.log("\n✅ ALL TESTS PASSED: Commercial Hardening Suite Complete");
}

runTests().catch(e => console.error(e));
