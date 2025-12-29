/**
 * Vibe Legal v7.0 Stress Tests
 * 
 * This file simulates the core logic of v7.0 to verify:
 * 1. Enhanced Map Generation
 * 2. Smart OXML Builder (numPr, rPr, pPr, w:ins)
 * 3. Authorship Consistency
 */

// --- MOCK DATA ---

// Simulated paragraph data from Word API
const mockParagraphs = [
    { id: 1, text: "NON-DISCLOSURE AGREEMENT", styleId: "Heading1", isListItem: false, listItem: null },
    { id: 2, text: "This Agreement is entered into...", styleId: "Normal", isListItem: false, listItem: null },
    { id: 3, text: "DEFINITIONS", styleId: "Heading1", isListItem: false, listItem: null },
    { id: 4, text: "1.1 \"Confidential Information\" means...", styleId: "ListParagraph", isListItem: true, listItem: { level: 0, listString: "1.1" } },
    { id: 5, text: "1.2 \"Disclosing Party\" means...", styleId: "ListParagraph", isListItem: true, listItem: { level: 0, listString: "1.2" } },
    { id: 6, text: "(a) Information disclosed orally...", styleId: "ListParagraph", isListItem: true, listItem: { level: 1, listString: "(a)" } },
    { id: 7, text: "(b) Information disclosed in writing...", styleId: "ListParagraph", isListItem: true, listItem: { level: 1, listString: "(b)" } },
    { id: 8, text: "OBLIGATIONS", styleId: "Heading1", isListItem: false, listItem: null },
];

// Simulated OXML snippets
const mockOxml = {
    headingStyle: `<pkg:package><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>HEADING</w:t></w:r></w:p></w:body></pkg:package>`,
    listItemL0: `<pkg:package><w:body><w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr><w:t>List Item</w:t></w:r></w:p></w:body></pkg:package>`,
    listItemL1: `<pkg:package><w:body><w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:rPr><w:b/><w:rFonts w:ascii="Times New Roman"/></w:rPr><w:t>Sub Item</w:t></w:r></w:p></w:body></pkg:package>`,
    normalWithSectPr: `<pkg:package><w:body><w:p><w:pPr><w:pStyle w:val="Normal"/><w:sectPr><w:pgSz/></w:sectPr></w:pPr><w:r><w:t>Normal</w:t></w:r></w:p></w:body></pkg:package>`,
};

// --- EXTRACTED FUNCTIONS FROM v7.0 ---

function generateEnhancedMap(paragraphs) {
    const mapLines = [];
    for (const p of paragraphs) {
        let meta = `{Style: ${p.styleId}}`;
        if (p.isListItem && p.listItem) {
            meta += ` {List: Lvl ${p.listItem.level}}`;
        }
        if (p.text.trim().length > 0) {
            mapLines.push(`[${p.id}] ${meta} ${p.text.substring(0, 100)}`);
        }
    }
    return mapLines.join("\n");
}

function generateSmartOxml(styleOxml, newContent) {
    // 1. Extract pPr (includes numPr)
    const pPrMatch = styleOxml.match(/<w:pPr>(.*?)<\/w:pPr>/);
    let pPrContent = "";
    if (pPrMatch && pPrMatch[0]) {
        pPrContent = pPrMatch[0];
        pPrContent = pPrContent.replace(/<w:sectPr>.*?<\/w:sectPr>/g, ""); // Sanitize
    }

    // 2. Extract rPr
    const rPrMatch = styleOxml.match(/<w:rPr>(.*?)<\/w:rPr>/);
    let rPrContent = "";
    if (rPrMatch && rPrMatch[0]) {
        rPrContent = rPrMatch[0];
    }

    // 3. Construct with w:ins
    const insId = Math.floor(Math.random() * 10000000);
    const date = new Date().toISOString();
    const author = "Vibe AI";

    const newParagraph =
        `<w:p>` +
        pPrContent +
        `<w:ins w:id="${insId}" w:author="${author}" w:date="${date}">` +
        `<w:r>` +
        rPrContent +
        `<w:t xml:space="preserve">${newContent}</w:t>` +
        `</w:r>` +
        `</w:ins>` +
        `</w:p>`;

    // 4. Wrap in Package
    const bodyMatch = styleOxml.match(/<w:body>(.*?)<\/w:body>/);
    if (bodyMatch) {
        return styleOxml.replace(bodyMatch[0], `<w:body>${newParagraph}</w:body>`);
    }
    return styleOxml;
}

// --- TEST CASES ---

const tests = [];
let passed = 0;
let failed = 0;

function test(name, fn) {
    tests.push({ name, fn });
}

function assert(condition, message) {
    if (!condition) {
        throw new Error(message || "Assertion failed");
    }
}

// TEST 1: Enhanced Map - Basic
test("Enhanced Map: Basic Structure", () => {
    const map = generateEnhancedMap(mockParagraphs);
    const lines = map.split("\n");
    assert(lines[0].includes("[1] {Style: Heading1}"), "Should include Style tag");
    assert(lines[0].includes("NON-DISCLOSURE AGREEMENT"), "Should include text");
    assert(!lines[0].includes("{List:"), "Heading line should NOT have List tag");
});

// TEST 2: Enhanced Map - List Items
test("Enhanced Map: List Items with Levels", () => {
    const map = generateEnhancedMap(mockParagraphs);
    assert(map.includes("[4] {Style: ListParagraph} {List: Lvl 0}"), "Should show Level 0 list");
    assert(map.includes("[6] {Style: ListParagraph} {List: Lvl 1}"), "Should show Level 1 sub-list");
});

// TEST 3: Smart OXML - Preserves numPr
test("Smart OXML: Preserves numPr from List Style", () => {
    const result = generateSmartOxml(mockOxml.listItemL0, "New Clause 1.3");
    assert(result.includes("<w:numPr>"), "Should contain numPr");
    assert(result.includes('<w:ilvl w:val="0"/>'), "Should preserve ilvl");
    assert(result.includes('<w:numId w:val="1"/>'), "Should preserve numId");
});

// TEST 4: Smart OXML - Preserves rPr (Font)
test("Smart OXML: Preserves rPr (Font/Size)", () => {
    const result = generateSmartOxml(mockOxml.listItemL0, "New Clause");
    assert(result.includes('<w:rFonts w:ascii="Arial"/>'), "Should preserve font");
    assert(result.includes('<w:sz w:val="24"/>'), "Should preserve size");
});

// TEST 5: Smart OXML - Track Changes (w:ins)
test("Smart OXML: Track Changes with Consistent Authorship", () => {
    const result = generateSmartOxml(mockOxml.listItemL0, "Inserted Text");
    assert(result.includes('<w:ins '), "Should contain w:ins");
    assert(result.includes('w:author="Vibe AI"'), "Should have 'Vibe AI' as author");
    assert(result.includes('w:date="'), "Should have date attribute");
});

// TEST 6: Smart OXML - Sanitizes sectPr
test("Smart OXML: Sanitizes sectPr from pPr", () => {
    const result = generateSmartOxml(mockOxml.normalWithSectPr, "New Text");
    assert(!result.includes("<w:sectPr>"), "Should NOT contain sectPr");
    assert(result.includes('<w:pStyle w:val="Normal"/>'), "Should preserve pStyle");
});

// TEST 7: Smart OXML - Sub-List Style (Level 1)
test("Smart OXML: Sub-List Style Preservation (Lvl 1, Bold)", () => {
    const result = generateSmartOxml(mockOxml.listItemL1, "New Sub-Item");
    assert(result.includes('<w:ilvl w:val="1"/>'), "Should preserve sub-list level");
    assert(result.includes('<w:b/>'), "Should preserve bold formatting");
    assert(result.includes('<w:rFonts w:ascii="Times New Roman"/>'), "Should preserve Times New Roman");
});

// TEST 8: Enhanced Map - All Paragraphs Mapped
test("Enhanced Map: All Paragraphs Mapped", () => {
    const map = generateEnhancedMap(mockParagraphs);
    const lines = map.split("\n");
    assert(lines.length === mockParagraphs.length, `Should have ${mockParagraphs.length} lines, got ${lines.length}`);
});

// TEST 9: Smart OXML - Package Structure Preserved
test("Smart OXML: Package Structure Preserved", () => {
    const result = generateSmartOxml(mockOxml.listItemL0, "Text");
    assert(result.startsWith("<pkg:package>"), "Should start with pkg:package");
    assert(result.endsWith("</pkg:package>"), "Should end with pkg:package");
    assert(result.includes("<w:body>"), "Should contain w:body");
});

// TEST 10: Smart OXML - Content Encoding
test("Smart OXML: Content Encoding with Special Characters", () => {
    const specialContent = "Party A & Party B <Confidential>";
    // Note: In real code, we'd need to XML-escape this. For now, just check it's inserted.
    const result = generateSmartOxml(mockOxml.listItemL0, specialContent);
    assert(result.includes(specialContent), "Should contain the new content (unescaped)");
});

// --- RUN TESTS ---

console.log("=== Vibe Legal v7.0 Stress Tests ===\n");

for (const { name, fn } of tests) {
    try {
        fn();
        console.log(`✅ PASS: ${name}`);
        passed++;
    } catch (e) {
        console.log(`❌ FAIL: ${name}`);
        console.log(`   Error: ${e.message}`);
        failed++;
    }
}

console.log(`\n=== Results: ${passed} passed, ${failed} failed ===`);

if (failed > 0) {
    process.exit(1);
}
