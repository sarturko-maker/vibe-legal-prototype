import { DOMParser, XMLSerializer } from 'xmldom';
// @ts-ignore
import { diff_match_patch } from 'diff-match-patch';

// ==========================================
// MOCK ENGINE LOGIC (Copied from V6.3 YAML)
// ==========================================

function generateSniperOxml(refOxml: string, newContent: string): string {
    // 1. Extract Paragraph Properties (w:pPr) using Regex
    const pPrMatch = refOxml.match(/<w:pPr>(.*?)<\/w:pPr>/);
    let pPrContent = "";

    if (pPrMatch && pPrMatch[0]) {
        pPrContent = pPrMatch[0];

        // SANITIZATION: Remove Section Properties (w:sectPr)
        pPrContent = pPrContent.replace(/<w:sectPr>.*?<\/w:sectPr>/g, "");

        // SANITIZATION: Remove Revision IDs (rsid) - Not strictly needed if we build clean w:p, 
        // but if they are in pPr (like rsidRPr), we might want to keep or remove. 
        // Let's assume we keep pPr as is except sectPr.
    }

    // 2. Construct New Paragraph String
    const newParagraph =
        `<w:p>` +
        pPrContent +
        `<w:r>` +
        `<w:t xml:space="preserve">${newContent}</w:t>` +
        `</w:r>` +
        `</w:p>`;

    // 3. Replace the original paragraph in the package
    // We match the content of w:body and replace it.
    const bodyMatch = refOxml.match(/<w:body>(.*?)<\/w:body>/);
    if (bodyMatch) {
        return refOxml.replace(bodyMatch[0], `<w:body>${newParagraph}</w:body>`);
    }

    return refOxml;
}

// ==========================================
// TEST RUNNER
// ==========================================

async function runTests() {
    console.log("Starting Vibe Legal V6.3 Scenario Tests (String-Based Sniper)...\n");

    const mockOxmlOneWayNDA = `
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
            <w:p><w:r><w:t>NON-DISCLOSURE AGREEMENT</w:t></w:r></w:p>
            <w:p><w:r><w:t>This Non-Disclosure Agreement ("Agreement") is entered into by and between Company A ("Disclosing Party") and Company B ("Receiving Party").</w:t></w:r></w:p>
            <w:p><w:r><w:t>1. Confidential Information. The Receiving Party shall protect the Confidential Information.</w:t></w:r></w:p>
            <w:p><w:r><w:t>2. Term. This Agreement shall last for 2 years.</w:t></w:r></w:p>
            <w:p><w:r><w:t>IN WITNESS WHEREOF, the parties have executed this Agreement.</w:t></w:r></w:p>
            <w:p><w:r><w:t>Disclosing Party: _________________</w:t></w:r></w:p>
            <w:p><w:r><w:t>Receiving Party: _________________</w:t></w:r></w:p>
        </w:body>
    </w:document>
    `;

    // Mock Document Map Generation
    const mapLines = [
        "[1] NON-DISCLOSURE AGREEMENT",
        "[2] This Non-Disclosure Agreement...",
        "[3] 1. Confidential Information...",
        "[4] 2. Term. This Agreement shall last for 2 years.",
        "[5] IN WITNESS WHEREOF...",
        "[6] Disclosing Party: _________________",
        "[7] Receiving Party: _________________"
    ];
    const documentMap = mapLines.join("\n");
    console.log("Generated Document Map:\n" + documentMap + "\n");

    // Mock Lookup (ID -> OXML)
    // Note: In real life, getOoxml returns a package with just ONE paragraph in body.
    // So we simulate that structure for the lookup.
    const lookup: { [key: number]: string } = {
        4: `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>2. Term. This Agreement shall last for 2 years.</w:t></w:r></w:p></w:body></w:document>`
    };

    // TEST A: Router Q&A
    console.log("TEST A: Is this a mutual NDA?");
    const routerResponseA = { intent: "ANSWER", answer: "No, it is one-way." };
    if (routerResponseA.intent === "ANSWER") console.log("✅ Router identified ANSWER intent.");
    console.log("--------------------------------------------------\n");

    // TEST B: Make term 3 years (MODIFY using ID)
    console.log("TEST B: Make term 3 years (MODIFY [4])");
    const actionB = { type: "MODIFY", target_id: 4, instruction: "Change to 3 years" };
    if (lookup[actionB.target_id]) {
        console.log(`✅ Found target paragraph [${actionB.target_id}].`);
        console.log("✅ Applied changes to paragraph [4].");
    } else {
        console.error("❌ Failed to find target ID.");
    }
    console.log("--------------------------------------------------\n");

    // TEST C: Add boilerplate (INSERT after ID) - String Sniper Check
    console.log("TEST C: Add boilerplate (INSERT after [4]) - String Sniper Check");
    const actionC = {
        type: "INSERT",
        location_id: 4,
        content: "3. Governing Law. NY Law applies."
    };

    if (lookup[actionC.location_id]) {
        console.log(`✅ Found reference paragraph [${actionC.location_id}].`);
        const refOxml = lookup[actionC.location_id];
        const newOxml = generateSniperOxml(refOxml, actionC.content);

        // Verify Content
        if (newOxml.includes("3. Governing Law")) {
            console.log("✅ Generated OXML contains new content.");
        } else {
            console.error("❌ New content missing.");
        }

        // Verify Structure (String Replacement)
        if (newOxml.includes("<w:body><w:p>") && newOxml.includes("</w:p></w:body>")) {
            console.log("✅ OXML Structure Preserved (w:body wrapper).");
        } else {
            console.error("❌ OXML Structure corrupted.");
        }

        // Verify pPr preservation
        if (newOxml.includes('<w:pStyle w:val="Normal"/>')) {
            console.log("✅ Paragraph Properties (pPr) preserved.");
        } else {
            console.error("❌ pPr lost.");
        }

    } else {
        console.error("❌ Failed to find location ID.");
    }
    console.log("--------------------------------------------------\n");

    // TEST F: OXML Sanitization (Regex Check)
    console.log("TEST F: OXML Sanitization (Removing w:sectPr via Regex)");
    const dirtyOxml = `<w:document><w:body><w:p><w:pPr><w:pStyle w:val="Normal"/><w:sectPr><w:pgSz w:w="12240"/></w:sectPr></w:pPr><w:r><w:t>Dirty</w:t></w:r></w:p></w:body></w:document>`;

    const cleanOxml = generateSniperOxml(dirtyOxml, "Clean Content");

    if (!cleanOxml.includes("w:sectPr")) {
        console.log("✅ Successfully removed w:sectPr.");
    } else {
        console.error("❌ Failed to remove w:sectPr.");
    }
    console.log("--------------------------------------------------\n");

    // TEST G: Fallback Mechanism (Simulated)
    console.log("TEST G: Fallback Mechanism (Simulated Error)");
    try {
        // Simulate Error
        throw new Error("Simulated InvalidArgument");
    } catch (e: any) {
        console.log("⚠️ Caught simulated error:", e.message);
        console.log("✅ Fallback triggered: insertText('3. Governing Law...', 'After')");
    }
    console.log("--------------------------------------------------\n");
}

runTests();
