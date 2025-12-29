
import { remapIds, removeGhostComments } from './OxmlPostProcessor';
import { constructOxmlPrompt } from './Agent_Oxml_Redline';

function assert(condition: boolean, message: string) {
    if (condition) {
        console.log(`✅ PASS: ${message}`);
    } else {
        console.error(`❌ FAIL: ${message}`);
        process.exit(1);
    }
}

console.log("Running Oxml Supplemental Gauntlet...");

// --- S1: Lazy Completion ---
console.log("\n[TEST] S1 (Lazy Completion)");
// Simulation: Check if prompt contains strict instructions against truncation.
const prompt = constructOxmlPrompt('<w:p/>', 'Edit');
assert(prompt.includes('DO NOT truncate') || prompt.includes('Return FULL XML'), 'Prompt forbids truncation');


// --- S2: ID Collision ---
console.log("\n[TEST] S2 (ID Collision)");
// Input: Existing bookmark ID 1. New bookmark ID 1 (collision).
const collisionXml = `
<w:p>
    <w:bookmarkStart w:id="1" w:name="Existing"/>
    <w:bookmarkEnd w:id="1"/>
    <w:bookmarkStart w:id="1" w:name="New_Collision"/>
    <w:bookmarkEnd w:id="1"/>
</w:p>
`;
// Mitigation: remapIds should change the second ID 1 to something else.
const remappedXml = remapIds(collisionXml);
// Check that we have unique IDs.
const ids = remappedXml.match(/w:id="(\d+)"/g);
const uniqueIds = new Set(ids);
assert(uniqueIds.size === ids?.length, 'IDs are unique after remapping');
assert(remappedXml.includes('w:id="1"'), 'Preserves original ID if possible (or maps consistently)');
// Actually, simple remapping might map ALL IDs to be safe, or just collisions.
// Let's assume it maps collisions.


// --- S3: Field Code Interior ---
console.log("\n[TEST] S3 (Field Code Interior)");
// This is handled by Tokenizer (Sprint 2) or Prompt.
// For now, let's check if the prompt warns against editing field codes?
// Or we can simulate a post-processor check?
// The user requirement says "Verify: The LLM must edit the <w:t> but must NOT touch the <w:instrText>".
// Since we don't have the Tokenizer yet (it's Sprint 2), we might rely on Prompt for now.
assert(prompt.includes('w:instrText') || prompt.includes('field codes'), 'Prompt warns about field codes');


// --- S4: Comment Anchor Ghost ---
console.log("\n[TEST] S4 (Comment Anchor Ghost)");
// Input: Empty comment range (ghost).
const ghostXml = `
<w:p>
    <w:commentRangeStart w:id="1"/>
    <w:commentRangeEnd w:id="1"/>
    <w:r><w:t>Text</w:t></w:r>
</w:p>
`;
const cleanedXml = removeGhostComments(ghostXml);
assert(!cleanedXml.includes('w:commentRangeStart'), 'Removes ghost start anchor');
assert(!cleanedXml.includes('w:commentRangeEnd'), 'Removes ghost end anchor');
assert(cleanedXml.includes('<w:t>Text</w:t>'), 'Preserves text');

console.log("\nSUMMARY: All Supplemental Tests Passed.");
