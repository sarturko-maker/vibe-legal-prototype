
import { determineIntent, SelectionState, DocumentContext, RouterResult } from './IntentRouter';

function assert(condition: boolean, message: string) {
    if (condition) {
        console.log(`✅ PASS: ${message}`);
    } else {
        console.error(`❌ FAIL: ${message}`);
        process.exit(1);
    }
}

console.log("Running Intent Router Test Suite...");

const emptySelection: SelectionState = { isEmpty: true, text: '' };
const textSelection: SelectionState = { isEmpty: false, text: 'Selected Text' };
const context: DocumentContext = {
    headers: ['Definitions', 'Indemnification', 'General Provisions', 'Miscellaneous'],
    cursorSection: 'Indemnification'
};

// R1: Nuclear Safety
console.log("\n[TEST] R1 (Nuclear Safety)");
const r1 = determineIntent('Delete everything', emptySelection, context);
assert(r1.intent === 'BLOCK', 'Blocks "Delete everything" with no selection');
assert(r1.message?.includes('Nuclear Safety') || false, 'Returns safety message');

const r1_safe = determineIntent('Delete everything', textSelection, context);
assert(r1_safe.intent === 'COMMAND', 'Allows "Delete everything" WITH selection');


// R2: Context Switch
console.log("\n[TEST] R2 (Context Switch)");
const r2_q = determineIntent('What is the indemnity?', emptySelection, context);
assert(r2_q.intent === 'QUESTION', 'Detects QUESTION');

const r2_cmd_sel = determineIntent('Make it mutual', textSelection, context);
assert(r2_cmd_sel.intent === 'COMMAND' && r2_cmd_sel.scope === 'SELECTION', 'Detects COMMAND + SELECTION');

const r2_cmd_doc = determineIntent('Make it mutual', emptySelection, context);
assert(r2_cmd_doc.intent === 'COMMAND' && r2_cmd_doc.scope === 'WHOLE_DOC', 'Detects COMMAND + WHOLE_DOC');


// R3: Auto-Place
console.log("\n[TEST] R3 (Auto-Place)");
const r3 = determineIntent('Add a Force Majeure clause', emptySelection, context);
assert(r3.intent === 'COMMAND', 'Detects COMMAND');
assert(r3.scope === 'AUTO_PLACE', 'Detects AUTO_PLACE');
assert(r3.targetSection === 'Miscellaneous' || r3.targetSection === 'General Provisions', `Targets correct section (Got: ${r3.targetSection})`);

const r3_explicit = determineIntent('Add to Definitions', emptySelection, context);
assert(r3_explicit.scope === 'AUTO_PLACE', 'Detects explicit AUTO_PLACE');
assert(r3_explicit.targetSection === 'Definitions', 'Targets explicit section');


// R4: Ambiguity
console.log("\n[TEST] R4 (Ambiguity)");
const r4 = determineIntent('Missing definitions', emptySelection, context);
if (r4.intent !== 'CLARIFY') console.log(`Got Intent: ${r4.intent}, Scope: ${r4.scope}`);
assert(r4.intent === 'CLARIFY', 'Detects Ambiguity');
assert(r4.message?.includes('Ambiguous') || false, 'Returns clarification message');


// R5: Read vs Write
console.log("\n[TEST] R5 (Read vs Write)");
const r5 = determineIntent('Summarise the indemnity clause', emptySelection, context);
assert(r5.intent === 'QUESTION', 'Summarise -> QUESTION');

console.log("\nSUMMARY: All Intent Router Tests Passed.");
