import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { applyRedlineToOxml } from './OxmlEngine';

// Polyfill for Node environment
(global as any).DOMParser = DOMParser;
(global as any).XMLSerializer = XMLSerializer;

const mockOxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p>
  <w:r>
    <w:rPr><w:color w:val="0000FF"/></w:rPr>
    <w:t>\${Seller Name}</w:t>
  </w:r>
  <w:r>
    <w:t>, a company</w:t>
  </w:r>
</w:p>
</w:body>
</w:document>
`;

const originalText = "\${Seller Name}, a company";
// Simulate AI returning LaTeX artifacts and changing text
const modifiedText = "$\\text{\${Seller Name}}$, a corporation";

console.log("Running OXML Debug Test (Formatting Preservation)...");

try {
  const result = applyRedlineToOxml(mockOxml, originalText, modifiedText);
  console.log("Result XML:");
  console.log(result);

  // Assertion: The blue color must be preserved for ${Seller Name}
  // The original engine destroys nodes and recreates them, likely losing the color property 
  // unless it happened to be the default (which it isn't here, it's specific to this run).

  if (result.includes('<w:color w:val="0000FF"/>')) {
    console.log("SUCCESS: Blue color tag found.");
  } else {
    console.error("FAILURE: Blue color tag MISSING. Formatting was lost.");
    process.exit(1);
  }

  // Assertion 2: Artifacts must be cleaned
  if (result.includes('$\\text{') || result.includes('}$')) {
    console.error("FAILURE: LaTeX artifacts found in output.");
    process.exit(1);
  } else {
    console.log("SUCCESS: Artifacts cleaned.");
  }

} catch (e) {
  console.error("ERROR:", e);
  process.exit(1);
}
