/**
 * Create a test DOCX file with English content
 */
import JSZip from "jszip";
import * as fs from "fs";
import * as path from "path";

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello, this is a test document.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Welcome to our translation service.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>This document contains English text that should be translated to Traditional Chinese.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>The quick brown fox jumps over the lazy dog.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Thank you for using our service!</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

async function createTestDocx() {
  const zip = new JSZip();

  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file("_rels/.rels", relsXml);
  zip.file("word/document.xml", documentXml);

  const buffer = await zip.generateAsync({ type: "nodebuffer" });

  const testDir = path.join(__dirname);
  if (!fs.existsSync(testDir)) {
    fs.mkdirSync(testDir, { recursive: true });
  }

  const outputPath = path.join(testDir, "test.docx");
  fs.writeFileSync(outputPath, buffer);
  console.log(`Test DOCX created: ${outputPath}`);
}

createTestDocx().catch(console.error);
