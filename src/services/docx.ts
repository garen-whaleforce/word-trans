/**
 * DOCX parsing and writing using JSZip and fast-xml-parser
 */

import * as fs from "fs";
import * as path from "path";
import JSZip from "jszip";
import { XMLParser, XMLBuilder } from "fast-xml-parser";

export interface DocxSegment {
  id: number;
  text: string;
  translated?: string;
}

export interface ParsedDocx {
  zip: JSZip;
  documentXml: any;
  segments: DocxSegment[];
}

// XML parser/builder options for preserving structure
const parserOptions = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  textNodeName: "#text",
  parseTagValue: false,
  trimValues: false,
};

const builderOptions = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  textNodeName: "#text",
  format: false,
  suppressEmptyNode: false,
  suppressBooleanAttributes: false,
};

/**
 * Recursively traverse XML tree and extract text from w:t nodes
 */
function extractSegments(
  node: any,
  segments: DocxSegment[],
  counter: { id: number }
): void {
  if (Array.isArray(node)) {
    for (const item of node) {
      extractSegments(item, segments, counter);
    }
    return;
  }

  if (typeof node !== "object" || node === null) {
    return;
  }

  // Check if this is a w:t node
  if ("w:t" in node) {
    const wtContent = node["w:t"];
    if (Array.isArray(wtContent)) {
      for (const item of wtContent) {
        if (typeof item === "object" && "#text" in item) {
          const text = String(item["#text"]);
          if (text.trim()) {
            segments.push({
              id: counter.id++,
              text: text,
            });
          }
        }
      }
    }
  }

  // Recurse into child nodes
  for (const key of Object.keys(node)) {
    if (key !== ":@" && key !== "#text") {
      extractSegments(node[key], segments, counter);
    }
  }
}

/**
 * Recursively traverse XML tree and replace text in w:t nodes
 */
function replaceSegments(
  node: any,
  segments: DocxSegment[],
  counter: { index: number }
): void {
  if (Array.isArray(node)) {
    for (const item of node) {
      replaceSegments(item, segments, counter);
    }
    return;
  }

  if (typeof node !== "object" || node === null) {
    return;
  }

  // Check if this is a w:t node
  if ("w:t" in node) {
    const wtContent = node["w:t"];
    if (Array.isArray(wtContent)) {
      for (const item of wtContent) {
        if (typeof item === "object" && "#text" in item) {
          const originalText = String(item["#text"]);
          if (originalText.trim()) {
            const segment = segments[counter.index];
            if (segment) {
              // Use translated text if available, otherwise keep original
              item["#text"] = segment.translated ?? segment.text;
              counter.index++;
            }
          }
        }
      }
    }
  }

  // Recurse into child nodes
  for (const key of Object.keys(node)) {
    if (key !== ":@" && key !== "#text") {
      replaceSegments(node[key], segments, counter);
    }
  }
}

/**
 * Parse a DOCX file and extract text segments
 */
export async function parseDocx(filePath: string): Promise<ParsedDocx> {
  // Read the DOCX file
  const buffer = fs.readFileSync(filePath);

  // Load as ZIP
  const zip = await JSZip.loadAsync(buffer);

  // Read word/document.xml
  const documentXmlFile = zip.file("word/document.xml");
  if (!documentXmlFile) {
    throw new Error("Invalid DOCX: word/document.xml not found");
  }

  const documentXmlString = await documentXmlFile.async("string");

  // Parse XML
  const parser = new XMLParser(parserOptions);
  const documentXml = parser.parse(documentXmlString);

  // Extract text segments
  const segments: DocxSegment[] = [];
  const counter = { id: 0 };
  extractSegments(documentXml, segments, counter);

  console.log(`Parsed DOCX: found ${segments.length} text segments`);

  return {
    zip,
    documentXml,
    segments,
  };
}

/**
 * Write translated segments back to DOCX
 */
export async function writeDocx(
  parsed: ParsedDocx,
  outputPath: string
): Promise<void> {
  // Clone the XML structure (deep copy)
  const documentXmlCopy = JSON.parse(JSON.stringify(parsed.documentXml));

  // Replace text with translations
  const counter = { index: 0 };
  replaceSegments(documentXmlCopy, parsed.segments, counter);

  // Build XML string
  const builder = new XMLBuilder(builderOptions);
  const newXmlString = builder.build(documentXmlCopy);

  // Update the zip with new document.xml
  parsed.zip.file("word/document.xml", newXmlString);

  // Generate new DOCX buffer
  const outputBuffer = await parsed.zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });

  // Ensure output directory exists
  const outputDir = path.dirname(outputPath);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // Write to file
  fs.writeFileSync(outputPath, outputBuffer);
  console.log(`Wrote translated DOCX: ${outputPath}`);
}
