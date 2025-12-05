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
  // Track the w:t nodes this segment spans
  wtNodes: any[];
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
 * Collect all w:t text nodes from a paragraph or run
 */
function collectWtNodes(node: any, wtNodes: any[]): void {
  if (Array.isArray(node)) {
    for (const item of node) {
      collectWtNodes(item, wtNodes);
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
          wtNodes.push(item);
        }
      }
    }
    return;
  }

  // Recurse into child nodes
  for (const key of Object.keys(node)) {
    if (key !== ":@" && key !== "#text") {
      collectWtNodes(node[key], wtNodes);
    }
  }
}

/**
 * Extract segments by paragraph (w:p) for better context
 */
function extractSegmentsByParagraph(
  node: any,
  segments: DocxSegment[],
  counter: { id: number }
): void {
  if (Array.isArray(node)) {
    for (const item of node) {
      extractSegmentsByParagraph(item, segments, counter);
    }
    return;
  }

  if (typeof node !== "object" || node === null) {
    return;
  }

  // Check if this is a w:p (paragraph) node
  if ("w:p" in node) {
    const wtNodes: any[] = [];
    collectWtNodes(node["w:p"], wtNodes);

    if (wtNodes.length > 0) {
      // Merge all text in this paragraph
      const mergedText = wtNodes.map((wt) => String(wt["#text"])).join("");

      if (mergedText.trim()) {
        segments.push({
          id: counter.id++,
          text: mergedText,
          wtNodes: wtNodes,
        });
      }
    }
    return; // Don't recurse into paragraph children (already processed)
  }

  // Recurse into other nodes
  for (const key of Object.keys(node)) {
    if (key !== ":@" && key !== "#text") {
      extractSegmentsByParagraph(node[key], segments, counter);
    }
  }
}

/**
 * Distribute translated text back to w:t nodes
 */
function distributeTranslation(segment: DocxSegment): void {
  if (!segment.translated || segment.wtNodes.length === 0) {
    return;
  }

  const translated = segment.translated;

  if (segment.wtNodes.length === 1) {
    // Simple case: single w:t node
    segment.wtNodes[0]["#text"] = translated;
  } else {
    // Multiple w:t nodes: put all text in first node, clear others
    segment.wtNodes[0]["#text"] = translated;
    for (let i = 1; i < segment.wtNodes.length; i++) {
      segment.wtNodes[i]["#text"] = "";
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

  // Extract text segments by paragraph
  const segments: DocxSegment[] = [];
  const counter = { id: 0 };
  extractSegmentsByParagraph(documentXml, segments, counter);

  console.log(`Parsed DOCX: found ${segments.length} paragraph segments`);

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
  // Apply translations to the original XML structure
  for (const segment of parsed.segments) {
    distributeTranslation(segment);
  }

  // Build XML string
  const builder = new XMLBuilder(builderOptions);
  const newXmlString = builder.build(parsed.documentXml);

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
