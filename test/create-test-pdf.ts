/**
 * Create a simple test PDF using PDFKit
 */
import PDFDocument from "pdfkit";
import * as fs from "fs";
import * as path from "path";

async function createTestPdf() {
  const doc = new PDFDocument();
  const outputPath = path.join(__dirname, "test.pdf");
  const writeStream = fs.createWriteStream(outputPath);

  doc.pipe(writeStream);

  doc.fontSize(16).text("Hello, this is a test PDF document.", 100, 100);
  doc.text("Welcome to our translation service.", 100, 130);
  doc.text("This document contains English text.", 100, 160);
  doc.text("The quick brown fox jumps over the lazy dog.", 100, 190);
  doc.text("Thank you for using our service!", 100, 220);

  doc.end();

  return new Promise<void>((resolve, reject) => {
    writeStream.on("finish", () => {
      console.log(`Test PDF created: ${outputPath}`);
      resolve();
    });
    writeStream.on("error", reject);
  });
}

createTestPdf().catch(console.error);
