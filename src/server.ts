/**
 * Express server with file upload and translation routes
 */

import "dotenv/config";
import express, { Request, Response } from "express";
import multer from "multer";
import path from "path";
import fs from "fs";
import { v4 as uuidv4 } from "uuid";

import {
  createJob,
  getJob,
  cancelJob,
  updateJob,
  finishJob,
  getElapsedSeconds,
  JobState,
} from "./jobs";
import { convertPdfToDocx } from "./services/adobe";
import { parseDocx, writeDocx } from "./services/docx";
import { translateSegments, qaAndRetranslate } from "./services/translator";

const app = express();
const PORT = process.env.PORT || 3000;

// Ensure directories exist
const UPLOAD_DIR = path.join(__dirname, "..", "uploads");
const WORK_DIR = path.join(__dirname, "..", "work");
const OUTPUT_DIR = path.join(__dirname, "..", "output");

[UPLOAD_DIR, WORK_DIR, OUTPUT_DIR].forEach((dir) => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, UPLOAD_DIR);
  },
  filename: (req, file, cb) => {
    const uniqueName = `${Date.now()}-${uuidv4()}${path.extname(file.originalname)}`;
    cb(null, uniqueName);
  },
});

const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === ".pdf" || ext === ".docx") {
      cb(null, true);
    } else {
      cb(new Error("Only .pdf and .docx files are allowed"));
    }
  },
  limits: {
    fileSize: 50 * 1024 * 1024, // 50MB limit
  },
});

// Serve static frontend
app.use(express.static(path.join(__dirname, "..", "public")));
app.use(express.json());

/**
 * Process a job in the background
 */
async function processJob(job: JobState, uploadedFilePath: string): Promise<void> {
  const ext = path.extname(job.fileName).toLowerCase();
  const baseName = path.basename(job.fileName, ext);

  try {
    job.startedAt = Date.now();
    let workingDocxPath: string;

    // Step 1: Convert PDF to DOCX if needed
    if (ext === ".pdf") {
      updateJob(job, {
        status: "converting",
        stepMessage: "正在將 PDF 轉換為 DOCX...",
        progress: 5,
      });

      workingDocxPath = path.join(WORK_DIR, `${uuidv4()}.docx`);
      await convertPdfToDocx(uploadedFilePath, workingDocxPath);

      if (job.cancelled) {
        console.log("Job cancelled after PDF conversion");
        return;
      }
    } else {
      // DOCX file - copy to work directory
      workingDocxPath = path.join(WORK_DIR, `${uuidv4()}.docx`);
      fs.copyFileSync(uploadedFilePath, workingDocxPath);
    }

    // Step 2: Parse DOCX
    updateJob(job, {
      status: "parsing-docx",
      stepMessage: "正在解析 DOCX 文件...",
      progress: 15,
    });

    const parsed = await parseDocx(workingDocxPath);
    job.totalSegments = parsed.segments.length;

    if (job.cancelled) {
      console.log("Job cancelled after parsing");
      return;
    }

    // Step 3: Translate
    updateJob(job, {
      status: "translating",
      stepMessage: "翻譯中...",
      progress: 20,
    });

    await translateSegments(job, parsed.segments, {
      sourceLang: "English",
      targetLang: "Traditional Chinese",
    });

    if (job.cancelled) {
      console.log("Job cancelled after translation");
      return;
    }

    // Step 4: QA and retranslate
    await qaAndRetranslate(job, parsed.segments, {
      sourceLang: "English",
      targetLang: "Traditional Chinese",
    });

    if (job.cancelled) {
      console.log("Job cancelled after QA");
      return;
    }

    // Step 5: Pack output DOCX
    updateJob(job, {
      status: "packing",
      stepMessage: "正在打包翻譯後的文件...",
      progress: 95,
    });

    const outputPath = path.join(OUTPUT_DIR, `${baseName}-translated.docx`);
    await writeDocx(parsed, outputPath);

    // Done
    job.outputPath = outputPath;
    finishJob(job, "done");
    updateJob(job, {
      stepMessage: "完成！",
    });

    console.log(`Job ${job.id} completed successfully`);

    // Cleanup: remove uploaded file and working file
    try {
      fs.unlinkSync(uploadedFilePath);
      fs.unlinkSync(workingDocxPath);
    } catch (e) {
      // Ignore cleanup errors
    }
  } catch (error: any) {
    if (job.cancelled) {
      console.log(`Job ${job.id} was cancelled`);
      return;
    }

    console.error(`Job ${job.id} failed:`, error);
    updateJob(job, {
      status: "error",
      errorMessage: error.message || "Unknown error",
      stepMessage: "處理失敗",
      finishedAt: Date.now(),
    });
  }
}

/**
 * POST /api/upload
 * Upload a PDF or DOCX file and start processing
 */
app.post("/api/upload", upload.single("file"), (req: Request, res: Response) => {
  if (!req.file) {
    res.status(400).json({ error: "No file uploaded" });
    return;
  }

  const jobId = uuidv4();
  const job = createJob(jobId, req.file.originalname);

  // Respond immediately with job ID
  res.json({ jobId });

  // Process in background
  (async () => {
    await processJob(job, req.file!.path);
  })();
});

/**
 * GET /api/status/:jobId
 * Get the current status of a job
 */
app.get("/api/status/:jobId", (req: Request, res: Response) => {
  const job = getJob(req.params.jobId);

  if (!job) {
    res.status(404).json({ error: "Job not found" });
    return;
  }

  const elapsedSeconds = getElapsedSeconds(job);
  const downloadable = job.status === "done" && !!job.outputPath;

  res.json({
    id: job.id,
    fileName: job.fileName,
    status: job.status,
    progress: job.progress,
    stepMessage: job.stepMessage,
    errorMessage: job.errorMessage,
    elapsedSeconds,
    usage: job.usage,
    costUSD: job.costUSD,
    downloadable,
  });
});

/**
 * POST /api/stop/:jobId
 * Cancel a running job
 */
app.post("/api/stop/:jobId", (req: Request, res: Response) => {
  const job = getJob(req.params.jobId);

  if (!job) {
    res.status(404).json({ error: "Job not found" });
    return;
  }

  cancelJob(req.params.jobId);
  res.json({ ok: true });
});

/**
 * GET /api/download/:jobId
 * Download the translated DOCX file
 */
app.get("/api/download/:jobId", (req: Request, res: Response) => {
  const job = getJob(req.params.jobId);

  if (!job || !job.outputPath || job.status !== "done") {
    res.status(404).json({ error: "File not available" });
    return;
  }

  if (!fs.existsSync(job.outputPath)) {
    res.status(404).json({ error: "Output file not found" });
    return;
  }

  const baseName = path.basename(job.fileName, path.extname(job.fileName));
  const downloadName = `${baseName}-translated.docx`;

  res.download(job.outputPath, downloadName);
});

// Error handling middleware
app.use((err: any, req: Request, res: Response, next: any) => {
  console.error("Error:", err);
  res.status(500).json({ error: err.message || "Internal server error" });
});

// Start server
app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
  console.log("Ready to accept file uploads for translation");
});
