import JSZip from "jszip";
import { Router } from "express";
import multer from "multer";
import path from "path";
import { parseUploadedFiles, TimetableRow, parseTextTable } from "../parsers/index";
import { mergeAndResolve, assignmentsToDocx } from "../solver/index";
import fs from "fs";

const router = Router();
const upload = multer({ dest: path.join(process.cwd(), "uploads/") });

// POST /api/upload
// Accepts multiple files (pdf/docx/xlsx/csv)
router.post("/", upload.array("files", 10), async (req, res) => {
  const files = req.files as Express.Multer.File[] | undefined;
  if (!files || files.length === 0) return res.status(400).json({ error: "No files uploaded" });

  const zip = new JSZip();
  let allParsed: TimetableRow[] = [];
  for (const file of files) {
    let parsed: TimetableRow[] = [];
    try {
      const ext = path.extname(file.originalname).toLowerCase();
      if (ext === ".docx" || ext === ".doc") {
        const buffer = fs.readFileSync(file.path);
        // For .docx, use mammoth; for .doc, treat as text
        if (ext === ".docx") {
          const result = await require("mammoth").extractRawText({ buffer });
          parsed = parseTextTable(result.value, file.originalname);
        } else {
          const content = fs.readFileSync(file.path, "utf8");
          parsed = parseTextTable(content, file.originalname);
        }
      } else {
        parsed = await parseUploadedFiles([file]);
      }
    } catch (err) {
      // skip errors, do not add to zip
    }
    allParsed = allParsed.concat(parsed);
  }

  // Resolve all timetables together
  const resolved = mergeAndResolve(allParsed);

  // Only add resolved .docx files to the zip
  for (const [sourceFile, assignments] of resolved.separatedTimetables.entries()) {
    const docxBuffer = await assignmentsToDocx(assignments, sourceFile);
    zip.file(`resolved_${sourceFile.replace(/\.[^/.]+$/, "")}.docx`, docxBuffer);
  }

  const zipBuffer = await zip.generateAsync({ type: "nodebuffer" });
  res.set({
    "Content-Type": "application/zip",
    "Content-Disposition": "attachment; filename=resolved_timetables.zip"
  });
  // Clean up uploaded files after reading
  files.forEach(file => {
    try { fs.unlinkSync(file.path); } catch {}
  });
  return res.send(zipBuffer);
});

export default router;
