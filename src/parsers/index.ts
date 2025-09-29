import fs from "fs";
import path from "path";
import ExcelJS from "exceljs";
import { parse as csvParse } from "csv-parse";
import mammoth from "mammoth";
import pdfParse from "pdf-parse";

export type TimetableRow = {
  day: string; // e.g., Mon
  period: string; // e.g., P1
  subject?: string;
  teacher?: string;
  group?: string; // can be comma-separated groups
  room?: string;
  sourceFile?: string;
};

export async function parseUploadedFiles(files: Express.Multer.File[]): Promise<TimetableRow[]> {
  const rows: TimetableRow[] = [];
  for (const f of files) {
    const ext = path.extname(f.originalname).toLowerCase();
    try {
      if (ext === ".csv") {
        const content = fs.readFileSync(f.path, "utf8");
        const recs: any[] = [];
        await new Promise<void>((resolve, reject) => {
          csvParse(content, { columns: true, skip_empty_lines: true })
            .on('data', (r) => recs.push(r))
            .on('end', () => resolve())
            .on('error', (err) => reject(err));
        });
        for (const r of recs) rows.push(normalizeRow(r, f.originalname));
      } else if (ext === ".xlsx" || ext === ".xls") {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(f.path);
        const sheet = workbook.worksheets[0];
        const header: string[] = [];
        sheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) {
            row.eachCell((cell) => header.push(String(cell.text).trim()));
          } else {
            const r: any = {};
            row.eachCell((cell, colNumber) => {
              r[header[colNumber - 1] || `col${colNumber}`] = String(cell.text).trim();
            });
            rows.push(normalizeRow(r, f.originalname));
          }
        });
      } else if (ext === ".docx") {
        const buffer = fs.readFileSync(f.path);
        const result = await mammoth.extractRawText({ buffer });
        rows.push(...parseTextTable(result.value, f.originalname));
      } else if (ext === ".pdf") {
        const buffer = fs.readFileSync(f.path);
        const data = await pdfParse(buffer);
        rows.push(...parseTextTable(data.text, f.originalname));
      } else {
        // try to treat as text
        const content = fs.readFileSync(f.path, "utf8");
        rows.push(...parseTextTable(content, f.originalname));
      }
    } catch (err) {
      console.warn("Failed to parse", f.originalname, err);
    } /* finally {
      // optionally unlink uploaded file to save space
      try { fs.unlinkSync(f.path); } catch {};
    } */
  }
  return rows;
}

function normalizeRow(raw: any, source: string): TimetableRow {
  // Attempt to map common column headers to our schema.
  const mapKey = (k: string) => k.toLowerCase().trim();
  const out: TimetableRow = {
    day: raw.day || raw.Day || raw['DAY'] || raw["Day"] || raw["day"] || raw["weekday"] || raw["Weekday"] || raw["DAY_OF_WEEK"] || "",
    period: raw.period || raw.Period || raw['PERIOD'] || raw["Time"] || raw["Slot"] || raw["Period"] || "",
    subject: raw.subject || raw.Subject || raw['SUBJECT'] || raw["Class"] || raw["Course"] || raw["Lesson"] || "",
    teacher: raw.teacher || raw.Teacher || raw["Staff"] || raw["Lecturer"] || "",
    group: raw.group || raw.Group || raw["Groups"] || raw["ClassGroup"] || "",
    room: raw.room || raw.Room || raw["Venue"] || "",
    sourceFile: source
  } as TimetableRow;
  // trim
  for (const k of Object.keys(out) as (keyof TimetableRow)[]) {
    if (typeof out[k] === "string") out[k] = (out[k] as string).trim();
  }
  return out;
}

function parseTextTable(text: string, source: string): TimetableRow[] {
  // Multi-line parser for vertical timetable format
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  // Find header positions
  const headerIdx = lines.findIndex(l => l.toLowerCase() === "day");
  if (headerIdx === -1) return [];
  const fields = ["day", "period", "subject", "teacher", "group", "room"];
  const rows: TimetableRow[] = [];
  // Start after header
  for (let i = headerIdx + fields.length; i + fields.length - 1 < lines.length; i += fields.length) {
    const r: any = {};
    for (let j = 0; j < fields.length; j++) {
      r[fields[j]] = lines[i + j] || "";
    }
    r.sourceFile = source;
    rows.push(normalizeRow(r, source));
  }
  return rows;
}

export { parseTextTable };
