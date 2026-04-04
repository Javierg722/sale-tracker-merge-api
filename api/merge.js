import formidable from "formidable";
import * as XLSX from "xlsx";
import fs from "fs";

export const config = {
  api: {
    bodyParser: false,
  },
};

const INPUT_COLUMNS = {
  ticker: "E",
  buyDate: "G",
  sharesBought: "H",
  costPerShare: "I",
  sellDate: "J",
  sharesSold: "K",
  salePricePerShare: "L",
  note: "N",
};

const START_ROW = 6;
const SHEET_NAME = "1_Data Entry";

function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
}

function overwriteCell(sheet, address, value) {
  if (value === undefined || value === null || value === "") {
    delete sheet[address];
    return;
  }
  sheet[address] = { v: value };
}

export default async function handler(req, res) {
  setCors(res);

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).send("Method not allowed");
  }

  const form = formidable({
    multiples: false,
    keepExtensions: true,
  });

  form.parse(req, async (err, fields, files) => {
    try {
      if (err) {
        return res.status(500).send(`Form parse error: ${err.message}`);
      }

      const workbookUpload = Array.isArray(files.workbook)
        ? files.workbook[0]
        : files.workbook;

      if (!workbookUpload || !workbookUpload.filepath) {
        return res.status(400).send("Missing workbook file");
      }

      const rawData = Array.isArray(fields.data) ? fields.data[0] : fields.data;
      if (!rawData) {
        return res.status(400).send("Missing data payload");
      }

      const parsed = JSON.parse(rawData);
      const rows = Array.isArray(parsed) ? parsed : parsed.rows;

      if (!Array.isArray(rows)) {
        return res.status(400).send("Invalid data payload");
      }

      const workbookBuffer = fs.readFileSync(workbookUpload.filepath);
      const workbook = XLSX.read(workbookBuffer, { type: "buffer" });

      const sheet = workbook.Sheets[SHEET_NAME];
      if (!sheet) {
        return res.status(400).send(`Sheet not found: ${SHEET_NAME}`);
      }

      for (let i = 0; i < rows.length; i++) {
        const excelRow = START_ROW + i;
        const row = rows[i] || {};

        for (const [field, col] of Object.entries(INPUT_COLUMNS)) {
          overwriteCell(sheet, `${col}${excelRow}`, row[field]);
        }
      }

      const outputBuffer = XLSX.write(workbook, {
        type: "buffer",
        bookType: "xlsx",
      });

      res.setHeader(
        "Content-Disposition",
        'attachment; filename="merged-workbook.xlsx"'
      );
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      return res.status(200).send(outputBuffer);
    } catch (e) {
      return res.status(500).send(`Merge failed: ${e.message}`);
    }
  });
}
