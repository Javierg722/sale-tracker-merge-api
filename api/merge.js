import formidable from "formidable";
import * as XLSX from "xlsx";

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

export default async function handler(req, res) {
  const form = formidable({ multiples: false });

  form.parse(req, async (err, fields, files) => {
    if (err) return res.status(500).send("Upload error");

    const workbookFile = files.workbook.filepath;
    const jsonData = JSON.parse(fields.data);

    const workbook = XLSX.readFile(workbookFile);
    const sheet = workbook.Sheets["1_Data Entry"];

    jsonData.forEach((row, i) => {
      const excelRow = START_ROW + i;

      Object.entries(INPUT_COLUMNS).forEach(([key, col]) => {
        const value = row[key];

        if (value !== undefined && value !== null) {
          const cellAddress = col + excelRow;
          sheet[cellAddress] = { v: value };
        }
      });
    });

    const buffer = XLSX.write(workbook, {
      type: "buffer",
      bookType: "xlsx",
    });

    res.setHeader(
      "Content-Disposition",
      "attachment; filename=merged.xlsx"
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.send(buffer);
  });
}
