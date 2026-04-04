const formidable = require("formidable");
const XLSX = require("xlsx");
const fs = require("fs");

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).send("Method not allowed");
  }

  const form = new formidable.IncomingForm();

  form.parse(req, async (err, fields, files) => {
    try {
      if (err) return res.status(500).send(err.message);

      const workbookFile = files.workbook;
      if (!workbookFile) {
        return res.status(400).send("Missing workbook");
      }

      const rawData = fields.data;
      if (!rawData) {
        return res.status(400).send("Missing data");
      }

      const rows = JSON.parse(rawData);

      const workbookBuffer = fs.readFileSync(workbookFile.filepath);
      const workbook = XLSX.read(workbookBuffer, { type: "buffer" });

      const sheet = workbook.Sheets["1_Data Entry"];

      rows.forEach((row, i) => {
        const r = 6 + i;

        if (row.ticker) sheet["E" + r] = { v: row.ticker };
        if (row.buyDate) sheet["G" + r] = { v: row.buyDate };
        if (row.sharesBought) sheet["H" + r] = { v: row.sharesBought };
        if (row.costPerShare) sheet["I" + r] = { v: row.costPerShare };
        if (row.sellDate) sheet["J" + r] = { v: row.sellDate };
        if (row.sharesSold) sheet["K" + r] = { v: row.sharesSold };
        if (row.salePricePerShare) sheet["L" + r] = { v: row.salePricePerShare };
        if (row.note) sheet["N" + r] = { v: row.note };
      });

      const output = XLSX.write(workbook, {
        type: "buffer",
        bookType: "xlsx",
      });

      res.setHeader(
        "Content-Disposition",
        'attachment; filename="merged.xlsx"'
      );
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      res.status(200).send(output);
    } catch (e) {
      res.status(500).send(e.message);
    }
  });
};
