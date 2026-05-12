import express from "express";
import multer from "multer";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS, GET");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

app.post("/generate-navette-paie", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).send("Fichier manquant.");

    const workbookSource = new ExcelJS.Workbook();
    await workbookSource.xlsx.load(req.file.buffer);
    const sheetSource = workbookSource.worksheets[0];

    const workbookDest = new ExcelJS.Workbook();
    const templatePath = path.join(__dirname, "navette_paie_template.xlsx");
    await workbookDest.xlsx.readFile(templatePath);
    const sheetDest = workbookDest.getWorksheet("Etat navette paie") || workbookDest.worksheets[0];

    const datesMap = {};
    const row11 = sheetSource.getRow(11);
    for (let col = 3; col <= 32; col++) {
      datesMap[col] = row11.getCell(col).value;
    }

    const startRowSource = 13;
    let currentDestRow = 5;

    // Variables pour les sommes
    let sumCA = 0;
    let sumMaladie = 0;
    let sumAT = 0;
    let sumAbs = 0;

    sheetSource.eachRow((row, rowNumber) => {
      if (rowNumber < startRowSource) return;

      const matriculeRaw = row.getCell(1).value;
      const nom = row.getCell(2).value;
      if (!matriculeRaw) return;

      const sequences = extractSequences(row, datesMap);

      const fillRow = (m, n, seq = null) => {
        const destRow = sheetDest.getRow(currentDestRow);
        
        const cellMat = destRow.getCell(1);
        cellMat.value = m.toString();
        cellMat.numFmt = '@';

        destRow.getCell(2).value = n;

        if (seq) {
          const colMap = { 'CA': 3, 'MALADIE': 6, 'AT': 9, 'ABS': 12 };
          const startCol = colMap[seq.type];

          if (startCol) {
            const count = Number(seq.count);
            
            // Calcul des sommes
            if (seq.type === 'CA') sumCA += count;
            if (seq.type === 'MALADIE') sumMaladie += count;
            if (seq.type === 'AT') sumAT += count;
            if (seq.type === 'ABS') sumAbs += count;

            const cellNb = destRow.getCell(startCol);
            cellNb.value = count;
            cellNb.numFmt = '0';

            const cellDu = destRow.getCell(startCol + 1);
            const cellAu = destRow.getCell(startCol + 2);
            cellDu.value = seq.start instanceof Date ? seq.start : new Date(seq.start);
            cellAu.value = seq.end instanceof Date ? seq.end : new Date(seq.end);
            cellDu.numFmt = 'dd/mm/yyyy';
            cellAu.numFmt = 'dd/mm/yyyy';
          }
        }
        currentDestRow++;
      };

      if (sequences.length === 0) {
        fillRow(matriculeRaw, nom);
      } else {
        const firstRowIndex = currentDestRow;
        sequences.forEach(s => fillRow(matriculeRaw, nom, s));

        if (sequences.length > 1) {
          sheetDest.mergeCells(firstRowIndex, 1, currentDestRow - 1, 1);
          sheetDest.mergeCells(firstRowIndex, 2, currentDestRow - 1, 2);
          sheetDest.getRow(firstRowIndex).getCell(1).alignment = { vertical: 'middle', horizontal: 'left' };
          sheetDest.getRow(firstRowIndex).getCell(2).alignment = { vertical: 'middle', horizontal: 'left' };
        }
      }
    });

    // --- INSERTION DES SOMMES À LA LIGNE 62 ---
    const row62 = sheetDest.getRow(62);
    
    // Colonne 3 (CA), 6 (Maladie), 9 (AT), 12 (Absence)
    row62.getCell(3).value = sumCA;
    row62.getCell(6).value = sumMaladie;
    row62.getCell(9).value = sumAT;
    row62.getCell(12).value = sumAbs;

    // Application du format nombre sans gras
    [3, 6, 9, 12].forEach(col => {
      row62.getCell(col).numFmt = '0';
      row62.getCell(col).font = { bold: false }; 
    });

    const buffer = await workbookDest.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=navette_paie.xlsx");
    return res.send(buffer);

  } catch (err) {
    res.status(500).send("Erreur: " + err.message);
  }
});

function extractSequences(row, datesMap) {
  const sequences = [];
  const targets = ['CA', 'MALADIE', 'AT', 'ABS'];
  let current = null;

  for (let col = 3; col <= 32; col++) {
    let val = row.getCell(col).value;
    val = val ? String(val).trim().toUpperCase() : null;

    if (targets.includes(val)) {
      const dateVal = datesMap[col];
      if (current && current.type === val) {
        current.count++;
        current.end = dateVal;
      } else {
        if (current) sequences.push(current);
        current = { type: val, start: dateVal, end: dateVal, count: 1 };
      }
    } else {
      if (current) {
        sequences.push(current);
        current = null;
      }
    }
  }
  if (current) sequences.push(current);
  return sequences;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Listening on ${PORT}`));