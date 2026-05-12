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

    // 1. Récupération des dates réelles (Ligne 11)
    const datesMap = {};
    const row11 = sheetSource.getRow(11);
    for (let col = 3; col <= 32; col++) {
      datesMap[col] = row11.getCell(col).value;
    }

    const startRowSource = 13;
    let currentDestRow = 5;

    sheetSource.eachRow((row, rowNumber) => {
      if (rowNumber < startRowSource) return;

      const matriculeRaw = row.getCell(1).value;
      const nom = row.getCell(2).value;
      if (!matriculeRaw) return;

      const sequences = extractSequences(row, datesMap);

      // Si aucune absence, on liste quand même l'employé
      if (sequences.length === 0) {
        const destRow = sheetDest.getRow(currentDestRow);
        
        // Force le matricule en format nombre pour éviter l'affichage "date"
        const cellMat = destRow.getCell(1);
        cellMat.value = Number(matriculeRaw);
        cellMat.numFmt = '0'; 
        
        destRow.getCell(2).value = nom;
        currentDestRow++;
      } else {
        const firstRowIndex = currentDestRow;

        sequences.forEach((seq) => {
          const destRow = sheetDest.getRow(currentDestRow);
          
          // Matricule forcé en nombre
          const cellMat = destRow.getCell(1);
          cellMat.value = Number(matriculeRaw);
          cellMat.numFmt = '0';

          destRow.getCell(2).value = nom;

          const colMap = { 'CA': 3, 'MALADIE': 6, 'AT': 9, 'ABS': 12 };
          const startCol = colMap[seq.type];

          if (startCol) {
            // NOMBRE DE JOURS : Forcé en nombre standard (ex: 2 au lieu de 02/01/1900)
            const cellNb = destRow.getCell(startCol);
            cellNb.value = Number(seq.count);
            cellNb.numFmt = '0'; 

            // DATES DU / AU : Format Date
            const cellDu = destRow.getCell(startCol + 1);
            const cellAu = destRow.getCell(startCol + 2);
            cellDu.value = seq.start;
            cellAu.value = seq.end;
            cellDu.numFmt = 'dd/mm/yyyy';
            cellAu.numFmt = 'dd/mm/yyyy';
          }
          currentDestRow++;
        });

        // Fusion Nom/Matricule si plusieurs lignes
        if (sequences.length > 1) {
          sheetDest.mergeCells(firstRowIndex, 1, currentDestRow - 1, 1);
          sheetDest.mergeCells(firstRowIndex, 2, currentDestRow - 1, 2);
          const align = { vertical: 'middle', horizontal: 'left' };
          sheetDest.getRow(firstRowIndex).getCell(1).alignment = align;
          sheetDest.getRow(firstRowIndex).getCell(2).alignment = align;
        }
      }
    });

    const buffer = await workbookDest.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=navette_paie_finale.xlsx");
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
    let cellVal = row.getCell(col).value;
    let val = cellVal ? String(cellVal).trim().toUpperCase() : null;

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
app.listen(PORT, () => console.log(`Serveur prêt sur port ${PORT}`));