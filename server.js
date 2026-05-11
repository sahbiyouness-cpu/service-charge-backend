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
    if (!req.file) return res.status(400).send("Fichier source manquant.");

    const workbookSource = new ExcelJS.Workbook();
    await workbookSource.xlsx.load(req.file.buffer);
    const sheetSource = workbookSource.worksheets[0];

    const workbookDest = new ExcelJS.Workbook();
    const templatePath = path.join(__dirname, "navette_paie_template.xlsx");
    if (!fs.existsSync(templatePath)) return res.status(500).send("Template introuvable.");
    
    await workbookDest.xlsx.readFile(templatePath);
    const sheetDest = workbookDest.getWorksheet("Etat navette paie") || workbookDest.worksheets[0];

    // 1. Récupérer les dates complètes de la ligne 11 (Colonnes C à AF)
    const datesMap = {};
    const headerRow = sheetSource.getRow(11);
    for (let col = 3; col <= 32; col++) {
      let dateVal = headerRow.getCell(col).value;
      // On stocke la date telle quelle pour garder le format complet
      datesMap[col] = dateVal; 
    }

    const startRowSource = 13;
    let currentDestRow = 5;

    sheetSource.eachRow((row, rowNumber) => {
      if (rowNumber < startRowSource) return;

      const matricule = row.getCell(1).value;
      const nom = row.getCell(2).value;
      if (!matricule) return;

      // 2. Extraire les séquences d'absences (gère les ruptures)
      const allSequences = extractAllSequences(row, datesMap);

      if (allSequences.length > 0) {
        const firstRowIndex = currentDestRow;
        
        allSequences.forEach((seq) => {
          const destRow = sheetDest.getRow(currentDestRow);
          
          // On injecte le matricule et le nom (ils seront fusionnés après)
          destRow.getCell(1).value = matricule;
          destRow.getCell(2).value = nom;

          const colMap = { 'CA': 3, 'MALADIE': 6, 'AT': 9, 'ABS': 12 };
          const startCol = colMap[seq.type];

          if (startCol) {
            // NB JR : Forcer en Number pour enlever les zéros (ex: 0009 -> 9)
            destRow.getCell(startCol).value = Number(seq.count);
            
            // "DU" et "À" : Format Date complet
            destRow.getCell(startCol + 1).value = seq.startDate;
            destRow.getCell(startCol + 2).value = seq.endDate;
            
            // Appliquer le format de date français dans la cellule
            destRow.getCell(startCol + 1).numFmt = 'dd/mm/yyyy';
            destRow.getCell(startCol + 2).numFmt = 'dd/mm/yyyy';
          }
          currentDestRow++;
        });

        // 3. Fusion (Merge) du Matricule et du Nom
        if (allSequences.length > 1) {
          sheetDest.mergeCells(firstRowIndex, 1, currentDestRow - 1, 1);
          sheetDest.mergeCells(firstRowIndex, 2, currentDestRow - 1, 2);
          
          // Alignement pour que ce soit propre
          const align = { vertical: 'middle', horizontal: 'left' };
          sheetDest.getRow(firstRowIndex).getCell(1).alignment = align;
          sheetDest.getRow(firstRowIndex).getCell(2).alignment = align;
        }
      }
    });

    const buffer = await workbookDest.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=navette_paie.xlsx");
    return res.send(buffer);

  } catch (err) {
    res.status(500).send("Erreur : " + err.message);
  }
});

function extractAllSequences(row, datesMap) {
  const sequences = [];
  const codes = ['CA', 'MALADIE', 'AT', 'ABS'];
  let currentSeq = null;

  for (let col = 3; col <= 32; col++) {
    let val = row.getCell(col).value;
    val = val ? String(val).trim().toUpperCase() : null;

    if (codes.includes(val)) {
      const fullDate = datesMap[col];
      
      if (currentSeq && currentSeq.type === val) {
        currentSeq.count++;
        currentSeq.endDate = fullDate;
      } else {
        if (currentSeq) sequences.push(currentSeq);
        currentSeq = { 
          type: val, 
          startDate: fullDate, 
          endDate: fullDate, 
          count: 1 
        };
      }
    } else {
      if (currentSeq) {
        sequences.push(currentSeq);
        currentSeq = null;
      }
    }
  }
  if (currentSeq) sequences.push(currentSeq);
  return sequences;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Serveur prêt sur port ${PORT}`));