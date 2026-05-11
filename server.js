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

// Debug : Liste les fichiers pour vérifier la présence du template
console.log("Fichiers au démarrage :", fs.readdirSync(__dirname));

app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS, GET");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

app.get("/", (req, res) => {
  res.send("Backend Navette Paie OK");
});

app.post("/generate-navette-paie", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).send("Fichier source manquant.");

    const workbookSource = new ExcelJS.Workbook();
    await workbookSource.xlsx.load(req.file.buffer);
    const sheetSource = workbookSource.worksheets[0];

    const workbookDest = new ExcelJS.Workbook();
    const templatePath = path.join(__dirname, "navette_paie_template.xlsx");

    if (!fs.existsSync(templatePath)) {
      return res.status(500).send(`Erreur : Le fichier template est introuvable.`);
    }

    await workbookDest.xlsx.readFile(templatePath);
    const sheetDest = workbookDest.getWorksheet("Etat navette paie") || workbookDest.worksheets[0];

    const startRowSource = 13;
    const startRowDest = 5;
    let currentDestRow = startRowDest;

    sheetSource.eachRow((row, rowNumber) => {
      if (rowNumber < startRowSource) return;

      // Récupération du matricule brut pour garder le format d'origine
      const matriculeCell = row.getCell(1);
      const matricule = matriculeCell.value;
      const nom = row.getCell(2).value;

      if (!matricule) return;

      const absences = extractAbsences(row);
      const destRow = sheetDest.getRow(currentDestRow);

      // Copie du matricule tel quel
      destRow.getCell(1).value = matricule;
      destRow.getCell(2).value = nom;

      // Fonction pour écrire sans le zéro devant (cast en Number)
      const writeData = (type, colStart) => {
        if (absences[type]) {
          destRow.getCell(colStart).value = Number(absences[type].total);     // NB JR
          destRow.getCell(colStart + 1).value = Number(absences[type].start); // DU
          destRow.getCell(colStart + 2).value = Number(absences[type].end);   // AU
        }
      };

      // Colonnes selon vos spécifications :
      writeData('CA', 3);       // C, D, E
      writeData('MALADIE', 6);  // F, G, H
      writeData('AT', 9);       // I, J, K
      writeData('ABS', 12);     // L, M, N

      destRow.commit();
      currentDestRow++;
    });

    const buffer = await workbookDest.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=navette_paie.xlsx");
    return res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).send("Erreur : " + err.message);
  }
});

function extractAbsences(row) {
  const res = {};
  const map = { 'CA': 'CA', 'MALADIE': 'MALADIE', 'AT': 'AT', 'ABS': 'ABS' };
  
  // Lecture de C(3) à AF(32)
  for (let col = 3; col <= 32; col++) {
    let val = row.getCell(col).value;
    if (!val) continue;
    
    val = String(val).trim().toUpperCase();

    if (map[val]) {
      const day = col - 2;
      if (!res[val]) {
        res[val] = { total: 0, start: day, end: day };
      }
      res[val].total++;
      res[val].end = day; 
    }
  }
  return res;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Serveur sur port ${PORT}`));