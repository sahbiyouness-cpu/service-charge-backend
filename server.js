import express from "express";
import multer from "multer";
import ExcelJS from "exceljs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Configuration des middlewares CORS pour l'interface
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, X-Results");
  res.setHeader("Access-Control-Expose-Headers", "X-Results");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS, GET, DELETE");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

app.get("/", (req, res) => {
  res.send("Backend Service Charge & Navette Paie OK");
});

// ==========================================
// ROUTE 1 : SERVICE CHARGE (Traitement XLSX)
// ==========================================
app.post("/process-xlsx", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).send("Fichier manquant.");

    const month = req.body.month;
    const year = req.body.year;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);
    const sheet = workbook.worksheets[0];

    const summaryResults = [];

    // Exemple de structure de traitement de la Service Charge basique
    // Parcourt les lignes pour renvoyer les statistiques à l'interface
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber < 2) return; // Ignore l'entête si nécessaire

      const nom = row.getCell(3).value; // Exemple colonne C
      const totalJrs = row.getCell(4).value; // Exemple colonne D
      const bloc = row.getCell(1).value; // Exemple colonne A

      if (nom) {
        summaryResults.push({
          section: bloc || "Général",
          rowNumber: rowNumber,
          name: String(nom),
          total: totalJrs || 0
        });
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=service_charge_traite.xlsx");
    res.setHeader("X-Results", encodeURIComponent(JSON.stringify(summaryResults)));
    return res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).send("Erreur Service Charge: " + err.message);
  }
});

// ==========================================
// ROUTE 2 : GENERATE NAVETTE PAIE
// ==========================================
app.post("/generate-navette-paie", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).send("Fichier manquant.");

    const workbookSource = new ExcelJS.Workbook();
    await workbookSource.xlsx.load(req.file.buffer);
    const sheetSource = workbookSource.worksheets[0];

    const workbookDest = new ExcelJS.Workbook();
    const templatePath = path.join(__dirname, "navette_paie_template.xlsx");
    
    try {
      await workbookDest.xlsx.readFile(templatePath);
    } catch (e) {
      return res.status(500).send("Template 'navette_paie_template.xlsx' introuvable sur le serveur.");
    }

    const sheetDest = workbookDest.getWorksheet("Etat navette paie") || workbookDest.worksheets[0];

    const startRowSource = 13;
    const startRowDest = 5;
    let currentDestRow = startRowDest;
    const summaryResults = [];

    // Extraction des en-têtes de colonnes pour les correspondances de dates (colonnes 3 à 32)
    const datesMap = {};
    const headerRow = sheetSource.getRow(11); 
    for (let col = 3; col <= 32; col++) {
      datesMap[col] = headerRow.getCell(col).value;
    }

    sheetSource.eachRow((row, rowNumber) => {
      if (rowNumber < startRowSource) return;

      const matricule = row.getCell(1).value; 
      const nom = row.getCell(2).value;       

      if (!matricule || String(matricule).trim() === "") return;

      // Logique d'extraction des séquences cumulées de ton code d'origine
      const absences = extractSequences(row, datesMap);

      const destRow = sheetDest.getRow(currentDestRow);
      destRow.getCell(1).value = matricule;
      destRow.getCell(2).value = nom;

      if (absences.CA) {
        destRow.getCell(3).value = absences.CA.total;
        destRow.getCell(4).value = absences.CA.start;
        destRow.getCell(5).value = absences.CA.end;
      }
      if (absences.MALADIE) {
        destRow.getCell(6).value = absences.MALADIE.total;
        destRow.getCell(7).value = absences.MALADIE.start;
        destRow.getCell(8).value = absences.MALADIE.end;
      }
      if (absences.AT) {
        destRow.getCell(9).value = absences.AT.total;
        destRow.getCell(10).value = absences.AT.start;
        destRow.getCell(11).value = absences.AT.end;
      }
      if (absences.ABS) {
        destRow.getCell(12).value = absences.ABS.total;
        destRow.getCell(13).value = absences.ABS.start;
        destRow.getCell(14).value = absences.ABS.end;
      }

      summaryResults.push({
        mat: matricule,
        name: nom,
        rowNumber: rowNumber,
        conge: absences.CA ? absences.CA.total : 0,
        maladie: absences.MALADIE ? absences.MALADIE.total : 0,
        at: absences.AT ? absences.AT.total : 0,
        abs: absences.ABS ? absences.ABS.total : 0
      });

      destRow.commit();
      currentDestRow++;
    });

    // Écritures des totaux globaux (Exemple basé sur ton code stable initial)
    const row62 = sheetDest.getRow(62);
    let sumCA = 0, sumMaladie = 0, sumAT = 0, sumAbs = 0;
    summaryResults.forEach(item => {
      sumCA += item.conge;
      sumMaladie += item.maladie;
      sumAT += item.at;
      sumAbs += item.abs;
    });

    row62.getCell(3).value = sumCA;
    row62.getCell(6).value = sumMaladie;
    row62.getCell(9).value = sumAT;
    row62.getCell(12).value = sumAbs;

    [3, 6, 9, 12].forEach(col => {
      row62.getCell(col).numFmt = '0';
      row62.getCell(col).font = { bold: false }; 
    });

    const buffer = await workbookDest.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=navette_paie.xlsx");
    res.setHeader("X-Results", encodeURIComponent(JSON.stringify(summaryResults)));
    return res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).send("Erreur Navette: " + err.message);
  }
});

// Analyseur de séquences pour la Navette
function extractSequences(row, datesMap) {
  const result = {};
  const targets = ['CA', 'MALADIE', 'AT', 'ABS'];

  for (let col = 3; col <= 32; col++) {
    let val = row.getCell(col).value;
    val = val ? String(val).trim().toUpperCase() : null;

    if (targets.includes(val)) {
      const dateVal = datesMap[col] || (col - 2);
      if (!result[val]) {
        result[val] = { total: 0, start: dateVal, end: dateVal };
      }
      result[val].total++;
      result[val].end = dateVal;
    }
  }
  return result;
}

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`Listening on ${PORT}`));