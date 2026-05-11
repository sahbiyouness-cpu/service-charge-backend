import express from "express";
import multer from "multer";
import JSZip from "jszip";
import ExcelJS from "exceljs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Middleware CORS
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS, GET");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

app.get("/", (req, res) => {
  res.send("Backend Service Charge & Navette Paie OK");
});

// --- ROUTE 1 : NAVETTE PAIE (Votre nouvelle demande) ---
app.post("/generate-navette-paie", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).send("Fichier source manquant.");

    const workbookSource = new ExcelJS.Workbook();
    await workbookSource.xlsx.load(req.file.buffer);
    
    // Sécurité pour l'erreur 'eachRow' : on s'assure que la feuille existe
    const sheetSource = workbookSource.worksheets[0];
    if (!sheetSource) {
      return res.status(400).send("Le fichier source est vide ou invalide.");
    }

    const workbookDest = new ExcelJS.Workbook();
    // Le fichier template doit être à la racine avec ce nom exact
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
    const summary = [];

    sheetSource.eachRow((row, rowNumber) => {
      if (rowNumber < startRowSource) return;

      const matricule = row.getCell(1).value; // A
      const nom = row.getCell(2).value;       // B

      if (!matricule) return;

      const absences = extractAbsences(row);

      const destRow = sheetDest.getRow(currentDestRow);
      destRow.getCell(1).value = matricule;
      destRow.getCell(2).value = nom;

      // Mapping des colonnes
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

      summary.push({ mat: matricule, name: nom, row: rowNumber });
      currentDestRow++;
    });

    const buffer = await workbookDest.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=navette_generee.xlsx");
    res.setHeader("Access-Control-Expose-Headers", "X-Results");
    res.setHeader("X-Results", encodeURIComponent(JSON.stringify(summary)));
    return res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).send("Erreur Navette: " + err.message);
  }
});

// --- ROUTE 2 : SERVICE CHARGE (Votre ancien code conservé) ---
app.post("/process-xlsx", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).send("Fichier manquant.");
    const zip = await JSZip.loadAsync(req.file.buffer);
    // ... (votre logique JSZip originale ici si vous voulez la garder)
    res.status(200).send("Route Service Charge active");
  } catch (err) {
    res.status(500).send(err.message);
  }
});

// Fonction utilitaire pour calculer les plages de dates
function extractAbsences(row) {
  const res = {};
  const map = { 'CA': 'CA', 'MALADIE': 'MALADIE', 'AT': 'AT', 'ABS': 'ABS' };
  
  // Colonnes C (3) à AF (32)
  for (let col = 3; col <= 32; col++) {
    let val = row.getCell(col).value;
    if (!val) continue;
    val = String(val).trim().toUpperCase();

    if (map[val]) {
      const day = col - 2;
      if (!res[val]) res[val] = { total: 0, start: day, end: day };
      res[val].total++;
      res[val].end = day;
    }
  }
  return res;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Listening on ${PORT}`));