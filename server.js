import express from "express";
import multer from "multer";
import ExcelJS from "exceljs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Middleware CORS pour permettre les requêtes depuis votre interface (Cloudflare)
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

// --- ROUTE : GENERATE NAVETTE PAIE ---
app.post("/generate-navette-paie", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send("Fichier source manquant.");
    }

    // 1. Charger le fichier source (Pointage CDI Diwan) envoyé par l'utilisateur
    const workbookSource = new ExcelJS.Workbook();
    await workbookSource.xlsx.load(req.file.buffer);
    const sheetSource = workbookSource.worksheets[0]; // Prend la première feuille

    // 2. Charger le template stocké sur le serveur (nommé navette_template.xlsx)
    const workbookDest = new ExcelJS.Workbook();
    const templatePath = path.join(__dirname, "navette_template.xlsx");
    await workbookDest.xlsx.readFile(templatePath);
    const sheetDest = workbookDest.getWorksheet("Etat navette paie");

    if (!sheetDest) {
      return res.status(500).send("L'onglet 'Etat navette paie' est introuvable dans le template.");
    }

    const startRowSource = 13;
    const startRowDest = 5;
    let currentDestRow = startRowDest;

    const summaryResults = [];

    // 3. Parcourir les lignes du fichier source
    sheetSource.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber < startRowSource) return;

      const matricule = row.getCell(1).value; // Colonne A
      const nom = row.getCell(2).value;       // Colonne B

      // On s'arrête si le matricule est vide
      if (!matricule || String(matricule).trim() === "") return;

      // Analyse des absences sur la ligne (Colonnes C à AF = 3 à 32)
      const absences = extractAbsenceSequences(row);

      // 4. Écrire les données dans le template
      const destRow = sheetDest.getRow(currentDestRow);
      destRow.getCell(1).value = matricule;
      destRow.getCell(2).value = nom;

      // Mapping des colonnes spécifié par vos soins :
      
      // CA (Congé Annuel) -> C(Total), D(Du), E(Au)
      if (absences.CA) {
        destRow.getCell(3).value = absences.CA.total;
        destRow.getCell(4).value = absences.CA.start;
        destRow.getCell(5).value = absences.CA.end;
      }
      // MALADIE -> F(Total), G(Du), H(Au)
      if (absences.MALADIE) {
        destRow.getCell(6).value = absences.MALADIE.total;
        destRow.getCell(7).value = absences.MALADIE.start;
        destRow.getCell(8).value = absences.MALADIE.end;
      }
      // AT (Accident Travail) -> I(Total), J(Du), K(Au)
      if (absences.AT) {
        destRow.getCell(9).value = absences.AT.total;
        destRow.getCell(10).value = absences.AT.start;
        destRow.getCell(11).value = absences.AT.end;
      }
      // ABS (Absences diverses) -> L(Total), M(Du), N(Au)
      if (absences.ABS) {
        destRow.getCell(12).value = absences.ABS.total;
        destRow.getCell(13).value = absences.ABS.start;
        destRow.getCell(14).value = absences.ABS.end;
      }

      // Pour l'affichage Debug/Résumé dans l'interface HTML
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

    // 5. Générer le fichier de sortie
    const buffer = await workbookDest.xlsx.writeBuffer();

    // Headers pour le téléchargement et l'envoi des résultats de debug
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=navette_paie_generee.xlsx");
    res.setHeader("Access-Control-Expose-Headers", "X-Results");
    res.setHeader("X-Results", encodeURIComponent(JSON.stringify(summaryResults)));

    return res.send(buffer);

  } catch (err) {
    console.error("Erreur serveur:", err);
    res.status(500).send("Erreur lors du traitement : " + err.message);
  }
});

/**
 * Fonction pour extraire les périodes et totaux d'absences.
 * Elle scanne de la colonne 3 (C) à la colonne 32 (AF).
 */
function extractAbsenceSequences(row) {
  const result = {};
  const codes = ['CA', 'MALADIE', 'AT', 'ABS'];

  for (let col = 3; col <= 32; col++) {
    let cellValue = row.getCell(col).value;
    if (!cellValue) continue;

    const val = String(cellValue).trim().toUpperCase();
    
    if (codes.includes(val)) {
      const dayNum = col - 2; // Exemple: Colonne C (3) = Jour 1

      if (!result[val]) {
        result[val] = { total: 0, start: dayNum, end: dayNum };
      }
      result[val].total += 1;
      
      // On met à jour le début et la fin pour avoir la plage "Du ... Au ..."
      if (dayNum < result[val].start) result[val].start = dayNum;
      if (dayNum > result[val].end) result[val].end = dayNum;
    }
  }
  return result;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Serveur démarré sur le port ${PORT}`);
});