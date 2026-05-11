import express from "express";
import multer from "multer";
import JSZip from "jszip";
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

app.get("/", (req, res) => {
  res.send("Backend RimH - Service Charge & Navette OK");
});

// =========================================================================
// ROUTE 1 : PROCESS XLSX (SERVICE CHARGE) - TON CODE D'ORIGINE INTACT
// =========================================================================
app.post("/process-xlsx", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).send("Fichier manquant.");
    const zip = await JSZip.loadAsync(req.file.buffer);
    const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    const sharedStringsXml = await zip.file("xl/sharedStrings.xml")?.async("string");

    if (!workbookXml || !relsXml) return res.status(400).send("Fichier XLSX invalide.");

    const firstSheetPath = resolveFirstWorksheetPath(workbookXml, relsXml);
    if (!firstSheetPath) return res.status(400).send("Première feuille introuvable.");

    let sheetXml = await zip.file(firstSheetPath).async("string");
    const sharedStrings = parseSharedStrings(sharedStringsXml || "");
    const rows = parseRows(sheetXml, sharedStrings);
    const result = buildAGUpdates(rows);

    for (const item of result.updates) {
      sheetXml = patchCellValueSafe(sheetXml, item.cellRef, item.value);
    }

    zip.file(firstSheetPath, sheetXml);
    const outputBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "STORE" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="service_charge_traite.xlsx"`);
    res.setHeader("X-Results", encodeURIComponent(JSON.stringify(result.results.slice(0, 300))));
    return res.send(outputBuffer);
  } catch (err) {
    console.error(err);
    return res.status(500).send("Erreur Service Charge: " + err.message);
  }
});

// =========================================================================
// ROUTE 2 : GENERATE NAVETTE PAIE (LOGIQUE RÉELLE DE TRANSFERT)
// =========================================================================
app.post("/generate-navette-paie", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).send("Fichier source manquant.");

    // 1. Lire le fichier source (ton tableau avec MAT et NOM)
    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.load(req.file.buffer);
    const sourceWs = sourceWorkbook.getWorksheet(1);
    
    const employeeData = [];

    // On parcourt le fichier source pour extraire les données
    sourceWs.eachRow((row, rowNumber) => {
      const mat = row.getCell(1).value; // Colonne A: Matricule
      const nom = row.getCell(2).value; // Colonne B: Nom
      
      if (mat && !isNaN(mat)) { // On s'assure que c'est une ligne de données (matricule numérique)
        let totalJours = 0;
        // On compte les présences (colonnes 3 à 32 comme pour ton service charge)
        for (let col = 3; col <= 32; col++) {
          const val = row.getCell(col).value;
          if (matchesWorkedValue(val)) {
            totalJours++;
          }
        }
        employeeData.push({ mat, nom, total: totalJours });
      }
    });

    // 2. Charger le template Navette
    const templatePath = path.join(process.cwd(), "templates", "navette_paie_template.xlsx");
    if (!fs.existsSync(templatePath)) throw new Error("Template non trouvé sur le serveur.");

    const templateWorkbook = new ExcelJS.Workbook();
    await templateWorkbook.xlsx.load(fs.readFileSync(templatePath));
    const targetWs = templateWorkbook.getWorksheet("Etat navette paie") || templateWorkbook.getWorksheet(1);

    // 3. Remplissage dynamique à partir de la ligne 5
    employeeData.forEach((emp, index) => {
      const targetRow = 5 + index;
      targetWs.getCell(`A${targetRow}`).value = emp.mat;
      targetWs.getCell(`B${targetRow}`).value = emp.nom;
      
      // On met le total calculé dans la colonne "Congé Payé / NB JR" (Colonne C)
      // Note : Tu pourras adapter ici si ce total doit aller ailleurs
      targetWs.getCell(`C${targetRow}`).value = emp.total;
    });

    // 4. Exportation
    const buffer = await templateWorkbook.xlsx.writeBuffer();
    
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="navette_paie_complete.xlsx"');
    return res.send(buffer);

  } catch (err) {
    console.error("Erreur Navette:", err);
    return res.status(500).send("Erreur lors de la génération: " + err.message);
  }
});

// =========================================================================
// UTILITAIRES (GARDÉS POUR LE SERVICE CHARGE ET LA NAVETTE)
// =========================================================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("Serveur actif sur port", PORT));

function resolveFirstWorksheetPath(workbookXml, relsXml) {
  const rid = workbookXml.match(/<sheet[^>]*r:id="([^"]+)"[^>]*\/>/)?.[1];
  const relMatch = relsXml.match(new RegExp(`<Relationship[^>]*Id="${rid}"[^>]*Target="([^"]+)"[^>]*/>`));
  if (!relMatch) return null;
  let target = relMatch[1].replace(/^\/+/, "");
  if (!target.startsWith("xl/")) target = "xl/" + target.replace(/^\.?\//, "");
  return target;
}

function parseSharedStrings(xml) {
  if (!xml) return [];
  return (xml.match(/<si[\s\S]*?<\/si>/g) || []).map(si => {
    return [...si.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(m => decodeXml(m[1])).join("");
  });
}

function parseRows(sheetXml, sharedStrings) {
  const rows = [];
  const rowRegex = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
  let m;
  while ((m = rowRegex.exec(sheetXml)) !== null) {
    const rowNum = m[1].match(/\br="(\d+)"/)?.[1];
    const cells = [];
    const cRegex = /<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g;
    let cm;
    while ((cm = cRegex.exec(m[2])) !== null) {
      const attrs = cm[1] || cm[3] || "";
      const valMatch = (cm[2] || "").match(/<v>([\s\S]*?)<\/v>/);
      const typeMatch = attrs.match(/\bt="s"/);
      let val = valMatch ? valMatch[1] : "";
      if (typeMatch) val = sharedStrings[parseInt(val, 10)] ?? "";
      cells.push({ ref: attrs.match(/\br="([A-Z]+[0-9]+)"/)?.[1], value: val });
    }
    rows.push({ rowNumber: rowNum ? parseInt(rowNum, 10) : null, cells });
  }
  return rows;
}

function buildAGUpdates(rows) {
  const updates = [], results = [];
  let block = 1, i = 0;
  while (i < rows.length) {
    let marker = -1;
    for (; i < rows.length; i++) {
      if (!isEmptyValue(getCellByColumn(rows[i], 33)?.value)) { marker = i; break; }
    }
    if (marker === -1) break;
    i = marker + 1;
    let count = 0;
    while (i < rows.length) {
      if (isRowEmptyAtoAF(rows[i])) break;
      if (hasAnyDataAtoAF(rows[i])) {
        let total = 0;
        for (let c = 3; c <= 32; c++) if (matchesWorkedValue(getCellByColumn(rows[i], c)?.value)) total++;
        updates.push({ cellRef: `AG${rows[i].rowNumber}`, value: total });
        results.push({ section: `BLOC ${block}`, rowNumber: rows[i].rowNumber, name: getCellByColumn(rows[i], 2)?.value, total });
        count++;
      }
      i++;
    }
    if (count > 0) block++;
    i++;
  }
  return { updates, results };
}

function patchCellValueSafe(xml, ref, val) {
  const rowNum = ref.match(/\d+$/)?.[0];
  const rowReg = new RegExp(`(<row\\b[^>]*\\br="${rowNum}"[^>]*>)([\\s\\S]*?)(<\\/row>)`);
  const m = xml.match(rowReg);
  if (!m) return xml;
  const cReg = new RegExp(`<c([^>]*\\br="${ref}"[^>]*)>([\\s\\S]*?)<\\/c>|<c([^>]*\\br="${ref}"[^>]*)\\/>`);
  let inner = m[2];
  const newVal = `<c r="${ref}"><v>${val}</v></c>`;
  if (cReg.test(inner)) inner = inner.replace(cReg, newVal);
  else inner += newVal;
  return xml.replace(rowReg, `${m[1]}${inner}${m[3]}`);
}

function getCellByColumn(row, col) { 
  return row.cells.find(c => {
    const letters = c.ref?.match(/^[A-Z]+/)?.[0];
    return colToIndex(letters) === col;
  });
}
function isRowEmptyAtoAF(row) { for (let c = 1; c <= 32; c++) if (!isEmptyValue(getCellByColumn(row, c)?.value)) return false; return true; }
function hasAnyDataAtoAF(row) { for (let c = 1; c <= 32; c++) if (!isEmptyValue(getCellByColumn(row, c)?.value)) return true; return false; }
function matchesWorkedValue(v) { 
  const s = String(v || "").trim().toLowerCase(); 
  return s === "1" || s === "mission"; 
}
function isEmptyValue(v) { return !v || String(v).trim() === ""; }
function colToIndex(c) { 
  if(!c) return 0;
  let n = 0; 
  for (let i = 0; i < c.length; i++) n = n * 26 + (c.charCodeAt(i) - 64); 
  return n; 
}
function decodeXml(s) { return String(s).replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'"); }