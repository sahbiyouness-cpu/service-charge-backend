import express from "express";
import multer from "multer";
import JSZip from "jszip";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";

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
  res.send("Backend Service Charge OK");
});

// =========================================================================
// ROUTE 1 : SERVICE CHARGE (JSZip - Gardée intacte)
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

    const sheetFile = zip.file(firstSheetPath);
    if (!sheetFile) return res.status(400).send("Feuille XML introuvable.");

    let sheetXml = await sheetFile.async("string");
    const sharedStrings = parseSharedStrings(sharedStringsXml || "");
    const rows = parseRows(sheetXml, sharedStrings);
    const result = buildAGUpdates(rows);

    for (const item of result.updates) {
      sheetXml = patchCellValueSafe(sheetXml, item.cellRef, item.value);
    }

    zip.file(firstSheetPath, sheetXml);
    const outputBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "STORE" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${buildOutputName(req.file.originalname)}"`);
    res.setHeader("X-Results", encodeURIComponent(JSON.stringify(result.results.slice(0, 300))));
    return res.send(outputBuffer);
  } catch (err) {
    console.error(err);
    return res.status(500).send("Erreur XLSX: " + err.message);
  }
});

// =========================================================================
// ROUTE 2 : NAVETTE PAIE (Approche hybride sécurisée)
// =========================================================================
app.post("/generate-navette-paie", upload.single("file"), async (req, res) => {
  try {
    const templatePath = path.join(process.cwd(), "templates", "navette_paie_template.xlsx");
    
    if (!fs.existsSync(templatePath)) {
      return res.status(500).send("Template introuvable sur le serveur.");
    }

    // On lit le fichier template comme un buffer binaire
    const templateBuffer = fs.readFileSync(templatePath);
    
    // On charge le workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateBuffer); // Utilisation de .load() au lieu de readFile pour éviter les conflits de flux

    // On force la suppression de TOUTES les feuilles sauf celle qu'on veut (si elles existent encore)
    workbook.eachSheet((sheet, id) => {
      if (sheet.name !== "JANVIER 2026" && workbook.worksheets.length > 1) {
        // workbook.removeWorksheet(id); // Optionnel : à n'utiliser que si tu veux vraiment nettoyer
      }
    });

    const ws = workbook.getWorksheet("JANVIER 2026") || workbook.getWorksheet(1);
    
    // Écriture de test
    ws.getCell("A5").value = "999";
    ws.getCell("B5").value = "MODIFIE LE " + new Date().toLocaleTimeString();
    
    const buffer = await workbook.xlsx.writeBuffer();
    
    console.log(`Génération Navette : Feuille utilisée [${ws.name}], Taille [${buffer.byteLength}]`);

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="navette_paie_propre.xlsx"');

    return res.send(buffer);
  } catch (err) {
    console.error("Erreur Navette:", err);
    return res.status(500).send("Erreur : " + err.message);
  }
});

// =========================================================================
// FONCTIONS UTILITAIRES (GARDER TEL QUEL)
// =========================================================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("Server on port", PORT));

function buildOutputName(name) { return String(name || "file.xlsx").replace(/\.xlsx$/i, "") + "_traite.xlsx"; }
function resolveFirstWorksheetPath(workbookXml, relsXml) {
  const rid = workbookXml.match(/<sheet[^>]*r:id="([^"]+)"[^>]*\/>/)?.[1];
  if (!rid) return null;
  const relMatch = relsXml.match(new RegExp(`<Relationship[^>]*Id="${rid}"[^>]*Target="([^"]+)"[^>]*/>`));
  if (!relMatch) return null;
  let t = relMatch[1].replace(/^\/+/, "");
  return t.startsWith("xl/") ? t : "xl/" + t;
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
      let val = valMatch ? valMatch[1] : "";
      if (attrs.includes('t="s"')) val = sharedStrings[parseInt(val, 10)] ?? "";
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
function getCellByColumn(row, col) { return row.cells.find(c => colToIndex(c.ref?.match(/^[A-Z]+/)?.[0]) === col); }
function isRowEmptyAtoAF(row) { for (let c = 1; c <= 32; c++) if (!isEmptyValue(getCellByColumn(row, c)?.value)) return false; return true; }
function hasAnyDataAtoAF(row) { for (let c = 1; c <= 32; c++) if (!isEmptyValue(getCellByColumn(row, c)?.value)) return true; return false; }
function matchesWorkedValue(v) { const s = String(v || "").trim().toLowerCase(); return s === "1" || s === "mission"; }
function isEmptyValue(v) { return !v || String(v).trim() === ""; }
function colToIndex(c) { 
  if(!c) return 0;
  let n = 0; 
  for (let i = 0; i < c.length; i++) n = n * 26 + (c.charCodeAt(i) - 64); 
  return n; 
}
function decodeXml(s) { return String(s).replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'"); }