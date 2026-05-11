import express from "express";
import multer from "multer";
import JSZip from "jszip";
import ExcelJS from "exceljs";
import path from "path";

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
  res.send("Backend Service Charge OK");
});

// --- ROUTE 1 : PROCESS XLSX (Traitement JSZip) ---
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
    return res.status(500).send("Erreur traitement XLSX: " + err.message);
  }
});

// --- ROUTE 2 : GENERATE NAVETTE PAIE (ExcelJS Corrigé) ---
app.post("/generate-navette-paie", upload.single("file"), async (req, res) => {
  try {
    console.log("--- NOUVELLE TENTATIVE GENERATION ---");
    const workbook = new ExcelJS.Workbook();
    
    // Chemin absolu vers la racine du projet sur Render
    const templatePath = path.join(process.cwd(), "templates", "navette_paie_template.xlsx");
    console.log("Recherche du fichier ici :", templatePath);

    // Lecture du fichier
    await workbook.xlsx.readFile(templatePath).catch(err => {
      throw new Error("Fichier introuvable sur le disque. Vérifie que le dossier 'templates' est bien poussé sur GitHub.");
    });

    // On liste les feuilles pour le debug dans Render
    const sheetNames = workbook.worksheets.map(w => w.name);
    console.log("Feuilles trouvées dans le template :", sheetNames);

    // D'après tes fichiers, on cherche "JANVIER 2026" ou la première feuille
    // On prend la première feuille visible pour être sûr
    const ws = workbook.getWorksheet(1); 
    
    if (!ws) {
      throw new Error("Aucune feuille de calcul trouvée dans le fichier.");
    }

    console.log("Écriture sur la feuille :", ws.name);

    // FORCE L'ÉCRITURE SUR DES CELLULES EXISTANTES (Basé sur ton CSV)
    // A5 = MAT, B5 = NOM, C5 = Congé Payé NB JR
    ws.getCell("A5").value = "999";
    ws.getCell("B5").value = "TEST NOM RENDER";
    ws.getCell("C5").value = 10;
    
    // On ajoute un test en A1 pour voir si le fichier est modifié
    ws.getCell("A1").value = "MODIFIÉ PAR LE SERVEUR";

    // Génération du buffer
    const buffer = await workbook.xlsx.writeBuffer();
    console.log("Génération réussie. Taille du buffer :", buffer.byteLength);

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="navette_paie_generee.xlsx"');
    
    return res.send(buffer);

  } catch (err) {
    console.error("ERREUR CRITIQUE :", err.message);
    return res.status(500).send("Erreur : " + err.message);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("Serveur en ligne sur le port", PORT);
});

// --- FONCTIONS UTILITAIRES ---

function buildOutputName(name) {
  return String(name || "service_charge.xlsx").replace(/\.xlsx$/i, "") + "_traite.xlsx";
}

function resolveFirstWorksheetPath(workbookXml, relsXml) {
  const sheetMatch = workbookXml.match(/<sheet[^>]*r:id="([^"]+)"[^>]*\/>/);
  if (!sheetMatch) return null;
  const rid = sheetMatch[1];
  const relRegex = new RegExp(`<Relationship[^>]*Id="${escapeRegExp(rid)}"[^>]*Target="([^"]+)"[^>]*/>`);
  const relMatch = relsXml.match(relRegex);
  if (!relMatch) return null;
  let target = relMatch[1].replace(/^\/+/, "");
  if (!target.startsWith("xl/")) target = "xl/" + target.replace(/^\.?\//, "");
  return target;
}

function parseSharedStrings(xml) {
  if (!xml) return [];
  const result = [];
  const siMatches = xml.match(/<si[\s\S]*?<\/si>/g) || [];
  for (const si of siMatches) {
    const tMatches = [...si.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)];
    result.push(tMatches.map(m => decodeXml(m[1])).join(""));
  }
  return result;
}

function parseRows(sheetXml, sharedStrings) {
  const rows = [];
  const rowRegex = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
  let rowMatch;
  while ((rowMatch = rowRegex.exec(sheetXml)) !== null) {
    const rowNumberMatch = rowMatch[1].match(/\br="(\d+)"/);
    const rowInner = rowMatch[2];
    const cells = [];
    const cellRegex = /<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g;
    let cellMatch;
    while ((cellMatch = cellRegex.exec(rowInner)) !== null) {
      const attrs = cellMatch[1] || cellMatch[3] || "";
      const inner = cellMatch[2] || "";
      const refMatch = attrs.match(/\br="([A-Z]+[0-9]+)"/);
      const typeMatch = attrs.match(/\bt="([^"]+)"/);
      let value = "";
      const vMatch = inner.match(/<v>([\s\S]*?)<\/v>/);
      if (vMatch) {
        value = typeMatch && typeMatch[1] === "s" ? sharedStrings[parseInt(vMatch[1], 10)] : vMatch[1];
      }
      cells.push({ ref: refMatch ? refMatch[1] : null, value });
    }
    rows.push({ rowNumber: rowNumberMatch ? parseInt(rowNumberMatch[1], 10) : null, cells });
  }
  return rows;
}

function buildAGUpdates(rows) {
  const FIRST_DATE_COL = 3, LAST_DATE_COL = 32, RESULT_COL = 33;
  const results = [], updates = [];
  let blockCount = 0, i = 0;
  while (i < rows.length) {
    let markerIndex = -1;
    for (; i < rows.length; i++) {
      const agCell = getCellByColumn(rows[i], RESULT_COL);
      if (agCell && !isEmptyValue(agCell.value)) { markerIndex = i; break; }
    }
    if (markerIndex === -1) break;
    i = markerIndex + 1;
    let processed = 0;
    while (i < rows.length) {
      if (isRowEmptyAtoAF(rows[i])) break;
      if (hasAnyDataAtoAF(rows[i])) {
        let total = 0;
        for (let col = FIRST_DATE_COL; col <= LAST_DATE_COL; col++) {
          if (matchesWorkedValue(getCellByColumn(rows[i], col)?.value)) total++;
        }
        updates.push({ cellRef: `AG${rows[i].rowNumber}`, value: total });
        results.push({ section: `BLOC ${blockCount + 1}`, rowNumber: rows[i].rowNumber, name: String(getCellByColumn(rows[i], 2)?.value || ""), total });
        processed++;
      }
      i++;
    }
    if (processed > 0) blockCount++;
    i++;
  }
  return { updates, results };
}

function patchCellValueSafe(sheetXml, cellRef, numericValue) {
  const rowNumber = parseInt(cellRef.match(/\d+$/)?.[0] || "", 10);
  const rowRegex = new RegExp(`(<row\\b[^>]*\\br="${rowNumber}"[^>]*>)([\\s\\S]*?)(<\\/row>)`);
  const rowMatch = sheetXml.match(rowRegex);
  if (!rowMatch) return sheetXml;
  const fullCellRegex = new RegExp(`<c([^>]*\\br="${escapeRegExp(cellRef)}"[^>]*)>([\\s\\S]*?)<\\/c>`);
  const selfCellRegex = new RegExp(`<c([^>]*\\br="${escapeRegExp(cellRef)}"[^>]*)\\/>`);
  let newInner = rowMatch[2];
  if (fullCellRegex.test(newInner)) {
    newInner = newInner.replace(fullCellRegex, (m, a) => `<c${a.replace(/\s+t="[^"]*"/g, "")}><v>${numericValue}</v></c>`);
  } else if (selfCellRegex.test(newInner)) {
    newInner = newInner.replace(selfCellRegex, (m, a) => `<c${a.replace(/\s+t="[^"]*"/g, "")}><v>${numericValue}</v></c>`);
  } else {
    newInner += `<c r="${cellRef}"><v>${numericValue}</v></c>`;
  }
  return sheetXml.replace(rowRegex, `${rowMatch[1]}${newInner}${rowMatch[3]}`);
}

function getCellByColumn(row, col) {
  return row.cells.find(c => {
    const letters = c.ref?.match(/^([A-Z]+)/)?.[1];
    return letters && colToIndex(letters) === col;
  });
}

function isRowEmptyAtoAF(row) {
  for (let c = 1; c <= 32; c++) if (!isEmptyValue(getCellByColumn(row, c)?.value)) return false;
  return true;
}

function hasAnyDataAtoAF(row) {
  for (let c = 1; c <= 32; c++) if (!isEmptyValue(getCellByColumn(row, c)?.value)) return true;
  return false;
}

function matchesWorkedValue(v) {
  const s = String(v || "").trim().toLowerCase();
  return s === "1" || s === "mission";
}

function isEmptyValue(v) { return !v || String(v).trim() === ""; }

function colToIndex(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
  return n;
}

function escapeRegExp(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

function decodeXml(s) {
  return String(s).replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'");
}