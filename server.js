import express from "express";
import multer from "multer";
import JSZip from "jszip";
import ExcelJS from "exceljs";
import path from "path";
import { fileURLToPath } from "url";

// Configuration pour récupérer le chemin du répertoire en mode ESM
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
  res.send("Backend Service Charge OK");
});

// --- ROUTE 1 : PROCESS XLSX (Traitement manuel XML) ---
app.post("/process-xlsx", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send("Fichier manquant.");
    }

    const zip = await JSZip.loadAsync(req.file.buffer);

    const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    const sharedStringsXml = await zip.file("xl/sharedStrings.xml")?.async("string");

    if (!workbookXml || !relsXml) {
      return res.status(400).send("Fichier XLSX invalide.");
    }

    const firstSheetPath = resolveFirstWorksheetPath(workbookXml, relsXml);
    if (!firstSheetPath) {
      return res.status(400).send("Première feuille introuvable.");
    }

    const sheetFile = zip.file(firstSheetPath);
    if (!sheetFile) {
      return res.status(400).send("Feuille XML introuvable.");
    }

    let sheetXml = await sheetFile.async("string");
    const sharedStrings = parseSharedStrings(sharedStringsXml || "");
    const rows = parseRows(sheetXml, sharedStrings);

    const result = buildAGUpdates(rows);

    for (const item of result.updates) {
      sheetXml = patchCellValueSafe(sheetXml, item.cellRef, item.value);
    }

    zip.file(firstSheetPath, sheetXml);

    const outputBuffer = await zip.generateAsync({
      type: "nodebuffer",
      compression: "STORE"
    });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${buildOutputName(req.file.originalname)}"`
    );
    res.setHeader(
      "X-Results",
      encodeURIComponent(JSON.stringify(result.results.slice(0, 300)))
    );

    return res.send(outputBuffer);
  } catch (err) {
    console.error(err);
    return res.status(500).send("Erreur traitement XLSX: " + err.message);
  }
});

// --- ROUTE 2 : GENERATE NAVETTE PAIE (ExcelJS corrigé) ---
app.post("/generate-navette-paie", upload.single("file"), async (req, res) => {
  try {
    const templateWb = new ExcelJS.Workbook();
    // Utilisation de path.join pour garantir le chemin sur Render
    const templatePath = path.join(__dirname, "templates", "navette_paie_template.xlsx");
    
    await templateWb.xlsx.readFile(templatePath);

    // Sélection de la première feuille de calcul
    const ws = templateWb.getWorksheet(1);
    if (!ws) {
      return res.status(500).send("Feuille template introuvable.");
    }

    // Ecriture explicite via les lignes pour garantir la visibilité
    ws.getCell("A1").value = "TEST A1";
    
    const row5 = ws.getRow(5);
    row5.getCell(1).value = "TEST MAT"; // A5
    row5.getCell(2).value = "TEST NOM"; // B5
    row5.getCell(3).value = 999;        // C5
    row5.getCell(4).value = "J1";       // D5
    row5.getCell(5).value = "J3";       // E5
    row5.commit();

    const row6 = ws.getRow(6);
    row6.getCell(1).value = "LIGNE 6";
    row6.getCell(2).value = "VISIBLE ?";
    row6.getCell(3).value = 123;
    row6.commit();

    const buffer = await templateWb.xlsx.writeBuffer();

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="navette_paie_generee.xlsx"'
    );

    return res.send(buffer);
  } catch (err) {
    console.error("Erreur Navette:", err);
    return res.status(500).send("Erreur génération navette paie: " + err.message);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("Server started on port", PORT);
});

// --- FONCTIONS UTILITAIRES ---

function buildOutputName(name) {
  return String(name || "service_charge.xlsx").replace(/\.xlsx$/i, "") + "_traite.xlsx";
}

function resolveFirstWorksheetPath(workbookXml, relsXml) {
  const sheetMatch = workbookXml.match(/<sheet[^>]*r:id="([^"]+)"[^>]*\/>/);
  if (!sheetMatch) return null;

  const rid = sheetMatch[1];
  const relRegex = new RegExp(
    `<Relationship[^>]*Id="${escapeRegExp(rid)}"[^>]*Target="([^"]+)"[^>]*/>`
  );
  const relMatch = relsXml.match(relRegex);
  if (!relMatch) return null;

  let target = relMatch[1];
  target = target.replace(/^\/+/, "");
  if (!target.startsWith("xl/")) {
    target = "xl/" + target.replace(/^\.?\//, "");
  }
  return target;
}

function parseSharedStrings(xml) {
  if (!xml) return [];
  const result = [];
  const siMatches = xml.match(/<si[\s\S]*?<\/si>/g) || [];
  for (const si of siMatches) {
    const tMatches = [...si.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)];
    const text = tMatches.map(m => decodeXml(m[1])).join("");
    result.push(text);
  }
  return result;
}

function parseRows(sheetXml, sharedStrings) {
  const rows = [];
  const rowRegex = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
  let rowMatch;

  while ((rowMatch = rowRegex.exec(sheetXml)) !== null) {
    const rowAttrs = rowMatch[1];
    const rowInner = rowMatch[2];
    const rowNumberMatch = rowAttrs.match(/\br="(\d+)"/);
    const rowNumber = rowNumberMatch ? parseInt(rowNumberMatch[1], 10) : null;

    const cells = [];
    const cellRegex = /<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g;
    let cellMatch;

    while ((cellMatch = cellRegex.exec(rowInner)) !== null) {
      const attrs = cellMatch[1] || cellMatch[3] || "";
      const inner = cellMatch[2] || "";
      const refMatch = attrs.match(/\br="([A-Z]+[0-9]+)"/);
      const typeMatch = attrs.match(/\bt="([^"]+)"/);

      const ref = refMatch ? refMatch[1] : null;
      const type = typeMatch ? typeMatch[1] : null;

      let value = "";
      const vMatch = inner.match(/<v>([\s\S]*?)<\/v>/);

      if (vMatch) {
        const raw = decodeXml(vMatch[1]);
        if (type === "s") {
          value = sharedStrings[parseInt(raw, 10)] ?? "";
        } else {
          value = raw;
        }
      } else {
        const isMatch = inner.match(/<is[\s\S]*?>([\s\S]*?)<\/is>/);
        if (isMatch) {
          const tMatches = [...isMatch[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)];
          value = tMatches.map(m => decodeXml(m[1])).join("");
        }
      }
      cells.push({ ref, type, value });
    }
    rows.push({ rowNumber, cells });
  }
  return rows;
}

function buildAGUpdates(rows) {
  const FIRST_DATE_COL = 3;
  const LAST_DATE_COL = 32;
  const RESULT_COL = 33;
  const results = [];
  const updates = [];
  let blockCount = 0;
  let i = 0;

  while (i < rows.length) {
    let markerIndex = -1;
    for (; i < rows.length; i++) {
      const row = rows[i];
      const agCell = getCellByColumn(row, RESULT_COL);
      if (agCell && !isEmptyValue(agCell.value)) {
        markerIndex = i;
        break;
      }
    }
    if (markerIndex === -1) break;

    i = markerIndex + 1;
    let processedInBlock = 0;
    while (i < rows.length) {
      const row = rows[i];
      if (isRowEmptyAtoAF(row)) break;
      if (hasAnyDataAtoAF(row)) {
        let total = 0;
        for (let col = FIRST_DATE_COL; col <= LAST_DATE_COL; col++) {
          const cell = getCellByColumn(row, col);
          if (matchesWorkedValue(cell?.value)) total++;
        }
        const cellRef = `AG${row.rowNumber}`;
        const nameCell = getCellByColumn(row, 2);
        updates.push({ cellRef, value: total });
        results.push({
          section: `BLOC ${blockCount + 1}`,
          rowNumber: row.rowNumber,
          name: String(nameCell?.value || ""),
          total
        });
        processedInBlock++;
      }
      i++;
    }
    if (processedInBlock > 0) blockCount++;
    i++;
  }
  return { updates, results };
}

function patchCellValueSafe(sheetXml, cellRef, numericValue) {
  const rowNumber = parseInt(cellRef.match(/\d+$/)?.[0] || "", 10);
  if (!rowNumber) return sheetXml;

  const rowRegex = new RegExp(`(<row\\b[^>]*\\br="${rowNumber}"[^>]*>)([\\s\\S]*?)(<\\/row>)`);
  const rowMatch = sheetXml.match(rowRegex);
  if (!rowMatch) return sheetXml;

  const rowStart = rowMatch[1];
  const rowInner = rowMatch[2];
  const rowEnd = rowMatch[3];

  const fullCellRegex = new RegExp(`<c([^>]*\\br="${escapeRegExp(cellRef)}"[^>]*)>([\\s\\S]*?)<\\/c>`);
  const selfCellRegex = new RegExp(`<c([^>]*\\br="${escapeRegExp(cellRef)}"[^>]*)\\/>`);

  let newRowInner = rowInner;
  if (fullCellRegex.test(rowInner)) {
    newRowInner = rowInner.replace(fullCellRegex, (match, attrs) => {
      const attrsNoType = attrs.replace(/\s+t="[^"]*"/g, "");
      return `<c${attrsNoType}><v>${numericValue}</v></c>`;
    });
  } else if (selfCellRegex.test(rowInner)) {
    newRowInner = rowInner.replace(selfCellRegex, (match, attrs) => {
      const attrsNoType = attrs.replace(/\s+t="[^"]*"/g, "");
      return `<c${attrsNoType}><v>${numericValue}</v></c>`;
    });
  } else {
    newRowInner += `<c r="${cellRef}"><v>${numericValue}</v></c>`;
  }
  return sheetXml.replace(rowRegex, `${rowStart}${newRowInner}${rowEnd}`);
}

function getCellByColumn(row, colNumber1Based) {
  for (const cell of row.cells) {
    const refCol = getColLetters(cell.ref);
    if (refCol && colToIndex(refCol) === colNumber1Based) {
      return cell;
    }
  }
  return null;
}

function isRowEmptyAtoAF(row) {
  for (let col = 1; col <= 32; col++) {
    const cell = getCellByColumn(row, col);
    if (!isEmptyValue(cell?.value)) return false;
  }
  return true;
}

function hasAnyDataAtoAF(row) {
  for (let col = 1; col <= 32; col++) {
    const cell = getCellByColumn(row, col);
    if (!isEmptyValue(cell?.value)) return true;
  }
  return false;
}

function matchesWorkedValue(value) {
  if (value === null || value === undefined) return false;
  const s = String(value).trim().toLowerCase();
  return s === "1" || s === "mission";
}

function isEmptyValue(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function getColLetters(ref) {
  const m = String(ref || "").match(/^([A-Z]+)/);
  return m ? m[1] : null;
}

function colToIndex(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
}

function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function decodeXml(str) {
  return String(str)
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}