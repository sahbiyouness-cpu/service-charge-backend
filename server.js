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

// Middleware CORS complet
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, X-Results");
  res.setHeader("Access-Control-Expose-Headers", "X-Results");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS, GET");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

app.get("/", (req, res) => {
  res.send("Backend Service Charge & Navette Paie OK");
});

// =========================================================================
// ROUTE 1 : SERVICE CHARGE (Analyse XML ultra-rapide par JSZip)
// =========================================================================
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

// =========================================================================
// ROUTE 2 : GENERATE NAVETTE PAIE (Via ExcelJS avec Template)
// =========================================================================
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

    // Récupération des en-têtes de dates (Ligne 11, colonnes C à AF)
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

      const absences = extractNavetteAbsences(row, datesMap);

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

    // Calcul et injection de la ligne des totaux globaux (Ligne 62)
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
    res.setHeader("Content-Disposition", "attachment; filename=navette_paie_generee.xlsx");
    res.setHeader("X-Results", encodeURIComponent(JSON.stringify(summaryResults)));
    return res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).send("Erreur Navette: " + err.message);
  }
});

// =========================================================================
// FONCTIONS UTILITAIRES (SERVICE CHARGE)
// =========================================================================
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
  const FIRST_DATE_COL = 3; // C
  const LAST_DATE_COL = 32; // AF
  const RESULT_COL = 33;    // AG

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

      if (isRowEmptyAtoAF(row)) {
        break;
      }

      if (hasAnyDataAtoAF(row)) {
        let total = 0;

        for (let col = FIRST_DATE_COL; col <= LAST_DATE_COL; col++) {
          const cell = getCellByColumn(row, col);
          if (matchesWorkedValue(cell?.value)) total++;
        }

        const cellRef = `AG${row.rowNumber}`;
        const nameCell = getCellByColumn(row, 2);

        updates.push({
          cellRef,
          value: total
        });

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

    if (processedInBlock > 0) {
      blockCount++;
    }
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

// Correction du nom pour éviter toute confusion de scope
function extractNavetteAbsences(row, datesMap) {
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

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`Listening on ${PORT}`));