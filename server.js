import express from "express";
import multer from "multer";
import JSZip from "jszip";
import ExcelJS from "exceljs";

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

app.post("/generate-navette-paie", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send("Fichier manquant.");
    }

    const planningWb = new ExcelJS.Workbook();
    await planningWb.xlsx.load(req.file.buffer);

    const planningWs = planningWb.worksheets[planningWb.worksheets.length - 1];
    if (!planningWs) {
      return res.status(400).send("Dernière feuille introuvable.");
    }

    const templateWb = new ExcelJS.Workbook();
    await templateWb.xlsx.readFile("templates/navette_paie_template.xlsx");

    const templateWs = templateWb.worksheets[0];
    if (!templateWs) {
      return res.status(500).send("Feuille template introuvable.");
    }

    const extraction = extractEmployeesFromPlanning(planningWs);
    const employees = extraction.employees;
    const summary = buildNavetteSheet(templateWs, employees);

    const debugInfo = {
      employeesCount: employees.length,
      firstEmployee: employees[0] || null,
      firstFiveEmployees: employees.slice(0, 5),
      dateDebug: extraction.dateDebug,
      firstRowsDebug: extraction.firstRowsDebug,
      summaryPreview: summary.slice(0, 20)
    };

    const buffer = await templateWb.xlsx.writeBuffer();

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${buildNavetteOutputName(req.file.originalname)}"`
    );
    res.setHeader(
      "X-Results",
      encodeURIComponent(JSON.stringify(summary.slice(0, 500)))
    );
    res.setHeader(
      "X-Debug-Navette",
      encodeURIComponent(JSON.stringify(debugInfo))
    );

    return res.send(Buffer.from(buffer));
  } catch (err) {
    console.error(err);
    return res.status(500).send("Erreur génération navette paie: " + err.message);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("Server started on port", PORT);
});

function buildOutputName(name) {
  return String(name || "service_charge.xlsx").replace(/\.xlsx$/i, "") + "_traite.xlsx";
}

function buildNavetteOutputName(name) {
  return String(name || "navette_paie.xlsx").replace(/\.xlsx$/i, "") + "_navette_paie.xlsx";
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

function extractEmployeesFromPlanning(ws) {
  const employees = [];
  const START_ROW = 13;
  const DATE_START_COL = 3;
  const DATE_END_COL = 32;

  const rawDateCells = [];
  for (let col = DATE_START_COL; col <= DATE_END_COL; col++) {
    rawDateCells.push(ws.getRow(11).getCell(col).value);
  }

  const year = detectYearFromDateRow(rawDateCells) || 2026;
  const dates = rawDateCells.map(v => normalizePlanningHeaderDate(v, year));

  const dateDebug = rawDateCells.map((v, i) => ({
    col: DATE_START_COL + i,
    raw: debugCellValue(v),
    normalized: dates[i]
  }));

  let rowNumber = START_ROW;
  const firstRowsDebug = [];

  while (true) {
    const row = ws.getRow(rowNumber);

    const mat = cleanCellText(row.getCell(1).value);
    const name = cleanCellText(row.getCell(2).value);

    let allDaysEmpty = true;
    for (let col = DATE_START_COL; col <= DATE_END_COL; col++) {
      const v = cleanCellText(row.getCell(col).value);
      if (v !== "") {
        allDaysEmpty = false;
        break;
      }
    }

    if (mat === "" && name === "" && allDaysEmpty) {
      break;
    }

    const rowCodesDebug = [];
    const blocks = [];
    let current = null;

    for (let idx = 0; idx < dates.length; idx++) {
      const col = DATE_START_COL + idx;
      const rawCell = row.getCell(col).value;
      const code = normalizeCode(rawCell);
      const date = dates[idx];

      rowCodesDebug.push({
        col,
        raw: debugCellValue(rawCell),
        code,
        date
      });

      if (!date || !isTrackedCode(code)) {
        if (current) {
          blocks.push(current);
          current = null;
        }
        continue;
      }

      if (!current) {
        current = {
          type: code,
          start: date,
          end: date,
          days: 1
        };
        continue;
      }

      if (current.type === code) {
        current.end = date;
        current.days += 1;
      } else {
        blocks.push(current);
        current = {
          type: code,
          start: date,
          end: date,
          days: 1
        };
      }
    }

    if (current) {
      blocks.push(current);
    }

    const employee = {
      mat,
      name,
      blocks: blocks.sort((a, b) => a.start.localeCompare(b.start))
    };

    employees.push(employee);

    if (firstRowsDebug.length < 3) {
      firstRowsDebug.push({
        rowNumber,
        mat,
        name,
        blocks: employee.blocks,
        cells: rowCodesDebug
      });
    }

    rowNumber++;
  }

  return { employees, dateDebug, firstRowsDebug };
}

function buildNavetteSheet(ws, employees) {
  const START_ROW = 5;
  const summary = [];

  unmergeDataArea(ws, START_ROW, 1000);

  let totalCA = 0;
  let totalMaladie = 0;
  let totalAT = 0;
  let totalABS = 0;

  let currentRow = START_ROW;

  for (const employee of employees) {
    const packedRows = packBlocksIntoRows(employee.blocks);
    const rowCount = Math.max(1, packedRows.length);

    if (currentRow > START_ROW) {
      for (let i = 0; i < rowCount; i++) {
        ws.insertRow(currentRow + i, []);
        cloneRowStyleFromPrevious(ws, currentRow + i);
      }
    }

    for (let i = 0; i < rowCount; i++) {
      const rowIndex = currentRow + i;
      const row = ws.getRow(rowIndex);
      const packed = packedRows[i] || { CA: null, MALADIE: null, AT: null, ABS: null };

      clearRowValues(row, 1, 20);

      if (i === 0) {
        row.getCell(1).value = employee.mat;
        row.getCell(2).value = employee.name;
      }

      writeBlockToRow(row, packed.CA, 3);
      writeBlockToRow(row, packed.MALADIE, 6);
      writeBlockToRow(row, packed.AT, 9);
      writeBlockToRow(row, packed.ABS, 12);

      if (packed.CA) totalCA += packed.CA.days;
      if (packed.MALADIE) totalMaladie += packed.MALADIE.days;
      if (packed.AT) totalAT += packed.AT.days;
      if (packed.ABS) totalABS += packed.ABS.days;

      summary.push({
        mat: employee.mat,
        name: employee.name,
        rowNumber: rowIndex,
        conge: packed.CA ? `${packed.CA.days}j ${formatDateFr(packed.CA.start)} -> ${formatDateFr(packed.CA.end)}` : "",
        maladie: packed.MALADIE ? `${packed.MALADIE.days}j ${formatDateFr(packed.MALADIE.start)} -> ${formatDateFr(packed.MALADIE.end)}` : "",
        at: packed.AT ? `${packed.AT.days}j ${formatDateFr(packed.AT.start)} -> ${formatDateFr(packed.AT.end)}` : "",
        abs: packed.ABS ? `${packed.ABS.days}j ${formatDateFr(packed.ABS.start)} -> ${formatDateFr(packed.ABS.end)}` : ""
      });
    }

    if (rowCount > 1) {
      ws.mergeCells(currentRow, 1, currentRow + rowCount - 1, 1);
      ws.mergeCells(currentRow, 2, currentRow + rowCount - 1, 2);
    }

    currentRow += rowCount;
  }

  const totalRow = currentRow;
  const preparedRow = currentRow + 2;
  const drhRow = currentRow + 3;

  ws.getRow(totalRow).getCell(3).value = totalCA;
  ws.getRow(totalRow).getCell(6).value = totalMaladie;
  ws.getRow(totalRow).getCell(9).value = totalAT;
  ws.getRow(totalRow).getCell(12).value = totalABS;
  ws.getRow(totalRow).getCell(15).value = 0;
  ws.getRow(totalRow).getCell(19).value = 0;
  ws.getRow(totalRow).getCell(20).value = 0;

  ws.getRow(preparedRow).getCell(2).value = "Préparé par :";
  ws.getRow(drhRow).getCell(2).value = "Direction Ressources Humaines";
  ws.getRow(drhRow).getCell(17).value = "Directeur Général";

  return summary;
}

function packBlocksIntoRows(blocks) {
  const rows = [];

  for (const block of blocks) {
    let placed = false;

    for (const row of rows) {
      if (!row[block.type]) {
        row[block.type] = block;
        placed = true;
        break;
      }
    }

    if (!placed) {
      rows.push({
        CA: null,
        MALADIE: null,
        AT: null,
        ABS: null
      });
      rows[rows.length - 1][block.type] = block;
    }
  }

  return rows;
}

function writeBlockToRow(row, block, startCol) {
  if (!block) return;
  row.getCell(startCol).value = block.days;
  row.getCell(startCol + 1).value = formatDateFr(block.start);
  row.getCell(startCol + 2).value = formatDateFr(block.end);
}

function clearRowValues(row, startCol, endCol) {
  for (let c = startCol; c <= endCol; c++) {
    row.getCell(c).value = null;
  }
}

function unmergeDataArea(ws, startRow, endRow) {
  const merges = Array.from(ws._merges || []);
  for (const merge of merges) {
    const m = typeof merge === "string" ? merge : merge.range;
    const match = String(m).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) continue;

    const r1 = parseInt(match[2], 10);
    const r2 = parseInt(match[4], 10);

    if (r2 >= startRow && r1 <= endRow) {
      try {
        ws.unMergeCells(m);
      } catch {}
    }
  }
}

function cloneRowStyleFromPrevious(ws, rowNumber) {
  const sourceRow = ws.getRow(rowNumber - 1);
  const targetRow = ws.getRow(rowNumber);

  targetRow.height = sourceRow.height;

  for (let c = 1; c <= 20; c++) {
    const sourceCell = sourceRow.getCell(c);
    const targetCell = targetRow.getCell(c);

    targetCell.style = JSON.parse(JSON.stringify(sourceCell.style || {}));
    if (sourceCell.numFmt) targetCell.numFmt = sourceCell.numFmt;
    if (sourceCell.alignment) targetCell.alignment = { ...sourceCell.alignment };
    if (sourceCell.font) targetCell.font = { ...sourceCell.font };
    if (sourceCell.border) targetCell.border = JSON.parse(JSON.stringify(sourceCell.border || {}));
    if (sourceCell.fill) targetCell.fill = JSON.parse(JSON.stringify(sourceCell.fill || {}));
  }
}

function cleanCellText(value) {
  if (value == null) return "";

  if (value instanceof Date) return value.toISOString().slice(0, 10);

  if (typeof value === "object") {
    if (value.text != null) return String(value.text).trim();
    if (Array.isArray(value.richText)) return value.richText.map(x => x.text || "").join("").trim();
    if (value.result != null) return cleanCellText(value.result);
    if (value.formula != null && value.result != null) return cleanCellText(value.result);
    if (value.hyperlink && value.text) return String(value.text).trim();
  }

  return String(value).trim();
}

function normalizeCode(value) {
  return cleanCellText(value).replace(/\s+/g, " ").trim().toUpperCase();
}

function isTrackedCode(code) {
  return code === "CA" || code === "MALADIE" || code === "AT" || code === "ABS";
}

function detectYearFromDateRow(values) {
  for (const v of values) {
    const year = extractYearFromAnyDateValue(v);
    if (year) return year;
  }
  return null;
}

function extractYearFromAnyDateValue(value) {
  if (!value) return null;

  if (value instanceof Date) return value.getFullYear();

  if (typeof value === "object" && value.result instanceof Date) {
    return value.result.getFullYear();
  }

  const s = cleanCellText(value);
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    let y = m[3];
    if (y.length === 2) y = "20" + y;
    return parseInt(y, 10);
  }

  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) return d.getFullYear();

  return null;
}

function normalizePlanningHeaderDate(value, fallbackYear) {
  if (!value) return null;

  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "object" && value.result instanceof Date) {
    return value.result.toISOString().slice(0, 10);
  }

  const s = cleanCellText(value);

  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    const dd = m[1].padStart(2, "0");
    const mm = m[2].padStart(2, "0");
    let yyyy = m[3];
    if (yyyy.length === 2) yyyy = "20" + yyyy;
    return `${yyyy}-${mm}-${dd}`;
  }

  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
  if (m) {
    const dd = m[1].padStart(2, "0");
    const mm = m[2].padStart(2, "0");
    return `${fallbackYear}-${mm}-${dd}`;
  }

  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) {
    return d.toISOString().slice(0, 10);
  }

  return null;
}

function debugCellValue(value) {
  if (value == null) return null;
  if (value instanceof Date) return { type: "Date", value: value.toISOString() };

  if (typeof value === "object") {
    const out = {};
    for (const k of Object.keys(value)) {
      const v = value[k];
      if (v instanceof Date) out[k] = v.toISOString();
      else out[k] = v;
    }
    return out;
  }

  return value;
}

function formatDateFr(isoDate) {
  if (!isoDate) return "";
  const [y, m, d] = isoDate.split("-");
  return `${d}/${m}/${y}`;
}