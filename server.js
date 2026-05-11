import express from "express";
import multer from "multer";
import JSZip from "jszip";
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

// --- ROUTE 1 : SERVICE CHARGE (Inchangée) ---
app.post("/process-xlsx", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).send("Fichier manquant.");
    const zip = await JSZip.loadAsync(req.file.buffer);
    const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    const sharedStringsXml = await zip.file("xl/sharedStrings.xml")?.async("string");

    const firstSheetPath = resolveFirstWorksheetPath(workbookXml, relsXml);
    let sheetXml = await zip.file(firstSheetPath).async("string");
    
    const sharedStrings = parseSharedStrings(sharedStringsXml || "");
    const rows = parseRows(sheetXml, sharedStrings);
    const result = buildAGUpdates(rows);

    for (const item of result.updates) {
      sheetXml = patchCellValueSafe(sheetXml, item.cellRef, item.value);
    }

    zip.file(firstSheetPath, sheetXml);
    const outputBuffer = await zip.generateAsync({ type: "nodebuffer" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    return res.send(outputBuffer);
  } catch (err) {
    res.status(500).send(err.message);
  }
});

// --- ROUTE 2 : NAVETTE PAIE (Version JSZip Ultra-Robuste) ---
app.post("/generate-navette-paie", upload.single("file"), async (req, res) => {
  try {
    console.log("--- DEBUT GENERATION NAVETTE (JSZip Mode) ---");
    const templatePath = path.join(process.cwd(), "templates", "navette_paie_template.xlsx");

    if (!fs.existsSync(templatePath)) {
      throw new Error("Fichier template introuvable sur le disque.");
    }

    // 1. Charger le template via JSZip
    const templateData = fs.readFileSync(templatePath);
    const zip = await JSZip.loadAsync(templateData);

    // 2. Trouver la feuille JANVIER 2026 (via workbook.xml)
    const workbookXml = await zip.file("xl/workbook.xml").async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    
    // On cherche l'ID de la feuille qui s'appelle "JANVIER 2026"
    const sheetMatch = workbookXml.match(/<sheet [^>]*name="JANVIER 2026"[^>]*r:id="([^"]+)"/i) 
                    || workbookXml.match(/<sheet [^>]*r:id="([^"]+)"[^>]*name="JANVIER 2026"/i);
    
    let sheetPath = "";
    if (sheetMatch) {
      const rId = sheetMatch[1];
      const relRegex = new RegExp(`Relationship [^>]*Id="${rId}" [^>]*Target="([^"]+)"`);
      const relMatch = relsXml.match(relRegex);
      if (relMatch) {
        sheetPath = "xl/" + relMatch[1].replace("../", "");
      }
    } else {
      // Si pas trouvé par nom, on prend la première feuille par défaut
      sheetPath = resolveFirstWorksheetPath(workbookXml, relsXml);
    }

    console.log("Cible détectée :", sheetPath);

    let sheetXml = await zip.file(sheetPath).async("string");

    // 3. Injection directe des données (comme dans ton autre route)
    // A5 = MAT, B5 = NOM, C5 = CP
    sheetXml = patchCellValueSafe(sheetXml, "A5", "999");
    sheetXml = patchCellValueSafe(sheetXml, "B5", "INJECTION JSZIP OK");
    sheetXml = patchCellValueSafe(sheetXml, "C5", "10");

    // 4. Reconstruction du fichier
    zip.file(sheetPath, sheetXml);
    const outputBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });

    console.log("Succès : Fichier généré via JSZip");

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="navette_paie.xlsx"');
    return res.send(outputBuffer);

  } catch (err) {
    console.error("ERREUR :", err.message);
    return res.status(500).send("Erreur : " + err.message);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("Server port:", PORT));

// --- UTILITAIRES RECOPIÉS ---
function resolveFirstWorksheetPath(workbookXml, relsXml) {
  const rid = workbookXml.match(/<sheet[^>]*r:id="([^"]+)"[^>]*\/>/)?.[1];
  const relMatch = relsXml.match(new RegExp(`<Relationship[^>]*Id="${rid}"[^>]*Target="([^"]+)"[^>]*/>`));
  let t = relMatch[1].replace(/^\/+/, "");
  return t.startsWith("xl/") ? t : "xl/" + t;
}
function patchCellValueSafe(xml, ref, val) {
  const rowNum = ref.match(/\d+$/)?.[0];
  const rowReg = new RegExp(`(<row\\b[^>]*\\br="${rowNum}"[^>]*>)([\\s\\S]*?)(<\\/row>)`);
  const m = xml.match(rowReg);
  if (!m) return xml;
  const cReg = new RegExp(`<c([^>]*\\br="${ref}"[^>]*)>([\\s\\S]*?)<\\/c>|<c([^>]*\\br="${ref}"[^>]*)\\/>`);
  let inner = m[2];
  // Si c'est un nombre, on ne met pas t="s", si c'est du texte, idéalement il faudrait sharedStrings, 
  // mais pour un test, on peut injecter en mode 'inlineStr' ou forcer la valeur
  const newVal = `<c r="${ref}"><v>${val}</v></c>`; 
  if (cReg.test(inner)) inner = inner.replace(cReg, newVal);
  else inner += newVal;
  return xml.replace(rowReg, `${m[1]}${inner}${m[3]}`);
}
function parseSharedStrings(xml) {
  if (!xml) return [];
  return (xml.match(/<si[\s\S]*?<\/si>/g) || []).map(si => {
    return [...si.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(m => m[1]).join("");
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
      cells.push({ ref: attrs.match(/\br="([A-Z]+[0-9]+)"/)?.[1], value: val });
    }
    rows.push({ rowNumber: rowNum ? parseInt(rowNum, 10) : null, cells });
  }
  return rows;
}
function buildAGUpdates(rows) { return { updates: [], results: [] }; } // Simplifié pour l'exemple