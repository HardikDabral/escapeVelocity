// Placement merge for the campaign Import sheet (.xlsm) from a Prerequisites
// media-plan workbook.
//
// Two phases, because the second step needs the user in the loop:
//
//   analyze()  — reads the media plan (Placement Name + Publisher), figures out
//                which placements need to be appended to the Import sheet, and
//                for each Publisher looks up matching sites in the "Site Lookups"
//                sheet. Returns one plan per placement, each with a list of
//                candidate sites (a Publisher like "Microsoft" matches many).
//
//   build()    — given the user's site choice per placement, appends the new
//                rows (Placement filled) and writes Site + Site_ID from the
//                chosen lookup. Re-zips the workbook untouched otherwise, so the
//                macro-enabled .xlsm (vbaProject.bin) stays valid.

import { unzipSync, zipSync, strToU8, strFromU8 } from "fflate";

const SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";

const SRC_PLACEMENT = "Placement Name";
const SRC_PUBLISHER = "Publisher";
const SRC_DIMENSION = "Ad dimension";
const SRC_LANDING = "Landing Page URL";
const SRC_START_LABEL = "Campaign Start Date";
const SRC_END_LABEL = "Campaign End Date";

const IMP_PLACEMENT = "Placement";
const IMP_PLACEMENT_ANCHOR = "Placement_ID";
const IMP_SITE = "Site";
const IMP_SITE_ID = "Site_ID";
const IMP_DIMENSIONS = "Dimensions";
const IMP_START = "Start_Date";
const IMP_END = "End_Date";
const IMP_CLICKTAG = "Clicktag_1";

const LOOKUP_SHEET = "Site Lookups";

const MONTHS: Record<string, string> = {
  jan: "01", feb: "02", mar: "03", apr: "04", may: "05", jun: "06",
  jul: "07", aug: "08", sep: "09", oct: "10", nov: "11", dec: "12",
};

// "5th Jun 2026" / "15th July 2026" -> "2026-06-05" / "2026-07-15".
// Already-ISO or unrecognised input is returned unchanged.
function toISODate(raw: string): string {
  const t = raw.trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return t;
  const m = /^(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\.?\s+(\d{4})$/.exec(t);
  if (!m) return t;
  const mon = MONTHS[m[2].toLowerCase().slice(0, 3)];
  if (!mon) return t;
  return `${m[3]}-${mon}-${m[1].padStart(2, "0")}`;
}

export interface SiteCandidate {
  site: string; // e.g. "Finecast (16742)"
  siteId: string; // e.g. "16742"
}

export interface PlacementPlan {
  name: string; // placement name
  publisher: string; // publisher from the media plan
  action: "append" | "existing"; // existing = already in the Import sheet
  candidates: SiteCandidate[]; // site lookup matches for this publisher
}

export interface AnalyzeResult {
  plans: PlacementPlan[];
  importSheetName: string;
  columns: { siteId: string; site: string; placement: string };
  startDate: string; // ISO, from the media plan header block
  endDate: string;
}

// site choice keyed by placement name; absent/null means "leave blank"
export type SiteSelections = Record<string, SiteCandidate | null>;

export interface BuildReport {
  appended: { name: string; row: number; site: string | null }[];
  alreadyPresent: string[];
  importSheetName: string;
  // 0-based indices within `headers`; -1 when the column is absent
  columns: {
    siteId: number; site: number; placement: number;
    dimensions: number; start: number; end: number; clicktag: number;
  };
  headers: string[];
  rows: string[][];
}

export interface BuildResult {
  data: Uint8Array;
  fileName: string;
  report: BuildReport;
}

// ----- low-level helpers ----------------------------------------------------

function colToNum(col: string): number {
  let n = 0;
  for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
function numToCol(n: number): string {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function refParts(ref: string): { col: string; row: number } | null {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  return m ? { col: m[1], row: parseInt(m[2], 10) } : null;
}

function parseSharedStrings(xml: string | undefined): string[] {
  if (!xml) return [];
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  return Array.from(doc.getElementsByTagNameNS(SML_NS, "si")).map((si) =>
    Array.from(si.getElementsByTagNameNS(SML_NS, "t"))
      .map((t) => t.textContent ?? "")
      .join("")
  );
}

function cellText(cell: Element, sst: string[]): string {
  const t = cell.getAttribute("t");
  if (t === "s") {
    const v = cell.getElementsByTagNameNS(SML_NS, "v")[0];
    const idx = v ? parseInt(v.textContent ?? "-1", 10) : -1;
    return idx >= 0 && idx < sst.length ? sst[idx] : "";
  }
  if (t === "inlineStr") {
    const is = cell.getElementsByTagNameNS(SML_NS, "is")[0];
    return is
      ? Array.from(is.getElementsByTagNameNS(SML_NS, "t"))
          .map((n) => n.textContent ?? "")
          .join("")
      : "";
  }
  const v = cell.getElementsByTagNameNS(SML_NS, "v")[0];
  return v?.textContent ?? "";
}

function setCellInline(doc: Document, cell: Element, text: string): void {
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  cell.setAttribute("t", "inlineStr");
  const is = doc.createElementNS(SML_NS, "is");
  const t = doc.createElementNS(SML_NS, "t");
  t.setAttributeNS(XML_NS, "xml:space", "preserve");
  t.textContent = text;
  is.appendChild(t);
  cell.appendChild(is);
}

function blankCell(cell: Element): void {
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  cell.removeAttribute("t");
}

function worksheetFiles(files: Record<string, Uint8Array>): string[] {
  return Object.keys(files)
    .filter((n) => /^xl\/worksheets\/sheet[^/]+\.xml$/.test(n))
    .sort();
}

// numbers in the lookup come through as "16742.0" — normalise to "16742"
function cleanId(raw: string): string {
  const t = raw.trim();
  const m = /^(\d+)(?:\.0+)?$/.exec(t);
  return m ? m[1] : t;
}

// ----- workbook lookups -----------------------------------------------------

function findSheetFileByName(
  files: Record<string, Uint8Array>,
  name: string
): string | null {
  const wb = files["xl/workbook.xml"] ? strFromU8(files["xl/workbook.xml"]) : "";
  const rels = files["xl/_rels/workbook.xml.rels"]
    ? strFromU8(files["xl/_rels/workbook.xml.rels"])
    : "";
  const sheet = new RegExp(`<sheet[^>]*name="${name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*/>`).exec(wb)?.[0];
  if (!sheet) return null;
  const relId = /r:id="(rId\d+)"/.exec(sheet)?.[1];
  if (!relId) return null;
  const target = new RegExp(`Id="${relId}"[^>]*Target="([^"]*)"`).exec(rels)?.[1];
  if (!target) return null;
  return target.startsWith("/") ? target.slice(1) : "xl/" + target.replace(/^\.\//, "");
}

function sheetNameForFile(files: Record<string, Uint8Array>, sheetFile: string): string {
  const wb = files["xl/workbook.xml"] ? strFromU8(files["xl/workbook.xml"]) : "";
  const rels = files["xl/_rels/workbook.xml.rels"]
    ? strFromU8(files["xl/_rels/workbook.xml.rels"])
    : "";
  const target = sheetFile.replace(/^xl\//, "");
  const relId = new RegExp(
    `Id="(rId\\d+)"[^>]*Target="(?:/xl/)?${target.replace(/[-/\\^$*+?.()|[\]{}]/g, "\\$&")}"`
  ).exec(rels)?.[1];
  if (relId) {
    const m = new RegExp(`<sheet[^>]*r:id="${relId}"[^>]*/>`).exec(wb)?.[0] ?? "";
    const name = /name="([^"]*)"/.exec(m)?.[1];
    if (name) return name;
  }
  return sheetFile;
}

interface HeaderHit {
  file: string;
  doc: Document;
  rows: Element[];
  headerRow: number;
  cols: Record<string, number>;
}

// Scans every worksheet for a header row that contains all `wanted` header
// labels, returning the column index of each.
function locateHeaders(
  files: Record<string, Uint8Array>,
  sst: string[],
  wanted: string[]
): HeaderHit | null {
  for (const file of worksheetFiles(files)) {
    const doc = new DOMParser().parseFromString(strFromU8(files[file]), "application/xml");
    const rows = Array.from(doc.getElementsByTagNameNS(SML_NS, "row"));
    for (const row of rows) {
      const cols: Record<string, number> = {};
      for (const cell of Array.from(row.getElementsByTagNameNS(SML_NS, "c"))) {
        const ref = refParts(cell.getAttribute("r") ?? "");
        if (!ref) continue;
        const text = cellText(cell, sst).trim();
        if (wanted.includes(text) && cols[text] === undefined) cols[text] = colToNum(ref.col);
      }
      if (wanted.every((w) => cols[w] !== undefined)) {
        return { file, doc, rows, headerRow: parseInt(row.getAttribute("r") ?? "-1", 10), cols };
      }
    }
  }
  return null;
}

// ----- source / lookup readers ----------------------------------------------

interface SourceEntry {
  name: string;
  publisher: string;
  dimension: string;
  landingUrl: string;
}

interface SourceData {
  entries: SourceEntry[];
  startDate: string; // ISO
  endDate: string;
}

// reads a labelled value from a header block: a cell == `label`, value to its right
function fieldValue(rows: Element[], sst: string[], label: string): string {
  for (const row of rows) {
    const byCol = new Map<number, Element>();
    let labelCol = -1;
    for (const cell of Array.from(row.getElementsByTagNameNS(SML_NS, "c"))) {
      const ref = refParts(cell.getAttribute("r") ?? "");
      if (!ref) continue;
      byCol.set(colToNum(ref.col), cell);
      if (cellText(cell, sst).trim() === label) labelCol = colToNum(ref.col);
    }
    if (labelCol > 0) {
      const v = byCol.get(labelCol + 1);
      return v ? cellText(v, sst).trim() : "";
    }
  }
  return "";
}

// finds a single column index within a header row by label (0 if absent)
function colInRow(rows: Element[], sst: string[], headerRow: number, label: string): number {
  const row = rows.find((r) => r.getAttribute("r") === String(headerRow));
  if (!row) return 0;
  for (const cell of Array.from(row.getElementsByTagNameNS(SML_NS, "c"))) {
    const ref = refParts(cell.getAttribute("r") ?? "");
    if (ref && cellText(cell, sst).trim() === label) return colToNum(ref.col);
  }
  return 0;
}

function readSource(buffer: ArrayBuffer): SourceData {
  const files = unzipSync(new Uint8Array(buffer));
  const sst = parseSharedStrings(
    files["xl/sharedStrings.xml"] ? strFromU8(files["xl/sharedStrings.xml"]) : undefined
  );
  const hit = locateHeaders(files, sst, [SRC_PLACEMENT, SRC_PUBLISHER]);
  if (!hit) throw new Error(`Couldn't find "${SRC_PLACEMENT}" / "${SRC_PUBLISHER}" columns in the first file.`);
  const nameCol = hit.cols[SRC_PLACEMENT];
  const pubCol = hit.cols[SRC_PUBLISHER];
  const dimCol = colInRow(hit.rows, sst, hit.headerRow, SRC_DIMENSION); // 0 if absent
  const landingCol = colInRow(hit.rows, sst, hit.headerRow, SRC_LANDING);

  const entries: SourceEntry[] = [];
  for (const row of hit.rows) {
    const rowNum = parseInt(row.getAttribute("r") ?? "-1", 10);
    if (rowNum <= hit.headerRow) continue;
    let name = "";
    let publisher = "";
    let dimension = "";
    let landingUrl = "";
    for (const cell of Array.from(row.getElementsByTagNameNS(SML_NS, "c"))) {
      const ref = refParts(cell.getAttribute("r") ?? "");
      if (!ref) continue;
      const ci = colToNum(ref.col);
      if (ci === nameCol) name = cellText(cell, sst).trim();
      else if (ci === pubCol) publisher = cellText(cell, sst).trim();
      else if (dimCol && ci === dimCol) dimension = cellText(cell, sst).trim();
      else if (landingCol && ci === landingCol) landingUrl = cellText(cell, sst).trim();
    }
    if (name) entries.push({ name, publisher, dimension, landingUrl });
  }

  return {
    entries,
    startDate: toISODate(fieldValue(hit.rows, sst, SRC_START_LABEL)),
    endDate: toISODate(fieldValue(hit.rows, sst, SRC_END_LABEL)),
  };
}

interface IndexedSite extends SiteCandidate {
  lower: string;
}

function readSiteLookups(files: Record<string, Uint8Array>, sst: string[]): IndexedSite[] {
  const sheetFile = findSheetFileByName(files, LOOKUP_SHEET);
  if (!sheetFile || !files[sheetFile]) return [];
  const doc = new DOMParser().parseFromString(strFromU8(files[sheetFile]), "application/xml");
  const rows = Array.from(doc.getElementsByTagNameNS(SML_NS, "row"));

  // header row of the lookup: Site + Site_ID
  let headerRow = -1;
  let siteCol = -1;
  let idCol = -1;
  for (const row of rows) {
    let s = -1;
    let i = -1;
    for (const cell of Array.from(row.getElementsByTagNameNS(SML_NS, "c"))) {
      const ref = refParts(cell.getAttribute("r") ?? "");
      if (!ref) continue;
      const text = cellText(cell, sst).trim();
      if (text === IMP_SITE) s = colToNum(ref.col);
      else if (text === IMP_SITE_ID) i = colToNum(ref.col);
    }
    if (s > 0 && i > 0) {
      headerRow = parseInt(row.getAttribute("r") ?? "-1", 10);
      siteCol = s;
      idCol = i;
      break;
    }
  }
  if (headerRow < 0) return [];

  const out: IndexedSite[] = [];
  for (const row of rows) {
    const rowNum = parseInt(row.getAttribute("r") ?? "-1", 10);
    if (rowNum <= headerRow) continue;
    let site = "";
    let siteId = "";
    for (const cell of Array.from(row.getElementsByTagNameNS(SML_NS, "c"))) {
      const ref = refParts(cell.getAttribute("r") ?? "");
      if (!ref) continue;
      const ci = colToNum(ref.col);
      if (ci === siteCol) site = cellText(cell, sst).trim();
      else if (ci === idCol) siteId = cleanId(cellText(cell, sst));
    }
    if (site) out.push({ site, siteId, lower: site.toLowerCase() });
  }
  return out;
}

function matchCandidates(publisher: string, index: IndexedSite[]): SiteCandidate[] {
  const needle = publisher.trim().toLowerCase();
  if (!needle) return [];
  const hits = index.filter((s) => s.lower.includes(needle));
  // sort: names that start with the publisher first, then alphabetical
  hits.sort((a, b) => {
    const as = a.lower.startsWith(needle) ? 0 : 1;
    const bs = b.lower.startsWith(needle) ? 0 : 1;
    if (as !== bs) return as - bs;
    return a.site.localeCompare(b.site);
  });
  return hits.map(({ site, siteId }) => ({ site, siteId }));
}

// ----- phase 1: analyze -----------------------------------------------------

export function analyze(sourceBuffer: ArrayBuffer, targetBuffer: ArrayBuffer): AnalyzeResult {
  const src = readSource(sourceBuffer);

  const tgtFiles = unzipSync(new Uint8Array(targetBuffer));
  const tgtSst = parseSharedStrings(
    tgtFiles["xl/sharedStrings.xml"] ? strFromU8(tgtFiles["xl/sharedStrings.xml"]) : undefined
  );
  const imp = locateHeaders(tgtFiles, tgtSst, [IMP_PLACEMENT, IMP_PLACEMENT_ANCHOR, IMP_SITE, IMP_SITE_ID]);
  if (!imp) throw new Error(`Couldn't find the Import sheet (needs "${IMP_PLACEMENT}" and "${IMP_SITE_ID}" columns) in the second file.`);

  const existing = existingPlacements(imp, tgtSst);
  const siteIndex = readSiteLookups(tgtFiles, tgtSst);

  const plans: PlacementPlan[] = src.entries.map((e) => ({
    name: e.name,
    publisher: e.publisher,
    action: existing.has(e.name) ? "existing" : "append",
    candidates: matchCandidates(e.publisher, siteIndex),
  }));

  return {
    plans,
    importSheetName: sheetNameForFile(tgtFiles, imp.file),
    columns: {
      siteId: numToCol(imp.cols[IMP_SITE_ID]),
      site: numToCol(imp.cols[IMP_SITE]),
      placement: numToCol(imp.cols[IMP_PLACEMENT]),
    },
    startDate: src.startDate,
    endDate: src.endDate,
  };
}

function existingPlacements(imp: HeaderHit, sst: string[]): Set<string> {
  const out = new Set<string>();
  const col = imp.cols[IMP_PLACEMENT];
  for (const row of imp.rows) {
    const rowNum = parseInt(row.getAttribute("r") ?? "-1", 10);
    if (rowNum <= imp.headerRow) continue;
    for (const cell of Array.from(row.getElementsByTagNameNS(SML_NS, "c"))) {
      const ref = refParts(cell.getAttribute("r") ?? "");
      if (!ref || colToNum(ref.col) !== col) continue;
      const v = cellText(cell, sst).trim();
      if (v) out.add(v);
    }
  }
  return out;
}

// ----- phase 2: build -------------------------------------------------------

export function build(
  sourceBuffer: ArrayBuffer,
  targetBuffer: ArrayBuffer,
  targetName: string,
  selections: SiteSelections
): BuildResult {
  const src = readSource(sourceBuffer);
  const entries = src.entries;

  const tgtFiles = unzipSync(new Uint8Array(targetBuffer));
  const tgtSst = parseSharedStrings(
    tgtFiles["xl/sharedStrings.xml"] ? strFromU8(tgtFiles["xl/sharedStrings.xml"]) : undefined
  );
  const imp = locateHeaders(tgtFiles, tgtSst, [IMP_PLACEMENT, IMP_PLACEMENT_ANCHOR, IMP_SITE, IMP_SITE_ID]);
  if (!imp) throw new Error("Couldn't find the Import sheet in the second file.");

  const { doc, rows } = imp;
  const placeCol = imp.cols[IMP_PLACEMENT];
  const siteCol = imp.cols[IMP_SITE];
  const idCol = imp.cols[IMP_SITE_ID];
  const placementIdCol = imp.cols[IMP_PLACEMENT_ANCHOR];
  // optional columns (0 = absent)
  const dimCol = colInRow(rows, tgtSst, imp.headerRow, IMP_DIMENSIONS);
  const startCol = colInRow(rows, tgtSst, imp.headerRow, IMP_START);
  const endCol = colInRow(rows, tgtSst, imp.headerRow, IMP_END);
  const clickCol = colInRow(rows, tgtSst, imp.headerRow, IMP_CLICKTAG);

  // Columns we set explicitly on new rows. Anything NOT in here is copied from
  // the cloned template row — that's how the constant columns (Placement_Type,
  // Status, Hidden, Stopped, Booked_Units) get their shared "same" values. A
  // managed column with no value for a given row is blanked, so a new row never
  // inherits the template's Site / Placement_ID / Clicktag.
  const managedCols = new Set<number>([placeCol, siteCol, idCol, placementIdCol]);
  for (const c of [dimCol, startCol, endCol, clickCol]) if (c) managedCols.add(c);

  // existing placements, last data row, template row, used row numbers
  const existing = new Set<string>();
  let lastDataRow = imp.headerRow;
  let templateRow: Element | null = null;
  const usedRowNums = new Set<number>();
  for (const row of rows) {
    const rowNum = parseInt(row.getAttribute("r") ?? "-1", 10);
    usedRowNums.add(rowNum);
    if (rowNum <= imp.headerRow) continue;
    for (const cell of Array.from(row.getElementsByTagNameNS(SML_NS, "c"))) {
      const ref = refParts(cell.getAttribute("r") ?? "");
      if (!ref || colToNum(ref.col) !== placeCol) continue;
      const v = cellText(cell, tgtSst).trim();
      if (v) {
        existing.add(v);
        if (rowNum > lastDataRow) lastDataRow = rowNum;
        templateRow = row;
      }
    }
  }

  const sheetData = doc.getElementsByTagNameNS(SML_NS, "sheetData")[0];
  const appended: { name: string; row: number; site: string | null }[] = [];
  const alreadyPresent: string[] = [];

  let nextRow = lastDataRow + 1;
  const seen = new Set(existing);
  for (const entry of entries) {
    if (seen.has(entry.name)) {
      alreadyPresent.push(entry.name);
      continue;
    }
    seen.add(entry.name);
    while (usedRowNums.has(nextRow)) nextRow++;
    const rowNum = nextRow;
    usedRowNums.add(rowNum);

    const choice = selections[entry.name] ?? null;
    const fills: Record<number, string> = { [placeCol]: entry.name };
    if (choice) {
      fills[siteCol] = choice.site;
      fills[idCol] = choice.siteId;
    }
    if (dimCol && entry.dimension) fills[dimCol] = entry.dimension;
    if (startCol && src.startDate) fills[startCol] = src.startDate;
    if (endCol && src.endDate) fills[endCol] = src.endDate;
    if (clickCol && entry.landingUrl) fills[clickCol] = entry.landingUrl;
    sheetData.appendChild(buildRow(doc, templateRow, rowNum, fills, managedCols));
    appended.push({ name: entry.name, row: rowNum, site: choice ? choice.site : null });
  }

  // keep rows ascending
  const allRows = Array.from(sheetData.getElementsByTagNameNS(SML_NS, "row"));
  allRows.sort(
    (a, b) => parseInt(a.getAttribute("r") ?? "0", 10) - parseInt(b.getAttribute("r") ?? "0", 10)
  );
  for (const r of allRows) sheetData.appendChild(r);

  updateDimension(doc);
  tgtFiles[imp.file] = strToU8(new XMLSerializer().serializeToString(doc));

  const { headers, previewRows } = buildPreview(doc, tgtSst, imp.headerRow);
  const data = zipSync(tgtFiles, { level: 6 });
  const ext = /\.xlsm$/i.test(targetName) ? ".xlsm" : ".xlsx";
  const fileName = targetName.replace(/\.(xlsm|xlsx)$/i, "") + "_placements" + ext;

  return {
    data,
    fileName,
    report: {
      appended,
      alreadyPresent,
      importSheetName: sheetNameForFile(tgtFiles, imp.file),
      columns: {
        siteId: idCol - 1,
        site: siteCol - 1,
        placement: placeCol - 1,
        dimensions: dimCol ? dimCol - 1 : -1,
        start: startCol ? startCol - 1 : -1,
        end: endCol ? endCol - 1 : -1,
        clicktag: clickCol ? clickCol - 1 : -1,
      },
      headers,
      rows: previewRows,
    },
  };
}

// Builds a new <row>, cloning a data row for styling when available. For each
// cloned cell: if the column has a fill -> set it; else if the column is
// "managed" -> blank it; otherwise keep the template value (the constant
// columns). `fills` and `managedCols` use 1-based column indices.
function buildRow(
  doc: Document,
  templateRow: Element | null,
  rowNum: number,
  fills: Record<number, string>,
  managedCols: Set<number>
): Element {
  if (templateRow) {
    const newRow = templateRow.cloneNode(true) as Element;
    newRow.setAttribute("r", String(rowNum));
    newRow.removeAttribute("spans");
    for (const cell of Array.from(newRow.getElementsByTagNameNS(SML_NS, "c"))) {
      const ref = refParts(cell.getAttribute("r") ?? "");
      if (!ref) continue;
      cell.setAttribute("r", ref.col + rowNum);
      const ci = colToNum(ref.col);
      const fill = fills[ci];
      if (fill !== undefined) setCellInline(doc, cell, fill);
      else if (managedCols.has(ci)) blankCell(cell);
      // otherwise keep the template's value (constant columns)
    }
    return newRow;
  }
  const row = doc.createElementNS(SML_NS, "row");
  row.setAttribute("r", String(rowNum));
  for (const [colStr, text] of Object.entries(fills)) {
    const cell = doc.createElementNS(SML_NS, "c");
    cell.setAttribute("r", numToCol(Number(colStr)) + rowNum);
    setCellInline(doc, cell, text);
    row.appendChild(cell);
  }
  return row;
}

function updateDimension(doc: Document): void {
  const dim = doc.getElementsByTagNameNS(SML_NS, "dimension")[0];
  if (!dim) return;
  const ref = dim.getAttribute("ref") ?? "";
  const [start, end] = ref.split(":");
  if (!end) return;
  const p = refParts(end);
  if (!p) return;
  let maxRow = p.row;
  for (const row of Array.from(doc.getElementsByTagNameNS(SML_NS, "row"))) {
    maxRow = Math.max(maxRow, parseInt(row.getAttribute("r") ?? "0", 10));
  }
  dim.setAttribute("ref", `${start}:${p.col}${maxRow}`);
}

function buildPreview(
  doc: Document,
  sst: string[],
  headerRow: number
): { headers: string[]; previewRows: string[][] } {
  const rows = Array.from(doc.getElementsByTagNameNS(SML_NS, "row"));
  const read = (rowNum: number): Record<number, string> => {
    const row = rows.find((r) => r.getAttribute("r") === String(rowNum));
    const out: Record<number, string> = {};
    if (!row) return out;
    for (const cell of Array.from(row.getElementsByTagNameNS(SML_NS, "c"))) {
      const ref = refParts(cell.getAttribute("r") ?? "");
      if (ref) out[colToNum(ref.col)] = cellText(cell, sst);
    }
    return out;
  };
  const headerMap = read(headerRow);
  const colIdxs = Object.keys(headerMap)
    .map(Number)
    .sort((a, b) => a - b);
  const headers = colIdxs.map((i) => headerMap[i]);
  const dataRowNums = rows
    .map((r) => parseInt(r.getAttribute("r") ?? "-1", 10))
    .filter((n) => n > headerRow)
    .sort((a, b) => a - b);
  const previewRows: string[][] = [];
  for (const rn of dataRowNums) {
    const m = read(rn);
    if (!colIdxs.some((i) => (m[i] ?? "").trim())) continue;
    previewRows.push(colIdxs.map((i) => m[i] ?? ""));
  }
  return { headers, previewRows };
}
