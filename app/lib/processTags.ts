// Core transform for Flashtalking "ft_tags" trafficking sheets.
//
// What it does, per worksheet, when it finds an `Update_Clicktag1` header:
//   1. In the `pixel` column, replace every cell below the header with a fresh
//      ft.event URL: ?<clientId>;<campaignId>;<placementId>;50126;201;[cachebuster]
//      where clientId comes from the "Client:" field, campaignId from the
//      "Campaign ID:" field, and placementId from each row's `Placement_ID`.
//   2. In the `Update_Clicktag1` column, strip the `us_privacy=${US_PRIVACY}`
//      query param (and its separator) from every URL below the header.
//   3. Delete the entire `Static_Clicktag1` column, shifting later columns left.
//
// The edit is done as surgery on the raw OOXML so everything else about the
// workbook (logo image, fonts, fills, borders, column widths) is preserved.

import { unzipSync, zipSync, strToU8, strFromU8 } from "fflate";

const HEADER_UPDATE = "Update_Clicktag1";
const HEADER_STATIC = "Static_Clicktag1";
const HEADER_PIXEL = "pixel";
const HEADER_PLACEMENT = "Placement_ID";

const LABEL_CLIENT = "Client:";
const LABEL_CAMPAIGN = "Campaign ID:";

// ft.event template: only the first three ;-separated fields are substituted;
// the tail (account;format;cachebuster) is kept exactly as supplied.
const EVENT_BASE = "https://ad-events.flashtalking.com/ft.event";
const EVENT_TAIL = "50126;201;[cachebuster]";

const SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";

export interface SheetReport {
  sheetFile: string;
  updateColumn: string; // original column letter of Update_Clicktag1
  staticColumn: string | null; // original column letter of Static_Clicktag1 (deleted)
  pixelColumn: string | null; // original column letter of the pixel column
  clientId: string | null;
  campaignId: string | null;
  pixelsReplaced: number;
  pixelSample: string | null;
  headerRow: number;
  urlsCleaned: number;
  sample: { before: string; after: string } | null;
  headers: string[];
  rows: string[][];
}

function buildEventUrl(
  clientId: string,
  campaignId: string,
  placementId: string
): string {
  return `${EVENT_BASE}?${clientId};${campaignId};${placementId};${EVENT_TAIL}`;
}

export interface ProcessResult {
  data: Uint8Array;
  fileName: string;
  sheets: SheetReport[];
}

// ----- column-letter helpers ------------------------------------------------

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
  if (!m) return null;
  return { col: m[1], row: parseInt(m[2], 10) };
}

// ----- the us_privacy stripper ----------------------------------------------

export function stripUsPrivacy(url: string): string {
  let out = url;
  // param in the middle / first position: "us_privacy=...&" -> drop param + trailing &
  out = out.replace(/us_privacy=[^&"]*&/gi, "");
  // param at the very end: "&us_privacy=..." / "?us_privacy=..." -> drop leading sep + param
  out = out.replace(/[?&]us_privacy=[^&"]*$/gi, "");
  // tidy any separators we may have left behind
  out = out.replace(/\?&/g, "?").replace(/&&+/g, "&").replace(/[?&]+$/g, "");
  return out;
}

// ----- shared strings -------------------------------------------------------

function parseSharedStrings(xml: string | undefined): string[] {
  if (!xml) return [];
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const sis = Array.from(doc.getElementsByTagNameNS(SML_NS, "si"));
  return sis.map((si) =>
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
    if (!is) return "";
    return Array.from(is.getElementsByTagNameNS(SML_NS, "t"))
      .map((n) => n.textContent ?? "")
      .join("");
  }
  // "str" (formula result) or numeric
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

// ----- header-block field lookups -------------------------------------------

// Finds a label cell (e.g. "Client:") and returns the text of the cell to its
// right — that's where the value sits in these sheets.
function getFieldValue(
  rows: Element[],
  sst: string[],
  label: string
): string | null {
  for (const row of rows) {
    const cells = Array.from(row.getElementsByTagNameNS(SML_NS, "c"));
    const byCol = new Map<number, Element>();
    let labelCol = -1;
    for (const cell of cells) {
      const ref = refParts(cell.getAttribute("r") ?? "");
      if (!ref) continue;
      const ci = colToNum(ref.col);
      byCol.set(ci, cell);
      if (cellText(cell, sst).trim() === label) labelCol = ci;
    }
    if (labelCol > 0) {
      const valueCell = byCol.get(labelCol + 1);
      const value = valueCell ? cellText(valueCell, sst).trim() : "";
      return value || null;
    }
  }
  return null;
}

// "JK Tyre_Ftrack (EV) [35615]" -> "35615"
function extractClientId(clientField: string): string | null {
  const bracket = /\[(\d+)\]/.exec(clientField);
  if (bracket) return bracket[1];
  const nums = clientField.match(/\d+/g);
  return nums ? nums[nums.length - 1] : null;
}

// ----- per-worksheet transform ----------------------------------------------

function processSheet(
  sheetFile: string,
  xml: string,
  sst: string[]
): { xml: string; report: SheetReport } | null {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const rows = Array.from(doc.getElementsByTagNameNS(SML_NS, "row"));

  // 1. locate the table header row (the one carrying Update_Clicktag1)
  let headerRowNum = -1;
  let updateColIdx = -1;
  let staticColIdx = -1;
  let pixelColIdx = -1;
  let placementColIdx = -1;

  for (const row of rows) {
    const cells = Array.from(row.getElementsByTagNameNS(SML_NS, "c"));
    let foundUpdate = -1;
    let foundStatic = -1;
    let foundPixel = -1;
    let foundPlacement = -1;
    for (const cell of cells) {
      const ref = refParts(cell.getAttribute("r") ?? "");
      if (!ref) continue;
      const text = cellText(cell, sst).trim();
      if (text === HEADER_UPDATE) foundUpdate = colToNum(ref.col);
      else if (text === HEADER_STATIC) foundStatic = colToNum(ref.col);
      else if (text === HEADER_PIXEL) foundPixel = colToNum(ref.col);
      else if (text === HEADER_PLACEMENT) foundPlacement = colToNum(ref.col);
    }
    if (foundUpdate > 0) {
      headerRowNum = parseInt(row.getAttribute("r") ?? "-1", 10);
      updateColIdx = foundUpdate;
      staticColIdx = foundStatic;
      pixelColIdx = foundPixel;
      placementColIdx = foundPlacement;
      break;
    }
  }

  if (headerRowNum < 0 || updateColIdx < 0) return null; // not a sheet we handle

  // Pull the campaign-wide IDs from the header block above the table.
  const clientField = getFieldValue(rows, sst, LABEL_CLIENT);
  const clientId = clientField ? extractClientId(clientField) : null;
  const campaignId = getFieldValue(rows, sst, LABEL_CAMPAIGN);

  // 2. clean every URL below the Update_Clicktag1 header (uses ORIGINAL columns)
  let urlsCleaned = 0;
  let sample: { before: string; after: string } | null = null;
  for (const row of rows) {
    const rowNum = parseInt(row.getAttribute("r") ?? "-1", 10);
    if (rowNum <= headerRowNum) continue;
    for (const cell of Array.from(row.getElementsByTagNameNS(SML_NS, "c"))) {
      const ref = refParts(cell.getAttribute("r") ?? "");
      if (!ref || colToNum(ref.col) !== updateColIdx) continue;
      const before = cellText(cell, sst);
      if (!before) continue;
      const after = stripUsPrivacy(before);
      if (after !== before) {
        setCellInline(doc, cell, after);
        urlsCleaned++;
        if (!sample) sample = { before, after };
      }
    }
  }

  // 2b. rebuild the pixel column with fresh ft.event URLs (uses ORIGINAL columns)
  let pixelsReplaced = 0;
  let pixelSample: string | null = null;
  if (pixelColIdx > 0 && placementColIdx > 0 && clientId && campaignId) {
    for (const row of rows) {
      const rowNum = parseInt(row.getAttribute("r") ?? "-1", 10);
      if (rowNum <= headerRowNum) continue;
      const cells = Array.from(row.getElementsByTagNameNS(SML_NS, "c"));
      let pixelCell: Element | null = null;
      let placementId = "";
      for (const cell of cells) {
        const ref = refParts(cell.getAttribute("r") ?? "");
        if (!ref) continue;
        const ci = colToNum(ref.col);
        if (ci === pixelColIdx) pixelCell = cell;
        else if (ci === placementColIdx) placementId = cellText(cell, sst).trim();
      }
      if (!pixelCell || !placementId) continue; // skip empty/padding rows
      const url = buildEventUrl(clientId, campaignId, placementId);
      setCellInline(doc, pixelCell, url);
      pixelsReplaced++;
      if (!pixelSample) pixelSample = url;
    }
  }

  // 3. delete the Static_Clicktag1 column and shift everything to its right left
  const delIdx = staticColIdx; // may be -1 if the column is absent
  if (delIdx > 0) {
    for (const row of rows) {
      for (const cell of Array.from(row.getElementsByTagNameNS(SML_NS, "c"))) {
        const ref = refParts(cell.getAttribute("r") ?? "");
        if (!ref) continue;
        const ci = colToNum(ref.col);
        if (ci === delIdx) {
          row.removeChild(cell);
        } else if (ci > delIdx) {
          cell.setAttribute("r", numToCol(ci - 1) + ref.row);
        }
      }
      // recompute the row span from the cells that remain
      const remaining = Array.from(row.getElementsByTagNameNS(SML_NS, "c"))
        .map((c) => refParts(c.getAttribute("r") ?? "")?.col)
        .filter((c): c is string => !!c)
        .map(colToNum);
      if (row.hasAttribute("spans") && remaining.length) {
        row.setAttribute(
          "spans",
          `${Math.min(...remaining)}:${Math.max(...remaining)}`
        );
      }
    }
    adjustCols(doc, delIdx);
    adjustDimension(doc, delIdx);
  }

  // build a small preview of the resulting table
  const { headers, previewRows } = buildPreview(doc, sst, headerRowNum);

  const serialized = new XMLSerializer().serializeToString(doc);
  return {
    xml: serialized,
    report: {
      sheetFile,
      updateColumn: numToCol(updateColIdx),
      staticColumn: delIdx > 0 ? numToCol(delIdx) : null,
      pixelColumn: pixelColIdx > 0 ? numToCol(pixelColIdx) : null,
      clientId,
      campaignId,
      pixelsReplaced,
      pixelSample,
      headerRow: headerRowNum,
      urlsCleaned,
      sample,
      headers,
      rows: previewRows,
    },
  };
}

function adjustCols(doc: Document, delIdx: number): void {
  const colsEl = doc.getElementsByTagNameNS(SML_NS, "cols")[0];
  if (!colsEl) return;
  for (const col of Array.from(colsEl.getElementsByTagNameNS(SML_NS, "col"))) {
    let min = parseInt(col.getAttribute("min") ?? "0", 10);
    let max = parseInt(col.getAttribute("max") ?? "0", 10);
    if (max < delIdx) {
      // fully before the deleted column — unchanged
    } else if (min > delIdx) {
      min -= 1;
      max -= 1;
    } else {
      // range straddles the deleted column
      max -= 1;
    }
    if (max < min) {
      colsEl.removeChild(col);
    } else {
      col.setAttribute("min", String(min));
      col.setAttribute("max", String(max));
    }
  }
  if (!colsEl.getElementsByTagNameNS(SML_NS, "col").length) {
    colsEl.parentNode?.removeChild(colsEl);
  }
}

function adjustDimension(doc: Document, delIdx: number): void {
  const dim = doc.getElementsByTagNameNS(SML_NS, "dimension")[0];
  if (!dim) return;
  const ref = dim.getAttribute("ref") ?? "";
  const [start, end] = ref.split(":");
  if (!end) return;
  const p = refParts(end);
  if (!p) return;
  const ci = colToNum(p.col);
  const newEnd = (ci >= delIdx ? numToCol(ci - 1) : p.col) + p.row;
  dim.setAttribute("ref", `${start}:${newEnd}`);
}

function buildPreview(
  doc: Document,
  sst: string[],
  headerRowNum: number
): { headers: string[]; previewRows: string[][] } {
  const rows = Array.from(doc.getElementsByTagNameNS(SML_NS, "row"));
  const readRow = (rowNum: number): Record<number, string> => {
    const row = rows.find((r) => r.getAttribute("r") === String(rowNum));
    const out: Record<number, string> = {};
    if (!row) return out;
    for (const cell of Array.from(row.getElementsByTagNameNS(SML_NS, "c"))) {
      const ref = refParts(cell.getAttribute("r") ?? "");
      if (ref) out[colToNum(ref.col)] = cellText(cell, sst);
    }
    return out;
  };

  const headerMap = readRow(headerRowNum);
  const colIdxs = Object.keys(headerMap)
    .map(Number)
    .sort((a, b) => a - b);
  const headers = colIdxs.map((i) => headerMap[i]);

  const previewRows: string[][] = [];
  for (let r = headerRowNum + 1; r <= headerRowNum + 25; r++) {
    const m = readRow(r);
    if (!Object.keys(m).length) continue;
    previewRows.push(colIdxs.map((i) => m[i] ?? ""));
  }
  return { headers, previewRows };
}

// ----- public entry point ---------------------------------------------------

export function processWorkbook(
  buffer: ArrayBuffer,
  originalName: string
): ProcessResult {
  const files = unzipSync(new Uint8Array(buffer));

  const sstName = "xl/sharedStrings.xml";
  const sst = parseSharedStrings(
    files[sstName] ? strFromU8(files[sstName]) : undefined
  );

  const sheetNames = Object.keys(files).filter((n) =>
    /^xl\/worksheets\/sheet[^/]+\.xml$/.test(n)
  );

  const reports: SheetReport[] = [];
  for (const name of sheetNames) {
    const result = processSheet(name, strFromU8(files[name]), sst);
    if (result) {
      files[name] = strToU8(result.xml);
      reports.push(result.report);
    }
  }

  if (!reports.length) {
    throw new Error(
      `No "${HEADER_UPDATE}" header found — this doesn't look like a Flashtalking tag sheet.`
    );
  }

  const data = zipSync(files, { level: 6 });
  const fileName = originalName.replace(/\.xlsx$/i, "") + "_cleaned.xlsx";
  return { data, fileName, sheets: reports };
}
