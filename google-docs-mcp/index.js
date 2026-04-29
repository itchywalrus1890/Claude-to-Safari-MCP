#!/usr/bin/env node
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { google } from "googleapis";
import fs from "fs/promises";
import path from "path";
import os from "os";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

async function loadAccounts() {
  const files = await fs.readdir(__dirname);
  const tokenFiles = files.filter((f) => /^token(?:\..+)?\.json$/.test(f)).sort();
  const accounts = [];
  for (const file of tokenFiles) {
    const m = file.match(/^token(?:\.(.+))?\.json$/);
    const label = m[1] || "default";
    const data = JSON.parse(await fs.readFile(path.join(__dirname, file), "utf8"));
    const auth = google.auth.fromJSON(data);
    accounts.push({
      label,
      email: data.email || null,
      auth,
      docs: google.docs({ version: "v1", auth }),
      sheets: google.sheets({ version: "v4", auth }),
    });
  }
  return accounts;
}

const accounts = await loadAccounts();
if (accounts.length === 0) {
  console.error(
    `[google-docs-mcp] No token files found in ${__dirname}. Run \`node setup.js [label]\` to authenticate at least one Google account.`
  );
  process.exit(1);
}

function accountList() {
  return accounts.map((a) => (a.email ? `${a.label} (${a.email})` : a.label)).join(", ");
}

function findAccount(hint) {
  return accounts.find((a) => a.label === hint || a.email === hint);
}

function extractId(idOrUrl) {
  if (!idOrUrl) return idOrUrl;
  const m = String(idOrUrl).match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : idOrUrl;
}

const docOwnerCache = new Map();
const sheetOwnerCache = new Map();

async function pickAccount(idOrUrl, hint, kind) {
  const id = extractId(idOrUrl);
  if (hint) {
    const acct = findAccount(hint);
    if (!acct) {
      throw new Error(`Unknown account "${hint}". Available: ${accountList()}`);
    }
    return { acct, id };
  }
  if (accounts.length === 1) return { acct: accounts[0], id };

  const cache = kind === "doc" ? docOwnerCache : sheetOwnerCache;
  if (cache.has(id)) {
    const acct = findAccount(cache.get(id));
    if (acct) return { acct, id };
  }

  let lastErr;
  for (const acct of accounts) {
    try {
      if (kind === "doc") {
        await acct.docs.documents.get({ documentId: id, fields: "documentId" });
      } else {
        await acct.sheets.spreadsheets.get({ spreadsheetId: id, fields: "spreadsheetId" });
      }
      cache.set(id, acct.label);
      return { acct, id };
    } catch (e) {
      const status = e.code || e.status || e.response?.status;
      if (status === 404 || status === 403 || status === 401) {
        lastErr = e;
        continue;
      }
      throw e;
    }
  }
  throw new Error(
    `No authenticated account can access this ${kind === "doc" ? "document" : "spreadsheet"}. Tried: ${accountList()}.${lastErr ? ` Last error: ${lastErr.message}` : ""}`
  );
}

function pickAccountForCreate(hint) {
  if (hint) {
    const acct = findAccount(hint);
    if (!acct) throw new Error(`Unknown account "${hint}". Available: ${accountList()}`);
    return acct;
  }
  return accounts[0];
}

function paragraphText(paragraph) {
  return (paragraph.elements || [])
    .map((e) => e.textRun?.content || "")
    .join("");
}

function cellPlainText(cell) {
  return (cell.content || [])
    .map((sub) => (sub.paragraph ? paragraphText(sub.paragraph) : ""))
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

function collectTopLevelTables(content) {
  const out = [];
  for (const el of content || []) {
    if (el.table) out.push(el.table);
  }
  return out;
}

function renderDocText(content) {
  let out = "";
  let tableIdx = 0;
  for (const el of content || []) {
    if (el.paragraph) {
      out += paragraphText(el.paragraph);
    } else if (el.table) {
      const myIdx = tableIdx++;
      const rows = el.table.tableRows || [];
      const nCols = rows[0]?.tableCells?.length || 0;
      out += `\n[TABLE ${myIdx} — ${rows.length} rows × ${nCols} cols]\n`;
      rows.forEach((row, rIdx) => {
        const cellTexts = (row.tableCells || []).map((cell) => cellPlainText(cell));
        out += `row ${rIdx}: | ${cellTexts.join(" | ")} |\n`;
      });
      out += `[/TABLE ${myIdx}]\n`;
    }
  }
  return out;
}

function getCellEditRange(cell) {
  if (!cell.content || cell.content.length === 0) return null;
  const start = cell.content[0].startIndex;
  const lastEnd = cell.content[cell.content.length - 1].endIndex;
  return { start, end: lastEnd - 1 };
}

// ---------- T3 helpers: A1 parsing, sheet lookup, chart specs ----------

function colA1ToIndex(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n - 1;
}

function splitSheetAndRange(rangeStr) {
  const bang = rangeStr.lastIndexOf("!");
  if (bang < 0) return { sheetName: null, range: rangeStr };
  let sheetName = rangeStr.slice(0, bang);
  if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
    sheetName = sheetName.slice(1, -1).replace(/''/g, "'");
  }
  return { sheetName, range: rangeStr.slice(bang + 1) };
}

async function getSheetIdMap(acct, spreadsheetId) {
  const r = await acct.sheets.spreadsheets.get({
    spreadsheetId,
    fields: "sheets.properties(sheetId,title)",
  });
  const map = new Map();
  let firstId = null;
  for (const s of r.data.sheets || []) {
    map.set(s.properties.title, s.properties.sheetId);
    if (firstId === null) firstId = s.properties.sheetId;
  }
  return { map, firstId };
}

function resolveSheetId(sheetName, sheetIdMap, firstSheetId) {
  if (!sheetName) {
    if (firstSheetId === null) throw new Error("Spreadsheet has no sheets");
    return firstSheetId;
  }
  if (!sheetIdMap.has(sheetName)) {
    throw new Error(
      `Sheet "${sheetName}" not found. Available: ${[...sheetIdMap.keys()].join(", ")}`
    );
  }
  return sheetIdMap.get(sheetName);
}

function parseGridRange(rangeStr, sheetIdMap, firstSheetId) {
  const { sheetName, range } = splitSheetAndRange(rangeStr);
  const sheetId = resolveSheetId(sheetName, sheetIdMap, firstSheetId);
  if (!range) return { sheetId };
  const [startTok, endTok] = range.split(":");
  const startMatch = startTok.match(/^([A-Z]*)(\d*)$/);
  const endMatch = (endTok || startTok).match(/^([A-Z]*)(\d*)$/);
  if (!startMatch || !endMatch) {
    throw new Error(`Could not parse A1 range "${rangeStr}"`);
  }
  const out = { sheetId };
  if (startMatch[1]) out.startColumnIndex = colA1ToIndex(startMatch[1]);
  if (startMatch[2]) out.startRowIndex = parseInt(startMatch[2], 10) - 1;
  if (endMatch[1]) out.endColumnIndex = colA1ToIndex(endMatch[1]) + 1;
  if (endMatch[2]) out.endRowIndex = parseInt(endMatch[2], 10);
  return out;
}

function parseAnchorCell(cellStr, sheetIdMap, firstSheetId) {
  const { sheetName, range } = splitSheetAndRange(cellStr);
  const sheetId = resolveSheetId(sheetName, sheetIdMap, firstSheetId);
  const m = (range || "").match(/^([A-Z]+)(\d+)$/);
  if (!m) throw new Error(`Anchor "${cellStr}" must be a single cell like "Sheet1!D2"`);
  return {
    sheetId,
    rowIndex: parseInt(m[2], 10) - 1,
    columnIndex: colA1ToIndex(m[1]),
  };
}

function buildChartSpec({ chartType, dataGrid, headerCount, title }) {
  const ct = (chartType || "COLUMN").toUpperCase();
  const domainGrid = {
    ...dataGrid,
    endColumnIndex:
      dataGrid.endColumnIndex !== undefined
        ? Math.min(dataGrid.endColumnIndex, (dataGrid.startColumnIndex ?? 0) + 1)
        : (dataGrid.startColumnIndex ?? 0) + 1,
  };
  const seriesGrid = {
    ...dataGrid,
    startColumnIndex: (dataGrid.startColumnIndex ?? 0) + 1,
  };
  const spec = { title: title || undefined };
  if (ct === "PIE") {
    spec.pieChart = {
      legendPosition: "RIGHT_LEGEND",
      domain: { sourceRange: { sources: [domainGrid] } },
      series: { sourceRange: { sources: [seriesGrid] } },
    };
  } else {
    spec.basicChart = {
      chartType: ct,
      legendPosition: "BOTTOM_LEGEND",
      headerCount: headerCount ?? 1,
      axis: [
        { position: "BOTTOM_AXIS" },
        { position: "LEFT_AXIS" },
      ],
      domains: [{ domain: { sourceRange: { sources: [domainGrid] } } }],
      series: [{ series: { sourceRange: { sources: [seriesGrid] } } }],
    };
  }
  return spec;
}

function buildTextFormat(style) {
  const tf = {};
  if (style.bold !== undefined) tf.bold = !!style.bold;
  if (style.italic !== undefined) tf.italic = !!style.italic;
  if (style.underline !== undefined) tf.underline = !!style.underline;
  if (style.font_size !== undefined) tf.fontSize = style.font_size;
  if (style.font_family) tf.fontFamily = style.font_family;
  if (style.foreground_color) tf.foregroundColor = parseHexColor(style.foreground_color);
  return tf;
}

function parseHexColor(hex) {
  const h = hex.replace("#", "");
  if (h.length !== 6) throw new Error(`Bad hex color "${hex}" (expected #RRGGBB)`);
  return {
    red: parseInt(h.slice(0, 2), 16) / 255,
    green: parseInt(h.slice(2, 4), 16) / 255,
    blue: parseInt(h.slice(4, 6), 16) / 255,
  };
}

const ACCOUNT_PARAM = {
  type: "string",
  description: `Optional Google account to use (label or email). Available accounts: ${accountList()}. If omitted, the MCP auto-detects which account owns the document by trying each in turn (cached after first success). For create operations, the first account is used by default.`,
};

const TOOLS = [
  {
    name: "sheets_read_range",
    description:
      "Read cell values from a Google Sheets spreadsheet. Use this whenever the user asks to read, inspect, or summarize data from a Google Sheet. Accepts either a spreadsheet ID or a full Sheets URL.",
    inputSchema: {
      type: "object",
      properties: {
        spreadsheet_id: {
          type: "string",
          description: "Spreadsheet ID or full URL (https://docs.google.com/spreadsheets/d/<id>/edit).",
        },
        range: {
          type: "string",
          description: "A1 notation, e.g. 'Sheet1!A1:D10' or 'A:A'.",
        },
        account: ACCOUNT_PARAM,
      },
      required: ["spreadsheet_id", "range"],
    },
  },
  {
    name: "sheets_write_range",
    description:
      "Write/overwrite cell values in a Google Sheets spreadsheet. Use this to fill in cells. Values is a 2D array — outer = rows, inner = cells in that row.",
    inputSchema: {
      type: "object",
      properties: {
        spreadsheet_id: { type: "string", description: "Spreadsheet ID or URL." },
        range: { type: "string", description: "A1 notation top-left anchor or full range." },
        values: {
          type: "array",
          description: "2D array. Each inner array is one row.",
          items: { type: "array", items: {} },
        },
        account: ACCOUNT_PARAM,
      },
      required: ["spreadsheet_id", "range", "values"],
    },
  },
  {
    name: "sheets_append_row",
    description:
      "Append one or more rows to the bottom of a table in a Google Sheet. Use this for logging entries, adding records, or extending a list.",
    inputSchema: {
      type: "object",
      properties: {
        spreadsheet_id: { type: "string", description: "Spreadsheet ID or URL." },
        range: {
          type: "string",
          description: "Range describing the table, e.g. 'Sheet1!A:D'. New rows go after the last filled row in this range.",
        },
        values: {
          type: "array",
          description: "2D array of rows to append.",
          items: { type: "array", items: {} },
        },
        account: ACCOUNT_PARAM,
      },
      required: ["spreadsheet_id", "range", "values"],
    },
  },
  {
    name: "sheets_create",
    description: "Create a brand new Google Spreadsheet. Returns the new spreadsheet ID and URL.",
    inputSchema: {
      type: "object",
      properties: {
        title: { type: "string" },
        account: ACCOUNT_PARAM,
      },
      required: ["title"],
    },
  },
  {
    name: "sheets_list_tabs",
    description: "List the tabs (sub-sheets) inside a Google Spreadsheet, with their titles and IDs.",
    inputSchema: {
      type: "object",
      properties: {
        spreadsheet_id: { type: "string", description: "Spreadsheet ID or URL." },
        account: ACCOUNT_PARAM,
      },
      required: ["spreadsheet_id"],
    },
  },
  {
    name: "docs_read",
    description:
      "Read the full text of a Google Doc. Returns paragraph text inline with table summaries. Tables are surfaced as `[TABLE N — R rows × C cols]` blocks with each row prefixed `row K:` and cells separated by `|`. Use the table index N when calling docs_fill_table.",
    inputSchema: {
      type: "object",
      properties: {
        document_id: { type: "string", description: "Document ID or full Docs URL." },
        account: ACCOUNT_PARAM,
      },
      required: ["document_id"],
    },
  },
  {
    name: "docs_create",
    description: "Create a new Google Doc. Returns the new document ID and URL.",
    inputSchema: {
      type: "object",
      properties: {
        title: { type: "string" },
        account: ACCOUNT_PARAM,
      },
      required: ["title"],
    },
  },
  {
    name: "docs_append_text",
    description: "Append text to the end of an existing Google Doc.",
    inputSchema: {
      type: "object",
      properties: {
        document_id: { type: "string", description: "Document ID or URL." },
        text: { type: "string" },
        account: ACCOUNT_PARAM,
      },
      required: ["document_id", "text"],
    },
  },
  {
    name: "docs_replace_text",
    description: "Find-and-replace text inside a Google Doc. Useful for filling templates with placeholders like {{name}}.",
    inputSchema: {
      type: "object",
      properties: {
        document_id: { type: "string", description: "Document ID or URL." },
        find: { type: "string" },
        replace: { type: "string" },
        match_case: { type: "boolean", description: "Default true." },
        account: ACCOUNT_PARAM,
      },
      required: ["document_id", "find", "replace"],
    },
  },
  {
    name: "docs_read_images",
    description:
      "Extract images embedded in a Google Doc. Saves each image to the OS temp directory and returns file paths plus metadata (alt text, byte size). Use the file_path values with the Read tool to view the image content. Only top-level inline images are extracted; positioned/floating images may not be returned.",
    inputSchema: {
      type: "object",
      properties: {
        document_id: { type: "string", description: "Document ID or URL." },
        account: ACCOUNT_PARAM,
      },
      required: ["document_id"],
    },
  },
  {
    name: "docs_fill_table",
    description:
      "Write text into specific cells of a table in a Google Doc. Call docs_read first to find the table_index and confirm row/column counts. `cells` is a list of {row, col, text} entries (0-indexed). `mode` defaults to 'replace' (clears the cell before writing); use 'insert' to prepend without clearing. Only top-level tables are supported (nested tables are not addressable).",
    inputSchema: {
      type: "object",
      properties: {
        document_id: { type: "string", description: "Document ID or URL." },
        table_index: {
          type: "number",
          description: "Which table in the document, 0-indexed in document order. Get this from docs_read output.",
        },
        cells: {
          type: "array",
          description: "Cells to fill. Each entry is {row, col, text} with 0-indexed row/col.",
          items: {
            type: "object",
            properties: {
              row: { type: "number" },
              col: { type: "number" },
              text: { type: "string" },
            },
            required: ["row", "col", "text"],
          },
        },
        mode: {
          type: "string",
          enum: ["replace", "insert"],
          description: "replace (default): clear cell content before writing. insert: prepend text to existing content.",
        },
        account: ACCOUNT_PARAM,
      },
      required: ["document_id", "table_index", "cells"],
    },
  },

  // ===== T3: charts =====
  {
    name: "sheets_list_charts",
    description:
      "[T3] List all charts in a Google Spreadsheet. Returns each chart's id, title, type, and the sheet it lives on. Use the chart_id with sheets_update_chart or sheets_delete_chart.",
    inputSchema: {
      type: "object",
      properties: {
        spreadsheet_id: { type: "string", description: "Spreadsheet ID or URL." },
        account: ACCOUNT_PARAM,
      },
      required: ["spreadsheet_id"],
    },
  },
  {
    name: "sheets_create_chart",
    description:
      "[T3] Create a chart in a Google Spreadsheet from a data range. The first column of the data range is used as the domain (x-axis / category), remaining columns become series. Returns the new chart_id.",
    inputSchema: {
      type: "object",
      properties: {
        spreadsheet_id: { type: "string", description: "Spreadsheet ID or URL." },
        data_range: {
          type: "string",
          description: "A1 data range, e.g. 'Sheet1!A1:C10'. First col = domain, rest = series.",
        },
        chart_type: {
          type: "string",
          enum: ["COLUMN", "BAR", "LINE", "AREA", "PIE", "SCATTER", "COMBO", "STEPPED_AREA"],
          description: "Default COLUMN.",
        },
        title: { type: "string", description: "Chart title (optional)." },
        anchor_cell: {
          type: "string",
          description: "Where to place the chart, e.g. 'Sheet1!E2'. If omitted, anchored at the first cell of the data sheet.",
        },
        header_count: {
          type: "number",
          description: "Number of header rows in the data range. Default 1.",
        },
        width_pixels: { type: "number", description: "Default 600." },
        height_pixels: { type: "number", description: "Default 371." },
        account: ACCOUNT_PARAM,
      },
      required: ["spreadsheet_id", "data_range"],
    },
  },
  {
    name: "sheets_update_chart",
    description:
      "[T3] Update an existing chart's title, type, or data range. Get chart_id from sheets_list_charts. Only fields you pass are changed; others preserved.",
    inputSchema: {
      type: "object",
      properties: {
        spreadsheet_id: { type: "string", description: "Spreadsheet ID or URL." },
        chart_id: { type: "number", description: "Chart ID from sheets_list_charts." },
        title: { type: "string" },
        chart_type: {
          type: "string",
          enum: ["COLUMN", "BAR", "LINE", "AREA", "PIE", "SCATTER", "COMBO", "STEPPED_AREA"],
        },
        data_range: { type: "string", description: "New A1 data range." },
        header_count: { type: "number" },
        account: ACCOUNT_PARAM,
      },
      required: ["spreadsheet_id", "chart_id"],
    },
  },
  {
    name: "sheets_delete_chart",
    description: "[T3] Delete a chart by chart_id. Get chart_id from sheets_list_charts.",
    inputSchema: {
      type: "object",
      properties: {
        spreadsheet_id: { type: "string", description: "Spreadsheet ID or URL." },
        chart_id: { type: "number", description: "Chart ID from sheets_list_charts." },
        account: ACCOUNT_PARAM,
      },
      required: ["spreadsheet_id", "chart_id"],
    },
  },

  // ===== T3: structural sheets ops =====
  {
    name: "sheets_insert_dimension",
    description:
      "[T3] Insert empty rows or columns at a position in a sheet. dimension='ROWS' or 'COLUMNS'. start_index is 0-indexed; 'count' rows/cols are inserted before that position.",
    inputSchema: {
      type: "object",
      properties: {
        spreadsheet_id: { type: "string" },
        sheet_name: { type: "string", description: "Tab name. Defaults to first sheet." },
        dimension: { type: "string", enum: ["ROWS", "COLUMNS"] },
        start_index: { type: "number", description: "0-indexed position to insert before." },
        count: { type: "number", description: "How many rows/cols to insert. Default 1." },
        account: ACCOUNT_PARAM,
      },
      required: ["spreadsheet_id", "dimension", "start_index"],
    },
  },
  {
    name: "sheets_delete_dimension",
    description:
      "[T3] Delete rows or columns from a sheet. dimension='ROWS' or 'COLUMNS'. Removes [start_index, start_index+count) (0-indexed, end-exclusive).",
    inputSchema: {
      type: "object",
      properties: {
        spreadsheet_id: { type: "string" },
        sheet_name: { type: "string" },
        dimension: { type: "string", enum: ["ROWS", "COLUMNS"] },
        start_index: { type: "number" },
        count: { type: "number", description: "Default 1." },
        account: ACCOUNT_PARAM,
      },
      required: ["spreadsheet_id", "dimension", "start_index"],
    },
  },
  {
    name: "sheets_format_range",
    description:
      "[T3] Apply formatting to a cell range: bold/italic/underline, font size/family, text color, background color, horizontal alignment, number format. Pass only the fields you want to change.",
    inputSchema: {
      type: "object",
      properties: {
        spreadsheet_id: { type: "string" },
        range: { type: "string", description: "A1 range, e.g. 'Sheet1!A1:B5'." },
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        underline: { type: "boolean" },
        font_size: { type: "number" },
        font_family: { type: "string" },
        foreground_color: { type: "string", description: "Hex like '#1A73E8'." },
        background_color: { type: "string", description: "Hex like '#FFF3E0'." },
        horizontal_alignment: {
          type: "string",
          enum: ["LEFT", "CENTER", "RIGHT"],
        },
        number_format: {
          type: "string",
          description: "Pattern like '0.00%', '$#,##0.00', 'yyyy-mm-dd'.",
        },
        number_format_type: {
          type: "string",
          enum: ["TEXT", "NUMBER", "PERCENT", "CURRENCY", "DATE", "TIME", "DATE_TIME", "SCIENTIFIC"],
          description: "Defaults to NUMBER if number_format is provided without a type.",
        },
        account: ACCOUNT_PARAM,
      },
      required: ["spreadsheet_id", "range"],
    },
  },
  {
    name: "sheets_merge_cells",
    description:
      "[T3] Merge a range of cells. merge_type='MERGE_ALL' (single merged cell), 'MERGE_COLUMNS' (merge each column), 'MERGE_ROWS' (merge each row). Default MERGE_ALL.",
    inputSchema: {
      type: "object",
      properties: {
        spreadsheet_id: { type: "string" },
        range: { type: "string", description: "A1 range to merge." },
        merge_type: {
          type: "string",
          enum: ["MERGE_ALL", "MERGE_COLUMNS", "MERGE_ROWS"],
        },
        account: ACCOUNT_PARAM,
      },
      required: ["spreadsheet_id", "range"],
    },
  },

  // ===== T3: structural docs ops =====
  {
    name: "docs_insert_table",
    description:
      "[T3] Insert a new empty table into a Google Doc. By default the table is appended to the end of the body. After insertion call docs_read to get the new table's index and use docs_fill_table to populate it.",
    inputSchema: {
      type: "object",
      properties: {
        document_id: { type: "string" },
        rows: { type: "number" },
        columns: { type: "number" },
        index: {
          type: "number",
          description: "Optional 1-indexed character position to insert at. If omitted, table is appended to end of doc.",
        },
        account: ACCOUNT_PARAM,
      },
      required: ["document_id", "rows", "columns"],
    },
  },
  {
    name: "docs_apply_text_style",
    description:
      "[T3] Apply text formatting (bold, italic, underline, font size, color) to text in a Google Doc. Locates targets via find string (all matches) — pass only style fields you want to set.",
    inputSchema: {
      type: "object",
      properties: {
        document_id: { type: "string" },
        find: { type: "string", description: "Text to locate and style. All occurrences are styled." },
        match_case: { type: "boolean", description: "Default true." },
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        underline: { type: "boolean" },
        font_size: { type: "number", description: "In points (e.g. 12)." },
        font_family: { type: "string" },
        foreground_color: { type: "string", description: "Hex like '#1A73E8'." },
        account: ACCOUNT_PARAM,
      },
      required: ["document_id", "find"],
    },
  },
];

const server = new Server(
  { name: "google-docs-mcp", version: "0.4.0" },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: TOOLS }));

server.setRequestHandler(CallToolRequestSchema, async (req) => {
  const { name, arguments: a } = req.params;
  try {
    let result;
    switch (name) {
      case "sheets_read_range": {
        const { acct, id } = await pickAccount(a.spreadsheet_id, a.account, "sheet");
        const r = await acct.sheets.spreadsheets.values.get({
          spreadsheetId: id,
          range: a.range,
        });
        result = r.data.values || [];
        break;
      }
      case "sheets_write_range": {
        const { acct, id } = await pickAccount(a.spreadsheet_id, a.account, "sheet");
        const r = await acct.sheets.spreadsheets.values.update({
          spreadsheetId: id,
          range: a.range,
          valueInputOption: "USER_ENTERED",
          requestBody: { values: a.values },
        });
        result = {
          updatedCells: r.data.updatedCells,
          updatedRange: r.data.updatedRange,
          account: acct.label,
        };
        break;
      }
      case "sheets_append_row": {
        const { acct, id } = await pickAccount(a.spreadsheet_id, a.account, "sheet");
        const r = await acct.sheets.spreadsheets.values.append({
          spreadsheetId: id,
          range: a.range,
          valueInputOption: "USER_ENTERED",
          insertDataOption: "INSERT_ROWS",
          requestBody: { values: a.values },
        });
        result = {
          updatedRange: r.data.updates.updatedRange,
          updatedRows: r.data.updates.updatedRows,
          account: acct.label,
        };
        break;
      }
      case "sheets_create": {
        const acct = pickAccountForCreate(a.account);
        const r = await acct.sheets.spreadsheets.create({
          requestBody: { properties: { title: a.title } },
        });
        result = {
          spreadsheet_id: r.data.spreadsheetId,
          url: r.data.spreadsheetUrl,
          account: acct.label,
        };
        break;
      }
      case "sheets_list_tabs": {
        const { acct, id } = await pickAccount(a.spreadsheet_id, a.account, "sheet");
        const r = await acct.sheets.spreadsheets.get({ spreadsheetId: id });
        result = (r.data.sheets || []).map((s) => ({
          title: s.properties.title,
          sheet_id: s.properties.sheetId,
        }));
        break;
      }
      case "docs_read": {
        const { acct, id } = await pickAccount(a.document_id, a.account, "doc");
        const r = await acct.docs.documents.get({ documentId: id });
        result = renderDocText(r.data.body?.content || []);
        break;
      }
      case "docs_create": {
        const acct = pickAccountForCreate(a.account);
        const r = await acct.docs.documents.create({
          requestBody: { title: a.title },
        });
        result = {
          document_id: r.data.documentId,
          url: `https://docs.google.com/document/d/${r.data.documentId}/edit`,
          account: acct.label,
        };
        break;
      }
      case "docs_append_text": {
        const { acct, id } = await pickAccount(a.document_id, a.account, "doc");
        const doc = await acct.docs.documents.get({ documentId: id });
        const content = doc.data.body?.content || [];
        const last = content[content.length - 1];
        const endIndex = (last?.endIndex ?? 1) - 1;
        await acct.docs.documents.batchUpdate({
          documentId: id,
          requestBody: {
            requests: [
              { insertText: { location: { index: endIndex }, text: a.text } },
            ],
          },
        });
        result = { ok: true, account: acct.label };
        break;
      }
      case "docs_replace_text": {
        const { acct, id } = await pickAccount(a.document_id, a.account, "doc");
        await acct.docs.documents.batchUpdate({
          documentId: id,
          requestBody: {
            requests: [
              {
                replaceAllText: {
                  containsText: {
                    text: a.find,
                    matchCase: a.match_case ?? true,
                  },
                  replaceText: a.replace,
                },
              },
            ],
          },
        });
        result = { ok: true, account: acct.label };
        break;
      }
      case "docs_read_images": {
        const { acct, id } = await pickAccount(a.document_id, a.account, "doc");
        const r = await acct.docs.documents.get({ documentId: id });
        const inlineObjects = r.data.inlineObjects || {};
        const images = [];
        let i = 0;
        for (const [objId, obj] of Object.entries(inlineObjects)) {
          const embed = obj.inlineObjectProperties?.embeddedObject;
          const props = embed?.imageProperties;
          if (!props?.contentUri) {
            i++;
            continue;
          }
          try {
            const res = await fetch(props.contentUri);
            if (!res.ok) {
              images.push({ image_index: i, object_id: objId, error: `HTTP ${res.status}` });
              i++;
              continue;
            }
            const buf = Buffer.from(await res.arrayBuffer());
            const ct = res.headers.get("content-type") || "";
            const ext = ct.includes("png")
              ? "png"
              : ct.includes("jpeg") || ct.includes("jpg")
              ? "jpg"
              : ct.includes("gif")
              ? "gif"
              : ct.includes("webp")
              ? "webp"
              : "bin";
            const filename = path.join(
              os.tmpdir(),
              `gdocs-${id.slice(0, 8)}-${i}.${ext}`
            );
            await fs.writeFile(filename, buf);
            images.push({
              image_index: i,
              object_id: objId,
              alt_text: embed?.title || embed?.description || null,
              file_path: filename,
              bytes: buf.length,
              content_type: ct,
            });
          } catch (e) {
            images.push({ image_index: i, object_id: objId, error: e.message });
          }
          i++;
        }
        result = images;
        break;
      }
      case "docs_fill_table": {
        const { acct, id } = await pickAccount(a.document_id, a.account, "doc");
        const doc = await acct.docs.documents.get({ documentId: id });
        const tables = collectTopLevelTables(doc.data.body?.content || []);
        if (a.table_index < 0 || a.table_index >= tables.length) {
          throw new Error(
            `table_index ${a.table_index} out of range. Document has ${tables.length} top-level table(s).`
          );
        }
        const table = tables[a.table_index];
        const mode = a.mode || "replace";

        const ops = [];
        for (const { row, col, text } of a.cells) {
          if (row < 0 || row >= table.tableRows.length) {
            throw new Error(
              `row ${row} out of range for table ${a.table_index} (${table.tableRows.length} rows).`
            );
          }
          const tableRow = table.tableRows[row];
          if (col < 0 || col >= tableRow.tableCells.length) {
            throw new Error(
              `col ${col} out of range for row ${row} (${tableRow.tableCells.length} cells).`
            );
          }
          const cell = tableRow.tableCells[col];
          const range = getCellEditRange(cell);
          if (!range) continue;
          ops.push({ range, text, mode });
        }

        // Sort descending by start so earlier ops don't shift later ops' indices.
        ops.sort((x, y) => y.range.start - x.range.start);

        const requests = [];
        for (const op of ops) {
          if (op.mode === "replace" && op.range.end > op.range.start) {
            requests.push({
              deleteContentRange: {
                range: { startIndex: op.range.start, endIndex: op.range.end },
              },
            });
          }
          if (op.text) {
            requests.push({
              insertText: {
                location: { index: op.range.start },
                text: op.text,
              },
            });
          }
        }

        if (requests.length === 0) {
          result = { ok: true, cells_filled: 0, account: acct.label };
          break;
        }

        await acct.docs.documents.batchUpdate({
          documentId: id,
          requestBody: { requests },
        });
        result = {
          ok: true,
          cells_filled: a.cells.length,
          requests_sent: requests.length,
          account: acct.label,
        };
        break;
      }
      // ===== T3: charts =====
      case "sheets_list_charts": {
        const { acct, id } = await pickAccount(a.spreadsheet_id, a.account, "sheet");
        const r = await acct.sheets.spreadsheets.get({
          spreadsheetId: id,
          fields: "sheets(properties(sheetId,title),charts(chartId,spec(title,basicChart.chartType,pieChart)))",
        });
        const charts = [];
        for (const sheet of r.data.sheets || []) {
          for (const c of sheet.charts || []) {
            const ct =
              c.spec?.basicChart?.chartType ||
              (c.spec?.pieChart ? "PIE" : "UNKNOWN");
            charts.push({
              chart_id: c.chartId,
              title: c.spec?.title || null,
              chart_type: ct,
              sheet_title: sheet.properties.title,
              sheet_id: sheet.properties.sheetId,
            });
          }
        }
        result = charts;
        break;
      }
      case "sheets_create_chart": {
        const { acct, id } = await pickAccount(a.spreadsheet_id, a.account, "sheet");
        const { map: sheetIdMap, firstId } = await getSheetIdMap(acct, id);
        const dataGrid = parseGridRange(a.data_range, sheetIdMap, firstId);
        const anchor = a.anchor_cell
          ? parseAnchorCell(a.anchor_cell, sheetIdMap, firstId)
          : { sheetId: dataGrid.sheetId, rowIndex: 0, columnIndex: 0 };
        const spec = buildChartSpec({
          chartType: a.chart_type,
          dataGrid,
          headerCount: a.header_count,
          title: a.title,
        });
        const r = await acct.sheets.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: {
            requests: [
              {
                addChart: {
                  chart: {
                    spec,
                    position: {
                      overlayPosition: {
                        anchorCell: anchor,
                        widthPixels: a.width_pixels ?? 600,
                        heightPixels: a.height_pixels ?? 371,
                      },
                    },
                  },
                },
              },
            ],
          },
        });
        const newChart = r.data.replies?.[0]?.addChart?.chart;
        result = {
          ok: true,
          chart_id: newChart?.chartId,
          account: acct.label,
        };
        break;
      }
      case "sheets_update_chart": {
        const { acct, id } = await pickAccount(a.spreadsheet_id, a.account, "sheet");
        const got = await acct.sheets.spreadsheets.get({
          spreadsheetId: id,
          fields: "sheets(properties(sheetId,title),charts(chartId,spec))",
        });
        let existing = null;
        for (const sheet of got.data.sheets || []) {
          for (const c of sheet.charts || []) {
            if (c.chartId === a.chart_id) {
              existing = c;
              break;
            }
          }
          if (existing) break;
        }
        if (!existing) throw new Error(`chart_id ${a.chart_id} not found`);
        const newSpec = JSON.parse(JSON.stringify(existing.spec || {}));
        if (a.title !== undefined) newSpec.title = a.title;
        if (a.chart_type || a.data_range || a.header_count !== undefined) {
          const { map: sheetIdMap, firstId } = await getSheetIdMap(acct, id);
          // Figure out data range: use new one or recover from existing spec.
          let dataGrid;
          if (a.data_range) {
            dataGrid = parseGridRange(a.data_range, sheetIdMap, firstId);
          } else {
            // Recover from existing domains/series sources.
            const oldDomain =
              newSpec.basicChart?.domains?.[0]?.domain?.sourceRange?.sources?.[0] ||
              newSpec.pieChart?.domain?.sourceRange?.sources?.[0];
            const oldSeries =
              newSpec.basicChart?.series?.[0]?.series?.sourceRange?.sources?.[0] ||
              newSpec.pieChart?.series?.sourceRange?.sources?.[0];
            if (!oldDomain) throw new Error("Cannot recover data range from existing chart; pass data_range explicitly");
            dataGrid = {
              sheetId: oldDomain.sheetId,
              startRowIndex: oldDomain.startRowIndex,
              endRowIndex: oldDomain.endRowIndex,
              startColumnIndex: oldDomain.startColumnIndex,
              endColumnIndex: oldSeries?.endColumnIndex ?? oldDomain.endColumnIndex,
            };
          }
          const ct =
            a.chart_type ||
            newSpec.basicChart?.chartType ||
            (newSpec.pieChart ? "PIE" : "COLUMN");
          const headerCount =
            a.header_count !== undefined
              ? a.header_count
              : newSpec.basicChart?.headerCount ?? 1;
          const rebuilt = buildChartSpec({
            chartType: ct,
            dataGrid,
            headerCount,
            title: newSpec.title,
          });
          delete newSpec.basicChart;
          delete newSpec.pieChart;
          Object.assign(newSpec, rebuilt);
        }
        await acct.sheets.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: {
            requests: [
              {
                updateChartSpec: {
                  chartId: a.chart_id,
                  spec: newSpec,
                },
              },
            ],
          },
        });
        result = { ok: true, chart_id: a.chart_id, account: acct.label };
        break;
      }
      case "sheets_delete_chart": {
        const { acct, id } = await pickAccount(a.spreadsheet_id, a.account, "sheet");
        await acct.sheets.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: {
            requests: [{ deleteEmbeddedObject: { objectId: a.chart_id } }],
          },
        });
        result = { ok: true, chart_id: a.chart_id, account: acct.label };
        break;
      }

      // ===== T3: structural sheets ops =====
      case "sheets_insert_dimension": {
        const { acct, id } = await pickAccount(a.spreadsheet_id, a.account, "sheet");
        const { map: sheetIdMap, firstId } = await getSheetIdMap(acct, id);
        const sheetId = a.sheet_name
          ? resolveSheetId(a.sheet_name, sheetIdMap, firstId)
          : firstId;
        const count = a.count ?? 1;
        await acct.sheets.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: {
            requests: [
              {
                insertDimension: {
                  range: {
                    sheetId,
                    dimension: a.dimension,
                    startIndex: a.start_index,
                    endIndex: a.start_index + count,
                  },
                  inheritFromBefore: a.start_index > 0,
                },
              },
            ],
          },
        });
        result = { ok: true, inserted: count, account: acct.label };
        break;
      }
      case "sheets_delete_dimension": {
        const { acct, id } = await pickAccount(a.spreadsheet_id, a.account, "sheet");
        const { map: sheetIdMap, firstId } = await getSheetIdMap(acct, id);
        const sheetId = a.sheet_name
          ? resolveSheetId(a.sheet_name, sheetIdMap, firstId)
          : firstId;
        const count = a.count ?? 1;
        await acct.sheets.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: {
            requests: [
              {
                deleteDimension: {
                  range: {
                    sheetId,
                    dimension: a.dimension,
                    startIndex: a.start_index,
                    endIndex: a.start_index + count,
                  },
                },
              },
            ],
          },
        });
        result = { ok: true, deleted: count, account: acct.label };
        break;
      }
      case "sheets_format_range": {
        const { acct, id } = await pickAccount(a.spreadsheet_id, a.account, "sheet");
        const { map: sheetIdMap, firstId } = await getSheetIdMap(acct, id);
        const grid = parseGridRange(a.range, sheetIdMap, firstId);
        const userEnteredFormat = {};
        const fields = [];
        const tf = buildTextFormat(a);
        if (Object.keys(tf).length) {
          userEnteredFormat.textFormat = tf;
          for (const k of Object.keys(tf)) fields.push(`textFormat.${k}`);
        }
        if (a.background_color) {
          userEnteredFormat.backgroundColor = parseHexColor(a.background_color);
          fields.push("backgroundColor");
        }
        if (a.horizontal_alignment) {
          userEnteredFormat.horizontalAlignment = a.horizontal_alignment;
          fields.push("horizontalAlignment");
        }
        if (a.number_format) {
          userEnteredFormat.numberFormat = {
            type: a.number_format_type || "NUMBER",
            pattern: a.number_format,
          };
          fields.push("numberFormat");
        }
        if (fields.length === 0) {
          result = { ok: true, fields_changed: 0, account: acct.label };
          break;
        }
        await acct.sheets.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: {
            requests: [
              {
                repeatCell: {
                  range: grid,
                  cell: { userEnteredFormat },
                  fields: `userEnteredFormat(${fields.join(",")})`,
                },
              },
            ],
          },
        });
        result = { ok: true, fields_changed: fields.length, account: acct.label };
        break;
      }
      case "sheets_merge_cells": {
        const { acct, id } = await pickAccount(a.spreadsheet_id, a.account, "sheet");
        const { map: sheetIdMap, firstId } = await getSheetIdMap(acct, id);
        const grid = parseGridRange(a.range, sheetIdMap, firstId);
        await acct.sheets.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: {
            requests: [
              {
                mergeCells: {
                  range: grid,
                  mergeType: a.merge_type || "MERGE_ALL",
                },
              },
            ],
          },
        });
        result = { ok: true, account: acct.label };
        break;
      }

      // ===== T3: structural docs ops =====
      case "docs_insert_table": {
        const { acct, id } = await pickAccount(a.document_id, a.account, "doc");
        let location;
        if (a.index !== undefined) {
          location = { location: { index: a.index } };
        } else {
          const doc = await acct.docs.documents.get({ documentId: id });
          const content = doc.data.body?.content || [];
          const last = content[content.length - 1];
          // endOfSegmentLocation appends; insertTable supports it.
          location = { endOfSegmentLocation: {} };
          // Fallback: if endOfSegmentLocation isn't accepted in some contexts, use index.
          if (!last) location = { location: { index: 1 } };
        }
        await acct.docs.documents.batchUpdate({
          documentId: id,
          requestBody: {
            requests: [
              {
                insertTable: {
                  rows: a.rows,
                  columns: a.columns,
                  ...location,
                },
              },
            ],
          },
        });
        result = {
          ok: true,
          rows: a.rows,
          columns: a.columns,
          account: acct.label,
        };
        break;
      }
      case "docs_apply_text_style": {
        const { acct, id } = await pickAccount(a.document_id, a.account, "doc");
        const doc = await acct.docs.documents.get({ documentId: id });
        const body = doc.data.body?.content || [];
        const matchCase = a.match_case ?? true;
        const needle = matchCase ? a.find : a.find.toLowerCase();
        const ranges = [];
        function walk(content) {
          for (const el of content || []) {
            if (el.paragraph) {
              for (const e of el.paragraph.elements || []) {
                if (e.textRun?.content) {
                  const hay = matchCase ? e.textRun.content : e.textRun.content.toLowerCase();
                  let from = 0;
                  while (true) {
                    const at = hay.indexOf(needle, from);
                    if (at < 0) break;
                    ranges.push({
                      startIndex: e.startIndex + at,
                      endIndex: e.startIndex + at + needle.length,
                    });
                    from = at + needle.length;
                  }
                }
              }
            } else if (el.table) {
              for (const row of el.table.tableRows || []) {
                for (const cell of row.tableCells || []) {
                  walk(cell.content);
                }
              }
            }
          }
        }
        walk(body);
        if (ranges.length === 0) {
          result = { ok: true, ranges_styled: 0, account: acct.label };
          break;
        }
        const textStyle = {};
        const fields = [];
        if (a.bold !== undefined) { textStyle.bold = a.bold; fields.push("bold"); }
        if (a.italic !== undefined) { textStyle.italic = a.italic; fields.push("italic"); }
        if (a.underline !== undefined) { textStyle.underline = a.underline; fields.push("underline"); }
        if (a.font_size !== undefined) {
          textStyle.fontSize = { magnitude: a.font_size, unit: "PT" };
          fields.push("fontSize");
        }
        if (a.font_family) {
          textStyle.weightedFontFamily = { fontFamily: a.font_family };
          fields.push("weightedFontFamily");
        }
        if (a.foreground_color) {
          textStyle.foregroundColor = { color: { rgbColor: parseHexColor(a.foreground_color) } };
          fields.push("foregroundColor");
        }
        if (fields.length === 0) {
          result = { ok: true, ranges_styled: 0, note: "no style fields provided", account: acct.label };
          break;
        }
        const requests = ranges.map((r) => ({
          updateTextStyle: {
            range: r,
            textStyle,
            fields: fields.join(","),
          },
        }));
        await acct.docs.documents.batchUpdate({
          documentId: id,
          requestBody: { requests },
        });
        result = { ok: true, ranges_styled: ranges.length, account: acct.label };
        break;
      }

      default:
        throw new Error(`Unknown tool: ${name}`);
    }
    return {
      content: [
        {
          type: "text",
          text: typeof result === "string" ? result : JSON.stringify(result, null, 2),
        },
      ],
    };
  } catch (err) {
    return {
      content: [{ type: "text", text: `Error: ${err.message || String(err)}` }],
      isError: true,
    };
  }
});

await server.connect(new StdioServerTransport());
