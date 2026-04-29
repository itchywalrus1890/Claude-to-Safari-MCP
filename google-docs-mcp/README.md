# google-docs-mcp

A local [Model Context Protocol](https://modelcontextprotocol.io) server that lets Claude (Desktop, Code, or any MCP client) read, edit, and structurally manipulate **Google Docs** and **Google Sheets** — including charts.

Multi-account aware. Auto-detects which Google account owns the document you ask about, so you can mix personal/work/school accounts in the same MCP.

## Why

Most "Google Docs MCP" servers stop at reading text. This one goes further:

- Read paragraphs *and* tables (with index labels) in a single call
- Write into specific table cells by `(row, col)` coordinates
- Pull inline images out of a doc and onto disk so Claude can see them
- Create / update / delete Sheets charts (column, bar, line, pie, scatter, combo, area, stepped area)
- Insert / delete rows and columns, format ranges, merge cells
- Insert tables and apply text styles inside Docs

## Tools

### T1 — read
| Tool | Purpose |
| --- | --- |
| `sheets_read_range` | Read cell values from an A1 range |
| `sheets_list_tabs` | List tabs in a spreadsheet |
| `docs_read` | Read a doc; tables surfaced as `[TABLE N — R rows × C cols]` blocks |
| `docs_read_images` | Pull inline images to OS temp dir (returns file paths) |

### T2 — basic edit
| Tool | Purpose |
| --- | --- |
| `sheets_write_range` | Write 2D values into a range |
| `sheets_append_row` | Append rows to bottom of a table |
| `sheets_create` | Create a new spreadsheet |
| `docs_create` | Create a new doc |
| `docs_append_text` | Append text to end of a doc |
| `docs_replace_text` | Find-and-replace inside a doc |
| `docs_fill_table` | Write into specific table cells by `{row, col, text}` |

### T3 — charts + structural
| Tool | Purpose |
| --- | --- |
| `sheets_list_charts` | Enumerate charts (id, title, type, sheet) |
| `sheets_create_chart` | Create a chart from a data range (8 chart types) |
| `sheets_update_chart` | Partial update: title, type, data range |
| `sheets_delete_chart` | Delete by chart id |
| `sheets_insert_dimension` | Insert rows or columns |
| `sheets_delete_dimension` | Delete rows or columns |
| `sheets_format_range` | Bold/italic/underline, font size/family, fg/bg color, alignment, number format |
| `sheets_merge_cells` | MERGE_ALL / MERGE_COLUMNS / MERGE_ROWS |
| `docs_insert_table` | Insert an empty NxM table |
| `docs_apply_text_style` | Style all matches of a text string |

## Install

```bash
git clone <this-repo>
cd <repo>/google-docs-mcp
npm install
```

### Get OAuth credentials

You need your own OAuth Desktop client. Free, takes a few minutes:

1. Go to <https://console.cloud.google.com/>
2. Create (or pick) a project
3. Enable both APIs:
   - Google Docs API
   - Google Sheets API
4. Go to **APIs & Services → Credentials → Create Credentials → OAuth client ID**
5. Application type: **Desktop**
6. Download the JSON, save it as `credentials.json` in this directory

### Authenticate

```bash
node setup.js               # default account
node setup.js work          # optional: a second account, labeled "work"
node setup.js school        # optional: a third
```

Each run opens a browser for Google sign-in. Tokens are saved as `token.<label>.json` (or `token.json` for the default unlabeled run).

### Register with Claude

**Claude Code:**
```bash
claude mcp add google-docs -- node /absolute/path/to/google-docs-mcp/index.js
```

**Claude Desktop** — edit `~/Library/Application Support/Claude/claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "google-docs": {
      "command": "node",
      "args": ["/absolute/path/to/google-docs-mcp/index.js"]
    }
  }
}
```

Restart Claude after registering.

## Multi-account

If you have more than one Google account authenticated, the MCP probes each in turn until one can open the document you asked about, then caches that mapping. You can override with the optional `account` argument on any tool call (label like `"work"` or full email).

## T3 examples (in Claude)

> *Create a column chart in this sheet from `Sheet1!A1:C20`, titled "Q1 revenue".*

> *Add 5 blank rows at row 10 in the "Forecast" tab.*

> *Make the header row of `Sheet1!A1:F1` bold with a light blue background.*

> *Style every occurrence of "DRAFT" in this doc as red and bold.*

## Security notes

- `credentials.json` and any `token.*.json` file grant API access to the Google account they represent. They are excluded from git via `.gitignore`. **Never commit them.**
- The OAuth scopes requested are `spreadsheets` (full read/write) and `documents` (full read/write). Charts and structural operations both fall under those scopes — no additional consent needed after the initial setup.
- Each MCP code change requires a Claude Code/Desktop restart for new tool schemas to load.

## Limitations

- Only top-level tables in Docs are addressable (nested tables aren't exposed).
- `docs_read_images` only handles inline images, not positioned/floating ones.
- `sheets_update_chart` rebuilds the chart spec when changing `chart_type` or `data_range`; some advanced styling options on the original chart will be reset to defaults.

## License

MIT.
