# Claude-to-Safari-MCP
A zero-dependency MCP server that lets Claude drive the Safari you already have open on macOS. One Node.js file, no npm install, no headless browser. 

# claude-safari-mcp

> Give Claude hands and eyes inside Safari. Zero dependencies, one file.

`claude-safari-mcp` is a [Model Context Protocol](https://modelcontextprotocol.io) server that lets Claude (Desktop, Code, or any MCP client) drive Safari on macOS — open URLs, read page text, run JavaScript, manage tabs, capture screenshots. The whole server is a single self-contained Node.js file. No `npm install`. No build step.

## Why

The web is still where most work happens. Browser-use frameworks like Playwright and Puppeteer launch their own headless Chromium, which means:

- You log in again. And again.
- Your sessions, cookies, and 2FA prompts live somewhere your agent can't see.
- It costs hundreds of MB of RAM per run.

This server piggybacks on the Safari you already have open, with the cookies and sessions you already trust. Claude works in *your* browser, not a sandboxed copy of it.

## Tools exposed

| Tool | What it does |
| --- | --- |
| `open_url` | Navigate the current tab to a URL |
| `new_tab` | Open a URL in a new tab |
| `get_current_tab` | Read URL + title of the active tab |
| `list_tabs` | List every tab across every window |
| `activate_tab` | Switch to a specific tab |
| `close_current_tab` | Close the active tab |
| `reload_page` | Reload the active tab |
| `go_back` / `go_forward` | History navigation |
| `get_page_text` | `document.body.innerText` of the active tab |
| `get_page_html` | Full outer HTML of the active tab |
| `run_javascript` | Execute arbitrary JS in the active tab and return the result |
| `search` | One-shot web search (google / duckduckgo / bing / brave) |
| `take_screenshot` | PNG of the front Safari window to a path you choose |

## Install

```bash
curl -O https://raw.githubusercontent.com/<your-username>/claude-safari-mcp/main/safari-mcp.js
chmod +x safari-mcp.js
```

That's it. Node 18+ is the only requirement — no `npm install`.

### One-time macOS setup

1. **Enable JS-from-AppleEvents** (only if you want `run_javascript`):
   Safari ▸ Settings ▸ Advanced ▸ "Show Develop menu in menu bar"
   then Develop ▸ "Allow JavaScript from Apple Events".
2. **Automation permission**: the first call will trigger a system prompt asking your terminal (or the Claude app) to control Safari. Approve it. You can revisit at System Settings ▸ Privacy & Security ▸ Automation.

## Wire it up to Claude

### Claude Desktop

Edit `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "safari": {
      "command": "node",
      "args": ["/absolute/path/to/safari-mcp.js"]
    }
  }
}
```

Restart Claude Desktop. You'll see "safari" in the tools menu.

### Claude Code

```bash
claude mcp add safari -- node /absolute/path/to/safari-mcp.js
```

### Any other MCP client

Point it at `node safari-mcp.js` over stdio. Protocol version `2024-11-05`.

## Try it

In Claude:

> *Open hacker news in Safari and tell me the top three story titles.*

> *Take a screenshot of the current tab and save it to ~/Desktop/page.png.*

> *In the active tab, run `document.querySelectorAll("a").length` and tell me how many links are on the page.*

## How it works

Everything routes through Apple's `osascript` — AppleScript for tab/window control, JXA (JavaScript for Automation) for richer reads like `list_tabs`. JSON-RPC 2.0 over stdio is implemented by hand, which is why there are no dependencies.

The whole protocol surface fits in ~30 lines at the bottom of [safari-mcp.js](./safari-mcp.js): `initialize`, `tools/list`, `tools/call`, plus the `notifications/initialized` ack. If you want to add a tool, drop a new entry in the `TOOLS` array and you're done.

## Limitations

- macOS only (Safari + osascript). For cross-platform browser control look at [puppeteer-mcp](https://github.com/modelcontextprotocol/servers) instead.
- `run_javascript` is gated behind the Develop-menu setting above. Apple requires explicit opt-in.
- The server controls *your real Safari*. Be thoughtful about what you let an agent click.

## Roadmap ideas

- [ ] Reader-mode text extraction (cleaner than `innerText`)
- [ ] Bookmark + Reading List read APIs
- [ ] Per-tab screenshot (not just front window)
- [ ] Form-fill helpers built on top of `run_javascript`
- [ ] Resource subscriptions: stream URL changes back to Claude

PRs welcome. The whole codebase is one file — easy to grok, easy to fork.

## License

MIT.
