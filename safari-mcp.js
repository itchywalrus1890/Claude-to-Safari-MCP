#!/usr/bin/env node
/**
 * claude-safari-mcp
 * A zero-dependency Model Context Protocol server that lets Claude drive Safari on macOS.
 *
 * Speaks MCP (JSON-RPC 2.0 over stdio) and controls Safari through `osascript`.
 *
 * Requirements:
 *   - macOS with Safari installed
 *   - Node.js 18+
 *   - For run_javascript: enable Safari ▸ Develop ▸ "Allow JavaScript from Apple Events"
 *   - On first call, macOS will prompt to grant your terminal/Claude app
 *     "Automation" permission for Safari (System Settings ▸ Privacy & Security ▸ Automation).
 *
 * Install:
 *   chmod +x safari-mcp.js
 *
 * Register with Claude Desktop (~/Library/Application Support/Claude/claude_desktop_config.json):
 *   {
 *     "mcpServers": {
 *       "safari": { "command": "node", "args": ["/absolute/path/to/safari-mcp.js"] }
 *     }
 *   }
 *
 * Register with Claude Code:
 *   claude mcp add safari -- node /absolute/path/to/safari-mcp.js
 */

const { execFile } = require("child_process");
const readline = require("readline");

const SERVER_NAME = "claude-safari-mcp";
const SERVER_VERSION = "0.1.0";
const PROTOCOL_VERSION = "2024-11-05";

// ---------- AppleScript helpers ----------

function runOsa(language, script) {
  return new Promise((resolve, reject) => {
    const args = ["-l", language, "-e", script];
    execFile("osascript", args, { maxBuffer: 32 * 1024 * 1024 }, (err, stdout, stderr) => {
      if (err) {
        const msg = (stderr && stderr.toString().trim()) || err.message;
        return reject(new Error(msg));
      }
      resolve(stdout.toString());
    });
  });
}

const applescript = (s) => runOsa("AppleScript", s);
const jxa = (s) => runOsa("JavaScript", s);

function escapeForAppleScript(str) {
  return String(str).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

// ---------- Safari tool implementations ----------

async function activateSafari() {
  await applescript('tell application "Safari" to activate');
}

async function openUrl({ url }) {
  if (!url) throw new Error("url is required");
  const safe = escapeForAppleScript(url);
  await applescript(`tell application "Safari"
    activate
    if (count of windows) = 0 then
      make new document with properties {URL:"${safe}"}
    else
      set URL of current tab of front window to "${safe}"
    end if
  end tell`);
  return { ok: true, url };
}

async function newTab({ url }) {
  const target = url ? escapeForAppleScript(url) : "about:blank";
  await applescript(`tell application "Safari"
    activate
    if (count of windows) = 0 then
      make new document with properties {URL:"${target}"}
    else
      tell front window to set current tab to (make new tab with properties {URL:"${target}"})
    end if
  end tell`);
  return { ok: true, url: url || "about:blank" };
}

async function getCurrentTab() {
  const out = await jxa(`
    const safari = Application("Safari");
    if (safari.windows.length === 0) { JSON.stringify({error: "no windows open"}); }
    else {
      const t = safari.windows[0].currentTab;
      JSON.stringify({ url: t.url(), title: t.name() });
    }
  `);
  return JSON.parse(out.trim());
}

async function listTabs() {
  const out = await jxa(`
    const safari = Application("Safari");
    const result = [];
    safari.windows().forEach((win, wi) => {
      win.tabs().forEach((tab, ti) => {
        result.push({
          window: wi,
          tab: ti + 1,
          url: tab.url(),
          title: tab.name(),
          active: tab.index() === win.currentTab.index()
        });
      });
    });
    JSON.stringify(result);
  `);
  return JSON.parse(out.trim());
}

async function activateTab({ window = 0, tab }) {
  if (!tab) throw new Error("tab (1-indexed) is required");
  await jxa(`
    const safari = Application("Safari");
    safari.activate();
    const win = safari.windows[${Number(window)}];
    win.currentTab = win.tabs[${Number(tab) - 1}];
  `);
  return { ok: true };
}

async function closeCurrentTab() {
  await applescript(`tell application "Safari" to tell front window to close current tab`);
  return { ok: true };
}

async function reloadPage() {
  await applescript(`tell application "Safari" to tell front window to set URL of current tab to (URL of current tab)`);
  return { ok: true };
}

async function goBack() {
  await runJavaScript({ code: "history.back();" });
  return { ok: true };
}

async function goForward() {
  await runJavaScript({ code: "history.forward();" });
  return { ok: true };
}

async function getPageText() {
  return await runJavaScript({ code: "document.body.innerText" });
}

async function getPageHtml() {
  return await runJavaScript({ code: "document.documentElement.outerHTML" });
}

async function runJavaScript({ code }) {
  if (!code) throw new Error("code is required");
  const safe = escapeForAppleScript(code);
  const out = await applescript(`tell application "Safari"
    set theResult to do JavaScript "${safe}" in current tab of front window
    return theResult as string
  end tell`);
  return out.trimEnd();
}

async function search({ query, engine = "google" }) {
  if (!query) throw new Error("query is required");
  const engines = {
    google: "https://www.google.com/search?q=",
    duckduckgo: "https://duckduckgo.com/?q=",
    bing: "https://www.bing.com/search?q=",
    brave: "https://search.brave.com/search?q=",
  };
  const base = engines[engine.toLowerCase()] || engines.google;
  return await openUrl({ url: base + encodeURIComponent(query) });
}

async function takeScreenshot({ path }) {
  if (!path) throw new Error("path is required (where to save the .png)");
  await activateSafari();
  await new Promise((r) => setTimeout(r, 250));
  await new Promise((resolve, reject) => {
    execFile("screencapture", ["-l", "$(osascript -e 'tell app \"Safari\" to id of front window')", path], {}, () => resolve());
  });
  // Fallback simpler call: capture frontmost window
  await new Promise((resolve, reject) => {
    execFile("/bin/bash", ["-c", `WID=$(osascript -e 'tell application "Safari" to id of front window'); screencapture -l "$WID" "${path.replace(/"/g, '\\"')}"`], (err) => {
      if (err) reject(err);
      else resolve();
    });
  });
  return { ok: true, path };
}

// ---------- MCP tool registry ----------

const TOOLS = [
  {
    name: "open_url",
    description: "Open a URL in the current Safari tab (creates a window if none exist).",
    inputSchema: { type: "object", properties: { url: { type: "string" } }, required: ["url"] },
    handler: openUrl,
  },
  {
    name: "new_tab",
    description: "Open a URL in a new Safari tab. URL is optional; defaults to about:blank.",
    inputSchema: { type: "object", properties: { url: { type: "string" } } },
    handler: newTab,
  },
  {
    name: "get_current_tab",
    description: "Return the URL and title of the active Safari tab.",
    inputSchema: { type: "object", properties: {} },
    handler: getCurrentTab,
  },
  {
    name: "list_tabs",
    description: "List every open Safari tab across all windows.",
    inputSchema: { type: "object", properties: {} },
    handler: listTabs,
  },
  {
    name: "activate_tab",
    description: "Switch to a specific tab. tab is 1-indexed within the given window (default: window 0).",
    inputSchema: {
      type: "object",
      properties: { window: { type: "number" }, tab: { type: "number" } },
      required: ["tab"],
    },
    handler: activateTab,
  },
  {
    name: "close_current_tab",
    description: "Close the currently active Safari tab.",
    inputSchema: { type: "object", properties: {} },
    handler: closeCurrentTab,
  },
  {
    name: "reload_page",
    description: "Reload the active tab.",
    inputSchema: { type: "object", properties: {} },
    handler: reloadPage,
  },
  {
    name: "go_back",
    description: "Navigate the active tab back one step in history.",
    inputSchema: { type: "object", properties: {} },
    handler: goBack,
  },
  {
    name: "go_forward",
    description: "Navigate the active tab forward one step in history.",
    inputSchema: { type: "object", properties: {} },
    handler: goForward,
  },
  {
    name: "get_page_text",
    description: "Return document.body.innerText of the active tab.",
    inputSchema: { type: "object", properties: {} },
    handler: getPageText,
  },
  {
    name: "get_page_html",
    description: "Return the full outer HTML of the active tab.",
    inputSchema: { type: "object", properties: {} },
    handler: getPageHtml,
  },
  {
    name: "run_javascript",
    description: "Execute JavaScript in the active tab and return its result. Requires Safari ▸ Develop ▸ 'Allow JavaScript from Apple Events'.",
    inputSchema: { type: "object", properties: { code: { type: "string" } }, required: ["code"] },
    handler: runJavaScript,
  },
  {
    name: "search",
    description: "Run a web search in the active tab. engine: google | duckduckgo | bing | brave (default google).",
    inputSchema: {
      type: "object",
      properties: { query: { type: "string" }, engine: { type: "string" } },
      required: ["query"],
    },
    handler: search,
  },
  {
    name: "take_screenshot",
    description: "Capture the front Safari window to a PNG at the given absolute path.",
    inputSchema: { type: "object", properties: { path: { type: "string" } }, required: ["path"] },
    handler: takeScreenshot,
  },
];

// ---------- JSON-RPC plumbing ----------

function send(msg) {
  process.stdout.write(JSON.stringify(msg) + "\n");
}

function reply(id, result) {
  send({ jsonrpc: "2.0", id, result });
}

function replyError(id, code, message) {
  send({ jsonrpc: "2.0", id, error: { code, message } });
}

async function handle(msg) {
  const { id, method, params } = msg;

  try {
    switch (method) {
      case "initialize":
        return reply(id, {
          protocolVersion: PROTOCOL_VERSION,
          capabilities: { tools: {} },
          serverInfo: { name: SERVER_NAME, version: SERVER_VERSION },
        });

      case "notifications/initialized":
        return; // notification, no reply

      case "ping":
        return reply(id, {});

      case "tools/list":
        return reply(id, {
          tools: TOOLS.map(({ name, description, inputSchema }) => ({ name, description, inputSchema })),
        });

      case "tools/call": {
        const { name, arguments: args = {} } = params || {};
        const tool = TOOLS.find((t) => t.name === name);
        if (!tool) return replyError(id, -32601, `Unknown tool: ${name}`);
        try {
          const result = await tool.handler(args);
          const text = typeof result === "string" ? result : JSON.stringify(result, null, 2);
          return reply(id, { content: [{ type: "text", text }] });
        } catch (err) {
          return reply(id, {
            isError: true,
            content: [{ type: "text", text: `Error in ${name}: ${err.message}` }],
          });
        }
      }

      default:
        if (id !== undefined) replyError(id, -32601, `Method not found: ${method}`);
    }
  } catch (err) {
    if (id !== undefined) replyError(id, -32603, err.message);
  }
}

const rl = readline.createInterface({ input: process.stdin });
rl.on("line", (line) => {
  const trimmed = line.trim();
  if (!trimmed) return;
  let msg;
  try {
    msg = JSON.parse(trimmed);
  } catch {
    return; // ignore malformed lines
  }
  handle(msg);
});

process.on("SIGINT", () => process.exit(0));
process.on("SIGTERM", () => process.exit(0));
