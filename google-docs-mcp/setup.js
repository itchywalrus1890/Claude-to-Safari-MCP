#!/usr/bin/env node
import { authenticate } from "@google-cloud/local-auth";
import { google } from "googleapis";
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CREDENTIALS_PATH = path.join(__dirname, "credentials.json");
const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/documents",
  "https://www.googleapis.com/auth/userinfo.email",
];

try {
  await fs.access(CREDENTIALS_PATH);
} catch {
  console.error(
    `Missing credentials.json. Download it from Google Cloud Console (OAuth Desktop client) and place it at:\n  ${CREDENTIALS_PATH}`
  );
  process.exit(1);
}

const labelArg = process.argv[2];

console.log("Opening browser for Google sign-in…");
console.log("Sign in with the Google account you want this token to represent.");
const client = await authenticate({ scopes: SCOPES, keyfilePath: CREDENTIALS_PATH });

const oauth2 = google.oauth2({ version: "v2", auth: client });
const { data: info } = await oauth2.userinfo.get();
const email = info.email || "";
const localPart = email.split("@")[0].replace(/[^a-zA-Z0-9_-]/g, "_") || "account";
const label = labelArg ? labelArg.replace(/[^a-zA-Z0-9_-]/g, "_") : localPart;

const keys = JSON.parse(await fs.readFile(CREDENTIALS_PATH, "utf8"));
const key = keys.installed || keys.web;
const payload = JSON.stringify(
  {
    type: "authorized_user",
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
    email,
  },
  null,
  2
);

const filename = `token.${label}.json`;
const tokenPath = path.join(__dirname, filename);
await fs.writeFile(tokenPath, payload);
console.log(`\nSaved token for ${email} as label "${label}" → ${filename}`);
console.log(`Run \`node setup.js <label>\` again with a different Google account to add more.`);
