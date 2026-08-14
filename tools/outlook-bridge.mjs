import { createServer } from "node:http";
import { spawn } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const host = "127.0.0.1";
const port = Number(process.env.OUTLOOK_BRIDGE_PORT || 47831);
const script = join(dirname(fileURLToPath(import.meta.url)), "read-outlook-calendar.ps1");
const scriptSource = readFileSync(script, "utf8").replace(/^\uFEFF/, "");
const allowedOrigins = new Set([
  "http://localhost:3000",
  "http://127.0.0.1:3000",
  "https://work-organizer-app.nunomoreira1993.chatgpt.site",
  ...(process.env.OUTLOOK_BRIDGE_ALLOWED_ORIGINS ?? "")
    .split(",")
    .map((origin) => origin.trim())
    .filter(Boolean),
]);

function isAllowedOrigin(origin) {
  if (allowedOrigins.has(origin)) return true;

  try {
    const { hostname, protocol } = new URL(origin);
    const isHttp = protocol === "http:" || protocol === "https:";
    return isHttp && ["localhost", "127.0.0.1", "[::1]", "terminal.local"].includes(hostname);
  } catch {
    return false;
  }
}

function send(response, status, body, origin) {
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    ...(isAllowedOrigin(origin) ? {
      "Access-Control-Allow-Origin": origin,
      "Access-Control-Allow-Private-Network": "true",
      Vary: "Origin",
    } : {}),
  });
  response.end(JSON.stringify(body));
}

function readCalendar(start, days) {
  return new Promise((resolve, reject) => {
    // Execute the checked-in source in memory. A downloaded repository can carry
    // Zone.Identifier metadata, and an enforced PowerShell policy may then reject
    // the .ps1 file even when the process requests ExecutionPolicy Bypass.
    const sourceBase64 = Buffer.from(scriptSource, "utf8").toString("base64");
    const command = [
      "[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)",
      `& ([scriptblock]::Create([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${sourceBase64}')))) -StartDate '${start}' -Days ${days}`,
    ].join("; ");
    const encodedCommand = Buffer.from(command, "utf16le").toString("base64");
    const child = spawn("powershell.exe", ["-NoLogo", "-NoProfile", "-NonInteractive", "-EncodedCommand", encodedCommand], { windowsHide: true });
    let stdout = "";
    let stderr = "";
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code !== 0) return reject(new Error(stderr.trim() || `PowerShell terminou com código ${code}.`));
      try { resolve(stdout.trim() ? JSON.parse(stdout) : []); }
      catch { reject(new Error("O Outlook devolveu dados num formato inesperado.")); }
    });
  });
}

const server = createServer(async (request, response) => {
  const origin = request.headers.origin || "";
  if (request.method === "OPTIONS") {
    response.writeHead(204, {
      ...(isAllowedOrigin(origin) ? { "Access-Control-Allow-Origin": origin } : {}),
      "Access-Control-Allow-Methods": "GET, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
      "Access-Control-Allow-Private-Network": "true",
      Vary: "Origin",
    });
    return response.end();
  }
  if (origin && !isAllowedOrigin(origin)) return send(response, 403, { ok: false, error: "Origem não autorizada." }, "");
  const url = new URL(request.url || "/", `http://${host}:${port}`);
  if (request.method === "GET" && url.pathname === "/health") {
    return send(response, 200, { ok: true }, origin);
  }
  if (request.method !== "GET" || url.pathname !== "/calendar") return send(response, 404, { ok: false, error: "Endpoint inexistente." }, origin);
  const days = Math.min(90, Math.max(1, Number(url.searchParams.get("days")) || 35));
  const requestedStart = url.searchParams.get("start");
  const start = /^\d{4}-\d{2}-\d{2}$/.test(requestedStart ?? "")
    ? requestedStart
    : new Date().toISOString().slice(0, 10);
  try {
    const events = await readCalendar(start, days);
    send(response, 200, { ok: true, events: Array.isArray(events) ? events : [events] }, origin);
  } catch (error) {
    send(response, 500, { ok: false, error: error instanceof Error ? error.message : "Falha ao ler o Outlook." }, origin);
  }
});

server.listen(port, host, () => {
  process.stdout.write(`Ponte Outlook ativa em http://${host}:${port}\nMantém esta janela aberta enquanto usas o Work Organizer.\n`);
});
