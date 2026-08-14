import { createReadStream, existsSync, mkdirSync, statSync } from "node:fs";
import { createServer } from "node:http";
import { DatabaseSync } from "node:sqlite";
import { dirname, extname, join, resolve, sep } from "node:path";
import { Readable } from "node:stream";
import { fileURLToPath, pathToFileURL } from "node:url";

process.env.WORK_ORGANIZER_LOCAL = "true";

const projectRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const clientRoot = join(projectRoot, "dist", "client");
const serverEntry = join(projectRoot, "dist", "server", "index.js");
const databaseDirectory = join(projectRoot, "data");
const databasePath = join(databaseDirectory, "work-organizer.sqlite");
const backupDirectory = join(projectRoot, "backups");
const host = "0.0.0.0";
const port = Number(process.env.PORT || 3000);
const ownerEmail = "local@work-organizer";
const maxStateBytes = 5 * 1024 * 1024;

if (!existsSync(serverEntry)) {
  throw new Error("A aplicação ainda não foi compilada. Executa npm run build:windows.");
}

mkdirSync(databaseDirectory, { recursive: true });
mkdirSync(backupDirectory, { recursive: true });
const database = new DatabaseSync(databasePath);
database.exec(`
  PRAGMA journal_mode = WAL;
  CREATE TABLE IF NOT EXISTS organizer_snapshots (
    owner_email TEXT PRIMARY KEY,
    state_json TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
  );
`);
const dailyBackup = join(
  backupDirectory,
  `work-organizer-auto-${new Date().toISOString().slice(0, 10)}.sqlite`,
);
if (!existsSync(dailyBackup)) {
  database.exec(`VACUUM INTO '${dailyBackup.replaceAll("'", "''")}'`);
}

const builtModule = await import(`${pathToFileURL(serverEntry).href}?local=${Date.now()}`);
const builtHandler =
  typeof builtModule.default === "function"
    ? builtModule.default
    : builtModule.default?.fetch?.bind(builtModule.default);
if (!builtHandler) throw new Error("O bundle local não contém um handler HTTP válido.");

const contentTypes = {
  ".css": "text/css; charset=utf-8",
  ".gif": "image/gif",
  ".html": "text/html; charset=utf-8",
  ".ico": "image/x-icon",
  ".jpeg": "image/jpeg",
  ".jpg": "image/jpeg",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".webp": "image/webp",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
};

function sendJson(response, status, value) {
  const body = JSON.stringify(value);
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body),
    "Cache-Control": "no-store",
  });
  response.end(body);
}

async function readBody(request) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > maxStateBytes) throw new Error("PAYLOAD_TOO_LARGE");
    chunks.push(chunk);
  }
  return Buffer.concat(chunks).toString("utf8");
}

async function handleState(request, response) {
  if (request.method === "GET") {
    const row = database
      .prepare("SELECT state_json, updated_at FROM organizer_snapshots WHERE owner_email = ?")
      .get(ownerEmail);
    return sendJson(response, 200, row
      ? { state: JSON.parse(row.state_json), updatedAt: row.updated_at, storage: "local-sqlite" }
      : { state: null, storage: "local-sqlite" });
  }
  if (request.method === "PUT") {
    try {
      const payload = JSON.parse(await readBody(request));
      if (!payload.state || typeof payload.state !== "object" || Array.isArray(payload.state)) {
        return sendJson(response, 400, { error: "Estado de planeamento inválido." });
      }
      const stateJson = JSON.stringify(payload.state);
      const now = new Date().toISOString();
      database.prepare(`
        INSERT INTO organizer_snapshots (owner_email, state_json, created_at, updated_at)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(owner_email) DO UPDATE SET state_json = excluded.state_json, updated_at = excluded.updated_at
      `).run(ownerEmail, stateJson, now, now);
      return sendJson(response, 200, { ok: true, updatedAt: now, storage: "local-sqlite" });
    } catch (error) {
      if (error instanceof Error && error.message === "PAYLOAD_TOO_LARGE") {
        return sendJson(response, 413, { error: "O planeamento excede o limite permitido." });
      }
      return sendJson(response, 400, { error: "Estado de planeamento inválido." });
    }
  }
  return sendJson(response, 405, { error: "Método não permitido." });
}

function staticFileFor(pathname) {
  if (pathname === "/" || pathname.startsWith("/.vite/")) return null;
  let decoded;
  try {
    decoded = decodeURIComponent(pathname);
  } catch {
    return null;
  }
  const file = resolve(clientRoot, `.${decoded}`);
  const rootPrefix = clientRoot.endsWith(sep) ? clientRoot : `${clientRoot}${sep}`;
  return file.startsWith(rootPrefix) && existsSync(file) ? file : null;
}

function toWebRequest(request) {
  const url = new URL(request.url || "/", `http://${request.headers.host || `localhost:${port}`}`);
  const init = { method: request.method, headers: request.headers };
  if (request.method !== "GET" && request.method !== "HEAD") {
    init.body = Readable.toWeb(request);
    init.duplex = "half";
  }
  return new Request(url, init);
}

async function sendWebResponse(nodeResponse, webResponse) {
  const headers = {};
  webResponse.headers.forEach((value, key) => { headers[key] = value; });
  const cookies = webResponse.headers.getSetCookie?.();
  if (cookies?.length) headers["set-cookie"] = cookies;
  nodeResponse.writeHead(webResponse.status, headers);
  if (!webResponse.body) return nodeResponse.end();
  Readable.fromWeb(webResponse.body).pipe(nodeResponse);
}

const server = createServer(async (request, response) => {
  try {
    const url = new URL(request.url || "/", `http://${request.headers.host || `localhost:${port}`}`);
    if (url.pathname === "/__local/health") {
      return sendJson(response, 200, { ok: true, storage: "local-sqlite" });
    }
    if (url.pathname === "/api/state") return await handleState(request, response);

    const staticFile = request.method === "GET" || request.method === "HEAD"
      ? staticFileFor(url.pathname)
      : null;
    if (staticFile) {
      const size = statSync(staticFile).size;
      response.writeHead(200, {
        "Content-Type": contentTypes[extname(staticFile).toLowerCase()] || "application/octet-stream",
        "Content-Length": size,
        "Cache-Control": url.pathname.startsWith("/assets/")
          ? "public, max-age=31536000, immutable"
          : "public, max-age=3600",
      });
      if (request.method === "HEAD") return response.end();
      return createReadStream(staticFile).pipe(response);
    }

    const webResponse = await builtHandler(toWebRequest(request), undefined, {
      waitUntil() {},
      passThroughOnException() {},
    });
    await sendWebResponse(response, webResponse);
  } catch (error) {
    sendJson(response, 500, {
      error: error instanceof Error ? error.message : "Erro inesperado no servidor local.",
    });
  }
});

server.listen(port, host, () => {
  process.stdout.write(`Work Organizer local ativo em http://localhost:${port}\n`);
});

function close() {
  server.close(() => {
    database.close();
    process.exit(0);
  });
}
process.on("SIGINT", close);
process.on("SIGTERM", close);
