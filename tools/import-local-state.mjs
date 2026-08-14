import { existsSync, mkdirSync, readFileSync } from "node:fs";
import { DatabaseSync } from "node:sqlite";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const sourcePath = resolve(process.argv[2] || "");
const targetDirectory = resolve(projectRoot, "data");
const targetPath = resolve(targetDirectory, "work-organizer.sqlite");
if (!process.argv[2] || !existsSync(sourcePath)) {
  throw new Error("Indica o caminho de uma exportação JSON existente.");
}

const exported = JSON.parse(readFileSync(sourcePath, "utf8"));
const state = exported.state ?? exported;
if (!state || typeof state !== "object" || Array.isArray(state)) {
  throw new Error("A exportação não contém um estado válido.");
}

mkdirSync(targetDirectory, { recursive: true });
const database = new DatabaseSync(targetPath);
database.exec(`
  PRAGMA journal_mode = WAL;
  CREATE TABLE IF NOT EXISTS organizer_snapshots (
    owner_email TEXT PRIMARY KEY,
    state_json TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
  );
`);
const now = new Date().toISOString();
database.prepare(`
  INSERT INTO organizer_snapshots (owner_email, state_json, created_at, updated_at)
  VALUES (?, ?, ?, ?)
  ON CONFLICT(owner_email) DO UPDATE SET state_json = excluded.state_json, updated_at = excluded.updated_at
`).run("local@work-organizer", JSON.stringify(state), exported.sourceCreatedAt || now, exported.sourceUpdatedAt || now);
database.close();

console.log(JSON.stringify({
  database: targetPath,
  source: sourcePath,
  tasks: state.tasks?.length ?? 0,
  issues: state.issues?.length ?? 0,
  logs: state.logs?.length ?? 0,
  meetings: state.meetings?.length ?? 0,
  projects: state.projectPreferences?.length ?? 0,
}, null, 2));
