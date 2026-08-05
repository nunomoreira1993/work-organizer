import { drizzle } from "drizzle-orm/d1";
import * as schema from "./schema";

export async function getDb() {
  // Keep the Cloudflare runtime import lazy. This lets the generated Worker be
  // inspected by the build validator in Node without trying to resolve a
  // platform-only module before a database request is actually handled.
  const { env } = await import("cloudflare:workers");
  if (!env.DB) {
    throw new Error(
      "Cloudflare D1 binding `DB` is unavailable. Set the `d1` field in .openai/hosting.json to `DB` or let your control plane inject the real binding values before using the database."
    );
  }

  // This application currently owns a single small table. Creating it lazily
  // keeps a fresh local D1 database (and a first deployment) usable even when
  // the migration command has not been run yet.
  await env.DB.exec(
    "CREATE TABLE IF NOT EXISTS organizer_snapshots (owner_email text PRIMARY KEY NOT NULL, state_json text NOT NULL, created_at text DEFAULT CURRENT_TIMESTAMP NOT NULL, updated_at text DEFAULT CURRENT_TIMESTAMP NOT NULL);",
  );

  return drizzle(env.DB, { schema });
}
