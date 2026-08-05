import { sql } from "drizzle-orm";
import { sqliteTable, text } from "drizzle-orm/sqlite-core";

/**
 * The organizer is a single-user workspace today, but ownership is still part
 * of the primary key so a signed-in viewer can never read or overwrite another
 * person's plan. The JSON document is versioned by the application and keeps
 * the client-side domain model atomic while it is still evolving quickly.
 */
export const organizerSnapshots = sqliteTable("organizer_snapshots", {
  ownerEmail: text("owner_email").primaryKey(),
  stateJson: text("state_json").notNull(),
  createdAt: text("created_at").notNull().default(sql`CURRENT_TIMESTAMP`),
  updatedAt: text("updated_at").notNull().default(sql`CURRENT_TIMESTAMP`),
});
