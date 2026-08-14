import { eq } from "drizzle-orm";
import { getDb } from "../../../db";
import { organizerSnapshots } from "../../../db/schema";
import { getChatGPTUser, isLocalOrganizer } from "../../chatgpt-auth";

const MAX_STATE_BYTES = 5 * 1024 * 1024;

async function ownerEmail() {
  const user = await getChatGPTUser();
  if (user?.email) return user.email.trim().toLowerCase();
  if (process.env.NODE_ENV !== "production") return "local-dev@work-organizer";
  return null;
}

function databaseError(error: unknown) {
  const message = error instanceof Error ? error.message : "Erro inesperado.";
  if (
    message.includes("no such table") &&
    message.includes("organizer_snapshots")
  ) {
    return "A base de dados ainda não foi inicializada.";
  }
  return message;
}

export async function GET() {
  if (isLocalOrganizer()) {
    return Response.json({ state: null, storage: "browser" });
  }

  const email = await ownerEmail();
  if (!email) return Response.json({ error: "Inicia sessão para aceder ao planeamento." }, { status: 401 });

  try {
    const db = await getDb();
    const rows = await db.select({ stateJson: organizerSnapshots.stateJson, updatedAt: organizerSnapshots.updatedAt })
      .from(organizerSnapshots)
      .where(eq(organizerSnapshots.ownerEmail, email))
      .limit(1);
    if (!rows[0]) return Response.json({ state: null });
    return Response.json({ state: JSON.parse(rows[0].stateJson), updatedAt: rows[0].updatedAt });
  } catch (error) {
    return Response.json({ error: databaseError(error) }, { status: 500 });
  }
}

export async function PUT(request: Request) {
  if (isLocalOrganizer()) {
    return Response.json({ ok: true, storage: "browser" });
  }

  const email = await ownerEmail();
  if (!email) return Response.json({ error: "Inicia sessão para guardar o planeamento." }, { status: 401 });

  try {
    const raw = await request.text();
    if (new TextEncoder().encode(raw).byteLength > MAX_STATE_BYTES) {
      return Response.json({ error: "O planeamento excede o limite permitido." }, { status: 413 });
    }
    const payload = JSON.parse(raw) as { state?: unknown };
    if (!payload.state || typeof payload.state !== "object" || Array.isArray(payload.state)) {
      return Response.json({ error: "Estado de planeamento inválido." }, { status: 400 });
    }
    const stateJson = JSON.stringify(payload.state);
    const now = new Date().toISOString();
    const db = await getDb();
    await db.insert(organizerSnapshots).values({ ownerEmail: email, stateJson, createdAt: now, updatedAt: now })
      .onConflictDoUpdate({ target: organizerSnapshots.ownerEmail, set: { stateJson, updatedAt: now } });
    return Response.json({ ok: true, updatedAt: now });
  } catch (error) {
    return Response.json({ error: databaseError(error) }, { status: 500 });
  }
}
