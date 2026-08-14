import { NextRequest, NextResponse } from "next/server";
import { getChatGPTUser } from "../../chatgpt-auth";

type GitLabRequest = {
  action?: "test" | "sync" | "addSpent" | "createIssue" | "projectLabels" | "renderMarkdown" | "updateIssueState" | "updateIssueDescription" | "updateIssueLabels" | "moveIssue";
  baseUrl?: string;
  token?: string;
  projectId?: number;
  toProjectId?: number;
  issueIid?: number;
  duration?: string;
  title?: string;
  description?: string;
  markdown?: string;
  estimateHours?: number;
  labels?: string[];
  relatedIssues?: { projectId?: number; issueIid?: number }[];
  state?: "opened" | "closed";
};

function normaliseBaseUrl(input: string) {
  const parsed = new URL(input);
  if (parsed.protocol !== "https:") {
    throw new Error("A ligação ao GitLab tem de usar HTTPS.");
  }
  if (
    parsed.username ||
    parsed.password ||
    ["localhost", "127.0.0.1", "0.0.0.0", "::1"].includes(parsed.hostname)
  ) {
    throw new Error("URL do GitLab inválido.");
  }
  return parsed.origin;
}

async function gitlabFetch(baseUrl: string, token: string, path: string, init?: RequestInit) {
  const response = await fetch(`${baseUrl}/api/v4${path}`, {
    ...init,
    headers: {
      Accept: "application/json",
      "PRIVATE-TOKEN": token,
      ...(init?.headers ?? {}),
    },
  });

  const text = await response.text();
  let data: unknown = null;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    data = text;
  }

  if (!response.ok) {
    const detail =
      typeof data === "object" && data && "message" in data
        ? String((data as { message: unknown }).message)
        : `Erro ${response.status}`;
    throw new Error(detail);
  }
  return data;
}

async function gitlabFetchAll(baseUrl: string, token: string, path: string) {
  const items: unknown[] = [];
  let page = 1;

  while (page <= 100) {
    const separator = path.includes("?") ? "&" : "?";
    const response = await fetch(`${baseUrl}/api/v4${path}${separator}page=${page}&per_page=100`, {
      headers: { Accept: "application/json", "PRIVATE-TOKEN": token },
    });
    const data = (await response.json()) as unknown;
    if (!response.ok) {
      const detail = typeof data === "object" && data && "message" in data
        ? String((data as { message: unknown }).message)
        : `Erro ${response.status}`;
      throw new Error(detail);
    }
    if (!Array.isArray(data)) throw new Error("Resposta inesperada do GitLab.");
    items.push(...data);
    const nextPage = response.headers.get("x-next-page");
    if (!nextPage) break;
    page = Number(nextPage);
    if (!Number.isFinite(page) || page < 1) break;
  }

  return items;
}

type GraphTimelog = {
  id?: string;
  spentAt?: string;
  timeSpent?: number;
  summary?: string | null;
  issue?: { iid?: string | number; title?: string; webUrl?: string; projectId?: string | number } | null;
};

function numericProjectId(value: string | number | undefined) {
  if (typeof value === "number") return value;
  const match = String(value ?? "").match(/(?:Project\/)?(\d+)$/);
  return match ? Number(match[1]) : 0;
}

async function gitlabGraphQL(baseUrl: string, token: string, query: string, variables: Record<string, unknown>) {
  const response = await fetch(`${baseUrl}/api/graphql`, {
    method: "POST",
    headers: { Accept: "application/json", "Content-Type": "application/json", Authorization: `Bearer ${token}` },
    body: JSON.stringify({ query, variables }),
  });
  const data = await response.json() as { data?: unknown; errors?: { message?: string }[] };
  if (!response.ok || data.errors?.length) throw new Error(data.errors?.[0]?.message ?? `GraphQL devolveu ${response.status}.`);
  return data.data as { timelogs?: { nodes?: GraphTimelog[]; pageInfo?: { hasNextPage?: boolean; endCursor?: string | null } } };
}

async function fetchMyTimelogs(baseUrl: string, token: string, username: string) {
  const query = `query WorkOrganizerTimelogs($username: String!, $after: String) {
    timelogs(username: $username, first: 100, after: $after) {
      nodes { id spentAt timeSpent summary issue { iid title webUrl projectId } }
      pageInfo { hasNextPage endCursor }
    }
  }`;
  const result: { id: string; project_id: number; issue_iid: number; issue_title: string; issue_web_url?: string; spent_at: string; time_spent: number; summary?: string }[] = [];
  let after: string | null = null;
  for (let page = 0; page < 100; page += 1) {
    const data = await gitlabGraphQL(baseUrl, token, query, { username, after });
    const connection = data.timelogs;
    for (const entry of connection?.nodes ?? []) {
      const projectId = numericProjectId(entry.issue?.projectId);
      const issueIid = Number(entry.issue?.iid ?? 0);
      if (!projectId || !issueIid || !entry.spentAt || !entry.timeSpent) continue;
      result.push({
        id: entry.id ?? `${projectId}-${issueIid}-${entry.spentAt}-${result.length}`,
        project_id: projectId,
        issue_iid: issueIid,
        issue_title: entry.issue?.title ?? `Issue #${issueIid}`,
        issue_web_url: entry.issue?.webUrl,
        spent_at: entry.spentAt,
        time_spent: entry.timeSpent,
        summary: entry.summary ?? undefined,
      });
    }
    if (!connection?.pageInfo?.hasNextPage || !connection.pageInfo.endCursor) break;
    after = connection.pageInfo.endCursor;
  }
  return result;
}

export async function POST(request: NextRequest) {
  try {
    const user = await getChatGPTUser();
    if (!user && process.env.NODE_ENV === "production") {
      return NextResponse.json({ ok: false, error: "Inicia sessão para usar a integração GitLab." }, { status: 401 });
    }
    const body = (await request.json()) as GitLabRequest;
    const baseUrl = normaliseBaseUrl(body.baseUrl?.trim() ?? "");
    const token = body.token?.trim() ?? "";
    if (!token) throw new Error("Indica um Personal Access Token.");

    if (body.action === "test") {
      const user = await gitlabFetch(baseUrl, token, "/user");
      return NextResponse.json({ ok: true, user });
    }

    if (body.action === "sync") {
      // Keep the planning dataset focused: all accessible Site Server issues
      // updated in the last month, regardless of their assignee.
      const [user, allProjects] = await Promise.all([
        gitlabFetch(baseUrl, token, "/user"),
        gitlabFetchAll(baseUrl, token, "/projects?membership=true&archived=false&with_issues_enabled=true&simple=true"),
      ]);
      const projects = (allProjects as {
        id?: number;
        name?: string;
        name_with_namespace?: string;
        path_with_namespace?: string;
      }[]).filter((project) =>
        `${project.name ?? ""} ${project.name_with_namespace ?? ""} ${project.path_with_namespace ?? ""}`
          .toLowerCase()
          .includes("site server"),
      );
      const projectIds = new Set(
        projects.map((project) => project.id).filter((id): id is number => Boolean(id)),
      );
      const updatedAfter = new Date();
      updatedAfter.setMonth(updatedAfter.getMonth() - 1);
      const query = new URLSearchParams({
        scope: "all",
        state: "all",
        non_archived: "true",
        updated_after: updatedAfter.toISOString(),
        order_by: "updated_at",
        sort: "desc",
      });
      const issues = (await gitlabFetchAll(baseUrl, token, `/issues?${query}`)).filter(
        (issue) => projectIds.has((issue as { project_id?: number }).project_id ?? 0),
      );
      let timelogs: Awaited<ReturnType<typeof fetchMyTimelogs>> = [];
      let timelogWarning: string | undefined;
      try {
        const username = String((user as { username?: string }).username ?? "");
        if (username) {
          timelogs = (await fetchMyTimelogs(baseUrl, token, username)).filter(
            (entry) =>
              projectIds.has(entry.project_id) &&
              entry.spent_at >= updatedAfter.toISOString(),
          );
        }
      } catch (error) {
        timelogWarning = error instanceof Error ? error.message : "Não foi possível consultar os timelogs pessoais.";
      }
      const issueProjectIds = [...new Set((issues as { project_id?: number }[]).map((issue) => issue.project_id).filter((id): id is number => Boolean(id)))];
      const labelEntries = await Promise.all(issueProjectIds.map(async (projectId) => {
        try {
          const labels = await gitlabFetchAll(baseUrl, token, `/projects/${projectId}/labels`);
          return [String(projectId), labels] as const;
        } catch {
          return [String(projectId), []] as const;
        }
      }));
      const labelCatalog = Object.fromEntries(labelEntries);
      return NextResponse.json({ ok: true, user, issues, projects, labelCatalog, timelogs, timelogsAvailable: !timelogWarning, timelogWarning });
    }

    if (body.action === "projectLabels") {
      if (!body.projectId) throw new Error("Seleciona um projeto GitLab.");
      const labels = await gitlabFetchAll(baseUrl, token, `/projects/${body.projectId}/labels`);
      return NextResponse.json({ ok: true, labels });
    }

    if (body.action === "renderMarkdown") {
      let project: string | undefined;
      if (body.projectId) {
        const context = (await gitlabFetch(
          baseUrl,
          token,
          `/projects/${body.projectId}`,
        )) as { path_with_namespace?: string };
        project = context.path_with_namespace;
      }
      const rendered = await gitlabFetch(baseUrl, token, "/markdown", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          text: body.markdown ?? "",
          gfm: true,
          ...(project ? { project } : {}),
        }),
      });
      return NextResponse.json({
        ok: true,
        html: (rendered as { html?: string }).html ?? "",
      });
    }

    if (body.action === "updateIssueState") {
      if (
        !body.projectId ||
        !body.issueIid ||
        (body.state !== "opened" && body.state !== "closed")
      ) {
        throw new Error("Faltam dados da US ou do estado.");
      }
      const issue = await gitlabFetch(
        baseUrl,
        token,
        `/projects/${body.projectId}/issues/${body.issueIid}`,
        {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            state_event: body.state === "closed" ? "close" : "reopen",
          }),
        },
      );
      return NextResponse.json({ ok: true, issue });
    }

    if (body.action === "updateIssueDescription") {
      if (!body.projectId || !body.issueIid) {
        throw new Error("Faltam dados da US.");
      }
      const issue = await gitlabFetch(
        baseUrl,
        token,
        `/projects/${body.projectId}/issues/${body.issueIid}`,
        {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ description: body.description ?? "" }),
        },
      );
      return NextResponse.json({ ok: true, issue });
    }

    if (body.action === "updateIssueLabels") {
      if (!body.projectId || !body.issueIid || !Array.isArray(body.labels)) {
        throw new Error("Faltam dados da US ou das labels.");
      }
      const issue = await gitlabFetch(
        baseUrl,
        token,
        `/projects/${body.projectId}/issues/${body.issueIid}`,
        {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            labels: body.labels
              .filter((label) => typeof label === "string" && label.trim())
              .join(","),
          }),
        },
      );
      return NextResponse.json({ ok: true, issue });
    }

    if (body.action === "moveIssue") {
      if (!body.projectId || !body.issueIid || !body.toProjectId) {
        throw new Error("Seleciona o projeto de destino da US.");
      }
      if (body.projectId === body.toProjectId) {
        throw new Error("A US já pertence ao projeto selecionado.");
      }
      const issue = await gitlabFetch(
        baseUrl,
        token,
        `/projects/${body.projectId}/issues/${body.issueIid}/move`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ to_project_id: body.toProjectId }),
        },
      );
      return NextResponse.json({ ok: true, issue });
    }

    if (body.action === "addSpent") {
      if (!body.projectId || !body.issueIid || !body.duration) {
        throw new Error("Faltam dados da Issue ou da duração.");
      }
      const duration = encodeURIComponent(body.duration);
      const result = await gitlabFetch(
        baseUrl,
        token,
        `/projects/${body.projectId}/issues/${body.issueIid}/add_spent_time?duration=${duration}`,
        { method: "POST" },
      );
      return NextResponse.json({ ok: true, result });
    }

    if (body.action === "createIssue") {
      if (!body.projectId || !body.title?.trim()) {
        throw new Error("Seleciona um projeto e indica o título da tarefa.");
      }
      const currentUser = await gitlabFetch(baseUrl, token, "/user") as { id?: number };
      if (!currentUser.id) throw new Error("Não foi possível identificar o teu utilizador GitLab.");
      const created = await gitlabFetch(
        baseUrl,
        token,
        `/projects/${body.projectId}/issues`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            title: body.title.trim(),
            description: body.description?.trim() ?? "",
            assignee_ids: [currentUser.id],
            labels: (body.labels ?? []).filter((label) => typeof label === "string" && label.trim()).join(","),
          }),
        },
      ) as { iid?: number };
      if (!created.iid) throw new Error("O GitLab criou a tarefa sem devolver o respetivo número.");
      const estimateHours = Math.max(0, Number(body.estimateHours) || 0);
      if (estimateHours > 0) {
        const duration = encodeURIComponent(`${Math.round(estimateHours * 60)}m`);
        await gitlabFetch(baseUrl, token, `/projects/${body.projectId}/issues/${created.iid}/time_estimate?duration=${duration}`, { method: "POST" });
      }
      const relatedIssues = (body.relatedIssues ?? []).filter(
        (related) => related.projectId && related.issueIid,
      );
      const linkResults = await Promise.allSettled(
        relatedIssues.map((related) =>
          gitlabFetch(
            baseUrl,
            token,
            `/projects/${body.projectId}/issues/${created.iid}/links`,
            {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                target_project_id: related.projectId,
                target_issue_iid: related.issueIid,
                link_type: "relates_to",
              }),
            },
          ),
        ),
      );
      const failedRelatedIssues = linkResults.filter(
        (result) => result.status === "rejected",
      ).length;
      const [issue, project] = await Promise.all([
        gitlabFetch(baseUrl, token, `/projects/${body.projectId}/issues/${created.iid}`),
        gitlabFetch(baseUrl, token, `/projects/${body.projectId}`),
      ]);
      return NextResponse.json({
        ok: true,
        issue,
        project,
        failedRelatedIssues,
      });
    }

    throw new Error("Ação desconhecida.");
  } catch (error) {
    return NextResponse.json(
      { ok: false, error: error instanceof Error ? error.message : "Erro inesperado." },
      { status: 400 },
    );
  }
}
