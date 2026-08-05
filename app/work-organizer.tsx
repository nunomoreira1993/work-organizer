"use client";

import { PublicClientApplication } from "@azure/msal-browser";
import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type DragEvent,
} from "react";

type View = "today" | "timeline" | "completed" | "analytics" | "settings";
type Zoom = "day" | "week" | "month";
type SettingsSection = "gitlab" | "calendar" | "capacity" | "clients";
type TaskSource = "gitlab" | "outlook";
type WorkPhase =
  | "Refinement"
  | "Desenvolvimento"
  | "QA"
  | "Deploy"
  | "Suporte"
  | "Reunião"
  | "Outro";

type GitLabLabel = {
  name: string;
  color: string;
  textColor: string;
  description?: string | null;
};

type Task = {
  id: string;
  rootId: string;
  allocationId: string;
  projectId: number;
  iid: number;
  title: string;
  client: string;
  project: string;
  color: string;
  date: string;
  start: number;
  estimate: number;
  spent: number;
  due: string;
  dueDate?: string | null;
  type: string;
  phase: WorkPhase;
  description?: string;
  priority: "Alta" | "Média" | "Normal";
  fixed?: boolean;
  webUrl?: string;
  source: TaskSource;
  hasEstimate?: boolean;
  labels?: GitLabLabel[];
};

type IssueRecord = {
  id: string;
  projectId: number;
  iid: number;
  title: string;
  client: string;
  project: string;
  color: string;
  state: "opened" | "closed";
  estimateTotal: number;
  teamSpent: number;
  personalSpent: number;
  assigneeCount: number;
  due: string;
  dueDate?: string | null;
  type: string;
  priority: Task["priority"];
  webUrl?: string;
  hasEstimate: boolean;
  labels: GitLabLabel[];
};

type AllocationDraft = {
  issueId: string;
  allocationId?: string;
  phase: WorkPhase;
  hours: number;
  date: string;
  start: number;
  distribution: "manual" | "automatic";
  description: string;
};

type CreateIssueDraft = {
  projectId: number;
  title: string;
  description: string;
  estimateHours: number;
  phase: WorkPhase;
  hours: number;
  date: string;
  start: number;
  distribution: "manual" | "automatic";
  labels: string[];
  meetingId?: string;
};

type WorkLog = {
  id: string;
  dateKey: string;
  client: string;
  project: string;
  task: string;
  type: string;
  hours: number;
  estimate?: number;
  webUrl?: string;
  projectId?: number;
  issueIid?: number;
  source?: "gitlab" | "timer";
};

type Config = { baseUrl: string; token: string };
type OutlookConfig = { tenantId: string; clientId: string };
type CapacityConfig = {
  dailyHours: number;
  startHour: number;
  workDays: number[];
};
type ProjectPreference = {
  projectId: number;
  project: string;
  client: string;
  color: string;
  enabled: boolean;
};
type CalendarMeeting = {
  id: string;
  subject: string;
  start: string;
  end: string;
  webUrl?: string;
  joinUrl?: string;
  projectId?: number;
  linkedIssueId?: string;
  linkedIssueIid?: number;
  linkedIssueWebUrl?: string;
};

type PersistedState = {
  version: 1;
  tasks: Task[];
  issues: IssueRecord[];
  logs: WorkLog[];
  meetings: CalendarMeeting[];
  capacity: CapacityConfig;
  projectPreferences: ProjectPreference[];
  labelCatalog: Record<string, GitLabLabel[]>;
};

type GitLabIssue = {
  project_id: number;
  iid: number;
  title: string;
  state: "opened" | "closed";
  web_url?: string;
  due_date?: string | null;
  closed_at?: string | null;
  updated_at?: string;
  labels?: string[];
  assignees?: { id: number; username?: string }[];
  time_stats?: { time_estimate?: number; total_time_spent?: number };
};

type GitLabTimelog = {
  id: string;
  project_id: number;
  issue_iid: number;
  issue_title: string;
  issue_web_url?: string;
  spent_at: string;
  time_spent: number;
  summary?: string;
};

type GitLabProject = {
  id: number;
  name?: string;
  name_with_namespace?: string;
  path_with_namespace?: string;
  web_url?: string;
};
type GitLabLabelResponse = {
  name?: string;
  color?: string;
  text_color?: string;
  description?: string | null;
};

const CLIENT_COLORS = [
  "#5f5bf6",
  "#e97042",
  "#1d9a78",
  "#c44770",
  "#3977d5",
  "#a06bd6",
  "#cf8d2f",
];
const WEEKDAYS = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SÁB"];
const MONTHS = [
  "JAN",
  "FEV",
  "MAR",
  "ABR",
  "MAI",
  "JUN",
  "JUL",
  "AGO",
  "SET",
  "OUT",
  "NOV",
  "DEZ",
];
const DEFAULT_WORK_DAYS = [1, 2, 3, 4, 5];
const DEFAULT_CAPACITY: CapacityConfig = {
  dailyHours: 8,
  startHour: 9,
  workDays: DEFAULT_WORK_DAYS,
};
const DEFAULT_GITLAB_CONFIG: Config = {
  baseUrl: "https://gitlab.ddsdev.deloitte.pt",
  token: "",
};
const DEFAULT_OUTLOOK_CONFIG: OutlookConfig = { tenantId: "", clientId: "" };
const priorityOrder = { Alta: 0, Média: 1, Normal: 2 } as const;

function dateKey(date: Date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function currentDateKey() {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/Lisbon",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date());
  const value = new Map(parts.map((part) => [part.type, part.value]));
  return `${value.get("year")}-${value.get("month")}-${value.get("day")}`;
}

function fromKey(key: string) {
  const [y, m, d] = key.split("-").map(Number);
  return new Date(y, m - 1, d, 12);
}

function addDays(key: string, amount: number) {
  const value = fromKey(key);
  value.setDate(value.getDate() + amount);
  return dateKey(value);
}

function isBusinessDay(key: string, workDays = DEFAULT_WORK_DAYS) {
  const day = fromKey(key).getDay();
  return workDays.includes(day);
}

function nextBusinessDay(
  key: string,
  includeCurrent = false,
  workDays = DEFAULT_WORK_DAYS,
) {
  let next = includeCurrent ? key : addDays(key, 1);
  while (!isBusinessDay(next, workDays)) next = addDays(next, 1);
  return next;
}

function businessDays(
  startKey: string,
  count: number,
  workDays = DEFAULT_WORK_DAYS,
) {
  const result: string[] = [];
  let cursor = nextBusinessDay(startKey, true, workDays);
  while (result.length < count) {
    result.push(cursor);
    cursor = nextBusinessDay(cursor, false, workDays);
  }
  return result;
}

function addBusinessDays(
  startKey: string,
  amount: number,
  workDays = DEFAULT_WORK_DAYS,
) {
  let cursor = nextBusinessDay(startKey, true, workDays);
  for (let index = 0; index < amount; index += 1)
    cursor = nextBusinessDay(cursor, false, workDays);
  return cursor;
}

function mondayFor(key: string) {
  const value = fromKey(key);
  const offset = value.getDay() === 0 ? -6 : 1 - value.getDay();
  value.setDate(value.getDate() + offset);
  return dateKey(value);
}

function monthBusinessDays(key: string, workDays = DEFAULT_WORK_DAYS) {
  const source = fromKey(key);
  const cursor = new Date(source.getFullYear(), source.getMonth(), 1, 12);
  const result: string[] = [];
  while (cursor.getMonth() === source.getMonth()) {
    const keyValue = dateKey(cursor);
    if (isBusinessDay(keyValue, workDays)) result.push(keyValue);
    cursor.setDate(cursor.getDate() + 1);
  }
  return result;
}

function visibleDaysFor(
  anchor: string,
  zoom: Zoom,
  workDays = DEFAULT_WORK_DAYS,
) {
  if (zoom === "day") return [nextBusinessDay(anchor, true, workDays)];
  if (zoom === "week")
    return businessDays(mondayFor(anchor), workDays.length, workDays);
  return monthBusinessDays(anchor, workDays);
}

function movePeriod(
  anchor: string,
  zoom: Zoom,
  direction: number,
  workDays = DEFAULT_WORK_DAYS,
) {
  const value = fromKey(anchor);
  if (zoom === "day" && direction > 0)
    return nextBusinessDay(anchor, false, workDays);
  if (zoom === "week") value.setDate(value.getDate() + direction * 7);
  if (zoom === "month") value.setMonth(value.getMonth() + direction, 1);
  if (zoom === "day" && direction < 0) {
    let previous = addDays(anchor, -1);
    while (!isBusinessDay(previous, workDays)) previous = addDays(previous, -1);
    return previous;
  }
  return dateKey(value);
}

function projectColor(projectId: number) {
  return CLIENT_COLORS[Math.abs(projectId) % CLIENT_COLORS.length];
}

function projectClient(project?: GitLabProject) {
  const parts = project?.path_with_namespace?.split("/") ?? [];
  const raw = parts.length > 1 ? parts.at(-2)! : (project?.name ?? "GitLab");
  return raw
    .replace(/[-_]/g, " ")
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function formatHours(hours: number) {
  const rounded = Math.round(hours * 100) / 100;
  return `${Number.isInteger(rounded) ? rounded : rounded.toFixed(2).replace(/0$/, "")} h`;
}

function formatClock(hours: number) {
  const totalMinutes = Math.round(hours * 60);
  const hour = Math.floor(totalMinutes / 60) % 24;
  const minute = totalMinutes % 60;
  return `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;
}

function formatTimer(seconds: number) {
  const h = Math.floor(seconds / 3600)
    .toString()
    .padStart(2, "0");
  const m = Math.floor((seconds % 3600) / 60)
    .toString()
    .padStart(2, "0");
  const s = Math.floor(seconds % 60)
    .toString()
    .padStart(2, "0");
  return `${h}:${m}:${s}`;
}

function formatDate(key: string, includeYear = false) {
  const value = fromKey(key);
  return `${value.getDate()} ${MONTHS[value.getMonth()].toLowerCase()}${includeYear ? ` ${value.getFullYear()}` : ""}`;
}

function taskType(labels: string[] = []) {
  const joined = labels.join(" ").toLowerCase();
  if (/meeting|reuni[aã]o/.test(joined)) return "Reunião";
  if (/support|suporte/.test(joined)) return "Suporte";
  if (/infra/.test(joined)) return "Infraestrutura";
  if (/analysis|an[aá]lise/.test(joined)) return "Análise";
  if (/planning|planeamento/.test(joined)) return "Planeamento";
  return "Desenvolvimento";
}

function taskPriority(labels: string[] = []): Task["priority"] {
  const joined = labels.join(" ").toLowerCase();
  if (/priority::high|priority:high|high priority|alta/.test(joined))
    return "Alta";
  if (/priority::medium|priority:medium|m[eé]dia/.test(joined)) return "Média";
  return "Normal";
}

function safeLabelColor(color?: string) {
  return /^#[0-9a-f]{6}$/i.test(color ?? "") ? color! : "#6b7280";
}

function normaliseLabel(label: GitLabLabelResponse | string): GitLabLabel {
  if (typeof label === "string")
    return { name: label, color: "#6b7280", textColor: "#ffffff" };
  return {
    name: label.name?.trim() || "Sem nome",
    color: safeLabelColor(label.color),
    textColor: safeLabelColor(label.text_color || "#ffffff"),
    description: label.description,
  };
}

function labelsForIssue(names: string[] = [], catalog: GitLabLabel[] = []) {
  const byName = new Map(catalog.map((label) => [label.name, label]));
  return names.map((name) => byName.get(name) ?? normaliseLabel(name));
}

function planTasks(
  input: Task[],
  startKey = dateKey(new Date()),
  capacity = DEFAULT_CAPACITY,
) {
  const fixed = input.filter((task) => task.fixed);
  const flexible = input.filter((task) => !task.fixed);
  const used = new Map<string, number>();
  fixed.forEach((task) =>
    used.set(task.date, (used.get(task.date) ?? 0) + task.estimate),
  );

  const roots = new Map<string, Task[]>();
  flexible.forEach((task) =>
    roots.set(task.allocationId, [
      ...(roots.get(task.allocationId) ?? []),
      task,
    ]),
  );
  const canonical = [...roots.values()]
    .map((segments) => {
      const base = segments[0];
      return {
        ...base,
        estimate: segments.reduce((sum, task) => sum + task.estimate, 0),
        spent: segments.reduce((sum, task) => sum + task.spent, 0),
      };
    })
    .sort((a, b) => {
      const dueA = a.dueDate
        ? fromKey(a.dueDate).getTime()
        : Number.MAX_SAFE_INTEGER;
      const dueB = b.dueDate
        ? fromKey(b.dueDate).getTime()
        : Number.MAX_SAFE_INTEGER;
      return (
        dueA - dueB || priorityOrder[a.priority] - priorityOrder[b.priority]
      );
    });

  const result = [...fixed];
  let cursor = nextBusinessDay(startKey, true, capacity.workDays);
  canonical.forEach((task) => {
    let remaining = Math.max(0.25, task.estimate);
    let remainingSpent = task.spent;
    let segment = 0;
    while (remaining > 0.001 && segment < 500) {
      cursor = nextBusinessDay(cursor, true, capacity.workDays);
      const occupied = used.get(cursor) ?? 0;
      const available = Math.max(0, capacity.dailyHours - occupied);
      if (available < 0.01) {
        cursor = nextBusinessDay(cursor, false, capacity.workDays);
        continue;
      }
      const hours = Math.min(remaining, available);
      const spent = Math.min(hours, remainingSpent);
      result.push({
        ...task,
        id: `${task.rootId}::${segment}`,
        date: cursor,
        start: capacity.startHour + occupied,
        estimate: hours,
        spent,
      });
      used.set(cursor, occupied + hours);
      remaining -= hours;
      remainingSpent = Math.max(0, remainingSpent - spent);
      segment += 1;
      if (remaining > 0.001)
        cursor = nextBusinessDay(cursor, false, capacity.workDays);
    }
  });
  return result;
}

const nav: { id: View; label: string; icon: string }[] = [
  { id: "today", label: "Hoje", icon: "⌂" },
  { id: "timeline", label: "Timeline", icon: "↔" },
  { id: "completed", label: "Trabalho realizado", icon: "✓" },
  { id: "analytics", label: "Analytics", icon: "⌁" },
];

function readStorage<T>(key: string, fallback: T): T {
  if (typeof window === "undefined") return fallback;
  const saved = window.localStorage.getItem(key);
  if (!saved) return fallback;
  try {
    return JSON.parse(saved) as T;
  } catch {
    return fallback;
  }
}

function readRealTasks() {
  return (
    readStorage<Task[]>("work-organizer.tasks", [])
      .filter(
        (task) => (task as unknown as { source?: string }).source !== "demo",
      )
      // Planeamentos criados pela versão antiga eram automáticos e não representam
      // uma decisão do utilizador. Mantemos apenas reuniões e alocações pessoais
      // criadas explicitamente no novo fluxo.
      .filter(
        (task) =>
          task.source === "outlook" ||
          Boolean(
            task.allocationId && !task.allocationId.startsWith("legacy-"),
          ),
      )
      .map((task) => ({
        ...task,
        allocationId: task.allocationId,
        phase:
          task.phase ??
          (task.type === "Reunião"
            ? "Reunião"
            : task.type === "Suporte"
              ? "Suporte"
              : "Desenvolvimento"),
      }))
  );
}

function readIssues() {
  return readStorage<IssueRecord[]>("work-organizer.issues", []);
}

function readRealLogs() {
  const demoIds = new Set(["l1", "l2", "l3", "l4"]);
  return readStorage<WorkLog[]>("work-organizer.logs", []).filter(
    (log) => !demoIds.has(log.id),
  );
}

function normaliseCapacity(saved: CapacityConfig) {
  return {
    dailyHours: Math.min(
      16,
      Math.max(1, Number(saved.dailyHours) || DEFAULT_CAPACITY.dailyHours),
    ),
    startHour: Math.min(23, Math.max(0, Number(saved.startHour) || 0)),
    workDays: saved.workDays?.length
      ? [...new Set(saved.workDays)]
          .filter((day) => day >= 0 && day <= 6)
          .sort()
      : DEFAULT_WORK_DAYS,
  };
}

function readCapacity() {
  return normaliseCapacity(
    readStorage<CapacityConfig>("work-organizer.capacity", DEFAULT_CAPACITY),
  );
}

export default function WorkOrganizer({
  user,
}: {
  user: { displayName: string; email: string };
}) {
  const todayKey = currentDateKey();
  const [view, setView] = useState<View>("today");
  // Keep the server and browser's first render identical. Browser-only data is
  // restored after hydration below.
  const [tasks, setTasks] = useState<Task[]>([]);
  const [issues, setIssues] = useState<IssueRecord[]>([]);
  const [logs, setLogs] = useState<WorkLog[]>([]);
  const [meetings, setMeetings] = useState<CalendarMeeting[]>([]);
  const [activeTask, setActiveTask] = useState<string | null>(null);
  const [elapsed, setElapsed] = useState(0);
  const [dragged, setDragged] = useState<string | null>(null);
  const [config, setConfig] = useState<Config>(DEFAULT_GITLAB_CONFIG);
  const [outlookConfig, setOutlookConfig] =
    useState<OutlookConfig>(DEFAULT_OUTLOOK_CONFIG);
  const [capacity, setCapacity] = useState<CapacityConfig>(DEFAULT_CAPACITY);
  const [projectPreferences, setProjectPreferences] = useState<
    ProjectPreference[]
  >([]);
  const [labelCatalog, setLabelCatalog] = useState<
    Record<string, GitLabLabel[]>
  >({});
  const [showToken, setShowToken] = useState(false);
  const [settingsSection, setSettingsSection] =
    useState<SettingsSection>("gitlab");
  const [connection, setConnection] = useState<{
    state: "idle" | "loading" | "ok" | "error";
    message: string;
  }>({ state: "idle", message: "" });
  const [outlookConnection, setOutlookConnection] = useState<{
    state: "idle" | "loading" | "ok" | "error";
    message: string;
  }>({ state: "idle", message: "" });
  const [syncing, setSyncing] = useState(false);
  const [syncingCalendar, setSyncingCalendar] = useState(false);
  const [toast, setToast] = useState("");
  const [searchOpen, setSearchOpen] = useState(false);
  const [notificationsOpen, setNotificationsOpen] = useState(false);
  const [globalSearch, setGlobalSearch] = useState("");
  const [persistenceStatus, setPersistenceStatus] = useState<
    "loading" | "ready" | "local"
  >("loading");
  const msalRef = useRef<PublicClientApplication | null>(null);

  useEffect(() => {
    let cancelled = false;
    async function hydrateFromDatabase() {
      const localState: PersistedState = {
        version: 1,
        tasks: readRealTasks(),
        issues: readIssues(),
        logs: readRealLogs(),
        meetings: readStorage("work-organizer.meetings", []),
        capacity: readCapacity(),
        projectPreferences: readStorage("work-organizer.projects", []),
        labelCatalog: readStorage("work-organizer.labels", {}),
      };
      setConfig(readStorage("work-organizer.gitlab", DEFAULT_GITLAB_CONFIG));
      setOutlookConfig(
        readStorage("work-organizer.outlook", DEFAULT_OUTLOOK_CONFIG),
      );
      try {
        const response = await fetch("/api/state", {
          headers: { Accept: "application/json" },
        });
        const data = (await response.json()) as {
          state?: Partial<PersistedState> | null;
          error?: string;
        };
        if (!response.ok)
          throw new Error(
            data.error ?? "Não foi possível abrir a base de dados.",
          );
        if (data.state) {
          if (Array.isArray(data.state.tasks)) setTasks(data.state.tasks);
          if (Array.isArray(data.state.issues)) setIssues(data.state.issues);
          if (Array.isArray(data.state.logs)) setLogs(data.state.logs);
          if (Array.isArray(data.state.meetings))
            setMeetings(data.state.meetings);
          if (data.state.capacity)
            setCapacity(normaliseCapacity(data.state.capacity));
          if (Array.isArray(data.state.projectPreferences))
            setProjectPreferences(data.state.projectPreferences);
          if (
            data.state.labelCatalog &&
            typeof data.state.labelCatalog === "object"
          )
            setLabelCatalog(data.state.labelCatalog);
        } else {
          if (!cancelled) {
            setTasks(localState.tasks);
            setIssues(localState.issues);
            setLogs(localState.logs);
            setMeetings(localState.meetings);
            setCapacity(localState.capacity);
            setProjectPreferences(localState.projectPreferences);
            setLabelCatalog(localState.labelCatalog);
          }
          const migration = await fetch("/api/state", {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ state: localState }),
          });
          const result = (await migration.json()) as { error?: string };
          if (!migration.ok)
            throw new Error(
              result.error ?? "Não foi possível migrar os dados existentes.",
            );
        }
        if (!cancelled) setPersistenceStatus("ready");
      } catch (error) {
        if (!cancelled) {
          setTasks(localState.tasks);
          setIssues(localState.issues);
          setLogs(localState.logs);
          setMeetings(localState.meetings);
          setCapacity(localState.capacity);
          setProjectPreferences(localState.projectPreferences);
          setLabelCatalog(localState.labelCatalog);
          setPersistenceStatus("local");
          setToast(
            `${error instanceof Error ? error.message : "A base de dados não está disponível."} Os dados continuam guardados neste browser.`,
          );
        }
      }
    }
    void hydrateFromDatabase();
    return () => {
      cancelled = true;
    };
  }, []);
  useEffect(() => {
    if (persistenceStatus === "loading") return;
    window.localStorage.setItem("work-organizer.tasks", JSON.stringify(tasks));
    window.localStorage.setItem("work-organizer.issues", JSON.stringify(issues));
    window.localStorage.setItem("work-organizer.logs", JSON.stringify(logs));
    window.localStorage.setItem(
      "work-organizer.meetings",
      JSON.stringify(meetings),
    );
    window.localStorage.setItem(
      "work-organizer.capacity",
      JSON.stringify(capacity),
    );
    window.localStorage.setItem(
      "work-organizer.projects",
      JSON.stringify(projectPreferences),
    );
    window.localStorage.setItem(
      "work-organizer.labels",
      JSON.stringify(labelCatalog),
    );
    const controller = new AbortController();
    const timer = window.setTimeout(async () => {
      try {
        const state: PersistedState = {
          version: 1,
          tasks,
          issues,
          logs,
          meetings,
          capacity,
          projectPreferences,
          labelCatalog,
        };
        const response = await fetch("/api/state", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ state }),
          signal: controller.signal,
        });
        if (!response.ok) {
          const data = (await response.json()) as { error?: string };
          throw new Error(
            data.error ?? "Não foi possível guardar as alterações.",
          );
        }
        setPersistenceStatus("ready");
      } catch (error) {
        if (controller.signal.aborted) return;
        setPersistenceStatus("local");
        setToast(
          error instanceof Error
            ? error.message
            : "Não foi possível guardar as alterações.",
        );
      }
    }, 500);
    return () => {
      window.clearTimeout(timer);
      controller.abort();
    };
  }, [
    tasks,
    issues,
    logs,
    meetings,
    capacity,
    projectPreferences,
    labelCatalog,
    persistenceStatus,
  ]);
  useEffect(() => {
    if (!activeTask) return;
    const interval = window.setInterval(
      () => setElapsed((value) => value + 1),
      1000,
    );
    return () => window.clearInterval(interval);
  }, [activeTask]);
  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(""), 3600);
    return () => window.clearTimeout(timer);
  }, [toast]);

  const issueStateById = useMemo(
    () => new Map(issues.map((issue) => [issue.id, issue.state])),
    [issues],
  );
  const planningTasks = useMemo(
    () =>
      tasks.filter(
        (task) =>
          task.source === "outlook" ||
          issueStateById.get(task.rootId) !== "closed",
      ),
    [tasks, issueStateById],
  );
  const todayTasks = planningTasks.filter((task) => task.date === todayKey);
  const todayPlanned = todayTasks.reduce((sum, task) => sum + task.estimate, 0);
  const todaySpent = logs
    .filter((log) => log.dateKey === todayKey)
    .reduce((sum, log) => sum + log.hours, 0);
  const active = planningTasks.find((task) => task.id === activeTask);
  const projects = useMemo(() => {
    const configured = projectPreferences.filter((project) => project.enabled);
    if (configured.length) return configured;
    return uniqueBy(planningTasks, (task) => String(task.projectId)).map(
      (task) => ({
        projectId: task.projectId,
        client: task.client,
        project: task.project,
        color: task.color,
        enabled: true,
      }),
    );
  }, [projectPreferences, planningTasks]);
  const weeklyDays = visibleDaysFor(todayKey, "week", capacity.workDays);
  const weeklyHours = planningTasks
    .filter((task) => weeklyDays.includes(task.date))
    .reduce((sum, task) => sum + task.estimate, 0);
  const weeklyCapacity = capacity.dailyHours * capacity.workDays.length;
  const overdue = issues.filter(
    (issue) =>
      issue.state === "opened" && issue.dueDate && issue.dueDate < todayKey,
  );
  const withoutEstimate = issues.filter(
    (issue) => issue.state === "opened" && !issue.hasEstimate,
  );
  const overloadedDays = [
    ...new Set(planningTasks.map((task) => task.date)),
  ].filter(
    (day) =>
      planningTasks
        .filter((task) => task.date === day)
        .reduce((sum, task) => sum + task.estimate, 0) > capacity.dailyHours,
  );
  const notificationCount =
    overdue.length +
    withoutEstimate.length +
    overloadedDays.length +
    (config.token ? 0 : 1);

  function notify(message: string) {
    setToast(message);
  }
  function saveConfig() {
    window.localStorage.setItem(
      "work-organizer.gitlab",
      JSON.stringify(config),
    );
    notify("Configuração GitLab guardada neste dispositivo.");
  }
  function saveOutlookConfig() {
    window.localStorage.setItem(
      "work-organizer.outlook",
      JSON.stringify(outlookConfig),
    );
    notify("Configuração Microsoft guardada neste dispositivo.");
  }
  function saveCapacity() {
    if (!capacity.workDays.length) {
      notify("Seleciona pelo menos um dia de trabalho.");
      return;
    }
    setTasks((all) => planTasks(all, todayKey, capacity));
    notify("Capacidade guardada e Timeline recalculada.");
  }
  function updateProjectPreference(
    projectId: number,
    changes: Partial<ProjectPreference>,
  ) {
    const current = projectPreferences.find(
      (project) => project.projectId === projectId,
    );
    if (!current) return;
    const next = { ...current, ...changes };
    setProjectPreferences((all) =>
      all.map((project) => (project.projectId === projectId ? next : project)),
    );
    setIssues((all) =>
      all.map((issue) =>
        issue.projectId === projectId
          ? {
              ...issue,
              client: next.client,
              project: next.project,
              color: next.color,
            }
          : issue,
      ),
    );
    setTasks((all) =>
      all.map((task) =>
        task.projectId === projectId
          ? {
              ...task,
              client: next.client,
              project: next.project,
              color: next.color,
            }
          : task,
      ),
    );
    setLogs((all) =>
      all.map((log) =>
        log.projectId === projectId
          ? { ...log, client: next.client, project: next.project }
          : log,
      ),
    );
  }

  async function callGitLab(
    action: string,
    extra: Record<string, unknown> = {},
  ) {
    const response = await fetch("/api/gitlab", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action, ...config, ...extra }),
    });
    const data = await response.json();
    if (!response.ok || !data.ok)
      throw new Error(data.error ?? "Não foi possível contactar o GitLab.");
    return data;
  }

  async function testConnection() {
    setConnection({ state: "loading", message: "A testar ligação…" });
    try {
      const data = await callGitLab("test");
      setConnection({
        state: "ok",
        message: `Ligação válida — ${data.user?.name ?? data.user?.username ?? "utilizador"}`,
      });
      saveConfig();
    } catch (error) {
      setConnection({
        state: "error",
        message: error instanceof Error ? error.message : "Ligação falhou.",
      });
    }
  }

  async function syncGitLab() {
    if (persistenceStatus === "loading") {
      notify("Aguarda enquanto o planeamento guardado é aberto.");
      return;
    }
    if (!config.token) {
      setView("settings");
      setSettingsSection("gitlab");
      notify("Configura primeiro o token do GitLab.");
      return;
    }
    setSyncing(true);
    try {
      const data = await callGitLab("sync");
      const projectMap = new Map<number, GitLabProject>(
        ((data.projects ?? []) as GitLabProject[]).map((project) => [
          project.id,
          project,
        ]),
      );
      const remoteIssues = (data.issues ?? []) as GitLabIssue[];
      const syncedLabelCatalog = Object.fromEntries(
        Object.entries(
          (data.labelCatalog ?? {}) as Record<string, GitLabLabelResponse[]>,
        ).map(([projectId, labels]) => [
          projectId,
          labels
            .map(normaliseLabel)
            .sort((a, b) => a.name.localeCompare(b.name)),
        ]),
      );
      const savedPreferences = new Map(
        projectPreferences.map((project) => [project.projectId, project]),
      );
      const nextPreferences: ProjectPreference[] = [...projectMap.values()].map(
        (project) =>
          savedPreferences.get(project.id) ?? {
            projectId: project.id,
            project:
              project.name ??
              project.path_with_namespace ??
              `Projeto ${project.id}`,
            client: projectClient(project),
            color: projectColor(project.id),
            enabled: true,
          },
      );
      const preferenceMap = new Map(
        nextPreferences.map((project) => [project.projectId, project]),
      );
      const syncedIssues = remoteIssues.filter(
        (issue) => preferenceMap.get(issue.project_id)?.enabled !== false,
      );
      const timelogs = (data.timelogs ?? []) as GitLabTimelog[];
      const personalSpentByIssue = new Map<string, number>();
      timelogs.forEach((entry) => {
        const key = `gitlab-${entry.project_id}-${entry.issue_iid}`;
        personalSpentByIssue.set(
          key,
          (personalSpentByIssue.get(key) ?? 0) + entry.time_spent / 3600,
        );
      });
      const nextIssues: IssueRecord[] = syncedIssues.map((issue) => {
        const project = projectMap.get(issue.project_id);
        const preference = preferenceMap.get(issue.project_id);
        const id = `gitlab-${issue.project_id}-${issue.iid}`;
        const rawEstimate = (issue.time_stats?.time_estimate ?? 0) / 3600;
        return {
          id,
          projectId: issue.project_id,
          iid: issue.iid,
          title: issue.title,
          client: preference?.client ?? projectClient(project),
          project:
            preference?.project ??
            project?.name ??
            `Projeto ${issue.project_id}`,
          color: preference?.color ?? projectColor(issue.project_id),
          state: issue.state,
          estimateTotal: rawEstimate,
          teamSpent: (issue.time_stats?.total_time_spent ?? 0) / 3600,
          personalSpent: personalSpentByIssue.get(id) ?? 0,
          assigneeCount: issue.assignees?.length ?? 0,
          due: issue.due_date ? formatDate(issue.due_date) : "Sem prazo",
          dueDate: issue.due_date,
          type: taskType(issue.labels),
          priority: taskPriority(issue.labels),
          webUrl: issue.web_url,
          hasEstimate: rawEstimate > 0,
          labels: labelsForIssue(
            issue.labels,
            syncedLabelCatalog[String(issue.project_id)] ??
              labelCatalog[String(issue.project_id)] ??
              [],
          ),
        };
      });
      const issueMap = new Map(nextIssues.map((issue) => [issue.id, issue]));
      const plannedByIssue = new Map<string, number>();
      tasks.forEach((task) =>
        plannedByIssue.set(
          task.rootId,
          (plannedByIssue.get(task.rootId) ?? 0) + task.estimate,
        ),
      );
      const importedLogs: WorkLog[] = timelogs.flatMap((entry) => {
        const issue = issueMap.get(
          `gitlab-${entry.project_id}-${entry.issue_iid}`,
        );
        if (!issue) return [];
        return {
          id: `gitlab-timelog-${entry.id}`,
          dateKey: entry.spent_at.slice(0, 10),
          client: issue.client,
          project: issue.project,
          task: entry.summary?.trim() || entry.issue_title || issue.title,
          type: issue.type,
          hours: entry.time_spent / 3600,
          estimate: plannedByIssue.get(issue.id) ?? 0,
          webUrl: entry.issue_web_url ?? issue.webUrl,
          projectId: issue.projectId,
          issueIid: issue.iid,
          source: "gitlab" as const,
        };
      });
      setIssues(nextIssues);
      setTasks((current) =>
        current.map((task) => {
          const issue = issueMap.get(task.rootId);
          return issue
            ? {
                ...task,
                title: issue.title,
                client: issue.client,
                project: issue.project,
                color: issue.color,
                due: issue.due,
                dueDate: issue.dueDate,
                priority: issue.priority,
                webUrl: issue.webUrl,
                hasEstimate: issue.hasEstimate,
                labels: issue.labels,
              }
            : task;
        }),
      );
      setLogs((current) => {
        const local = current.filter(
          (log) =>
            log.source !== "gitlab" &&
            !log.id.startsWith("gitlab-timelog-") &&
            !(
              data.timelogsAvailable &&
              log.source === "timer" &&
              issueMap.has(`gitlab-${log.projectId}-${log.issueIid}`)
            ),
        );
        return [...importedLogs, ...local];
      });
      setProjectPreferences(nextPreferences);
      setLabelCatalog((current) => ({ ...current, ...syncedLabelCatalog }));
      saveConfig();
      const openedCount = nextIssues.filter(
        (issue) => issue.state === "opened",
      ).length;
      const suffix = data.timelogWarning
        ? " As horas pessoais não ficaram disponíveis nesta versão do GitLab."
        : "";
      notify(
        `${openedCount} US abertas disponíveis por planear, ${nextIssues.length - openedCount} concluídas e ${importedLogs.length} registos pessoais sincronizados.${suffix}`,
      );
    } catch (error) {
      notify(
        error instanceof Error ? error.message : "A sincronização falhou.",
      );
    } finally {
      setSyncing(false);
    }
  }

  async function loadProjectLabels(projectId: number) {
    if (!config.token)
      throw new Error("Configura primeiro a ligação ao GitLab.");
    const data = await callGitLab("projectLabels", { projectId });
    const labels = ((data.labels ?? []) as GitLabLabelResponse[])
      .map(normaliseLabel)
      .sort((a, b) => a.name.localeCompare(b.name));
    setLabelCatalog((current) => ({ ...current, [String(projectId)]: labels }));
    return labels;
  }

  async function getMsal() {
    if (!outlookConfig.clientId || !outlookConfig.tenantId)
      throw new Error(
        "Indica o Tenant ID e o Client ID da aplicação Microsoft.",
      );
    if (!msalRef.current) {
      msalRef.current = new PublicClientApplication({
        auth: {
          clientId: outlookConfig.clientId.trim(),
          authority: `https://login.microsoftonline.com/${outlookConfig.tenantId.trim()}`,
          redirectUri: window.location.origin,
        },
        cache: { cacheLocation: "sessionStorage" },
      });
      await msalRef.current.initialize();
    }
    return msalRef.current;
  }

  async function connectOutlook() {
    setOutlookConnection({
      state: "loading",
      message: "A abrir autenticação Microsoft…",
    });
    try {
      msalRef.current = null;
      const msal = await getMsal();
      const result = await msal.loginPopup({
        scopes: ["User.Read", "Calendars.Read"],
      });
      window.localStorage.setItem(
        "work-organizer.outlook",
        JSON.stringify(outlookConfig),
      );
      setOutlookConnection({
        state: "ok",
        message: `Ligado a ${result.account?.username ?? "Microsoft 365"}`,
      });
    } catch (error) {
      setOutlookConnection({
        state: "error",
        message:
          error instanceof Error
            ? error.message
            : "Não foi possível ligar ao Microsoft 365.",
      });
    }
  }

  function replaceImportedMeetings(events: CalendarMeeting[]) {
    const previous = new Map(meetings.map((meeting) => [meeting.id, meeting]));
    const imported = events.map((event) => {
      const saved = previous.get(event.id);
      return {
        ...event,
        projectId: saved?.projectId,
        linkedIssueId: saved?.linkedIssueId,
        linkedIssueIid: saved?.linkedIssueIid,
        linkedIssueWebUrl: saved?.linkedIssueWebUrl,
      };
    });
    const importedRootIds = new Set(
      imported.map((meeting) => `outlook-${meeting.id}`),
    );
    const importedMeetingIds = new Set(
      imported.map((meeting) => meeting.id),
    );
    const importEnd = addDays(todayKey, 36);
    const isInImportWindow = (value: string) =>
      value >= todayKey && value < importEnd;
    const removedLinkedIssueIds = new Set(
      meetings
        .filter(
          (meeting) =>
            isInImportWindow(meeting.start.slice(0, 10)) &&
            !importedMeetingIds.has(meeting.id) &&
            meeting.linkedIssueId,
        )
        .map((meeting) => meeting.linkedIssueId!),
    );
    const meetingTasks = imported.flatMap((meeting): Task[] => {
      if (meeting.linkedIssueId) return [];
      const project = projects.find(
        (item) => item.projectId === meeting.projectId,
      );
      const start = new Date(meeting.start);
      const end = new Date(meeting.end);
      return [
        {
          id: `outlook-${meeting.id}`,
          rootId: `outlook-${meeting.id}`,
          allocationId: `outlook-${meeting.id}`,
          projectId: project?.projectId ?? 0,
          iid: 0,
          title: meeting.subject,
          client: project?.client ?? "Outlook",
          project: project?.project ?? "Reuniões",
          color: project?.color ?? "#3977d5",
          date: meeting.start.slice(0, 10),
          start: start.getHours() + start.getMinutes() / 60,
          estimate: Math.max(
            0.25,
            (end.getTime() - start.getTime()) / 3_600_000,
          ),
          spent: 0,
          due: "Reunião",
          dueDate: meeting.start.slice(0, 10),
          type: "Reunião",
          phase: "Reunião",
          priority: "Normal",
          fixed: true,
          webUrl: meeting.joinUrl ?? meeting.webUrl,
          source: "outlook",
        },
      ];
    });
    setMeetings(imported);
    setTasks((current) => [
      ...current.filter((task) => {
        if (importedRootIds.has(task.rootId)) return false;
        if (task.source === "outlook" && isInImportWindow(task.date))
          return false;
        if (
          removedLinkedIssueIds.has(task.rootId) &&
          task.fixed &&
          task.phase === "Reunião"
        )
          return false;
        return true;
      }),
      ...meetingTasks,
    ]);
    return imported;
  }

  async function syncOutlook() {
    if (persistenceStatus === "loading") {
      notify("Aguarda enquanto o planeamento guardado é aberto.");
      return;
    }
    setSyncingCalendar(true);
    try {
      const msal = await getMsal();
      const account = msal.getAllAccounts()[0];
      const auth = account
        ? await msal
            .acquireTokenSilent({
              account,
              scopes: ["User.Read", "Calendars.Read"],
            })
            .catch(() =>
              msal.acquireTokenPopup({
                account,
                scopes: ["User.Read", "Calendars.Read"],
              }),
            )
        : await msal.loginPopup({ scopes: ["User.Read", "Calendars.Read"] });
      const start = new Date();
      const end = new Date();
      end.setDate(end.getDate() + 35);
      const params = new URLSearchParams({
        startDateTime: start.toISOString(),
        endDateTime: end.toISOString(),
        $orderby: "start/dateTime",
        $top: "200",
        $select:
          "id,subject,start,end,webLink,isOnlineMeeting,onlineMeeting,isCancelled,responseStatus",
      });
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/calendarView?${params}`,
        {
          headers: {
            Authorization: `Bearer ${auth.accessToken}`,
            Prefer: 'outlook.timezone="Europe/Lisbon"',
          },
        },
      );
      const data = await response.json();
      if (!response.ok)
        throw new Error(
          data.error?.message ?? "Não foi possível ler o calendário.",
        );
      const outlookEvents = (data.value ?? []) as Array<{
          id: string;
          subject?: string;
          start: { dateTime: string };
          end: { dateTime: string };
          webLink?: string;
          onlineMeeting?: { joinUrl?: string };
          isCancelled?: boolean;
          responseStatus?: { response?: string };
        }>;
      const events = outlookEvents
        .filter(
          (event) =>
            !event.isCancelled &&
            ["accepted", "organizer"].includes(
              event.responseStatus?.response ?? "",
            ),
        )
        .map((event) => ({
          id: event.id,
          subject: event.subject || "Reunião sem título",
          start: event.start.dateTime,
          end: event.end.dateTime,
          webUrl: event.webLink,
          joinUrl: event.onlineMeeting?.joinUrl,
        }));
      const imported = replaceImportedMeetings(events);
      setOutlookConnection({
        state: "ok",
        message: `${imported.length} reuniões importadas para os próximos 35 dias.`,
      });
      notify(
        `${imported.length} reuniões atualizadas a partir do Outlook/Teams.`,
      );
    } catch (error) {
      setOutlookConnection({
        state: "error",
        message:
          error instanceof Error
            ? error.message
            : "A sincronização do calendário falhou.",
      });
    } finally {
      setSyncingCalendar(false);
    }
  }

  async function syncLocalOutlook() {
    if (persistenceStatus === "loading") {
      notify("Aguarda enquanto o planeamento guardado é aberto.");
      return;
    }
    setSyncingCalendar(true);
    setOutlookConnection({
      state: "loading",
      message: "A ler o Outlook clássico instalado…",
    });
    try {
      const response = await fetch("http://127.0.0.1:47831/calendar?days=35", {
        signal: AbortSignal.timeout(30000),
      });
      const data = await response.json();
      if (!response.ok || !data.ok)
        throw new Error(
          data.error ?? "A ponte local do Outlook devolveu um erro.",
        );
      const imported = replaceImportedMeetings(data.events ?? []);
      setOutlookConnection({
        state: "ok",
        message: `${imported.length} reuniões lidas do Outlook clássico.`,
      });
      notify(
        `${imported.length} reuniões atualizadas a partir do Outlook instalado.`,
      );
    } catch (error) {
      const message =
        error instanceof TypeError && error.message === "Failed to fetch"
          ? "Não foi possível contactar a ponte local. Autoriza o acesso à rede local se o browser o pedir."
          : error instanceof Error
          ? error.message
          : "Não foi possível contactar a ponte local.";
      setOutlookConnection({
        state: "error",
        message: `${message} Confirma que executaste npm run outlook:bridge.`,
      });
    } finally {
      setSyncingCalendar(false);
    }
  }

  function associateMeeting(meetingId: string, projectId: number | undefined) {
    const meeting = meetings.find((item) => item.id === meetingId);
    setMeetings((all) =>
      all.map((item) =>
        item.id === meetingId ? { ...item, projectId } : item,
      ),
    );
    setTasks((all) => {
      const without = all.filter(
        (task) => task.rootId !== `outlook-${meetingId}`,
      );
      if (!meeting || !projectId || meeting.linkedIssueId) return without;
      const project = projects.find((item) => item.projectId === projectId);
      if (!project) return without;
      const start = new Date(meeting.start);
      const end = new Date(meeting.end);
      const estimate = Math.max(
        0.25,
        (end.getTime() - start.getTime()) / 3_600_000,
      );
      const task: Task = {
        id: `outlook-${meetingId}`,
        rootId: `outlook-${meetingId}`,
        allocationId: `outlook-${meetingId}`,
        projectId,
        iid: 0,
        title: meeting.subject,
        client: project.client,
        project: project.project,
        color: project.color,
        date: meeting.start.slice(0, 10),
        start: start.getHours() + start.getMinutes() / 60,
        estimate,
        spent: 0,
        due: "Reunião",
        dueDate: meeting.start.slice(0, 10),
        type: "Reunião",
        phase: "Reunião",
        priority: "Normal",
        fixed: true,
        webUrl: meeting.joinUrl ?? meeting.webUrl,
        source: "outlook",
      };
      return [...without, task];
    });
    notify(
      projectId
        ? "Reunião associada ao projeto e adicionada à Timeline."
        : "Associação da reunião removida.",
    );
  }

  function saveAllocationForIssue(issue: IssueRecord, draft: AllocationDraft) {
    const allocationId =
      draft.allocationId ??
      `allocation-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
    const hours = Math.max(0.25, Math.round(draft.hours * 4) / 4);
    setTasks((current) => {
      const withoutCurrent = current.filter(
        (task) => task.allocationId !== allocationId,
      );
      const base: Omit<Task, "id" | "date" | "start" | "estimate" | "spent"> = {
        rootId: issue.id,
        allocationId,
        projectId: issue.projectId,
        iid: issue.iid,
        title: issue.title,
        client: issue.client,
        project: issue.project,
        color: issue.color,
        due: issue.due,
        dueDate: issue.dueDate,
        type: draft.phase,
        phase: draft.phase,
        description: draft.description.trim(),
        priority: issue.priority,
        fixed: false,
        webUrl: issue.webUrl,
        source: "gitlab",
        hasEstimate: issue.hasEstimate,
        labels: issue.labels,
      };
      if (draft.distribution === "manual") {
        return [
          ...withoutCurrent,
          {
            ...base,
            id: `${allocationId}::0`,
            date: draft.date,
            start: draft.start,
            estimate: hours,
            spent: 0,
          },
        ];
      }
      const used = new Map<string, number>();
      withoutCurrent.forEach((task) =>
        used.set(task.date, (used.get(task.date) ?? 0) + task.estimate),
      );
      const segments: Task[] = [];
      let remaining = hours;
      let cursor = nextBusinessDay(draft.date, true, capacity.workDays);
      let segment = 0;
      while (remaining > 0.001 && segment < 500) {
        const occupied = used.get(cursor) ?? 0;
        const available = Math.max(0, capacity.dailyHours - occupied);
        if (available < 0.01) {
          cursor = nextBusinessDay(cursor, false, capacity.workDays);
          continue;
        }
        const segmentHours = Math.min(remaining, available);
        segments.push({
          ...base,
          id: `${allocationId}::${segment}`,
          date: cursor,
          start: capacity.startHour + occupied,
          estimate: segmentHours,
          spent: 0,
        });
        used.set(cursor, occupied + segmentHours);
        remaining -= segmentHours;
        segment += 1;
        if (remaining > 0.001)
          cursor = nextBusinessDay(cursor, false, capacity.workDays);
      }
      return [...withoutCurrent, ...segments];
    });
    notify(
      draft.allocationId
        ? "Alocação atualizada."
        : "US adicionada ao teu planeamento.",
    );
  }

  function saveAllocation(draft: AllocationDraft) {
    const issue = issues.find((item) => item.id === draft.issueId);
    if (!issue) {
      notify("A US já não está disponível. Sincroniza novamente o GitLab.");
      return;
    }
    if (!draft.description.trim()) {
      notify("Descreve o que vais fazer nesta alocação.");
      return;
    }
    saveAllocationForIssue(issue, draft);
  }

  async function createIssueAndAllocate(draft: CreateIssueDraft) {
    if (!config.token)
      throw new Error("Configura primeiro a ligação ao GitLab.");
    if (!draft.title.trim() || !draft.description.trim())
      throw new Error("Indica o título e descreve o trabalho a realizar.");
    const data = await callGitLab("createIssue", {
      projectId: draft.projectId,
      title: draft.title,
      description: draft.description,
      estimateHours: draft.estimateHours,
      labels: draft.labels,
    });
    const remote = data.issue as GitLabIssue;
    const remoteProject = data.project as GitLabProject;
    const preference = projectPreferences.find(
      (project) => project.projectId === draft.projectId,
    );
    const rawEstimate =
      (remote.time_stats?.time_estimate ??
        Math.round(draft.estimateHours * 3600)) / 3600;
    const issue: IssueRecord = {
      id: `gitlab-${remote.project_id}-${remote.iid}`,
      projectId: remote.project_id,
      iid: remote.iid,
      title: remote.title,
      client: preference?.client ?? projectClient(remoteProject),
      project:
        preference?.project ??
        remoteProject.name ??
        `Projeto ${remote.project_id}`,
      color: preference?.color ?? projectColor(remote.project_id),
      state: remote.state,
      estimateTotal: rawEstimate,
      teamSpent: (remote.time_stats?.total_time_spent ?? 0) / 3600,
      personalSpent: 0,
      assigneeCount: remote.assignees?.length ?? 1,
      due: remote.due_date ? formatDate(remote.due_date) : "Sem prazo",
      dueDate: remote.due_date,
      type: taskType(remote.labels),
      priority: taskPriority(remote.labels),
      webUrl: remote.web_url,
      hasEstimate: rawEstimate > 0,
      labels: labelsForIssue(
        remote.labels ?? draft.labels,
        labelCatalog[String(draft.projectId)] ?? [],
      ),
    };
    setIssues((current) => [
      issue,
      ...current.filter((item) => item.id !== issue.id),
    ]);
    if (draft.meetingId) {
      const meetingRootId = `outlook-${draft.meetingId}`;
      setMeetings((current) =>
        current.map((meeting) =>
          meeting.id === draft.meetingId
            ? {
                ...meeting,
                projectId: issue.projectId,
                linkedIssueId: issue.id,
                linkedIssueIid: issue.iid,
                linkedIssueWebUrl: issue.webUrl,
              }
            : meeting,
        ),
      );
      setTasks((current) =>
        current.map((task) =>
          task.rootId === meetingRootId
            ? {
                ...task,
                rootId: issue.id,
                projectId: issue.projectId,
                iid: issue.iid,
                title: issue.title,
                client: issue.client,
                project: issue.project,
                color: issue.color,
                due: issue.due,
                dueDate: issue.dueDate,
                type: "Reunião",
                phase: "Reunião",
                description: draft.description.trim(),
                priority: issue.priority,
                fixed: true,
                webUrl: issue.webUrl,
                source: "gitlab",
                hasEstimate: issue.hasEstimate,
                labels: issue.labels,
              }
            : task,
        ),
      );
      notify(
        `US #${issue.iid} criada para a reunião. O cronómetro passa a registar spent no GitLab.`,
      );
    } else {
      saveAllocationForIssue(issue, {
        issueId: issue.id,
        phase: draft.phase,
        hours: draft.hours,
        date: draft.date,
        start: draft.start,
        distribution: draft.distribution,
        description: draft.description,
      });
      notify(`Tarefa #${issue.iid} criada no GitLab e adicionada à Timeline.`);
    }
  }

  function deleteAllocation(allocationId: string) {
    setTasks((current) =>
      current.filter((task) => task.allocationId !== allocationId),
    );
    notify("Alocação removida da Timeline. A US continua no backlog.");
  }

  function startTimer(id: string) {
    if (activeTask !== id) {
      setActiveTask(id);
      setElapsed(0);
    }
  }
  async function stopTimer() {
    if (!active || elapsed < 1) {
      setActiveTask(null);
      setElapsed(0);
      return;
    }
    const minutes = Math.max(1, Math.round(elapsed / 60));
    try {
      if (config.token && active.source === "gitlab")
        await callGitLab("addSpent", {
          projectId: active.projectId,
          issueIid: active.iid,
          duration: `${minutes}m`,
        });
      const hours = minutes / 60;
      setTasks((all) =>
        all.map((task) =>
          task.id === active.id ? { ...task, spent: task.spent + hours } : task,
        ),
      );
      setLogs((all) => [
        {
          id: `${Date.now()}`,
          dateKey: todayKey,
          client: active.client,
          project: active.project,
          task: active.title,
          type: active.type,
          hours,
          estimate: active.estimate,
          webUrl: active.webUrl,
          projectId: active.projectId,
          issueIid: active.iid,
          source: "timer",
        },
        ...all,
      ]);
      notify(
        config.token && active.source === "gitlab"
          ? `${minutes} min registados no GitLab.`
          : `${minutes} min registados localmente.`,
      );
    } catch (error) {
      notify(
        error instanceof Error
          ? error.message
          : "Não foi possível registar o tempo.",
      );
      return;
    }
    setActiveTask(null);
    setElapsed(0);
  }

  function moveTask(
    allocationId: string,
    targetDate: string,
    targetStart?: number,
  ) {
    setTasks((all) => {
      const rootTasks = all
        .filter((task) => task.allocationId === allocationId)
        .sort((a, b) => a.date.localeCompare(b.date) || a.start - b.start);
      const startDelta =
        targetStart == null || !rootTasks.length
          ? 0
          : targetStart - rootTasks[0].start;
      const sourceDates = [...new Set(rootTasks.map((task) => task.date))];
      const targetBySource = new Map(
        sourceDates.map((sourceDate, index) => [
          sourceDate,
          addBusinessDays(targetDate, index, capacity.workDays),
        ]),
      );
      return all.map((task) =>
        task.allocationId === allocationId
          ? {
              ...task,
              date: targetBySource.get(task.date) ?? targetDate,
              start: Math.max(0, Math.min(23.75, task.start + startDelta)),
            }
          : task,
      );
    });
    setDragged(null);
    notify("Planeamento manual guardado.");
  }
  function autoPlan() {
    setTasks((all) => planTasks(all, todayKey, capacity));
    notify(
      "Timeline recalculada. A carga continua automaticamente nas semanas seguintes.",
    );
  }

  const searchNeedle = globalSearch.trim().toLowerCase();
  const searchIssues = searchNeedle
    ? issues
        .filter(
          (issue) =>
            issue.state === "opened" &&
            `${issue.title} ${issue.client} ${issue.project} ${issue.iid} ${(issue.labels ?? []).map((label) => label.name).join(" ")}`
              .toLowerCase()
              .includes(searchNeedle),
        )
        .slice(0, 8)
    : [];
  const searchLogs = searchNeedle
    ? logs
        .filter((log) =>
          `${log.task} ${log.client} ${log.project}`
            .toLowerCase()
            .includes(searchNeedle),
        )
        .slice(0, 5)
    : [];

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <span className="brand-mark">W</span>
          <div>
            <strong>Work Organizer</strong>
            <small>Personal capacity</small>
          </div>
        </div>
        <nav className="main-nav" aria-label="Navegação principal">
          <span className="nav-label">ESPAÇO DE TRABALHO</span>
          {nav.map((item) => (
            <button
              key={item.id}
              className={view === item.id ? "active" : ""}
              onClick={() => setView(item.id)}
            >
              <span>{item.icon}</span>
              {item.label}
              {item.id === "today" && <b>{todayTasks.length}</b>}
            </button>
          ))}
          <span className="nav-label nav-section">SISTEMA</span>
          <button
            className={view === "settings" ? "active" : ""}
            onClick={() => setView("settings")}
          >
            <span>⚙</span>Definições
            {config.token ? (
              <i className="connection-dot ok" />
            ) : (
              <i className="connection-dot" />
            )}
          </button>
        </nav>
        <div className="sidebar-card">
          <div className="sidebar-card-head">
            <span>CAPACIDADE SEMANAL</span>
            <strong>
              {formatHours(weeklyHours)} / {formatHours(weeklyCapacity)}
            </strong>
          </div>
          <div className="mini-progress">
            <i
              style={{
                width: `${Math.min(100, weeklyCapacity ? (weeklyHours / weeklyCapacity) * 100 : 0)}%`,
              }}
            />
          </div>
          <small>Planeamento distribuído por capacidade</small>
        </div>
        <div className="profile">
          <div className="avatar">
            {user.displayName
              .split(/\s+/)
              .filter(Boolean)
              .slice(0, 2)
              .map((part) => part[0])
              .join("")
              .toUpperCase() || "NM"}
          </div>
          <div>
            <strong>{user.displayName}</strong>
            <small>
              {persistenceStatus === "ready"
                ? "Dados guardados na BD"
                : persistenceStatus === "loading"
                  ? "A abrir dados…"
                  : "Guardado neste browser"}
            </small>
          </div>
          <button
            aria-label="Abrir definições"
            onClick={() => setView("settings")}
          >
            •••
          </button>
        </div>
      </aside>
      <main className="workspace">
        <header className="topbar">
          <div className="breadcrumb">
            <span>Work Organizer</span>
            <b>/</b>
            <strong>
              {view === "today"
                ? "Hoje"
                : view === "timeline"
                  ? "Timeline"
                  : view === "completed"
                    ? "Trabalho realizado"
                    : view === "analytics"
                      ? "Analytics"
                      : "Definições"}
            </strong>
          </div>
          <div className="top-actions">
            <button
              className="icon-button"
              aria-label="Pesquisar"
              aria-expanded={searchOpen}
              onClick={() => {
                setSearchOpen(!searchOpen);
                setNotificationsOpen(false);
              }}
            >
              ⌕
            </button>
            <button
              className="icon-button notification"
              aria-label="Notificações"
              aria-expanded={notificationsOpen}
              onClick={() => {
                setNotificationsOpen(!notificationsOpen);
                setSearchOpen(false);
              }}
            >
              ♢
              {notificationCount > 0 && (
                <i>
                  <span>{Math.min(9, notificationCount)}</span>
                </i>
              )}
            </button>
            <button
              className="sync-button"
              onClick={syncGitLab}
              disabled={syncing}
            >
              <span className={syncing ? "spin" : ""}>↻</span>
              {syncing ? "A sincronizar…" : "Sincronizar GitLab"}
            </button>
            {searchOpen && (
              <div className="top-popover search-popover">
                <div className="popover-title">
                  <strong>Pesquisar trabalho</strong>
                  <button
                    onClick={() => setSearchOpen(false)}
                    aria-label="Fechar"
                  >
                    ×
                  </button>
                </div>
                <div className="popover-search">
                  ⌕
                  <input
                    autoFocus
                    value={globalSearch}
                    onChange={(event) => setGlobalSearch(event.target.value)}
                    placeholder="Issue, cliente ou projeto…"
                  />
                </div>
                <div className="popover-results">
                  {!searchNeedle ? (
                    <p>
                      Escreve para pesquisar nas Issues sincronizadas e no
                      trabalho realizado.
                    </p>
                  ) : (
                    <>
                      {searchIssues.map((issue) => (
                        <button
                          key={issue.id}
                          onClick={() => {
                            setView("timeline");
                            setSearchOpen(false);
                          }}
                        >
                          <span style={{ background: issue.color }} />
                          <div>
                            <strong>{issue.title}</strong>
                            <small>
                              {issue.client} · {issue.project} · #{issue.iid}
                            </small>
                          </div>
                        </button>
                      ))}
                      {searchLogs.map((log) => (
                        <button
                          key={log.id}
                          onClick={() => {
                            setView("completed");
                            setSearchOpen(false);
                          }}
                        >
                          <span className="result-done">✓</span>
                          <div>
                            <strong>{log.task}</strong>
                            <small>
                              {log.client} · {formatHours(log.hours)} realizadas
                            </small>
                          </div>
                        </button>
                      ))}
                      {!searchIssues.length && !searchLogs.length && (
                        <p>Não encontrei resultados reais sincronizados.</p>
                      )}
                    </>
                  )}
                </div>
              </div>
            )}
            {notificationsOpen && (
              <div className="top-popover notifications-popover">
                <div className="popover-title">
                  <strong>Notificações</strong>
                  <button
                    onClick={() => setNotificationsOpen(false)}
                    aria-label="Fechar"
                  >
                    ×
                  </button>
                </div>
                <div className="notification-list">
                  {!config.token && (
                    <button
                      onClick={() => {
                        setView("settings");
                        setSettingsSection("gitlab");
                        setNotificationsOpen(false);
                      }}
                    >
                      <b>!</b>
                      <div>
                        <strong>GitLab por configurar</strong>
                        <small>
                          Adiciona o token para sincronizar dados reais.
                        </small>
                      </div>
                    </button>
                  )}
                  {overdue.length > 0 && (
                    <button
                      onClick={() => {
                        setView("timeline");
                        setNotificationsOpen(false);
                      }}
                    >
                      <b>!</b>
                      <div>
                        <strong>
                          {overdue.length}{" "}
                          {overdue.length === 1
                            ? "Issue atrasada"
                            : "Issues atrasadas"}
                        </strong>
                        <small>O prazo já terminou e continuam abertas.</small>
                      </div>
                    </button>
                  )}
                  {withoutEstimate.length > 0 && (
                    <button
                      onClick={() => {
                        setView("timeline");
                        setNotificationsOpen(false);
                      }}
                    >
                      <b>i</b>
                      <div>
                        <strong>{withoutEstimate.length} sem estimativa</strong>
                        <small>
                          Define manualmente as horas da tua participação.
                        </small>
                      </div>
                    </button>
                  )}
                  {overloadedDays.length > 0 && (
                    <button
                      onClick={() => {
                        setView("timeline");
                        setNotificationsOpen(false);
                      }}
                    >
                      <b>↑</b>
                      <div>
                        <strong>
                          {overloadedDays.length}{" "}
                          {overloadedDays.length === 1
                            ? "dia sobrecarregado"
                            : "dias sobrecarregados"}
                        </strong>
                        <small>Revê tarefas fixas ou volta a planear.</small>
                      </div>
                    </button>
                  )}
                  {notificationCount === 0 && <p>Sem alertas neste momento.</p>}
                </div>
              </div>
            )}
          </div>
        </header>
        {view === "today" && (
          <TodayView
            tasks={todayTasks}
            planned={todayPlanned}
            spent={todaySpent}
            dailyHours={capacity.dailyHours}
            todayKey={todayKey}
            activeTask={activeTask}
            elapsed={elapsed}
            startTimer={startTimer}
            stopTimer={stopTimer}
            setView={setView}
          />
        )}
        {view === "timeline" && (
          <TimelineView
            issues={issues}
            tasks={planningTasks}
            gitlabProjects={projects}
            labelCatalog={labelCatalog}
            loadProjectLabels={loadProjectLabels}
            dragged={dragged}
            setDragged={setDragged}
            moveTask={moveTask}
            saveAllocation={saveAllocation}
            createIssueAndAllocate={createIssueAndAllocate}
            deleteAllocation={deleteAllocation}
            autoPlan={autoPlan}
            todayKey={todayKey}
            capacity={capacity}
          />
        )}
        {view === "completed" && (
          <CompletedView logs={logs} todayKey={todayKey} />
        )}
        {view === "analytics" && (
          <AnalyticsView
            issues={issues}
            logs={logs}
            tasks={tasks}
            todayKey={todayKey}
          />
        )}
        {view === "settings" && (
          <SettingsView
            section={settingsSection}
            setSection={setSettingsSection}
            config={config}
            setConfig={setConfig}
            showToken={showToken}
            setShowToken={setShowToken}
            connection={connection}
            saveConfig={saveConfig}
            testConnection={testConnection}
            outlookConfig={outlookConfig}
            setOutlookConfig={setOutlookConfig}
            outlookConnection={outlookConnection}
            saveOutlookConfig={saveOutlookConfig}
            connectOutlook={connectOutlook}
            syncOutlook={syncOutlook}
            syncLocalOutlook={syncLocalOutlook}
            syncingCalendar={syncingCalendar}
            meetings={meetings}
            projects={projects}
            associateMeeting={associateMeeting}
            capacity={capacity}
            setCapacity={setCapacity}
            saveCapacity={saveCapacity}
            projectPreferences={projectPreferences}
            updateProjectPreference={updateProjectPreference}
            syncGitLab={syncGitLab}
            syncing={syncing}
          />
        )}
      </main>
      {toast && (
        <div className="toast">
          <span>✓</span>
          {toast}
        </div>
      )}
    </div>
  );
}

function PageIntro({
  eyebrow,
  title,
  description,
  children,
}: {
  eyebrow: string;
  title: string;
  description: string;
  children?: React.ReactNode;
}) {
  return (
    <div className="page-intro">
      <div>
        <span>{eyebrow}</span>
        <h1>{title}</h1>
        <p>{description}</p>
      </div>
      {children && <div className="intro-actions">{children}</div>}
    </div>
  );
}

function IssueTitle({
  task,
  compact = false,
}: {
  task: Task;
  compact?: boolean;
}) {
  if (!task.webUrl)
    return compact ? <strong>{task.title}</strong> : <h3>{task.title}</h3>;
  return compact ? (
    <a
      className="issue-link compact"
      href={task.webUrl}
      target="_blank"
      rel="noreferrer"
      onClick={(event) => event.stopPropagation()}
    >
      <strong>{task.title}</strong>
      <span>↗</span>
    </a>
  ) : (
    <h3>
      <a
        className="issue-link"
        href={task.webUrl}
        target="_blank"
        rel="noreferrer"
      >
        {task.title}
        <span>↗</span>
      </a>
    </h3>
  );
}

function IssueLabels({
  labels,
  compact = false,
}: {
  labels?: GitLabLabel[];
  compact?: boolean;
}) {
  if (!labels?.length) return null;
  return (
    <div className={`issue-labels ${compact ? "compact" : ""}`}>
      {labels.map((label) => (
        <span
          key={label.name}
          title={label.description || label.name}
          style={{
            backgroundColor: label.color,
            color: label.textColor,
            borderColor: label.color,
          }}
        >
          {label.name}
        </span>
      ))}
    </div>
  );
}

function TodayView({
  tasks,
  planned,
  spent,
  dailyHours,
  todayKey,
  activeTask,
  elapsed,
  startTimer,
  stopTimer,
  setView,
}: {
  tasks: Task[];
  planned: number;
  spent: number;
  dailyHours: number;
  todayKey: string;
  activeTask: string | null;
  elapsed: number;
  startTimer: (id: string) => void;
  stopTimer: () => void;
  setView: (view: View) => void;
}) {
  const remaining = Math.max(0, dailyHours - spent);
  return (
    <div className="page">
      <PageIntro
        eyebrow={`${WEEKDAYS[fromKey(todayKey).getDay()]} · ${formatDate(todayKey, true).toUpperCase()}`}
        title="Bom dia, Nuno."
        description={
          planned >= dailyHours
            ? "Tens o dia totalmente planeado."
            : `Ainda tens ${formatHours(dailyHours - planned)} de capacidade disponível.`
        }
      >
        <button
          className="secondary-button"
          onClick={() => setView("timeline")}
        >
          Ver timeline <span>→</span>
        </button>
      </PageIntro>
      <section className="metrics-grid">
        <article>
          <div className="metric-top">
            <span className="metric-icon purple">◷</span>
            <small>PLANEADO HOJE</small>
          </div>
          <strong>{formatHours(planned)}</strong>
          <p>
            {planned > dailyHours
              ? "Capacidade excedida"
              : `Capacidade diária de ${formatHours(dailyHours)}`}
          </p>
        </article>
        <article>
          <div className="metric-top">
            <span className="metric-icon green">✓</span>
            <small>REALIZADO</small>
          </div>
          <strong>{formatHours(spent)}</strong>
          <div className="metric-progress">
            <i
              style={{ width: `${Math.min(100, (spent / dailyHours) * 100)}%` }}
            />
          </div>
          <p>{Math.round((spent / dailyHours) * 100)}% da meta diária</p>
        </article>
        <article>
          <div className="metric-top">
            <span className="metric-icon orange">◌</span>
            <small>POR REGISTAR</small>
          </div>
          <strong>{formatHours(remaining)}</strong>
          <p className="warning-text">
            Para completar {formatHours(dailyHours)}
          </p>
        </article>
        <article>
          <div className="metric-top">
            <span className="metric-icon blue">◎</span>
            <small>TAREFAS</small>
          </div>
          <strong>{tasks.length}</strong>
          <p>
            <i className="done-dot" />{" "}
            {tasks.filter((task) => task.spent >= task.estimate).length}{" "}
            concluídas
          </p>
        </article>
      </section>
      <section className="today-layout">
        <div className="panel schedule-panel">
          <div className="panel-head">
            <div>
              <h2>Plano de hoje</h2>
              <p>
                {formatHours(planned)} distribuídas por{" "}
                {new Set(tasks.map((task) => task.projectId)).size} projetos
              </p>
            </div>
            <button className="text-button" onClick={() => setView("timeline")}>
              Ajustar plano
            </button>
          </div>
          <div className="schedule-list">
            {tasks.length ? (
              [...tasks]
                .sort((a, b) => a.start - b.start)
                .map((task, index) => {
                  const running = activeTask === task.id;
                  const complete = task.spent >= task.estimate;
                  return (
                    <div
                      className={`schedule-row ${running ? "running" : ""}`}
                      key={task.id}
                    >
                      <div className="time-column">
                        <strong>
                          {String(Math.floor(task.start)).padStart(2, "0")}:
                          {task.start % 1 ? "30" : "00"}
                        </strong>
                        <span>{formatHours(task.estimate)}</span>
                      </div>
                      <div className="schedule-line">
                        <i style={{ background: task.color }} />
                        {index < tasks.length - 1 && <span />}
                      </div>
                      <div className="task-main">
                        <div className="task-meta">
                          <span
                            style={{
                              color: task.color,
                              background: `${task.color}14`,
                            }}
                          >
                            {task.client}
                          </span>
                          <small>{task.iid ? `#${task.iid}` : "Outlook"}</small>
                          {task.fixed && <small>Horário fixo</small>}
                        </div>
                        <IssueTitle task={task} />
                        <IssueLabels labels={task.labels} compact />
                        <p>
                          {task.description
                            ? `${task.description} · ${task.phase}`
                            : `${task.project} · ${task.type}`}
                        </p>
                      </div>
                      <div className="task-state">
                        {running ? (
                          <>
                            <strong className="live-time">
                              {formatTimer(elapsed)}
                            </strong>
                            <button className="stop-button" onClick={stopTimer}>
                              Parar
                            </button>
                          </>
                        ) : task.source === "outlook" ? (
                          <button
                            className="play-button"
                            onClick={() => setView("timeline")}
                          >
                            Criar US
                          </button>
                        ) : complete ? (
                          <span className="complete-pill">✓ Realizada</span>
                        ) : (
                          <button
                            className="play-button"
                            onClick={() => startTimer(task.id)}
                          >
                            <span>▶</span> Iniciar
                          </button>
                        )}
                      </div>
                    </div>
                  );
                })
            ) : (
              <div className="empty-state">
                <span>◷</span>
                <strong>Sem trabalho planeado para hoje</strong>
                <p>Abre a Timeline e planeia uma US do backlog.</p>
              </div>
            )}
          </div>
        </div>
        <aside className="right-column">
          <div className="panel focus-card">
            <span className="focus-label">FOCO ATUAL</span>
            <div className="focus-art">
              <span>◉</span>
            </div>
            <h3>
              {activeTask
                ? tasks.find((task) => task.id === activeTask)?.title
                : "Pronto para começar?"}
            </h3>
            <p>
              {activeTask
                ? "O cronómetro está a registar esta sessão."
                : "Inicia uma tarefa do plano para registar o tempo automaticamente."}
            </p>
            {activeTask ? (
              <button className="primary-button danger" onClick={stopTimer}>
                ■ Parar e registar
              </button>
            ) : (
              <button
                className="primary-button"
                disabled={!tasks[0]}
                onClick={() => tasks[0] && startTimer(tasks[0].id)}
              >
                ▶ Iniciar próxima tarefa
              </button>
            )}
          </div>
          <div className="panel alerts">
            <div className="panel-head compact">
              <h2>Atenção</h2>
              <span>{remaining > 0 ? 1 : 0}</span>
            </div>
            <div className="alert-item">
              <i className="alert-yellow">i</i>
              <div>
                <strong>Tempo por classificar</strong>
                <p>Faltam {formatHours(remaining)} para completar o dia.</p>
              </div>
            </div>
          </div>
        </aside>
      </section>
    </div>
  );
}

type TimelineBar = {
  allocationId: string;
  task: Task;
  left: number;
  width: number;
  lane: number;
  totalHours: number;
  totalDays: number;
};

type AgendaItem = {
  task: Task;
  column: number;
  columnCount: number;
};

function layoutAgendaDay(tasks: Task[]): AgendaItem[] {
  const ordered = [...tasks].sort(
    (a, b) => a.start - b.start || b.estimate - a.estimate,
  );
  const clusters: Task[][] = [];
  let cluster: Task[] = [];
  let clusterEnd = -1;
  ordered.forEach((task) => {
    if (cluster.length && task.start >= clusterEnd - 0.001) {
      clusters.push(cluster);
      cluster = [];
      clusterEnd = -1;
    }
    cluster.push(task);
    clusterEnd = Math.max(clusterEnd, task.start + task.estimate);
  });
  if (cluster.length) clusters.push(cluster);

  return clusters.flatMap((items) => {
    const columnEnds: number[] = [];
    const positioned = items.map((task) => {
      let column = columnEnds.findIndex((end) => end <= task.start + 0.001);
      if (column < 0) column = columnEnds.length;
      columnEnds[column] = task.start + task.estimate;
      return { task, column };
    });
    const columnCount = Math.max(1, columnEnds.length);
    return positioned.map((item) => ({ ...item, columnCount }));
  });
}

function buildTimelineBars(
  tasks: Task[],
  days: string[],
  capacity: CapacityConfig,
): TimelineBar[] {
  if (!days.length) return [];
  const dayIndex = new Map(days.map((day, index) => [day, index]));
  const grouped = new Map<string, Task[]>();
  tasks.forEach((task) =>
    grouped.set(task.allocationId, [
      ...(grouped.get(task.allocationId) ?? []),
      task,
    ]),
  );

  const ranges = [...grouped.entries()]
    .flatMap(([allocationId, rootTasks]) => {
      const ordered = [...rootTasks].sort(
        (a, b) => a.date.localeCompare(b.date) || a.start - b.start,
      );
      const visible = ordered.filter((task) => dayIndex.has(task.date));
      if (!visible.length) return [];

      const first = visible[0];
      const last = visible[visible.length - 1];
      const startsBefore = ordered[0].date < days[0];
      const endsAfter =
        ordered[ordered.length - 1].date > days[days.length - 1];
      const startOffset = startsBefore
        ? 0
        : Math.max(
            0,
            Math.min(
              1,
              (first.start - capacity.startHour) / capacity.dailyHours,
            ),
          );
      const endOffset = endsAfter
        ? 1
        : Math.max(
            0,
            Math.min(
              1,
              (last.start - capacity.startHour + last.estimate) /
                capacity.dailyHours,
            ),
          );
      const start = startsBefore
        ? 0
        : (dayIndex.get(first.date) ?? 0) + startOffset;
      const end = endsAfter
        ? days.length
        : (dayIndex.get(last.date) ?? 0) +
          Math.max(endOffset, startOffset + 0.02);

      return [
        {
          allocationId,
          task: ordered[0],
          start,
          end: Math.max(start + 0.02, end),
          totalHours: ordered.reduce((sum, task) => sum + task.estimate, 0),
          totalDays: new Set(ordered.map((task) => task.date)).size,
        },
      ];
    })
    .sort((a, b) => a.start - b.start || a.end - b.end);

  const laneEnds: number[] = [];
  return ranges.map((range) => {
    let lane = laneEnds.findIndex((end) => end <= range.start + 0.001);
    if (lane < 0) lane = laneEnds.length;
    laneEnds[lane] = range.end;
    return {
      allocationId: range.allocationId,
      task: range.task,
      left: (range.start / days.length) * 100,
      width: ((range.end - range.start) / days.length) * 100,
      lane,
      totalHours: range.totalHours,
      totalDays: range.totalDays,
    };
  });
}

const WORK_PHASES: WorkPhase[] = [
  "Refinement",
  "Desenvolvimento",
  "QA",
  "Deploy",
  "Suporte",
  "Reunião",
  "Outro",
];

function defaultPhase(issue: IssueRecord): WorkPhase {
  if (issue.type === "Reunião") return "Reunião";
  if (issue.type === "Suporte") return "Suporte";
  if (issue.type === "Planeamento" || issue.type === "Análise")
    return "Refinement";
  return "Desenvolvimento";
}

function TimelineView({
  issues,
  tasks,
  gitlabProjects,
  labelCatalog,
  loadProjectLabels,
  dragged,
  setDragged,
  moveTask,
  saveAllocation,
  createIssueAndAllocate,
  deleteAllocation,
  autoPlan,
  todayKey,
  capacity,
}: {
  issues: IssueRecord[];
  tasks: Task[];
  gitlabProjects: ProjectPreference[];
  labelCatalog: Record<string, GitLabLabel[]>;
  loadProjectLabels: (projectId: number) => Promise<GitLabLabel[]>;
  dragged: string | null;
  setDragged: (id: string | null) => void;
  moveTask: (id: string, date: string, start?: number) => void;
  saveAllocation: (draft: AllocationDraft) => void;
  createIssueAndAllocate: (draft: CreateIssueDraft) => Promise<void>;
  deleteAllocation: (id: string) => void;
  autoPlan: () => void;
  todayKey: string;
  capacity: CapacityConfig;
}) {
  const [zoom, setZoom] = useState<Zoom>("week");
  const [anchor, setAnchor] = useState(todayKey);
  const [backlogSearch, setBacklogSearch] = useState("");
  const [backlogFilter, setBacklogFilter] = useState<
    "all" | "unplanned" | "planned"
  >("unplanned");
  const [editor, setEditor] = useState<AllocationDraft | null>(null);
  const [createEditor, setCreateEditor] = useState<CreateIssueDraft | null>(
    null,
  );
  const [creatingIssue, setCreatingIssue] = useState(false);
  const [labelsLoading, setLabelsLoading] = useState(false);
  const [createError, setCreateError] = useState("");
  const days = visibleDaysFor(anchor, zoom, capacity.workDays);
  const openedIssues = issues.filter((issue) => issue.state === "opened");
  const visibleDaySet = new Set(days);
  const visibleTasks = tasks.filter((task) => visibleDaySet.has(task.date));
  const visibleProjects = uniqueBy(
    visibleTasks,
    (item) => `${item.projectId}:${item.project}`,
  ).sort(
    (a, b) =>
      a.client.localeCompare(b.client) || a.project.localeCompare(b.project),
  );
  const total = tasks.reduce((sum, task) => sum + task.estimate, 0);
  const finalDate = tasks.length
    ? [...tasks].sort((a, b) => b.date.localeCompare(a.date))[0].date
    : todayKey;
  const minWidth =
    210 + days.length * (zoom === "month" ? 105 : zoom === "day" ? 360 : 142);
  const gridStyle = {
    minWidth,
    gridTemplateColumns: `210px repeat(${days.length}, minmax(${zoom === "month" ? 105 : zoom === "day" ? 360 : 136}px, 1fr))`,
  };
  const plannedByIssue = useMemo(() => {
    const values = new Map<string, number>();
    tasks.forEach((task) =>
      values.set(task.rootId, (values.get(task.rootId) ?? 0) + task.estimate),
    );
    return values;
  }, [tasks]);
  const needle = backlogSearch.trim().toLowerCase();
  const backlog = openedIssues
    .filter((issue) => {
      const planned = plannedByIssue.get(issue.id) ?? 0;
      if (backlogFilter === "unplanned" && planned > 0) return false;
      if (backlogFilter === "planned" && planned <= 0) return false;
      return (
        !needle ||
        `${issue.title} ${issue.client} ${issue.project} ${issue.iid} ${(issue.labels ?? []).map((label) => label.name).join(" ")}`
          .toLowerCase()
          .includes(needle)
      );
    })
    .sort(
      (a, b) =>
        priorityOrder[a.priority] - priorityOrder[b.priority] ||
        a.title.localeCompare(b.title),
    );

  function openNew(
    issue: IssueRecord,
    date = todayKey,
    start = capacity.startHour,
  ) {
    const alreadyPlanned = plannedByIssue.get(issue.id) ?? 0;
    const suggestion =
      issue.estimateTotal > 0
        ? Math.max(0.25, issue.estimateTotal - alreadyPlanned)
        : 1;
    setEditor({
      issueId: issue.id,
      phase: defaultPhase(issue),
      hours: suggestion,
      date,
      start,
      distribution: suggestion > capacity.dailyHours ? "automatic" : "manual",
      description: "",
    });
  }

  function openEdit(allocationId: string) {
    const segments = tasks
      .filter((task) => task.allocationId === allocationId)
      .sort((a, b) => a.date.localeCompare(b.date) || a.start - b.start);
    if (!segments.length) return;
    if (segments[0].source === "outlook") {
      const task = segments[0];
      const project =
        gitlabProjects.find((item) => item.projectId === task.projectId) ??
        gitlabProjects.find((item) => item.projectId > 0);
      if (!project) return;
      setCreateError("");
      setCreateEditor({
        projectId: project.projectId,
        title: task.title,
        description: `Reunião Outlook de ${formatHours(task.estimate)} em ${formatDate(task.date, true)}.`,
        estimateHours: task.estimate,
        phase: "Reunião",
        hours: task.estimate,
        date: task.date,
        start: task.start,
        distribution: "manual",
        labels: [],
        meetingId: task.rootId.slice("outlook-".length),
      });
      refreshProjectLabels(project.projectId);
      return;
    }
    setEditor({
      issueId: segments[0].rootId,
      allocationId,
      phase: segments[0].phase,
      hours: segments.reduce((sum, task) => sum + task.estimate, 0),
      date: segments[0].date,
      start: segments[0].start,
      distribution: segments.length > 1 ? "automatic" : "manual",
      description: segments[0].description ?? "",
    });
  }

  function openCreateIssue(date = todayKey) {
    const firstProject = gitlabProjects.find(
      (project) => project.projectId > 0,
    );
    if (!firstProject) return;
    setCreateError("");
    setCreateEditor({
      projectId: firstProject.projectId,
      title: "",
      description: "",
      estimateHours: 1,
      phase: "Desenvolvimento",
      hours: 1,
      date,
      start: capacity.startHour,
      distribution: "manual",
      labels: [],
    });
    refreshProjectLabels(firstProject.projectId);
  }

  function refreshProjectLabels(projectId: number) {
    setLabelsLoading(true);
    setCreateError("");
    void loadProjectLabels(projectId)
      .catch((error) => {
        setCreateError(
          error instanceof Error
            ? error.message
            : "Não foi possível carregar as labels do projeto.",
        );
      })
      .finally(() => setLabelsLoading(false));
  }

  function toggleCreateLabel(name: string) {
    if (!createEditor) return;
    const selected = createEditor.labels.includes(name);
    setCreateEditor({
      ...createEditor,
      labels: selected
        ? createEditor.labels.filter((label) => label !== name)
        : [...createEditor.labels, name],
    });
  }

  async function submitCreateIssue() {
    if (!createEditor) return;
    setCreateError("");
    setCreatingIssue(true);
    try {
      await createIssueAndAllocate(createEditor);
      setCreateEditor(null);
    } catch (error) {
      setCreateError(
        error instanceof Error
          ? error.message
          : "Não foi possível criar a tarefa no GitLab.",
      );
    } finally {
      setCreatingIssue(false);
    }
  }

  function dropOnDay(day: string) {
    if (!dragged) return;
    if (dragged.startsWith("issue:")) {
      const issue = issues.find((item) => item.id === dragged.slice(6));
      if (issue) openNew(issue, day);
    } else if (dragged.startsWith("allocation:")) {
      moveTask(dragged.slice(11), day);
    }
    setDragged(null);
  }

  function dropOnAgenda(
    day: string,
    event: DragEvent<HTMLDivElement>,
    agendaStart: number,
    hourHeight: number,
  ) {
    if (!dragged) return;
    const bounds = event.currentTarget.getBoundingClientRect();
    const rawStart = agendaStart + (event.clientY - bounds.top) / hourHeight;
    const start = Math.max(
      agendaStart,
      Math.round(rawStart * 4) / 4,
    );
    if (dragged.startsWith("issue:")) {
      const issue = issues.find((item) => item.id === dragged.slice(6));
      if (issue) openNew(issue, day, start);
    } else if (dragged.startsWith("allocation:")) {
      moveTask(dragged.slice(11), day, start);
    }
    setDragged(null);
  }

  const editingIssue = editor
    ? issues.find((issue) => issue.id === editor.issueId)
    : undefined;
  const agendaHourHeight = 64;
  // A calendar is always a complete day. Capacity is an alert, never a visual
  // boundary for meetings or planned work.
  const agendaStart = 0;
  const agendaEnd = 24;
  const agendaHours = Array.from(
    { length: agendaEnd - agendaStart + 1 },
    (_, index) => agendaStart + index,
  );
  const agendaHeight = (agendaEnd - agendaStart) * agendaHourHeight;
  const agendaGridStyle = {
    minWidth: 68 + days.length * (zoom === "day" ? 420 : 155),
    gridTemplateColumns: `68px repeat(${days.length}, minmax(${zoom === "day" ? 420 : 155}px, 1fr))`,
  };

  return (
    <div className="page wide-page">
      <PageIntro
        eyebrow="US GITLAB · ALOCAÇÕES PESSOAIS"
        title="Linha cronológica"
        description={`${formatHours(total)} do teu tempo planeado${tasks.length ? ` até ${formatDate(finalDate, true)}` : ""}. A estimativa GitLab continua separada da tua capacidade.`}
      >
        <button className="secondary-button" onClick={autoPlan}>
          ↻ Replanear alocações
        </button>
        <button
          className="primary-button"
          onClick={() => openCreateIssue()}
          disabled={!gitlabProjects.some((project) => project.projectId > 0)}
        >
          + Criar tarefa GitLab
        </button>
      </PageIntro>
      <div className="planning-layout">
        <aside className="issue-backlog panel">
          <div className="backlog-head">
            <div>
              <span>BACKLOG</span>
              <strong>US abertas</strong>
            </div>
            <b>{openedIssues.length}</b>
          </div>
          <div className="backlog-search">
            ⌕
            <input
              value={backlogSearch}
              onChange={(event) => setBacklogSearch(event.target.value)}
              placeholder="Pesquisar US ou label…"
            />
          </div>
          <select
            className="backlog-filter"
            value={backlogFilter}
            onChange={(event) =>
              setBacklogFilter(event.target.value as typeof backlogFilter)
            }
          >
            <option value="unplanned">Por planear</option>
            <option value="planned">Planeadas por mim</option>
            <option value="all">Todas as abertas</option>
          </select>
          <div className="backlog-list">
            {backlog.map((issue) => {
              const planned = plannedByIssue.get(issue.id) ?? 0;
              return (
                <article
                  className="backlog-card"
                  key={issue.id}
                  draggable
                  onDragStart={() => setDragged(`issue:${issue.id}`)}
                  onDragEnd={() => setDragged(null)}
                  style={{ borderLeftColor: issue.color }}
                >
                  <div className="backlog-meta">
                    <span>{planned > 0 ? "PLANEADA" : "POR PLANEAR"}</span>
                    <b>#{issue.iid}</b>
                  </div>
                  {issue.webUrl ? (
                    <a href={issue.webUrl} target="_blank" rel="noreferrer">
                      {issue.title}
                      <span>↗</span>
                    </a>
                  ) : (
                    <strong>{issue.title}</strong>
                  )}
                  <p>
                    {issue.client} · {issue.project}
                  </p>
                  <IssueLabels labels={issue.labels} />
                  <div className="backlog-hours">
                    <span>
                      <small>Estimate equipa</small>
                      <b>
                        {issue.hasEstimate
                          ? formatHours(issue.estimateTotal)
                          : "—"}
                      </b>
                    </span>
                    <span>
                      <small>Meu plano</small>
                      <b>{formatHours(planned)}</b>
                    </span>
                  </div>
                  {issue.assigneeCount > 1 && (
                    <em>
                      Estimativa partilhada por {issue.assigneeCount} pessoas
                    </em>
                  )}
                  <button onClick={() => openNew(issue)}>
                    + Planear participação
                  </button>
                </article>
              );
            })}
            {!backlog.length && (
              <div className="backlog-empty">
                <span>✓</span>
                <p>
                  {backlogFilter === "unplanned"
                    ? "Não tens US abertas por planear."
                    : "Nenhuma US neste filtro."}
                </p>
              </div>
            )}
          </div>
        </aside>
        <div className="timeline-column">
          <section className="timeline-toolbar panel">
            <div className="legend">
              <span>
                <i className="legend-fixed" /> Reunião fixa
              </span>
              <span>
                <i className="legend-due" /> Alocação pessoal
              </span>
              <span>
                <i className="legend-drag" /> Arrasta US ou bloco
              </span>
            </div>
            <div className="period-controls">
              <button
                className="period-arrow"
                aria-label="Período anterior"
                onClick={() =>
                  setAnchor(movePeriod(anchor, zoom, -1, capacity.workDays))
                }
              >
                ‹
              </button>
              <button
                className="today-button"
                onClick={() => setAnchor(todayKey)}
              >
                Hoje
              </button>
              <button
                className="period-arrow"
                aria-label="Período seguinte"
                onClick={() =>
                  setAnchor(movePeriod(anchor, zoom, 1, capacity.workDays))
                }
              >
                ›
              </button>
              <div className="zoom">
                {(["day", "week", "month"] as Zoom[]).map((option) => (
                  <button
                    key={option}
                    className={zoom === option ? "active" : ""}
                    onClick={() => {
                      setZoom(option);
                      setAnchor(todayKey);
                    }}
                  >
                    {option === "day"
                      ? "Dia"
                      : option === "week"
                        ? "Semana"
                        : "Mês"}
                  </button>
                ))}
              </div>
            </div>
          </section>
          {zoom !== "month" && (
            <section className="agenda-timeline panel">
              <div className="agenda-head" style={agendaGridStyle}>
                <div className="agenda-time-head">HORA</div>
                {days.map((key) => {
                  const day = fromKey(key);
                  const hours = tasks
                    .filter((task) => task.date === key)
                    .reduce((sum, task) => sum + task.estimate, 0);
                  return (
                    <div
                      className={`agenda-day-head ${key === todayKey ? "today" : ""}`}
                      key={key}
                    >
                      <span>{WEEKDAYS[day.getDay()]}</span>
                      <strong>{String(day.getDate()).padStart(2, "0")}</strong>
                      <small>{MONTHS[day.getMonth()]}</small>
                      <em
                        className={
                          hours > capacity.dailyHours ? "over" : undefined
                        }
                      >
                        {formatHours(hours)} / {formatHours(capacity.dailyHours)}
                      </em>
                    </div>
                  );
                })}
              </div>
              <div className="agenda-scroll">
                <div
                  className="agenda-body"
                  style={{ ...agendaGridStyle, height: agendaHeight }}
                >
                  <div className="agenda-time-axis">
                    {agendaHours.map((hour) => (
                      <span
                        key={hour}
                        style={{ top: (hour - agendaStart) * agendaHourHeight }}
                      >
                        {formatClock(hour)}
                      </span>
                    ))}
                  </div>
                  {days.map((day) => {
                    const items = layoutAgendaDay(
                      tasks.filter((task) => task.date === day),
                    );
                    return (
                      <div
                        key={day}
                        className={`agenda-day-column ${day === todayKey ? "today" : ""} ${dragged ? "drop-ready" : ""}`}
                        onDragOver={(event) => event.preventDefault()}
                        onDrop={(event) =>
                          dropOnAgenda(
                            day,
                            event,
                            agendaStart,
                            agendaHourHeight,
                          )
                        }
                      >
                        {items.map(({ task, column, columnCount }) => {
                          const start = Math.max(task.start, agendaStart);
                          const end = Math.min(
                            task.start + task.estimate,
                            agendaEnd,
                          );
                          return (
                            <article
                              key={task.id}
                              className={`agenda-event ${task.fixed ? "fixed" : ""}`}
                              style={{
                                borderColor: task.color,
                                top: (start - agendaStart) * agendaHourHeight,
                                height: Math.max(
                                  24,
                                  (end - start) * agendaHourHeight,
                                ),
                                left: `calc(${(column / columnCount) * 100}% + 3px)`,
                                width: `calc(${100 / columnCount}% - 6px)`,
                              }}
                              draggable={!task.fixed}
                              onClick={() => openEdit(task.allocationId)}
                              onDragStart={() => {
                                if (!task.fixed)
                                  setDragged(
                                    `allocation:${task.allocationId}`,
                                  );
                              }}
                              onDragEnd={() => setDragged(null)}
                              title={`${formatClock(task.start)}–${formatClock(task.start + task.estimate)} · ${task.title}`}
                            >
                              <div>
                                <span>
                                  {formatClock(task.start)}–
                                  {formatClock(task.start + task.estimate)}
                                </span>
                                <small>{formatHours(task.estimate)}</small>
                              </div>
                              <strong>{task.title}</strong>
                              <p>
                                {task.client} · {task.project}
                                {task.source === "outlook"
                                  ? " · Criar US"
                                  : ` · #${task.iid}`}
                              </p>
                            </article>
                          );
                        })}
                      </div>
                    );
                  })}
                </div>
              </div>
            </section>
          )}
          {zoom === "month" && (
          <section className="timeline panel">
            <div className="timeline-head" style={gridStyle}>
              <div className="timeline-client-head">CLIENTE / PROJETO</div>
              {days.map((key) => {
                const day = fromKey(key);
                return (
                  <div
                    className={key === todayKey ? "day-head today" : "day-head"}
                    key={key}
                  >
                    <span>{WEEKDAYS[day.getDay()]}</span>
                    <strong>{String(day.getDate()).padStart(2, "0")}</strong>
                    <small>{MONTHS[day.getMonth()]}</small>
                    {key === todayKey && <i>HOJE</i>}
                  </div>
                );
              })}
            </div>
            <div className="timeline-body">
              {dragged?.startsWith("issue:") && (
                <div className="timeline-drop-row" style={gridStyle}>
                  <div className="timeline-client">
                    <div>
                      <strong>Adicionar US</strong>
                      <span>Larga no dia pretendido</span>
                    </div>
                  </div>
                  {days.map((day) => (
                    <div
                      key={day}
                      className={`timeline-cell ${day === todayKey ? "today" : ""}`}
                      onDragOver={(event) => event.preventDefault()}
                      onDrop={() => dropOnDay(day)}
                      aria-label={`Adicionar US em ${formatDate(day, true)}`}
                    />
                  ))}
                </div>
              )}
              {visibleProjects.map((project) => {
                const projectTasks = tasks.filter(
                  (task) =>
                    task.projectId === project.projectId &&
                    task.project === project.project,
                );
                const visibleProjectTasks = projectTasks.filter((task) =>
                  visibleDaySet.has(task.date),
                );
                const bars = buildTimelineBars(projectTasks, days, capacity);
        const rowHeight = Math.max(
          106,
          18 + Math.max(1, ...bars.map((bar) => bar.lane + 1)) * 86,
        );
                return (
                  <div
                    className="timeline-row"
                    style={{ ...gridStyle, minHeight: rowHeight }}
                    key={`${project.projectId}:${project.project}`}
                  >
                    <div className="timeline-client">
                      <i style={{ background: project.color }} />
                      <div>
                        <strong>{project.client}</strong>
                        <span>{project.project}</span>
                      </div>
                      <small>
                        {formatHours(
                          visibleProjectTasks.reduce(
                            (sum, task) => sum + task.estimate,
                            0,
                          ),
                        )}
                      </small>
                    </div>
                    {days.map((day) => (
                      <div
                        key={day}
                        className={`timeline-cell ${day === todayKey ? "today" : ""} ${dragged ? "drop-ready" : ""}`}
                        onDragOver={(event) => event.preventDefault()}
                        onDrop={() => dropOnDay(day)}
                      />
                    ))}
                    <div
                      className={`timeline-bars ${dragged ? "dragging" : ""}`}
                      aria-label={`Planeamento de ${project.project}`}
                    >
                      {bars.map((bar) => (
                        <article
                          key={bar.allocationId}
                          className={`timeline-task timeline-span ${bar.task.fixed ? "fixed" : ""}`}
                          style={{
                            borderColor: bar.task.color,
                            left: `calc(${bar.left}% + 4px)`,
                            width: `max(30px, calc(${bar.width}% - 8px))`,
                            top: 9 + bar.lane * 86,
                          }}
                          draggable
                          onClick={() => openEdit(bar.allocationId)}
                          onDragStart={() =>
                            setDragged(`allocation:${bar.allocationId}`)
                          }
                          onDragEnd={() => setDragged(null)}
                          title={`${bar.task.description ? `${bar.task.description} · ` : ""}${bar.task.phase} · ${formatHours(bar.totalHours)} em ${bar.totalDays} ${bar.totalDays === 1 ? "dia" : "dias"}`}
                        >
                          <div>
                            <span style={{ color: bar.task.color }}>
                              {bar.task.iid
                                ? `#${bar.task.iid} · ${bar.task.phase}`
                                : "OUTLOOK · CRIAR US"}
                            </span>
                            <small>{formatHours(bar.totalHours)}</small>
                          </div>
                          <IssueTitle task={bar.task} compact />
                          <IssueLabels labels={bar.task.labels} compact />
                          <p>
                            {bar.task.description ||
                              (bar.task.due === "Reunião"
                                ? "Reunião fixa"
                                : `${bar.task.phase} · ${bar.totalDays} ${bar.totalDays === 1 ? "dia" : "dias"}`)}
                          </p>
                        </article>
                      ))}
                    </div>
                  </div>
                );
              })}
              {!visibleProjects.length && !dragged?.startsWith("issue:") && (
                <div className="timeline-empty" style={gridStyle}>
                  <div className="empty-state">
                    <span>↔</span>
                    <strong>Sem horas planeadas neste período</strong>
                    <p>
                      Arrasta uma US do backlog ou cria uma tarefa GitLab e
                      planeia-a imediatamente.
                    </p>
                  </div>
                </div>
              )}
              <div className="capacity-row" style={gridStyle}>
                <div className="timeline-client">
                  <strong>CAPACIDADE DIÁRIA</strong>
                </div>
                {days.map((day) => {
                  const hours = tasks
                    .filter((task) => task.date === day)
                    .reduce((sum, task) => sum + task.estimate, 0);
                  return (
                    <div className="capacity-cell" key={day}>
                      <strong
                        className={
                          hours > capacity.dailyHours
                            ? "over"
                            : hours === capacity.dailyHours
                              ? "full"
                              : ""
                        }
                      >
                        {formatHours(hours)}
                      </strong>
                      <div>
                        <i
                          style={{
                            width: `${Math.min(100, (hours / capacity.dailyHours) * 100)}%`,
                          }}
                        />
                      </div>
                      <small>
                        {hours > capacity.dailyHours
                          ? "Sobrecarga"
                          : hours === capacity.dailyHours
                            ? "Completo"
                            : `${formatHours(capacity.dailyHours - hours)} livres`}
                      </small>
                    </div>
                  );
                })}
              </div>
            </div>
          </section>
          )}
          <p className="timeline-hint">
            <span>↔</span>{" "}
            {zoom === "month"
              ? "Arrasta uma US do backlog para um dia."
              : "Arrasta uma US para uma hora livre; o horário ajusta-se em intervalos de 15 minutos."}
          </p>
        </div>
      </div>
      {editor && editingIssue && (
        <div
          className="modal-backdrop"
          role="presentation"
          onMouseDown={() => setEditor(null)}
        >
          <section
            className="allocation-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="allocation-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="allocation-modal-head">
              <div>
                <span>
                  {editor.allocationId
                    ? "EDITAR ALOCAÇÃO"
                    : "NOVA ALOCAÇÃO PESSOAL"}
                </span>
                <h2 id="allocation-title">
                  #{editingIssue.iid} · {editingIssue.title}
                </h2>
                <p>
                  {editingIssue.client} · {editingIssue.project}
                </p>
              </div>
              <button onClick={() => setEditor(null)} aria-label="Fechar">
                ×
              </button>
            </div>
            <div className="allocation-issue-labels">
              <IssueLabels labels={editingIssue.labels} />
            </div>
            <div className="allocation-context">
              <span>
                <small>Estimate da equipa</small>
                <strong>
                  {editingIssue.hasEstimate
                    ? formatHours(editingIssue.estimateTotal)
                    : "Sem estimate"}
                </strong>
              </span>
              <span>
                <small>Meu planeado</small>
                <strong>
                  {formatHours(plannedByIssue.get(editingIssue.id) ?? 0)}
                </strong>
              </span>
              <span>
                <small>Meu realizado</small>
                <strong>{formatHours(editingIssue.personalSpent)}</strong>
              </span>
            </div>
            {editingIssue.assigneeCount > 1 && (
              <div className="allocation-warning">
                <b>!</b>
                <p>
                  Esta é a estimativa total da US e existem{" "}
                  {editingIssue.assigneeCount} pessoas atribuídas. Confirma
                  apenas as horas da tua participação.
                </p>
              </div>
            )}
            <div className="allocation-form">
              <label className="allocation-description">
                <span>O que vou fazer nesta alocação</span>
                <textarea
                  autoFocus={!editor.allocationId}
                  rows={3}
                  value={editor.description}
                  onChange={(event) =>
                    setEditor({ ...editor, description: event.target.value })
                  }
                  placeholder="Ex.: rever a estrutura da landing page, ajustar os componentes e validar a copy em mobile."
                />
              </label>
              <label>
                <span>Minha fase</span>
                <select
                  value={editor.phase}
                  onChange={(event) =>
                    setEditor({
                      ...editor,
                      phase: event.target.value as WorkPhase,
                    })
                  }
                >
                  {WORK_PHASES.map((phase) => (
                    <option key={phase}>{phase}</option>
                  ))}
                </select>
              </label>
              <label>
                <span>Minhas horas</span>
                <input
                  type="number"
                  min="0.25"
                  max="500"
                  step="0.25"
                  value={editor.hours}
                  onChange={(event) =>
                    setEditor({
                      ...editor,
                      hours: Math.max(0.25, Number(event.target.value) || 0.25),
                    })
                  }
                />
              </label>
              <label>
                <span>Data inicial</span>
                <input
                  type="date"
                  value={editor.date}
                  onChange={(event) =>
                    setEditor({ ...editor, date: event.target.value })
                  }
                />
              </label>
              <label>
                <span>Hora inicial</span>
                <input
                  type="time"
                  step="900"
                  value={formatClock(editor.start)}
                  onChange={(event) => {
                    const [hour, minute] = event.target.value
                      .split(":")
                      .map(Number);
                    setEditor({ ...editor, start: hour + minute / 60 });
                  }}
                  disabled={editor.distribution === "automatic"}
                />
              </label>
              <label className="allocation-distribution">
                <span>Distribuição</span>
                <div>
                  <button
                    type="button"
                    className={editor.distribution === "manual" ? "active" : ""}
                    onClick={() =>
                      setEditor({ ...editor, distribution: "manual" })
                    }
                  >
                    Manual · um bloco
                  </button>
                  <button
                    type="button"
                    className={
                      editor.distribution === "automatic" ? "active" : ""
                    }
                    onClick={() =>
                      setEditor({ ...editor, distribution: "automatic" })
                    }
                  >
                    Automática · pela capacidade
                  </button>
                </div>
              </label>
            </div>
            <div className="allocation-modal-actions">
              {editor.allocationId && (
                <button
                  className="delete-button"
                  onClick={() => {
                    deleteAllocation(editor.allocationId!);
                    setEditor(null);
                  }}
                >
                  Remover
                </button>
              )}
              <span />
              <button
                className="secondary-button"
                onClick={() => setEditor(null)}
              >
                Cancelar
              </button>
              <button
                className="primary-button"
                disabled={!editor.description.trim()}
                onClick={() => {
                  saveAllocation(editor);
                  setEditor(null);
                }}
              >
                {editor.allocationId
                  ? "Guardar alterações"
                  : "Adicionar à Timeline"}
              </button>
            </div>
          </section>
        </div>
      )}
      {createEditor && (
        <div
          className="modal-backdrop"
          role="presentation"
          onMouseDown={() => !creatingIssue && setCreateEditor(null)}
        >
          <section
            className="allocation-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="create-task-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="allocation-modal-head">
              <div>
                <span>
                  {createEditor.meetingId
                    ? "REUNIÃO OUTLOOK → US GITLAB"
                    : "NOVA TAREFA GITLAB + ALOCAÇÃO"}
                </span>
                <h2 id="create-task-title">
                  {createEditor.meetingId
                    ? "Criar US para esta reunião"
                    : "Criar e planear num único passo"}
                </h2>
                <p>
                  A tarefa fica atribuída a ti no GitLab e entra imediatamente
                  na Timeline.
                </p>
              </div>
              <button
                onClick={() => setCreateEditor(null)}
                aria-label="Fechar"
                disabled={creatingIssue}
              >
                ×
              </button>
            </div>
            <div className="allocation-form">
              <label className="allocation-description">
                <span>Projeto GitLab</span>
                <select
                  value={createEditor.projectId}
                  onChange={(event) => {
                    const projectId = Number(event.target.value);
                    setCreateEditor({
                      ...createEditor,
                      projectId,
                      labels: [],
                    });
                    refreshProjectLabels(projectId);
                  }}
                >
                  {gitlabProjects
                    .filter((project) => project.projectId > 0)
                    .map((project) => (
                      <option
                        key={project.projectId}
                        value={project.projectId}
                      >
                        {project.client} · {project.project}
                      </option>
                    ))}
                </select>
              </label>
              <label className="allocation-description">
                <span>Título da tarefa</span>
                <input
                  autoFocus
                  value={createEditor.title}
                  onChange={(event) =>
                    setCreateEditor({
                      ...createEditor,
                      title: event.target.value,
                    })
                  }
                  placeholder="Ex.: Ajustar estrutura da landing page"
                />
              </label>
              <label className="allocation-description">
                <span>Descrição do trabalho</span>
                <textarea
                  rows={4}
                  value={createEditor.description}
                  onChange={(event) =>
                    setCreateEditor({
                      ...createEditor,
                      description: event.target.value,
                    })
                  }
                  placeholder="Descreve concretamente o que vais fazer. Esta descrição fica também na Issue do GitLab."
                />
              </label>
              <div className="allocation-description label-picker-field">
                <span>Labels GitLab</span>
                {labelsLoading ? (
                  <div className="label-picker-state">A carregar labels do projeto…</div>
                ) : (labelCatalog[String(createEditor.projectId)] ?? []).length ? (
                  <div className="label-picker" role="group" aria-label="Labels da nova tarefa">
                    {(labelCatalog[String(createEditor.projectId)] ?? []).map((label) => {
                      const selected = createEditor.labels.includes(label.name);
                      return (
                        <button
                          type="button"
                          key={label.name}
                          className={selected ? "selected" : ""}
                          onClick={() => toggleCreateLabel(label.name)}
                          title={label.description || label.name}
                          aria-pressed={selected}
                        >
                          <i style={{ backgroundColor: label.color }} />
                          <span>{label.name}</span>
                          <b>{selected ? "✓" : "+"}</b>
                        </button>
                      );
                    })}
                  </div>
                ) : (
                  <div className="label-picker-state">Este projeto não tem labels definidas.</div>
                )}
              </div>
              <label>
                <span>Estimate GitLab</span>
                <input
                  type="number"
                  min="0.25"
                  max="500"
                  step="0.25"
                  value={createEditor.estimateHours}
                  onChange={(event) =>
                    setCreateEditor({
                      ...createEditor,
                      estimateHours: Math.max(
                        0.25,
                        Number(event.target.value) || 0.25,
                      ),
                    })
                  }
                />
              </label>
              <label>
                <span>Minhas horas planeadas</span>
                <input
                  type="number"
                  min="0.25"
                  max="500"
                  step="0.25"
                  value={createEditor.hours}
                  onChange={(event) =>
                    setCreateEditor({
                      ...createEditor,
                      hours: Math.max(0.25, Number(event.target.value) || 0.25),
                    })
                  }
                />
              </label>
              <label>
                <span>Minha fase</span>
                <select
                  value={createEditor.phase}
                  onChange={(event) =>
                    setCreateEditor({
                      ...createEditor,
                      phase: event.target.value as WorkPhase,
                    })
                  }
                >
                  {WORK_PHASES.map((phase) => (
                    <option key={phase}>{phase}</option>
                  ))}
                </select>
              </label>
              <label>
                <span>Data inicial</span>
                <input
                  type="date"
                  value={createEditor.date}
                  onChange={(event) =>
                    setCreateEditor({
                      ...createEditor,
                      date: event.target.value,
                    })
                  }
                />
              </label>
              <label>
                <span>Hora inicial</span>
                <input
                  type="time"
                  step="900"
                  value={formatClock(createEditor.start)}
                  onChange={(event) => {
                    const [hour, minute] = event.target.value
                      .split(":")
                      .map(Number);
                    setCreateEditor({
                      ...createEditor,
                      start: hour + minute / 60,
                    });
                  }}
                  disabled={createEditor.distribution === "automatic"}
                />
              </label>
              <label className="allocation-distribution">
                <span>Distribuição</span>
                <div>
                  <button
                    type="button"
                    className={
                      createEditor.distribution === "manual" ? "active" : ""
                    }
                    onClick={() =>
                      setCreateEditor({
                        ...createEditor,
                        distribution: "manual",
                      })
                    }
                  >
                    Manual · um bloco
                  </button>
                  <button
                    type="button"
                    className={
                      createEditor.distribution === "automatic" ? "active" : ""
                    }
                    onClick={() =>
                      setCreateEditor({
                        ...createEditor,
                        distribution: "automatic",
                      })
                    }
                  >
                    Automática · pela capacidade
                  </button>
                </div>
              </label>
              {createError && (
                <div className="create-task-error">{createError}</div>
              )}
            </div>
            <div className="allocation-modal-actions">
              <span />
              <span />
              <button
                className="secondary-button"
                onClick={() => setCreateEditor(null)}
                disabled={creatingIssue}
              >
                Cancelar
              </button>
              <button
                className="primary-button"
                disabled={
                  creatingIssue ||
                  !createEditor.title.trim() ||
                  !createEditor.description.trim()
                }
                onClick={submitCreateIssue}
              >
                {creatingIssue
                  ? "A criar no GitLab…"
                  : createEditor.meetingId
                    ? "Criar US da reunião"
                    : "Criar e adicionar à Timeline"}
              </button>
            </div>
          </section>
        </div>
      )}
    </div>
  );
}

function periodStart(todayKey: string, period: Zoom | "all") {
  if (period === "all") return "0000-00-00";
  if (period === "day") return todayKey;
  if (period === "week") return mondayFor(todayKey);
  const value = fromKey(todayKey);
  value.setDate(1);
  return dateKey(value);
}

function PeriodSelector({
  value,
  onChange,
  allowAll = false,
}: {
  value: Zoom | "all";
  onChange: (value: Zoom | "all") => void;
  allowAll?: boolean;
}) {
  const options: (Zoom | "all")[] = allowAll
    ? ["day", "week", "month", "all"]
    : ["day", "week", "month"];
  return (
    <div className="zoom period-filter">
      {options.map((option) => (
        <button
          key={option}
          className={value === option ? "active" : ""}
          onClick={() => onChange(option)}
        >
          {option === "day"
            ? "Dia"
            : option === "week"
              ? "Semana"
              : option === "month"
                ? "Mês"
                : "Tudo"}
        </button>
      ))}
    </div>
  );
}

function CompletedView({
  logs,
  todayKey,
}: {
  logs: WorkLog[];
  todayKey: string;
}) {
  const [period, setPeriod] = useState<Zoom | "all">("week");
  const [search, setSearch] = useState("");
  const [client, setClient] = useState("all");
  const [type, setType] = useState("all");
  const start = periodStart(todayKey, period);
  const filtered = logs.filter(
    (log) =>
      log.dateKey >= start &&
      (client === "all" || log.client === client) &&
      (type === "all" || log.type === type) &&
      `${log.client} ${log.project} ${log.task}`
        .toLowerCase()
        .includes(search.toLowerCase()),
  );
  const total = filtered.reduce((sum, log) => sum + log.hours, 0);
  const todayTotal = logs
    .filter((log) => log.dateKey === todayKey)
    .reduce((sum, log) => sum + log.hours, 0);
  return (
    <div className="page">
      <PageIntro
        eyebrow="REGISTO DE TEMPO"
        title="Trabalho realizado"
        description="Issues concluídas e sessões registadas, organizadas por cliente e projeto."
      >
        <PeriodSelector value={period} onChange={setPeriod} allowAll />
        <button
          className="secondary-button"
          onClick={() => exportCsv(filtered)}
        >
          ⇩ Exportar CSV
        </button>
      </PageIntro>
      <section className="summary-strip panel">
        <div>
          <small>PERÍODO VISÍVEL</small>
          <strong>{formatHours(total)}</strong>
          <span>{filtered.length} registos</span>
        </div>
        <div>
          <small>HOJE</small>
          <strong>{formatHours(todayTotal)}</strong>
          <span>de 8 h</span>
        </div>
        <div>
          <small>CLIENTES</small>
          <strong>{new Set(filtered.map((log) => log.client)).size}</strong>
          <span>ativos</span>
        </div>
        <div>
          <small>PRECISÃO</small>
          <strong>{accuracy(filtered)}%</strong>
          <span>estimativa vs. real</span>
        </div>
      </section>
      <section className="panel table-panel">
        <div className="table-toolbar">
          <div className="search-field">
            ⌕{" "}
            <input
              aria-label="Pesquisar"
              placeholder="Pesquisar tarefa ou cliente…"
              value={search}
              onChange={(event) => setSearch(event.target.value)}
            />
          </div>
          <select
            value={client}
            onChange={(event) => setClient(event.target.value)}
            aria-label="Filtrar por cliente"
          >
            <option value="all">Todos os clientes</option>
            {[...new Set(logs.map((log) => log.client))].sort().map((value) => (
              <option key={value}>{value}</option>
            ))}
          </select>
          <select
            value={type}
            onChange={(event) => setType(event.target.value)}
            aria-label="Filtrar por tipo"
          >
            <option value="all">Todos os tipos</option>
            {[...new Set(logs.map((log) => log.type))].sort().map((value) => (
              <option key={value}>{value}</option>
            ))}
          </select>
        </div>
        <div className="work-table">
          <div className="table-row table-head">
            <span>DATA</span>
            <span>CLIENTE / PROJETO</span>
            <span>TAREFA</span>
            <span>TIPO</span>
            <span>HORAS</span>
            <span />
          </div>
          {filtered.map((log) => (
            <div className="table-row" key={log.id}>
              <span>{formatDate(log.dateKey)}</span>
              <span>
                <strong>{log.client}</strong>
                <small>{log.project}</small>
              </span>
              <span>
                {log.webUrl ? (
                  <a
                    className="table-link"
                    href={log.webUrl}
                    target="_blank"
                    rel="noreferrer"
                  >
                    {log.task}
                    <b>↗</b>
                  </a>
                ) : (
                  log.task
                )}
              </span>
              <span>
                <i className={`type-dot ${log.type.toLowerCase()}`} />
                {log.type}
              </span>
              <span>
                <strong>{formatHours(log.hours)}</strong>
              </span>
              <span>•••</span>
            </div>
          ))}
          {!filtered.length && (
            <div className="empty-state table-empty">
              <span>⌕</span>
              <strong>Sem registos neste período</strong>
              <p>Altera os filtros ou sincroniza novamente o GitLab.</p>
            </div>
          )}
        </div>
        <div className="table-footer">
          <span>A mostrar {filtered.length} registos</span>
          <strong>
            Total visível <b>{formatHours(total)}</b>
          </strong>
        </div>
      </section>
    </div>
  );
}

function accuracy(logs: WorkLog[]) {
  const withEstimate = logs.filter(
    (log) => (log.estimate ?? 0) > 0 && log.hours > 0,
  );
  if (!withEstimate.length) return 0;
  const score =
    withEstimate.reduce(
      (sum, log) =>
        sum +
        Math.min(log.hours, log.estimate!) / Math.max(log.hours, log.estimate!),
      0,
    ) / withEstimate.length;
  return Math.round(score * 100);
}

function AnalyticsView({
  issues,
  logs,
  tasks,
  todayKey,
}: {
  issues: IssueRecord[];
  logs: WorkLog[];
  tasks: Task[];
  todayKey: string;
}) {
  const [period, setPeriod] = useState<Zoom | "all">("month");
  const filtered = logs.filter(
    (log) => log.dateKey >= periodStart(todayKey, period),
  );
  const total = filtered.reduce((sum, log) => sum + log.hours, 0);
  const grouped = new Map<string, number>();
  filtered.forEach((log) =>
    grouped.set(log.client, (grouped.get(log.client) ?? 0) + log.hours),
  );
  const clientTotals = [...grouped.entries()].sort((a, b) => b[1] - a[1]);
  const maxClient = Math.max(1, ...clientTotals.map((entry) => entry[1]));
  const teamEstimate = issues
    .filter((issue) => issue.state === "opened")
    .reduce((sum, issue) => sum + issue.estimateTotal, 0);
  const personalPlanned = tasks.reduce((sum, task) => sum + task.estimate, 0);
  const personalRemaining = Math.max(0, personalPlanned - total);
  return (
    <div className="page">
      <PageIntro
        eyebrow="ANÁLISE DE CAPACIDADE"
        title="Onde está o teu tempo?"
        description="Estimate da equipa, planeamento pessoal e tempo realizado são tratados como medidas independentes."
      >
        <PeriodSelector value={period} onChange={setPeriod} allowAll />
      </PageIntro>
      <section className="metrics-grid analytics-metrics">
        <article>
          <small>ESTIMATE GITLAB · EQUIPA</small>
          <strong>{formatHours(teamEstimate)}</strong>
          <p>
            {issues.filter((issue) => issue.state === "opened").length} US
            abertas
          </p>
        </article>
        <article>
          <small>MEU PLANEADO</small>
          <strong>{formatHours(personalPlanned)}</strong>
          <p>
            {new Set(tasks.map((task) => task.allocationId)).size} alocações
            pessoais
          </p>
        </article>
        <article>
          <small>MEU REALIZADO</small>
          <strong>{formatHours(total)}</strong>
          <p>{filtered.length} timelogs no período</p>
        </article>
        <article>
          <small>RESTANTE PESSOAL</small>
          <strong>{formatHours(personalRemaining)}</strong>
          <p>Planeado menos realizado</p>
        </article>
      </section>
      <section className="analytics-grid">
        <div className="panel chart-card">
          <div className="panel-head">
            <div>
              <h2>Horas por cliente</h2>
              <p>Distribuição do meu tempo registado</p>
            </div>
            <span className="total-pill">{formatHours(total)}</span>
          </div>
          <div className="bar-chart">
            {clientTotals.map(([client, hours], index) => (
              <div className="bar-row" key={client}>
                <span>{client}</span>
                <div>
                  <i
                    style={{
                      width: `${(hours / maxClient) * 100}%`,
                      background: CLIENT_COLORS[index % CLIENT_COLORS.length],
                    }}
                  />
                </div>
                <strong>{formatHours(hours)}</strong>
              </div>
            ))}
            {!clientTotals.length && (
              <div className="empty-state">
                <span>⌁</span>
                <strong>Sem dados para analisar</strong>
                <p>Sincroniza os teus timelogs ou altera o período.</p>
              </div>
            )}
          </div>
        </div>
        <div className="panel chart-card">
          <div className="panel-head">
            <div>
              <h2>Meu planeado por projeto</h2>
              <p>Apenas alocações pessoais na Timeline</p>
            </div>
          </div>
          <div className="bar-chart">
            {projectTotals(tasks)
              .slice(0, 8)
              .map(([project, hours], index) => (
                <div className="bar-row" key={project}>
                  <span>{project}</span>
                  <div>
                    <i
                      style={{
                        width: `${(hours / Math.max(1, ...projectTotals(tasks).map((entry) => entry[1]))) * 100}%`,
                        background: CLIENT_COLORS[index % CLIENT_COLORS.length],
                      }}
                    />
                  </div>
                  <strong>{formatHours(hours)}</strong>
                </div>
              ))}
          </div>
        </div>
      </section>
      <section className="panel insight">
        <span>✦</span>
        <div>
          <strong>Leitura automática</strong>
          <p>
            {total
              ? `${clientTotals[0]?.[0] ?? "O principal cliente"} representa ${Math.round(((clientTotals[0]?.[1] ?? 0) / total) * 100)}% do teu tempo registado no período.`
              : `${formatHours(personalPlanned)} planeadas sem tempo pessoal suficiente para gerar um insight.`}
          </p>
        </div>
      </section>
    </div>
  );
}

function SettingsView(props: {
  section: SettingsSection;
  setSection: (section: SettingsSection) => void;
  config: Config;
  setConfig: (config: Config) => void;
  showToken: boolean;
  setShowToken: (show: boolean) => void;
  connection: { state: string; message: string };
  saveConfig: () => void;
  testConnection: () => void;
  outlookConfig: OutlookConfig;
  setOutlookConfig: (config: OutlookConfig) => void;
  outlookConnection: { state: string; message: string };
  saveOutlookConfig: () => void;
  connectOutlook: () => void;
  syncOutlook: () => void;
  syncLocalOutlook: () => void;
  syncingCalendar: boolean;
  meetings: CalendarMeeting[];
  projects: {
    projectId: number;
    client: string;
    project: string;
    color: string;
  }[];
  associateMeeting: (meetingId: string, projectId: number | undefined) => void;
  capacity: CapacityConfig;
  setCapacity: (capacity: CapacityConfig) => void;
  saveCapacity: () => void;
  projectPreferences: ProjectPreference[];
  updateProjectPreference: (
    projectId: number,
    changes: Partial<ProjectPreference>,
  ) => void;
  syncGitLab: () => void;
  syncing: boolean;
}) {
  const {
    section,
    setSection,
    config,
    setConfig,
    showToken,
    setShowToken,
    connection,
    saveConfig,
    testConnection,
    outlookConfig,
    setOutlookConfig,
    outlookConnection,
    saveOutlookConfig,
    connectOutlook,
    syncOutlook,
    syncLocalOutlook,
    syncingCalendar,
    meetings,
    projects,
    associateMeeting,
    capacity,
    setCapacity,
    saveCapacity,
    projectPreferences,
    updateProjectPreference,
    syncGitLab,
    syncing,
  } = props;
  const titles: Record<SettingsSection, [string, string]> = {
    gitlab: [
      "Liga o teu GitLab",
      "Gere a integração sem abrir ou alterar ficheiros .env.",
    ],
    calendar: [
      "Liga o Outlook e Teams",
      "Importa reuniões da semana e associa cada uma a um projeto GitLab.",
    ],
    capacity: [
      "Define a tua capacidade",
      "Configura quantas horas trabalhas e quais os dias que entram no planeamento.",
    ],
    clients: [
      "Organiza clientes e projetos",
      "Ajusta os nomes e cores descobertos através da sincronização real do GitLab.",
    ],
  };

  let content: React.ReactNode;
  if (section === "gitlab") {
    content = (
      <section className="panel settings-card">
        <div className="settings-title">
          <div className="gitlab-logo">◆</div>
          <div>
            <h2>Ligação ao GitLab</h2>
            <p>
              Descobre projetos, sincroniza Issues abertas e concluídas e
              regista spent time.
            </p>
          </div>
          <span
            className={config.token ? "status-pill connected" : "status-pill"}
          >
            {config.token ? "● Configurado" : "○ Por configurar"}
          </span>
        </div>
        <div className="security-note">
          <span>▣</span>
          <div>
            <strong>A credencial fica neste dispositivo</strong>
            <p>
              O token é guardado no armazenamento local do teu browser e enviado
              apenas quando sincronizas ou registas tempo.
            </p>
          </div>
        </div>
        <label className="form-field">
          <span>URL da instância GitLab</span>
          <input
            type="url"
            value={config.baseUrl}
            onChange={(event) =>
              setConfig({ ...config, baseUrl: event.target.value })
            }
            placeholder="https://gitlab.empresa.pt"
          />
        </label>
        <label className="form-field">
          <span>Personal Access Token</span>
          <div className="secret-field">
            <input
              type={showToken ? "text" : "password"}
              value={config.token}
              onChange={(event) =>
                setConfig({ ...config, token: event.target.value })
              }
              placeholder="glpat-••••••••••••••••••••"
              autoComplete="off"
            />
            <button type="button" onClick={() => setShowToken(!showToken)}>
              {showToken ? "Ocultar" : "Mostrar"}
            </button>
          </div>
          <small>
            Scope necessário: <code>api</code>.
          </small>
        </label>
        {connection.message && (
          <div className={`connection-result ${connection.state}`}>
            <span>
              {connection.state === "ok"
                ? "✓"
                : connection.state === "error"
                  ? "!"
                  : "↻"}
            </span>
            {connection.message}
          </div>
        )}
        <div className="settings-actions">
          <button
            className="secondary-button"
            disabled={
              !config.baseUrl || !config.token || connection.state === "loading"
            }
            onClick={testConnection}
          >
            {connection.state === "loading" ? "A testar…" : "Testar ligação"}
          </button>
          <button
            className="primary-button"
            disabled={!config.baseUrl || !config.token}
            onClick={saveConfig}
          >
            Guardar configuração
          </button>
        </div>
        <div className="settings-foot">
          <div>
            <span>↻</span>
            <div>
              <strong>Sincronização completa</strong>
              <p>
                Importa todas as páginas de projetos e Issues abertas/fechadas.
              </p>
            </div>
          </div>
          <div>
            <span>⌁</span>
            <div>
              <strong>Sem dados fictícios</strong>
              <p>
                O dashboard só mostra informação recebida do GitLab ou do teu
                cronómetro.
              </p>
            </div>
          </div>
        </div>
      </section>
    );
  } else if (section === "calendar") {
    content = (
      <section className="panel settings-card">
        <div className="settings-title">
          <div className="outlook-logo">O</div>
          <div>
            <h2>Outlook / Teams</h2>
            <p>
              Importa reuniões do Outlook clássico instalado ou, em alternativa,
              pelo Microsoft Graph.
            </p>
          </div>
          <span
            className={
              outlookConnection.state === "ok"
                ? "status-pill connected"
                : "status-pill"
            }
          >
            {outlookConnection.state === "ok" ? "● Ligado" : "○ Por configurar"}
          </span>
        </div>
        <div className="security-note">
          <span>▣</span>
          <div>
            <strong>Outlook clássico — sem App Registration</strong>
            <p>
              No Windows executa <code>npm run outlook:bridge</code> e usa o
              botão abaixo. A ponte lê apenas os próximos 35 dias do calendário
              local.
            </p>
          </div>
        </div>
        {outlookConnection.message && (
          <div className={`connection-result ${outlookConnection.state}`}>
            <span>
              {outlookConnection.state === "ok"
                ? "✓"
                : outlookConnection.state === "error"
                  ? "!"
                  : "↻"}
            </span>
            {outlookConnection.message}
          </div>
        )}
        <div className="settings-actions">
          <button
            className="primary-button"
            disabled={syncingCalendar}
            onClick={syncLocalOutlook}
          >
            {syncingCalendar ? "A importar…" : "Importar do Outlook instalado"}
          </button>
        </div>
        <div className="settings-divider">
          <span>ou usa Microsoft Graph</span>
        </div>
        <div className="form-grid">
          <label className="form-field">
            <span>Microsoft Entra Tenant ID</span>
            <input
              value={outlookConfig.tenantId}
              onChange={(event) =>
                setOutlookConfig({
                  ...outlookConfig,
                  tenantId: event.target.value,
                })
              }
              placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
            />
          </label>
          <label className="form-field">
            <span>Application (Client) ID</span>
            <input
              value={outlookConfig.clientId}
              onChange={(event) =>
                setOutlookConfig({
                  ...outlookConfig,
                  clientId: event.target.value,
                })
              }
              placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
            />
          </label>
        </div>
        <label className="form-field">
          <span>Redirect URI a registar como SPA</span>
          <div className="readonly-field">
            {typeof window !== "undefined"
              ? window.location.origin
              : "http://localhost:3000"}
          </div>
        </label>
        <div className="settings-actions">
          <button
            className="secondary-button"
            disabled={!outlookConfig.tenantId || !outlookConfig.clientId}
            onClick={saveOutlookConfig}
          >
            Guardar
          </button>
          <button
            className="secondary-button"
            disabled={
              !outlookConfig.tenantId ||
              !outlookConfig.clientId ||
              outlookConnection.state === "loading"
            }
            onClick={connectOutlook}
          >
            Ligar Microsoft
          </button>
          <button
            className="secondary-button"
            disabled={
              !outlookConfig.tenantId ||
              !outlookConfig.clientId ||
              syncingCalendar
            }
            onClick={syncOutlook}
          >
            Importar via Graph
          </button>
        </div>
        {meetings.length > 0 && (
          <div className="meetings-section">
            <div className="panel-head">
              <div>
                <h2>Reuniões importadas</h2>
                <p>
                  Já estão na Timeline. Associa um projeto e clica no bloco para
                  criar a respetiva US.
                </p>
              </div>
              <span className="total-pill">{meetings.length}</span>
            </div>
            <div className="meeting-list">
              {meetings.map((meeting) => (
                <div className="meeting-row" key={meeting.id}>
                  <div className="meeting-time">
                    <strong>{formatDate(meeting.start.slice(0, 10))}</strong>
                    <span>
                      {meeting.start.slice(11, 16)}–{meeting.end.slice(11, 16)}
                    </span>
                  </div>
                  <div className="meeting-main">
                    <strong>{meeting.subject}</strong>
                    <small>
                      {meeting.joinUrl ? "Microsoft Teams" : "Outlook"}
                    </small>
                  </div>
                  <select
                    value={meeting.projectId ?? ""}
                    disabled={Boolean(meeting.linkedIssueId)}
                    onChange={(event) =>
                      associateMeeting(
                        meeting.id,
                        event.target.value
                          ? Number(event.target.value)
                          : undefined,
                      )
                    }
                  >
                    <option value="">Sem projeto</option>
                    {projects
                      .filter((project) => project.projectId > 0)
                      .map((project) => (
                        <option
                          key={project.projectId}
                          value={project.projectId}
                        >
                          {project.client} · {project.project}
                        </option>
                      ))}
                  </select>
                  {meeting.linkedIssueWebUrl && meeting.linkedIssueIid && (
                    <a
                      href={meeting.linkedIssueWebUrl}
                      target="_blank"
                      rel="noreferrer"
                      aria-label={`Abrir US ${meeting.linkedIssueIid}`}
                    >
                      #{meeting.linkedIssueIid}
                    </a>
                  )}
                  {(meeting.joinUrl || meeting.webUrl) && (
                    <a
                      href={meeting.joinUrl ?? meeting.webUrl}
                      target="_blank"
                      rel="noreferrer"
                      aria-label="Abrir reunião"
                    >
                      ↗
                    </a>
                  )}
                </div>
              ))}
            </div>
          </div>
        )}
        <div className="settings-foot">
          <div>
            <span>1</span>
            <div>
              <strong>Ponte local recomendada</strong>
              <p>
                Funciona com o Outlook clássico que já confirmaste estar
                acessível por COM.
              </p>
            </div>
          </div>
          <div>
            <span>2</span>
            <div>
              <strong>Dados no teu computador</strong>
              <p>
                Não usa Client Secret nem requer permissões de administrador no
                Entra.
              </p>
            </div>
          </div>
        </div>
      </section>
    );
  } else if (section === "capacity") {
    content = (
      <section className="panel settings-card">
        <div className="settings-title">
          <div className="capacity-logo">◷</div>
          <div>
            <h2>Capacidade de trabalho</h2>
            <p>
              Estas regras controlam o planeamento automático, os alertas e a
              escala visual.
            </p>
          </div>
          <span className="status-pill connected">
            {formatHours(capacity.dailyHours)} / dia
          </span>
        </div>
        <div className="form-grid">
          <label className="form-field">
            <span>Horas por dia</span>
            <input
              type="number"
              min="1"
              max="16"
              step="0.5"
              value={capacity.dailyHours}
              onChange={(event) =>
                setCapacity({
                  ...capacity,
                  dailyHours: Math.min(
                    16,
                    Math.max(1, Number(event.target.value) || 1),
                  ),
                })
              }
            />
            <small>É o limite antes de um dia ficar em sobrecarga.</small>
          </label>
          <label className="form-field">
            <span>Hora de início</span>
            <input
              type="number"
              min="0"
              max="23"
              step="0.5"
              value={capacity.startHour}
              onChange={(event) =>
                setCapacity({
                  ...capacity,
                  startHour: Math.min(
                    23,
                    Math.max(0, Number(event.target.value) || 0),
                  ),
                })
              }
            />
            <small>
              Fim previsto:{" "}
              {String(
                Math.floor((capacity.startHour + capacity.dailyHours) % 24),
              ).padStart(2, "0")}
              :{(capacity.startHour + capacity.dailyHours) % 1 ? "30" : "00"}.
            </small>
          </label>
        </div>
        <div className="workdays-field">
          <span>Dias de trabalho</span>
          <div>
            {WEEKDAYS.map((label, day) => (
              <button
                key={label}
                className={capacity.workDays.includes(day) ? "active" : ""}
                onClick={() => {
                  const selected = capacity.workDays.includes(day);
                  if (selected && capacity.workDays.length === 1) return;
                  setCapacity({
                    ...capacity,
                    workDays: selected
                      ? capacity.workDays.filter((value) => value !== day)
                      : [...capacity.workDays, day].sort(),
                  });
                }}
              >
                {label}
              </button>
            ))}
          </div>
          <small>A Timeline semanal apresenta apenas estes dias.</small>
        </div>
        <div className="settings-actions">
          <button className="primary-button" onClick={saveCapacity}>
            Guardar e replanear
          </button>
        </div>
        <div className="settings-foot">
          <div>
            <span>↻</span>
            <div>
              <strong>Replaneamento imediato</strong>
              <p>
                As tarefas flexíveis são novamente distribuídas pela capacidade
                disponível.
              </p>
            </div>
          </div>
          <div>
            <span>↔</span>
            <div>
              <strong>Escala coerente</strong>
              <p>
                A largura de cada tarefa corresponde às horas sobre a capacidade
                diária.
              </p>
            </div>
          </div>
        </div>
      </section>
    );
  } else {
    content = (
      <section className="panel settings-card">
        <div className="settings-title">
          <div className="clients-logo">◉</div>
          <div>
            <h2>Clientes e projetos</h2>
            <p>Projetos reais encontrados através do teu utilizador GitLab.</p>
          </div>
          <span className="status-pill connected">
            {projectPreferences.filter((project) => project.enabled).length}{" "}
            ativos
          </span>
        </div>
        {projectPreferences.length ? (
          <div className="projects-config-list">
            {[...projectPreferences]
              .sort(
                (a, b) =>
                  a.client.localeCompare(b.client) ||
                  a.project.localeCompare(b.project),
              )
              .map((project) => (
                <div
                  className={`project-config-row ${project.enabled ? "" : "disabled"}`}
                  key={project.projectId}
                >
                  <label className="project-toggle">
                    <input
                      type="checkbox"
                      checked={project.enabled}
                      onChange={(event) =>
                        updateProjectPreference(project.projectId, {
                          enabled: event.target.checked,
                        })
                      }
                    />
                    <span />
                  </label>
                  <input
                    className="project-color"
                    type="color"
                    value={project.color}
                    aria-label={`Cor de ${project.project}`}
                    onChange={(event) =>
                      updateProjectPreference(project.projectId, {
                        color: event.target.value,
                      })
                    }
                  />
                  <div className="project-identity">
                    <strong>{project.project}</strong>
                    <small>GitLab project #{project.projectId}</small>
                  </div>
                  <label>
                    <span>Cliente</span>
                    <input
                      value={project.client}
                      onChange={(event) =>
                        updateProjectPreference(project.projectId, {
                          client: event.target.value,
                        })
                      }
                    />
                  </label>
                </div>
              ))}
          </div>
        ) : (
          <div className="empty-state">
            <span>↻</span>
            <strong>Sincroniza primeiro o GitLab</strong>
            <p>
              Os clientes e projetos aparecerão aqui automaticamente; não
              existem registos fictícios.
            </p>
          </div>
        )}
        <div className="settings-actions">
          <button
            className="primary-button"
            onClick={syncGitLab}
            disabled={syncing}
          >
            {syncing ? "A sincronizar…" : "Sincronizar projetos"}
          </button>
        </div>
        <div className="settings-foot">
          <div>
            <span>◉</span>
            <div>
              <strong>Cliente ajustável</strong>
              <p>Podes corrigir o cliente sugerido a partir do namespace.</p>
            </div>
          </div>
          <div>
            <span>○</span>
            <div>
              <strong>Projetos opcionais</strong>
              <p>
                Desativa projetos irrelevantes e sincroniza para atualizar todos
                os módulos.
              </p>
            </div>
          </div>
        </div>
      </section>
    );
  }

  return (
    <div className="page settings-page">
      <PageIntro
        eyebrow="CONFIGURAÇÃO"
        title={titles[section][0]}
        description={titles[section][1]}
      />
      <div className="settings-layout">
        <nav className="settings-nav">
          <button
            className={section === "gitlab" ? "active" : ""}
            onClick={() => setSection("gitlab")}
          >
            <span>⌘</span>
            <div>
              <strong>Integração GitLab</strong>
              <small>Issues, projetos e spent time</small>
            </div>
          </button>
          <button
            className={section === "calendar" ? "active" : ""}
            onClick={() => setSection("calendar")}
          >
            <span>▦</span>
            <div>
              <strong>Outlook / Teams</strong>
              <small>Reuniões e projetos</small>
            </div>
          </button>
          <button
            className={section === "capacity" ? "active" : ""}
            onClick={() => setSection("capacity")}
          >
            <span>◷</span>
            <div>
              <strong>Capacidade</strong>
              <small>Horário e dias úteis</small>
            </div>
          </button>
          <button
            className={section === "clients" ? "active" : ""}
            onClick={() => setSection("clients")}
          >
            <span>◉</span>
            <div>
              <strong>Clientes</strong>
              <small>Projetos e cores</small>
            </div>
          </button>
        </nav>
        {content}
      </div>
    </div>
  );
}

function projectTotals(tasks: Task[]) {
  const grouped = new Map<string, number>();
  tasks.forEach((task) =>
    grouped.set(task.project, (grouped.get(task.project) ?? 0) + task.estimate),
  );
  return [...grouped.entries()].sort((a, b) => b[1] - a[1]);
}

function exportCsv(logs: WorkLog[]) {
  const csv = [
    "Data;Cliente;Projeto;Tarefa;Tipo;Horas",
    ...logs.map((log) =>
      [log.dateKey, log.client, log.project, log.task, log.type, log.hours]
        .map((value) => `"${String(value).replace(/"/g, '""')}"`)
        .join(";"),
    ),
  ].join("\n");
  const link = document.createElement("a");
  link.href = URL.createObjectURL(
    new Blob([csv], { type: "text/csv;charset=utf-8" }),
  );
  link.download = "trabalho-realizado.csv";
  link.click();
  URL.revokeObjectURL(link.href);
}

function uniqueBy<T>(items: T[], getter: (item: T) => string) {
  const seen = new Set<string>();
  return items.filter((item) => {
    const value = getter(item);
    if (seen.has(value)) return false;
    seen.add(value);
    return true;
  });
}
