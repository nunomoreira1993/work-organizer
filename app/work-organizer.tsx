"use client";

import { PublicClientApplication } from "@azure/msal-browser";
import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type Dispatch,
  type DragEvent,
  type SetStateAction,
} from "react";

type View = "today" | "issues" | "team" | "timeline" | "completed" | "analytics" | "settings";
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

type GitLabAssignee = {
  id: number;
  name?: string;
  username?: string;
  avatarUrl?: string;
  webUrl?: string;
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
  closedAt?: string | null;
  type: string;
  phase: WorkPhase;
  description?: string;
  priority: "Alta" | "Média" | "Normal";
  fixed?: boolean;
  webUrl?: string;
  source: TaskSource;
  hasEstimate?: boolean;
  labels?: GitLabLabel[];
  assignees?: GitLabAssignee[];
};

type IssueRecord = {
  id: string;
  projectId: number;
  iid: number;
  title: string;
  description?: string;
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
  closedAt?: string | null;
  type: string;
  priority: Task["priority"];
  webUrl?: string;
  hasEstimate: boolean;
  labels: GitLabLabel[];
  assignees: GitLabAssignee[];
};

type AllocationDraft = {
  issueId: string;
  projectId?: number;
  labels?: string[];
  issueDescription?: string;
  allocationId?: string;
  phase: WorkPhase;
  hours: number;
  date: string;
  start: number;
  distribution: "manual" | "automatic";
  description: string;
  issueState?: "opened" | "closed";
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
  relatedIssueIds: string[];
  meetingId?: string;
};

type WorkLog = {
  id: string;
  dateKey: string;
  client: string;
  project: string;
  task: string;
  description?: string;
  type: string;
  hours: number;
  estimate?: number;
  webUrl?: string;
  projectId?: number;
  issueIid?: number;
  source?: "gitlab" | "timer" | "closed" | "timeline";
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

type TeamMember = {
  id: string;
  name: string;
  role?: string;
};

type TeamAbsenceType = "vacation" | "absence" | "training" | "other";

type TeamAllocation = {
  id: string;
  memberId: string;
  issueId?: string;
  customTitle?: string;
  customColor?: string;
  absenceType?: TeamAbsenceType;
  weekStart: string;
  startSlot: number;
  hours: number;
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
  teamMembers: TeamMember[];
  teamAllocations: TeamAllocation[];
};

type GitLabIssue = {
  project_id: number;
  iid: number;
  title: string;
  description?: string | null;
  state: "opened" | "closed";
  web_url?: string;
  due_date?: string | null;
  closed_at?: string | null;
  updated_at?: string;
  labels?: string[];
  assignees?: {
    id: number;
    name?: string;
    username?: string;
    avatar_url?: string;
    web_url?: string;
  }[];
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
const DEFAULT_TEAM_MEMBERS: TeamMember[] = [
  { id: "neuza-santos", name: "Neuza Santos" },
  { id: "nuno-moreira", name: "Nuno Moreira" },
  { id: "andre-carvalho", name: "André Carvalho" },
  { id: "fabio-cerqueira", name: "Fabio Cerqueira", role: "Dev 5" },
  { id: "goncalo-duarte", name: "Gonçalo Duarte", role: "Dev 5" },
];
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

function assigneesForIssue(issue: GitLabIssue): GitLabAssignee[] {
  return (issue.assignees ?? []).map((assignee) => ({
    id: assignee.id,
    name: assignee.name,
    username: assignee.username,
    avatarUrl: assignee.avatar_url,
    webUrl: assignee.web_url,
  }));
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

const RELATED_STOP_WORDS = new Set([
  "para", "com", "sem", "uma", "uns", "das", "dos", "que", "por", "nos",
  "nas", "este", "esta", "isto", "isso", "como", "mais", "menos", "deve",
  "ser", "ter", "fazer", "criar", "adicionar", "alterar", "quando", "onde",
  "the", "and", "for", "from", "with", "this", "that", "into", "issue",
]);

function relatedTerms(text: string) {
  const plain = text
    .replace(/`[^`]*`/g, " ")
    .replace(/https?:\/\/\S+/g, " ")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
  return new Set(
    plain
      .split(/[^a-z0-9]+/)
      .filter((term) => term.length >= 3 && !RELATED_STOP_WORDS.has(term))
      .map((term) =>
        term.length > 6
          ? term.replace(
              /(coes|cao|mente|amentos|amento|idades|idade|ados|adas|idos|idas|ar|er|ir|s)$/,
              "",
            )
          : term,
      )
      .filter((term) => term.length >= 3),
  );
}

function relatedIssueScore(
  title: string,
  description: string,
  selectedLabels: string[],
  issue: IssueRecord,
) {
  const sourceTitle = relatedTerms(title);
  const sourceDescription = relatedTerms(description);
  const candidateTitle = relatedTerms(issue.title);
  const candidateDescription = relatedTerms(issue.description ?? "");
  const matched = new Set<string>();
  let score = 0;
  candidateTitle.forEach((term) => {
    if (sourceTitle.has(term)) {
      score += 5;
      matched.add(term);
    } else if (sourceDescription.has(term)) {
      score += 3;
      matched.add(term);
    }
  });
  candidateDescription.forEach((term) => {
    if (sourceTitle.has(term)) {
      score += 2;
      matched.add(term);
    } else if (sourceDescription.has(term)) {
      score += 1;
      matched.add(term);
    }
  });
  issue.labels.forEach((item) => {
    if (selectedLabels.includes(item.name)) score += 3;
  });
  return { score, matches: [...matched].slice(0, 4) };
}

const nav: { id: View; label: string; icon: string }[] = [
  { id: "today", label: "Hoje", icon: "⌂" },
  { id: "issues", label: "Explorar US", icon: "☷" },
  { id: "team", label: "Alocação da equipa", icon: "▦" },
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

function writeStorage(key: string, value: unknown) {
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch {
    return false;
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
  const [teamMembers, setTeamMembers] = useState<TeamMember[]>(
    DEFAULT_TEAM_MEMBERS,
  );
  const [teamAllocations, setTeamAllocations] = useState<TeamAllocation[]>([]);
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
    const requested = new URLSearchParams(window.location.search).get("view");
    if (
      requested &&
      [...nav.map((item) => item.id), "settings"].includes(requested as View)
    ) {
      setView(requested as View);
    }
  }, []);

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
        teamMembers: readStorage(
          "work-organizer.team-members",
          DEFAULT_TEAM_MEMBERS,
        ),
        teamAllocations: readStorage("work-organizer.team-allocations", []),
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
          if (Array.isArray(data.state.teamMembers))
            setTeamMembers(data.state.teamMembers);
          if (Array.isArray(data.state.teamAllocations))
            setTeamAllocations(data.state.teamAllocations);
        } else {
          if (!cancelled) {
            setTasks(localState.tasks);
            setIssues(localState.issues);
            setLogs(localState.logs);
            setMeetings(localState.meetings);
            setCapacity(localState.capacity);
            setProjectPreferences(localState.projectPreferences);
            setLabelCatalog(localState.labelCatalog);
            setTeamMembers(localState.teamMembers);
            setTeamAllocations(localState.teamAllocations);
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
          setTeamMembers(localState.teamMembers);
          setTeamAllocations(localState.teamAllocations);
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
    const localCacheFailed = [
      writeStorage("work-organizer.tasks", tasks),
      writeStorage("work-organizer.issues", issues),
      writeStorage("work-organizer.logs", logs),
      writeStorage("work-organizer.meetings", meetings),
      writeStorage("work-organizer.capacity", capacity),
      writeStorage("work-organizer.projects", projectPreferences),
      writeStorage("work-organizer.labels", labelCatalog),
      writeStorage("work-organizer.team-members", teamMembers),
      writeStorage("work-organizer.team-allocations", teamAllocations),
    ].some((saved) => !saved);
    if (localCacheFailed) {
      setToast(
        "O cache deste browser ficou sem espaço. Os dados continuam a ser guardados na base de dados.",
      );
    }
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
          teamMembers,
          teamAllocations,
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
    teamMembers,
    teamAllocations,
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

  // Closing a US finishes it in GitLab, but its allocation remains part of the
  // planning history and must stay visible in the Timeline.
  const planningTasks = tasks;
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
          description: issue.description ?? "",
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
          assignees: assigneesForIssue(issue),
          due: issue.due_date ? formatDate(issue.due_date) : "Sem prazo",
          dueDate: issue.due_date,
          closedAt: issue.closed_at,
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
                assignees: issue.assignees,
                closedAt: issue.closedAt,
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

  async function updateIssueState(
    issue: IssueRecord,
    state: "opened" | "closed",
  ) {
    if (issue.state === state) return issue;
    if (!config.token)
      throw new Error("Configura primeiro a ligação ao GitLab.");
    const data = await callGitLab("updateIssueState", {
      projectId: issue.projectId,
      issueIid: issue.iid,
      state,
    });
    const remote = data.issue as GitLabIssue;
    setIssues((current) =>
      current.map((item) =>
        item.id === issue.id
          ? {
              ...item,
              state: remote.state,
              closedAt: remote.closed_at,
            }
          : item,
      ),
    );
    setTasks((current) =>
      current.map((task) =>
        task.rootId === issue.id
          ? { ...task, closedAt: remote.closed_at }
          : task,
      ),
    );
    return { ...issue, state: remote.state, closedAt: remote.closed_at };
  }

  async function moveIssueProject(issue: IssueRecord, toProjectId: number) {
    if (issue.projectId === toProjectId) return issue;
    if (!config.token)
      throw new Error("Configura primeiro a ligação ao GitLab.");
    const destination = projectPreferences.find(
      (project) => project.projectId === toProjectId,
    );
    if (!destination) throw new Error("O projeto de destino já não está disponível.");
    const data = await callGitLab("moveIssue", {
      projectId: issue.projectId,
      issueIid: issue.iid,
      toProjectId,
    });
    const remote = data.issue as GitLabIssue;
    const nextId = `gitlab-${remote.project_id}-${remote.iid}`;
    const moved: IssueRecord = {
      ...issue,
      id: nextId,
      projectId: remote.project_id,
      iid: remote.iid,
      title: remote.title,
      description: remote.description ?? issue.description ?? "",
      client: destination.client,
      project: destination.project,
      color: destination.color,
      state: remote.state,
      webUrl: remote.web_url,
      due: remote.due_date ? formatDate(remote.due_date) : "Sem prazo",
      dueDate: remote.due_date,
      closedAt: remote.closed_at,
      labels: labelsForIssue(
        remote.labels,
        labelCatalog[String(remote.project_id)] ?? [],
      ),
      assignees: assigneesForIssue(remote),
      assigneeCount: remote.assignees?.length ?? 0,
    };
    setIssues((current) =>
      current.map((item) => (item.id === issue.id ? moved : item)),
    );
    setTasks((current) =>
      current.map((task) =>
        task.rootId === issue.id
          ? {
              ...task,
              rootId: nextId,
              projectId: moved.projectId,
              iid: moved.iid,
              title: moved.title,
              client: moved.client,
              project: moved.project,
              color: moved.color,
              webUrl: moved.webUrl,
              due: moved.due,
              dueDate: moved.dueDate,
              closedAt: moved.closedAt,
              labels: moved.labels,
              assignees: moved.assignees,
            }
          : task,
      ),
    );
    setLogs((current) =>
      current.map((log) =>
        log.projectId === issue.projectId && log.issueIid === issue.iid
          ? {
              ...log,
              projectId: moved.projectId,
              issueIid: moved.iid,
              client: moved.client,
              project: moved.project,
              webUrl: moved.webUrl,
            }
          : log,
      ),
    );
    setMeetings((current) =>
      current.map((meeting) =>
        meeting.linkedIssueId === issue.id
          ? {
              ...meeting,
              projectId: moved.projectId,
              linkedIssueId: nextId,
              linkedIssueIid: moved.iid,
              linkedIssueWebUrl: moved.webUrl,
            }
          : meeting,
      ),
    );
    return moved;
  }

  async function updateIssueLabels(issue: IssueRecord, labelNames: string[]) {
    const currentNames = issue.labels.map((label) => label.name).sort();
    const nextNames = [...new Set(labelNames)].sort();
    if (
      currentNames.length === nextNames.length &&
      currentNames.every((name, index) => name === nextNames[index])
    )
      return issue;
    if (!config.token)
      throw new Error("Configura primeiro a ligação ao GitLab.");
    const data = await callGitLab("updateIssueLabels", {
      projectId: issue.projectId,
      issueIid: issue.iid,
      labels: nextNames,
    });
    const remote = data.issue as GitLabIssue;
    const labels = labelsForIssue(
      remote.labels,
      labelCatalog[String(issue.projectId)] ?? [],
    );
    const updated = {
      ...issue,
      labels,
      type: taskType(remote.labels),
      priority: taskPriority(remote.labels),
    };
    setIssues((current) =>
      current.map((item) => (item.id === issue.id ? updated : item)),
    );
    setTasks((current) =>
      current.map((task) =>
        task.rootId === issue.id
          ? {
              ...task,
              labels,
              type: updated.type,
              priority: updated.priority,
            }
          : task,
      ),
    );
    return updated;
  }

  async function updateIssueDescription(
    issue: IssueRecord,
    description: string,
  ) {
    if ((issue.description ?? "") === description) return issue;
    if (!config.token)
      throw new Error("Configura primeiro a ligação ao GitLab.");
    const data = await callGitLab("updateIssueDescription", {
      projectId: issue.projectId,
      issueIid: issue.iid,
      description,
    });
    const remote = data.issue as GitLabIssue;
    const updated = {
      ...issue,
      description: remote.description ?? description,
    };
    setIssues((current) =>
      current.map((item) => (item.id === issue.id ? updated : item)),
    );
    return updated;
  }

  async function renderGitLabMarkdown(markdown: string, projectId?: number) {
    if (!config.token)
      throw new Error("Configura primeiro a ligação ao GitLab.");
    const data = await callGitLab("renderMarkdown", { markdown, projectId });
    return String(data.html ?? "");
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

  function mergeImportedMeetings(
    events: CalendarMeeting[],
    importStartKey: string,
    importEndKey: string,
  ) {
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
    const isInImportRange = (date: string) =>
      date >= importStartKey && date < importEndKey;
    // This is a complete calendar slice. Any saved Outlook meeting inside the
    // same range but absent from the response was cancelled or is no longer
    // accepted, so it must disappear from the calendar as well.
    const merged = [
      ...meetings.filter(
        (meeting) => !isInImportRange(meeting.start.slice(0, 10)),
      ),
      ...imported,
    ].sort((left, right) => left.start.localeCompare(right.start));
    setMeetings(merged);
    setTasks((current) => [
      ...current.filter((task) => {
        if (importedRootIds.has(task.rootId)) return false;
        if (task.source === "outlook" && isInImportRange(task.date))
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
      const importStartKey = `${todayKey.slice(0, 8)}01`;
      const importEndKey = addDays(todayKey, 36);
      const params = new URLSearchParams({
        startDateTime: `${importStartKey}T00:00:00Z`,
        endDateTime: `${importEndKey}T00:00:00Z`,
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
      const imported = mergeImportedMeetings(
        events,
        importStartKey,
        importEndKey,
      );
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
      const importStartKey = `${todayKey.slice(0, 8)}01`;
      const importEndKey = addDays(todayKey, 36);
      const importDays = Math.round(
        (fromKey(importEndKey).getTime() - fromKey(importStartKey).getTime()) /
          86_400_000,
      );
      const response = await fetch(`http://127.0.0.1:47831/calendar?start=${importStartKey}&days=${importDays}`, {
        signal: AbortSignal.timeout(30000),
      });
      const data = await response.json();
      if (!response.ok || !data.ok)
        throw new Error(
          data.error ?? "A ponte local do Outlook devolveu um erro.",
        );
      const imported = mergeImportedMeetings(
        data.events ?? [],
        importStartKey,
        importEndKey,
      );
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
        assignees: issue.assignees,
        closedAt: issue.closedAt,
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

  function saveAllocation(draft: AllocationDraft, issueOverride?: IssueRecord) {
    const issue =
      issueOverride ?? issues.find((item) => item.id === draft.issueId);
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
      relatedIssues: draft.relatedIssueIds.flatMap((id) => {
        const related = issues.find((issue) => issue.id === id);
        return related
          ? [{ projectId: related.projectId, issueIid: related.iid }]
          : [];
      }),
    });
    const remote = data.issue as GitLabIssue;
    const remoteProject = data.project as GitLabProject;
    const failedRelatedIssues = Number(data.failedRelatedIssues ?? 0);
    const relationSuffix = failedRelatedIssues
      ? ` ${failedRelatedIssues} relação(ões) não puderam ser criada(s) no GitLab.`
      : "";
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
      description: remote.description ?? draft.description,
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
      assignees: assigneesForIssue(remote),
      due: remote.due_date ? formatDate(remote.due_date) : "Sem prazo",
      dueDate: remote.due_date,
      closedAt: remote.closed_at,
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
                assignees: issue.assignees,
                closedAt: issue.closedAt,
              }
            : task,
        ),
      );
      notify(
        `US #${issue.iid} criada para a reunião. O cronómetro passa a registar spent no GitLab.${relationSuffix}`,
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
      notify(
        `Tarefa #${issue.iid} criada no GitLab e adicionada à Timeline.${relationSuffix}`,
      );
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
                : view === "issues"
                  ? "Explorar US"
                  : view === "team"
                    ? "Alocação da equipa"
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
            updateIssueState={updateIssueState}
            moveIssueProject={moveIssueProject}
            updateIssueLabels={updateIssueLabels}
            updateIssueDescription={updateIssueDescription}
            renderMarkdown={renderGitLabMarkdown}
            createIssueAndAllocate={createIssueAndAllocate}
            deleteAllocation={deleteAllocation}
            autoPlan={autoPlan}
            todayKey={todayKey}
            capacity={capacity}
          />
        )}
        {view === "issues" && <IssuesView issues={issues} tasks={tasks} />}
        {view === "team" && (
          <TeamAllocationView
            issues={issues}
            members={teamMembers}
            allocations={teamAllocations}
            setMembers={setTeamMembers}
            setAllocations={setTeamAllocations}
            todayKey={todayKey}
            setToast={setToast}
          />
        )}
        {view === "completed" && (
          <CompletedView tasks={tasks} todayKey={todayKey} />
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

function MarkdownEditor({
  value,
  onChange,
  renderMarkdown,
  rows = 7,
  placeholder,
  autoFocus = false,
}: {
  value: string;
  onChange: (value: string) => void;
  renderMarkdown: (markdown: string, projectId?: number) => Promise<string>;
  rows?: number;
  placeholder?: string;
  autoFocus?: boolean;
}) {
  const [mode, setMode] = useState<"write" | "preview">("write");
  const [preview, setPreview] = useState("");
  const [previewSource, setPreviewSource] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function showPreview() {
    setMode("preview");
    if (value === previewSource && !error) return;
    setLoading(true);
    setError("");
    try {
      setPreview(await renderMarkdown(value));
      setPreviewSource(value);
    } catch (previewError) {
      setError(
        previewError instanceof Error
          ? previewError.message
          : "Não foi possível gerar a pré-visualização.",
      );
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="markdown-editor">
      <div className="markdown-editor-tabs">
        <button
          type="button"
          className={mode === "write" ? "active" : ""}
          onClick={() => setMode("write")}
        >
          Escrever
        </button>
        <button
          type="button"
          className={mode === "preview" ? "active" : ""}
          onClick={() => void showPreview()}
        >
          Pré-visualizar
        </button>
        <span>Markdown do GitLab</span>
      </div>
      {mode === "write" ? (
        <textarea
          autoFocus={autoFocus}
          rows={rows}
          value={value}
          onChange={(event) => onChange(event.target.value)}
          placeholder={placeholder}
        />
      ) : (
        <div className="markdown-preview">
          {loading ? (
            <p className="markdown-preview-state">A gerar pré-visualização…</p>
          ) : error ? (
            <p className="markdown-preview-error">{error}</p>
          ) : preview ? (
            <div dangerouslySetInnerHTML={{ __html: preview }} />
          ) : (
            <p className="markdown-preview-state">Sem conteúdo para apresentar.</p>
          )}
        </div>
      )}
    </div>
  );
}

function IssueAssignees({
  assignees,
  compact = false,
}: {
  assignees?: GitLabAssignee[];
  compact?: boolean;
}) {
  if (!assignees?.length) {
    return compact ? null : (
      <div className="issue-assignees unassigned">Sem responsável</div>
    );
  }
  return (
    <div className={`issue-assignees ${compact ? "compact" : ""}`}>
      {assignees.map((assignee) => {
        const label = assignee.name?.trim() || assignee.username || "Utilizador";
        const initials = label
          .split(/\s+/)
          .slice(0, 2)
          .map((part) => part[0])
          .join("")
          .toUpperCase();
        return (
          <span className="issue-assignee" key={assignee.id} title={label}>
            <i
              aria-hidden="true"
              className={assignee.avatarUrl ? "has-avatar" : ""}
              style={
                assignee.avatarUrl
                  ? { backgroundImage: `url("${assignee.avatarUrl.replace(/"/g, "%22")}")` }
                  : undefined
              }
            >
              {!assignee.avatarUrl && initials}
            </i>
            <b>{label}</b>
          </span>
        );
      })}
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
                        <IssueAssignees assignees={task.assignees} compact />
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
  updateIssueState,
  moveIssueProject,
  updateIssueLabels,
  updateIssueDescription,
  renderMarkdown,
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
  saveAllocation: (draft: AllocationDraft, issueOverride?: IssueRecord) => void;
  updateIssueState: (
    issue: IssueRecord,
    state: "opened" | "closed",
  ) => Promise<IssueRecord>;
  moveIssueProject: (
    issue: IssueRecord,
    toProjectId: number,
  ) => Promise<IssueRecord>;
  updateIssueLabels: (
    issue: IssueRecord,
    labels: string[],
  ) => Promise<IssueRecord>;
  updateIssueDescription: (
    issue: IssueRecord,
    description: string,
  ) => Promise<IssueRecord>;
  renderMarkdown: (markdown: string, projectId?: number) => Promise<string>;
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
  >("all");
  const [assigneeFilter, setAssigneeFilter] = useState("all");
  const [hiddenProjectIds, setHiddenProjectIds] = useState<number[]>([]);
  const [hiddenLabels, setHiddenLabels] = useState<string[]>([]);
  const [editor, setEditor] = useState<AllocationDraft | null>(null);
  const [createEditor, setCreateEditor] = useState<CreateIssueDraft | null>(
    null,
  );
  const [creatingIssue, setCreatingIssue] = useState(false);
  const [labelsLoading, setLabelsLoading] = useState(false);
  const [createError, setCreateError] = useState("");
  const [editorError, setEditorError] = useState("");
  const [savingEditor, setSavingEditor] = useState(false);
  const [projectSearch, setProjectSearch] = useState("");
  const [editorProjectSearch, setEditorProjectSearch] = useState("");
  const [editorLabelsLoading, setEditorLabelsLoading] = useState(false);
  const days = visibleDaysFor(anchor, zoom, capacity.workDays);
  const openedIssues = issues.filter((issue) => issue.state === "opened");
  const closedIssueIds = new Set(
    issues
      .filter((issue) => issue.state === "closed")
      .map((issue) => issue.id),
  );
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
  const assigneeOptions = useMemo(() => {
    const byId = new Map<number, GitLabAssignee>();
    issues
      .filter((issue) => issue.state === "opened")
      .forEach((issue) =>
        (issue.assignees ?? []).forEach((assignee) =>
          byId.set(assignee.id, assignee),
        ),
      );
    return [...byId.values()].sort((left, right) =>
      (left.name || left.username || "").localeCompare(
        right.name || right.username || "",
      ),
    );
  }, [issues]);
  useEffect(() => {
    if (
      assigneeFilter !== "all" &&
      assigneeFilter !== "unassigned" &&
      !assigneeOptions.some(
        (assignee) => String(assignee.id) === assigneeFilter,
      )
    ) {
      setAssigneeFilter("all");
    }
  }, [assigneeFilter, assigneeOptions]);
  const backlogProjectOptions = useMemo(
    () =>
      uniqueBy(
        issues.filter((issue) => issue.state === "opened"),
        (issue) => String(issue.projectId),
      ).sort(
        (left, right) =>
          left.client.localeCompare(right.client) ||
          left.project.localeCompare(right.project),
      ),
    [issues],
  );
  const backlogLabelOptions = useMemo(() => {
    const byName = new Map<string, GitLabLabel>();
    issues
      .filter((issue) => issue.state === "opened")
      .forEach((issue) =>
        (issue.labels ?? []).forEach((label) => byName.set(label.name, label)),
      );
    return [...byName.values()].sort((left, right) =>
      left.name.localeCompare(right.name),
    );
  }, [issues]);
  const projectSearchNeedle = projectSearch.trim().toLowerCase();
  const selectableProjects = gitlabProjects
    .filter((project) => project.projectId > 0)
    .sort(
      (left, right) =>
        left.client.localeCompare(right.client) ||
        left.project.localeCompare(right.project),
    );
  const matchingProjects = projectSearchNeedle
    ? selectableProjects
        .filter((project) =>
          `${project.client} ${project.project} ${project.projectId}`
            .toLowerCase()
            .includes(projectSearchNeedle),
        )
        .slice(0, 50)
    : [];
  const editorProjectNeedle = editorProjectSearch.trim().toLowerCase();
  const matchingEditorProjects = editorProjectNeedle
    ? selectableProjects
        .filter((project) =>
          `${project.client} ${project.project} ${project.projectId}`
            .toLowerCase()
            .includes(editorProjectNeedle),
        )
        .slice(0, 50)
    : [];
  const selectedEditorProject = editor
    ? selectableProjects.find(
        (project) => project.projectId === editor.projectId,
      )
    : undefined;
  const selectedCreateProject = createEditor
    ? selectableProjects.find(
        (project) => project.projectId === createEditor.projectId,
      )
    : undefined;
  const relatedSuggestions = useMemo(() => {
    if (!createEditor) return [];
    const hasEnoughText =
      relatedTerms(`${createEditor.title} ${createEditor.description}`).size > 0;
    if (!hasEnoughText) return [];
    return issues
      .map((issue) => ({
        issue,
        ...relatedIssueScore(
          createEditor.title,
          createEditor.description,
          createEditor.labels,
          issue,
        ),
      }))
      .filter((item) => item.score >= 2)
      .sort(
        (left, right) =>
          right.score - left.score ||
          Number(right.issue.state === "opened") -
            Number(left.issue.state === "opened"),
      )
      .slice(0, 6);
  }, [createEditor, issues]);
  const selectedRelatedIssues = createEditor
    ? createEditor.relatedIssueIds
        .map((id) => issues.find((issue) => issue.id === id))
        .filter((issue): issue is IssueRecord => Boolean(issue))
    : [];
  const backlog = openedIssues
    .filter((issue) => {
      const planned = plannedByIssue.get(issue.id) ?? 0;
      if (backlogFilter === "unplanned" && planned > 0) return false;
      if (backlogFilter === "planned" && planned <= 0) return false;
      if (hiddenProjectIds.includes(issue.projectId)) return false;
      if (
        issue.labels?.some((label) => hiddenLabels.includes(label.name))
      )
        return false;
      if (assigneeFilter === "unassigned" && issue.assignees?.length)
        return false;
      if (
        assigneeFilter !== "all" &&
        assigneeFilter !== "unassigned" &&
        !issue.assignees?.some(
          (assignee) => String(assignee.id) === assigneeFilter,
        )
      )
        return false;
      return (
        !needle ||
        `${issue.title} ${issue.client} ${issue.project} ${issue.iid} ${(issue.labels ?? []).map((label) => label.name).join(" ")} ${(issue.assignees ?? []).map((assignee) => `${assignee.name ?? ""} ${assignee.username ?? ""}`).join(" ")}`
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
    setEditorError("");
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
    setEditorError("");
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
      setProjectSearch("");
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
        relatedIssueIds: [],
        meetingId: task.rootId.slice("outlook-".length),
      });
      refreshProjectLabels(project.projectId);
      return;
    }
    setEditor({
      issueId: segments[0].rootId,
      projectId: segments[0].projectId,
      labels:
        issues
          .find((issue) => issue.id === segments[0].rootId)
          ?.labels.map((label) => label.name) ?? [],
      issueDescription:
        issues.find((issue) => issue.id === segments[0].rootId)?.description ??
        "",
      allocationId,
      issueState:
        issues.find((issue) => issue.id === segments[0].rootId)?.state ??
        "opened",
      phase: segments[0].phase,
      hours: segments.reduce((sum, task) => sum + task.estimate, 0),
      date: segments[0].date,
      start: segments[0].start,
      distribution: segments.length > 1 ? "automatic" : "manual",
      description: segments[0].description ?? "",
    });
    setEditorProjectSearch("");
    refreshEditorProjectLabels(segments[0].projectId);
  }

  async function submitEditor() {
    if (!editor || !editingIssue || !editor.description.trim()) return;
    setSavingEditor(true);
    try {
      let savedIssue = editingIssue;
      if (
        editor.allocationId &&
        editor.projectId &&
        editor.projectId !== savedIssue.projectId
      ) {
        savedIssue = await moveIssueProject(savedIssue, editor.projectId);
      }
      if (editor.allocationId && editor.labels) {
        savedIssue = await updateIssueLabels(savedIssue, editor.labels);
      }
      if (editor.allocationId && editor.issueDescription !== undefined) {
        savedIssue = await updateIssueDescription(
          savedIssue,
          editor.issueDescription,
        );
      }
      if (editor.allocationId && editor.issueState) {
        savedIssue = await updateIssueState(savedIssue, editor.issueState);
      }
      saveAllocation({ ...editor, issueId: savedIssue.id }, savedIssue);
      setEditor(null);
    } catch (error) {
      setEditorError(
        error instanceof Error
          ? error.message
          : "Não foi possível atualizar a US no GitLab.",
      );
    } finally {
      setSavingEditor(false);
    }
  }

  function openCreateIssue(date = todayKey) {
    const firstProject = gitlabProjects.find(
      (project) => project.projectId > 0,
    );
    if (!firstProject) return;
    setCreateError("");
    setProjectSearch("");
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
      relatedIssueIds: [],
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

  function refreshEditorProjectLabels(projectId: number) {
    setEditorLabelsLoading(true);
    setEditorError("");
    void loadProjectLabels(projectId)
      .catch((error) => {
        setEditorError(
          error instanceof Error
            ? error.message
            : "Não foi possível carregar as labels do projeto.",
        );
      })
      .finally(() => setEditorLabelsLoading(false));
  }

  function toggleEditorLabel(name: string) {
    if (!editor) return;
    const labels = editor.labels ?? [];
    setEditor({
      ...editor,
      labels: labels.includes(name)
        ? labels.filter((label) => label !== name)
        : [...labels, name],
    });
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

  function toggleRelatedIssue(issueId: string) {
    if (!createEditor) return;
    const selected = createEditor.relatedIssueIds.includes(issueId);
    setCreateEditor({
      ...createEditor,
      relatedIssueIds: selected
        ? createEditor.relatedIssueIds.filter((id) => id !== issueId)
        : [...createEditor.relatedIssueIds, issueId],
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
            <b>{backlog.length}</b>
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
            <option value="all">Todas as abertas</option>
            <option value="unplanned">Por planear</option>
            <option value="planned">Planeadas por mim</option>
          </select>
          <select
            className="backlog-filter assignee-filter"
            value={assigneeFilter}
            onChange={(event) => setAssigneeFilter(event.target.value)}
            aria-label="Filtrar US por responsável"
          >
            <option value="all">Todos os responsáveis</option>
            <option value="unassigned">Sem responsável</option>
            {assigneeOptions.map((assignee) => (
              <option key={assignee.id} value={assignee.id}>
                {assignee.name ||
                  assignee.username ||
                  `Utilizador ${assignee.id}`}
              </option>
            ))}
          </select>
          <div className="backlog-multi-filters">
            <details className="backlog-multi-filter">
              <summary>
                <span>Projetos</span>
                <b>
                  {
                    backlogProjectOptions.filter(
                      (project) => !hiddenProjectIds.includes(project.projectId),
                    ).length
                  }
                  /
                  {backlogProjectOptions.length}
                </b>
              </summary>
              <div className="backlog-filter-popover">
                <div className="backlog-filter-actions">
                  <button
                    type="button"
                    onClick={() => setHiddenProjectIds([])}
                  >
                    Mostrar todos
                  </button>
                  <button
                    type="button"
                    onClick={() =>
                      setHiddenProjectIds(
                        backlogProjectOptions.map((issue) => issue.projectId),
                      )
                    }
                  >
                    Ocultar todos
                  </button>
                </div>
                <div className="backlog-filter-options">
                  {backlogProjectOptions.map((project) => (
                    <label key={project.projectId}>
                      <input
                        type="checkbox"
                        checked={!hiddenProjectIds.includes(project.projectId)}
                        onChange={() =>
                          setHiddenProjectIds((current) =>
                            current.includes(project.projectId)
                              ? current.filter((id) => id !== project.projectId)
                              : [...current, project.projectId],
                          )
                        }
                      />
                      <i style={{ backgroundColor: project.color }} />
                      <span>
                        <strong>{project.project}</strong>
                        <small>{project.client}</small>
                      </span>
                    </label>
                  ))}
                </div>
              </div>
            </details>
            <details className="backlog-multi-filter">
              <summary>
                <span>Labels</span>
                <b>
                  {
                    backlogLabelOptions.filter(
                      (label) => !hiddenLabels.includes(label.name),
                    ).length
                  }
                  /
                  {backlogLabelOptions.length}
                </b>
              </summary>
              <div className="backlog-filter-popover">
                <div className="backlog-filter-actions">
                  <button type="button" onClick={() => setHiddenLabels([])}>
                    Mostrar todas
                  </button>
                  <button
                    type="button"
                    onClick={() =>
                      setHiddenLabels(
                        backlogLabelOptions.map((label) => label.name),
                      )
                    }
                  >
                    Ocultar todas
                  </button>
                </div>
                <div className="backlog-filter-options">
                  {backlogLabelOptions.map((label) => (
                    <label key={label.name}>
                      <input
                        type="checkbox"
                        checked={!hiddenLabels.includes(label.name)}
                        onChange={() =>
                          setHiddenLabels((current) =>
                            current.includes(label.name)
                              ? current.filter((name) => name !== label.name)
                              : [...current, label.name],
                          )
                        }
                      />
                      <i style={{ backgroundColor: label.color }} />
                      <span>
                        <strong>{label.name}</strong>
                      </span>
                    </label>
                  ))}
                  {!backlogLabelOptions.length && <p>Sem labels disponíveis.</p>}
                </div>
              </div>
            </details>
          </div>
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
                  <IssueAssignees assignees={issue.assignees} />
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
                              className={`agenda-event ${task.fixed ? "fixed" : ""} ${closedIssueIds.has(task.rootId) ? "completed" : ""}`}
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
                          className={`timeline-task timeline-span ${bar.task.fixed ? "fixed" : ""} ${closedIssueIds.has(bar.task.rootId) ? "completed" : ""}`}
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
                          <IssueAssignees
                            assignees={bar.task.assignees}
                            compact
                          />
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
              <IssueAssignees assignees={editingIssue.assignees} />
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
              {editor.allocationId && (
                <div className="allocation-description markdown-field">
                  <span>Descrição da US no GitLab</span>
                  <MarkdownEditor
                    value={editor.issueDescription ?? ""}
                    onChange={(issueDescription) =>
                      setEditor({ ...editor, issueDescription })
                    }
                    renderMarkdown={(markdown) =>
                      renderMarkdown(markdown, editor.projectId)
                    }
                    placeholder={"## Objetivo\n\nDescreve o contexto e os critérios de aceitação…"}
                  />
                </div>
              )}
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
              {editor.allocationId && (
                <div className="allocation-description project-picker-field">
                  <span>Projeto da US no GitLab</span>
                  <div className="project-picker">
                    <div className="project-picker-selected">
                      <i style={{ backgroundColor: selectedEditorProject?.color }} />
                      <span>
                        {selectedEditorProject
                          ? `${selectedEditorProject.client} · ${selectedEditorProject.project}`
                          : "Escolhe um projeto"}
                      </span>
                      {selectedEditorProject && (
                        <small>#{selectedEditorProject.projectId}</small>
                      )}
                    </div>
                    <div className="project-picker-search">
                      <span>⌕</span>
                      <input
                        value={editorProjectSearch}
                        onChange={(event) =>
                          setEditorProjectSearch(event.target.value)
                        }
                        placeholder="Escreve para filtrar projetos…"
                        aria-label="Filtrar projeto da US"
                      />
                      {editorProjectSearch && (
                        <button
                          type="button"
                          onClick={() => setEditorProjectSearch("")}
                          aria-label="Limpar pesquisa de projetos"
                        >
                          ×
                        </button>
                      )}
                    </div>
                    {editorProjectNeedle && (
                      <div className="project-picker-results" role="listbox">
                        {matchingEditorProjects.map((project) => (
                          <button
                            type="button"
                            key={project.projectId}
                            role="option"
                            aria-selected={project.projectId === editor.projectId}
                            className={project.projectId === editor.projectId ? "selected" : ""}
                            onClick={() => {
                              setEditor({
                                ...editor,
                                projectId: project.projectId,
                                labels: [],
                              });
                              setEditorProjectSearch("");
                              refreshEditorProjectLabels(project.projectId);
                            }}
                          >
                            <span>
                              <strong>{project.project}</strong>
                              <small>{project.client}</small>
                            </span>
                            <em>#{project.projectId}</em>
                          </button>
                        ))}
                        {!matchingEditorProjects.length && (
                          <p>Nenhum projeto encontrado.</p>
                        )}
                      </div>
                    )}
                  </div>
                </div>
              )}
              {editor.allocationId &&
                editor.projectId !== undefined &&
                editor.projectId !== editingIssue.projectId && (
                  <div className="allocation-warning">
                    <b>!</b>
                    <p>
                      Ao guardar, o GitLab move a US para o projeto escolhido e
                      pode atribuir-lhe um novo número.
                    </p>
                  </div>
                )}
              {editor.allocationId && (
                <div className="allocation-description label-picker-field">
                  <span>Labels GitLab</span>
                  {editorLabelsLoading ? (
                    <div className="label-picker-state">A carregar labels do projeto…</div>
                  ) : (labelCatalog[String(editor.projectId)] ?? []).length ? (
                    <div className="label-picker" role="group" aria-label="Labels da US">
                      {(labelCatalog[String(editor.projectId)] ?? []).map((label) => {
                        const selected = (editor.labels ?? []).includes(label.name);
                        return (
                          <button
                            type="button"
                            key={label.name}
                            className={selected ? "selected" : ""}
                            onClick={() => toggleEditorLabel(label.name)}
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
              )}
              {editor.allocationId && (
                <label>
                  <span>Estado da US no GitLab</span>
                  <select
                    value={editor.issueState ?? editingIssue.state}
                    onChange={(event) =>
                      setEditor({
                        ...editor,
                        issueState: event.target.value as "opened" | "closed",
                      })
                    }
                  >
                    <option value="opened">Open</option>
                    <option value="closed">Closed</option>
                  </select>
                </label>
              )}
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
              {editorError && (
                <div className="create-task-error">{editorError}</div>
              )}
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
                disabled={!editor.description.trim() || savingEditor}
                onClick={() => void submitEditor()}
              >
                {savingEditor
                  ? "A guardar…"
                  : editor.allocationId
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
              <div className="allocation-description project-picker-field">
                <span>Projeto GitLab</span>
                <div className="project-picker">
                  <div className="project-picker-selected">
                    <i />
                    <span>
                      {selectedCreateProject
                        ? `${selectedCreateProject.client} · ${selectedCreateProject.project}`
                        : "Escolhe um projeto"}
                    </span>
                    {selectedCreateProject && (
                      <small>#{selectedCreateProject.projectId}</small>
                    )}
                  </div>
                  <div className="project-picker-search">
                    <span>⌕</span>
                    <input
                      value={projectSearch}
                      onChange={(event) => setProjectSearch(event.target.value)}
                      placeholder="Pesquisar cliente, projeto ou ID…"
                      aria-label="Pesquisar projeto GitLab"
                    />
                    {projectSearch && (
                      <button
                        type="button"
                        onClick={() => setProjectSearch("")}
                        aria-label="Limpar pesquisa de projetos"
                      >
                        ×
                      </button>
                    )}
                  </div>
                  {projectSearchNeedle && (
                    <div className="project-picker-results" role="listbox">
                      {matchingProjects.map((project) => (
                        <button
                          type="button"
                          key={project.projectId}
                          role="option"
                          aria-selected={
                            project.projectId === createEditor.projectId
                          }
                          className={
                            project.projectId === createEditor.projectId
                              ? "selected"
                              : ""
                          }
                          onClick={() => {
                            setCreateEditor({
                              ...createEditor,
                              projectId: project.projectId,
                              labels: [],
                            });
                            setProjectSearch("");
                            refreshProjectLabels(project.projectId);
                          }}
                        >
                          <span>
                            <strong>{project.project}</strong>
                            <small>{project.client}</small>
                          </span>
                          <em>#{project.projectId}</em>
                        </button>
                      ))}
                      {!matchingProjects.length && (
                        <p>Nenhum projeto encontrado.</p>
                      )}
                    </div>
                  )}
                </div>
              </div>
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
                <MarkdownEditor
                  value={createEditor.description}
                  onChange={(description) =>
                    setCreateEditor({
                      ...createEditor,
                      description,
                    })
                  }
                  renderMarkdown={(markdown) =>
                    renderMarkdown(markdown, createEditor.projectId)
                  }
                  placeholder={"## Objetivo\n\nDescreve o trabalho a realizar. Podes usar listas, tabelas, links, checklists e outros elementos Markdown."}
                />
              </label>
              <div className="allocation-description related-issues-field">
                <div className="related-issues-title">
                  <span>US possivelmente relacionadas</span>
                  <small>
                    Sugestões calculadas pelo tema do título e da descrição
                  </small>
                </div>
                {selectedRelatedIssues.length > 0 && (
                  <div className="related-issues-selected">
                    {selectedRelatedIssues.map((issue) => (
                      <button
                        type="button"
                        key={issue.id}
                        onClick={() => toggleRelatedIssue(issue.id)}
                        title="Remover relação"
                      >
                        #{issue.iid} · {issue.title} <b>×</b>
                      </button>
                    ))}
                  </div>
                )}
                {relatedSuggestions.length ? (
                  <div className="related-issue-suggestions">
                    {relatedSuggestions.map(({ issue, matches }) => {
                      const selected = createEditor.relatedIssueIds.includes(
                        issue.id,
                      );
                      return (
                        <button
                          type="button"
                          key={issue.id}
                          className={selected ? "selected" : ""}
                          onClick={() => toggleRelatedIssue(issue.id)}
                          aria-pressed={selected}
                        >
                          <span>
                            <small>
                              {issue.client} · {issue.project} · #{issue.iid}
                            </small>
                            <strong>{issue.title}</strong>
                            {matches.length > 0 && (
                              <em>Em comum: {matches.join(", ")}</em>
                            )}
                          </span>
                          <b>{selected ? "✓ Relacionar" : "+ Relacionar"}</b>
                        </button>
                      );
                    })}
                  </div>
                ) : createEditor.title.trim().length > 2 ||
                  createEditor.description.trim().length > 10 ? (
                  <div className="related-issues-empty">
                    Ainda não foram encontradas US com um tema semelhante.
                  </div>
                ) : (
                  <div className="related-issues-empty">
                    Começa a escrever para receber sugestões.
                  </div>
                )}
              </div>
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

type FilterOption = { value: string; label: string; meta?: string };

function MultiSelectFilter({
  label,
  options,
  values,
  onChange,
  placeholder = "Pesquisar…",
}: {
  label: string;
  options: FilterOption[];
  values: string[];
  onChange: (values: string[]) => void;
  placeholder?: string;
}) {
  const [search, setSearch] = useState("");
  const detailsRef = useRef<HTMLDetailsElement>(null);
  const needle = search.trim().toLowerCase();
  const visible = options.filter((option) =>
    `${option.label} ${option.meta ?? ""}`.toLowerCase().includes(needle),
  );
  const selectedLabels = options
    .filter((option) => values.includes(option.value))
    .map((option) => option.label);

  function toggle(value: string) {
    onChange(
      values.includes(value)
        ? values.filter((item) => item !== value)
        : [...values, value],
    );
  }

  useEffect(() => {
    function closeOnOutsideClick(event: PointerEvent) {
      if (!detailsRef.current?.contains(event.target as Node)) {
        detailsRef.current?.removeAttribute("open");
      }
    }
    function closeOnEscape(event: KeyboardEvent) {
      if (event.key === "Escape") {
        detailsRef.current?.removeAttribute("open");
      }
    }
    document.addEventListener("pointerdown", closeOnOutsideClick);
    document.addEventListener("keydown", closeOnEscape);
    return () => {
      document.removeEventListener("pointerdown", closeOnOutsideClick);
      document.removeEventListener("keydown", closeOnEscape);
    };
  }, []);

  return (
    <label className="multi-filter-label">
      <span>{label}</span>
      <details className="multi-filter" ref={detailsRef}>
        <summary title={selectedLabels.join(", ")}>
          <span>
            {values.length === 0
              ? `Todos os ${label.toLowerCase()}`
              : values.length === 1
                ? selectedLabels[0]
                : `${values.length} selecionados`}
          </span>
          <b>⌄</b>
        </summary>
        <div className="multi-filter-popover">
          <button
            type="button"
            className="multi-filter-close"
            aria-label={`Fechar filtro ${label}`}
            onClick={() => detailsRef.current?.removeAttribute("open")}
          >
            ×
          </button>
          <div className="multi-filter-search">
            <span>⌕</span>
            <input
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder={placeholder}
              onClick={(event) => event.stopPropagation()}
            />
          </div>
          <div className="multi-filter-actions">
            <button type="button" onClick={() => onChange(visible.map((item) => item.value))}>
              Selecionar visíveis
            </button>
            <button type="button" onClick={() => onChange([])}>Limpar</button>
          </div>
          <div className="multi-filter-options">
            {visible.map((option) => (
              <label key={option.value}>
                <input
                  type="checkbox"
                  checked={values.includes(option.value)}
                  onChange={() => toggle(option.value)}
                />
                <span>{option.label}</span>
                {option.meta && <small>{option.meta}</small>}
              </label>
            ))}
            {!visible.length && <p>Sem resultados.</p>}
          </div>
        </div>
      </details>
    </label>
  );
}

function IssuesView({ issues, tasks }: { issues: IssueRecord[]; tasks: Task[] }) {
  const [search, setSearch] = useState("");
  const [projectIds, setProjectIds] = useState<string[]>([]);
  const [state, setState] = useState<"all" | "opened" | "closed">("opened");
  const [selectedLabels, setSelectedLabels] = useState<string[]>([]);
  const [assigneeIds, setAssigneeIds] = useState<string[]>([]);
  const [allocation, setAllocation] = useState<
    "all" | "unplanned" | "partial" | "allocated"
  >("all");
  const plannedByIssue = useMemo(() => {
    const totals = new Map<string, number>();
    tasks.forEach((task) =>
      totals.set(task.rootId, (totals.get(task.rootId) ?? 0) + task.estimate),
    );
    return totals;
  }, [tasks]);
  const projects = useMemo(
    () =>
      uniqueBy(issues, (issue) => String(issue.projectId)).sort(
        (left, right) =>
          left.client.localeCompare(right.client) ||
          left.project.localeCompare(right.project),
      ),
    [issues],
  );
  const labels = useMemo(
    () =>
      [...new Set(issues.flatMap((issue) => issue.labels.map((item) => item.name)))].sort(
        (left, right) => left.localeCompare(right),
      ),
    [issues],
  );
  const assignees = useMemo(() => {
    const values = new Map<number, GitLabAssignee>();
    issues.forEach((issue) =>
      issue.assignees.forEach((item) => values.set(item.id, item)),
    );
    return [...values.values()].sort((left, right) =>
      (left.name || left.username || "").localeCompare(
        right.name || right.username || "",
      ),
    );
  }, [issues]);
  const needle = search.trim().toLowerCase();
  const filtered = issues.filter((issue) => {
    const planned = plannedByIssue.get(issue.id) ?? 0;
    const remaining = Math.max(0, issue.estimateTotal - planned);
    if (state !== "all" && issue.state !== state) return false;
    if (projectIds.length && !projectIds.includes(String(issue.projectId)))
      return false;
    if (
      selectedLabels.length &&
      !issue.labels.some((item) => selectedLabels.includes(item.name))
    ) return false;
    if (assigneeIds.length && !(
      (assigneeIds.includes("unassigned") && issue.assignees.length === 0) ||
      issue.assignees.some((item) => assigneeIds.includes(String(item.id)))
    ))
      return false;
    if (allocation === "unplanned" && planned > 0) return false;
    if (allocation === "partial" && !(planned > 0 && remaining > 0.001))
      return false;
    if (allocation === "allocated" && remaining > 0.001) return false;
    return (
      !needle ||
      `${issue.iid} ${issue.title} ${issue.client} ${issue.project} ${issue.labels.map((item) => item.name).join(" ")} ${issue.assignees.map((item) => `${item.name ?? ""} ${item.username ?? ""}`).join(" ")}`
        .toLowerCase()
        .includes(needle)
    );
  });
  const teamHours = filtered.reduce(
    (sum, issue) => sum + issue.estimateTotal,
    0,
  );
  const plannedHours = filtered.reduce(
    (sum, issue) => sum + (plannedByIssue.get(issue.id) ?? 0),
    0,
  );
  const hoursToAllocate = filtered.reduce(
    (sum, issue) =>
      sum + Math.max(0, issue.estimateTotal - (plannedByIssue.get(issue.id) ?? 0)),
    0,
  );

  function clearFilters() {
    setSearch("");
    setProjectIds([]);
    setState("opened");
    setSelectedLabels([]);
    setAssigneeIds([]);
    setAllocation("all");
  }

  return (
    <div className="page wide-page issues-view">
      <PageIntro
        eyebrow="PESQUISA E CAPACIDADE"
        title="Explorar US"
        description="Combina filtros para localizar trabalho e perceber rapidamente quantas horas ainda precisam de alocação."
      >
        <button className="secondary-button" onClick={clearFilters}>
          Limpar filtros
        </button>
      </PageIntro>
      <section className="metrics-grid issue-metrics">
        <article><small>US ENCONTRADAS</small><strong>{filtered.length}</strong><p>de {issues.length} sincronizadas</p></article>
        <article><small>ESTIMATE DA EQUIPA</small><strong>{formatHours(teamHours)}</strong><p>nas US filtradas</p></article>
        <article><small>JÁ ALOCADO</small><strong>{formatHours(plannedHours)}</strong><p>no planeamento pessoal</p></article>
        <article><small>POR ALOCAR</small><strong>{formatHours(hoursToAllocate)}</strong><p>estimate ainda sem cobertura</p></article>
      </section>
      <section className="panel issue-explorer">
        <div className="issue-explorer-filters">
          <label className="issue-search-field">
            <span>Pesquisar</span>
            <input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Título, número, projeto, label ou pessoa…" />
          </label>
          <MultiSelectFilter label="Projetos" values={projectIds} onChange={setProjectIds} placeholder="Pesquisar projeto…" options={projects.map((project) => ({ value: String(project.projectId), label: project.project, meta: project.client }))} />
          <label><span>Estado</span><select value={state} onChange={(event) => setState(event.target.value as typeof state)}><option value="all">Todos</option><option value="opened">Open</option><option value="closed">Closed</option></select></label>
          <MultiSelectFilter label="Labels" values={selectedLabels} onChange={setSelectedLabels} placeholder="Pesquisar label…" options={labels.map((name) => ({ value: name, label: name }))} />
          <MultiSelectFilter label="Responsáveis" values={assigneeIds} onChange={setAssigneeIds} placeholder="Pesquisar responsável…" options={[{ value: "unassigned", label: "Sem responsável" }, ...assignees.map((item) => ({ value: String(item.id), label: item.name || item.username || `Utilizador ${item.id}`, meta: item.username }))]} />
          <label><span>Alocação</span><select value={allocation} onChange={(event) => setAllocation(event.target.value as typeof allocation)}><option value="all">Qualquer situação</option><option value="unplanned">Sem alocação</option><option value="partial">Parcialmente alocada</option><option value="allocated">Totalmente alocada</option></select></label>
        </div>
        <div className="issue-results-head"><strong>{filtered.length} US</strong><span>Os filtros são combinados entre si</span></div>
        <div className="issue-results">
          {filtered.map((issue) => {
            const planned = plannedByIssue.get(issue.id) ?? 0;
            const remaining = Math.max(0, issue.estimateTotal - planned);
            return (
              <article key={issue.id} className="issue-result-row">
                <div className="issue-result-main">
                  <small>{issue.client} · {issue.project} · #{issue.iid}</small>
                  {issue.webUrl ? <a href={issue.webUrl} target="_blank" rel="noreferrer">{issue.title} <span>↗</span></a> : <strong>{issue.title}</strong>}
                  <IssueLabels labels={issue.labels} />
                  <IssueAssignees assignees={issue.assignees} />
                </div>
                <div className="issue-result-hours"><span><small>Estimate</small><b>{issue.hasEstimate ? formatHours(issue.estimateTotal) : "—"}</b></span><span><small>Alocado</small><b>{formatHours(planned)}</b></span><span className={remaining > 0 ? "pending" : "done"}><small>Por alocar</small><b>{formatHours(remaining)}</b></span></div>
              </article>
            );
          })}
          {!filtered.length && <div className="empty-state"><span>⌕</span><strong>Nenhuma US corresponde aos filtros</strong><p>Altera ou limpa os filtros para alargar a pesquisa.</p></div>}
        </div>
      </section>
    </div>
  );
}

const TEAM_WEEKDAYS = ["2ª Feira", "3ª Feira", "4ª Feira", "5ª Feira", "6ª Feira"];
const TEAM_ABSENCE_TYPES: Record<TeamAbsenceType, { label: string; color: string }> = {
  vacation: { label: "Férias", color: "#d88a24" },
  absence: { label: "Ausência", color: "#d05260" },
  training: { label: "Formação", color: "#278b78" },
  other: { label: "Outro", color: "#778092" },
};

function teamWeekLabel(weekStart: string) {
  const end = addDays(weekStart, 4);
  return `${new Intl.DateTimeFormat("pt-PT", { day: "2-digit", month: "short" }).format(fromKey(weekStart))} — ${new Intl.DateTimeFormat("pt-PT", { day: "2-digit", month: "short", year: "numeric" }).format(fromKey(end))}`;
}

function escapeHtml(value: string) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function TeamAllocationView({
  issues,
  members,
  allocations,
  setMembers,
  setAllocations,
  todayKey,
  setToast,
}: {
  issues: IssueRecord[];
  members: TeamMember[];
  allocations: TeamAllocation[];
  setMembers: Dispatch<SetStateAction<TeamMember[]>>;
  setAllocations: Dispatch<SetStateAction<TeamAllocation[]>>;
  todayKey: string;
  setToast: (message: string) => void;
}) {
  const [weekStart, setWeekStart] = useState(() => mondayFor(todayKey));
  const [search, setSearch] = useState("");
  const [projectIds, setProjectIds] = useState<string[]>([]);
  const [selectedLabels, setSelectedLabels] = useState<string[]>([]);
  const [assigneeIds, setAssigneeIds] = useState<string[]>([]);
  const [state, setState] = useState<"all" | "opened" | "closed">("opened");
  const [allocationFilter, setAllocationFilter] = useState<
    "all" | "unplanned" | "partial" | "allocated"
  >("all");
  const [draftHours, setDraftHours] = useState<Record<string, number>>({});
  const [draggedIssueId, setDraggedIssueId] = useState<string | null>(null);
  const [draggedAllocationId, setDraggedAllocationId] = useState<string | null>(null);
  const [dropPreview, setDropPreview] = useState<{
    memberId: string;
    startSlot: number;
    hours: number;
    valid: boolean;
  } | null>(null);
  const [selectedAllocationId, setSelectedAllocationId] = useState<string | null>(null);
  const [newMemberName, setNewMemberName] = useState("");
  const [customTitle, setCustomTitle] = useState("");
  const [absenceType, setAbsenceType] = useState<TeamAbsenceType>("vacation");
  const [absenceMemberId, setAbsenceMemberId] = useState(() => members[0]?.id ?? "");
  const [absenceDay, setAbsenceDay] = useState(0);
  const [absenceStartHour, setAbsenceStartHour] = useState(0);
  const [customHours, setCustomHours] = useState(8);

  useEffect(() => {
    if (!members.some((member) => member.id === absenceMemberId)) {
      setAbsenceMemberId(members[0]?.id ?? "");
    }
  }, [absenceMemberId, members]);

  const issueById = useMemo(
    () => new Map(issues.map((issue) => [issue.id, issue])),
    [issues],
  );
  const currentAllocations = allocations.filter(
    (allocation) => allocation.weekStart === weekStart,
  );
  const allocatedByIssue = useMemo(() => {
    const totals = new Map<string, number>();
    allocations
      .filter((allocation) => allocation.weekStart === weekStart && allocation.issueId)
      .forEach((allocation) =>
        totals.set(
          allocation.issueId!,
          (totals.get(allocation.issueId!) ?? 0) + allocation.hours,
        ),
      );
    return totals;
  }, [allocations, weekStart]);
  const projects = useMemo(
    () =>
      uniqueBy(issues, (issue) => String(issue.projectId)).sort(
        (left, right) =>
          left.client.localeCompare(right.client) ||
          left.project.localeCompare(right.project),
      ),
    [issues],
  );
  const labels = useMemo(
    () =>
      [...new Set(issues.flatMap((issue) => issue.labels.map((item) => item.name)))].sort(
        (left, right) => left.localeCompare(right),
      ),
    [issues],
  );
  const assignees = useMemo(() => {
    const values = new Map<number, GitLabAssignee>();
    issues.forEach((issue) =>
      issue.assignees.forEach((item) => values.set(item.id, item)),
    );
    return [...values.values()].sort((left, right) =>
      (left.name || left.username || "").localeCompare(
        right.name || right.username || "",
      ),
    );
  }, [issues]);
  const needle = search.trim().toLowerCase();
  const filteredIssues = issues.filter((issue) => {
    const allocated = allocatedByIssue.get(issue.id) ?? 0;
    const estimate = Math.max(0, issue.estimateTotal);
    const remaining = Math.max(0, estimate - allocated);
    if (state !== "all" && issue.state !== state) return false;
    if (projectIds.length && !projectIds.includes(String(issue.projectId))) return false;
    if (
      selectedLabels.length &&
      !issue.labels.some((item) => selectedLabels.includes(item.name))
    ) return false;
    if (assigneeIds.length && !(
      (assigneeIds.includes("unassigned") && issue.assignees.length === 0) ||
      issue.assignees.some((item) => assigneeIds.includes(String(item.id)))
    )) return false;
    if (allocationFilter === "unplanned" && allocated > 0) return false;
    if (allocationFilter === "partial" && !(allocated > 0 && remaining > 0.001)) return false;
    if (allocationFilter === "allocated" && (estimate <= 0 || remaining > 0.001)) return false;
    return !needle || `${issue.iid} ${issue.title} ${issue.client} ${issue.project} ${issue.labels.map((item) => item.name).join(" ")} ${issue.assignees.map((item) => `${item.name ?? ""} ${item.username ?? ""}`).join(" ")}`.toLowerCase().includes(needle);
  });
  const selectedAllocation = allocations.find(
    (allocation) => allocation.id === selectedAllocationId,
  );

  function defaultHours(issue: IssueRecord) {
    return Math.min(40, Math.max(1, Math.round(issue.estimateTotal || 1)));
  }

  function hasCollision(
    memberId: string,
    startSlot: number,
    hours: number,
    ignoreId?: string,
  ) {
    const end = startSlot + hours;
    return currentAllocations.some(
      (allocation) =>
        allocation.memberId === memberId &&
        allocation.id !== ignoreId &&
        startSlot < allocation.startSlot + allocation.hours &&
        end > allocation.startSlot,
    );
  }

  function normalizePlacement(startSlot: number, requestedHours: number) {
    const safeStart = Math.min(39, Math.max(0, Math.round(startSlot)));
    const hours = Math.min(40, Math.max(1, Math.round(requestedHours)));
    return {
      startSlot: Math.min(safeStart, 40 - hours),
      hours,
    };
  }

  function findAvailablePlacement(
    memberId: string,
    preferredStart: number,
    requestedHours: number,
    ignoreId?: string,
  ) {
    const preferred = normalizePlacement(preferredStart, requestedHours);
    if (!hasCollision(memberId, preferred.startSlot, preferred.hours, ignoreId)) {
      return preferred;
    }
    const candidates = Array.from(
      { length: 41 - preferred.hours },
      (_, startSlot) => startSlot,
    ).sort((left, right) => {
      const distance = Math.abs(left - preferred.startSlot) - Math.abs(right - preferred.startSlot);
      return distance || left - right;
    });
    const startSlot = candidates.find(
      (candidate) => !hasCollision(memberId, candidate, preferred.hours, ignoreId),
    );
    return startSlot === undefined
      ? null
      : { startSlot, hours: preferred.hours };
  }

  function previewDrop(memberId: string, hoveredSlot: number) {
    const draggedAllocation = draggedAllocationId
      ? allocations.find((allocation) => allocation.id === draggedAllocationId)
      : null;
    const draggedIssue = draggedIssueId ? issueById.get(draggedIssueId) : null;
    const requestedHours = draggedAllocation
      ? draggedAllocation.hours
      : draggedIssue
        ? draftHours[draggedIssue.id] ?? defaultHours(draggedIssue)
        : 1;
    const placement = findAvailablePlacement(
      memberId,
      hoveredSlot,
      requestedHours,
      draggedAllocation?.id,
    );
    const next = placement
      ? { memberId, ...placement, valid: true }
      : { memberId, startSlot: hoveredSlot, hours: 1, valid: false };
    setDropPreview((current) =>
      current &&
      current.memberId === next.memberId &&
      current.startSlot === next.startSlot &&
      current.hours === next.hours &&
      current.valid === next.valid
        ? current
        : next,
    );
  }

  function addAllocation(memberId: string, startSlot: number, issueId: string) {
    const issue = issueById.get(issueId);
    if (!issue) return;
    const placement = findAvailablePlacement(
      memberId,
      startSlot,
      draftHours[issueId] ?? defaultHours(issue),
    );
    if (!placement) {
      setToast("Essa pessoa não tem espaço livre suficiente para esta US na semana.");
      return;
    }
    const allocation: TeamAllocation = {
      id: crypto.randomUUID(),
      memberId,
      issueId,
      weekStart,
      ...placement,
    };
    setAllocations((current) => [...current, allocation]);
    setSelectedAllocationId(allocation.id);
  }

  function addCustomAllocation() {
    const member = members.find((item) => item.id === absenceMemberId);
    const typeConfig = TEAM_ABSENCE_TYPES[absenceType];
    const title = customTitle.trim() || typeConfig.label;
    if (!member) {
      setToast("Escolhe uma pessoa para registar a ausência.");
      return;
    }
    const placement = findAvailablePlacement(
      member.id,
      absenceDay * 8 + absenceStartHour,
      customHours,
    );
    if (!placement) {
      setToast("Essa pessoa não tem espaço livre suficiente para esta ausência na semana.");
      return;
    }
    const allocation: TeamAllocation = {
      id: crypto.randomUUID(),
      memberId: member.id,
      customTitle: title,
      customColor: typeConfig.color,
      absenceType,
      weekStart,
      ...placement,
    };
    setAllocations((current) => [...current, allocation]);
    setSelectedAllocationId(allocation.id);
    setCustomTitle("");
    setToast(`${typeConfig.label} registada para ${member.name}.`);
  }

  function updateAllocation(
    allocationId: string,
    patch: Partial<TeamAllocation>,
  ) {
    setAllocations((current) =>
      current.map((allocation) => {
        if (allocation.id !== allocationId) return allocation;
        const memberId = patch.memberId ?? allocation.memberId;
        const placement = normalizePlacement(
          patch.startSlot ?? allocation.startSlot,
          patch.hours ?? allocation.hours,
        );
        if (hasCollision(
          memberId,
          placement.startSlot,
          placement.hours,
          allocationId,
        )) {
          setToast("A alteração sobrepõe outra alocação.");
          return allocation;
        }
        return { ...allocation, ...patch, memberId, ...placement };
      }),
    );
  }

  function moveAllocation(allocationId: string, memberId: string, startSlot: number) {
    const allocation = allocations.find((item) => item.id === allocationId);
    if (!allocation) return;
    const placement = findAvailablePlacement(
      memberId,
      startSlot,
      allocation.hours,
      allocationId,
    );
    if (!placement) {
      setDraggedAllocationId(null);
      setToast("Essa pessoa não tem espaço livre suficiente para mover esta alocação.");
      return;
    }
    setAllocations((current) =>
      current.map((item) =>
        item.id === allocationId
          ? { ...item, memberId, ...placement }
          : item,
      ),
    );
    setSelectedAllocationId(allocationId);
  }

  function removeAllocation(allocationId: string) {
    setAllocations((current) => current.filter((item) => item.id !== allocationId));
    setSelectedAllocationId((current) => current === allocationId ? null : current);
    setToast("Alocação removida.");
  }

  function addMember() {
    const name = newMemberName.trim();
    if (!name) return;
    setMembers((current) => [
      ...current,
      { id: crypto.randomUUID(), name },
    ]);
    setNewMemberName("");
  }

  function removeMember(member: TeamMember) {
    if (!window.confirm(`Remover ${member.name} e as respetivas alocações?`)) return;
    setMembers((current) => current.filter((item) => item.id !== member.id));
    setAllocations((current) =>
      current.filter((allocation) => allocation.memberId !== member.id),
    );
    setSelectedAllocationId(null);
  }

  function allocationText(allocation: TeamAllocation) {
    const issue = allocation.issueId ? issueById.get(allocation.issueId) : null;
    return issue
      ? `${issue.project} · #${issue.iid} ${issue.title}`
      : allocation.customTitle ?? "Bloco livre";
  }

  function buildEmailTable() {
    const border = "#dfe2e8";
    const percentRow = Array.from({ length: 40 }, () => `<th style="width:30px;height:24px;padding:0;text-align:center;color:#8d94a2;border-right:1px solid #eceef2;border-bottom:1px solid ${border};font-size:9px;font-weight:400">2.5%</th>`).join("");
    const hourRow = Array.from({ length: 40 }, (_, index) => `<th style="height:28px;padding:0;text-align:center;color:#5b6371;border-right:1px solid #eceef2;border-bottom:1px solid ${border};font-size:10px">${(index % 8) + 1}</th>`).join("");
    const dayRow = TEAM_WEEKDAYS.map((day, index) => `<th colspan="8" style="height:38px;padding:0 9px;text-align:left;color:#4c5463;border-right:1px solid #cdd1da;border-bottom:1px solid ${border};background:#f5f6f9;font-size:11px">${day}<span style="float:right;color:#9299a7;font-size:9px;font-weight:400">${escapeHtml(formatDate(addDays(weekStart, index)))}</span></th>`).join("");
    const corner = (content: string) => `<th style="width:210px;padding:0 12px;text-align:left;color:#4d5564;border-right:1px solid ${border};border-bottom:1px solid ${border};background:#fafbfc;font-size:10px">${content}</th>`;
    const columns = `<colgroup><col style="width:210px">${Array.from({ length: 40 }, () => '<col style="width:30px">').join("")}</colgroup>`;
    const body = members.map((member) => {
      const memberAllocations = currentAllocations.filter(
        (allocation) => allocation.memberId === member.id,
      );
      const total = memberAllocations.reduce((sum, allocation) => sum + allocation.hours, 0);
      const byStart = new Map(
        memberAllocations
          .sort((left, right) => left.startSlot - right.startSlot)
          .map((allocation) => [allocation.startSlot, allocation]),
      );
      const cells: string[] = [];
      for (let slot = 0; slot < 40;) {
        const allocation = byStart.get(slot);
        if (!allocation) {
          cells.push(`<td style="height:62px;padding:0;border-right:1px solid #eef0f3;border-bottom:1px solid ${border};background:#fff"></td>`);
          slot += 1;
          continue;
        }
        const issue = allocation.issueId ? issueById.get(allocation.issueId) : null;
        const project = issue?.project ?? allocation.customTitle ?? "Outro";
        const issueLine = issue
          ? issue.webUrl
            ? `<br><a href="${escapeHtml(issue.webUrl)}" style="color:#596171;text-decoration:none;font-size:10px">#${issue.iid} ${escapeHtml(issue.title)}</a>`
            : `<br><span style="color:#596171;font-size:10px">#${issue.iid} ${escapeHtml(issue.title)}</span>`
          : `<br><span style="color:#747b87;font-size:9px;font-style:italic">${TEAM_ABSENCE_TYPES[allocation.absenceType ?? "other"].label} · sem cliente/US</span>`;
        const color = issue?.color ?? allocation.customColor ?? "#8b91a7";
        const background = issue ? `${color}20` : allocation.absenceType === "vacation" ? "#fff0d5" : allocation.absenceType === "absence" ? "#fbe1e4" : allocation.absenceType === "training" ? "#ddf3ed" : "#eceef2";
        cells.push(`<td colspan="${allocation.hours}" style="height:50px;padding:6px 8px;vertical-align:middle;color:#3e4655;border-right:1px solid ${border};border-bottom:1px solid ${border};border-left:4px ${issue ? "solid" : "dashed"} ${color};background:${background};font-family:Arial,sans-serif"><strong style="display:block;font-size:11px">${escapeHtml(project)}</strong>${issueLine}<span style="float:right;color:#6a65dd;font-size:9px">${formatHours(allocation.hours)}</span></td>`);
        slot += allocation.hours;
      }
      return `<tr><th style="height:62px;padding:8px 10px;text-align:left;vertical-align:middle;white-space:nowrap;color:#343c4b;border-right:1px solid ${border};border-bottom:1px solid ${border};background:#fff;font-size:11px">${escapeHtml(member.name)}${member.role ? `<br><span style="color:#9198a5;font-size:9px;font-weight:400">${escapeHtml(member.role)}</span>` : ""}<br><span style="color:#6560dc;font-size:9px;font-weight:400">${formatHours(total)} · ${Math.round((total / 40) * 100)}%</span></th>${cells.join("")}</tr>`;
    }).join("");
    return `<div style="font-family:Arial,sans-serif;color:#343c4b"><div style="margin:0 0 10px"><strong style="font-size:18px">Alocação semanal da equipa</strong><br><span style="color:#7c8492;font-size:11px">${escapeHtml(teamWeekLabel(weekStart))} · 1 célula = 1 h = 2,5%</span></div><table role="presentation" cellspacing="0" cellpadding="0" style="width:1410px;table-layout:fixed;border-collapse:collapse;border:1px solid ${border};font-family:Arial,sans-serif;background:#fff">${columns}<thead><tr>${corner("Pessoa")}${percentRow}</tr><tr>${corner("Semana")}${dayRow}</tr><tr>${corner("Hora")}${hourRow}</tr></thead><tbody>${body}</tbody></table></div>`;
  }

  async function copyForEmail() {
    const html = buildEmailTable();
    const plainRows = members.map((member) => {
      const slots = Array.from({ length: 40 }, () => "");
      currentAllocations
        .filter((allocation) => allocation.memberId === member.id)
        .forEach((allocation) => {
          slots[allocation.startSlot] = allocationText(allocation);
        });
      return [member.name, ...slots].join("\t");
    });
    const plain = [
      ["", ...Array.from({ length: 40 }, () => "2.5%")].join("\t"),
      ["", ...Array.from({ length: 40 }, (_, index) => String((index % 8) + 1))].join("\t"),
      ...plainRows,
    ].join("\n");
    try {
      if (typeof ClipboardItem !== "undefined" && navigator.clipboard.write) {
        await navigator.clipboard.write([
          new ClipboardItem({
            "text/html": new Blob([html], { type: "text/html" }),
            "text/plain": new Blob([plain], { type: "text/plain" }),
          }),
        ]);
      } else {
        await navigator.clipboard.writeText(plain);
      }
      setToast("Tabela copiada. Já podes colá-la no corpo do e-mail.");
    } catch {
      setToast("O browser bloqueou a cópia. Autoriza o acesso à área de transferência.");
    }
  }

  async function copyForExcel() {
    const dayCells = TEAM_WEEKDAYS.flatMap((day) =>
      Array.from({ length: 8 }, () => day),
    );
    const rows = members.map((member) => {
      const slots = Array.from({ length: 40 }, () => "");
      currentAllocations
        .filter((allocation) => allocation.memberId === member.id)
        .forEach((allocation) => {
          const text = allocationText(allocation);
          for (
            let slot = allocation.startSlot;
            slot < allocation.startSlot + allocation.hours;
            slot++
          ) {
            slots[slot] = text;
          }
        });
      return [member.name, ...slots];
    });
    const tsv = [
      [teamWeekLabel(weekStart), ...Array.from({ length: 40 }, () => "2.5%")],
      ["Dia", ...dayCells],
      ["Hora", ...Array.from({ length: 40 }, (_, index) => String((index % 8) + 1))],
      ...rows,
    ].map((row) => row.join("\t")).join("\n");
    const html = buildEmailTable();
    try {
      if (typeof ClipboardItem !== "undefined" && navigator.clipboard.write) {
        await navigator.clipboard.write([
          new ClipboardItem({
            "text/html": new Blob([html], { type: "text/html" }),
            "text/plain": new Blob([tsv], { type: "text/plain" }),
          }),
        ]);
      } else {
        await navigator.clipboard.writeText(tsv);
      }
      setToast("Tabela copiada para Excel. Seleciona a célula A1 e cola.");
    } catch {
      setToast("O browser bloqueou a cópia. Autoriza o acesso à área de transferência.");
    }
  }

  function exportTeamCsv() {
    const header = [
      "Pessoa",
      ...TEAM_WEEKDAYS.flatMap((day) =>
        Array.from({ length: 8 }, (_, index) => `${day} ${index + 1}`),
      ),
    ];
    const rows = members.map((member) => {
      const slots = Array.from({ length: 40 }, () => "");
      currentAllocations
        .filter((allocation) => allocation.memberId === member.id)
        .forEach((allocation) => {
          const text = allocationText(allocation);
          for (let slot = allocation.startSlot; slot < allocation.startSlot + allocation.hours; slot++) {
            slots[slot] = text;
          }
        });
      return [member.name, ...slots];
    });
    const csv = [header, ...rows]
      .map((row) => row.map((value) => `"${String(value).replaceAll('"', '""')}"`).join(";"))
      .join("\n");
    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([`\uFEFF${csv}`], { type: "text/csv;charset=utf-8" }));
    link.download = `alocacao-equipa-${weekStart}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  return (
    <div className="page team-allocation-view">
      <PageIntro
        eyebrow="CAPACIDADE DA EQUIPA"
        title="Alocação semanal"
        description="Arrasta US para a grelha de 40 horas. Cada célula representa 1 hora e 2,5% da semana."
      >
        <div className="team-week-nav">
          <button onClick={() => setWeekStart(addDays(weekStart, -7))}>←</button>
          <strong>{teamWeekLabel(weekStart)}</strong>
          <button onClick={() => setWeekStart(addDays(weekStart, 7))}>→</button>
          <button className="secondary-button" onClick={() => setWeekStart(mondayFor(todayKey))}>Esta semana</button>
        </div>
      </PageIntro>

      <section className="team-allocation-layout">
        <aside className="panel team-issue-palette">
          <div className="team-palette-head">
            <div><small>BACKLOG</small><strong>{filteredIssues.length} US disponíveis</strong></div>
          </div>
          <div className="team-issue-filters">
            <label className="issue-search-field"><span>Pesquisar</span><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Título, número ou projeto…" /></label>
            <MultiSelectFilter label="Projetos" values={projectIds} onChange={setProjectIds} placeholder="Pesquisar projeto…" options={projects.map((project) => ({ value: String(project.projectId), label: project.project, meta: project.client }))} />
            <MultiSelectFilter label="Labels" values={selectedLabels} onChange={setSelectedLabels} placeholder="Pesquisar label…" options={labels.map((name) => ({ value: name, label: name }))} />
            <MultiSelectFilter label="Responsáveis" values={assigneeIds} onChange={setAssigneeIds} placeholder="Pesquisar responsável…" options={[{ value: "unassigned", label: "Sem responsável" }, ...assignees.map((item) => ({ value: String(item.id), label: item.name || item.username || `Utilizador ${item.id}`, meta: item.username }))]} />
            <label><span>Estado</span><select value={state} onChange={(event) => setState(event.target.value as typeof state)}><option value="all">Todos</option><option value="opened">Open</option><option value="closed">Closed</option></select></label>
            <label><span>Alocação semanal</span><select value={allocationFilter} onChange={(event) => setAllocationFilter(event.target.value as typeof allocationFilter)}><option value="all">Qualquer situação</option><option value="unplanned">Sem alocação</option><option value="partial">Parcial</option><option value="allocated">Total</option></select></label>
          </div>
          <div className="team-issue-list">
            {filteredIssues.map((issue) => {
              const hours = draftHours[issue.id] ?? defaultHours(issue);
              const allocated = allocatedByIssue.get(issue.id) ?? 0;
              const allocationState = allocated <= 0
                ? "unallocated"
                : issue.estimateTotal > 0 && allocated + 0.001 >= issue.estimateTotal
                  ? "fully-allocated"
                  : "partially-allocated";
              return (
                <article
                  key={issue.id}
                  className={`team-issue-card ${allocationState} ${draggedIssueId === issue.id ? "dragging" : ""}`}
                  draggable
                  onDragStart={(event) => {
                    event.dataTransfer.setData("application/x-work-organizer-issue", issue.id);
                    event.dataTransfer.effectAllowed = "copy";
                    setDraggedIssueId(issue.id);
                  }}
                  onDragEnd={() => { setDraggedIssueId(null); setDropPreview(null); }}
                >
                  <span className="team-issue-color" style={{ background: issue.color }} />
                  <div><small>{issue.client} · {issue.project} · #{issue.iid}</small><strong>{issue.title}</strong><IssueLabels labels={issue.labels} /></div>
                  <label onClick={(event) => event.stopPropagation()}><span>Horas</span><input type="number" min="1" max="40" step="1" value={hours} onChange={(event) => setDraftHours((current) => ({ ...current, [issue.id]: Math.min(40, Math.max(1, Number(event.target.value) || 1)) }))} /></label>
                  <small className="team-allocated-note"><b>{allocationState === "fully-allocated" ? "✓ Totalmente alocada" : allocationState === "partially-allocated" ? "◐ Parcialmente alocada" : "○ Ainda não alocada"}</b><span>{formatHours(allocated)} nesta semana{issue.estimateTotal > 0 ? ` / ${formatHours(issue.estimateTotal)}` : ""}</span></small>
                </article>
              );
            })}
            {!filteredIssues.length && <div className="empty-state compact"><span>⌕</span><strong>Sem US para estes filtros</strong></div>}
          </div>
          <div className="custom-allocation-form">
            <strong>Férias e ausências</strong>
            <label><span>Tipo</span><select value={absenceType} onChange={(event) => setAbsenceType(event.target.value as TeamAbsenceType)}>{Object.entries(TEAM_ABSENCE_TYPES).map(([value, config]) => <option key={value} value={value}>{config.label}</option>)}</select></label>
            <label><span>Pessoa</span><select value={absenceMemberId} onChange={(event) => setAbsenceMemberId(event.target.value)}><option value="">Escolher…</option>{members.map((member) => <option key={member.id} value={member.id}>{member.name}</option>)}</select></label>
            <label><span>Dia</span><select value={absenceDay} onChange={(event) => setAbsenceDay(Number(event.target.value))}>{TEAM_WEEKDAYS.map((day, index) => <option key={day} value={index}>{day}</option>)}</select></label>
            <label><span>Início</span><select value={absenceStartHour} onChange={(event) => setAbsenceStartHour(Number(event.target.value))}>{Array.from({ length: 8 }, (_, index) => <option key={index} value={index}>Hora {index + 1}</option>)}</select></label>
            <label><span>Horas</span><input type="number" min="1" max="40" value={customHours} onChange={(event) => setCustomHours(Number(event.target.value) || 1)} /></label>
            <label className="custom-allocation-note"><span>Nota opcional</span><input value={customTitle} onChange={(event) => setCustomTitle(event.target.value)} placeholder={TEAM_ABSENCE_TYPES[absenceType].label} /></label>
            <button className="secondary-button" onClick={addCustomAllocation}>Adicionar à semana</button>
          </div>
        </aside>

        <section className="panel team-board-panel">
          <div className="team-board-toolbar">
            <div className="team-add-member"><input value={newMemberName} onChange={(event) => setNewMemberName(event.target.value)} onKeyDown={(event) => { if (event.key === "Enter") addMember(); }} placeholder="Nome da pessoa" /><button onClick={addMember}>+ Adicionar pessoa</button></div>
            <div><button className="secondary-button" onClick={exportTeamCsv}>Exportar CSV</button><button className="secondary-button" onClick={copyForExcel}>Copiar para Excel</button><button className="primary-button" onClick={copyForEmail}>Copiar para e-mail</button></div>
          </div>

          {selectedAllocation && (
            <div className="team-allocation-editor">
              <strong>{allocationText(selectedAllocation)}</strong>
              <label><span>Pessoa</span><select value={selectedAllocation.memberId} onChange={(event) => updateAllocation(selectedAllocation.id, { memberId: event.target.value })}>{members.map((member) => <option key={member.id} value={member.id}>{member.name}</option>)}</select></label>
              <label><span>Dia</span><select value={Math.floor(selectedAllocation.startSlot / 8)} onChange={(event) => updateAllocation(selectedAllocation.id, { startSlot: Number(event.target.value) * 8 + (selectedAllocation.startSlot % 8) })}>{TEAM_WEEKDAYS.map((day, index) => <option key={day} value={index}>{day}</option>)}</select></label>
              <label><span>Hora</span><select value={selectedAllocation.startSlot % 8} onChange={(event) => updateAllocation(selectedAllocation.id, { startSlot: Math.floor(selectedAllocation.startSlot / 8) * 8 + Number(event.target.value) })}>{Array.from({ length: 8 }, (_, index) => <option key={index} value={index}>{index + 1}</option>)}</select></label>
              <label><span>Horas</span><input type="number" min="1" max={40 - selectedAllocation.startSlot} value={selectedAllocation.hours} onChange={(event) => updateAllocation(selectedAllocation.id, { hours: Number(event.target.value) || 1 })} /></label>
              <button className="danger-button" onClick={() => removeAllocation(selectedAllocation.id)}>Remover</button>
              <button className="icon-button" aria-label="Fechar editor" onClick={() => setSelectedAllocationId(null)}>×</button>
            </div>
          )}

          <div className="team-board-scroll">
            <div className="team-hours-grid team-grid-header">
              <div className="team-grid-corner" style={{ gridRow: 1 }}><strong>Pessoa</strong><small>40 h semanais</small></div>
              {Array.from({ length: 40 }, (_, slot) => <div key={`percent-${slot}`} className={`team-percent-cell ${(slot + 1) % 8 === 0 ? "day-end" : ""}`}>2.5%</div>)}
              <div className="team-grid-corner secondary" style={{ gridRow: 2 }}><span>Semana</span></div>
              {TEAM_WEEKDAYS.map((day, dayIndex) => <div key={day} className="team-day-cell" style={{ gridColumn: `${dayIndex * 8 + 2} / span 8`, gridRow: 2 }}><strong>{day}</strong><small>{formatDate(addDays(weekStart, dayIndex))}</small></div>)}
              <div className="team-grid-corner secondary" style={{ gridRow: 3 }}><span>Hora</span></div>
              {Array.from({ length: 40 }, (_, slot) => <div key={`hour-${slot}`} className={`team-hour-cell ${(slot + 1) % 8 === 0 ? "day-end" : ""}`}>{(slot % 8) + 1}</div>)}
            </div>
            {members.map((member) => {
              const memberAllocations = currentAllocations.filter((allocation) => allocation.memberId === member.id);
              const total = memberAllocations.reduce((sum, allocation) => sum + allocation.hours, 0);
              return (
                <div className="team-hours-grid team-member-row" key={member.id}>
                  <div className="team-member-cell"><div><strong>{member.name}</strong>{member.role && <small>{member.role}</small>}<span>{formatHours(total)} · {Math.round((total / 40) * 100)}%</span></div><button aria-label={`Remover ${member.name}`} onClick={() => removeMember(member)}>×</button></div>
                  {Array.from({ length: 40 }, (_, slot) => {
                    const isPreview = dropPreview?.memberId === member.id && slot >= dropPreview.startSlot && slot < dropPreview.startSlot + dropPreview.hours;
                    return <button key={slot} style={{ gridColumn: slot + 2 }} className={`team-drop-cell ${(slot + 1) % 8 === 0 ? "day-end" : ""} ${isPreview ? dropPreview.valid ? "drop-preview" : "drop-preview-invalid" : ""} ${isPreview && slot === dropPreview.startSlot ? "drop-preview-start" : ""}`} aria-label={`${member.name}, ${TEAM_WEEKDAYS[Math.floor(slot / 8)]}, hora ${(slot % 8) + 1}`} onDragEnter={(event) => { event.preventDefault(); previewDrop(member.id, slot); }} onDragOver={(event) => { event.preventDefault(); event.dataTransfer.dropEffect = draggedAllocationId ? "move" : "copy"; previewDrop(member.id, slot); }} onDrop={(event) => { event.preventDefault(); setDropPreview(null); const allocationId = event.dataTransfer.getData("application/x-work-organizer-allocation") || draggedAllocationId; if (allocationId) { moveAllocation(allocationId, member.id, slot); return; } const issueId = event.dataTransfer.getData("application/x-work-organizer-issue") || draggedIssueId; if (issueId) addAllocation(member.id, slot, issueId); }} />;
                  })}
                  {[7, 15, 23, 31].map((slot) => <div key={`divider-${slot}`} className="team-member-day-divider" style={{ gridColumn: slot + 2 }} />)}
                  {memberAllocations.map((allocation) => {
                    const issue = allocation.issueId ? issueById.get(allocation.issueId) : null;
                    return <div key={allocation.id} role="button" tabIndex={0} draggable className={`team-allocation-block ${draggedAllocationId === allocation.id ? "dragging" : ""} ${!issue ? `absence absence-${allocation.absenceType ?? "other"}` : ""} ${selectedAllocationId === allocation.id ? "selected" : ""}`} style={{ gridColumn: `${allocation.startSlot + 2} / span ${allocation.hours}`, borderColor: issue?.color ?? allocation.customColor ?? "#8b91a7", background: `${issue?.color ?? allocation.customColor ?? "#8b91a7"}20` }} onDragStart={(event) => { event.dataTransfer.setData("application/x-work-organizer-allocation", allocation.id); event.dataTransfer.effectAllowed = "move"; setDraggedAllocationId(allocation.id); }} onDragEnd={() => { setDraggedAllocationId(null); setDropPreview(null); }} onClick={() => setSelectedAllocationId(allocation.id)} onKeyDown={(event) => { if (event.key === "Enter" || event.key === " ") setSelectedAllocationId(allocation.id); if (event.key === "Delete" || event.key === "Backspace") removeAllocation(allocation.id); }} title={`${allocationText(allocation)} · ${formatHours(allocation.hours)} · arrasta para mover`}><button type="button" className="team-allocation-remove" aria-label={`Remover ${allocationText(allocation)}`} title="Remover da alocação" onClick={(event) => { event.stopPropagation(); removeAllocation(allocation.id); }}>×</button><strong>{issue?.project ?? allocation.customTitle}</strong>{issue && (issue.webUrl ? <a href={issue.webUrl} target="_blank" rel="noreferrer" onClick={(event) => event.stopPropagation()}>#{issue.iid} {issue.title}</a> : <span>#{issue.iid} {issue.title}</span>)}{!issue && <span>{TEAM_ABSENCE_TYPES[allocation.absenceType ?? "other"].label} · sem cliente/US</span>}<small>{formatHours(allocation.hours)}</small></div>;
                  })}
                </div>
              );
            })}
            {!members.length && <div className="empty-state"><span>＋</span><strong>Adiciona a primeira pessoa da equipa</strong><p>As pessoas formam as linhas da tabela semanal.</p></div>}
          </div>
          <div className="team-board-legend"><span><i />1 célula = 1 h = 2,5%</span><span>{currentAllocations.length} blocos nesta semana</span><span>{formatHours(currentAllocations.reduce((sum, allocation) => sum + allocation.hours, 0))} alocadas</span></div>
        </section>
      </section>
    </div>
  );
}

function CompletedView({
  tasks,
  todayKey,
}: {
  tasks: Task[];
  todayKey: string;
}) {
  const [period, setPeriod] = useState<Zoom | "all">("week");
  const [search, setSearch] = useState("");
  const [client, setClient] = useState("all");
  const [type, setType] = useState("all");
  const start = periodStart(todayKey, period);
  // This is a record of work allocated in the Timeline. It must not depend on
  // an issue's current state or a later GitLab closure date.
  const timelineRecords: WorkLog[] = tasks
    .map((task) => ({
      id: `timeline-${task.id}`,
      dateKey: task.date,
      client: task.client,
      project: task.project,
      task: task.title,
      description: task.description,
      type: task.phase,
      hours: task.estimate,
      estimate: task.estimate,
      webUrl: task.webUrl,
      projectId: task.projectId,
      issueIid: task.iid,
      source: "timeline" as const,
    }))
    .sort((left, right) => right.dateKey.localeCompare(left.dateKey));
  const filtered = timelineRecords.filter(
    (log) =>
      log.dateKey >= start &&
      (client === "all" || log.client === client) &&
      (type === "all" || log.type === type) &&
      `${log.client} ${log.project} ${log.task}`
        .toLowerCase()
        .includes(search.toLowerCase()),
  );
  const total = filtered.reduce((sum, log) => sum + log.hours, 0);
  const todayTotal = timelineRecords
    .filter((log) => log.dateKey === todayKey)
    .reduce((sum, log) => sum + log.hours, 0);
  return (
    <div className="page">
      <PageIntro
        eyebrow="REGISTO DE TEMPO"
        title="Trabalho realizado"
        description="Alocações registadas na Timeline, organizadas por cliente e projeto."
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
            {[...new Set(timelineRecords.map((log) => log.client))]
              .sort()
              .map((value) => (
              <option key={value}>{value}</option>
              ))}
          </select>
          <select
            value={type}
            onChange={(event) => setType(event.target.value)}
            aria-label="Filtrar por tipo"
          >
            <option value="all">Todos os tipos</option>
            {[...new Set(timelineRecords.map((log) => log.type))]
              .sort()
              .map((value) => (
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
                {log.description && <small>{log.description}</small>}
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
              placeholder="Introduz o token de acesso"
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
    "Data;Projeto;Tipo de Tarefa;Titulo;Descrição;Tempo Gasto",
    ...logs.map((log) =>
      [
        log.dateKey,
        log.project,
        log.type,
        log.task,
        log.description ?? "",
        log.hours,
      ]
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
