"use client";

import { useEffect, useMemo, useRef, useState, type FormEvent } from "react";
import { parseExcelFile } from "./excel-import";
import { exportDailyWorkbook, exportMonthlyWorkbook } from "./excel-export";

type Category = "wechat" | "remix" | "original";
type TaskCategory = Category | "duty" | "bonus";
type TaskType =
  | "wechat_article"
  | "video_edit"
  | "field_video"
  | "drama_video"
  | "agency_wechat"
  | "weibo"
  | "live"
  | "platform_upload"
  | "co_micro"
  | "ai_video"
  | "park_hr_micro"
  | "duty_editor"
  | "wechat_weekly"
  | "wechat_forward"
  | "video_replay_upload"
  | "news_program_upload"
  | "command_phone";

type Entry = {
  id: string;
  title: string;
  staff: string;
  views: string;
  duration: string;
  manualPoints: number | null;
  sourcePoints: number | null;
  notes: string;
  taskType?: TaskType;
};

type Day = {
  date: string;
  weekday: string;
  wechatEditor: string;
  videoEditor: string;
  dutyEditor: string;
  dutyDirector: string;
  supervisor: string;
  sections: Record<Category, Entry[]>;
};

function exportEntryPoints(category: Category, entry: Entry, reward = 0) {
  return entryPoints(category, entry, reward);
}

type LegacyDay = Omit<Day, "wechatEditor" | "videoEditor" | "dutyEditor" | "dutyDirector" | "supervisor"> &
  Partial<Pick<Day, "wechatEditor" | "videoEditor" | "dutyEditor" | "dutyDirector" | "supervisor">> & {
  microEditor?: string;
  duty?: string;
};

type SheetData = {
  title: string;
  members: string[];
  days: Day[];
  lastImport?: {
    fileName: string;
    importedAt: string;
    months: string[];
  };
};

type PersonScore = {
  name: string;
  duty: number;
  wechat: number;
  remix: number;
  original: number;
  fixed: number;
  total: number;
  pieces: number;
};

type PersonTask = {
  date: string;
  category: TaskCategory;
  entry: Entry;
  totalPoints: number;
  personalPoints: number;
  rewardPoints: number;
};

type RankedVideo = {
  date: string;
  category: "remix" | "original";
  entry: Entry;
  views: number;
  people: string[];
  rewardPoints: number;
};

type NewsCenterScore = {
  name: string;
  order: number;
  forwardCount: number;
  forward: number;
  micro: number;
  video: number;
  justNow: number;
  total: number;
};

type ImportMode = "day" | "month";

const STORAGE_KEY = "nanping-media-rundown-calendar-v2";
const LEGACY_STORAGE_KEY = "nanping-media-rundown-2026-07-v1";
const RMB_PER_POINT = 25;
const NEW_MEDIA_TEAM = [
  "张瑞君",
  "周婷",
  "张笑弛",
  "王馨",
  "吴轲宇",
  "高洁",
  "黄琦",
  "周俊",
  "牛文静",
  "徐卓凡",
  "龚启涛",
  "梁斌",
  "李晨雨",
  "陈晓强",
  "刘乐",
] as const;
const NEW_MEDIA_TEAM_SET = new Set<string>(NEW_MEDIA_TEAM);
const NEWS_CENTER_TEAM = [
  "蒋超", "徐卓凡", "余华尊", "倪婷婷", "黄琦", "寿洪清", "郑晖", "郑晟", "陈晓强", "危义铭",
  "黄益平", "周俊", "吴丹", "黎志刚", "包剑武", "张瑞君", "许文婷", "胡志雄", "吴骁", "伍道微",
  "杨志林", "胡宗榕", "陈慧强", "张笑弛", "翁崇毅", "邱太文", "林俊涛", "周婷", "周颖", "王馨",
  "王小川", "林世利", "黄慧娟", "吴轲宇", "陈玲珑", "周琪圆", "刘乐", "龚启涛", "梁斌", "张云婷",
  "陈仁", "王晓飞", "潘东浩", "牛文静", "李晨雨",
] as const;
const NEWS_CENTER_TEAM_SET = new Set<string>(NEWS_CENTER_TEAM);
const WECHAT_WEEKLY_POINTS = 1200 / RMB_PER_POINT;
const DEFAULT_FORWARD_COUNT = 40;
const DEFAULT_FORWARD_POINTS = 1.6;
const ADMIN_SESSION_KEY = "nanping-media-admin-until";
const ADMIN_SESSION_MS = 2 * 60 * 60 * 1000;
const ADMIN_PASSWORD_HASH = "5600715f42bf51c40dc330d750cd996f58fead4ddea56466ce7498d17801b3a5";
const DUTY_EDITOR_CANDIDATES = ["张瑞君", "周婷", "张笑弛"] as const;
const WECHAT_EDITOR_CANDIDATES = ["王馨", "吴轲宇", "高洁"] as const;
const MONTHLY_FIXED_BONUSES = new Map<string, { taskType: TaskType; points: number }>([
  ["周婷", { taskType: "video_replay_upload", points: 24 }],
  ["刘乐", { taskType: "news_program_upload", points: 36 }],
  ["吴轲宇", { taskType: "command_phone", points: 20 }],
]);

const categoryMeta: Record<
  Category,
  { label: string; short: string; description: string; className: string }
> = {
  wechat: {
    label: "微信公众号",
    short: "微信",
    description: "文章、微刊与刚刚帖",
    className: "wechat",
  },
  remix: {
    label: "小编二创短视频",
    short: "二创",
    description: "含制作封面、包框",
    className: "remix",
  },
  original: {
    label: "记者原创短视频",
    short: "原创",
    description: "新媒体首发",
    className: "original",
  },
};

const taskMeta: Record<TaskType, { label: string; unit: string; points: number }> = {
  wechat_article: { label: "微信/微刊", unit: "阅读奖励", points: 0 },
  video_edit: { label: "短视频编辑", unit: "1D/条", points: 2 },
  field_video: { label: "新闻外采", unit: "1C/条", points: 5 },
  drama_video: { label: "剧情短视频", unit: "A+B起评", points: 20 },
  agency_wechat: { label: "公众号代运营", unit: "3D/组", points: 6 },
  weibo: { label: "微博发布", unit: "0.5D/条", points: 1 },
  live: { label: "连线/直播拉流", unit: "1D/场", points: 2 },
  platform_upload: { label: "整档上传平台", unit: "3A/月", points: 36 },
  co_micro: { label: "服务共建微刊", unit: "3D/组", points: 6 },
  ai_video: { label: "策划纯AI视频", unit: "1B/条", points: 8 },
  park_hr_micro: { label: "工业园区/人社局微刊", unit: "3D/天", points: 6 },
  duty_editor: { label: "值班责编", unit: "1B/天", points: 8 },
  wechat_weekly: { label: "微信公众号值班小编", unit: "¥1,200/周", points: WECHAT_WEEKLY_POINTS },
  wechat_forward: { label: "微信转发", unit: "默认40条/1.6分", points: DEFAULT_FORWARD_POINTS },
  video_replay_upload: { label: "视频回看每日上传", unit: "24分/月", points: 24 },
  news_program_upload: { label: "新闻节目上传", unit: "36分/月", points: 36 },
  command_phone: { label: "接听指令电话", unit: "20分/月", points: 20 },
};

const entryTaskTypes = (Object.keys(taskMeta) as TaskType[]).filter(
  (taskType) => !["duty_editor", "wechat_weekly", "wechat_forward", "video_replay_upload", "news_program_upload", "command_phone"].includes(taskType),
);

function defaultTaskType(category: Category): TaskType {
  if (category === "wechat") return "wechat_article";
  if (category === "remix") return "video_edit";
  return "field_video";
}

function parseViews(value: string) {
  const normalized = String(value ?? "").replace(/,/g, "").trim();
  const wan = normalized.match(/(\d+(?:\.\d+)?)\s*万/);
  if (wan) return Number(wan[1]) * 10000;
  const plain = normalized.match(/\d+(?:\.\d+)?/);
  return plain ? Number(plain[0]) : 0;
}

function microReadingReward(viewsValue: string) {
  const views = parseViews(viewsValue);
  if (views >= 1000000) return 24;
  if (views >= 100000) return 12;
  if (views >= 10000) return 5;
  if (views > 5000) return 2;
  return 0;
}

function justNowPoints(value: string) {
  const normalized = String(value ?? "").replace(/,/g, "").trim();
  const matched = normalized.match(/\d+(?:\.\d+)?/);
  return matched ? Number(matched[0]) : 0;
}

function weekStart(dateString: string) {
  const [year, month, day] = dateString.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  const offset = (date.getUTCDay() + 6) % 7;
  date.setUTCDate(date.getUTCDate() - offset);
  return dateKey(date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate());
}

function entryPoints(category: Category, entry: Entry, monthlyVideoReward = 0) {
  if (category === "remix" && !entry.staff.trim()) return 0;
  if (entry.manualPoints !== null) return entry.manualPoints;
  const taskType = entry.taskType ?? defaultTaskType(category);
  const base = taskMeta[taskType].points;
  const readingReward = category === "wechat" ? microReadingReward(entry.views) : monthlyVideoReward;
  return base + readingReward;
}

function round(value: number) {
  return Math.round((value + Number.EPSILON) * 10) / 10;
}

function formatPoints(value: number) {
  return Number.isInteger(round(value)) ? String(round(value)) : round(value).toFixed(1);
}

function formatMoney(points: number) {
  return new Intl.NumberFormat("zh-CN", {
    style: "currency",
    currency: "CNY",
    maximumFractionDigits: 0,
  }).format(points * RMB_PER_POINT);
}

function unitForPoints(points: number) {
  const known = new Map([
    [1, "0.5D"],
    [2, "1D"],
    [5, "1C"],
    [6, "3D"],
    [8, "1B"],
    [12, "1A"],
    [20, "A+B"],
    [24, "2A"],
    [36, "3A"],
  ]);
  return known.get(round(points)) ?? `${formatPoints(points)}分`;
}

function peopleFor(staff: string, members: string[]) {
  const cleaned = String(staff ?? "")
    .replace(/(AI|ai)制作[：:]?/g, "")
    .replace(/(摄制|编辑|采编|记者)[：:]?/g, "")
    .trim();
  const found = members.filter((member) => cleaned.includes(member));
  if (found.length) return [...new Set(found)];
  if (!cleaned || cleaned === "转载" || cleaned === "无") return [];
  return cleaned
    .split(/[、，,；;\/\s]+/)
    .map((name) => name.trim())
    .filter((name) => name.length >= 2 && name.length <= 4);
}

function entryPeople(entry: Entry, day: Day, members: string[]) {
  if (entry.taskType === "park_hr_micro") {
    return peopleFor(day.wechatEditor, members);
  }
  return peopleFor(entry.staff, members);
}

function dayLabel(dateString: string) {
  return Number(dateString.slice(-2));
}

function formatDate(dateString: string) {
  const [year, month, day] = dateString.split("-");
  return `${year}年${Number(month)}月${Number(day)}日`;
}

function dateKey(year: number, month: number, day: number) {
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function monthKey(year: number, month: number) {
  return `${year}-${String(month).padStart(2, "0")}`;
}

function legacyDutyName(value: string, label: "值班责编" | "值班主任" | "监审") {
  const normalized = String(value ?? "").replace(/[\r\n]+/g, " ").trim();
  if (!normalized) return "";
  const nextLabels = label === "值班责编" ? "值班主任|监审" : label === "值班主任" ? "监审|值班责编" : "值班责编|值班主任";
  const match = normalized.match(
    new RegExp(`${label}\\s*[：:]?\\s*(.*?)(?=\\s*(?:${nextLabels})\\s*[：:]|$)`),
  );
  return match?.[1]?.replace(/^[，,；;|\s]+|[，,；;|\s]+$/g, "").trim() ?? "";
}

function normalizeDay(day: LegacyDay): Day {
  const legacyDuty = day.duty ?? "";
  const fallbackEditor = !/(值班责编|值班主任|监审)/.test(legacyDuty) ? legacyDuty.trim() : "";
  return {
    date: day.date,
    weekday: day.weekday,
    wechatEditor: day.wechatEditor ?? day.microEditor ?? "",
    videoEditor: day.videoEditor ?? "",
    dutyEditor: day.dutyEditor ?? (legacyDutyName(legacyDuty, "值班责编") || fallbackEditor),
    dutyDirector: day.dutyDirector ?? legacyDutyName(legacyDuty, "值班主任"),
    supervisor: day.supervisor ?? legacyDutyName(legacyDuty, "监审"),
    sections: day.sections,
  };
}

function normalizeSheet(source: SheetData): SheetData {
  return {
    ...source,
    days: source.days.map((day) => normalizeDay(day as LegacyDay)),
  };
}

function blankDay(year: number, month: number, day: number): Day {
  const date = new Date(year, month - 1, day);
  const week = ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"];
  return {
    date: dateKey(year, month, day),
    weekday: week[date.getDay()],
    wechatEditor: "",
    videoEditor: "",
    dutyEditor: "",
    dutyDirector: "",
    supervisor: "",
    sections: { wechat: [], remix: [], original: [] },
  };
}

function ensureMonth(source: SheetData, year: number, month: number) {
  const daysInMonth = new Date(year, month, 0).getDate();
  const existing = new Set(source.days.map((day) => day.date));
  const missing: Day[] = [];
  for (let day = 1; day <= daysInMonth; day += 1) {
    const key = dateKey(year, month, day);
    if (!existing.has(key)) missing.push(blankDay(year, month, day));
  }
  if (!missing.length) return source;
  return {
    ...source,
    title: "新媒体全年串联单",
    days: [...source.days, ...missing].sort((a, b) => a.date.localeCompare(b.date)),
  };
}

function dayHasContent(day: Day) {
  return Boolean(
    day.wechatEditor ||
      day.dutyEditor ||
      day.dutyDirector ||
      day.supervisor ||
      day.sections.wechat.length ||
      day.sections.remix.length ||
      day.sections.original.length,
  );
}

function downloadText(filename: string, content: string, type: string) {
  const blob = new Blob(["\uFEFF", content], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function csvCell(value: string | number) {
  return `"${String(value).replace(/"/g, '""')}"`;
}

async function passwordHash(value: string) {
  const bytes = new TextEncoder().encode(value);
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

function adminAccessActive() {
  return typeof window !== "undefined" && Number(localStorage.getItem(ADMIN_SESSION_KEY) ?? 0) > Date.now();
}

export default function Home() {
  const today = useRef(new Date());
  const [data, setData] = useState<SheetData | null>(null);
  const initialData = useRef<SheetData | null>(null);
  const [selectedYear, setSelectedYear] = useState(today.current.getFullYear());
  const [selectedMonth, setSelectedMonth] = useState(today.current.getMonth() + 1);
  const [selectedDate, setSelectedDate] = useState(
    dateKey(today.current.getFullYear(), today.current.getMonth() + 1, today.current.getDate()),
  );
  const [view, setView] = useState<"daily" | "summary">("daily");
  const [query, setQuery] = useState("");
  const [personQuery, setPersonQuery] = useState("");
  const [saved, setSaved] = useState(true);
  const [showRules, setShowRules] = useState(false);
  const [toast, setToast] = useState("");
  const [isImporting, setIsImporting] = useState(false);
  const [isMusicPlaying, setIsMusicPlaying] = useState(false);
  const [isMusicHidden, setIsMusicHidden] = useState(false);
  const [isAdmin, setIsAdmin] = useState(false);
  const [showAdminLogin, setShowAdminLogin] = useState(false);
  const [adminPassword, setAdminPassword] = useState("");
  const [adminError, setAdminError] = useState("");
  const [openCandidatePicker, setOpenCandidatePicker] = useState<"duty" | "wechat" | null>(null);
  const hydrated = useRef(false);
  const dayFileInputRef = useRef<HTMLInputElement>(null);
  const monthFileInputRef = useRef<HTMLInputElement>(null);
  const audioRef = useRef<HTMLAudioElement>(null);
  const pendingAdminAction = useRef<null | (() => void)>(null);

  useEffect(() => {
    const checkAdmin = () => {
      const adminUntil = Number(localStorage.getItem(ADMIN_SESSION_KEY) ?? 0);
      setIsAdmin(adminUntil > Date.now());
      if (adminUntil && adminUntil <= Date.now()) localStorage.removeItem(ADMIN_SESSION_KEY);
    };
    checkAdmin();
    const timer = window.setInterval(checkAdmin, 60_000);
    return () => window.clearInterval(timer);
  }, []);

  useEffect(() => {
    fetch(`${import.meta.env.BASE_URL}seed.json`)
      .then((response) => response.json())
      .then((seed: SheetData) => {
        const normalizedSeed = normalizeSheet({ ...seed, title: "新媒体全年串联单" });
        initialData.current = normalizedSeed;
        const stored = localStorage.getItem(STORAGE_KEY) ?? localStorage.getItem(LEGACY_STORAGE_KEY);
        if (stored) {
          try {
            setData(ensureMonth(normalizeSheet(JSON.parse(stored)), selectedYear, selectedMonth));
          } catch {
            setData(ensureMonth(normalizedSeed, selectedYear, selectedMonth));
          }
        } else {
          setData(ensureMonth(normalizedSeed, selectedYear, selectedMonth));
        }
        hydrated.current = true;
      });
    // Initial import happens once; later month changes expand the same working calendar.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (!data) return;
    const expanded = ensureMonth(data, selectedYear, selectedMonth);
    if (expanded !== data) setData(expanded);
    const desiredDay =
      selectedYear === today.current.getFullYear() && selectedMonth === today.current.getMonth() + 1
        ? today.current.getDate()
        : 1;
    setSelectedDate(dateKey(selectedYear, selectedMonth, desiredDay));
    // Month changes intentionally select today or the first day of that month.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedYear, selectedMonth]);

  useEffect(() => {
    if (!data || !hydrated.current) return;
    setSaved(false);
    const timer = window.setTimeout(() => {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
      setSaved(true);
    }, 420);
    return () => window.clearTimeout(timer);
  }, [data]);

  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(""), 2400);
    return () => window.clearTimeout(timer);
  }, [toast]);

  useEffect(() => {
    const audio = audioRef.current;
    if (!audio) return;
    audio.volume = 0.7;
    void audio.play().catch(() => {
      // Browsers may block audible autoplay until the visitor interacts with the page.
      setIsMusicPlaying(false);
    });
  }, []);

  function toggleMusic() {
    const audio = audioRef.current;
    if (!audio) return;
    if (audio.paused) {
      void audio.play().catch(() => setIsMusicPlaying(false));
    } else {
      audio.pause();
    }
  }

  function requireAdmin(action?: () => void) {
    const adminUntil = Number(localStorage.getItem(ADMIN_SESSION_KEY) ?? 0);
    if (adminUntil > Date.now()) {
      setIsAdmin(true);
      action?.();
      return true;
    }
    pendingAdminAction.current = action ?? null;
    setAdminPassword("");
    setAdminError("");
    setShowAdminLogin(true);
    return false;
  }

  async function submitAdminPassword(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if ((await passwordHash(adminPassword)) !== ADMIN_PASSWORD_HASH) {
      setAdminError("密码不正确，请重新输入");
      return;
    }
    localStorage.setItem(ADMIN_SESSION_KEY, String(Date.now() + ADMIN_SESSION_MS));
    setIsAdmin(true);
    setShowAdminLogin(false);
    setAdminPassword("");
    setToast("管理员验证成功，2小时内可编辑");
    const action = pendingAdminAction.current;
    pendingAdminAction.current = null;
    action?.();
  }

  function requestEditAccess() {
    if (!isAdmin) requireAdmin();
  }

  function toggleCandidatePicker(picker: "duty" | "wechat") {
    if (!adminAccessActive()) {
      requireAdmin(() => setOpenCandidatePicker(picker));
      return;
    }
    setOpenCandidatePicker((current) => current === picker ? null : picker);
  }

  function chooseCandidate(field: "dutyEditor" | "wechatEditor", name: string) {
    updateDay({ [field]: name });
    setOpenCandidatePicker(null);
  }

  const selectedDay = useMemo(
    () => data?.days.find((day) => day.date === selectedDate),
    [data, selectedDate],
  );

  const visibleDays = useMemo(() => {
    const prefix = monthKey(selectedYear, selectedMonth);
    return data?.days.filter((day) => day.date.startsWith(prefix)) ?? [];
  }, [data, selectedMonth, selectedYear]);

  const monthlyTopVideos = useMemo<RankedVideo[]>(() => {
    if (!data) return [];
    return visibleDays
      .flatMap((day) => [
        ...day.sections.remix.map((entry) => ({
          date: day.date,
          category: "remix" as const,
          entry,
          views: parseViews(entry.views),
          people: entryPeople(entry, day, data.members),
        })),
        ...day.sections.original.map((entry) => ({
          date: day.date,
          category: "original" as const,
          entry,
          views: parseViews(entry.views),
          people: entryPeople(entry, day, data.members),
        })),
      ])
      .filter(({ views, people }) => views > 0 && people.length > 0)
      .sort((a, b) => b.views - a.views || a.date.localeCompare(b.date))
      .slice(0, 10)
      .map((video) => ({
        ...video,
        rewardPoints:
          video.views >= 10000000 ? 36 :
          video.views >= 1000000 ? 24 :
          video.views >= 100000 ? 12 : 8,
      }));
  }, [data, visibleDays]);

  const monthlyVideoRewards = useMemo(() => {
    const rewards = new Map<string, number>();
    monthlyTopVideos.forEach(({ entry, rewardPoints }) => {
      rewards.set(entry.id, rewardPoints);
    });
    return rewards;
  }, [monthlyTopVideos]);

  const scores = useMemo<PersonScore[]>(() => {
    if (!data) return [];
    const totals = new Map<string, PersonScore>();
    visibleDays.forEach((day) => {
      const dutyPeople = peopleFor(day.dutyEditor, data.members);
      const dutyShare = dutyPeople.length ? taskMeta.duty_editor.points / dutyPeople.length : 0;
      dutyPeople.forEach((name) => {
        const current = totals.get(name) ?? {
          name,
          duty: 0,
          wechat: 0,
          remix: 0,
          original: 0,
          fixed: 0,
          total: 0,
          pieces: 0,
        };
        current.duty += dutyShare;
        current.total += dutyShare;
        current.pieces += 1;
        totals.set(name, current);
      });
      (Object.keys(day.sections) as Category[]).forEach((category) => {
        day.sections[category].forEach((entry) => {
          const people = entryPeople(entry, day, data.members);
          if (!people.length) return;
          const share = entryPoints(category, entry, monthlyVideoRewards.get(entry.id) ?? 0) / people.length;
          people.forEach((name) => {
            const current = totals.get(name) ?? {
              name,
              duty: 0,
              wechat: 0,
              remix: 0,
              original: 0,
              fixed: 0,
              total: 0,
              pieces: 0,
            };
            current[category] += share;
            current.total += share;
            current.pieces += 1;
            totals.set(name, current);
          });
        });
      });
    });
    const weeklyEditors = new Map<string, Map<string, number>>();
    visibleDays.forEach((day) => {
      const editors = peopleFor(day.wechatEditor, data.members);
      if (!editors.length) return;
      const start = weekStart(day.date);
      editors.forEach((name) => {
        const weeks = weeklyEditors.get(name) ?? new Map<string, number>();
        weeks.set(start, (weeks.get(start) ?? 0) + 1 / editors.length);
        weeklyEditors.set(name, weeks);
      });
    });
    weeklyEditors.forEach((weeks, name) => {
      const weeklyPoints = round(
        [...weeks.values()].reduce((sum, assignedDays) => sum + (assignedDays / 7) * WECHAT_WEEKLY_POINTS, 0),
      );
      const current = totals.get(name) ?? {
        name,
        duty: 0,
        wechat: 0,
        remix: 0,
        original: 0,
        fixed: 0,
        total: 0,
        pieces: 0,
      };
      current.wechat += weeklyPoints;
      current.total += weeklyPoints;
      current.pieces += weeks.size;
      totals.set(name, current);
    });
    NEW_MEDIA_TEAM.forEach((name) => {
      if (!totals.has(name)) {
        totals.set(name, {
          name,
          duty: 0,
          wechat: 0,
          remix: 0,
          original: 0,
          fixed: 0,
          total: 0,
          pieces: 0,
        });
      }
    });
    MONTHLY_FIXED_BONUSES.forEach(({ points }, name) => {
      const current = totals.get(name);
      if (!current) return;
      current.fixed += points;
      current.total += points;
      current.pieces += 1;
    });
    return [...totals.values()].sort((a, b) => b.total - a.total || a.name.localeCompare(b.name, "zh-CN"));
  }, [data, monthlyVideoRewards, visibleDays]);

  const newsCenterScores = useMemo<NewsCenterScore[]>(() => {
    if (!data) return [];
    const totals = new Map<string, Omit<NewsCenterScore, "name" | "order" | "forwardCount" | "forward" | "total">>();
    NEWS_CENTER_TEAM.forEach((name) => totals.set(name, { micro: 0, video: 0, justNow: 0 }));
    visibleDays.forEach((day) => {
      day.sections.wechat.forEach((entry) => {
        const people = entryPeople(entry, day, data.members).filter((name) => NEWS_CENTER_TEAM_SET.has(name));
        if (!people.length) return;
        const rewardShare = microReadingReward(entry.views) / people.length;
        const justNowShare = justNowPoints(entry.duration) / people.length;
        people.forEach((name) => {
          if (NEW_MEDIA_TEAM_SET.has(name)) return;
          const current = totals.get(name);
          if (!current) return;
          current.micro += rewardShare;
          current.justNow += justNowShare;
        });
      });
      (["remix", "original"] as Category[]).forEach((category) => {
        day.sections[category].forEach((entry) => {
          const people = entryPeople(entry, day, data.members).filter((name) => NEWS_CENTER_TEAM_SET.has(name));
          if (!people.length) return;
          const share = entryPoints(category, entry, monthlyVideoRewards.get(entry.id) ?? 0) / people.length;
          people.forEach((name) => {
            if (NEW_MEDIA_TEAM_SET.has(name)) return;
            const current = totals.get(name);
            if (current) current.video += share;
          });
        });
      });
    });
    return NEWS_CENTER_TEAM.map((name, index) => {
      const current = totals.get(name) ?? { micro: 0, video: 0, justNow: 0 };
      const micro = round(current.micro);
      const video = round(current.video);
      const justNow = round(current.justNow);
      return {
        name,
        order: index + 1,
        forwardCount: DEFAULT_FORWARD_COUNT,
        forward: DEFAULT_FORWARD_POINTS,
        micro,
        video,
        justNow,
        total: round(DEFAULT_FORWARD_POINTS + micro + video + justNow),
      };
    });
  }, [data, monthlyVideoRewards, visibleDays]);

  const totalEntries = useMemo(() => {
    if (!data) return 0;
    return visibleDays.reduce(
      (sum, day) =>
        sum +
        day.sections.wechat.length +
        day.sections.remix.length +
        day.sections.original.length,
      0,
    );
  }, [data, visibleDays]);

  const totalPoints = useMemo(
    () =>
      scores.filter((person) => NEW_MEDIA_TEAM_SET.has(person.name)).reduce((sum, person) => sum + person.total, 0) +
      newsCenterScores.reduce((sum, person) => sum + person.total, 0),
    [newsCenterScores, scores],
  );

  const scoreGroups = useMemo(
    () => [
      {
        key: "new-media",
        title: "新媒体人员绩效",
        description: "指定新媒体人员 · 按当月串单计算工分排名",
        people: scores
          .filter((person) => NEW_MEDIA_TEAM_SET.has(person.name))
          .map((person, index) => ({ person, rank: index + 1 })),
      },
    ],
    [scores],
  );

  const filteredScoreGroups = useMemo(
    () =>
      scoreGroups.map((group) => ({
        ...group,
        people: group.people.filter(({ person }) => person.name.includes(personQuery.trim())),
      })),
    [personQuery, scoreGroups],
  );

  const filteredScores = useMemo(
    () => filteredScoreGroups.flatMap((group) => group.people),
    [filteredScoreGroups],
  );

  const filteredNewsCenterScores = useMemo(
    () => newsCenterScores.filter((person) => person.name.includes(personQuery.trim())),
    [newsCenterScores, personQuery],
  );

  const rankedScores = useMemo(
    () => scoreGroups.flatMap((group) => group.people.map((item) => ({ ...item, group: group.title }))),
    [scoreGroups],
  );

  const focusedPerson = useMemo(() => {
    const exact = scores.find((person) => person.name === personQuery.trim()) ??
      newsCenterScores.find((person) => person.name === personQuery.trim());
    if (exact) return exact.name;
    return personQuery.trim() && filteredScores.length === 1 ? filteredScores[0].person.name : "";
  }, [filteredScores, newsCenterScores, personQuery, scores]);

  const personTasks = useMemo<PersonTask[]>(() => {
    if (!data || !focusedPerson) return [];
    const tasks: PersonTask[] = [];
    const isNewMediaPerson = NEW_MEDIA_TEAM_SET.has(focusedPerson);
    visibleDays.forEach((day) => {
      const dutyPeople = peopleFor(day.dutyEditor, data.members);
      if (isNewMediaPerson && dutyPeople.includes(focusedPerson)) {
        tasks.push({
          date: day.date,
          category: "duty",
          entry: {
            id: `${day.date}-duty-editor`,
            title: "当日值班责编",
            staff: day.dutyEditor,
            views: "",
            duration: "",
            manualPoints: null,
            sourcePoints: null,
            notes: [day.dutyDirector && `值班主任：${day.dutyDirector}`, day.supervisor && `监审：${day.supervisor}`]
              .filter(Boolean)
              .join("；"),
            taskType: "duty_editor",
          },
          totalPoints: taskMeta.duty_editor.points,
          personalPoints: taskMeta.duty_editor.points / dutyPeople.length,
          rewardPoints: 0,
        });
      }
      (Object.keys(day.sections) as Category[]).forEach((category) => {
        day.sections[category].forEach((entry) => {
          const people = entryPeople(entry, day, data.members);
          if (!people.includes(focusedPerson)) return;
          const rewardPoints = monthlyVideoRewards.get(entry.id) ?? 0;
          if (!isNewMediaPerson && category === "wechat") {
            const total = microReadingReward(entry.views) + justNowPoints(entry.duration);
            if (!total) return;
            tasks.push({
              date: day.date,
              category,
              entry,
              totalPoints: total,
              personalPoints: total / people.length,
              rewardPoints: microReadingReward(entry.views),
            });
            return;
          }
          const total = entryPoints(category, entry, rewardPoints);
          tasks.push({
            date: day.date,
            category,
            entry,
            totalPoints: total,
            personalPoints: total / people.length,
            rewardPoints: entry.manualPoints === null ? rewardPoints : 0,
          });
        });
      });
    });
    if (isNewMediaPerson) {
      const editorWeeks = new Map<string, { days: number; firstDate: string }>();
      visibleDays.forEach((day) => {
        const editors = peopleFor(day.wechatEditor, data.members);
        if (!editors.includes(focusedPerson)) return;
        const start = weekStart(day.date);
        const current = editorWeeks.get(start) ?? { days: 0, firstDate: day.date };
        current.days += 1 / editors.length;
        editorWeeks.set(start, current);
      });
      editorWeeks.forEach(({ days, firstDate }, start) => {
        const points = round((days / 7) * WECHAT_WEEKLY_POINTS);
        tasks.push({
          date: firstDate,
          category: "wechat",
          entry: {
            id: `${start}-wechat-weekly-${focusedPerson}`,
            title: `微信公众号值班小编（${start}起，折算${formatPoints(days)}天）`,
            staff: focusedPerson,
            views: "",
            duration: "",
            manualPoints: points,
            sourcePoints: points,
            notes: "¥1,200/周，按当月实际值班天数折算",
            taskType: "wechat_weekly",
          },
          totalPoints: points,
          personalPoints: points,
          rewardPoints: 0,
        });
      });
      const fixedBonus = MONTHLY_FIXED_BONUSES.get(focusedPerson);
      if (fixedBonus) {
        tasks.push({
          date: dateKey(selectedYear, selectedMonth, 1),
          category: "bonus",
          entry: {
            id: `${monthKey(selectedYear, selectedMonth)}-${fixedBonus.taskType}`,
            title: taskMeta[fixedBonus.taskType].label,
            staff: focusedPerson,
            views: "",
            duration: "",
            manualPoints: fixedBonus.points,
            sourcePoints: fixedBonus.points,
            notes: "每月固定绩效",
            taskType: fixedBonus.taskType,
          },
          totalPoints: fixedBonus.points,
          personalPoints: fixedBonus.points,
          rewardPoints: 0,
        });
      }
    }
    return tasks.sort((a, b) => a.date.localeCompare(b.date));
  }, [data, focusedPerson, monthlyVideoRewards, visibleDays]);

  function updateDay(patch: Partial<Day>) {
    if (!data || !selectedDay || !adminAccessActive()) return;
    setData({
      ...data,
      days: data.days.map((day) =>
        day.date === selectedDay.date ? { ...day, ...patch } : day,
      ),
    });
  }

  function updateEntry(category: Category, id: string, patch: Partial<Entry>) {
    if (!selectedDay || !adminAccessActive()) return;
    updateDay({
      sections: {
        ...selectedDay.sections,
        [category]: selectedDay.sections[category].map((entry) =>
          entry.id === id ? { ...entry, ...patch } : entry,
        ),
      },
    });
  }

  function addEntry(category: Category, taskType = defaultTaskType(category)) {
    if (!selectedDay || !adminAccessActive()) return;
    const entry: Entry = {
      id: `${selectedDay.date}-${category}-${Date.now()}`,
      title: "",
      staff: "",
      views: "",
      duration: "",
      manualPoints: null,
      sourcePoints: null,
      notes: "",
      taskType,
    };
    updateDay({
      sections: {
        ...selectedDay.sections,
        [category]: [...selectedDay.sections[category], entry],
      },
    });
    setToast(`已新增：${taskMeta[taskType].label}（${taskMeta[taskType].unit}）`);
  }

  function removeEntry(category: Category, id: string) {
    if (!selectedDay || !adminAccessActive()) return;
    updateDay({
      sections: {
        ...selectedDay.sections,
        [category]: selectedDay.sections[category].filter((entry) => entry.id !== id),
      },
    });
    setToast("已删除，可通过“恢复原始数据”找回");
  }

  function resetData() {
    if (!initialData.current) return;
    if (!adminAccessActive()) {
      requireAdmin(resetData);
      return;
    }
    if (!window.confirm("确认恢复原始数据？全年所有修改都会被清除，2026年7月将恢复为上传的 Excel 内容。")) return;
    setData(ensureMonth(structuredClone(initialData.current), selectedYear, selectedMonth));
    localStorage.removeItem(STORAGE_KEY);
    localStorage.removeItem(LEGACY_STORAGE_KEY);
    setToast("已恢复原始数据");
  }

  function exportSummaryCsv() {
    if (!data) return;
    const rows = [
      ["绩效分类", "序号/排名", "姓名", "微信转发", "微刊（分）", "短视频（分）", "刚刚帖（分）", "值班责编（分）", "新媒体微信（分）", "二创（分）", "原创（分）", "固定月绩效（分）", "参与任务", "合计（分）", "折合人民币"],
      ...rankedScores.map(({ person, rank, group }) => [
        group,
        rank,
        person.name,
        "",
        "",
        "",
        "",
        formatPoints(person.duty),
        formatPoints(person.wechat),
        formatPoints(person.remix),
        formatPoints(person.original),
        formatPoints(person.fixed),
        person.pieces,
        formatPoints(person.total),
        formatMoney(person.total),
      ]),
      ...newsCenterScores.map((person) => [
        "新闻中心绩效",
        person.order,
        person.name,
        `${person.forwardCount}条 / ${formatPoints(person.forward)}分`,
        person.micro,
        person.video,
        person.justNow,
        "",
        "",
        "",
        "",
        "",
        "",
        person.total,
        formatMoney(person.total),
      ]),
    ];
    downloadText(
      `${selectedYear}年${selectedMonth}月新媒体工分汇总.csv`,
      rows.map((row) => row.map(csvCell).join(",")).join("\n"),
      "text/csv;charset=utf-8",
    );
  }

  async function exportDailyRundown() {
    if (!selectedDay) return;
    try {
      await exportDailyWorkbook({
        ...selectedDay,
        sections: {
          wechat: selectedDay.sections.wechat.map((entry) => ({ ...entry, exportPoints: exportEntryPoints("wechat", entry) })),
          remix: selectedDay.sections.remix.map((entry) => ({ ...entry, exportPoints: exportEntryPoints("remix", entry, monthlyVideoRewards.get(entry.id) ?? 0) })),
          original: selectedDay.sections.original.map((entry) => ({ ...entry, exportPoints: exportEntryPoints("original", entry, monthlyVideoRewards.get(entry.id) ?? 0) })),
        },
      });
      setToast(`已按模板导出 ${formatDate(selectedDay.date)} 日串单`);
    } catch (error) {
      window.alert(`导出日串单失败：${error instanceof Error ? error.message : "模板无法读取"}`);
    }
  }

  async function exportTotalWorkbook() {
    if (!data) return;
    try {
      const wechatPerformanceMap = new Map<string, { dutyDays: number; parkHrCount: number }>();
      WECHAT_EDITOR_CANDIDATES.forEach((name) => wechatPerformanceMap.set(name, { dutyDays: 0, parkHrCount: 0 }));
      const videoPerformanceMap = new Map<string, { pieces: number; videoPoints: number; dutyDays: number }>();
      NEW_MEDIA_TEAM.forEach((name) => videoPerformanceMap.set(name, { pieces: 0, videoPoints: 0, dutyDays: 0 }));

      visibleDays.forEach((day) => {
        const wechatEditors = peopleFor(day.wechatEditor, data.members);
        wechatEditors.forEach((name) => {
          const current = wechatPerformanceMap.get(name) ?? { dutyDays: 0, parkHrCount: 0 };
          current.dutyDays += 1 / wechatEditors.length;
          wechatPerformanceMap.set(name, current);
        });
        const parkHrEntries = day.sections.wechat.filter((entry) => entry.taskType === "park_hr_micro");
        wechatEditors.forEach((name) => {
          const current = wechatPerformanceMap.get(name) ?? { dutyDays: 0, parkHrCount: 0 };
          current.parkHrCount += parkHrEntries.length / wechatEditors.length;
          wechatPerformanceMap.set(name, current);
        });

        const allDutyEditors = peopleFor(day.dutyEditor, data.members);
        allDutyEditors.filter((name) => NEW_MEDIA_TEAM_SET.has(name)).forEach((name) => {
          const current = videoPerformanceMap.get(name)!;
          current.dutyDays += 1 / allDutyEditors.length;
        });
        (["remix", "original"] as Category[]).forEach((category) => {
          day.sections[category].forEach((entry) => {
            const allPeople = entryPeople(entry, day, data.members);
            if (!allPeople.length) return;
            const share = entryPoints(category, entry, monthlyVideoRewards.get(entry.id) ?? 0) / allPeople.length;
            allPeople.filter((name) => NEW_MEDIA_TEAM_SET.has(name)).forEach((name) => {
              const current = videoPerformanceMap.get(name)!;
              current.pieces += 1;
              current.videoPoints += share;
            });
          });
        });
      });

      const exportNewsCenterScores = [...newsCenterScores];
      NEW_MEDIA_TEAM.forEach((name) => {
        if (exportNewsCenterScores.some((person) => person.name === name)) return;
        exportNewsCenterScores.push({
          name,
          order: exportNewsCenterScores.length + 1,
          forwardCount: DEFAULT_FORWARD_COUNT,
          forward: DEFAULT_FORWARD_POINTS,
          micro: 0,
          video: 0,
          justNow: 0,
          total: DEFAULT_FORWARD_POINTS,
        });
      });

      await exportMonthlyWorkbook({
        year: selectedYear,
        month: selectedMonth,
        days: visibleDays.map((day) => ({
          ...day,
          sections: {
            wechat: day.sections.wechat.map((entry) => ({ ...entry, exportPoints: exportEntryPoints("wechat", entry) })),
            remix: day.sections.remix.map((entry) => ({ ...entry, exportPoints: exportEntryPoints("remix", entry, monthlyVideoRewards.get(entry.id) ?? 0) })),
            original: day.sections.original.map((entry) => ({ ...entry, exportPoints: exportEntryPoints("original", entry, monthlyVideoRewards.get(entry.id) ?? 0) })),
          },
        })),
        newMediaScores: NEW_MEDIA_TEAM.map((name, index) => ({
          ...(scores.find((person) => person.name === name) ?? {
            name,
            duty: 0,
            wechat: 0,
            remix: 0,
            original: 0,
            fixed: 0,
            total: 0,
            pieces: 0,
          }),
          rank: index + 1,
        })),
        nonMediaScores: scores
          .filter((person) => !NEW_MEDIA_TEAM_SET.has(person.name) && person.total > 0)
          .map((person, index) => ({ ...person, rank: index + 1 })),
        newsCenterScores: exportNewsCenterScores,
        microRewards: visibleDays.flatMap((day) => day.sections.wechat.flatMap((entry) => {
          const views = parseViews(entry.views);
          if (views <= 5000) return [];
          const authors = peopleFor(entry.staff, data.members);
          if (!authors.length) return [];
          const reward = microReadingReward(entry.views);
          return [{
            title: entry.title,
            views,
            staff: authors.join("、"),
            reward,
            notes: entry.notes,
          }];
        })),
        wechatPerformance: [...wechatPerformanceMap.entries()]
          .map(([name, performance]) => ({
            name,
            dutyDays: round(performance.dutyDays),
            parkHrCount: round(performance.parkHrCount),
          })),
        videoPerformance: NEW_MEDIA_TEAM.map((name) => {
          const performance = videoPerformanceMap.get(name)!;
          return {
            name,
            pieces: round(performance.pieces),
            videoPoints: round(performance.videoPoints),
            dutyDays: round(performance.dutyDays),
          };
        }),
        newMediaNames: [...NEW_MEDIA_TEAM],
      });
      setToast(`已导出 ${selectedYear}年${selectedMonth}月绩效`);
    } catch (error) {
      window.alert(`导出当月绩效失败：${error instanceof Error ? error.message : "模板无法读取"}`);
    }
  }

  async function handleExcelImport(file: File, mode: ImportMode) {
    if (!data) return;
    if (!adminAccessActive()) {
      const inputRef = mode === "day" ? dayFileInputRef : monthFileInputRef;
      if (inputRef.current) inputRef.current.value = "";
      requireAdmin(() => inputRef.current?.click());
      return;
    }
    setIsImporting(true);
    try {
      const result = await parseExcelFile(file);
      if (!result.days.length && !result.scoreCount) {
        throw new Error("没有识别到串单日期或工分统计表");
      }
      const currentMonth = monthKey(selectedYear, selectedMonth);
      const importedForCurrentMonth = result.days.filter((day) => day.date.startsWith(currentMonth));
      const sourceDay = result.days.find((day) => day.date === selectedDate) ?? importedForCurrentMonth[0] ?? result.days[0];
      if (mode === "day" && !sourceDay) throw new Error("Excel 中没有识别到可导入的日串单");
      if (mode === "month" && !importedForCurrentMonth.length) {
        throw new Error(`Excel 中没有识别到 ${selectedYear}年${selectedMonth}月 的串单`);
      }
      const detail = mode === "day"
        ? `将 Excel 中识别到的 ${formatDate(sourceDay.date)} 串单导入该日期`
        : `将用 Excel 中的 ${importedForCurrentMonth.length} 天数据覆盖 ${selectedYear}年${selectedMonth}月整月串单`;
      if (!window.confirm(`${detail}。\n\n此操作会替换对应范围内已有串单内容，Excel 原表工分不会用于排名。是否继续？`)) {
        return;
      }

      const replacementDays: Day[] = mode === "day"
        ? [normalizeDay(sourceDay as LegacyDay)]
        : visibleDays.map((existingDay) => {
            const importedDay = importedForCurrentMonth.find((day) => day.date === existingDay.date);
            return importedDay ? normalizeDay(importedDay as LegacyDay) : blankDay(selectedYear, selectedMonth, dayLabel(existingDay.date));
          });
      const replacementDates = new Set(replacementDays.map((day) => day.date));
      const nextData: SheetData = {
        ...data,
        members: [...new Set([...data.members, ...result.members])],
        days: [
          ...data.days.filter((day) => !replacementDates.has(day.date)),
          ...replacementDays,
        ].sort((a, b) => a.date.localeCompare(b.date)),
        lastImport: {
          fileName: file.name,
          importedAt: new Date().toISOString(),
          months: [currentMonth],
        },
      };
      setData(nextData);
      if (mode === "day") {
        const [importedYear, importedMonth] = sourceDay.date.split("-").map(Number);
        setSelectedYear(importedYear);
        setSelectedMonth(importedMonth);
        setSelectedDate(sourceDay.date);
      }
      setView("daily");
      setToast(mode === "day" ? `已导入当前日串单：${file.name}` : `已覆盖当前月串单：${file.name}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : "文件无法读取";
      window.alert(`Excel 导入失败：${message}\n\n请确认文件包含串联单日期标题，或带有姓名的人员统计表。`);
    } finally {
      setIsImporting(false);
      if (dayFileInputRef.current) dayFileInputRef.current.value = "";
      if (monthFileInputRef.current) monthFileInputRef.current.value = "";
    }
  }

  function changeMonth(year: number, month: number) {
    let nextYear = year;
    let nextMonth = month;
    if (nextMonth < 1) {
      nextYear -= 1;
      nextMonth = 12;
    } else if (nextMonth > 12) {
      nextYear += 1;
      nextMonth = 1;
    }
    setSelectedYear(nextYear);
    setSelectedMonth(nextMonth);
    setView("daily");
    setQuery("");
    setPersonQuery("");
  }

  if (!data || !selectedDay) {
    return (
      <main className="loading-screen">
        <div className="loading-mark" />
        <p>正在载入全年串联单…</p>
      </main>
    );
  }

  const dayPoints =
    (peopleFor(selectedDay.dutyEditor, data.members).length ? taskMeta.duty_editor.points : 0) +
    (Object.keys(selectedDay.sections) as Category[]).reduce(
      (sum, category) =>
        sum +
        selectedDay.sections[category].reduce(
          (sectionSum, entry) =>
            sectionSum + entryPoints(category, entry, monthlyVideoRewards.get(entry.id) ?? 0),
          0,
        ),
      0,
    );
  const calendarOffset = (new Date(selectedYear, selectedMonth - 1, 1).getDay() + 6) % 7;

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark">串</div>
          <div>
            <h1>新媒体串单</h1>
            <p>每日编辑 · 自动计分</p>
          </div>
        </div>
        <div className="top-actions">
          <span className={`save-status ${saved ? "is-saved" : ""}`}>
            <i /> {saved ? "已自动保存" : "保存中…"}
          </span>
          <button
            className={`access-status ${isAdmin ? "is-admin" : ""}`}
            onClick={requestEditAccess}
            title={isAdmin ? "管理员模式将在验证后2小时到期" : "当前为只读模式，点击验证管理员"}
          >
            {isAdmin ? "管理员 · 可编辑" : "访客 · 只读"}
          </button>
          <button className="ghost-button" onClick={() => setShowRules(true)}>
            计分规则
          </button>
          <input
            ref={dayFileInputRef}
            className="file-input"
            type="file"
            accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) void handleExcelImport(file, "day");
            }}
          />
          <input
            ref={monthFileInputRef}
            className="file-input"
            type="file"
            accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) void handleExcelImport(file, "month");
            }}
          />
          <button
            className="ghost-button import-button"
            onClick={() => requireAdmin(() => dayFileInputRef.current?.click())}
            disabled={isImporting}
          >
            {isImporting ? "读取中…" : "导入日串单"}
          </button>
          <button
            className="ghost-button import-button"
            onClick={() => requireAdmin(() => monthFileInputRef.current?.click())}
            disabled={isImporting}
          >
            {isImporting ? "读取中…" : "导入月串单"}
          </button>
          <button className="primary-button" onClick={() => void exportDailyRundown()}>
            导出日串单
          </button>
        </div>
      </header>

      <div className="workspace">
        <aside className="sidebar">
          <div className="month-heading">
            <button aria-label="上一个月" title="上一个月" onClick={() => changeMonth(selectedYear, selectedMonth - 1)}>
              ‹
            </button>
            <div className="year-month-display">
              <label>
                <input
                  className="year-input"
                  type="number"
                  min="1900"
                  max="2199"
                  value={selectedYear}
                  onChange={(event) => {
                    const year = Number(event.target.value);
                    if (year >= 1900 && year <= 2199) setSelectedYear(year);
                  }}
                  aria-label="年份"
                />
                <span>年</span>
              </label>
              <strong>{String(selectedMonth).padStart(2, "0")} 月</strong>
            </div>
            <button aria-label="下一个月" title="下一个月" onClick={() => changeMonth(selectedYear, selectedMonth + 1)}>
              ›
            </button>
            <button className="reset-data" aria-label="恢复原始数据" title="恢复原始数据" onClick={resetData}>
              ↺
            </button>
          </div>
          <nav className="view-switch" aria-label="视图选择">
            <button
              className={view === "daily" ? "active" : ""}
              onClick={() => setView("daily")}
            >
              日串单
            </button>
            <button
              className={view === "summary" ? "active" : ""}
              onClick={() => setView("summary")}
            >
              工分汇总
            </button>
          </nav>
          <div className="month-grid" aria-label="月份选择">
            {Array.from({ length: 12 }, (_, index) => index + 1).map((month) => (
              <button
                key={month}
                className={selectedMonth === month ? "selected" : ""}
                onClick={() => changeMonth(selectedYear, month)}
              >
                {month}月
              </button>
            ))}
          </div>
          <div className="weekday-row" aria-hidden="true">
            {["一", "二", "三", "四", "五", "六", "日"].map((weekday) => (
              <span key={weekday}>{weekday}</span>
            ))}
          </div>
          <div className="calendar-grid" aria-label="日期选择">
            {Array.from({ length: calendarOffset }, (_, index) => (
              <span className="calendar-spacer" key={`space-${index}`} />
            ))}
            {visibleDays.map((day) => {
              const count =
                day.sections.wechat.length +
                day.sections.remix.length +
                day.sections.original.length;
              return (
                <button
                  key={day.date}
                  className={selectedDate === day.date ? "selected" : ""}
                  onClick={() => {
                    setSelectedDate(day.date);
                    setView("daily");
                  }}
                  aria-label={`${formatDate(day.date)}，${count}条内容`}
                >
                  <span>{dayLabel(day.date)}</span>
                  <i className={dayHasContent(day) ? "has-data" : ""} />
                </button>
              );
            })}
          </div>
          <div className="sidebar-note">
            <span>全年串单日历</span>
            <strong>{selectedYear} 年 · 按月独立计分</strong>
            <small>2026 年 7 月已导入原表数据</small>
          </div>
        </aside>

        <section className="content">
          {data.lastImport?.months.includes(monthKey(selectedYear, selectedMonth)) && (
            <div className="import-banner">
              <span>Excel</span>
              <div>
                <strong>当前月份已从 {data.lastImport.fileName} 导入</strong>
                <small>仅导入串单与人员数据；最终排名始终以串单计算工分为准</small>
              </div>
            </div>
          )}
          {view === "daily" ? (
            <>
              <div className="page-heading">
                <div>
                  <span className="eyebrow">DAILY RUNDOWN</span>
                  <h2>
                    {formatDate(selectedDay.date)}
                    <small>{selectedDay.weekday}</small>
                  </h2>
                  <p>共 {selectedDay.sections.wechat.length + selectedDay.sections.remix.length + selectedDay.sections.original.length} 条内容，合计 {formatPoints(dayPoints)} 工分</p>
                </div>
                <label className="search-box">
                  <span>⌕</span>
                  <input
                    value={query}
                    onChange={(event) => setQuery(event.target.value)}
                    placeholder="搜索标题或人员"
                  />
                </label>
              </div>

              <div className="duty-card">
                <label className="candidate-field">
                  <span>值班责编</span>
                  <button
                    type="button"
                    className="candidate-trigger"
                    onClick={() => toggleCandidatePicker("duty")}
                    aria-haspopup="listbox"
                    aria-expanded={openCandidatePicker === "duty"}
                  >
                    <span>{selectedDay.dutyEditor || "点击选择值班责编"}</span><i>⌄</i>
                  </button>
                  {openCandidatePicker === "duty" && (
                    <div className="candidate-menu" role="listbox" aria-label="值班责编候选人员">
                      {DUTY_EDITOR_CANDIDATES.map((name) => (
                        <button key={name} type="button" role="option" aria-selected={selectedDay.dutyEditor === name} onClick={() => chooseCandidate("dutyEditor", name)}>{name}</button>
                      ))}
                    </div>
                  )}
                  <input
                    className="candidate-custom-input"
                    readOnly={!isAdmin}
                    onClick={requestEditAccess}
                    value={selectedDay.dutyEditor ?? ""}
                    onChange={(event) => updateDay({ dutyEditor: event.target.value })}
                    placeholder="也可手动输入姓名"
                    aria-label="值班责编姓名，每天自动计1B"
                  />
                </label>
                <label>
                  <span>值班主任</span>
                  <input
                    readOnly={!isAdmin}
                    onClick={requestEditAccess}
                    value={selectedDay.dutyDirector ?? ""}
                    onChange={(event) => updateDay({ dutyDirector: event.target.value })}
                    placeholder="值班主任姓名"
                    aria-label="值班主任姓名"
                  />
                </label>
                <label>
                  <span>监审</span>
                  <input
                    readOnly={!isAdmin}
                    onClick={requestEditAccess}
                    value={selectedDay.supervisor ?? ""}
                    onChange={(event) => updateDay({ supervisor: event.target.value })}
                    placeholder="监审姓名"
                    aria-label="监审姓名"
                  />
                </label>
              </div>

              {(Object.keys(categoryMeta) as Category[]).map((category) => {
                const meta = categoryMeta[category];
                const entries = selectedDay.sections[category].filter((entry) => {
                  const creditedEditor = entry.taskType === "park_hr_micro" ? selectedDay.wechatEditor : "";
                  const haystack = `${entry.title} ${entry.staff} ${creditedEditor}`.toLowerCase();
                  return haystack.includes(query.trim().toLowerCase());
                });
                const sectionPoints = selectedDay.sections[category].reduce(
                  (sum, entry) =>
                    sum + entryPoints(category, entry, monthlyVideoRewards.get(entry.id) ?? 0),
                  0,
                );
                return (
                  <article className={`section-card ${meta.className}`} key={category}>
                    <div className="section-header">
                      <div className="section-title">
                        <span className="section-icon">{category === "wechat" ? "微" : category === "remix" ? "创" : "摄"}</span>
                        <div>
                          <h3>{meta.label}</h3>
                          <p>{meta.description}</p>
                        </div>
                      </div>
                      {category === "wechat" && (
                        <label className="wechat-editor-field candidate-field">
                          <span>微信公众号值班小编</span>
                          <button
                            type="button"
                            className="candidate-trigger"
                            onClick={() => toggleCandidatePicker("wechat")}
                            aria-haspopup="listbox"
                            aria-expanded={openCandidatePicker === "wechat"}
                          >
                            <span>{selectedDay.wechatEditor || "点击选择小编"}</span><i>⌄</i>
                          </button>
                          {openCandidatePicker === "wechat" && (
                            <div className="candidate-menu" role="listbox" aria-label="微信公众号值班小编候选人员">
                              {WECHAT_EDITOR_CANDIDATES.map((name) => (
                                <button key={name} type="button" role="option" aria-selected={selectedDay.wechatEditor === name} onClick={() => chooseCandidate("wechatEditor", name)}>{name}</button>
                              ))}
                            </div>
                          )}
                          <input
                            className="candidate-custom-input"
                            readOnly={!isAdmin}
                            onClick={requestEditAccess}
                            value={selectedDay.wechatEditor ?? ""}
                            onChange={(event) => updateDay({ wechatEditor: event.target.value })}
                            placeholder="也可手动输入"
                            aria-label="微信公众号值班小编姓名，每周绩效1200元"
                          />
                        </label>
                      )}
                      <div className="section-total">
                        <span>{selectedDay.sections[category].length} 条</span>
                        <strong>{formatPoints(sectionPoints)} <small>分</small></strong>
                      </div>
                    </div>
                    <div className="table-wrap">
                      <table>
                        <thead>
                          <tr>
                            <th className="number-col">序号</th>
                            <th className="title-col">节目标题</th>
                            <th className="staff-col">{category === "remix" ? "编辑" : "记者"}</th>
                            <th className="views-col">{category === "wechat" ? "阅读量" : "播放量"}</th>
                            <th className="duration-col">{category === "wechat" ? "刚刚帖" : "时长"}</th>
                            <th className="points-col">工分</th>
                            <th className="notes-col">备注</th>
                            <th className="action-col" />
                          </tr>
                        </thead>
                        <tbody>
                          {entries.map((entry) => {
                            const originalIndex = selectedDay.sections[category].findIndex(
                              (item) => item.id === entry.id,
                            );
                            const people = entryPeople(entry, selectedDay, data.members);
                            const rewardPoints = monthlyVideoRewards.get(entry.id) ?? 0;
                            const points = entryPoints(category, entry, rewardPoints);
                            return (
                              <tr key={entry.id}>
                                <td className="row-number">{originalIndex + 1}</td>
                                <td>
                                  <textarea
                                    readOnly={!isAdmin}
                                    onClick={requestEditAccess}
                                    className="title-input"
                                    rows={2}
                                    value={entry.title}
                                    onChange={(event) => updateEntry(category, entry.id, { title: event.target.value })}
                                    placeholder="输入节目标题"
                                    aria-label={`${meta.label}第${originalIndex + 1}条标题`}
                                  />
                                </td>
                                <td>
                                  <input
                                    readOnly={!isAdmin}
                                    onClick={requestEditAccess}
                                    value={entry.staff}
                                    onChange={(event) => updateEntry(category, entry.id, { staff: event.target.value })}
                                    placeholder="姓名"
                                    aria-label={`${meta.label}第${originalIndex + 1}条人员`}
                                  />
                                </td>
                                <td>
                                  <input
                                    readOnly={!isAdmin}
                                    onClick={requestEditAccess}
                                    value={entry.views}
                                    onChange={(event) => updateEntry(category, entry.id, { views: event.target.value })}
                                    placeholder="0"
                                    aria-label={`${meta.label}第${originalIndex + 1}条${category === "wechat" ? "阅读量" : "播放量"}`}
                                  />
                                </td>
                                <td>
                                  <input
                                    readOnly={!isAdmin}
                                    onClick={requestEditAccess}
                                    value={entry.duration}
                                    onChange={(event) => updateEntry(category, entry.id, { duration: event.target.value })}
                                    placeholder={category === "wechat" ? "—" : "00:00"}
                                    aria-label={`${meta.label}第${originalIndex + 1}条时长`}
                                  />
                                </td>
                                <td>
                                  <div className="points-input" onClick={!isAdmin ? requestEditAccess : undefined}>
                                    <select
                                      disabled={!isAdmin}
                                      onClick={requestEditAccess}
                                      value={entry.taskType ?? defaultTaskType(category)}
                                      onChange={(event) =>
                                        updateEntry(category, entry.id, {
                                          taskType: event.target.value as TaskType,
                                          manualPoints: null,
                                        })
                                      }
                                      aria-label={`${meta.label}第${originalIndex + 1}条任务类型`}
                                    >
                                      {entryTaskTypes.map((taskType) => (
                                        <option key={taskType} value={taskType}>
                                          {taskMeta[taskType].label} · {taskMeta[taskType].unit}
                                        </option>
                                      ))}
                                    </select>
                                    <input
                                      readOnly={!isAdmin}
                                      onClick={requestEditAccess}
                                      type="number"
                                      step="0.1"
                                      value={formatPoints(points)}
                                      onChange={(event) =>
                                        updateEntry(category, entry.id, {
                                          manualPoints:
                                            event.target.value === "" ? null : Number(event.target.value),
                                        })
                                      }
                                      aria-label={`${meta.label}第${originalIndex + 1}条工分`}
                                    />
                                    {entry.manualPoints === null ? (
                                      <small>
                                        自动 · {taskMeta[entry.taskType ?? defaultTaskType(category)].unit}
                                        {rewardPoints > 0 ? ` + 播放奖励${unitForPoints(rewardPoints)}` : ""}
                                        {entry.taskType === "park_hr_micro"
                                          ? selectedDay.wechatEditor
                                            ? ` · 归属${selectedDay.wechatEditor}`
                                            : " · 请先填写值班小编"
                                          : ""}
                                      </small>
                                    ) : (
                                      <button
                                        title="恢复自动计分"
                                        onClick={() => requireAdmin(() => updateEntry(category, entry.id, { manualPoints: null }))}
                                      >
                                        手动 ↺
                                      </button>
                                    )}
                                    {people.length > 1 && (
                                      <em>{formatPoints(points / people.length)}/人</em>
                                    )}
                                  </div>
                                </td>
                                <td>
                                  <input
                                    readOnly={!isAdmin}
                                    onClick={requestEditAccess}
                                    value={entry.notes}
                                    onChange={(event) => updateEntry(category, entry.id, { notes: event.target.value })}
                                    placeholder="选填"
                                    aria-label={`${meta.label}第${originalIndex + 1}条备注`}
                                  />
                                </td>
                                <td>
                                  <button
                                    className="delete-button"
                                    onClick={() => requireAdmin(() => removeEntry(category, entry.id))}
                                    aria-label={`删除${entry.title || "此条内容"}`}
                                    title="删除"
                                  >
                                    ×
                                  </button>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                    {entries.length === 0 && query && (
                      <div className="empty-search">本栏没有匹配“{query}”的内容</div>
                    )}
                    <div className="add-row-actions">
                      <button className="add-row" onClick={() => requireAdmin(() => addEntry(category))}>
                        <span>＋</span> 添加一条{meta.short}内容
                      </button>
                      {category === "wechat" && (
                        <button className="add-row quick-performance" onClick={() => requireAdmin(() => addEntry("wechat", "park_hr_micro"))}>
                          <span>＋</span> 工业园区/人社局微刊 · 3D/天
                        </button>
                      )}
                    </div>
                  </article>
                );
              })}
            </>
          ) : (
            <section className="summary-view">
              <div className="page-heading">
                <div>
                  <span className="eyebrow">MONTHLY SCOREBOARD</span>
                  <h2>{selectedYear} 年 {selectedMonth} 月工分汇总</h2>
                  <p>仅统计当前月份；多人参与项目按人数平均分配，个人明细按工分从高到低排名</p>
                </div>
              </div>
              <div className="conversion-card">
                <div className="conversion-formula">
                  <span>统一换算公式</span>
                  <strong>人民币 = 工分 × 25 元</strong>
                </div>
                <div className="conversion-units">
                  <span><b>1D</b> = 2分 = ¥50</span>
                  <span><b>1C</b> = 5分 = ¥125</span>
                  <span><b>1B</b> = 8分 = ¥200</span>
                  <span><b>1A</b> = 12分 = ¥300</span>
                </div>
                <div className="duty-pay">
                  <span>值班责编绩效</span>
                  <strong>1B / 天 · ¥200</strong>
                  <small>公众号值班小编 ¥1,200 / 周 = 48分/周，按天折算至0.1分</small>
                </div>
              </div>
              <div className="stat-grid">
                <div className="stat-card highlight">
                  <span>串单计算总工分</span>
                  <strong>{formatPoints(totalPoints)}</strong>
                  <small>折合 {formatMoney(totalPoints)} · 最终排名依据</small>
                </div>
                <div className="stat-card">
                  <span>串单内容</span>
                  <strong>{totalEntries}</strong>
                  <small>条节目</small>
                </div>
                <div className="stat-card">
                  <span>汇总人员</span>
                  <strong>{NEW_MEDIA_TEAM.length + NEWS_CENTER_TEAM.length}</strong>
                  <small>15名新媒体 + 45名新闻中心</small>
                </div>
                <div className="stat-card">
                  <span>当月天数</span>
                  <strong>{visibleDays.length}</strong>
                  <small>天</small>
                </div>
              </div>
              <div className="score-toolbar">
                <div>
                  <h3>个人工分明细</h3>
                  <p>新媒体人员按工分排名；新闻中心按指定名单顺序汇总</p>
                </div>
                <div className="score-actions">
                  <label className="person-search">
                    <span>⌕</span>
                    <input
                      value={personQuery}
                      onChange={(event) => setPersonQuery(event.target.value)}
                      placeholder="检索人名"
                      aria-label="检索人员姓名"
                    />
                    {personQuery && <button onClick={() => setPersonQuery("")} aria-label="清除检索">×</button>}
                  </label>
                  <button className="ghost-button" onClick={exportSummaryCsv}>导出汇总</button>
                  <button className="primary-button total-export-button" onClick={() => void exportTotalWorkbook()}>导出当月绩效</button>
                </div>
              </div>
              <div className="performance-groups">
                <div className="score-card performance-card video-ranking">
                  <div className="score-card-heading">
                    <div>
                      <span className="performance-label">TOP 10 VIDEOS</span>
                      <h3>短视频播放量前十</h3>
                      <p>仅统计填写记者/编辑姓名的二创与原创 · 按播放量从高到低自动排名</p>
                    </div>
                    <div className="group-total">
                      <span>{monthlyTopVideos.length} 条入榜</span>
                      <strong>{formatPoints(monthlyTopVideos.reduce((sum, video) => sum + video.rewardPoints, 0))} 分奖励</strong>
                    </div>
                  </div>
                  <div className="score-table-wrap">
                    <table className="score-table video-ranking-table">
                      <thead>
                        <tr>
                          <th>排名</th>
                          <th>日期</th>
                          <th>类别</th>
                          <th>节目标题</th>
                          <th>记者/编辑</th>
                          <th>播放量</th>
                          <th>奖励工分</th>
                        </tr>
                      </thead>
                      <tbody>
                        {monthlyTopVideos.map((video, index) => (
                          <tr key={`${video.date}-${video.entry.id}`}>
                            <td><span className={`rank rank-${index + 1}`}>{index + 1}</span></td>
                            <td>{formatDate(video.date)}</td>
                            <td><span className={`category-pill ${video.category}`}>{categoryMeta[video.category].short}</span></td>
                            <td><strong>{video.entry.title || "未填写标题"}</strong></td>
                            <td><strong>{video.people.join("、")}</strong></td>
                            <td><strong className="video-views">{new Intl.NumberFormat("zh-CN").format(video.views)}</strong></td>
                            <td><strong className="total-score">{unitForPoints(video.rewardPoints)} · {formatPoints(video.rewardPoints)}分</strong></td>
                          </tr>
                        ))}
                        {!monthlyTopVideos.length && (
                          <tr><td className="no-person-result" colSpan={7}>当月暂无填写姓名和播放量的二创或原创短视频</td></tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
                {filteredScoreGroups.map((group) => {
                  const originalGroup = scoreGroups.find((item) => item.key === group.key) ?? group;
                  const groupPoints = originalGroup.people.reduce((sum, item) => sum + item.person.total, 0);
                  return (
                    <div className={`score-card performance-card ${group.key}`} key={group.key}>
                      <div className="score-card-heading">
                        <div>
                          <span className="performance-label">{group.key === "new-media" ? "NEW MEDIA" : "NEWS CENTER"}</span>
                          <h3>{group.title}</h3>
                          <p>{group.description}</p>
                        </div>
                        <div className="group-total">
                          <span>{originalGroup.people.length} 人 · {formatPoints(groupPoints)} 分</span>
                          <strong>{formatMoney(groupPoints)}</strong>
                        </div>
                      </div>
                      <div className="score-table-wrap">
                        <table className="score-table">
                          <thead>
                            <tr>
                              <th>排名</th>
                              <th>姓名</th>
                              <th>值班责编</th>
                              <th>微信</th>
                              <th>二创</th>
                              <th>原创</th>
                              <th>固定月绩效</th>
                              <th>参与条数</th>
                              <th>串单计算</th>
                              <th>折合人民币</th>
                            </tr>
                          </thead>
                          <tbody>
                            {group.people.map(({ person, rank }) => (
                              <tr
                                key={person.name}
                                className={focusedPerson === person.name ? "person-selected" : "person-row"}
                                onClick={() => setPersonQuery(person.name)}
                                onKeyDown={(event) => {
                                  if (event.key === "Enter" || event.key === " ") setPersonQuery(person.name);
                                }}
                                tabIndex={0}
                              >
                                <td>
                                  <span className={`rank rank-${rank}`}>
                                    {rank === 1 && <b className="crown" aria-label="本月工分冠军">♛</b>}
                                    {rank}
                                  </span>
                                </td>
                                <td><strong>{person.name}</strong></td>
                                <td>{formatPoints(person.duty)}</td>
                                <td>{formatPoints(person.wechat)}</td>
                                <td>{formatPoints(person.remix)}</td>
                                <td>{formatPoints(person.original)}</td>
                                <td>{formatPoints(person.fixed)}</td>
                                <td>{person.pieces}</td>
                                <td><strong className="total-score">{formatPoints(person.total)}</strong></td>
                                <td><strong className="money-score">{formatMoney(person.total)}</strong></td>
                              </tr>
                            ))}
                            {!group.people.length && (
                              <tr>
                                <td className="no-person-result" colSpan={10}>
                                  {personQuery ? `此表没有找到“${personQuery}”` : "本月暂无串单工分"}
                                </td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  );
                })}
                <div className="score-card performance-card news-center">
                  <div className="score-card-heading">
                    <div>
                      <span className="performance-label">NEWS CENTER</span>
                      <h3>新闻中心绩效</h3>
                      <p>固定名单顺序 · 微信转发每人默认40条/1.6分</p>
                    </div>
                    <div className="group-total">
                      <span>{NEWS_CENTER_TEAM.length} 人 · {formatPoints(newsCenterScores.reduce((sum, person) => sum + person.total, 0))} 分</span>
                      <strong>{formatMoney(newsCenterScores.reduce((sum, person) => sum + person.total, 0))}</strong>
                    </div>
                  </div>
                  <div className="score-table-wrap">
                    <table className="score-table news-center-table">
                      <thead>
                        <tr>
                          <th>序号</th>
                          <th>姓名</th>
                          <th>微信转发</th>
                          <th>微刊</th>
                          <th>短视频</th>
                          <th>刚刚帖</th>
                          <th>合计</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredNewsCenterScores.map((person) => (
                          <tr
                            key={person.name}
                            className={focusedPerson === person.name ? "person-selected" : "person-row"}
                            onClick={() => setPersonQuery(person.name)}
                            onKeyDown={(event) => {
                              if (event.key === "Enter" || event.key === " ") setPersonQuery(person.name);
                            }}
                            tabIndex={0}
                          >
                            <td><span className="fixed-order">{person.order}</span></td>
                            <td>
                              <strong>{person.name}</strong>
                              {NEW_MEDIA_TEAM_SET.has(person.name) && <small className="new-media-note">仅保留转发分</small>}
                            </td>
                            <td><strong>{person.forwardCount}条 · {formatPoints(person.forward)}分</strong></td>
                            <td>{formatPoints(person.micro)}</td>
                            <td>{formatPoints(person.video)}</td>
                            <td>{formatPoints(person.justNow)}</td>
                            <td><strong className="total-score">{formatPoints(person.total)}</strong></td>
                          </tr>
                        ))}
                        {!filteredNewsCenterScores.length && (
                          <tr><td className="no-person-result" colSpan={7}>此表没有找到“{personQuery}”</td></tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
              {focusedPerson && (
                <div className="person-detail-card">
                  <div className="person-detail-heading">
                    <div>
                      <span className="eyebrow">PERSONAL TASKS</span>
                      <h3>{focusedPerson} · 当月任务与工分</h3>
                      <p>{personTasks.length} 项任务，共 {formatPoints(personTasks.reduce((sum, task) => sum + task.personalPoints, 0))} 工分</p>
                    </div>
                    <strong>{formatMoney(personTasks.reduce((sum, task) => sum + task.personalPoints, 0))}</strong>
                  </div>
                  {personTasks.length ? (
                    <div className="person-task-wrap">
                      <table className="person-task-table">
                        <thead>
                          <tr>
                            <th>日期</th>
                            <th>分类</th>
                            <th>任务标准</th>
                            <th>串单任务</th>
                            <th>项目工分</th>
                            <th>个人工分</th>
                            <th>人民币</th>
                          </tr>
                        </thead>
                        <tbody>
                          {personTasks.map((task) => {
                            const taskType =
                              task.entry.taskType ??
                              (task.category === "duty" ? "duty_editor" : task.category === "bonus" ? "video_replay_upload" : defaultTaskType(task.category));
                            return (
                              <tr key={`${task.date}-${task.entry.id}`}>
                                <td>{Number(task.date.slice(-2))}日</td>
                                <td>
                                  <span className={`category-pill ${task.category}`}>
                                    {task.category === "duty" ? "值班" : task.category === "bonus" ? "固定" : categoryMeta[task.category].short}
                                  </span>
                                </td>
                                <td>
                                  <strong>{taskMeta[taskType].label}</strong>
                                  <small>{taskMeta[taskType].unit}</small>
                                </td>
                                <td>
                                  <strong>{task.entry.title || "未填写标题"}</strong>
                                  {task.rewardPoints > 0 && <small className="reward-note">月度播放量前十奖励 +{unitForPoints(task.rewardPoints)}</small>}
                                  {task.entry.notes && <small>{task.entry.notes}</small>}
                                </td>
                                <td>{formatPoints(task.totalPoints)}</td>
                                <td><strong className="total-score">{formatPoints(task.personalPoints)}</strong></td>
                                <td><strong className="money-score">{formatMoney(task.personalPoints)}</strong></td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <div className="no-person-tasks">当前月份串单中未识别到该人员的任务。</div>
                  )}
                </div>
              )}
            </section>
          )}
        </section>
      </div>

      {showRules && (
        <div className="modal-backdrop" role="presentation" onMouseDown={() => setShowRules(false)}>
          <section className="rules-modal" role="dialog" aria-modal="true" aria-labelledby="rules-title" onMouseDown={(event) => event.stopPropagation()}>
            <button className="modal-close" onClick={() => setShowRules(false)} aria-label="关闭">×</button>
            <span className="eyebrow">SCORING RULES</span>
            <h2 id="rules-title">新媒体工分标准</h2>
            <p className="rules-intro">选择串单中的“任务类型”后自动计分；特殊情况可直接修改工分。多人参与时按人数平均分配。</p>
            <div className="modal-conversion">
              <strong>人民币 = 工分 × 25 元</strong>
              <span>D=2分/¥50 · C=5分/¥125 · B=8分/¥200 · A=12分/¥300</span>
            </div>
            <ol className="standards-list">
              <li><strong>新武夷微信公众号值班</strong><span>¥1,200 / 周 = 48工分/周；跨月或不足一周按实际值班天数折算，四舍五入至0.1分</span></li>
              <li><strong>当日值班责编</strong><span>1B / 天（8分，¥200），填写姓名后自动计入</span></li>
              <li><strong>短视频编辑</strong><span>1D / 条（2分，¥50）</span></li>
              <li><strong>新闻活动类外采短视频</strong><span>1C / 条（5分，¥125）</span></li>
              <li><strong>剧情类短视频</strong><span>A+B 起评（20分，¥500）</span></li>
              <li><strong>短视频播放量奖励</strong><span>每月自动统计带记者/编辑姓名的二创、原创；播放量前十奖励1B起评，超10万1A、超100万2A、超1000万3A</span></li>
              <li><strong>微信公众号代运营编辑</strong><span>3D / 组（6分，¥150）</span></li>
              <li><strong>每月微博发布</strong><span>0.5D / 条（1分，¥25）</span></li>
              <li><strong>视频连线、直播拉流</strong><span>1D / 场（2分，¥50）</span></li>
              <li><strong>节目整档上传网络平台</strong><span>3A / 月（36分，¥900）</span></li>
              <li><strong>微刊阅读量奖励</strong><span>超1万1C；超10万1A；超100万2A</span></li>
              <li><strong>服务共建类微刊</strong><span>3D / 组（6分，¥150）</span></li>
              <li><strong>工业园区/人社局微刊</strong><span>3D / 天（6分，¥150），网页手动加入后归属当日公众号值班小编</span></li>
              <li><strong>策划类纯AI短视频</strong><span>1B / 条（8分，¥200）</span></li>
            </ol>
            <button className="primary-button modal-action" onClick={() => setShowRules(false)}>知道了</button>
          </section>
        </div>
      )}

      {showAdminLogin && (
        <div className="modal-backdrop admin-backdrop" role="presentation" onMouseDown={() => {
          pendingAdminAction.current = null;
          setShowAdminLogin(false);
        }}>
          <form className="admin-modal" role="dialog" aria-modal="true" aria-labelledby="admin-title" onSubmit={submitAdminPassword} onMouseDown={(event) => event.stopPropagation()}>
            <button
              className="modal-close"
              type="button"
              onClick={() => {
                pendingAdminAction.current = null;
                setShowAdminLogin(false);
              }}
              aria-label="关闭"
            >×</button>
            <span className="admin-lock" aria-hidden="true">锁</span>
            <span className="eyebrow">ADMIN ACCESS</span>
            <h2 id="admin-title">管理员验证</h2>
            <p>当前为只读模式。验证后可添加、修改或删除串单内容，权限在本浏览器保留2小时。</p>
            <label className="admin-password-field">
              <span>管理密码</span>
              <input
                type="password"
                value={adminPassword}
                onChange={(event) => {
                  setAdminPassword(event.target.value);
                  setAdminError("");
                }}
                placeholder="请输入管理密码"
                autoFocus
                autoComplete="current-password"
              />
            </label>
            {adminError && <div className="admin-error" role="alert">{adminError}</div>}
            <button className="primary-button admin-submit" type="submit">验证并进入编辑模式</button>
          </form>
        </div>
      )}

      <aside className={`music-player ${isMusicHidden ? "is-hidden" : ""}`} aria-label="背景音乐播放器">
        <audio
          ref={audioRef}
          src={`${import.meta.env.BASE_URL}music.mp3`}
          autoPlay
          loop
          preload="auto"
          onPlay={() => setIsMusicPlaying(true)}
          onPause={() => setIsMusicPlaying(false)}
        />
        <button
          className="music-collapse"
          type="button"
          onClick={() => setIsMusicHidden((hidden) => !hidden)}
          aria-label={isMusicHidden ? "显示音乐播放器" : "隐藏音乐播放器"}
          title={isMusicHidden ? "显示播放器" : "隐藏播放器"}
        >
          {isMusicHidden ? "♫" : "⌄"}
        </button>
        <button
          className={`music-toggle ${isMusicPlaying ? "is-playing" : ""}`}
          type="button"
          onClick={toggleMusic}
          aria-label={isMusicPlaying ? "暂停音乐" : "播放音乐"}
          title={isMusicPlaying ? "暂停" : "播放"}
        >
          {isMusicPlaying ? "Ⅱ" : "▶"}
        </button>
        <div className="music-cover" aria-hidden="true">念</div>
        <div className="music-info">
          <span>{isMusicPlaying ? "正在播放" : "已暂停"}</span>
          <strong>念张师</strong>
          <small>尚春 · 张雪峰老师 我还记得你</small>
        </div>
        <div className={`music-bars ${isMusicPlaying ? "is-playing" : ""}`} aria-hidden="true">
          <i /><i /><i />
        </div>
      </aside>

      {toast && <div className="toast">{toast}</div>}
    </main>
  );
}
