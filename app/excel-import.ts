export type ImportedEntry = {
  id: string;
  title: string;
  staff: string;
  views: string;
  duration: string;
  manualPoints: number | null;
  sourcePoints: number | null;
  notes: string;
};

export type ImportedDay = {
  date: string;
  weekday: string;
  wechatEditor: string;
  videoEditor: string;
  dutyEditor: string;
  dutyDirector: string;
  supervisor: string;
  sections: {
    wechat: ImportedEntry[];
    remix: ImportedEntry[];
    original: ImportedEntry[];
  };
};

export type ImportedScore = {
  name: string;
  total: number;
  wechat: number | null;
  video: number | null;
};

export type ExcelImportResult = {
  days: ImportedDay[];
  scoresByMonth: Record<string, ImportedScore[]>;
  members: string[];
  entryCount: number;
  scoreCount: number;
  months: string[];
};

type Category = "wechat" | "remix" | "original";
type Row = unknown[];

function text(value: unknown) {
  if (value === null || value === undefined) return "";
  return String(value).replace(/\u00a0/g, " ").trim();
}

function numeric(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const raw = text(value).replace(/,/g, "");
  if (!raw) return null;
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : null;
}

function dateKey(year: number, month: number, day: number) {
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function monthKey(year: number, month: number) {
  return `${year}-${String(month).padStart(2, "0")}`;
}

function parseHeaderDate(value: unknown) {
  const raw = text(value);
  if (!raw || !/(串|片单|联单)/.test(raw)) return null;
  const numbers = raw.match(/\d+/g)?.map(Number) ?? [];
  const yearIndex = numbers.findIndex((part) => part >= 1900 && part <= 2199);
  if (yearIndex < 0 || numbers.length < yearIndex + 3) return null;
  const [year, month, day] = numbers.slice(yearIndex, yearIndex + 3);
  if (month < 1 || month > 12 || day < 1 || day > new Date(year, month, 0).getDate()) return null;
  return { year, month, day };
}

function weekdayFor(date: string) {
  const [year, month, day] = date.split("-").map(Number);
  return ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"][
    new Date(year, month - 1, day).getDay()
  ];
}

function blankImportedDay(date: string): ImportedDay {
  return {
    date,
    weekday: weekdayFor(date),
    wechatEditor: "",
    videoEditor: "",
    dutyEditor: "",
    dutyDirector: "",
    supervisor: "",
    sections: { wechat: [], remix: [], original: [] },
  };
}

function dutyName(value: string, label: "值班责编" | "值班主任" | "监审") {
  const normalized = value.replace(/[\r\n]+/g, " ").trim();
  const nextLabels = label === "值班责编" ? "值班主任|监审" : label === "值班主任" ? "监审|值班责编" : "值班责编|值班主任";
  const match = normalized.match(
    new RegExp(`${label}\\s*[：:]?\\s*(.*?)(?=\\s*(?:${nextLabels})\\s*[：:]|$)`),
  );
  return match?.[1]?.replace(/^[，,；;|\s]+|[，,；;|\s]+$/g, "").trim() ?? "";
}

function findEditor(label: string) {
  return label.match(/小编[：:]\s*(.*)$/)?.[1]?.trim() ?? "";
}

function parseRundownRows(rows: Row[], sourceTag: string) {
  const dayMap = new Map<string, ImportedDay>();
  let current: ImportedDay | null = null;
  let category: Category = "wechat";

  rows.forEach((row, rowIndex) => {
    const first = text(row[0]);
    const title = text(row[1]);
    const parsedDate = parseHeaderDate(row[0]);
    if (parsedDate) {
      const date = dateKey(parsedDate.year, parsedDate.month, parsedDate.day);
      current = blankImportedDay(date);
      dayMap.set(date, current);
      category = "wechat";
      return;
    }
    if (!current) return;

    if (first.startsWith("微信公众号")) {
      current.wechatEditor = findEditor(first);
      category = "wechat";
      return;
    }
    if (first.includes("小编二创短视频")) {
      current.videoEditor = findEditor(first);
      category = "remix";
      return;
    }
    if (first.includes("记者原创短视频")) {
      category = "original";
      return;
    }
    if (first.startsWith("序号")) return;
    if (/(值班责编|值班主任|监审)/.test(first)) {
      const dutyLine = row.map(text).filter(Boolean).join(" ");
      current.dutyEditor = dutyName(dutyLine, "值班责编") || current.dutyEditor;
      current.dutyDirector = dutyName(dutyLine, "值班主任") || current.dutyDirector;
      current.supervisor = dutyName(dutyLine, "监审") || current.supervisor;
      return;
    }
    if (!title) return;

    const fullPoints = numeric(row[5]);
    const perPersonPoints = numeric(row[7]);
    const points = fullPoints ?? perPersonPoints;
    current.sections[category].push({
      id: `${sourceTag}-${current.date}-${category}-${rowIndex}`,
      title,
      staff: text(row[2]),
      views: text(row[3]),
      duration: text(row[4]),
      manualPoints: points,
      sourcePoints: points,
      notes: text(row[6]),
    });
  });

  return [...dayMap.values()];
}

function scoreSheetMonth(sheetName: string, fallbackYear: number, fallbackMonth: number) {
  const year = Number(sheetName.match(/(19\d{2}|20\d{2}|21\d{2})/)?.[1] ?? fallbackYear);
  const month = Number(sheetName.match(/(\d{1,2})\s*月/)?.[1] ?? fallbackMonth);
  return monthKey(year, month >= 1 && month <= 12 ? month : fallbackMonth);
}

function parseScoreRows(rows: Row[]) {
  const headerIndex = rows.findIndex((row) => row.some((cell) => text(cell) === "姓名") && row.some((cell) => text(cell).includes("合计")));
  if (headerIndex < 0) return [];
  const header = rows[headerIndex].map(text);
  const nameColumn = header.findIndex((cell) => cell === "姓名");
  const totalColumn = header.findIndex((cell) => cell.includes("合计"));
  const forwardColumn = header.findIndex((cell) => cell.includes("微信转发"));
  const microColumn = header.findIndex((cell) => cell.includes("微刊"));
  const justNowColumn = header.findIndex((cell) => cell.includes("刚刚帖"));
  const videoColumn = header.findIndex((cell) => cell.includes("短视频"));
  if (nameColumn < 0 || totalColumn < 0) return [];

  const scores: ImportedScore[] = [];
  rows.slice(headerIndex + 1).forEach((row) => {
    const name = text(row[nameColumn]);
    const total = numeric(row[totalColumn]);
    if (!name || total === null || name === "姓名" || name.length > 12) return;
    const forward = forwardColumn >= 0 ? numeric(row[forwardColumn]) ?? 0 : 0;
    const micro = microColumn >= 0 ? numeric(row[microColumn]) ?? 0 : 0;
    const justNow = justNowColumn >= 0 ? numeric(row[justNowColumn]) ?? 0 : 0;
    const video = videoColumn >= 0 ? numeric(row[videoColumn]) : null;
    scores.push({
      name,
      total,
      wechat: forward / 25 + micro + justNow,
      video,
    });
  });
  return scores;
}

export async function parseExcelFile(file: File): Promise<ExcelImportResult> {
  const XLSX = await import("xlsx");
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const rowsBySheet = new Map<string, Row[]>();
  workbook.SheetNames.forEach((sheetName) => {
    rowsBySheet.set(
      sheetName,
      XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null }) as Row[],
    );
  });

  const parsedDays: ImportedDay[] = [];
  rowsBySheet.forEach((rows, sheetName) => {
    if (rows.some((row) => parseHeaderDate(row[0]))) {
      parsedDays.push(...parseRundownRows(rows, `${file.name}-${sheetName}`));
    }
  });

  const latestDays = new Map<string, ImportedDay>();
  parsedDays.forEach((day) => latestDays.set(day.date, day));
  const days = [...latestDays.values()].sort((a, b) => a.date.localeCompare(b.date));
  const firstDate = days[0]?.date?.split("-").map(Number);
  const fallbackYear = firstDate?.[0] ?? new Date().getFullYear();
  const fallbackMonth = firstDate?.[1] ?? new Date().getMonth() + 1;
  const scoresByMonth: Record<string, ImportedScore[]> = {};

  rowsBySheet.forEach((rows, sheetName) => {
    if (!/(工分|公分)/.test(sheetName)) return;
    const scores = parseScoreRows(rows);
    if (scores.length) scoresByMonth[scoreSheetMonth(sheetName, fallbackYear, fallbackMonth)] = scores;
  });

  const members = new Set<string>();
  Object.values(scoresByMonth).flat().forEach((score) => members.add(score.name));
  const months = new Set<string>();
  days.forEach((day) => months.add(day.date.slice(0, 7)));
  Object.keys(scoresByMonth).forEach((month) => months.add(month));

  return {
    days,
    scoresByMonth,
    members: [...members],
    entryCount: days.reduce(
      (sum, day) => sum + day.sections.wechat.length + day.sections.remix.length + day.sections.original.length,
      0,
    ),
    scoreCount: Object.values(scoresByMonth).reduce((sum, scores) => sum + scores.length, 0),
    months: [...months].sort(),
  };
}
