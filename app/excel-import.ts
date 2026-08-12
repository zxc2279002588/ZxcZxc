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

function pointsValue(value: unknown) {
  const direct = numeric(value);
  if (direct !== null) return direct;
  const raw = text(value).toUpperCase().replace(/\s+/g, "");
  if (!raw) return null;
  const units: Record<string, number> = { A: 12, B: 8, C: 5, D: 2 };
  let total = 0;
  let matched = false;
  for (const match of raw.matchAll(/(\d+(?:\.\d+)?)?([ABCD])/g)) {
    matched = true;
    total += Number(match[1] ?? 1) * units[match[2]];
  }
  return matched ? total : null;
}

function dateKey(year: number, month: number, day: number) {
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function monthKey(year: number, month: number) {
  return `${year}-${String(month).padStart(2, "0")}`;
}

function parseHeaderDate(value: unknown) {
  if (value instanceof Date && Number.isFinite(value.getTime())) {
    return { year: value.getFullYear(), month: value.getMonth() + 1, day: value.getDate() };
  }
  const raw = text(value);
  if (!raw || !/(串|片单|联单)/.test(raw)) return null;
  const match = raw.match(/(19\d{2}|20\d{2}|21\d{2})\s*(?:年|[./-])\s*(\d{1,2})\s*(?:月|[./-])[-\s]*(\d{1,2})\s*日?/);
  if (!match) return null;
  const [, yearText, monthText, dayText] = match;
  const [year, month, day] = [Number(yearText), Number(monthText), Number(dayText)];
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
  return label.match(/(?:微刊|主班)?小编\s*[：:]\s*(.*)$/)?.[1]?.trim() ?? "";
}

type RundownColumns = {
  title: number;
  staff: number;
  views: number;
  duration: number;
  points: number;
  notes: number;
  perPersonPoints: number;
};

const defaultColumns: RundownColumns = {
  title: 1,
  staff: 2,
  views: 3,
  duration: 4,
  points: 5,
  notes: 6,
  perPersonPoints: 7,
};

function columnsFromHeader(row: Row, previous: RundownColumns) {
  const headers = row.map(text);
  const normalizedHeaders = headers.map((cell) => cell.replace(/\s+/g, ""));
  const title = normalizedHeaders.findIndex((cell) => /节目标题|标题/.test(cell));
  if (title < 0) return null;
  const find = (pattern: RegExp, fallback: number) => {
    const index = normalizedHeaders.findIndex((cell) => pattern.test(cell));
    return index >= 0 ? index : fallback;
  };
  const points = find(/^(工分|公分)$/, previous.points);
  return {
    title,
    staff: find(/^(记者|采编|编辑|小编)$/, previous.staff),
    views: find(/阅读量/, previous.views),
    duration: find(/^(刚刚帖|时长)$/, previous.duration),
    points,
    notes: find(/备注/, previous.notes),
    perPersonPoints: normalizedHeaders.findIndex((cell, index) => index > points && /^(工分|公分|个人工分)$/.test(cell)),
  };
}

function parseRundownRows(rows: Row[], sourceTag: string) {
  const dayMap = new Map<string, ImportedDay>();
  let current: ImportedDay | null = null;
  let category: Category = "wechat";
  let columns = { ...defaultColumns };

  rows.forEach((row, rowIndex) => {
    const rowText = row.map(text).filter(Boolean).join(" ");
    const parsedDate = row.map(parseHeaderDate).find((value): value is NonNullable<ReturnType<typeof parseHeaderDate>> => Boolean(value)) ?? null;
    if (parsedDate) {
      const date = dateKey(parsedDate.year, parsedDate.month, parsedDate.day);
      current = blankImportedDay(date);
      dayMap.set(date, current);
      category = "wechat";
      columns = { ...defaultColumns };
      return;
    }
    if (!current) return;

    if (rowText.includes("微信公众号")) {
      current.wechatEditor = findEditor(rowText);
      category = "wechat";
      return;
    }
    if (rowText.includes("小编二创短视频")) {
      current.videoEditor = findEditor(rowText);
      category = "remix";
      return;
    }
    if (rowText.includes("记者原创短视频")) {
      category = "original";
      return;
    }
    const detectedColumns = columnsFromHeader(row, columns);
    if (detectedColumns) {
      columns = detectedColumns;
      return;
    }
    if (/(值班责编|值班主任|监审)/.test(rowText)) {
      current.dutyEditor = dutyName(rowText, "值班责编") || current.dutyEditor;
      current.dutyDirector = dutyName(rowText, "值班主任") || current.dutyDirector;
      current.supervisor = dutyName(rowText, "监审") || current.supervisor;
      return;
    }
    const title = text(row[columns.title]);
    if (!title) return;

    const fullPoints = pointsValue(row[columns.points]);
    const perPersonPoints = columns.perPersonPoints >= 0 ? pointsValue(row[columns.perPersonPoints]) : null;
    const points = fullPoints ?? perPersonPoints;
    current.sections[category].push({
      id: `${sourceTag}-${current.date}-${category}-${rowIndex}`,
      title,
      staff: text(row[columns.staff]),
      views: text(row[columns.views]),
      duration: text(row[columns.duration]),
      manualPoints: points,
      sourcePoints: points,
      notes: text(row[columns.notes]),
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
  days.forEach((day) => {
    [day.wechatEditor, day.videoEditor, day.dutyEditor, day.dutyDirector, day.supervisor]
      .flatMap((value) => text(value).split(/[、，,；;\/\s]+/))
      .filter((name) => name.length >= 2 && name.length <= 4)
      .forEach((name) => members.add(name));
    Object.values(day.sections).flat().forEach((entry) => {
      text(entry.staff).replace(/(AI|ai)制作[：:]?/g, "").split(/[、，,；;\/\s]+/)
        .filter((name) => name.length >= 2 && name.length <= 4)
        .forEach((name) => members.add(name));
    });
  });
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
