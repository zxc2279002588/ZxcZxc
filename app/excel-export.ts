type ExportEntry = {
  title: string;
  staff: string;
  views: string;
  duration: string;
  notes: string;
  exportPoints: number;
};

type ExportDay = {
  date: string;
  weekday: string;
  wechatEditor: string;
  dutyEditor: string;
  dutyDirector: string;
  supervisor: string;
  sections: Record<"wechat" | "remix" | "original", ExportEntry[]>;
};

type NewMediaScore = {
  name: string;
  rank: number;
  duty: number;
  wechat: number;
  remix: number;
  original: number;
  fixed: number;
  total: number;
};

type NewsCenterScore = {
  name: string;
  order: number;
  forwardCount: number;
  micro: number;
  video: number;
  justNow: number;
};

type MicroReward = { title: string; views: string; staff: string; reward: number; notes: string };

type ExportPayload = {
  year: number;
  month: number;
  days: ExportDay[];
  newMediaScores: NewMediaScore[];
  newsCenterScores: NewsCenterScore[];
  microRewards: MicroReward[];
};

function clone<T>(value: T): T {
  return value === undefined ? value : JSON.parse(JSON.stringify(value));
}

function writeCell(sheet: any, address: string, value: unknown, templateCell?: any) {
  const target = sheet[address] ?? {};
  if (templateCell?.s !== undefined) target.s = clone(templateCell.s);
  if (templateCell?.z !== undefined) target.z = templateCell.z;
  delete target.f;
  target.v = value ?? "";
  target.t = typeof value === "number" ? "n" : "s";
  sheet[address] = target;
}

function copyStyledRow(XLSX: any, sheet: any, templates: Record<number, any[]>, sourceRow: number, targetRow: number, values: unknown[]) {
  const styleCells = templates[sourceRow];
  for (let column = 0; column < 8; column += 1) {
    const address = XLSX.utils.encode_cell({ r: targetRow - 1, c: column });
    writeCell(sheet, address, values[column] ?? "", styleCells[column]);
  }
}

function buildRundownSheet(XLSX: any, sheet: any, payload: ExportPayload) {
  const templateRows: Record<number, any[]> = {};
  [1, 2, 3, 4, 9, 10, 11, 13, 14, 15, 17, 18, 19].forEach((row) => {
    templateRows[row] = Array.from({ length: 8 }, (_, column) => clone(sheet[XLSX.utils.encode_cell({ r: row - 1, c: column })] ?? {}));
  });
  Object.keys(sheet).filter((key) => /^[A-H]\d+$/.test(key)).forEach((key) => delete sheet[key]);
  sheet["!merges"] = [];
  sheet["!rows"] = [];
  let row = 1;
  const merge = (at: number) => sheet["!merges"].push({ s: { r: at - 1, c: 0 }, e: { r: at - 1, c: 6 } });
  payload.days.forEach((day) => {
    const [year, month, date] = day.date.split("-").map(Number);
    copyStyledRow(XLSX, sheet, templateRows, 1, row, [`新媒体串片单  ${year}-${month}-${date} ${day.weekday}`]); merge(row); row += 1;
    copyStyledRow(XLSX, sheet, templateRows, 2, row, [`微信公众号  本周微刊小编：${day.wechatEditor || ""}`]); merge(row); row += 1;
    copyStyledRow(XLSX, sheet, templateRows, 3, row, ["序号", " 节 目 标 题", "记者", "阅读量", "刚刚帖", "工分", "备注"]); row += 1;
    day.sections.wechat.forEach((entry, index) => {
      copyStyledRow(XLSX, sheet, templateRows, 4, row, [index + 1, entry.title, entry.staff, entry.views, entry.duration, entry.exportPoints || "", entry.notes, entry.exportPoints || ""]); row += 1;
    });
    copyStyledRow(XLSX, sheet, templateRows, 9, row, ["小编二创短视频（含制作封面、包框）"]); merge(row); row += 1;
    copyStyledRow(XLSX, sheet, templateRows, 10, row, ["序号", " 节 目 标 题", "编辑", "阅读量", "时长", "工分", "备注"]); row += 1;
    day.sections.remix.forEach((entry, index) => {
      copyStyledRow(XLSX, sheet, templateRows, 11, row, [index + 1, entry.title, entry.staff, entry.views, entry.duration, entry.exportPoints || "", entry.notes, entry.exportPoints || ""]); row += 1;
    });
    copyStyledRow(XLSX, sheet, templateRows, 13, row, ["记者原创短视频（新媒体首发）"]); merge(row); row += 1;
    copyStyledRow(XLSX, sheet, templateRows, 14, row, ["序号", " 节 目 标 题", "采编", "阅读量", "时长", "工分", "备注"]); row += 1;
    day.sections.original.forEach((entry, index) => {
      copyStyledRow(XLSX, sheet, templateRows, 15, row, [index + 1, entry.title, entry.staff, entry.views, entry.duration, entry.exportPoints || "", entry.notes, entry.exportPoints || ""]); row += 1;
    });
    copyStyledRow(XLSX, sheet, templateRows, 17, row, [`值班责编：${day.dutyEditor || ""}    值班主任：${day.dutyDirector || ""}    监审：${day.supervisor || ""}`]); merge(row); row += 1;
    copyStyledRow(XLSX, sheet, templateRows, 18, row, []); row += 1;
    copyStyledRow(XLSX, sheet, templateRows, 19, row, []); row += 1;
  });
  sheet["!ref"] = `A1:H${Math.max(1, row - 1)}`;
}

function updateNewMediaSheet(XLSX: any, sheet: any, payload: ExportPayload) {
  writeCell(sheet, "A1", `${payload.month}月新媒体人员绩效`, sheet.A1);
  writeCell(sheet, "B26", "新媒体人员合计分", sheet.B26);
  const templateName = clone(sheet.B34 ?? sheet.B27 ?? {});
  const templateTotal = clone(sheet.C34 ?? sheet.C27 ?? {});
  for (let row = 27; row <= 41; row += 1) {
    const score = payload.newMediaScores[row - 27];
    writeCell(sheet, `B${row}`, score?.name ?? "", templateName);
    writeCell(sheet, `C${row}`, score?.total ?? "", templateTotal);
  }
  const bonusRows = [
    [18, "周婷", "【新闻回看】视频回看每日上传绩效：24分；", 24],
    [20, "刘乐", "新闻节目上传：36分", 36],
    [22, "吴轲宇", "接听指令电话", 20],
  ] as const;
  bonusRows.forEach(([row, name, label, points]) => {
    writeCell(sheet, `A${row}`, name, sheet[`A${row}`]);
    writeCell(sheet, `B${row}`, label, sheet[`B${row}`]);
    writeCell(sheet, `F${row}`, points, sheet[`F${row}`]);
  });
  sheet["!ref"] = `A1:F52`;
}

function updateMicroSheet(XLSX: any, sheet: any, payload: ExportPayload) {
  writeCell(sheet, "A1", `${payload.month}月份刚刚帖及微刊阅读量奖励明细`, sheet.A1);
  const template = Array.from({ length: 6 }, (_, column) => clone(sheet[XLSX.utils.encode_cell({ r: 2, c: column })] ?? {}));
  Object.keys(sheet).filter((key) => /^[A-F]\d+$/.test(key) && Number(key.match(/\d+/)?.[0]) >= 3).forEach((key) => delete sheet[key]);
  payload.microRewards.forEach((item, index) => {
    const row = index + 3;
    const values = [index + 1, item.title, item.views, item.staff, item.reward, item.notes];
    values.forEach((value, column) => writeCell(sheet, XLSX.utils.encode_cell({ r: row - 1, c: column }), value, template[column]));
  });
  sheet["!ref"] = `A1:F${Math.max(3, payload.microRewards.length + 2)}`;
}

function updateNewsCenterSheet(XLSX: any, sheet: any, payload: ExportPayload) {
  writeCell(sheet, "A1", `${payload.month}月新媒体工分`, sheet.A1);
  const styleRows = Array.from({ length: 7 }, (_, column) => clone(sheet[XLSX.utils.encode_cell({ r: 2, c: column })] ?? {}));
  payload.newsCenterScores.forEach((person, index) => {
    const row = index + 3;
    const values = [person.order, person.name, person.forwardCount, person.micro || "", person.video || "", person.justNow || ""];
    values.forEach((value, column) => writeCell(sheet, XLSX.utils.encode_cell({ r: row - 1, c: column }), value, styleRows[column]));
    const totalCell = sheet[`G${row}`] ?? clone(styleRows[6]);
    totalCell.t = "n";
    totalCell.f = `C${row}/25+F${row}+E${row}+D${row}`;
    delete totalCell.v;
    sheet[`G${row}`] = totalCell;
  });
  const totalRow = payload.newsCenterScores.length + 3;
  writeCell(sheet, `F${totalRow}`, "合计", sheet.F47 ?? styleRows[5]);
  const totalCell = clone(sheet.G47 ?? styleRows[6]);
  totalCell.t = "n";
  totalCell.f = `SUM(G3:G${totalRow - 1})`;
  delete totalCell.v;
  sheet[`G${totalRow}`] = totalCell;
  sheet["!ref"] = `A1:G${totalRow}`;
}

export async function exportMonthlyWorkbook(payload: ExportPayload) {
  const XLSX = await import("xlsx");
  const response = await fetch(`${import.meta.env.BASE_URL}monthly-export-template.xlsx`);
  if (!response.ok) throw new Error("无法读取月度导出模板");
  const workbook = XLSX.read(await response.arrayBuffer(), { type: "array", cellStyles: true, cellFormula: true, bookVBA: true });
  const [rundownName, performanceName, microName, newsName] = workbook.SheetNames;
  buildRundownSheet(XLSX, workbook.Sheets[rundownName], payload);
  updateNewMediaSheet(XLSX, workbook.Sheets[performanceName], payload);
  updateMicroSheet(XLSX, workbook.Sheets[microName], payload);
  updateNewsCenterSheet(XLSX, workbook.Sheets[newsName], payload);
  const renamed = workbook.SheetNames.map((name) => name.replace(/^7月/, `${payload.month}月`));
  workbook.SheetNames.forEach((oldName, index) => {
    if (renamed[index] !== oldName) {
      workbook.Sheets[renamed[index]] = workbook.Sheets[oldName];
      delete workbook.Sheets[oldName];
    }
  });
  workbook.SheetNames = renamed;
  XLSX.writeFile(workbook, `新媒体${payload.year}年${payload.month}月工分与串单总表.xlsx`, { cellStyles: true, bookType: "xlsx" });
}
