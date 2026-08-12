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
  videoEditor: string;
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

type WechatPerformance = {
  name: string;
  dutyDays: number;
  parkHrCount: number;
};

type VideoPerformance = {
  name: string;
  pieces: number;
  videoPoints: number;
  dutyDays: number;
};

type ExportPayload = {
  year: number;
  month: number;
  days: ExportDay[];
  newMediaScores: NewMediaScore[];
  nonMediaScores: NewMediaScore[];
  newsCenterScores: NewsCenterScore[];
  microRewards: MicroReward[];
  wechatPerformance: WechatPerformance[];
  videoPerformance: VideoPerformance[];
  newMediaNames: string[];
};

type DailyExportPayload = ExportDay;

function clone<T>(value: T): T {
  return value === undefined ? value : JSON.parse(JSON.stringify(value));
}

function dailyBodySourceRow(category: "wechat" | "remix" | "original", index: number) {
  if (category === "wechat") return index === 0 ? 4 : index < 4 ? 5 : 8;
  if (category === "remix") return index === 0 ? 11 : 12;
  return index === 0 ? 15 : 16;
}

function buildDailySheet(sheet: any, day: DailyExportPayload) {
  const sourceRows = [1, 2, 3, 4, 5, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18];
  const templateCells: Record<number, any[]> = {};
  const templateHeights: Record<number, number | undefined> = {};
  sourceRows.forEach((row) => {
    templateCells[row] = Array.from({ length: 7 }, (_, column) => clone(sheet.getRow(row).getCell(column + 1).style));
    templateHeights[row] = sheet.getRow(row).height;
  });
  const columnWidths = Array.from({ length: 7 }, (_, column) => sheet.getColumn(column + 1).width);
  Object.values(sheet._merges ?? {}).forEach((merge: any) => {
    if (merge?.range) sheet.unMergeCells(merge.range);
  });
  sheet.spliceRows(1, sheet.rowCount);
  const writeRow = (sourceRow: number, targetRow: number, values: unknown[]) => {
    const target = sheet.getRow(targetRow);
    for (let column = 1; column <= 7; column += 1) {
      const cell = target.getCell(column);
      cell.style = clone(templateCells[sourceRow][column - 1]);
      cell.value = values[column - 1] ?? "";
    }
    if (templateHeights[sourceRow] !== undefined) target.height = templateHeights[sourceRow];
  };
  const merge = (row: number) => sheet.mergeCells(row, 1, row, 7);

  const [year, month, date] = day.date.split("-").map(Number);
  let row = 1;
  writeRow(1, row, [`新媒体串片单  ${year}-${month}-${date}${day.weekday}`]);
  merge(row);
  row += 1;
  writeRow(2, row, [`微信公众号  本周微刊小编：${day.wechatEditor || ""}`]);
  merge(row);
  row += 1;
  writeRow(3, row, ["序号", " 节 目 标 题", "记者", "阅读量", "刚刚帖", "工分", "备注"]);
  row += 1;
  day.sections.wechat.forEach((entry, index) => {
    writeRow(dailyBodySourceRow("wechat", index), row, [
      index + 1,
      entry.title,
      entry.staff,
      entry.views,
      entry.duration,
      entry.exportPoints || "",
      entry.notes,
    ]);
    row += 1;
  });

  writeRow(9, row, [
    `小编二创短视频（含制作封面、包框） 本周短视频主班小编：${day.videoEditor || ""}`,
  ]);
  merge(row);
  row += 1;
  writeRow(10, row, ["序号", " 节 目 标 题", "编辑", "阅读量", "时长", "工分", "备注"]);
  row += 1;
  day.sections.remix.forEach((entry, index) => {
    writeRow(dailyBodySourceRow("remix", index), row, [
      index + 1,
      entry.title,
      entry.staff,
      entry.views,
      entry.duration,
      entry.exportPoints || "",
      entry.notes,
    ]);
    row += 1;
  });

  writeRow(13, row, ["记者原创短视频（新媒体首发）"]);
  merge(row);
  row += 1;
  writeRow(14, row, ["序号", " 节 目 标 题", "采编", "阅读量", "时长", "工分", "备注"]);
  row += 1;
  day.sections.original.forEach((entry, index) => {
    writeRow(dailyBodySourceRow("original", index), row, [
      index + 1,
      entry.title,
      entry.staff,
      entry.views,
      entry.duration,
      entry.exportPoints || "",
      entry.notes,
    ]);
    row += 1;
  });

  writeRow(17, row, [
    `   值班责编:${day.dutyEditor || ""}     值班主任： ${day.dutyDirector || ""}        监审：${day.supervisor || ""}`,
  ]);
  merge(row);
  row += 1;
  writeRow(18, row, []);
  columnWidths.forEach((width, column) => {
    if (width !== undefined) sheet.getColumn(column + 1).width = width;
  });
}

export async function createDailyWorkbook(day: DailyExportPayload) {
  const ExcelJS = (await import("exceljs")).default;
  const response = await fetch(`${import.meta.env.BASE_URL}daily-export-template.xlsx`);
  if (!response.ok) throw new Error("无法读取日串单导出模板");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await response.arrayBuffer());
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error("日串单模板中没有工作表");
  buildDailySheet(sheet, day);
  return workbook;
}

export async function exportDailyWorkbook(day: DailyExportPayload) {
  const workbook = await createDailyWorkbook(day);
  const [year, month, date] = day.date.split("-").map(Number);
  const content = await workbook.xlsx.writeBuffer();
  const blob = new Blob([content as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `${year}年${month}月${date}日串单.xlsx`;
  anchor.click();
  URL.revokeObjectURL(url);
}

function buildRundownSheet(sheet: any, payload: ExportPayload) {
  const sourceRows = [1, 2, 3, 4, 5, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19];
  const templateCells: Record<number, any[]> = {};
  const templateHeights: Record<number, number | undefined> = {};
  sourceRows.forEach((row) => {
    templateCells[row] = Array.from({ length: 7 }, (_, column) => clone(sheet.getCell(row, column + 1).style));
    templateHeights[row] = sheet.getRow(row).height;
  });
  const columnWidths = Array.from({ length: 7 }, (_, column) => sheet.getColumn(column + 1).width);
  Object.values(sheet._merges ?? {}).forEach((merge: any) => {
    if (merge?.range) sheet.unMergeCells(merge.range);
  });
  sheet.spliceRows(1, sheet.rowCount);
  const writeRow = (sourceRow: number, targetRow: number, values: unknown[]) => {
    const target = sheet.getRow(targetRow);
    for (let column = 1; column <= 7; column += 1) {
      target.getCell(column).style = clone(templateCells[sourceRow][column - 1]);
      target.getCell(column).value = values[column - 1] ?? "";
    }
    if (templateHeights[sourceRow] !== undefined) target.height = templateHeights[sourceRow];
  };
  const merge = (row: number) => sheet.mergeCells(row, 1, row, 7);
  let row = 1;
  payload.days.forEach((day) => {
    const [year, month, date] = day.date.split("-").map(Number);
    writeRow(1, row, [`新媒体串片单  ${year}-${month}-${date}${day.weekday}`]); merge(row); row += 1;
    writeRow(2, row, [`微信公众号  本周微刊小编：${day.wechatEditor || ""}`]); merge(row); row += 1;
    writeRow(3, row, ["序号", " 节 目 标 题", "记者", "阅读量", "刚刚帖", "工分", "备注"]); row += 1;
    day.sections.wechat.forEach((entry, index) => {
      writeRow(dailyBodySourceRow("wechat", index), row, [index + 1, entry.title, entry.staff, entry.views, entry.duration, entry.exportPoints || "", entry.notes]);
      row += 1;
    });
    writeRow(9, row, [`小编二创短视频（含制作封面、包框） 本周短视频主班小编：${day.videoEditor || ""}`]); merge(row); row += 1;
    writeRow(10, row, ["序号", " 节 目 标 题", "编辑", "阅读量", "时长", "工分", "备注"]); row += 1;
    day.sections.remix.forEach((entry, index) => {
      writeRow(dailyBodySourceRow("remix", index), row, [index + 1, entry.title, entry.staff, entry.views, entry.duration, entry.exportPoints || "", entry.notes]);
      row += 1;
    });
    writeRow(13, row, ["记者原创短视频（新媒体首发）"]); merge(row); row += 1;
    writeRow(14, row, ["序号", " 节 目 标 题", "采编", "阅读量", "时长", "工分", "备注"]); row += 1;
    day.sections.original.forEach((entry, index) => {
      writeRow(dailyBodySourceRow("original", index), row, [index + 1, entry.title, entry.staff, entry.views, entry.duration, entry.exportPoints || "", entry.notes]);
      row += 1;
    });
    writeRow(17, row, [`   值班责编:${day.dutyEditor || ""}     值班主任： ${day.dutyDirector || ""}        监审：${day.supervisor || ""}`]); merge(row); row += 1;
    writeRow(18, row, []); row += 1;
    writeRow(19, row, []); row += 1;
  });
  columnWidths.forEach((width, column) => {
    if (width !== undefined) sheet.getColumn(column + 1).width = width;
  });
}

function updateNewMediaSheet(sheet: any, payload: ExportPayload) {
  const sourceRows = [1, 2, 3, 5, 6, 7, 17, 18, 26, 27, 37, 38];
  const templateCells: Record<number, any[]> = {};
  const templateHeights: Record<number, number | undefined> = {};
  sourceRows.forEach((row) => {
    templateCells[row] = Array.from({ length: 6 }, (_, column) => clone(sheet.getCell(row, column + 1).style));
    templateHeights[row] = sheet.getRow(row).height;
  });
  const columnWidths = Array.from({ length: 6 }, (_, column) => sheet.getColumn(column + 1).width);
  Object.values(sheet._merges ?? {}).forEach((merge: any) => {
    if (merge?.range) sheet.unMergeCells(merge.range);
  });
  const previousRowCount = sheet.rowCount;
  sheet.spliceRows(1, sheet.rowCount);
  for (let row = 1; row <= previousRowCount; row += 1) {
    for (let column = 1; column <= 6; column += 1) sheet.getCell(row, column).value = null;
  }

  const writeRow = (sourceRow: number, targetRow: number, values: unknown[]) => {
    const target = sheet.getRow(targetRow);
    for (let column = 1; column <= 6; column += 1) {
      const cell = target.getCell(column);
      cell.style = clone(templateCells[sourceRow][column - 1]);
      cell.value = null;
      cell.value = values[column - 1] ?? "";
    }
    if (templateHeights[sourceRow] !== undefined) target.height = templateHeights[sourceRow];
  };
  const mergeTitle = (row: number) => sheet.mergeCells(row, 1, row, 6);
  let row = 1;

  writeRow(1, row, ["微信公众号绩效"]); mergeTitle(row); row += 1;
  writeRow(2, row, ["姓名", "系数", "基础绩效（分/周）", "工业园区/人社局微刊（次）", "公众号值班天数", "合计（分）"]); row += 1;
  payload.wechatPerformance.forEach((person) => {
    writeRow(3, row, [
      person.name,
      1,
      48,
      person.parkHrCount,
      person.dutyDays,
      { formula: `ROUND(C${row}*B${row}/7*E${row}+D${row}*6,1)` },
    ]);
    row += 1;
  });
  row += 1;

  writeRow(5, row, ["抖音、视频号绩效"]); mergeTitle(row); row += 1;
  writeRow(6, row, ["姓名", "系数1", "条数", "工分", "责编天数", "合计（分）"]); row += 1;
  payload.videoPerformance.forEach((person) => {
    writeRow(7, row, [
      person.name,
      1,
      person.pieces,
      person.videoPoints,
      person.dutyDays,
      { formula: `ROUND(D${row}*B${row}+E${row}*8,1)` },
    ]);
    row += 1;
  });
  row += 1;

  writeRow(17, row, ["其他绩效"]); mergeTitle(row); row += 1;
  const bonusRows = [
    ["周婷", "【新闻回看】视频回看每日上传绩效：24分", 24],
    ["刘乐", "新闻节目上传：36分", 36],
    ["吴轲宇", "接听指令电话：20分", 20],
  ] as const;
  bonusRows.forEach(([name, label, points]) => {
    writeRow(18, row, [name, label, "", "", "", points]);
    sheet.mergeCells(row, 2, row, 5);
    row += 1;
  });
  row += 1;

  writeRow(26, row, ["", "新媒体人员合计分"]);
  sheet.mergeCells(row, 2, row, 3);
  row += 1;
  payload.newMediaScores.forEach((score) => {
    writeRow(27, row, ["", score.name, { formula: `ROUND(SUM(${score.duty},${score.wechat},${score.remix},${score.original},${score.fixed}),1)` }]);
    row += 1;
  });
  row += 1;

  writeRow(37, row, ["", "非新媒体人员合计分"]);
  sheet.mergeCells(row, 2, row, 3);
  row += 1;
  payload.nonMediaScores.forEach((score) => {
    writeRow(38, row, ["", score.name, score.total]);
    row += 1;
  });
  if (!payload.nonMediaScores.length) writeRow(38, row, ["", "", 0]);

  columnWidths.forEach((width, column) => {
    if (width !== undefined) sheet.getColumn(column + 1).width = width;
  });
}

function updateMicroSheet(sheet: any, payload: ExportPayload) {
  sheet.getCell("A1").value = `${payload.month}月份刚刚帖及微刊阅读量超5000部分`;
  const template = Array.from({ length: 6 }, (_, column) => clone(sheet.getCell(3, column + 1).style));
  const templateHeight = sheet.getRow(3).height;
  const rowsToClear = Math.max(sheet.rowCount, payload.microRewards.length + 2);
  for (let row = 3; row <= rowsToClear; row += 1) {
    for (let column = 1; column <= 6; column += 1) {
      const cell = sheet.getCell(row, column);
      cell.style = clone(template[column - 1]);
      cell.value = "";
    }
    if (templateHeight !== undefined) sheet.getRow(row).height = templateHeight;
  }
  payload.microRewards.forEach((item, index) => {
    const row = index + 3;
    const values = [index + 1, item.title, item.views, item.staff, item.reward, item.notes];
    values.forEach((value, column) => {
      sheet.getCell(row, column + 1).value = value;
    });
  });
}

function updateNewsCenterSheet(sheet: any, payload: ExportPayload) {
  sheet.getCell("A1").value = `${payload.month}月新媒体工分`;
  const bodyStyles = Array.from({ length: 7 }, (_, column) => clone(sheet.getCell(3, column + 1).style));
  const bodyHeight = sheet.getRow(3).height;
  const totalStyles = Array.from({ length: 7 }, (_, column) => clone(sheet.getCell(48, column + 1).style));
  const totalHeight = sheet.getRow(48).height;
  for (let row = 3; row <= sheet.rowCount; row += 1) {
    for (let column = 1; column <= 7; column += 1) sheet.getCell(row, column).value = null;
  }
  payload.newsCenterScores.forEach((person, index) => {
    const row = index + 3;
    const isNewMedia = payload.newMediaNames.includes(person.name);
    const values = [
      person.order,
      person.name,
      person.forwardCount,
      isNewMedia ? 0 : person.micro || 0,
      isNewMedia ? 0 : person.video || 0,
      isNewMedia ? 0 : person.justNow || 0,
    ];
    values.forEach((value, column) => {
      const cell = sheet.getCell(row, column + 1);
      cell.style = clone(bodyStyles[column]);
      cell.value = null;
      cell.value = value;
    });
    if (bodyHeight !== undefined) sheet.getRow(row).height = bodyHeight;
    sheet.getCell(`G${row}`).style = clone(bodyStyles[6]);
    sheet.getCell(`G${row}`).value = { formula: `C${row}/25+E${row}+F${row}` };
  });
  const totalRow = payload.newsCenterScores.length + 3;
  for (let column = 1; column <= 7; column += 1) {
    const cell = sheet.getCell(totalRow, column);
    cell.style = clone(totalStyles[column - 1]);
    cell.value = null;
    cell.value = "";
  }
  if (totalHeight !== undefined) sheet.getRow(totalRow).height = totalHeight;
  sheet.getCell(`G${totalRow}`).value = { formula: `SUM(G3:G${totalRow - 1})` };
}

export function populateMonthlyWorkbook(workbook: any, payload: ExportPayload) {
  const [rundownSheet, performanceSheet, microSheet, newsSheet] = workbook.worksheets;
  if (!rundownSheet || !performanceSheet || !microSheet || !newsSheet) throw new Error("月度模板必须包含四个工作表");
  const originalSheetNames = workbook.worksheets.map((sheet: any) => sheet.name);
  const temporaryNames = workbook.worksheets.map((sheet: any, index: number) => `__monthly_export_${index + 1}__`);
  workbook.worksheets.forEach((sheet: any, index: number) => {
    sheet.name = temporaryNames[index];
  });
  buildRundownSheet(rundownSheet, payload);
  updateNewMediaSheet(performanceSheet, payload);
  updateMicroSheet(microSheet, payload);
  updateNewsCenterSheet(newsSheet, payload);
  workbook.worksheets.forEach((sheet: any, index: number) => {
    sheet.name = originalSheetNames[index].replace(/^7月/, `${payload.month}月`);
  });
  workbook.calcProperties.fullCalcOnLoad = true;
  return workbook;
}

export async function exportMonthlyWorkbook(payload: ExportPayload) {
  const ExcelJS = (await import("exceljs")).default;
  const response = await fetch(`${import.meta.env.BASE_URL}monthly-performance-template.xlsx`);
  if (!response.ok) throw new Error("无法读取月度导出模板");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await response.arrayBuffer());
  populateMonthlyWorkbook(workbook, payload);
  const content = await workbook.xlsx.writeBuffer();
  const blob = new Blob([content as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `新媒体${payload.year}年${payload.month}月绩效.xlsx`;
  anchor.click();
  URL.revokeObjectURL(url);
}
