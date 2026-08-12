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

type ExportPayload = {
  year: number;
  month: number;
  days: ExportDay[];
  newMediaScores: NewMediaScore[];
  nonMediaScores: NewMediaScore[];
  newsCenterScores: NewsCenterScore[];
  microRewards: MicroReward[];
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
  sheet.getCell("B26").value = "新媒体人员合计分";
  const newMediaNameStyle = clone(sheet.getCell("B27").style);
  const newMediaTotalStyle = clone(sheet.getCell("C27").style);
  const newMediaHeight = sheet.getRow(27).height;
  const nonMediaHeaderStyle = Array.from({ length: 6 }, (_, column) => clone(sheet.getCell(37, column + 1).style));
  const nonMediaHeaderHeight = sheet.getRow(37).height;
  const nonMediaNameStyle = clone(sheet.getCell("B38").style);
  const nonMediaTotalStyle = clone(sheet.getCell("C38").style);
  const nonMediaHeight = sheet.getRow(38).height;
  if (sheet.getCell("B37").isMerged) sheet.unMergeCells("B37:C37");
  for (let row = 27; row <= 41; row += 1) {
    const score = payload.newMediaScores[row - 27];
    const targetRow = sheet.getRow(row);
    targetRow.getCell(2).style = clone(newMediaNameStyle);
    targetRow.getCell(3).style = clone(newMediaTotalStyle);
    targetRow.getCell(2).value = score?.name ?? "";
    targetRow.getCell(3).value = score
      ? { formula: `SUM(${score.duty},${score.wechat},${score.remix},${score.original},${score.fixed})` }
      : "";
    if (newMediaHeight !== undefined) targetRow.height = newMediaHeight;
  }
  const nonMediaHeaderRow = 43;
  for (let column = 1; column <= 6; column += 1) {
    const cell = sheet.getCell(nonMediaHeaderRow, column);
    cell.style = clone(nonMediaHeaderStyle[column - 1]);
    cell.value = "";
  }
  if (nonMediaHeaderHeight !== undefined) sheet.getRow(nonMediaHeaderRow).height = nonMediaHeaderHeight;
  sheet.mergeCells(`B${nonMediaHeaderRow}:C${nonMediaHeaderRow}`);
  sheet.getCell(`B${nonMediaHeaderRow}`).value = "非新媒体人员合计分";
  const firstNonMediaRow = nonMediaHeaderRow + 1;
  const neededLastRow = firstNonMediaRow + Math.max(15, payload.nonMediaScores.length) - 1;
  for (let row = firstNonMediaRow; row <= neededLastRow; row += 1) {
    const score = payload.nonMediaScores[row - firstNonMediaRow];
    const targetRow = sheet.getRow(row);
    targetRow.getCell(2).style = clone(nonMediaNameStyle);
    targetRow.getCell(3).style = clone(nonMediaTotalStyle);
    targetRow.getCell(2).value = score?.name ?? "";
    targetRow.getCell(3).value = score?.total ?? "";
    if (nonMediaHeight !== undefined) targetRow.height = nonMediaHeight;
  }
  const bonusRows = [
    [18, "周婷", "【新闻回看】视频回看每日上传绩效：24分；", 24],
    [20, "刘乐", "新闻节目上传：36分", 36],
    [22, "吴轲宇", "接听指令电话", 20],
  ] as const;
  bonusRows.forEach(([row, name, label, points]) => {
    sheet.getCell(`A${row}`).value = name;
    sheet.getCell(`B${row}`).value = label;
    sheet.getCell(`F${row}`).value = points;
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
  payload.newsCenterScores.forEach((person, index) => {
    const row = index + 3;
    const values = [person.order, person.name, person.forwardCount, person.micro || "", person.video || "", person.justNow || ""];
    values.forEach((value, column) => {
      sheet.getCell(row, column + 1).value = value;
    });
    sheet.getCell(`G${row}`).value = { formula: `C${row}/25+F${row}+E${row}+D${row}` };
  });
  const totalRow = payload.newsCenterScores.length + 3;
  sheet.getCell(`G${totalRow}`).value = { formula: `SUM(G3:G${totalRow - 1})` };
}

export async function exportMonthlyWorkbook(payload: ExportPayload) {
  const ExcelJS = (await import("exceljs")).default;
  const response = await fetch(`${import.meta.env.BASE_URL}monthly-performance-template.xlsx`);
  if (!response.ok) throw new Error("无法读取月度导出模板");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await response.arrayBuffer());
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
  const content = await workbook.xlsx.writeBuffer();
  const blob = new Blob([content as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `新媒体${payload.year}年${payload.month}月绩效.xlsx`;
  anchor.click();
  URL.revokeObjectURL(url);
}
