import ExcelJS from "exceljs";
import { addDays, format, startOfWeek, differenceInCalendarDays } from "date-fns";
import type { ProjectInputs } from "./types";
import { computeTimeline, groupByMonth } from "./engine";
import { getTemplate, REPORT_TYPE_LABELS } from "./library";

const COLOR_HEADER_BG = "FF1F4E78";
const COLOR_HEADER_FG = "FFFFFFFF";
const COLOR_SUBHEADER_BG = "FFD9E1F2";
const COLOR_MONTH_BG = "FFFCE4D6";
const COLOR_BORDER = "FF808080";
const COLOR_BAR = "FF4F81BD";
const COLOR_BAR_CLIENT = "FFE8A33D";
const COLOR_BAR_BOTH = "FF6FAE6B";

function thinBorder() {
  return {
    top: { style: "thin" as const, color: { argb: COLOR_BORDER } },
    left: { style: "thin" as const, color: { argb: COLOR_BORDER } },
    right: { style: "thin" as const, color: { argb: COLOR_BORDER } },
    bottom: { style: "thin" as const, color: { argb: COLOR_BORDER } },
  };
}

function colLetter(col: number): string {
  let s = "";
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

function ownerLabel(r: string): string {
  if (r === "client") return "Client";
  if (r === "wp") return "WP";
  return "Both";
}

export async function buildTimelineWorkbook(inputs: ProjectInputs): Promise<ExcelJS.Workbook> {
  const { anchors, tasks } = computeTimeline(inputs);
  const template = getTemplate(inputs.reportType);
  const phaseName = (id: string) => template.phases.find((p) => p.id === id)?.name ?? id;

  const wb = new ExcelJS.Workbook();
  wb.creator = "Timeline Generator";
  wb.created = new Date();

  // ----- Sheet 1: Timelines (TEIL-style readable plan) -----
  const ws = wb.addWorksheet("Timelines", { views: [{ state: "frozen", ySplit: 9 }] });
  ws.columns = [
    { width: 2 },
    { width: 18 },
    { width: 60 },
    { width: 10 },
    { width: 10 },
    { width: 13 },
    { width: 13 },
    { width: 10 },
  ];

  ws.mergeCells("B2:H2");
  const title = ws.getCell("B2");
  title.value = `Tentative Timeline | ${inputs.clientName} ${REPORT_TYPE_LABELS[inputs.reportType]} ${inputs.reportingPeriod}`;
  title.font = { bold: true, size: 14, color: { argb: COLOR_HEADER_FG } };
  title.alignment = { horizontal: "center", vertical: "middle" };
  title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_HEADER_BG } };
  ws.getRow(2).height = 26;

  ws.mergeCells("B3:C3");
  ws.getCell("B3").value = "Key Dates";
  ws.getCell("B3").font = { bold: true };
  ws.getCell("B3").fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_SUBHEADER_BG } };
  ws.mergeCells("E3:H3");
  ws.getCell("E3").value = "Single Point of Contact";
  ws.getCell("E3").font = { bold: true };
  ws.getCell("E3").fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_SUBHEADER_BG } };

  ws.getCell("B4").value = "Kick-off";
  ws.getCell("C4").value = format(anchors.kick_off, "dd-MM-yyyy");
  ws.getCell("B5").value = "Closure";
  ws.getCell("C5").value = format(anchors.closure, "dd-MM-yyyy");
  if (anchors.board_meeting) {
    ws.getCell("B6").value = "Board Meeting";
    ws.getCell("C6").value = format(anchors.board_meeting, "dd-MM-yyyy");
  }
  if (anchors.agm) {
    ws.getCell("B7").value = "AGM";
    ws.getCell("C7").value = format(anchors.agm, "dd-MM-yyyy");
  }
  if (inputs.spoc) {
    ws.getCell("E4").value = inputs.spoc.team;
    ws.getCell("F4").value = inputs.spoc.name;
  }

  const headerRow = ws.getRow(9);
  headerRow.values = ["", "Month", "Process Description", "Client", "WP", "Start", "End", "Duration"];
  headerRow.eachCell((cell, col) => {
    if (col === 1) return;
    cell.font = { bold: true, color: { argb: COLOR_HEADER_FG } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_HEADER_BG } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = thinBorder();
  });
  headerRow.height = 22;

  const grouped = groupByMonth(tasks);
  let row = 10;
  const taskRowMap = new Map<string, number>();
  for (const [month, monthTasks] of grouped) {
    let firstInMonth = true;
    for (const t of monthTasks) {
      const r = ws.getRow(row);
      r.getCell(2).value = firstInMonth ? month : "";
      r.getCell(3).value = t.name;
      r.getCell(4).value = t.responsibility === "client" || t.responsibility === "both" ? "P" : "";
      r.getCell(5).value = t.responsibility === "wp" || t.responsibility === "both" ? "P" : "";
      r.getCell(6).value = t.startDate;
      r.getCell(6).numFmt = "dd-mm-yyyy";
      r.getCell(7).value = t.endDate;
      r.getCell(7).numFmt = "dd-mm-yyyy";
      r.getCell(8).value = { formula: `G${row}-F${row}+1`, result: differenceInCalendarDays(t.endDate, t.startDate) + 1 };

      if (firstInMonth) {
        r.getCell(2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_MONTH_BG } };
        r.getCell(2).font = { bold: true };
      }
      r.getCell(3).alignment = { vertical: "middle", wrapText: true };
      r.getCell(4).alignment = { horizontal: "center" };
      r.getCell(5).alignment = { horizontal: "center" };
      r.getCell(6).alignment = { horizontal: "center" };
      r.getCell(7).alignment = { horizontal: "center" };
      r.getCell(8).alignment = { horizontal: "center" };
      [2, 3, 4, 5, 6, 7, 8].forEach((c) => (r.getCell(c).border = thinBorder()));

      taskRowMap.set(t.id, row);
      firstInMonth = false;
      row++;
    }
  }

  // ----- Sheet 2: Gantt (interactive weekly Gantt with conditional formatting) -----
  const gs = wb.addWorksheet("Gantt", { views: [{ state: "frozen", xSplit: 7, ySplit: 5 }] });

  // earliest start, latest end across all tasks
  const earliest = tasks.reduce((m, t) => (t.startDate < m ? t.startDate : m), tasks[0].startDate);
  const latest = tasks.reduce((m, t) => (t.endDate > m ? t.endDate : m), tasks[0].endDate);
  const firstWeek = startOfWeek(earliest, { weekStartsOn: 1 });
  const lastWeek = startOfWeek(latest, { weekStartsOn: 1 });
  const weeks: Date[] = [];
  for (let d = firstWeek; d <= lastWeek; d = addDays(d, 7)) weeks.push(d);

  // Column widths: 1 spacer, 2 task, 3 phase, 4 owner, 5 start, 6 end, 7 dur, 8+ weeks
  const cols: Partial<ExcelJS.Column>[] = [
    { width: 2 },
    { width: 55 },
    { width: 22 },
    { width: 8 },
    { width: 12 },
    { width: 12 },
    { width: 9 },
  ];
  weeks.forEach(() => cols.push({ width: 4 }));
  gs.columns = cols;

  // Title
  gs.mergeCells(2, 2, 2, 7 + weeks.length);
  const gTitle = gs.getCell(2, 2);
  gTitle.value = `Gantt | ${inputs.clientName} ${REPORT_TYPE_LABELS[inputs.reportType]} ${inputs.reportingPeriod}`;
  gTitle.font = { bold: true, size: 14, color: { argb: COLOR_HEADER_FG } };
  gTitle.alignment = { horizontal: "center", vertical: "middle" };
  gTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_HEADER_BG } };
  gs.getRow(2).height = 26;

  // Month band (row 4) above week headers (row 5)
  const headerR = gs.getRow(5);
  ["", "Task", "Phase", "Owner", "Start", "End", "Days"].forEach((h, i) => {
    if (i === 0) return;
    const c = headerR.getCell(i + 1);
    c.value = h;
    c.font = { bold: true, color: { argb: COLOR_HEADER_FG } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_HEADER_BG } };
    c.alignment = { horizontal: "center", vertical: "middle" };
    c.border = thinBorder();
  });

  // Month band row
  const monthR = gs.getRow(4);
  let monthStart = -1;
  let currentMonth = "";
  weeks.forEach((w, i) => {
    const colIdx = 8 + i;
    const m = format(w, "MMM yyyy");
    if (m !== currentMonth) {
      if (monthStart >= 0) {
        gs.mergeCells(4, monthStart, 4, colIdx - 1);
        const mc = gs.getCell(4, monthStart);
        mc.value = currentMonth;
        mc.font = { bold: true, color: { argb: COLOR_HEADER_FG } };
        mc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_HEADER_BG } };
        mc.alignment = { horizontal: "center" };
        mc.border = thinBorder();
      }
      currentMonth = m;
      monthStart = colIdx;
    }
    if (i === weeks.length - 1) {
      gs.mergeCells(4, monthStart, 4, colIdx);
      const mc = gs.getCell(4, monthStart);
      mc.value = currentMonth;
      mc.font = { bold: true, color: { argb: COLOR_HEADER_FG } };
      mc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_HEADER_BG } };
      mc.alignment = { horizontal: "center" };
      mc.border = thinBorder();
    }
  });

  // Week headers (row 5) — store actual Date so conditional formatting works
  weeks.forEach((w, i) => {
    const c = gs.getCell(5, 8 + i);
    c.value = w;
    c.numFmt = "dd-mmm";
    c.font = { bold: true, color: { argb: COLOR_HEADER_FG }, size: 9 };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_HEADER_BG } };
    c.alignment = { horizontal: "center", vertical: "middle", textRotation: 90 };
    c.border = thinBorder();
  });
  gs.getRow(5).height = 60;

  // Task rows
  const ganttStartRow = 6;
  tasks.forEach((t, i) => {
    const r = ganttStartRow + i;
    const row = gs.getRow(r);
    row.getCell(2).value = t.name;
    row.getCell(2).alignment = { vertical: "middle", wrapText: true };
    row.getCell(3).value = phaseName(t.phase);
    row.getCell(3).alignment = { vertical: "middle" };
    row.getCell(4).value = ownerLabel(t.responsibility);
    row.getCell(4).alignment = { horizontal: "center" };
    row.getCell(5).value = t.startDate;
    row.getCell(5).numFmt = "dd-mm-yyyy";
    row.getCell(6).value = t.endDate;
    row.getCell(6).numFmt = "dd-mm-yyyy";
    row.getCell(7).value = { formula: `F${r}-E${r}+1`, result: differenceInCalendarDays(t.endDate, t.startDate) + 1 };
    row.getCell(7).alignment = { horizontal: "center" };
    [2, 3, 4, 5, 6, 7].forEach((c) => (row.getCell(c).border = thinBorder()));
    // empty borders for week cells
    for (let w = 0; w < weeks.length; w++) {
      row.getCell(8 + w).border = thinBorder();
    }
  });

  // Conditional formatting: highlight week-cell if week overlaps task range
  // Range: H6 : <lastCol><lastRow>
  const firstWeekColL = colLetter(8);
  const lastWeekColL = colLetter(7 + weeks.length);
  const lastTaskRow = ganttStartRow + tasks.length - 1;
  const ganttRange = `${firstWeekColL}${ganttStartRow}:${lastWeekColL}${lastTaskRow}`;

  // Owner-aware coloring: 3 rules so client/wp/both get distinct colors
  gs.addConditionalFormatting({
    ref: ganttRange,
    rules: [
      {
        type: "expression",
        priority: 1,
        formulae: [
          `AND(${firstWeekColL}$5<=$F${ganttStartRow}, ${firstWeekColL}$5+6>=$E${ganttStartRow}, $D${ganttStartRow}="Both")`,
        ],
        style: {
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: COLOR_BAR_BOTH } },
        },
      },
      {
        type: "expression",
        priority: 2,
        formulae: [
          `AND(${firstWeekColL}$5<=$F${ganttStartRow}, ${firstWeekColL}$5+6>=$E${ganttStartRow}, $D${ganttStartRow}="Client")`,
        ],
        style: {
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: COLOR_BAR_CLIENT } },
        },
      },
      {
        type: "expression",
        priority: 3,
        formulae: [
          `AND(${firstWeekColL}$5<=$F${ganttStartRow}, ${firstWeekColL}$5+6>=$E${ganttStartRow})`,
        ],
        style: {
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: COLOR_BAR } },
        },
      },
    ],
  });

  // Legend below
  const legendRow = lastTaskRow + 2;
  gs.getCell(legendRow, 2).value = "Legend:";
  gs.getCell(legendRow, 2).font = { bold: true };
  gs.getCell(legendRow, 3).value = "WP";
  gs.getCell(legendRow, 3).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_BAR } };
  gs.getCell(legendRow, 3).font = { color: { argb: COLOR_HEADER_FG }, bold: true };
  gs.getCell(legendRow, 3).alignment = { horizontal: "center" };
  gs.getCell(legendRow, 4).value = "Client";
  gs.getCell(legendRow, 4).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_BAR_CLIENT } };
  gs.getCell(legendRow, 4).font = { color: { argb: COLOR_HEADER_FG }, bold: true };
  gs.getCell(legendRow, 4).alignment = { horizontal: "center" };
  gs.getCell(legendRow, 5).value = "Both";
  gs.getCell(legendRow, 5).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_BAR_BOTH } };
  gs.getCell(legendRow, 5).font = { color: { argb: COLOR_HEADER_FG }, bold: true };
  gs.getCell(legendRow, 5).alignment = { horizontal: "center" };

  gs.getCell(legendRow + 2, 2).value =
    "Tip: edit Start / End dates and the bars update automatically.";
  gs.getCell(legendRow + 2, 2).font = { italic: true, color: { argb: "FF666666" } };

  // ----- Sheet 3: Phases -----
  const ps = wb.addWorksheet("Phases");
  ps.columns = [{ width: 4 }, { width: 30 }, { width: 14 }, { width: 14 }, { width: 14 }];
  ps.getRow(2).values = ["", "Phase", "Tasks", "Start", "End"];
  ps.getRow(2).font = { bold: true };
  ps.getRow(2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_SUBHEADER_BG } };

  let pr = 3;
  for (const phase of template.phases) {
    const phaseTasks = tasks.filter((t) => t.phase === phase.id);
    if (!phaseTasks.length) continue;
    const start = phaseTasks.reduce((m, t) => (t.startDate < m ? t.startDate : m), phaseTasks[0].startDate);
    const end = phaseTasks.reduce((m, t) => (t.endDate > m ? t.endDate : m), phaseTasks[0].endDate);
    ps.getRow(pr).values = ["", phase.name, phaseTasks.length, format(start, "dd-MM-yyyy"), format(end, "dd-MM-yyyy")];
    pr++;
  }

  ps.mergeCells(`B${pr + 2}:E${pr + 2}`);
  ps.getCell(`B${pr + 2}`).value = "Please note:";
  ps.getCell(`B${pr + 2}`).font = { bold: true };
  ps.mergeCells(`B${pr + 3}:E${pr + 3}`);
  ps.getCell(`B${pr + 3}`).value =
    "Deliverables are subject to receipt of necessary information and inputs from the client at the right time.";
  ps.getCell(`B${pr + 3}`).alignment = { wrapText: true };
  if (inputs.reportType === "annual_report" || inputs.reportType === "integrated_report") {
    ps.mergeCells(`B${pr + 4}:E${pr + 4}`);
    ps.getCell(`B${pr + 4}`).value =
      "Print files will be provided 3-4 working days post the release of the web upload file.";
    ps.getCell(`B${pr + 4}`).alignment = { wrapText: true };
  }

  return wb;
}

export async function timelineXlsxBuffer(inputs: ProjectInputs): Promise<Buffer> {
  const wb = await buildTimelineWorkbook(inputs);
  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf as ArrayBuffer);
}
