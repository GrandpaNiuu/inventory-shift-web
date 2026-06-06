"use client";

import { useMemo, useState, type ChangeEvent, type ReactNode } from "react";
import * as XLSX from "xlsx";

type Tab = "overview" | "daily" | "summary" | "handover" | "count";
type IssueType = "疑似未扣早班销量" | "早班交夜班" | "次日接班" | "金额核对";
type CardTone = "default" | "danger";

type SmokeRow = {
  date: string;
  sheetName: string;
  rowIndex: number;
  rowNumber: number;
  no: number | string;
  name: string;
  price: number;
  purchase: number;
  morningStock: number;
  morningSold: number;
  morningAmount: number;
  nightStock: number;
  nightSold: number;
  nightAmount: number;
  totalAmount: number;
  endingStock: number;
  sameDayExpectedNightStock: number;
  sameDayStockDiff: number;
  morningAmountDiff: number;
  nightAmountDiff: number;
  totalAmountDiff: number;
  hasMorningStock: boolean;
  hasMorningSold: boolean;
  hasNightStock: boolean;
  hasNightSold: boolean;
  hasPurchase: boolean;
  hasMorningAmount: boolean;
  hasNightAmount: boolean;
  hasTotalAmount: boolean;
};

type DailySummary = {
  date: string;
  purchase: number;
  morningSold: number;
  nightSold: number;
  totalSold: number;
  morningAmount: number;
  nightAmount: number;
  totalAmount: number;
  abnormalCount: number;
};

type ProductSummary = {
  name: string;
  price: number;
  purchase: number;
  morningSold: number;
  nightSold: number;
  totalSold: number;
  totalAmount: number;
  endingStock: number;
};

type WriteBackTarget = {
  key: string;
  date: string;
  sheetName: string;
  rowIndex: number;
  rowNumber: number;
  colIndex: number;
  cell: string;
  columnName: string;
  name: string;
  price: number;
  before: number;
  after: number;
};

type HandoverIssue = {
  id: string;
  type: IssueType;
  date: string;
  nextDate?: string;
  name: string;
  expected: number;
  actual: number;
  diff: number;
  note: string;
  writeBack?: WriteBackTarget;
};

type CorrectionRecord = WriteBackTarget & {
  diff: number;
  reason: string;
  correctedAt: string;
};

type IssueStats = {
  totalIssueCount: number;
  inventoryIssueCount: number;
  amountIssueCount: number;
  writableCount: number;
  preparedCount: number;
};

type CellWithStyle = XLSX.CellObject & {
  s?: Record<string, unknown>;
  z?: string;
  f?: string;
  w?: string;
};

const nav: Array<[Tab, string]> = [["overview", "总览"], ["daily", "每日烟报"], ["summary", "商品月汇总"], ["handover", "交接核对"], ["count", "现场点烟"]];
const STORE_NAME = "SY11";
const MORNING_STOCK_COLUMN = 4; // E列：早班库存
const NIGHT_STOCK_COLUMN = 7; // H列：夜班库存
const HEADER_STYLE = { font: { bold: true }, alignment: { horizontal: "center", vertical: "center" }, fill: { fgColor: { rgb: "E2E8F0" } } };
const TITLE_STYLE = { font: { bold: true, sz: 16 }, alignment: { horizontal: "center", vertical: "center" } };
const TOTAL_STYLE = { font: { bold: true }, fill: { fgColor: { rgb: "FEF3C7" } } };

function hasValue(value: unknown) { return value !== null && value !== undefined && String(value).trim() !== ""; }
function n(value: unknown) { if (typeof value === "number") return Number.isFinite(value) ? value : 0; const text = String(value ?? "").replace(/,/g, "").replace(/[￥¥元\s]/g, "").trim(); if (!text) return 0; const parsed = Number(text); return Number.isFinite(parsed) ? parsed : 0; }
function round2(value: number) { return Math.round((value + Number.EPSILON) * 100) / 100; }
function money(value: number) { return round2(value).toFixed(2); }
function rowKey(row: SmokeRow, date = row.date) { return `${date}__${row.sheetName}__${row.rowNumber}__${row.name}__${row.price}`; }
function correctionKey(sheetName: string, rowIndex: number, colIndex: number) { return `${sheetName}__${rowIndex}__${colIndex}`; }
function rowCellKey(row: SmokeRow, colIndex: number) { return correctionKey(row.sheetName, row.rowIndex, colIndex); }
function isMissedMorningDeduction(row: SmokeRow) { return row.hasNightStock && row.morningSold > 0 && Math.abs(row.sameDayStockDiff - row.morningSold) < 0.01; }
function correctedNightStock(row: SmokeRow) { return row.sameDayExpectedNightStock; }
function safeFileName(name: string) { return name.replace(/\.xlsx?$/i, "").replace(/[\\/:*?\"<>|]/g, "-"); }
function reportDateRange(daily: DailySummary[]) { if (!daily.length) return ""; return daily.length === 1 ? daily[0].date : `${daily[0].date} 至 ${daily[daily.length - 1].date}`; }
function queryTime() { return new Date().toLocaleString("zh-CN", { hour12: false }); }

function extractDate(sheetName: string, title: unknown, fallbackIndex: number) {
  const raw = `${String(title ?? "")} ${sheetName}`;
  const full = raw.match(/(20\d{2}|19\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?/);
  if (full) return `${full[1]}-${full[2].padStart(2, "0")}-${full[3].padStart(2, "0")}`;
  const yearMonth = raw.match(/(20\d{2}|19\d{2})\s*年\s*(\d{1,2})\s*月/);
  const numericSheet = sheetName.match(/^(\d{1,2})$/);
  const day = numericSheet ? Number(numericSheet[1]) : fallbackIndex + 1;
  const year = yearMonth ? yearMonth[1] : String(new Date().getFullYear());
  const month = yearMonth ? yearMonth[2] : String(new Date().getMonth() + 1);
  return `${year}-${month.padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function parseWorkbook(workbook: XLSX.WorkBook) {
  const parsed: SmokeRow[] = [];
  workbook.SheetNames.forEach((sheetName, sheetIndex) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;
    const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null });
    const date = extractDate(sheetName, grid[0]?.[0], sheetIndex);

    for (let i = 3; i < Math.min(grid.length, 90); i += 1) {
      const row = grid[i] ?? [];
      const name = String(row[1] ?? "").trim();
      const price = n(row[2]);
      if (!name || !price) continue;
      if (name.includes("合计") || name.includes("收款")) continue;

      const purchaseRaw = row[3], morningStockRaw = row[4], morningSoldRaw = row[5], morningAmountRaw = row[6], nightStockRaw = row[7], nightSoldRaw = row[8], nightAmountRaw = row[9], totalAmountRaw = row[10];
      const purchase = n(purchaseRaw), morningStock = n(morningStockRaw), morningSold = n(morningSoldRaw), nightStock = n(nightStockRaw), nightSold = n(nightSoldRaw);
      const hasPurchase = hasValue(purchaseRaw), hasMorningStock = hasValue(morningStockRaw), hasMorningSold = hasValue(morningSoldRaw), hasNightStock = hasValue(nightStockRaw), hasNightSold = hasValue(nightSoldRaw);
      const hasMorningAmount = hasValue(morningAmountRaw) && n(morningAmountRaw) !== 0, hasNightAmount = hasValue(nightAmountRaw) && n(nightAmountRaw) !== 0, hasTotalAmount = hasValue(totalAmountRaw) && n(totalAmountRaw) !== 0;
      const active = hasPurchase || hasMorningStock || hasMorningSold || hasNightStock || hasNightSold || hasMorningAmount || hasNightAmount || hasTotalAmount;
      if (!active) continue;

      const morningAmount = n(morningAmountRaw) || round2(morningSold * price);
      const nightAmount = n(nightAmountRaw) || round2(nightSold * price);
      const totalAmount = n(totalAmountRaw) || round2(morningAmount + nightAmount);
      const sameDayExpectedNightStock = round2(morningStock + purchase - morningSold);
      const endingStock = hasNightStock ? round2(nightStock - nightSold) : sameDayExpectedNightStock;

      parsed.push({ date, sheetName, rowIndex: i, rowNumber: i + 1, no: row[0] as number | string, name, price, purchase, morningStock, morningSold, morningAmount, nightStock, nightSold, nightAmount, totalAmount, endingStock, sameDayExpectedNightStock, sameDayStockDiff: hasNightStock ? round2(nightStock - sameDayExpectedNightStock) : 0, morningAmountDiff: hasMorningSold || hasMorningAmount ? round2(morningAmount - morningSold * price) : 0, nightAmountDiff: hasNightSold || hasNightAmount ? round2(nightAmount - nightSold * price) : 0, totalAmountDiff: hasMorningAmount || hasNightAmount || hasTotalAmount ? round2(totalAmount - morningAmount - nightAmount) : 0, hasMorningStock, hasMorningSold, hasNightStock, hasNightSold, hasPurchase, hasMorningAmount, hasNightAmount, hasTotalAmount });
    }
  });
  return parsed.sort((a, b) => `${a.date}-${a.rowNumber}`.localeCompare(`${b.date}-${b.rowNumber}`));
}

function hasRowIssue(row: SmokeRow) { return (row.hasNightStock && Math.abs(row.sameDayStockDiff) > 0.01) || Math.abs(row.morningAmountDiff) > 0.01 || Math.abs(row.nightAmountDiff) > 0.01 || Math.abs(row.totalAmountDiff) > 0.01; }
function columnName(colIndex: number) { return colIndex === MORNING_STOCK_COLUMN ? "E列 早班库存" : "H列 夜班库存"; }
function makeWriteBack(row: SmokeRow, colIndex: number, before: number, after: number): WriteBackTarget {
  const cell = XLSX.utils.encode_cell({ r: row.rowIndex, c: colIndex });
  return { key: rowCellKey(row, colIndex), date: row.date, sheetName: row.sheetName, rowIndex: row.rowIndex, rowNumber: row.rowNumber, colIndex, cell, columnName: columnName(colIndex), name: row.name, price: row.price, before, after: round2(after) };
}
function writeCell(sheet: XLSX.WorkSheet, rowIndex: number, colIndex: number, value: number) { const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex }); const cell = (sheet[address] ?? {}) as CellWithStyle; delete cell.f; delete cell.w; cell.t = "n"; cell.v = value; sheet[address] = cell; return address; }
function downloadWorkbook(fileName: string, sheets: Record<string, Record<string, unknown>[]>) { const workbook = XLSX.utils.book_new(); Object.entries(sheets).forEach(([name, rows]) => XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), name.slice(0, 31))); XLSX.writeFile(workbook, fileName); }
function setCellStyle(sheet: XLSX.WorkSheet, row: number, col: number, style: Record<string, unknown>) { const address = XLSX.utils.encode_cell({ r: row, c: col }); const cell = sheet[address] as CellWithStyle | undefined; if (cell) cell.s = { ...(cell.s ?? {}), ...style }; }
function setNumFormat(sheet: XLSX.WorkSheet, row: number, col: number, format: string) { const address = XLSX.utils.encode_cell({ r: row, c: col }); const cell = sheet[address] as CellWithStyle | undefined; if (cell) cell.z = format; }
function styleRow(sheet: XLSX.WorkSheet, row: number, colCount: number, style: Record<string, unknown>) { for (let col = 0; col < colCount; col += 1) setCellStyle(sheet, row, col, style); }
function writeBackLocation(issue: HandoverIssue) { return issue.writeBack ? `${issue.writeBack.sheetName}!${issue.writeBack.cell}（${issue.writeBack.columnName}）` : "不自动写回"; }

function Card({ title, value, desc, tone = "default" }: { title: string; value: string; desc?: string; tone?: CardTone }) {
  const cls = tone === "danger" ? "border-red-500 bg-red-600 text-white shadow-sm" : "border bg-white text-slate-900 shadow-sm";
  const titleCls = tone === "danger" ? "text-red-100" : "text-slate-500";
  const descCls = tone === "danger" ? "text-red-100" : "text-slate-500";
  return <div className={`rounded-2xl p-5 ${cls}`}><div className={`text-sm ${titleCls}`}>{title}</div><div className="mt-2 text-2xl font-bold">{value}</div>{desc ? <div className={`mt-1 text-xs ${descCls}`}>{desc}</div> : null}</div>;
}

function StatsCards({ stats }: { stats: IssueStats }) {
  return <div className="grid gap-4 md:grid-cols-5"><Card tone="danger" title="总异常数" value={`${stats.totalIssueCount} 条`} desc="库存 + 金额" /><Card title="库存异常数" value={`${stats.inventoryIssueCount} 条`} /><Card title="金额异常数" value={`${stats.amountIssueCount} 条`} /><Card title="可写回原表数量" value={`${stats.writableCount} 条`} desc="只含库存类异常" /><Card title="已准备写回数量" value={`${stats.preparedCount} 条`} /></div>;
}

function Badge({ children, tone = "gray" }: { children: ReactNode; tone?: "gray" | "green" | "red" | "yellow" }) { const cls = { gray: "bg-slate-100 text-slate-700", green: "bg-green-100 text-green-700", red: "bg-red-100 text-red-700", yellow: "bg-yellow-100 text-yellow-800" }[tone]; return <span className={`rounded-full px-2 py-1 text-xs font-medium ${cls}`}>{children}</span>; }

export default function Home() {
  const [tab, setTab] = useState<Tab>("overview");
  const [fileName, setFileName] = useState("");
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [rows, setRows] = useState<SmokeRow[]>([]);
  const [countInput, setCountInput] = useState<Record<string, string>>({});
  const [corrections, setCorrections] = useState<Record<string, CorrectionRecord>>({});

  async function handleFile(file: File) {
    const buffer = await file.arrayBuffer();
    const nextWorkbook = XLSX.read(buffer, { type: "array", cellDates: true, cellFormula: true, cellNF: true, cellStyles: true });
    setWorkbook(nextWorkbook);
    setRows(parseWorkbook(nextWorkbook));
    setFileName(file.name);
    setCountInput({});
    setCorrections({});
  }

  const daily = useMemo<DailySummary[]>(() => {
    const map = new Map<string, DailySummary>();
    rows.forEach((row) => { const item = map.get(row.date) ?? { date: row.date, purchase: 0, morningSold: 0, nightSold: 0, totalSold: 0, morningAmount: 0, nightAmount: 0, totalAmount: 0, abnormalCount: 0 }; item.purchase += row.purchase; item.morningSold += row.morningSold; item.nightSold += row.nightSold; item.totalSold += row.morningSold + row.nightSold; item.morningAmount += row.morningAmount; item.nightAmount += row.nightAmount; item.totalAmount += row.totalAmount; item.abnormalCount += hasRowIssue(row) ? 1 : 0; map.set(row.date, item); });
    return Array.from(map.values()).map((item) => ({ ...item, purchase: round2(item.purchase), morningSold: round2(item.morningSold), nightSold: round2(item.nightSold), totalSold: round2(item.totalSold), morningAmount: round2(item.morningAmount), nightAmount: round2(item.nightAmount), totalAmount: round2(item.totalAmount) })).sort((a, b) => a.date.localeCompare(b.date));
  }, [rows]);

  const productSummary = useMemo<ProductSummary[]>(() => {
    const map = new Map<string, ProductSummary>();
    rows.forEach((row) => { const key = `${row.name}__${row.price}`; const item = map.get(key) ?? { name: row.name, price: row.price, purchase: 0, morningSold: 0, nightSold: 0, totalSold: 0, totalAmount: 0, endingStock: 0 }; item.purchase += row.purchase; item.morningSold += row.morningSold; item.nightSold += row.nightSold; item.totalSold += row.morningSold + row.nightSold; item.totalAmount += row.totalAmount; item.endingStock = row.endingStock; map.set(key, item); });
    return Array.from(map.values()).map((item) => ({ ...item, purchase: round2(item.purchase), morningSold: round2(item.morningSold), nightSold: round2(item.nightSold), totalSold: round2(item.totalSold), totalAmount: round2(item.totalAmount), endingStock: round2(item.endingStock) })).sort((a, b) => b.totalAmount - a.totalAmount);
  }, [rows]);

  const handovers = useMemo<HandoverIssue[]>(() => {
    const issues: HandoverIssue[] = [];
    rows.forEach((row) => {
      if (row.hasNightStock && Math.abs(row.sameDayStockDiff) > 0.01) {
        const type: IssueType = isMissedMorningDeduction(row) ? "疑似未扣早班销量" : "早班交夜班";
        const note = type === "疑似未扣早班销量" ? "夜班库存比正确值多出的数量，正好等于早班销量。高度疑似交班时忘记扣早班售卖数量。" : "夜班库存应等于：早班库存 + 进货 - 早班售卖数量。";
        const writeBack = makeWriteBack(row, NIGHT_STOCK_COLUMN, row.nightStock, correctedNightStock(row));
        issues.push({ id: `${type}__${writeBack.key}`, type, date: row.date, name: row.name, expected: row.sameDayExpectedNightStock, actual: row.nightStock, diff: row.sameDayStockDiff, note, writeBack });
      }
      if (Math.abs(row.morningAmountDiff) > 0.01) issues.push({ id: `金额核对__早班__${rowKey(row)}`, type: "金额核对", date: row.date, name: row.name, expected: round2(row.morningSold * row.price), actual: row.morningAmount, diff: row.morningAmountDiff, note: "早班金额不等于：早班售卖数量 × 价格。金额核对异常只提示，不自动改原表。" });
      if (Math.abs(row.nightAmountDiff) > 0.01) issues.push({ id: `金额核对__夜班__${rowKey(row)}`, type: "金额核对", date: row.date, name: row.name, expected: round2(row.nightSold * row.price), actual: row.nightAmount, diff: row.nightAmountDiff, note: "夜班金额不等于：夜班售卖数量 × 价格。金额核对异常只提示，不自动改原表。" });
      if (Math.abs(row.totalAmountDiff) > 0.01) issues.push({ id: `金额核对__合计__${rowKey(row)}`, type: "金额核对", date: row.date, name: row.name, expected: round2(row.morningAmount + row.nightAmount), actual: row.totalAmount, diff: row.totalAmountDiff, note: "合计金额不等于：早班金额 + 夜班金额。金额核对异常只提示，不自动改原表。" });
    });

    const byProduct = new Map<string, SmokeRow[]>();
    rows.forEach((row) => { const key = `${row.name}__${row.price}`; const list = byProduct.get(key) ?? []; list.push(row); byProduct.set(key, list); });
    byProduct.forEach((list) => {
      const sorted = [...list].sort((a, b) => a.date.localeCompare(b.date));
      for (let i = 0; i < sorted.length - 1; i += 1) {
        const current = sorted[i], next = sorted[i + 1];
        if (!current.hasNightStock || !next.hasMorningStock) continue;
        const expected = round2(current.nightStock - current.nightSold);
        const diff = round2(next.morningStock - expected);
        if (Math.abs(diff) > 0.01) {
          const writeBack = makeWriteBack(next, MORNING_STOCK_COLUMN, next.morningStock, expected);
          issues.push({ id: `次日接班__${writeBack.key}__${current.date}`, type: "次日接班", date: current.date, nextDate: next.date, name: current.name, expected, actual: next.morningStock, diff, note: "次日早班库存应等于：前一日夜班库存 - 前一日夜班售卖数量。", writeBack });
        }
      }
    });
    return issues.sort((a, b) => `${a.date}-${a.name}-${a.type}`.localeCompare(`${b.date}-${b.name}-${b.type}`));
  }, [rows]);

  const latestDate = daily.length ? daily[daily.length - 1].date : "";
  const latestRows = rows.filter((row) => row.date === latestDate);
  const missedMorningRows = rows.filter(isMissedMorningDeduction);
  const latestMissedMorningRows = latestRows.filter(isMissedMorningDeduction);
  const totalAmount = daily.reduce((sum, item) => sum + item.totalAmount, 0);
  const totalSold = daily.reduce((sum, item) => sum + item.totalSold, 0);
  const correctionRows = Object.values(corrections).sort((a, b) => `${a.date}-${a.name}-${a.cell}`.localeCompare(`${b.date}-${b.name}-${b.cell}`));
  const stats = useMemo<IssueStats>(() => {
    const inventoryIssueCount = handovers.filter((item) => item.type !== "金额核对").length;
    const amountIssueCount = handovers.filter((item) => item.type === "金额核对").length;
    const writableCount = handovers.filter((item) => item.writeBack).length;
    return { totalIssueCount: handovers.length, inventoryIssueCount, amountIssueCount, writableCount, preparedCount: correctionRows.length };
  }, [handovers, correctionRows.length]);

  function makeCorrection(target: WriteBackTarget, reason: string): CorrectionRecord {
    return { ...target, after: round2(target.after), diff: round2(target.after - target.before), reason, correctedAt: new Date().toLocaleString("zh-CN", { hour12: false }) };
  }
  function setCorrection(target: WriteBackTarget, reason: string, old: Record<string, CorrectionRecord>) { return { ...old, [target.key]: makeCorrection(target, reason) }; }
  function correctedExpectedEndingStock(row: SmokeRow) { const correction = corrections[rowCellKey(row, NIGHT_STOCK_COLUMN)]; return correction ? round2(correction.after - row.nightSold) : row.endingStock; }
  function addOnsiteCorrection(row: SmokeRow, actualEndingStock: number) { const fixedNightStock = round2(actualEndingStock + row.nightSold); const target = makeWriteBack(row, NIGHT_STOCK_COLUMN, row.nightStock, fixedNightStock); setCorrections((old) => setCorrection(target, "按现场实点反推夜班库存，写回H列夜班库存", old)); setCountInput((old) => ({ ...old, [row.name]: String(actualEndingStock) })); }
  function applyIssueCorrection(issue: HandoverIssue) { if (!issue.writeBack) return alert("金额核对异常只提示，不自动改原表。"); setCorrections((old) => setCorrection(issue.writeBack as WriteBackTarget, issue.note, old)); }
  function undoCorrectionByKey(key: string) { setCorrections((old) => { const copy = { ...old }; delete copy[key]; return copy; }); }
  function undoCorrection(row: SmokeRow) { undoCorrectionByKey(rowCellKey(row, NIGHT_STOCK_COLUMN)); }
  function applyWritableHandoverCorrections(scope: "all" | "missedMorning") {
    const writable = handovers.filter((issue) => issue.writeBack && (scope === "all" || issue.type === "疑似未扣早班销量") && !corrections[issue.writeBack.key]);
    if (!writable.length) return alert("没有新的库存异常需要准备写回。");
    const ok = confirm(`确定准备写回 ${writable.length} 条库存异常吗？金额核对异常不会自动改原表。`);
    if (!ok) return;
    setCorrections((old) => writable.reduce((acc, issue) => setCorrection(issue.writeBack as WriteBackTarget, issue.note, acc), old));
  }

  function downloadCorrectedOriginal() {
    if (!workbook) return alert("请先上传原始烟报 Excel");
    const correctionList = Object.values(corrections);
    if (!correctionList.length) return alert("还没有纠正内容。请先在交接核对或现场点烟里准备写回。");
    const fixed = XLSX.read(XLSX.write(workbook, { bookType: "xlsx", type: "array" }), { type: "array", cellDates: true, cellFormula: true, cellNF: true, cellStyles: true });
    correctionList.forEach((item) => { const sheet = fixed.Sheets[item.sheetName]; if (sheet) writeCell(sheet, item.rowIndex, item.colIndex, item.after); });
    XLSX.writeFile(fixed, `已纠正-${safeFileName(fileName || "手工烟报")}.xlsx`);
  }

  function exportSalesSummaryReport() {
    if (!rows.length) return alert("请先上传原始烟报 Excel");
    const totalQty = round2(productSummary.reduce((sum, item) => sum + item.totalSold, 0));
    const totalMoney = round2(productSummary.reduce((sum, item) => sum + item.totalAmount, 0));
    const aoa: Array<Array<string | number>> = [
      ["SY11 香烟销售汇总报表", "", "", "", "", ""],
      [`门店：${STORE_NAME}`, "", `营业日范围：${reportDateRange(daily)}`, "", `最后查询时间：${queryTime()}`, ""],
      ["", "", "", "", "", ""],
      ["序号", "商品分类", "商品名称", "平均价格", "数量", "金额"],
      ...productSummary.map((item, index) => [index + 1, "香烟(酒店库存)", item.name, round2(item.price), round2(item.totalSold), round2(item.totalAmount)]),
      ["总计", "", "", "", totalQty, totalMoney]
    ];
    const sheet = XLSX.utils.aoa_to_sheet(aoa);
    const totalRow = aoa.length - 1;
    sheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }, { s: { r: 1, c: 2 }, e: { r: 1, c: 3 } }, { s: { r: 1, c: 4 }, e: { r: 1, c: 5 } }];
    sheet["!cols"] = [{ wch: 8 }, { wch: 20 }, { wch: 32 }, { wch: 12 }, { wch: 12 }, { wch: 14 }];
    setCellStyle(sheet, 0, 0, TITLE_STYLE);
    styleRow(sheet, 3, 6, HEADER_STYLE);
    styleRow(sheet, totalRow, 6, TOTAL_STYLE);
    for (let row = 4; row <= totalRow; row += 1) { setNumFormat(sheet, row, 3, "0.00"); setNumFormat(sheet, row, 5, "0.00"); }
    const reportWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(reportWorkbook, sheet, "销售汇总报表");
    XLSX.writeFile(reportWorkbook, `${STORE_NAME}-香烟销售汇总报表.xlsx`);
  }

  function exportExceptionReport() {
    if (!rows.length) return alert("请先上传原始烟报 Excel");
    const aoa: Array<Array<string | number>> = [
      ["烟报异常核对报表", "", "", "", "", "", "", "", "", ""],
      [`门店：${STORE_NAME}`, `日期范围：${reportDateRange(daily)}`, `总异常数：${stats.totalIssueCount}`, `库存异常数：${stats.inventoryIssueCount}`, `金额异常数：${stats.amountIssueCount}`, "", "", "", "", ""],
      ["", "", "", "", "", "", "", "", "", ""],
      ["序号", "异常类型", "日期", "次日", "商品名称", "应为", "实际", "差异", "写回位置", "说明"],
      ...handovers.map((item, index) => [index + 1, item.type, item.date, item.nextDate ?? "", item.name, item.expected, item.actual, item.diff, writeBackLocation(item), item.note]),
      ["总计", `共 ${stats.totalIssueCount} 条`, "", "", `库存异常 ${stats.inventoryIssueCount} 条`, `金额异常 ${stats.amountIssueCount} 条`, "", "", `可写回 ${stats.writableCount} 条`, ""]
    ];
    const sheet = XLSX.utils.aoa_to_sheet(aoa);
    const totalRow = aoa.length - 1;
    sheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 9 } }];
    sheet["!cols"] = [{ wch: 8 }, { wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 28 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 24 }, { wch: 52 }];
    setCellStyle(sheet, 0, 0, TITLE_STYLE);
    styleRow(sheet, 3, 10, HEADER_STYLE);
    styleRow(sheet, totalRow, 10, TOTAL_STYLE);
    for (let row = 4; row < totalRow; row += 1) { setNumFormat(sheet, row, 5, "0.00"); setNumFormat(sheet, row, 6, "0.00"); setNumFormat(sheet, row, 7, "0.00"); }
    const reportWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(reportWorkbook, sheet, "异常核对报表");
    XLSX.writeFile(reportWorkbook, "烟报异常核对报表.xlsx");
  }

  function exportOnsiteCount() { if (!latestRows.length) return; const countRows = latestRows.map((row) => { const actualText = countInput[row.name] ?? ""; const actual = actualText === "" ? "" : n(actualText); const expected = correctedExpectedEndingStock(row); const diff = actual === "" ? "" : round2(Number(actual) - expected); const correction = corrections[rowCellKey(row, NIGHT_STOCK_COLUMN)]; return { 日期: latestDate, 商品名称: row.name, 价格: row.price, 原应剩数量: row.endingStock, 当前应剩数量: expected, 现场实点: actual, 差异: diff, 是否已准备写回: correction ? "是" : "否", 写回位置: correction ? `${correction.sheetName}!${correction.cell}` : "" }; }); downloadWorkbook(`${latestDate}-现场点烟表.xlsx`, { 现场点烟: countRows, 差异商品: countRows.filter((row) => row.差异 !== "" && row.差异 !== 0) }); }

  return (
    <main className="min-h-screen p-4 md:p-8"><div className="mx-auto max-w-7xl">
      <header className="mb-6 rounded-3xl bg-slate-900 p-6 text-white shadow-sm"><p className="text-sm text-slate-300">Cigarette Report Reconciliation</p><h1 className="mt-1 text-2xl font-bold md:text-4xl">烟报自动核对助手</h1><p className="mt-2 max-w-3xl text-sm text-slate-300">按你的原始烟报模板解析、核对，并支持把库存类异常写回原表对应单元格后下载。</p></header>
      <section className="mb-6 rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">导入手工烟报 Excel</h2><p className="mt-1 text-sm text-slate-500">上传原始烟报后，系统会解析香烟区域。下载已纠正原表时，只修改原工作簿对应单元格，不新增纠正记录 Sheet。</p></div><label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-slate-900 px-4 py-3 text-sm font-medium text-white">选择烟报文件<input className="hidden" type="file" accept=".xlsx,.xls" onChange={(event: ChangeEvent<HTMLInputElement>) => event.target.files?.[0] && handleFile(event.target.files[0])} /></label></div>{fileName ? <p className="mt-3 text-sm text-slate-600">已导入：<b>{fileName}</b>，共解析 {rows.length} 条有效香烟明细。</p> : null}</section>
      <nav className="mb-6 flex flex-wrap gap-2">{nav.map(([key, label]) => <button key={key} onClick={() => setTab(key)} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === key ? "bg-slate-900 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>{label}</button>)}</nav>
      {!rows.length ? <section className="rounded-2xl border bg-white p-6 shadow-sm"><h2 className="text-lg font-bold">等待导入烟报</h2><p className="mt-2 text-sm text-slate-600">请先上传手工烟报 Excel。导入后会自动显示汇总、异常和原表纠正功能。</p></section> : null}
      {rows.length > 0 && tab === "overview" ? <section className="space-y-6"><StatsCards stats={stats} /><div className="grid gap-4 md:grid-cols-4"><Card title="日期数量" value={`${daily.length} 天`} desc={`最新日期 ${latestDate}`} /><Card title="香烟品规" value={`${productSummary.length} 个`} /><Card title="月销量" value={`${round2(totalSold)} 包`} /><Card title="月销售金额" value={`¥${money(totalAmount)}`} /></div><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">导出</h2><p className="mt-1 text-sm text-slate-500">库存类异常可写回原表；金额核对异常只出现在提示和异常报表里，不自动改原表。</p></div><div className="flex flex-wrap gap-2"><button onClick={downloadCorrectedOriginal} className="rounded-xl bg-green-700 px-4 py-2 text-sm font-medium text-white">下载已纠正原表</button><button onClick={exportSalesSummaryReport} className="rounded-xl border px-4 py-2 text-sm font-medium">导出销售汇总报表</button><button onClick={exportExceptionReport} className="rounded-xl border px-4 py-2 text-sm font-medium">导出异常核对报表</button></div></div></div></section> : null}
      {rows.length > 0 && tab === "daily" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">每日烟报汇总</h2><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["日期", "进货", "早班销量", "夜班销量", "总销量", "早班金额", "夜班金额", "总金额", "异常"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{daily.map((item) => <tr key={item.date} className="border-t"><td className="p-3">{item.date}</td><td className="p-3">{item.purchase}</td><td className="p-3">{item.morningSold}</td><td className="p-3">{item.nightSold}</td><td className="p-3">{item.totalSold}</td><td className="p-3">¥{money(item.morningAmount)}</td><td className="p-3">¥{money(item.nightAmount)}</td><td className="p-3">¥{money(item.totalAmount)}</td><td className="p-3">{item.abnormalCount ? <Badge tone="red">{item.abnormalCount} 条</Badge> : <Badge tone="green">正常</Badge>}</td></tr>)}</tbody></table></div></section> : null}
      {rows.length > 0 && tab === "summary" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">商品月汇总</h2><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["商品", "价格", "月进货", "早班销量", "夜班销量", "月销量", "月销售金额", "月末结存"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{productSummary.map((item) => <tr key={`${item.name}-${item.price}`} className="border-t"><td className="p-3">{item.name}</td><td className="p-3">¥{money(item.price)}</td><td className="p-3">{item.purchase}</td><td className="p-3">{item.morningSold}</td><td className="p-3">{item.nightSold}</td><td className="p-3">{item.totalSold}</td><td className="p-3">¥{money(item.totalAmount)}</td><td className="p-3">{item.endingStock}</td></tr>)}</tbody></table></div></section> : null}
      {rows.length > 0 && tab === "handover" ? <section className="space-y-5"><StatsCards stats={stats} /><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">交接和金额异常</h2><p className="mt-1 text-sm text-slate-500">早班交夜班、疑似未扣早班销量写回当天 H列夜班库存；次日接班写回次日 E列早班库存；金额核对只提示。</p></div><div className="flex flex-wrap gap-2"><button onClick={() => applyWritableHandoverCorrections("all")} className="rounded-xl bg-yellow-700 px-4 py-2 text-sm text-white">准备写回全部库存异常</button><button onClick={() => applyWritableHandoverCorrections("missedMorning")} className="rounded-xl border px-4 py-2 text-sm">只准备疑似未扣</button><button onClick={downloadCorrectedOriginal} className="rounded-xl bg-green-700 px-4 py-2 text-sm text-white">下载已纠正原表</button></div></div><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["类型", "日期", "次日", "商品", "应为", "实际", "差异", "写回位置", "说明", "操作"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{handovers.map((item) => { const prepared = item.writeBack ? corrections[item.writeBack.key] : null; return <tr key={item.id} className="border-t"><td className="p-3"><Badge tone={item.type === "金额核对" ? "red" : item.type === "疑似未扣早班销量" ? "yellow" : "red"}>{item.type}</Badge></td><td className="p-3">{item.date}</td><td className="p-3">{item.nextDate ?? ""}</td><td className="p-3">{item.name}</td><td className="p-3">{item.expected}</td><td className="p-3">{item.actual}</td><td className="p-3">{item.diff}</td><td className="p-3">{writeBackLocation(item)}</td><td className="p-3">{item.note}</td><td className="p-3">{item.writeBack ? prepared ? <button className="rounded-lg border px-3 py-1 text-xs" onClick={() => undoCorrectionByKey(item.writeBack!.key)}>撤销</button> : <button className="rounded-lg bg-slate-900 px-3 py-1 text-xs text-white" onClick={() => applyIssueCorrection(item)}>准备写回</button> : <Badge>只提示</Badge>}</td></tr>; })}</tbody></table></div>{correctionRows.length ? <div className="mt-5 rounded-xl bg-yellow-50 p-4 text-sm text-yellow-900"><b>已准备写回原表 {correctionRows.length} 条：</b>{correctionRows.slice(0, 20).map((item) => <div key={item.key}>{item.sheetName}!{item.cell} {item.name}（{item.columnName}）：{item.before} → {item.after}</div>)}</div> : null}</div></section> : null}
      {rows.length > 0 && tab === "count" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">现场点烟清单</h2><p className="mt-1 text-sm text-slate-500">现场实点有差异时，可以按实点反推并修正原表 H列 夜班库存。</p></div><button onClick={exportOnsiteCount} className="rounded-xl border px-4 py-2 text-sm font-medium">导出现场点烟明细</button></div>{latestMissedMorningRows.length ? <div className="mt-4 rounded-xl bg-yellow-50 p-4 text-sm text-yellow-900"><b>最新日期检测到 {latestMissedMorningRows.length} 条疑似交班忘扣早班销量。</b><button className="ml-0 mt-3 rounded-lg bg-yellow-700 px-4 py-2 text-sm text-white md:ml-3 md:mt-0" onClick={() => applyWritableHandoverCorrections("missedMorning")}>准备写回疑似未扣</button></div> : null}<div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["商品", "价格", "当前应剩", "现场实点", "差异", "状态", "纠正原表"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{latestRows.map((row) => { const key = rowKey(row); const actualText = countInput[row.name] ?? ""; const actual = actualText === "" ? null : n(actualText); const expected = correctedExpectedEndingStock(row); const diff = actual === null ? null : round2(actual - expected); const correction = corrections[rowCellKey(row, NIGHT_STOCK_COLUMN)]; return <tr key={key} className="border-t"><td className="p-3">{row.name}{isMissedMorningDeduction(row) ? <div className="text-xs text-yellow-700">疑似未扣早班销量</div> : null}</td><td className="p-3">¥{money(row.price)}</td><td className="p-3 font-bold">{expected}{correction ? <div className="text-xs font-normal text-slate-500">写回夜班库存：{correction.before} → {correction.after}</div> : null}</td><td className="p-3"><input className="w-28 rounded-xl border p-2" inputMode="decimal" value={actualText} onChange={(event: ChangeEvent<HTMLInputElement>) => setCountInput({ ...countInput, [row.name]: event.target.value })} /></td><td className="p-3">{diff === null ? "" : diff}</td><td className="p-3">{correction ? <Badge tone="yellow">待写回原表</Badge> : diff === null ? <Badge>未点</Badge> : diff === 0 ? <Badge tone="green">正常</Badge> : <Badge tone="red">有差异</Badge>}</td><td className="p-3">{actual !== null && diff !== 0 ? <button className="rounded-lg bg-slate-900 px-3 py-1 text-xs text-white" onClick={() => addOnsiteCorrection(row, actual)}>按实点纠正H列</button> : null}{correction ? <button className="ml-2 rounded-lg border px-3 py-1 text-xs" onClick={() => undoCorrection(row)}>撤销</button> : null}</td></tr>; })}</tbody></table></div>{correctionRows.length ? <div className="mt-5 rounded-xl bg-yellow-50 p-4 text-sm text-yellow-900"><b>待写回原表 {correctionRows.length} 条。</b><button className="ml-3 rounded-lg bg-green-700 px-4 py-2 text-sm text-white" onClick={downloadCorrectedOriginal}>下载已纠正原表</button></div> : null}</section> : null}
    </div></main>
  );
}
