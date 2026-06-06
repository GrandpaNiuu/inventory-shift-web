"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type Tab = "overview" | "daily" | "summary" | "handover" | "count";

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

type DailySummary = { date: string; purchase: number; morningSold: number; nightSold: number; totalSold: number; morningAmount: number; nightAmount: number; totalAmount: number; abnormalCount: number };
type ProductSummary = { name: string; price: number; purchase: number; morningSold: number; nightSold: number; totalSold: number; totalAmount: number; endingStock: number };
type HandoverIssue = { type: "疑似未扣早班销量" | "早班交夜班" | "次日接班" | "金额核对"; date: string; nextDate?: string; name: string; expected: number; actual: number; diff: number; note: string };
type CorrectionRecord = { date: string; sheetName: string; rowNumber: number; cell: string; name: string; price: number; before: number; after: number; diff: number; reason: string; correctedAt: string };

const nav: Array<[Tab, string]> = [["overview", "总览"], ["daily", "每日烟报"], ["summary", "商品月汇总"], ["handover", "交接核对"], ["count", "现场点烟"]];
const NIGHT_STOCK_COLUMN = 7; // H列：夜班库存

function hasValue(value: unknown) { return value !== null && value !== undefined && String(value).trim() !== ""; }
function n(value: unknown) { if (typeof value === "number") return Number.isFinite(value) ? value : 0; const text = String(value ?? "").replace(/,/g, "").replace(/[￥¥元\s]/g, "").trim(); if (!text) return 0; const parsed = Number(text); return Number.isFinite(parsed) ? parsed : 0; }
function round2(value: number) { return Math.round((value + Number.EPSILON) * 100) / 100; }
function money(value: number) { return round2(value).toFixed(2); }
function rowKey(row: SmokeRow, date = row.date) { return `${date}__${row.sheetName}__${row.rowNumber}__${row.name}__${row.price}`; }
function isMissedMorningDeduction(row: SmokeRow) { return row.hasNightStock && row.morningSold > 0 && Math.abs(row.sameDayStockDiff - row.morningSold) < 0.01; }
function correctedNightStock(row: SmokeRow) { return row.sameDayExpectedNightStock; }
function correctedEndingStock(row: SmokeRow) { return round2(correctedNightStock(row) - row.nightSold); }
function safeFileName(name: string) { return name.replace(/\.xlsx?$/i, "").replace(/[\\/:*?\"<>|]/g, "-"); }

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
function writeCell(sheet: XLSX.WorkSheet, rowIndex: number, colIndex: number, value: number) { const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex }); sheet[address] = { t: "n", v: value }; return address; }
function downloadWorkbook(fileName: string, sheets: Record<string, Record<string, unknown>[]>) { const workbook = XLSX.utils.book_new(); Object.entries(sheets).forEach(([name, rows]) => XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), name.slice(0, 31))); XLSX.writeFile(workbook, fileName); }
function Card({ title, value, desc }: { title: string; value: string; desc?: string }) { return <div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="text-sm text-slate-500">{title}</div><div className="mt-2 text-2xl font-bold text-slate-900">{value}</div>{desc ? <div className="mt-1 text-xs text-slate-500">{desc}</div> : null}</div>; }
function Badge({ children, tone = "gray" }: { children: React.ReactNode; tone?: "gray" | "green" | "red" | "yellow" }) { const cls = { gray: "bg-slate-100 text-slate-700", green: "bg-green-100 text-green-700", red: "bg-red-100 text-red-700", yellow: "bg-yellow-100 text-yellow-800" }[tone]; return <span className={`rounded-full px-2 py-1 text-xs font-medium ${cls}`}>{children}</span>; }

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
    rows.forEach((row) => { if (row.hasNightStock && Math.abs(row.sameDayStockDiff) > 0.01) { if (isMissedMorningDeduction(row)) issues.push({ type: "疑似未扣早班销量", date: row.date, name: row.name, expected: row.sameDayExpectedNightStock, actual: row.nightStock, diff: row.sameDayStockDiff, note: "夜班库存比正确值多出的数量，正好等于早班销量。高度疑似交班时忘记扣早班售卖数量。" }); else issues.push({ type: "早班交夜班", date: row.date, name: row.name, expected: row.sameDayExpectedNightStock, actual: row.nightStock, diff: row.sameDayStockDiff, note: "夜班库存应等于：早班库存 + 进货 - 早班售卖数量。" }); } if (Math.abs(row.morningAmountDiff) > 0.01) issues.push({ type: "金额核对", date: row.date, name: row.name, expected: round2(row.morningSold * row.price), actual: row.morningAmount, diff: row.morningAmountDiff, note: "早班金额不等于：早班售卖数量 × 价格" }); if (Math.abs(row.nightAmountDiff) > 0.01) issues.push({ type: "金额核对", date: row.date, name: row.name, expected: round2(row.nightSold * row.price), actual: row.nightAmount, diff: row.nightAmountDiff, note: "夜班金额不等于：夜班售卖数量 × 价格" }); if (Math.abs(row.totalAmountDiff) > 0.01) issues.push({ type: "金额核对", date: row.date, name: row.name, expected: round2(row.morningAmount + row.nightAmount), actual: row.totalAmount, diff: row.totalAmountDiff, note: "合计金额不等于：早班金额 + 夜班金额" }); });
    const byProduct = new Map<string, SmokeRow[]>();
    rows.forEach((row) => { const list = byProduct.get(row.name) ?? []; list.push(row); byProduct.set(row.name, list); });
    byProduct.forEach((list) => { const sorted = [...list].sort((a, b) => a.date.localeCompare(b.date)); for (let i = 0; i < sorted.length - 1; i += 1) { const current = sorted[i], next = sorted[i + 1]; if (!current.hasNightStock || !next.hasMorningStock) continue; const diff = round2(next.morningStock - current.endingStock); if (Math.abs(diff) > 0.01) issues.push({ type: "次日接班", date: current.date, nextDate: next.date, name: current.name, expected: current.endingStock, actual: next.morningStock, diff, note: "次日早班库存应等于：前一日夜班库存 - 前一日夜班售卖数量" }); } });
    return issues.sort((a, b) => `${a.date}-${a.name}`.localeCompare(`${b.date}-${b.name}`));
  }, [rows]);

  const latestDate = daily.length ? daily[daily.length - 1].date : "";
  const latestRows = rows.filter((row) => row.date === latestDate);
  const missedMorningRows = rows.filter(isMissedMorningDeduction);
  const latestMissedMorningRows = latestRows.filter(isMissedMorningDeduction);
  const totalAmount = daily.reduce((sum, item) => sum + item.totalAmount, 0);
  const totalSold = daily.reduce((sum, item) => sum + item.totalSold, 0);
  const correctionRows = Object.values(corrections).sort((a, b) => `${a.date}-${a.name}`.localeCompare(`${b.date}-${b.name}`));

  function correctedExpected(row: SmokeRow) { return corrections[rowKey(row)]?.after ?? row.endingStock; }
  function addCorrection(row: SmokeRow, after: number, reason: string, old: Record<string, CorrectionRecord>) { const cell = XLSX.utils.encode_cell({ r: row.rowIndex, c: NIGHT_STOCK_COLUMN }); return { ...old, [rowKey(row)]: { date: row.date, sheetName: row.sheetName, rowNumber: row.rowNumber, cell, name: row.name, price: row.price, before: row.nightStock, after, diff: round2(after - row.nightStock), reason, correctedAt: new Date().toLocaleString() } }; }
  function applyCorrection(row: SmokeRow, actualEndingStock: number) { const fixedNightStock = round2(actualEndingStock + row.nightSold); setCorrections((old) => addCorrection(row, fixedNightStock, "按现场实点反推夜班库存", old)); setCountInput((old) => ({ ...old, [row.name]: String(actualEndingStock) })); }
  function applyMissedMorningBatchFix(scope: "latest" | "all") { const target = scope === "latest" ? latestMissedMorningRows : missedMorningRows; if (!target.length) return; const ok = confirm(`检测到 ${target.length} 条疑似“交班忘扣早班销量”。确定在原表 H列 夜班库存 中批量修正吗？`); if (!ok) return; setCorrections((old) => target.reduce((acc, row) => addCorrection(row, correctedNightStock(row), "疑似交班未扣早班销量，修正H列夜班库存", acc), old)); }
  function undoCorrection(row: SmokeRow) { const key = rowKey(row); setCorrections((old) => { const copy = { ...old }; delete copy[key]; return copy; }); }

  function downloadCorrectedOriginal() {
    if (!workbook) return alert("请先上传原始烟报 Excel");
    const correctionList = Object.values(corrections);
    if (!correctionList.length) return alert("还没有纠正内容。请先在交接核对或现场点烟里执行纠正。");
    const fixed = XLSX.read(XLSX.write(workbook, { bookType: "xlsx", type: "array" }), { type: "array", cellDates: true, cellFormula: true, cellNF: true, cellStyles: true });
    correctionList.forEach((item) => { const sheet = fixed.Sheets[item.sheetName]; if (sheet) { const decoded = XLSX.utils.decode_cell(item.cell); writeCell(sheet, decoded.r, decoded.c, item.after); } });
    const logRows = correctionList.map((item) => ({ 日期: item.date, Sheet: item.sheetName, 行号: item.rowNumber, 修正单元格: item.cell, 商品名称: item.name, 价格: item.price, 原夜班库存: item.before, 修正后夜班库存: item.after, 修正差异: item.diff, 原因: item.reason, 修正时间: item.correctedAt }));
    const logName = "纠正记录";
    fixed.SheetNames = fixed.SheetNames.filter((name) => name !== logName);
    delete fixed.Sheets[logName];
    XLSX.utils.book_append_sheet(fixed, XLSX.utils.json_to_sheet(logRows), logName);
    XLSX.writeFile(fixed, `已纠正-${safeFileName(fileName || "手工烟报")}.xlsx`);
  }

  function exportAll() { downloadWorkbook("烟报自动核对汇总.xlsx", { 每日汇总: daily.map((item) => ({ 日期: item.date, 进货: item.purchase, 早班销量: item.morningSold, 夜班销量: item.nightSold, 总销量: item.totalSold, 早班金额: item.morningAmount, 夜班金额: item.nightAmount, 总金额: item.totalAmount, 当日异常数: item.abnormalCount })), 商品月汇总: productSummary.map((item) => ({ 商品名称: item.name, 价格: item.price, 月进货: item.purchase, 早班销量: item.morningSold, 夜班销量: item.nightSold, 月销量: item.totalSold, 月销售金额: item.totalAmount, 月末结存: item.endingStock })), 交接异常: handovers.map((item) => ({ 类型: item.type, 日期: item.date, 次日: item.nextDate ?? "", 商品名称: item.name, 应为: item.expected, 实际: item.actual, 差异: item.diff, 说明: item.note })), 纠正记录: correctionRows.length ? correctionRows.map((item) => ({ 日期: item.date, Sheet: item.sheetName, 行号: item.rowNumber, 单元格: item.cell, 商品名称: item.name, 原夜班库存: item.before, 修正后夜班库存: item.after, 原因: item.reason, 修正时间: item.correctedAt })) : [{ 说明: "暂无纠正记录" }] }); }
  function exportOnsiteCount() { if (!latestRows.length) return; const countRows = latestRows.map((row) => { const actualText = countInput[row.name] ?? ""; const actual = actualText === "" ? "" : n(actualText); const expected = correctedExpected(row); const diff = actual === "" ? "" : round2(Number(actual) - expected); const correction = corrections[rowKey(row)]; return { 日期: latestDate, 商品名称: row.name, 价格: row.price, 原应剩数量: row.endingStock, 当前应剩数量: expected, 现场实点: actual, 差异: diff, 是否已纠正: correction ? "是" : "否" }; }); downloadWorkbook(`${latestDate}-现场点烟表.xlsx`, { 现场点烟: countRows, 差异商品: countRows.filter((row) => row.差异 !== "" && row.差异 !== 0), 纠正记录: correctionRows.length ? correctionRows.map((item) => ({ 日期: item.date, Sheet: item.sheetName, 行号: item.rowNumber, 单元格: item.cell, 商品名称: item.name, 原夜班库存: item.before, 修正后夜班库存: item.after, 原因: item.reason, 修正时间: item.correctedAt })) : [{ 说明: "暂无纠正记录" }] }); }

  return (
    <main className="min-h-screen p-4 md:p-8"><div className="mx-auto max-w-7xl">
      <header className="mb-6 rounded-3xl bg-slate-900 p-6 text-white shadow-sm"><p className="text-sm text-slate-300">Cigarette Report Reconciliation</p><h1 className="mt-1 text-2xl font-bold md:text-4xl">烟报自动核对助手</h1><p className="mt-2 max-w-3xl text-sm text-slate-300">按你的原始烟报模板解析、核对，并支持直接纠正原表 H列 夜班库存 后下载。</p></header>
      <section className="mb-6 rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">导入手工烟报 Excel</h2><p className="mt-1 text-sm text-slate-500">上传原始烟报后，系统会解析香烟区域。纠正后可以下载“已纠正-原文件名.xlsx”，尽量保留原表结构。</p></div><label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-slate-900 px-4 py-3 text-sm font-medium text-white">选择烟报文件<input className="hidden" type="file" accept=".xlsx,.xls" onChange={(event) => event.target.files?.[0] && handleFile(event.target.files[0])} /></label></div>{fileName ? <p className="mt-3 text-sm text-slate-600">已导入：<b>{fileName}</b>，共解析 {rows.length} 条有效香烟明细。</p> : null}</section>
      <nav className="mb-6 flex flex-wrap gap-2">{nav.map(([key, label]) => <button key={key} onClick={() => setTab(key)} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === key ? "bg-slate-900 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>{label}</button>)}</nav>
      {!rows.length ? <section className="rounded-2xl border bg-white p-6 shadow-sm"><h2 className="text-lg font-bold">等待导入烟报</h2><p className="mt-2 text-sm text-slate-600">请先上传手工烟报 Excel。导入后会自动显示汇总、异常和原表纠正功能。</p></section> : null}
      {rows.length > 0 && tab === "overview" ? <section className="space-y-6"><div className="grid gap-4 md:grid-cols-4"><Card title="日期数量" value={`${daily.length} 天`} desc={`最新日期 ${latestDate}`} /><Card title="香烟品规" value={`${productSummary.length} 个`} /><Card title="月销量" value={`${round2(totalSold)} 包`} /><Card title="月销售金额" value={`¥${money(totalAmount)}`} desc={`异常 ${handovers.length} 条，原表纠正 ${correctionRows.length} 条`} /></div><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">导出</h2><p className="mt-1 text-sm text-slate-500">“下载已纠正原表”会在你的原烟报模板里改 H列 夜班库存，并附加纠正记录 Sheet。</p></div><div className="flex flex-wrap gap-2"><button onClick={downloadCorrectedOriginal} className="rounded-xl bg-green-700 px-4 py-2 text-sm font-medium text-white">下载已纠正原表</button><button onClick={exportAll} className="rounded-xl border px-4 py-2 text-sm font-medium">导出核对汇总</button></div></div></div></section> : null}
      {rows.length > 0 && tab === "daily" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">每日烟报汇总</h2><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["日期", "进货", "早班销量", "夜班销量", "总销量", "早班金额", "夜班金额", "总金额", "异常"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{daily.map((item) => <tr key={item.date} className="border-t"><td className="p-3">{item.date}</td><td className="p-3">{item.purchase}</td><td className="p-3">{item.morningSold}</td><td className="p-3">{item.nightSold}</td><td className="p-3">{item.totalSold}</td><td className="p-3">¥{money(item.morningAmount)}</td><td className="p-3">¥{money(item.nightAmount)}</td><td className="p-3">¥{money(item.totalAmount)}</td><td className="p-3">{item.abnormalCount ? <Badge tone="red">{item.abnormalCount} 条</Badge> : <Badge tone="green">正常</Badge>}</td></tr>)}</tbody></table></div></section> : null}
      {rows.length > 0 && tab === "summary" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">商品月汇总</h2><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["商品", "价格", "月进货", "早班销量", "夜班销量", "月销量", "月销售金额", "月末结存"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{productSummary.map((item) => <tr key={`${item.name}-${item.price}`} className="border-t"><td className="p-3">{item.name}</td><td className="p-3">¥{money(item.price)}</td><td className="p-3">{item.purchase}</td><td className="p-3">{item.morningSold}</td><td className="p-3">{item.nightSold}</td><td className="p-3">{item.totalSold}</td><td className="p-3">¥{money(item.totalAmount)}</td><td className="p-3">{item.endingStock}</td></tr>)}</tbody></table></div></section> : null}
      {rows.length > 0 && tab === "handover" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">交接和金额异常</h2><p className="mt-1 text-sm text-slate-500">类型为“疑似未扣早班销量”时，可一键修正原表 H列 夜班库存。</p></div><div className="flex flex-wrap gap-2"><button onClick={() => applyMissedMorningBatchFix("all")} className="rounded-xl bg-yellow-700 px-4 py-2 text-sm text-white">一键修正全部疑似未扣</button><button onClick={downloadCorrectedOriginal} className="rounded-xl bg-green-700 px-4 py-2 text-sm text-white">下载已纠正原表</button></div></div><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["类型", "日期", "次日", "商品", "应为", "实际", "差异", "说明"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{handovers.map((item, index) => <tr key={`${item.type}-${item.date}-${item.name}-${index}`} className="border-t"><td className="p-3"><Badge tone={item.type === "疑似未扣早班销量" ? "yellow" : "red"}>{item.type}</Badge></td><td className="p-3">{item.date}</td><td className="p-3">{item.nextDate ?? ""}</td><td className="p-3">{item.name}</td><td className="p-3">{item.expected}</td><td className="p-3">{item.actual}</td><td className="p-3">{item.diff}</td><td className="p-3">{item.note}</td></tr>)}</tbody></table></div>{correctionRows.length ? <div className="mt-5 rounded-xl bg-yellow-50 p-4 text-sm text-yellow-900"><b>已准备写回原表 {correctionRows.length} 条：</b>{correctionRows.slice(0, 20).map((item) => <div key={`${item.sheetName}-${item.cell}-${item.name}`}>{item.sheetName}!{item.cell} {item.name}：{item.before} → {item.after}</div>)}</div> : null}</section> : null}
      {rows.length > 0 && tab === "count" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">现场点烟清单</h2><p className="mt-1 text-sm text-slate-500">现场实点有差异时，可以按实点反推并修正原表 H列 夜班库存。</p></div><button onClick={exportOnsiteCount} className="rounded-xl border px-4 py-2 text-sm font-medium">导出现场点烟明细</button></div>{latestMissedMorningRows.length ? <div className="mt-4 rounded-xl bg-yellow-50 p-4 text-sm text-yellow-900"><b>最新日期检测到 {latestMissedMorningRows.length} 条疑似交班忘扣早班销量。</b><button className="ml-0 mt-3 rounded-lg bg-yellow-700 px-4 py-2 text-sm text-white md:ml-3 md:mt-0" onClick={() => applyMissedMorningBatchFix("latest")}>只修正最新日期</button></div> : null}<div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["商品", "价格", "当前应剩", "现场实点", "差异", "状态", "纠正原表"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{latestRows.map((row) => { const key = rowKey(row); const actualText = countInput[row.name] ?? ""; const actual = actualText === "" ? null : n(actualText); const expected = correctedExpected(row); const diff = actual === null ? null : round2(actual - expected); const correction = corrections[key]; return <tr key={key} className="border-t"><td className="p-3">{row.name}{isMissedMorningDeduction(row) ? <div className="text-xs text-yellow-700">疑似未扣早班销量</div> : null}</td><td className="p-3">¥{money(row.price)}</td><td className="p-3 font-bold">{expected}{correction ? <div className="text-xs font-normal text-slate-500">原夜班库存：{correction.before}</div> : null}</td><td className="p-3"><input className="w-28 rounded-xl border p-2" inputMode="decimal" value={actualText} onChange={(event) => setCountInput({ ...countInput, [row.name]: event.target.value })} /></td><td className="p-3">{diff === null ? "" : diff}</td><td className="p-3">{correction ? <Badge tone="yellow">待写回原表</Badge> : diff === null ? <Badge>未点</Badge> : diff === 0 ? <Badge tone="green">正常</Badge> : <Badge tone="red">有差异</Badge>}</td><td className="p-3">{actual !== null && diff !== 0 ? <button className="rounded-lg bg-slate-900 px-3 py-1 text-xs text-white" onClick={() => applyCorrection(row, actual)}>按实点纠正H列</button> : null}{correction ? <button className="ml-2 rounded-lg border px-3 py-1 text-xs" onClick={() => undoCorrection(row)}>撤销</button> : null}</td></tr>; })}</tbody></table></div>{correctionRows.length ? <div className="mt-5 rounded-xl bg-yellow-50 p-4 text-sm text-yellow-900"><b>待写回原表 {correctionRows.length} 条。</b><button className="ml-3 rounded-lg bg-green-700 px-4 py-2 text-sm text-white" onClick={downloadCorrectedOriginal}>下载已纠正原表</button></div> : null}</section> : null}
    </div></main>
  );
}
