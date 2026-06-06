"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type Tab = "overview" | "daily" | "summary" | "handover" | "count";

type SmokeRow = {
  date: string;
  sheetName: string;
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

type HandoverIssue = {
  type: "早班交夜班" | "次日接班" | "金额核对";
  date: string;
  nextDate?: string;
  name: string;
  expected: number;
  actual: number;
  diff: number;
  note: string;
};

type CorrectionRecord = {
  date: string;
  name: string;
  price: number;
  before: number;
  after: number;
  diff: number;
  correctedAt: string;
};

const nav: Array<[Tab, string]> = [["overview", "总览"], ["daily", "每日烟报"], ["summary", "商品月汇总"], ["handover", "交接核对"], ["count", "现场点烟"]];

function hasValue(value: unknown) {
  return value !== null && value !== undefined && String(value).trim() !== "";
}

function n(value: unknown) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const text = String(value ?? "").replace(/,/g, "").replace(/[￥¥元\s]/g, "").trim();
  if (!text) return 0;
  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : 0;
}

function round2(value: number) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function money(value: number) {
  return round2(value).toFixed(2);
}

function rowKey(row: SmokeRow, date: string) {
  return `${date}__${row.name}__${row.price}`;
}

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

      const purchaseRaw = row[3];
      const morningStockRaw = row[4];
      const morningSoldRaw = row[5];
      const morningAmountRaw = row[6];
      const nightStockRaw = row[7];
      const nightSoldRaw = row[8];
      const nightAmountRaw = row[9];
      const totalAmountRaw = row[10];

      const purchase = n(purchaseRaw);
      const morningStock = n(morningStockRaw);
      const morningSold = n(morningSoldRaw);
      const nightStock = n(nightStockRaw);
      const nightSold = n(nightSoldRaw);

      const hasPurchase = hasValue(purchaseRaw);
      const hasMorningStock = hasValue(morningStockRaw);
      const hasMorningSold = hasValue(morningSoldRaw);
      const hasNightStock = hasValue(nightStockRaw);
      const hasNightSold = hasValue(nightSoldRaw);
      const hasMorningAmount = hasValue(morningAmountRaw) && n(morningAmountRaw) !== 0;
      const hasNightAmount = hasValue(nightAmountRaw) && n(nightAmountRaw) !== 0;
      const hasTotalAmount = hasValue(totalAmountRaw) && n(totalAmountRaw) !== 0;
      const active = hasPurchase || hasMorningStock || hasMorningSold || hasNightStock || hasNightSold || hasMorningAmount || hasNightAmount || hasTotalAmount;
      if (!active) continue;

      const morningAmount = n(morningAmountRaw) || round2(morningSold * price);
      const nightAmount = n(nightAmountRaw) || round2(nightSold * price);
      const totalAmount = n(totalAmountRaw) || round2(morningAmount + nightAmount);
      const sameDayExpectedNightStock = round2(morningStock + purchase - morningSold);
      const endingStock = hasNightStock ? round2(nightStock - nightSold) : sameDayExpectedNightStock;

      parsed.push({
        date,
        sheetName,
        rowNumber: i + 1,
        no: row[0] as number | string,
        name,
        price,
        purchase,
        morningStock,
        morningSold,
        morningAmount,
        nightStock,
        nightSold,
        nightAmount,
        totalAmount,
        endingStock,
        sameDayExpectedNightStock,
        sameDayStockDiff: hasNightStock ? round2(nightStock - sameDayExpectedNightStock) : 0,
        morningAmountDiff: hasMorningSold || hasMorningAmount ? round2(morningAmount - morningSold * price) : 0,
        nightAmountDiff: hasNightSold || hasNightAmount ? round2(nightAmount - nightSold * price) : 0,
        totalAmountDiff: hasMorningAmount || hasNightAmount || hasTotalAmount ? round2(totalAmount - morningAmount - nightAmount) : 0,
        hasMorningStock,
        hasMorningSold,
        hasNightStock,
        hasNightSold,
        hasPurchase,
        hasMorningAmount,
        hasNightAmount,
        hasTotalAmount
      });
    }
  });
  return parsed.sort((a, b) => `${a.date}-${a.rowNumber}`.localeCompare(`${b.date}-${b.rowNumber}`));
}

function hasRowIssue(row: SmokeRow) {
  return (row.hasNightStock && Math.abs(row.sameDayStockDiff) > 0.01) || Math.abs(row.morningAmountDiff) > 0.01 || Math.abs(row.nightAmountDiff) > 0.01 || Math.abs(row.totalAmountDiff) > 0.01;
}

function downloadWorkbook(fileName: string, sheets: Record<string, Record<string, unknown>[]>) {
  const workbook = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name, sheetRows]) => XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(sheetRows), name.slice(0, 31)));
  XLSX.writeFile(workbook, fileName);
}

function Card({ title, value, desc }: { title: string; value: string; desc?: string }) {
  return <div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="text-sm text-slate-500">{title}</div><div className="mt-2 text-2xl font-bold text-slate-900">{value}</div>{desc ? <div className="mt-1 text-xs text-slate-500">{desc}</div> : null}</div>;
}

function Badge({ children, tone = "gray" }: { children: React.ReactNode; tone?: "gray" | "green" | "red" | "yellow" }) {
  const cls = { gray: "bg-slate-100 text-slate-700", green: "bg-green-100 text-green-700", red: "bg-red-100 text-red-700", yellow: "bg-yellow-100 text-yellow-800" }[tone];
  return <span className={`rounded-full px-2 py-1 text-xs font-medium ${cls}`}>{children}</span>;
}

export default function Home() {
  const [tab, setTab] = useState<Tab>("overview");
  const [fileName, setFileName] = useState("");
  const [rows, setRows] = useState<SmokeRow[]>([]);
  const [countInput, setCountInput] = useState<Record<string, string>>({});
  const [corrections, setCorrections] = useState<Record<string, CorrectionRecord>>({});

  async function handleFile(file: File) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    setRows(parseWorkbook(workbook));
    setFileName(file.name);
    setCountInput({});
    setCorrections({});
  }

  const daily = useMemo<DailySummary[]>(() => {
    const map = new Map<string, DailySummary>();
    rows.forEach((row) => {
      const item = map.get(row.date) ?? { date: row.date, purchase: 0, morningSold: 0, nightSold: 0, totalSold: 0, morningAmount: 0, nightAmount: 0, totalAmount: 0, abnormalCount: 0 };
      item.purchase += row.purchase;
      item.morningSold += row.morningSold;
      item.nightSold += row.nightSold;
      item.totalSold += row.morningSold + row.nightSold;
      item.morningAmount += row.morningAmount;
      item.nightAmount += row.nightAmount;
      item.totalAmount += row.totalAmount;
      item.abnormalCount += hasRowIssue(row) ? 1 : 0;
      map.set(row.date, item);
    });
    return Array.from(map.values()).map((item) => ({ ...item, purchase: round2(item.purchase), morningSold: round2(item.morningSold), nightSold: round2(item.nightSold), totalSold: round2(item.totalSold), morningAmount: round2(item.morningAmount), nightAmount: round2(item.nightAmount), totalAmount: round2(item.totalAmount) })).sort((a, b) => a.date.localeCompare(b.date));
  }, [rows]);

  const productSummary = useMemo<ProductSummary[]>(() => {
    const map = new Map<string, ProductSummary>();
    rows.forEach((row) => {
      const key = `${row.name}__${row.price}`;
      const item = map.get(key) ?? { name: row.name, price: row.price, purchase: 0, morningSold: 0, nightSold: 0, totalSold: 0, totalAmount: 0, endingStock: 0 };
      item.purchase += row.purchase;
      item.morningSold += row.morningSold;
      item.nightSold += row.nightSold;
      item.totalSold += row.morningSold + row.nightSold;
      item.totalAmount += row.totalAmount;
      item.endingStock = row.endingStock;
      map.set(key, item);
    });
    return Array.from(map.values()).map((item) => ({ ...item, purchase: round2(item.purchase), morningSold: round2(item.morningSold), nightSold: round2(item.nightSold), totalSold: round2(item.totalSold), totalAmount: round2(item.totalAmount), endingStock: round2(item.endingStock) })).sort((a, b) => b.totalAmount - a.totalAmount);
  }, [rows]);

  const handovers = useMemo<HandoverIssue[]>(() => {
    const issues: HandoverIssue[] = [];
    rows.forEach((row) => {
      if (row.hasNightStock && Math.abs(row.sameDayStockDiff) > 0.01) issues.push({ type: "早班交夜班", date: row.date, name: row.name, expected: row.sameDayExpectedNightStock, actual: row.nightStock, diff: row.sameDayStockDiff, note: "夜班库存应等于：早班库存 + 进货 - 早班售卖数量。夜班库存为空时不核对。" });
      if (Math.abs(row.morningAmountDiff) > 0.01) issues.push({ type: "金额核对", date: row.date, name: row.name, expected: round2(row.morningSold * row.price), actual: row.morningAmount, diff: row.morningAmountDiff, note: "早班金额不等于：早班售卖数量 × 价格" });
      if (Math.abs(row.nightAmountDiff) > 0.01) issues.push({ type: "金额核对", date: row.date, name: row.name, expected: round2(row.nightSold * row.price), actual: row.nightAmount, diff: row.nightAmountDiff, note: "夜班金额不等于：夜班售卖数量 × 价格" });
      if (Math.abs(row.totalAmountDiff) > 0.01) issues.push({ type: "金额核对", date: row.date, name: row.name, expected: round2(row.morningAmount + row.nightAmount), actual: row.totalAmount, diff: row.totalAmountDiff, note: "合计金额不等于：早班金额 + 夜班金额" });
    });
    const byProduct = new Map<string, SmokeRow[]>();
    rows.forEach((row) => {
      const list = byProduct.get(row.name) ?? [];
      list.push(row);
      byProduct.set(row.name, list);
    });
    byProduct.forEach((list) => {
      const sorted = [...list].sort((a, b) => a.date.localeCompare(b.date));
      for (let i = 0; i < sorted.length - 1; i += 1) {
        const current = sorted[i];
        const next = sorted[i + 1];
        if (!current.hasNightStock || !next.hasMorningStock) continue;
        const diff = round2(next.morningStock - current.endingStock);
        if (Math.abs(diff) > 0.01) issues.push({ type: "次日接班", date: current.date, nextDate: next.date, name: current.name, expected: current.endingStock, actual: next.morningStock, diff, note: "次日早班库存应等于：前一日夜班库存 - 前一日夜班售卖数量" });
      }
    });
    return issues.sort((a, b) => `${a.date}-${a.name}`.localeCompare(`${b.date}-${b.name}`));
  }, [rows]);

  const latestDate = daily.length ? daily[daily.length - 1].date : "";
  const latestRows = rows.filter((row) => row.date === latestDate);
  const totalAmount = daily.reduce((sum, item) => sum + item.totalAmount, 0);
  const totalSold = daily.reduce((sum, item) => sum + item.totalSold, 0);
  const correctionRows = Object.values(corrections).sort((a, b) => `${a.date}-${a.name}`.localeCompare(`${b.date}-${b.name}`));

  function correctedExpected(row: SmokeRow) {
    return corrections[rowKey(row, latestDate)]?.after ?? row.endingStock;
  }

  function applyCorrection(row: SmokeRow, actual: number) {
    const key = rowKey(row, latestDate);
    setCorrections((old) => ({
      ...old,
      [key]: { date: latestDate, name: row.name, price: row.price, before: row.endingStock, after: actual, diff: round2(actual - row.endingStock), correctedAt: new Date().toLocaleString() }
    }));
    setCountInput((old) => ({ ...old, [row.name]: String(actual) }));
  }

  function undoCorrection(row: SmokeRow) {
    const key = rowKey(row, latestDate);
    setCorrections((old) => {
      const copy = { ...old };
      delete copy[key];
      return copy;
    });
  }

  function exportAll() {
    downloadWorkbook("烟报自动核对汇总.xlsx", {
      每日汇总: daily.map((item) => ({ 日期: item.date, 进货: item.purchase, 早班销量: item.morningSold, 夜班销量: item.nightSold, 总销量: item.totalSold, 早班金额: item.morningAmount, 夜班金额: item.nightAmount, 总金额: item.totalAmount, 当日异常数: item.abnormalCount })),
      商品月汇总: productSummary.map((item) => ({ 商品名称: item.name, 价格: item.price, 月进货: item.purchase, 早班销量: item.morningSold, 夜班销量: item.nightSold, 月销量: item.totalSold, 月销售金额: item.totalAmount, 月末结存: item.endingStock })),
      交接异常: handovers.map((item) => ({ 类型: item.type, 日期: item.date, 次日: item.nextDate ?? "", 商品名称: item.name, 应为: item.expected, 实际: item.actual, 差异: item.diff, 说明: item.note })),
      纠正记录: correctionRows.length ? correctionRows.map((item) => ({ 日期: item.date, 商品名称: item.name, 价格: item.price, 原应剩: item.before, 纠正后: item.after, 纠正差异: item.diff, 纠正时间: item.correctedAt })) : [{ 说明: "暂无纠正记录" }],
      原始明细: rows.map((row) => ({ 日期: row.date, Sheet: row.sheetName, 行号: row.rowNumber, 序号: row.no, 商品名称: row.name, 价格: row.price, 进货: row.purchase, 早班库存: row.morningStock, 早班售卖数量: row.morningSold, 早班售卖金额: row.morningAmount, 夜班库存: row.hasNightStock ? row.nightStock : "未填", 夜班售卖数量: row.nightSold, 夜班售卖金额: row.nightAmount, 合计金额: row.totalAmount, 当前结存: row.endingStock }))
    });
  }

  function exportOnsiteCount() {
    if (!latestRows.length) return;
    const countRows = latestRows.map((row) => {
      const actualText = countInput[row.name] ?? "";
      const actual = actualText === "" ? "" : n(actualText);
      const expected = correctedExpected(row);
      const diff = actual === "" ? "" : round2(Number(actual) - expected);
      const correction = corrections[rowKey(row, latestDate)];
      return { 日期: latestDate, 商品名称: row.name, 价格: row.price, 原应剩数量: row.endingStock, 当前应剩数量: expected, 现场实点: actual, 差异: diff, 是否已纠正: correction ? "是" : "否" };
    });
    downloadWorkbook(`${latestDate}-现场点烟表.xlsx`, {
      现场点烟: countRows,
      差异商品: countRows.filter((row) => row.差异 !== "" && row.差异 !== 0),
      纠正记录: correctionRows.length ? correctionRows.map((item) => ({ 日期: item.date, 商品名称: item.name, 价格: item.price, 原应剩: item.before, 纠正后: item.after, 纠正差异: item.diff, 纠正时间: item.correctedAt })) : [{ 说明: "暂无纠正记录" }]
    });
  }

  return (
    <main className="min-h-screen p-4 md:p-8">
      <div className="mx-auto max-w-7xl">
        <header className="mb-6 rounded-3xl bg-slate-900 p-6 text-white shadow-sm">
          <p className="text-sm text-slate-300">Cigarette Report Reconciliation</p>
          <h1 className="mt-1 text-2xl font-bold md:text-4xl">烟报自动核对助手</h1>
          <p className="mt-2 max-w-3xl text-sm text-slate-300">按你上传的手工烟报模板解析：左侧 A-K 香烟区域，自动生成每日汇总、商品月汇总、交接异常和现场点烟清单。</p>
        </header>

        <section className="mb-6 rounded-2xl border bg-white p-5 shadow-sm">
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">导入整个月手工烟报 Excel</h2><p className="mt-1 text-sm text-slate-500">直接上传手工烟报 Excel 文件。系统会按最新日期算出当前每种烟应剩数量，方便你现场点烟。</p></div><label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-slate-900 px-4 py-3 text-sm font-medium text-white">选择烟报文件<input className="hidden" type="file" accept=".xlsx,.xls" onChange={(event) => event.target.files?.[0] && handleFile(event.target.files[0])} /></label></div>
          {fileName ? <p className="mt-3 text-sm text-slate-600">已导入：<b>{fileName}</b>，共解析 {rows.length} 条有效香烟明细。</p> : null}
        </section>

        <nav className="mb-6 flex flex-wrap gap-2">{nav.map(([key, label]) => <button key={key} onClick={() => setTab(key)} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === key ? "bg-slate-900 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>{label}</button>)}</nav>

        {!rows.length ? <section className="rounded-2xl border bg-white p-6 shadow-sm"><h2 className="text-lg font-bold">等待导入烟报</h2><p className="mt-2 text-sm text-slate-600">请先上传手工烟报 Excel。导入后会自动显示汇总、异常和现场点烟清单。</p></section> : null}

        {rows.length > 0 && tab === "overview" ? <section className="space-y-6"><div className="grid gap-4 md:grid-cols-4"><Card title="日期数量" value={`${daily.length} 天`} desc={`最新日期 ${latestDate}`} /><Card title="香烟品规" value={`${productSummary.length} 个`} /><Card title="月销量" value={`${round2(totalSold)} 包`} /><Card title="月销售金额" value={`¥${money(totalAmount)}`} desc={`异常 ${handovers.length} 条，纠正 ${correctionRows.length} 条`} /></div><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">一键导出核对结果</h2><p className="mt-1 text-sm text-slate-500">导出后包含：每日汇总、商品月汇总、交接异常、纠正记录、原始明细。</p></div><button onClick={exportAll} className="rounded-xl bg-green-700 px-4 py-2 text-sm font-medium text-white">导出烟报汇总 Excel</button></div></div></section> : null}

        {rows.length > 0 && tab === "daily" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">每日烟报汇总</h2><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["日期", "进货", "早班销量", "夜班销量", "总销量", "早班金额", "夜班金额", "总金额", "异常"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{daily.map((item) => <tr key={item.date} className="border-t"><td className="p-3">{item.date}</td><td className="p-3">{item.purchase}</td><td className="p-3">{item.morningSold}</td><td className="p-3">{item.nightSold}</td><td className="p-3">{item.totalSold}</td><td className="p-3">¥{money(item.morningAmount)}</td><td className="p-3">¥{money(item.nightAmount)}</td><td className="p-3">¥{money(item.totalAmount)}</td><td className="p-3">{item.abnormalCount ? <Badge tone="red">{item.abnormalCount} 条</Badge> : <Badge tone="green">正常</Badge>}</td></tr>)}</tbody></table></div></section> : null}

        {rows.length > 0 && tab === "summary" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">商品月汇总</h2><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["商品", "价格", "月进货", "早班销量", "夜班销量", "月销量", "月销售金额", "月末结存"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{productSummary.map((item) => <tr key={`${item.name}-${item.price}`} className="border-t"><td className="p-3">{item.name}</td><td className="p-3">¥{money(item.price)}</td><td className="p-3">{item.purchase}</td><td className="p-3">{item.morningSold}</td><td className="p-3">{item.nightSold}</td><td className="p-3">{item.totalSold}</td><td className="p-3">¥{money(item.totalAmount)}</td><td className="p-3">{item.endingStock}</td></tr>)}</tbody></table></div></section> : null}

        {rows.length > 0 && tab === "handover" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">交接和金额异常</h2><p className="mt-1 text-sm text-slate-500">夜班库存为空时，说明该班次未填完，不再把空白当作 0 来误报异常。</p><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["类型", "日期", "次日", "商品", "应为", "实际", "差异", "说明"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{handovers.map((item, index) => <tr key={`${item.type}-${item.date}-${item.name}-${index}`} className="border-t"><td className="p-3"><Badge tone="red">{item.type}</Badge></td><td className="p-3">{item.date}</td><td className="p-3">{item.nextDate ?? ""}</td><td className="p-3">{item.name}</td><td className="p-3">{item.expected}</td><td className="p-3">{item.actual}</td><td className="p-3">{item.diff}</td><td className="p-3">{item.note}</td></tr>)}</tbody></table></div></section> : null}

        {rows.length > 0 && tab === "count" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">现场点烟清单</h2><p className="mt-1 text-sm text-slate-500">输入现场实点数量。有差异时点“按实点纠正”，系统会把当前应剩改成实点数并记录纠正记录。</p></div><button onClick={exportOnsiteCount} className="rounded-xl bg-slate-900 px-4 py-2 text-sm font-medium text-white">导出现场点烟表</button></div><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["商品", "价格", "当前应剩", "现场实点", "差异", "状态", "纠正"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{latestRows.map((row) => { const key = rowKey(row, latestDate); const actualText = countInput[row.name] ?? ""; const actual = actualText === "" ? null : n(actualText); const expected = correctedExpected(row); const diff = actual === null ? null : round2(actual - expected); const correction = corrections[key]; return <tr key={key} className="border-t"><td className="p-3">{row.name}</td><td className="p-3">¥{money(row.price)}</td><td className="p-3 font-bold">{expected}{correction ? <div className="text-xs font-normal text-slate-500">原：{row.endingStock}</div> : null}</td><td className="p-3"><input className="w-28 rounded-xl border p-2" inputMode="decimal" value={actualText} onChange={(event) => setCountInput({ ...countInput, [row.name]: event.target.value })} /></td><td className="p-3">{diff === null ? "" : diff}</td><td className="p-3">{correction ? <Badge tone="yellow">已纠正</Badge> : diff === null ? <Badge>未点</Badge> : diff === 0 ? <Badge tone="green">正常</Badge> : <Badge tone="red">有差异</Badge>}</td><td className="p-3">{actual !== null && diff !== 0 ? <button className="rounded-lg bg-slate-900 px-3 py-1 text-xs text-white" onClick={() => applyCorrection(row, actual)}>按实点纠正</button> : null}{correction ? <button className="ml-2 rounded-lg border px-3 py-1 text-xs" onClick={() => undoCorrection(row)}>撤销</button> : null}</td></tr>; })}</tbody></table></div>{correctionRows.length ? <div className="mt-5 rounded-xl bg-yellow-50 p-4 text-sm text-yellow-900"><b>已纠正 {correctionRows.length} 条：</b>{correctionRows.map((item) => <div key={`${item.date}-${item.name}-${item.price}`}>{item.name}：{item.before} → {item.after}，差异 {item.diff}</div>)}</div> : null}</section> : null}
      </div>
    </main>
  );
}
