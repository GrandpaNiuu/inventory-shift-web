"use client";

import { useMemo, useState, type ChangeEvent } from "react";
import * as XLSX from "xlsx";

type Tab = "overview" | "detect" | "daily" | "summary" | "payment";
type Shift = "早班" | "夜班" | "合计" | "未知";
type ProductRow = { date: string; sheet: string; name: string; price: number; purchase: number; morningStock: number; morningSold: number; morningAmount: number; nightStock: number; nightSold: number; nightAmount: number; totalAmount: number; order: number };
type PaymentRow = { date: string; sheet: string; shift: Shift; channel: string; amount: number; order: number };
type SheetInfo = { sheet: string; date: string; productRows: number; paymentRows: number; skipped: boolean };

const tabs: Array<[Tab, string]> = [["overview", "总览"], ["detect", "识别预览"], ["daily", "每日报表"], ["summary", "商品总汇"], ["payment", "收款统计"]];
const channels = ["小Y", "现金", "微信", "扫码盒子", "支付宝"];

function text(v: unknown) { return String(v ?? "").trim(); }
function compact(v: unknown) { return text(v).replace(/\s/g, ""); }
function num(v: unknown) { if (typeof v === "number") return Number.isFinite(v) ? v : 0; const n = Number(text(v).replace(/,/g, "").replace(/[￥¥元\s]/g, "")); return Number.isFinite(n) ? n : 0; }
function money(n: number) { return (Math.round((n + Number.EPSILON) * 100) / 100).toFixed(2); }
function rowText(row: unknown[] | undefined) { return (row ?? []).map(text).filter(Boolean).join(" "); }
function hasFullDate(sheet: string, title: unknown) { return /(20\d{2}|19\d{2})\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日?/.test(`${title ?? ""} ${sheet}`); }
function shouldReadSheet(sheet: string, grid: unknown[][]) { return /^\d{1,2}$/.test(sheet) || hasFullDate(sheet, grid[0]?.[0]); }
function getDate(sheet: string, title: unknown, index: number) { const raw = `${title ?? ""} ${sheet}`; const full = raw.match(/(20\d{2}|19\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?/); if (full) return `${full[1]}-${full[2].padStart(2, "0")}-${full[3].padStart(2, "0")}`; const ym = raw.match(/(20\d{2}|19\d{2})\s*年\s*(\d{1,2})\s*月/); const day = sheet.match(/^\d{1,2}$/) ? Number(sheet) : index + 1; return `${ym?.[1] ?? new Date().getFullYear()}-${String(ym?.[2] ?? new Date().getMonth() + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`; }
function cleanChannel(v: unknown) { const c = compact(v); if (!c || /^\d+(\.\d+)?$/.test(c)) return ""; if (/^(收款渠道|渠道|金额|合计|总计|早班|白班|夜班|售卖金额|售卖数量|备注)$/.test(c)) return ""; return c; }

function shiftFromPaymentArea(grid: unknown[][], rowIndex: number, col: number): Shift {
  const fromCol = Math.max(0, col - 2);
  const toCol = col + 8;
  for (let r = rowIndex - 1; r >= Math.max(0, rowIndex - 8); r -= 1) {
    const line = (grid[r] ?? []).slice(fromCol, toCol).map(compact).filter(Boolean).join("");
    if (!line) continue;
    if (/(早班|白班)/.test(line)) return "早班";
    if (/夜班/.test(line)) return "夜班";
    if (/^(合计售卖金额|合计售卖数量|总计售卖金额|总计售卖数量|收款合计|总计|合计)/.test(line)) return "合计";
  }
  return "未知";
}

function parsePayments(grid: unknown[][], sheet: string, date: string) {
  const out: PaymentRow[] = [];
  grid.forEach((row, r) => row.forEach((cell, c) => {
    if (text(cell) !== "收款渠道") return;
    const shift = shiftFromPaymentArea(grid, r, c);
    for (let col = c + 1; col < row.length - 1; col += 2) {
      const channel = cleanChannel(row[col]);
      const amount = num(row[col + 1]);
      if (channel && amount) out.push({ date, sheet, shift, channel, amount, order: r * 1000 + col });
    }
  }));
  return out;
}

function findProductHeaders(grid: unknown[][]) {
  const groups: Array<{ h: number; c: number }> = [];
  grid.forEach((row, r) => row.forEach((cell, c) => {
    const a = compact(cell);
    const b = compact(row[c + 1]);
    if ((a === "商品名称" || a === "售卖商品") && (b === "价格" || b === "售价")) groups.push({ h: r, c });
  }));
  return groups;
}

function parseProducts(grid: unknown[][], sheet: string, date: string, sheetIndex: number) {
  const out: ProductRow[] = [];
  findProductHeaders(grid).forEach(({ h, c }) => {
    for (let r = h + 1; r < grid.length; r += 1) {
      const row = grid[r] ?? [];
      const name = text(row[c]);
      const price = num(row[c + 1]);
      if (!name || !price || /合计|总计|小计|收款|金额|售卖数量|售卖金额|渠道|备注/.test(name)) continue;
      const purchase = num(row[c + 2]);
      const morningStock = num(row[c + 3]);
      const morningSold = num(row[c + 4]);
      const morningAmount = num(row[c + 5]) || morningSold * price;
      const nightStock = num(row[c + 6]);
      const nightSold = num(row[c + 7]);
      const nightAmount = num(row[c + 8]) || nightSold * price;
      const totalAmount = num(row[c + 9]) || morningAmount + nightAmount;
      if (!(purchase || morningStock || morningSold || nightStock || nightSold || totalAmount)) continue;
      out.push({ date, sheet, name, price, purchase, morningStock, morningSold, morningAmount, nightStock, nightSold, nightAmount, totalAmount, order: sheetIndex * 1000000 + c * 10000 + r });
    }
  });
  return out;
}

function validPayments(rows: PaymentRow[]) {
  const byDate = new Map<string, PaymentRow[]>();
  rows.forEach((r) => byDate.set(r.date, [...(byDate.get(r.date) ?? []), r]));
  return Array.from(byDate.values()).flatMap((items) => {
    const totals = items.filter((r) => r.shift === "合计");
    return totals.length ? totals : items.filter((r) => r.shift !== "合计");
  });
}

export default function InventoryPageConfigurableFixed() {
  const [tab, setTab] = useState<Tab>("overview");
  const [fileName, setFileName] = useState("");
  const [products, setProducts] = useState<ProductRow[]>([]);
  const [payments, setPayments] = useState<PaymentRow[]>([]);
  const [sheets, setSheets] = useState<SheetInfo[]>([]);

  async function handleFile(file: File) {
    const wb = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
    const productRows: ProductRow[] = [];
    const paymentRows: PaymentRow[] = [];
    const sheetInfos: SheetInfo[] = [];
    wb.SheetNames.forEach((sheetName, index) => {
      const sheet = wb.Sheets[sheetName];
      if (!sheet) return;
      const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null });
      if (!shouldReadSheet(sheetName, grid)) { sheetInfos.push({ sheet: sheetName, date: "已跳过", productRows: 0, paymentRows: 0, skipped: true }); return; }
      const date = getDate(sheetName, grid[0]?.[0], index);
      const pr = parseProducts(grid, sheetName, date, index);
      const py = parsePayments(grid, sheetName, date);
      productRows.push(...pr);
      paymentRows.push(...py);
      sheetInfos.push({ sheet: sheetName, date, productRows: pr.length, paymentRows: py.length, skipped: false });
    });
    setFileName(file.name);
    setProducts(productRows);
    setPayments(paymentRows);
    setSheets(sheetInfos);
  }

  const effectivePayments = useMemo(() => validPayments(payments), [payments]);
  const salesAmount = products.reduce((s, r) => s + r.totalAmount, 0);
  const salesQty = products.reduce((s, r) => s + r.morningSold + r.nightSold, 0);
  const paymentAmount = effectivePayments.reduce((s, r) => s + r.amount, 0);
  const diff = paymentAmount - salesAmount;
  const channelSummary = channels.map((channel) => ({ channel, amount: effectivePayments.filter((r) => r.channel === channel).reduce((s, r) => s + r.amount, 0) }));
  const daily = Array.from(new Set(products.map((r) => r.date))).sort().map((date) => {
    const rows = products.filter((r) => r.date === date);
    return { date, qty: rows.reduce((s, r) => s + r.morningSold + r.nightSold, 0), amount: rows.reduce((s, r) => s + r.totalAmount, 0), payment: effectivePayments.filter((r) => r.date === date).reduce((s, r) => s + r.amount, 0) };
  });
  const productSummary = Array.from(new Map(products.map((row) => [`${row.name}-${row.price}`, row])).keys()).map((key) => {
    const rows = products.filter((r) => `${r.name}-${r.price}` === key);
    return { name: rows[0]?.name ?? "", price: rows[0]?.price ?? 0, qty: rows.reduce((s, r) => s + r.morningSold + r.nightSold, 0), amount: rows.reduce((s, r) => s + r.totalAmount, 0) };
  });

  return <main className="min-h-screen bg-slate-50 p-4 md:p-8">
    <div className="mx-auto max-w-7xl">
      <header className="rounded-3xl bg-slate-950 p-6 text-white shadow-sm">
        <p className="text-sm text-slate-300">修复版商品详细核对</p>
        <h1 className="mt-1 text-2xl font-bold md:text-4xl">商品日报智能核对系统</h1>
        <p className="mt-2 text-sm text-slate-300">收款按最近标题识别：有总合计只用总合计；没有总合计才用早班 + 夜班。</p>
        <div className="mt-5 flex flex-wrap gap-3"><label className="inline-flex cursor-pointer items-center rounded-xl bg-white px-4 py-3 text-sm font-medium text-slate-950"><input className="hidden" type="file" accept=".xlsx,.xls" onChange={(e: ChangeEvent<HTMLInputElement>) => e.target.files?.[0] && handleFile(e.target.files[0])} />导入商品报表 Excel</label>{fileName ? <span className="rounded-xl bg-slate-800 px-4 py-3 text-sm text-slate-200">当前文件：{fileName}</span> : null}</div>
      </header>
      <nav className="my-5 flex flex-wrap gap-2">{tabs.map(([k, label]) => <button key={k} onClick={() => setTab(k)} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === k ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm"}`}>{label}</button>)}</nav>
      <section className="mb-5 grid gap-4 md:grid-cols-4"><Card title="销售金额" value={`¥${money(salesAmount)}`} desc={`${salesQty} 件`} /><Card title="渠道收款" value={`¥${money(paymentAmount)}`} desc="总合计优先" /><Card title="收款差异" value={`¥${money(diff)}`} desc="收款 - 销售" danger={Math.abs(diff) > 0.01} /><Card title="有效收款条数" value={String(effectivePayments.length)} desc="已剔除早夜班重复" /></section>
      {tab === "overview" ? <Panel title="总览"><div className="grid gap-4 md:grid-cols-3"><Card title="识别商品行" value={String(products.length)} /><Card title="商品品规" value={String(productSummary.length)} /><Card title="工作表" value={String(sheets.length)} /></div></Panel> : null}
      {tab === "detect" ? <Panel title="识别预览"><Table headers={["Sheet", "日期", "商品行", "收款记录", "状态"]} rows={sheets.map((s) => ({ Sheet: s.sheet, 日期: s.date, 商品行: s.productRows, 收款记录: s.paymentRows, 状态: s.skipped ? "已跳过" : "已解析" }))} /></Panel> : null}
      {tab === "daily" ? <Panel title="每日报表"><Table headers={["日期", "销量", "销售金额", "有效收款", "差异"]} rows={daily.map((d) => ({ 日期: d.date, 销量: d.qty, 销售金额: money(d.amount), 有效收款: money(d.payment), 差异: money(d.payment - d.amount) }))} /></Panel> : null}
      {tab === "summary" ? <Panel title="商品总汇"><Table headers={["商品", "价格", "销量", "金额"]} rows={productSummary.map((p) => ({ 商品: p.name, 价格: p.price, 销量: p.qty, 金额: money(p.amount) }))} /></Panel> : null}
      {tab === "payment" ? <Panel title="收款统计"><p className="mb-3 text-sm text-slate-500">这里只统计有效收款：当天有总合计收款时，只显示总合计。</p><Table headers={["收款渠道", "月收入"]} rows={channelSummary.map((c) => ({ 收款渠道: c.channel, 月收入: money(c.amount) }))} /><Table headers={["日期", "班次", "收款渠道", "金额"]} rows={effectivePayments.map((p) => ({ 日期: p.date, 班次: p.shift, 收款渠道: p.channel, 金额: money(p.amount) }))} /></Panel> : null}
    </div>
  </main>;
}

function Card({ title, value, desc, danger = false }: { title: string; value: string; desc?: string; danger?: boolean }) { return <div className={`rounded-2xl p-5 shadow-sm ${danger ? "bg-red-600 text-white" : "bg-white text-slate-900"}`}><div className="text-sm opacity-70">{title}</div><div className="mt-2 text-2xl font-bold">{value}</div>{desc ? <div className="mt-1 text-xs opacity-70">{desc}</div> : null}</div>; }
function Panel({ title, children }: { title: string; children: React.ReactNode }) { return <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="mb-3 text-lg font-bold text-slate-900">{title}</h2>{children}</section>; }
function Table({ headers, rows }: { headers: string[]; rows: Record<string, string | number>[] }) { if (!rows.length) return <p className="rounded-xl bg-slate-50 p-4 text-sm text-slate-500">暂无数据</p>; return <div className="overflow-auto rounded-xl border"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{headers.map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{rows.map((row, i) => <tr key={i} className="border-t">{headers.map((h) => <td key={h} className="p-3 whitespace-nowrap">{row[h]}</td>)}</tr>)}</tbody></table></div>; }
