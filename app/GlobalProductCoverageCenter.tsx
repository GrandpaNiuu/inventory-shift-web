"use client";

import { useMemo, useState, type ChangeEvent } from "react";
import * as XLSX from "xlsx";
import { GLOBAL_PRODUCT_CATEGORIES, classifyGlobalProduct } from "./globalProductCatalog";

type ProductRecord = {
  sheetName: string;
  rowNumber: number;
  productName: string;
  sku: string;
  price: number;
  quantity: number;
  amount: number;
  category: string;
  confidence: number;
  matchedKeyword: string;
};

type SheetScan = {
  sheetName: string;
  headerRow: number;
  productCol: number;
  skuCol: number;
  priceCol: number;
  quantityCol: number;
  amountCol: number;
  rows: ProductRecord[];
};

function text(value: unknown) { return String(value ?? "").trim(); }
function norm(value: unknown) { return text(value).replace(/\s/g, "").toLowerCase(); }
function toNumber(value: unknown) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const cleaned = text(value).replace(/,/g, "").replace(/[￥¥元\s]/g, "");
  if (!cleaned) return 0;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}
function money(value: number) { return Math.round((value + Number.EPSILON) * 100) / 100; }
function isLikelyProductHeader(value: unknown) {
  const v = norm(value);
  return /商品|产品|品名|货品|物品|名称|标题|name|product|item|товар|название|منتج|اسم/.test(v);
}
function scoreHeaderRow(row: unknown[]) {
  const joined = norm(row.map(text).filter(Boolean).join(" "));
  let score = row.filter((cell) => text(cell)).length;
  if (/商品|产品|品名|货品|名称|标题|product|item|name|товар|название|منتج/.test(joined)) score += 12;
  if (/价格|单价|售价|price|цена|السعر/.test(joined)) score += 5;
  if (/数量|库存|销量|qty|quantity|stock|количество|كمية/.test(joined)) score += 5;
  if (/sku|货号|编码|条码|barcode|артикул|رمز/.test(joined)) score += 4;
  if (/金额|总价|合计|amount|total|сумма|المبلغ/.test(joined)) score += 4;
  if (/合计|总计|小计/.test(joined)) score -= 4;
  return score;
}
function findHeaderRow(grid: unknown[][]) {
  let best = 0;
  let bestScore = -Infinity;
  for (let i = 0; i < Math.min(35, grid.length); i += 1) {
    const score = scoreHeaderRow(grid[i] ?? []);
    if (score > bestScore) {
      best = i;
      bestScore = score;
    }
  }
  return best;
}
function findColumn(headers: unknown[], patterns: RegExp[]) {
  for (let i = 0; i < headers.length; i += 1) {
    const value = norm(headers[i]);
    if (patterns.some((pattern) => pattern.test(value))) return i;
  }
  return -1;
}
function detectProductColumn(headers: unknown[]) {
  const direct = findColumn(headers, [/商品|产品|品名|货品|物品|标题|product|item|товар|название|منتج/]);
  if (direct >= 0) return direct;
  return headers.findIndex(isLikelyProductHeader);
}
function scanWorkbook(workbook: XLSX.WorkBook) {
  const sheets: SheetScan[] = [];
  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;
    const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null });
    const headerRow = findHeaderRow(grid);
    const headers = grid[headerRow] ?? [];
    const productCol = detectProductColumn(headers);
    const skuCol = findColumn(headers, [/sku|货号|编码|条码|barcode|артикул|رمز/]);
    const priceCol = findColumn(headers, [/价格|单价|售价|price|цена|السعر/]);
    const quantityCol = findColumn(headers, [/数量|销量|库存|件数|qty|quantity|stock|количество|كمية/]);
    const amountCol = findColumn(headers, [/金额|总价|合计|小计|amount|total|сумма|المبلغ/]);
    const rows: ProductRecord[] = [];

    if (productCol < 0) return;
    for (let r = headerRow + 1; r < grid.length; r += 1) {
      const row = grid[r] ?? [];
      const productName = text(row[productCol]);
      if (!productName || /合计|总计|小计|备注|收款|金额/.test(productName)) continue;
      const sku = skuCol >= 0 ? text(row[skuCol]) : "";
      const price = priceCol >= 0 ? toNumber(row[priceCol]) : 0;
      const quantity = quantityCol >= 0 ? toNumber(row[quantityCol]) : 0;
      const amount = amountCol >= 0 ? toNumber(row[amountCol]) : 0;
      const matched = classifyGlobalProduct(`${productName} ${sku}`);
      rows.push({ sheetName, rowNumber: r + 1, productName, sku, price, quantity, amount, category: matched.category, confidence: matched.confidence, matchedKeyword: matched.matchedKeyword });
    }

    sheets.push({ sheetName, headerRow, productCol, skuCol, priceCol, quantityCol, amountCol, rows });
  });
  return sheets;
}
function exportCoverage(fileName: string, rows: ProductRecord[]) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows.map((row, index) => ({ 序号: index + 1, Sheet: row.sheetName, 行号: row.rowNumber, 商品名称: row.productName, SKU: row.sku, 自动分类: row.category, 置信度: Math.round(row.confidence * 100) + "%", 匹配词: row.matchedKeyword, 价格: row.price, 数量: row.quantity, 金额: row.amount }))), "商品识别明细");
  const byCategory = Array.from(rows.reduce((map, row) => {
    const old = map.get(row.category) ?? { category: row.category, count: 0, amount: 0, quantity: 0 };
    old.count += 1;
    old.amount += row.amount;
    old.quantity += row.quantity;
    map.set(row.category, old);
    return map;
  }, new Map<string, { category: string; count: number; amount: number; quantity: number }>()).values());
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(byCategory.map((item, index) => ({ 序号: index + 1, 分类: item.category, 商品数: item.count, 数量合计: money(item.quantity), 金额合计: money(item.amount) }))), "分类覆盖汇总");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows.filter((row) => row.category === "未识别" || row.confidence < 0.6).map((row, index) => ({ 序号: index + 1, Sheet: row.sheetName, 行号: row.rowNumber, 商品名称: row.productName, SKU: row.sku, 建议: "补充别名或加入商品库" }))), "未识别商品");
  XLSX.writeFile(workbook, `${fileName.replace(/\.xlsx?$/i, "")}-商品识别库报告.xlsx`);
}

export default function GlobalProductCoverageCenter() {
  const [fileName, setFileName] = useState("");
  const [sheets, setSheets] = useState<SheetScan[]>([]);

  async function handleFile(file: File) {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true, cellFormula: true, cellNF: true, cellStyles: true });
    setFileName(file.name);
    setSheets(scanWorkbook(workbook));
  }

  const rows = useMemo(() => sheets.flatMap((sheet) => sheet.rows), [sheets]);
  const recognizedRows = rows.filter((row) => row.category !== "未识别" && row.confidence >= 0.55);
  const coverageRate = rows.length ? Math.round((recognizedRows.length / rows.length) * 100) : 0;
  const categorySummary = useMemo(() => Array.from(rows.reduce((map, row) => {
    const old = map.get(row.category) ?? { category: row.category, count: 0, amount: 0, quantity: 0 };
    old.count += 1;
    old.amount += row.amount;
    old.quantity += row.quantity;
    map.set(row.category, old);
    return map;
  }, new Map<string, { category: string; count: number; amount: number; quantity: number }>()).values()).sort((a, b) => b.count - a.count), [rows]);
  const unmatchedRows = rows.filter((row) => row.category === "未识别" || row.confidence < 0.6);

  return <section className="space-y-6">
    <div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">商品识别库</h2><p className="mt-1 text-sm text-slate-500">上传包含商品名称的 Excel，系统会按商品词库自动识别分类、覆盖率和未识别商品。</p></div><label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-slate-950 px-4 py-3 text-sm font-medium text-white">选择商品表<input className="hidden" type="file" accept=".xlsx,.xls" onChange={(event: ChangeEvent<HTMLInputElement>) => event.target.files?.[0] && handleFile(event.target.files[0])} /></label></div>{fileName ? <p className="mt-3 text-sm text-slate-600">已导入：<b>{fileName}</b>，识别 {sheets.length} 个 Sheet，{rows.length} 个商品行。</p> : null}</div>

    {!rows.length ? <div className="rounded-2xl border bg-white p-6 shadow-sm"><h3 className="font-bold">等待上传商品表</h3><p className="mt-2 text-sm text-slate-600">支持中文、英文、俄文、阿拉伯语关键词。优先识别商品名称、SKU/货号、价格、数量、金额。</p><div className="mt-4 grid gap-3 md:grid-cols-4">{GLOBAL_PRODUCT_CATEGORIES.map((category) => <div key={category.id} className="rounded-xl bg-slate-50 p-3"><div className="font-medium">{category.name}</div><div className="text-xs text-slate-500">{category.globalName}</div></div>)}</div></div> : null}

    {rows.length ? <><div className="grid gap-4 md:grid-cols-4"><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="text-sm text-slate-500">商品行</div><div className="mt-2 text-2xl font-bold">{rows.length}</div></div><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="text-sm text-slate-500">识别覆盖率</div><div className="mt-2 text-2xl font-bold">{coverageRate}%</div></div><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="text-sm text-slate-500">分类数</div><div className="mt-2 text-2xl font-bold">{categorySummary.length}</div></div><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="text-sm text-slate-500">待补充商品</div><div className="mt-2 text-2xl font-bold">{unmatchedRows.length}</div></div></div>
    <div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-wrap items-center justify-between gap-3"><div><h3 className="font-bold">分类覆盖汇总</h3><p className="mt-1 text-sm text-slate-500">按自动识别分类统计商品数量、数量和金额。</p></div><button onClick={() => exportCoverage(fileName || "商品表", rows)} className="rounded-xl bg-green-700 px-4 py-2 text-sm text-white">导出识别报告</button></div><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["分类", "商品数", "数量合计", "金额合计"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{categorySummary.map((item) => <tr key={item.category} className="border-t"><td className="p-3 font-medium">{item.category}</td><td className="p-3">{item.count}</td><td className="p-3">{money(item.quantity)}</td><td className="p-3">¥{money(item.amount)}</td></tr>)}</tbody></table></div></div>
    <div className="rounded-2xl border bg-white p-5 shadow-sm"><h3 className="font-bold">商品识别明细</h3><div className="table-scroll mt-4 max-h-[520px] overflow-auto"><table className="min-w-full text-sm"><thead className="sticky top-0 bg-slate-50 text-left"><tr>{["Sheet", "行号", "商品", "SKU", "分类", "置信度", "匹配词", "价格", "数量", "金额"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{rows.map((row) => <tr key={`${row.sheetName}-${row.rowNumber}-${row.productName}`} className="border-t"><td className="p-3">{row.sheetName}</td><td className="p-3">{row.rowNumber}</td><td className="p-3">{row.productName}</td><td className="p-3">{row.sku}</td><td className="p-3">{row.category}</td><td className="p-3">{Math.round(row.confidence * 100)}%</td><td className="p-3">{row.matchedKeyword}</td><td className="p-3">{row.price ? `¥${money(row.price)}` : ""}</td><td className="p-3">{row.quantity || ""}</td><td className="p-3">{row.amount ? `¥${money(row.amount)}` : ""}</td></tr>)}</tbody></table></div></div>
    {unmatchedRows.length ? <div className="rounded-2xl border bg-yellow-50 p-5 text-yellow-900 shadow-sm"><h3 className="font-bold">待补充商品</h3><p className="mt-1 text-sm">这些商品没有高置信度匹配。后续可把它们加入商品别名库或商品词库，系统会越用越准。</p><div className="mt-3 flex flex-wrap gap-2">{unmatchedRows.slice(0, 80).map((row) => <span key={`${row.sheetName}-${row.rowNumber}`} className="rounded-full bg-white px-3 py-1 text-xs shadow-sm">{row.productName}</span>)}</div></div> : null}</> : null}
  </section>;
}
