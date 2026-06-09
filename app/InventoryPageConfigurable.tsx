"use client";

import { useMemo, useState, type ChangeEvent, type ReactNode } from "react";
import * as XLSX from "xlsx";

type Tab = "overview" | "rules" | "detect" | "daily" | "summary" | "handover" | "inventory" | "payment";
type IssueType = "早班交夜班" | "疑似未扣早班销量" | "次日接班" | "金额核对";
type PaymentShift = "早班" | "夜班" | "合计" | "未知";

type CategoryRule = {
  id: string;
  name: string;
  keywords: string;
  includeSales: boolean;
  includeStockCheck: boolean;
  includeInventory: boolean;
};

type RawItem = {
  date: string;
  sheetName: string;
  rowIndex: number;
  rowNumber: number;
  sourceOrder: number;
  rawCategory: string;
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
  hasMorningAmount: boolean;
  hasNightAmount: boolean;
  hasTotalAmount: boolean;
  morningStockCol: number;
  nightStockCol: number;
};

type StockItem = RawItem & { category: string };
type PaymentRecord = { date: string; sheetName: string; shift: PaymentShift; channel: string; amount: number };
type SheetPreview = { sheetName: string; date: string; rowCount: number; paymentCount: number; tableCount: number };
type ParsedWorkbook = { rows: RawItem[]; payments: PaymentRecord[]; sheets: SheetPreview[] };
type DailySummary = { date: string; purchase: number; morningSold: number; nightSold: number; totalSold: number; morningAmount: number; nightAmount: number; totalAmount: number; abnormalCount: number };
type ProductSummary = { category: string; name: string; price: number; purchase: number; morningSold: number; nightSold: number; totalSold: number; totalAmount: number; endingStock: number; latestDate: string; latestRow?: StockItem; sourceOrder: number };
type CategorySummary = { category: string; totalSold: number; totalAmount: number; count: number; sourceOrder: number };
type PaymentDailySummary = { date: string; totals: Record<string, number>; total: number };
type WriteBackTarget = { key: string; date: string; sheetName: string; rowIndex: number; rowNumber: number; colIndex: number; cell: string; columnName: string; category: string; name: string; price: number; before: number; after: number };
type Issue = { id: string; type: IssueType; date: string; nextDate?: string; category: string; name: string; price: number; expected: number; actual: number; diff: number; note: string; writeBack?: WriteBackTarget };
type Correction = WriteBackTarget & { diff: number; reason: string; correctedAt: string };
type TableGroup = { headerRow: number; startCol: number; nameCol: number; priceCol: number; purchaseCol: number; morningStockCol: number; morningSoldCol: number; morningAmountCol: number; nightStockCol: number; nightSoldCol: number; nightAmountCol: number; totalAmountCol: number; defaultCategory: string };
type CellWithStyle = XLSX.CellObject & { f?: string; w?: string; z?: string; s?: Record<string, unknown> };

const STORE_NAME = "SY11";
const UNKNOWN_CATEGORY = "未分类";
const tabs: Array<[Tab, string]> = [["overview", "总览"], ["rules", "分类规则"], ["detect", "识别预览"], ["daily", "每日报表"], ["summary", "商品总汇"], ["handover", "交接核对"], ["inventory", "自动库存盘点"], ["payment", "收款统计"]];
const DEFAULT_RULES: CategoryRule[] = [
  { id: "cigarette", name: "香烟(酒店库存)", keywords: "烟,香烟,卷烟,烟草,中华,芙蓉王,利群,玉溪,黄鹤楼,云烟,南京,双喜,红塔山,白沙,娇子,黄金叶,苏烟,泰山,七匹狼,中南海,牡丹,贵烟,真龙,钻石,煊赫门,万宝路,黄山,长白山,延安,兰州,宽窄,荷花,天子,红河,红金龙,金圣,人民大会堂,细支,中支,雨花石", includeSales: true, includeStockCheck: true, includeInventory: true },
  { id: "drink", name: "饮料", keywords: "饮料,矿泉水,纯净水,可乐,雪碧,芬达,苏打水,气泡水,红牛,东鹏,脉动,外星人,农夫山泉,怡宝,百岁山,王老吉,加多宝,茶,冰红茶,绿茶,乌龙茶,咖啡,牛奶,酸奶,椰汁,果汁,柠檬茶,阿萨姆,海之言,娃哈哈,康师傅,打火机,火机,火柴,点烟器,方便面,泡面,螺蛳粉,零食,口香糖,杂项,其他,小商品,日用品,百货,纸巾,扑克牌,牙刷,牙膏,剃须,充电器,数据线,雨伞", includeSales: true, includeStockCheck: true, includeInventory: true }
];

function text(value: unknown) { return String(value ?? "").trim(); }
function hasValue(value: unknown) { return value !== null && value !== undefined && text(value) !== ""; }
function n(value: unknown) { if (typeof value === "number") return Number.isFinite(value) ? value : 0; const cleaned = text(value).replace(/,/g, "").replace(/[￥¥元\s]/g, ""); if (!cleaned) return 0; const parsed = Number(cleaned); return Number.isFinite(parsed) ? parsed : 0; }
function round2(value: number) { return Math.round((value + Number.EPSILON) * 100) / 100; }
function money(value: number) { return round2(value).toFixed(2); }
function rowText(row: unknown[] | undefined) { return (row ?? []).map((cell) => text(cell)).filter(Boolean).join(" "); }
function safeFileName(name: string) { return name.replace(/\.xlsx?$/i, "").replace(/[\\/:*?\"<>|]/g, "-"); }
function byOriginalOrder(a: { sourceOrder: number; name: string }, b: { sourceOrder: number; name: string }) { return a.sourceOrder - b.sourceOrder || a.name.localeCompare(b.name); }
function ruleKeywords(rule: CategoryRule) { return rule.keywords.split(/[,，、\n]/).map((item) => item.trim()).filter(Boolean); }
function ruleByName(rules: CategoryRule[], name: string) { return rules.find((rule) => rule.name === name); }
function makeEmptyTotals(channels: string[]) { return Object.fromEntries(channels.map((channel) => [channel, 0])) as Record<string, number>; }
function reportDateRange(daily: DailySummary[]) { if (!daily.length) return ""; return daily.length === 1 ? daily[0].date : `${daily[0].date} 至 ${daily[daily.length - 1].date}`; }
function queryTime() { return new Date().toLocaleString("zh-CN", { hour12: false }); }
function rowKey(row: StockItem) { return `${row.date}__${row.sheetName}__${row.rowNumber}__${row.category}__${row.name}__${row.price}`; }
function correctionKey(sheetName: string, rowIndex: number, colIndex: number) { return `${sheetName}__${rowIndex}__${colIndex}`; }
function rowCellKey(row: StockItem, colIndex: number) { return correctionKey(row.sheetName, row.rowIndex, colIndex); }
function productKey(category: string, name: string, price: number) { return `${category}__${name}__${price}`; }
function isSummaryLikeName(name: string) { return /合计|收款|备注|售卖金额|售卖数量|渠道|小计|总计/.test(name); }
function isAmountLikeLabel(value: string) { return /^(金额|收款金额|合计|总计|早班|夜班|收款渠道|渠道|售卖金额)$/.test(value) || /^\d+(\.\d+)?$/.test(value); }
function cleanChannelName(value: unknown) { const channel = text(value).replace(/\s/g, ""); return channel && !isAmountLikeLabel(channel) ? channel : ""; }

function categorize(source: string, rules: CategoryRule[]) {
  const value = source.replace(/\s/g, "");
  if (!value) return UNKNOWN_CATEGORY;
  for (const rule of rules) {
    if (ruleKeywords(rule).some((keyword) => value.includes(keyword.replace(/\s/g, "")))) return rule.name;
  }
  return UNKNOWN_CATEGORY;
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

function inferCategoryFromContext(grid: unknown[][], headerRow: number, startCol: number, endCol: number, sheetName: string, rules: CategoryRule[]) {
  const samples = [sheetName];
  for (let row = 0; row <= headerRow; row += 1) for (let col = startCol; col <= endCol; col += 1) samples.push(text(grid[row]?.[col]));
  return categorize(samples.join(" "), rules);
}

function findTableGroups(grid: unknown[][], sheetName: string, rules: CategoryRule[]) {
  const groups: TableGroup[] = [];
  const seen = new Set<string>();
  grid.forEach((row, headerRow) => {
    row.forEach((cell, col) => {
      if (!/^(商品名称|售卖商品)$/.test(text(cell)) || !/^(价格|售价)$/.test(text(row[col + 1]))) return;
      const startCol = Math.max(0, col - 1);
      const key = `${headerRow}_${startCol}`;
      if (seen.has(key)) return;
      seen.add(key);
      groups.push({ headerRow, startCol, nameCol: col, priceCol: col + 1, purchaseCol: col + 2, morningStockCol: col + 3, morningSoldCol: col + 4, morningAmountCol: col + 5, nightStockCol: col + 6, nightSoldCol: col + 7, nightAmountCol: col + 8, totalAmountCol: col + 9, defaultCategory: inferCategoryFromContext(grid, headerRow, startCol, col + 9, sheetName, rules) });
    });
  });
  return groups;
}

function paymentShiftFromContext(grid: unknown[][], rowIndex: number): PaymentShift {
  const current = rowText(grid[rowIndex]);
  if (/合计售卖金额|总计|合计/.test(current)) return "合计";
  if (/早班/.test(current)) return "早班";
  if (/夜班/.test(current)) return "夜班";
  const context = `${rowText(grid[rowIndex - 2])} ${rowText(grid[rowIndex - 1])} ${current}`;
  if (/合计售卖金额|总计|合计/.test(context)) return "合计";
  if (/早班/.test(context)) return "早班";
  if (/夜班/.test(context)) return "夜班";
  return "未知";
}

function parsePaymentRecords(grid: unknown[][], sheetName: string, date: string) {
  const records: PaymentRecord[] = [];
  grid.forEach((row, rowIndex) => {
    row.forEach((cell, col) => {
      if (text(cell) !== "收款渠道") return;
      const shift = paymentShiftFromContext(grid, rowIndex);
      for (let c = col + 1; c < row.length - 1; c += 2) {
        const channel = cleanChannelName(row[c]);
        if (channel) records.push({ date, sheetName, shift, channel, amount: round2(n(row[c + 1])) });
      }
    });
  });
  return records;
}

function parseWorkbook(workbook: XLSX.WorkBook, rules: CategoryRule[]): ParsedWorkbook {
  const parsedRows: RawItem[] = [];
  const payments: PaymentRecord[] = [];
  const sheets: SheetPreview[] = [];
  workbook.SheetNames.forEach((sheetName, sheetIndex) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;
    const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null });
    const date = extractDate(sheetName, grid[0]?.[0], sheetIndex);
    const sheetPayments = parsePaymentRecords(grid, sheetName, date);
    payments.push(...sheetPayments);
    let sheetRows = 0;
    const groups = findTableGroups(grid, sheetName, rules);
    groups.forEach((group) => {
      let currentCategory = group.defaultCategory;
      for (let i = group.headerRow + 1; i < grid.length; i += 1) {
        const row = grid[i] ?? [];
        const lineText = rowText(row.slice(group.startCol, group.totalAmountCol + 1));
        const lineCategory = categorize(lineText, rules);
        const name = text(row[group.nameCol]);
        const price = n(row[group.priceCol]);
        if (lineCategory !== UNKNOWN_CATEGORY && (!name || !price || /分类|类别|品类|小计|合计|库存|销售|报表/.test(lineText))) currentCategory = lineCategory;
        if (!name || !price || isSummaryLikeName(name)) continue;

        const purchaseRaw = row[group.purchaseCol];
        const morningStockRaw = row[group.morningStockCol];
        const morningSoldRaw = row[group.morningSoldCol];
        const morningAmountRaw = row[group.morningAmountCol];
        const nightStockRaw = row[group.nightStockCol];
        const nightSoldRaw = row[group.nightSoldCol];
        const nightAmountRaw = row[group.nightAmountCol];
        const totalAmountRaw = row[group.totalAmountCol];
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

        const purchase = n(purchaseRaw);
        const morningStock = n(morningStockRaw);
        const morningSold = n(morningSoldRaw);
        const nightStock = n(nightStockRaw);
        const nightSold = n(nightSoldRaw);
        const morningAmount = n(morningAmountRaw) || round2(morningSold * price);
        const nightAmount = n(nightAmountRaw) || round2(nightSold * price);
        const totalAmount = n(totalAmountRaw) || round2(morningAmount + nightAmount);
        const sameDayExpectedNightStock = round2(morningStock + purchase - morningSold);
        const endingStock = hasNightStock ? round2(nightStock - nightSold) : sameDayExpectedNightStock;
        const rawCategory = categorize(`${name} ${currentCategory}`, rules) === UNKNOWN_CATEGORY ? currentCategory : categorize(`${name} ${currentCategory}`, rules);
        parsedRows.push({ date, sheetName, rowIndex: i, rowNumber: i + 1, sourceOrder: sheetIndex * 1_000_000 + group.startCol * 10_000 + i, rawCategory, name, price, purchase, morningStock, morningSold, morningAmount, nightStock, nightSold, nightAmount, totalAmount, endingStock, sameDayExpectedNightStock, sameDayStockDiff: hasNightStock ? round2(nightStock - sameDayExpectedNightStock) : 0, morningAmountDiff: hasMorningSold || hasMorningAmount ? round2(morningAmount - morningSold * price) : 0, nightAmountDiff: hasNightSold || hasNightAmount ? round2(nightAmount - nightSold * price) : 0, totalAmountDiff: hasMorningAmount || hasNightAmount || hasTotalAmount ? round2(totalAmount - morningAmount - nightAmount) : 0, hasMorningStock, hasMorningSold, hasNightStock, hasNightSold, hasMorningAmount, hasNightAmount, hasTotalAmount, morningStockCol: group.morningStockCol, nightStockCol: group.nightStockCol });
        sheetRows += 1;
      }
    });
    sheets.push({ sheetName, date, rowCount: sheetRows, paymentCount: sheetPayments.length, tableCount: groups.length });
  });
  return { rows: parsedRows.sort((a, b) => a.date.localeCompare(b.date) || byOriginalOrder(a, b)), payments: payments.sort((a, b) => `${a.date}-${a.shift}-${a.channel}`.localeCompare(`${b.date}-${b.shift}-${b.channel}`)), sheets };
}

function hasRowIssue(row: StockItem) { return (row.hasNightStock && Math.abs(row.sameDayStockDiff) > 0.01) || Math.abs(row.morningAmountDiff) > 0.01 || Math.abs(row.nightAmountDiff) > 0.01 || Math.abs(row.totalAmountDiff) > 0.01; }
function columnName(row: StockItem, colIndex: number) { const col = `${XLSX.utils.encode_col(colIndex)}列`; if (colIndex === row.morningStockCol) return `${col} 早班库存`; if (colIndex === row.nightStockCol) return `${col} 夜班库存`; return col; }
function makeWriteBack(row: StockItem, colIndex: number, before: number, after: number): WriteBackTarget { const cell = XLSX.utils.encode_cell({ r: row.rowIndex, c: colIndex }); return { key: rowCellKey(row, colIndex), date: row.date, sheetName: row.sheetName, rowIndex: row.rowIndex, rowNumber: row.rowNumber, colIndex, cell, columnName: columnName(row, colIndex), category: row.category, name: row.name, price: row.price, before, after: round2(after) }; }
function writeCell(sheet: XLSX.WorkSheet, rowIndex: number, colIndex: number, value: number) { const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex }); const cell = (sheet[address] ?? {}) as CellWithStyle; delete cell.f; delete cell.w; cell.t = "n"; cell.v = value; sheet[address] = cell; }
function writeBackLocation(issue: Issue) { return issue.writeBack ? `${issue.writeBack.sheetName}!${issue.writeBack.cell}（${issue.writeBack.columnName}）` : "不自动写回"; }
function downloadWorkbook(fileName: string, sheets: Record<string, Record<string, unknown>[]>) { const workbook = XLSX.utils.book_new(); Object.entries(sheets).forEach(([name, data]) => XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(data), name.slice(0, 31))); XLSX.writeFile(workbook, fileName); }

function Card({ title, value, desc, tone = "default" }: { title: string; value: string; desc?: string; tone?: "default" | "danger" }) {
  const cls = tone === "danger" ? "border-red-500 bg-red-600 text-white shadow-sm" : "border bg-white text-slate-900 shadow-sm";
  return <div className={`rounded-2xl p-5 ${cls}`}><div className={`text-sm ${tone === "danger" ? "text-red-100" : "text-slate-500"}`}>{title}</div><div className="mt-2 text-2xl font-bold">{value}</div>{desc ? <div className={`mt-1 text-xs ${tone === "danger" ? "text-red-100" : "text-slate-500"}`}>{desc}</div> : null}</div>;
}
function Badge({ children, tone = "gray" }: { children: ReactNode; tone?: "gray" | "green" | "red" | "yellow" }) { const cls = { gray: "bg-slate-100 text-slate-700", green: "bg-green-100 text-green-700", red: "bg-red-100 text-red-700", yellow: "bg-yellow-100 text-yellow-800" }[tone]; return <span className={`rounded-full px-2 py-1 text-xs font-medium ${cls}`}>{children}</span>; }

export default function InventoryPageConfigurable() {
  const [tab, setTab] = useState<Tab>("overview");
  const [rules, setRules] = useState<CategoryRule[]>(DEFAULT_RULES);
  const [rawRows, setRawRows] = useState<RawItem[]>([]);
  const [paymentRecords, setPaymentRecords] = useState<PaymentRecord[]>([]);
  const [sheetPreview, setSheetPreview] = useState<SheetPreview[]>([]);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [fileName, setFileName] = useState("");
  const [corrections, setCorrections] = useState<Record<string, Correction>>({});

  const rows = useMemo<StockItem[]>(() => rawRows.map((row) => ({ ...row, category: categorize(`${row.name} ${row.rawCategory}`, rules) })), [rawRows, rules]);
  async function handleFile(file: File) { const nextWorkbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true, cellFormula: true, cellNF: true, cellStyles: true }); const parsed = parseWorkbook(nextWorkbook, rules); setWorkbook(nextWorkbook); setRawRows(parsed.rows); setPaymentRecords(parsed.payments); setSheetPreview(parsed.sheets); setFileName(file.name); setCorrections({}); }

  const salesRows = useMemo(() => rows.filter((row) => ruleByName(rules, row.category)?.includeSales ?? true), [rows, rules]);
  const stockCheckRows = useMemo(() => rows.filter((row) => ruleByName(rules, row.category)?.includeStockCheck ?? false), [rows, rules]);
  const inventoryCategoryNames = useMemo(() => new Set(rules.filter((rule) => rule.includeInventory).map((rule) => rule.name)), [rules]);
  const paymentChannels = useMemo(() => { const channels: string[] = []; paymentRecords.forEach((record) => { if (record.channel && !channels.includes(record.channel)) channels.push(record.channel); }); return channels; }, [paymentRecords]);
  const paymentDaily = useMemo<PaymentDailySummary[]>(() => { const byDate = new Map<string, PaymentRecord[]>(); paymentRecords.forEach((record) => byDate.set(record.date, [...(byDate.get(record.date) ?? []), record])); return Array.from(byDate.entries()).map(([date, records]) => { const totalRecords = records.filter((record) => record.shift === "合计"); const source = totalRecords.length ? totalRecords : records.filter((record) => record.shift !== "合计"); const totals = makeEmptyTotals(paymentChannels); source.forEach((record) => { totals[record.channel] = round2((totals[record.channel] ?? 0) + record.amount); }); return { date, totals, total: round2(paymentChannels.reduce((sum, channel) => sum + (totals[channel] ?? 0), 0)) }; }).sort((a, b) => a.date.localeCompare(b.date)); }, [paymentRecords, paymentChannels]);
  const paymentSummary = useMemo(() => paymentChannels.map((channel) => ({ channel, amount: round2(paymentDaily.reduce((sum, day) => sum + (day.totals[channel] ?? 0), 0)) })), [paymentChannels, paymentDaily]);
  const reportProgressDate = useMemo(() => [...rows.map((row) => row.date), ...paymentDaily.filter((item) => item.total !== 0).map((item) => item.date)].sort().at(-1) ?? "", [rows, paymentDaily]);

  const daily = useMemo<DailySummary[]>(() => { const map = new Map<string, DailySummary>(); salesRows.forEach((row) => { const item = map.get(row.date) ?? { date: row.date, purchase: 0, morningSold: 0, nightSold: 0, totalSold: 0, morningAmount: 0, nightAmount: 0, totalAmount: 0, abnormalCount: 0 }; item.purchase += row.purchase; item.morningSold += row.morningSold; item.nightSold += row.nightSold; item.totalSold += row.morningSold + row.nightSold; item.morningAmount += row.morningAmount; item.nightAmount += row.nightAmount; item.totalAmount += row.totalAmount; item.abnormalCount += hasRowIssue(row) ? 1 : 0; map.set(row.date, item); }); return Array.from(map.values()).map((item) => ({ ...item, purchase: round2(item.purchase), morningSold: round2(item.morningSold), nightSold: round2(item.nightSold), totalSold: round2(item.totalSold), morningAmount: round2(item.morningAmount), nightAmount: round2(item.nightAmount), totalAmount: round2(item.totalAmount) })).sort((a, b) => a.date.localeCompare(b.date)); }, [salesRows]);

  const productSummary = useMemo<ProductSummary[]>(() => { const map = new Map<string, ProductSummary>(); salesRows.forEach((row) => { const key = productKey(row.category, row.name, row.price); const item = map.get(key) ?? { category: row.category, name: row.name, price: row.price, purchase: 0, morningSold: 0, nightSold: 0, totalSold: 0, totalAmount: 0, endingStock: 0, latestDate: row.date, latestRow: row, sourceOrder: row.sourceOrder }; item.purchase += row.purchase; item.morningSold += row.morningSold; item.nightSold += row.nightSold; item.totalSold += row.morningSold + row.nightSold; item.totalAmount += row.totalAmount; if (row.date >= item.latestDate) { item.latestDate = row.date; item.latestRow = row; item.endingStock = row.endingStock; } item.sourceOrder = Math.min(item.sourceOrder, row.sourceOrder); map.set(key, item); }); return Array.from(map.values()).map((item) => ({ ...item, purchase: round2(item.purchase), morningSold: round2(item.morningSold), nightSold: round2(item.nightSold), totalSold: round2(item.totalSold), totalAmount: round2(item.totalAmount), endingStock: round2(item.endingStock) })).sort((a, b) => a.category.localeCompare(b.category) || a.sourceOrder - b.sourceOrder); }, [salesRows]);
  const inventoryRows = useMemo(() => productSummary.filter((item) => inventoryCategoryNames.has(item.category)), [productSummary, inventoryCategoryNames]);
  const categorySummary = useMemo<CategorySummary[]>(() => { const map = new Map<string, CategorySummary>(); productSummary.forEach((item) => { const current = map.get(item.category) ?? { category: item.category, totalSold: 0, totalAmount: 0, count: 0, sourceOrder: item.sourceOrder }; current.totalSold += item.totalSold; current.totalAmount += item.totalAmount; current.count += 1; current.sourceOrder = Math.min(current.sourceOrder, item.sourceOrder); map.set(item.category, current); }); return Array.from(map.values()).map((item) => ({ ...item, totalSold: round2(item.totalSold), totalAmount: round2(item.totalAmount) })).sort((a, b) => a.sourceOrder - b.sourceOrder); }, [productSummary]);

  const issues = useMemo<Issue[]>(() => { const out: Issue[] = []; stockCheckRows.forEach((row) => { if (row.hasNightStock && Math.abs(row.sameDayStockDiff) > 0.01) { const type: IssueType = row.morningSold > 0 && Math.abs(row.sameDayStockDiff - row.morningSold) < 0.01 ? "疑似未扣早班销量" : "早班交夜班"; const target = makeWriteBack(row, row.nightStockCol, row.nightStock, row.sameDayExpectedNightStock); out.push({ id: `${type}_${target.key}`, type, date: row.date, category: row.category, name: row.name, price: row.price, expected: row.sameDayExpectedNightStock, actual: row.nightStock, diff: row.sameDayStockDiff, note: type === "疑似未扣早班销量" ? "夜班库存比正确值多出的数量，正好等于早班销量。" : "夜班库存应等于：早班库存 + 进货 - 早班售卖数量。", writeBack: target }); } if (Math.abs(row.morningAmountDiff) > 0.01) out.push({ id: `金额核对_早班_${rowKey(row)}`, type: "金额核对", date: row.date, category: row.category, name: row.name, price: row.price, expected: round2(row.morningSold * row.price), actual: row.morningAmount, diff: row.morningAmountDiff, note: "早班金额不等于：早班售卖数量 × 价格。" }); if (Math.abs(row.nightAmountDiff) > 0.01) out.push({ id: `金额核对_夜班_${rowKey(row)}`, type: "金额核对", date: row.date, category: row.category, name: row.name, price: row.price, expected: round2(row.nightSold * row.price), actual: row.nightAmount, diff: row.nightAmountDiff, note: "夜班金额不等于：夜班售卖数量 × 价格。" }); }); const byProduct = new Map<string, StockItem[]>(); stockCheckRows.forEach((row) => byProduct.set(productKey(row.category, row.name, row.price), [...(byProduct.get(productKey(row.category, row.name, row.price)) ?? []), row])); byProduct.forEach((list) => { const sorted = [...list].sort((a, b) => a.date.localeCompare(b.date)); for (let i = 0; i < sorted.length - 1; i += 1) { const current = sorted[i]; const next = sorted[i + 1]; if (!current.hasNightStock || !next.hasMorningStock) continue; const expected = round2(current.nightStock - current.nightSold); const diff = round2(next.morningStock - expected); if (Math.abs(diff) > 0.01) { const target = makeWriteBack(next, next.morningStockCol, next.morningStock, expected); out.push({ id: `次日接班_${target.key}_${current.date}`, type: "次日接班", date: current.date, nextDate: next.date, category: current.category, name: current.name, price: current.price, expected, actual: next.morningStock, diff, note: "次日早班库存应等于：前一日夜班库存 - 前一日夜班售卖数量。", writeBack: target }); } } }); return out.sort((a, b) => `${a.date}-${a.category}-${a.name}-${a.type}`.localeCompare(`${b.date}-${b.category}-${b.name}-${b.type}`)); }, [stockCheckRows]);
  const productIssueCount = useMemo(() => { const map = new Map<string, number>(); issues.forEach((issue) => { const key = productKey(issue.category, issue.name, issue.price); map.set(key, (map.get(key) ?? 0) + 1); }); return map; }, [issues]);

  const stats = { totalIssueCount: issues.length, inventoryIssueCount: issues.filter((issue) => issue.type !== "金额核对").length, amountIssueCount: issues.filter((issue) => issue.type === "金额核对").length, writableCount: issues.filter((issue) => issue.writeBack).length, preparedCount: Object.keys(corrections).length };
  const totalAmount = daily.reduce((sum, item) => sum + item.totalAmount, 0);
  const totalSold = daily.reduce((sum, item) => sum + item.totalSold, 0);
  const totalPaymentAmount = paymentSummary.reduce((sum, item) => sum + item.amount, 0);

  function makeCorrection(target: WriteBackTarget, reason: string): Correction { return { ...target, after: round2(target.after), diff: round2(target.after - target.before), reason, correctedAt: new Date().toLocaleString("zh-CN", { hour12: false }) }; }
  function setCorrection(target: WriteBackTarget, reason: string) { setCorrections((old) => ({ ...old, [target.key]: makeCorrection(target, reason) })); }
  function undoCorrection(key: string) { setCorrections((old) => { const copy = { ...old }; delete copy[key]; return copy; }); }
  function applyIssueCorrection(issue: Issue) { if (!issue.writeBack) return alert("金额核对异常只提示，不自动改原表。"); setCorrection(issue.writeBack, issue.note); }
  function applyAllWritable() { const writable = issues.filter((issue) => issue.writeBack && !corrections[issue.writeBack.key]); if (!writable.length) return alert("没有新的库存异常需要准备写回。"); if (!confirm(`确定准备写回 ${writable.length} 条库存异常吗？`)) return; setCorrections((old) => writable.reduce((acc, issue) => issue.writeBack ? { ...acc, [issue.writeBack.key]: makeCorrection(issue.writeBack, issue.note) } : acc, old)); }
  function downloadCorrectedOriginal() { if (!workbook) return alert("请先上传原始报表 Excel"); const list = Object.values(corrections); if (!list.length) return alert("还没有纠正内容。请先在交接核对里准备写回。"); const fixed = XLSX.read(XLSX.write(workbook, { bookType: "xlsx", type: "array" }), { type: "array", cellDates: true, cellFormula: true, cellNF: true, cellStyles: true }); list.forEach((item) => { const sheet = fixed.Sheets[item.sheetName]; if (sheet) writeCell(sheet, item.rowIndex, item.colIndex, item.after); }); XLSX.writeFile(fixed, `已纠正-${safeFileName(fileName || "手工报表")}.xlsx`); }
  function updateRule(id: string, patch: Partial<CategoryRule>) { setRules((old) => old.map((rule) => rule.id === id ? { ...rule, ...patch } : rule)); }
  function addRule() { setRules((old) => [...old, { id: `rule_${Date.now()}`, name: "新分类", keywords: "", includeSales: true, includeStockCheck: false, includeInventory: true }]); }
  function removeRule(id: string) { setRules((old) => old.filter((rule) => rule.id !== id)); }

  function exportSalesSummaryReport() { if (!productSummary.length) return alert("请先上传原始报表 Excel"); downloadWorkbook(`${STORE_NAME}-销售汇总报表.xlsx`, { 销售汇总报表: productSummary.map((item, index) => ({ 序号: index + 1, 商品分类: item.category, 商品名称: item.name, 平均价格: item.price, 数量: item.totalSold, 金额: item.totalAmount })), 分类汇总: categorySummary.map((item, index) => ({ 序号: index + 1, 商品分类: item.category, 品规数: item.count, 销量: item.totalSold, 销售金额: item.totalAmount })), 收款渠道月统计: paymentSummary.map((item, index) => ({ 序号: index + 1, 收款渠道: item.channel, 月收入: item.amount })) }); }
  function exportExceptionReport() { if (!issues.length) return alert("当前没有异常。"); downloadWorkbook("报表异常核对报表.xlsx", { 异常核对报表: issues.map((item, index) => ({ 序号: index + 1, 异常类型: item.type, 商品分类: item.category, 日期: item.date, 次日: item.nextDate ?? "", 商品名称: item.name, 应为: item.expected, 实际: item.actual, 差异: item.diff, 写回位置: writeBackLocation(item), 说明: item.note })) }); }
  function exportAutoInventoryReport() { if (!inventoryRows.length) return alert("当前没有可自动盘点的商品。请在分类规则里勾选自动盘点。"); downloadWorkbook(`${reportProgressDate || "本月"}-自动库存盘点表.xlsx`, { 自动库存盘点: inventoryRows.map((item, index) => ({ 序号: index + 1, 分类: item.category, 商品名称: item.name, 价格: item.price, 本月进货: item.purchase, 早班销量: item.morningSold, 夜班销量: item.nightSold, 本月总销量: item.totalSold, 月末账面应剩: item.endingStock, 最新日期: item.latestDate, 异常数: productIssueCount.get(productKey(item.category, item.name, item.price)) ?? 0, 状态: (productIssueCount.get(productKey(item.category, item.name, item.price)) ?? 0) > 0 ? "需核对" : "正常" })), 异常盘点: inventoryRows.filter((item) => (productIssueCount.get(productKey(item.category, item.name, item.price)) ?? 0) > 0).map((item, index) => ({ 序号: index + 1, 分类: item.category, 商品名称: item.name, 价格: item.price, 月末账面应剩: item.endingStock, 最新日期: item.latestDate, 异常数: productIssueCount.get(productKey(item.category, item.name, item.price)) ?? 0 })) }); }

  return <main className="min-h-screen p-4 md:p-8"><div className="mx-auto max-w-7xl">
    <header className="mb-6 rounded-3xl bg-slate-900 p-6 text-white shadow-sm"><p className="text-sm text-slate-300">Configurable Report Reconciliation</p><h1 className="mt-1 text-2xl font-bold md:text-4xl">报表自动核对助手</h1><p className="mt-2 max-w-3xl text-sm text-slate-300">导入 Excel 后自动计算账面库存盘点、销售汇总、收款统计和异常核对。</p></header>
    <section className="mb-6 rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">导入手工报表 Excel</h2><p className="mt-1 text-sm text-slate-500">系统会自动识别商品、分类和收款渠道，并生成自动库存盘点结果。无需手动输入实点数量。</p></div><label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-slate-900 px-4 py-3 text-sm font-medium text-white">选择报表文件<input className="hidden" type="file" accept=".xlsx,.xls" onChange={(event: ChangeEvent<HTMLInputElement>) => event.target.files?.[0] && handleFile(event.target.files[0])} /></label></div>{fileName ? <p className="mt-3 text-sm text-slate-600">已导入：<b>{fileName}</b>，共解析 {rows.length} 条商品明细，{categorySummary.length} 个分类，{paymentChannels.length} 个收款渠道。本月做到：<b>{reportProgressDate || "未识别"}</b>。</p> : null}</section>
    <nav className="mb-6 flex flex-wrap gap-2">{tabs.map(([key, label]) => <button key={key} onClick={() => setTab(key)} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === key ? "bg-slate-900 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>{label}</button>)}</nav>
    {!rows.length ? <section className="rounded-2xl border bg-white p-6 shadow-sm"><h2 className="text-lg font-bold">等待导入报表</h2><p className="mt-2 text-sm text-slate-600">可先到“分类规则”确认哪些分类进入自动库存盘点。</p></section> : null}

    {tab === "overview" && rows.length > 0 ? <section className="space-y-6"><div className="grid gap-4 md:grid-cols-5"><Card tone="danger" title="总异常数" value={`${stats.totalIssueCount} 条`} desc="库存 + 金额" /><Card title="本月做到日期" value={reportProgressDate || "未识别"} desc={`自动盘点 ${inventoryRows.length} 个品规`} /><Card title="商品品规" value={`${productSummary.length} 个`} desc={`${categorySummary.length} 个分类`} /><Card title="总销量" value={`${round2(totalSold)} 件`} /><Card title="总销售金额" value={`¥${money(totalAmount)}`} desc={`收款 ¥${money(totalPaymentAmount)}`} /></div><div className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">分类汇总</h2><div className="mt-3 grid gap-3 md:grid-cols-4">{categorySummary.map((item) => <div key={item.category} className="rounded-xl bg-slate-50 p-4"><div className="font-bold">{item.category}</div><div className="mt-1 text-sm text-slate-600">品规 {item.count} / 销量 {item.totalSold} / ¥{money(item.totalAmount)}</div></div>)}</div></div><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-wrap gap-2"><button onClick={downloadCorrectedOriginal} className="rounded-xl bg-green-700 px-4 py-2 text-sm font-medium text-white">下载已纠正原表</button><button onClick={exportAutoInventoryReport} className="rounded-xl border px-4 py-2 text-sm font-medium">导出自动库存盘点表</button><button onClick={exportSalesSummaryReport} className="rounded-xl border px-4 py-2 text-sm font-medium">导出销售汇总报表</button><button onClick={exportExceptionReport} className="rounded-xl border px-4 py-2 text-sm font-medium">导出异常核对报表</button></div></div></section> : null}

    {tab === "rules" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex items-center justify-between"><div><h2 className="text-lg font-bold">分类规则配置</h2><p className="mt-1 text-sm text-slate-500">关键词用逗号分隔。勾选“自动盘点”的分类会进入账面库存盘点结果。</p></div><button onClick={addRule} className="rounded-xl bg-slate-900 px-4 py-2 text-sm text-white">新增分类</button></div><div className="mt-4 space-y-3">{rules.map((rule) => <div key={rule.id} className="rounded-xl border p-4"><div className="grid gap-3 md:grid-cols-5"><input className="rounded-xl border p-2" value={rule.name} onChange={(event) => updateRule(rule.id, { name: event.target.value })} /><input className="rounded-xl border p-2 md:col-span-2" value={rule.keywords} onChange={(event) => updateRule(rule.id, { keywords: event.target.value })} /><label className="text-sm"><input type="checkbox" checked={rule.includeSales} onChange={(event) => updateRule(rule.id, { includeSales: event.target.checked })} /> 销售汇总</label><label className="text-sm"><input type="checkbox" checked={rule.includeInventory} onChange={(event) => updateRule(rule.id, { includeInventory: event.target.checked })} /> 自动盘点</label></div><div className="mt-2 flex gap-4 text-sm"><label><input type="checkbox" checked={rule.includeStockCheck} onChange={(event) => updateRule(rule.id, { includeStockCheck: event.target.checked })} /> 库存核对</label><button className="text-red-600" onClick={() => removeRule(rule.id)}>删除</button></div></div>)}</div></section> : null}

    {tab === "detect" && rows.length > 0 ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">工作表识别预览</h2><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["Sheet", "日期", "商品行", "商品表", "收款记录"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{sheetPreview.map((sheet) => <tr key={sheet.sheetName} className="border-t"><td className="p-3">{sheet.sheetName}</td><td className="p-3">{sheet.date}</td><td className="p-3">{sheet.rowCount}</td><td className="p-3">{sheet.tableCount}</td><td className="p-3">{sheet.paymentCount}</td></tr>)}</tbody></table></div></section> : null}

    {tab === "daily" && rows.length > 0 ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">每日报表汇总</h2><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["日期", "进货", "早班销量", "夜班销量", "总销量", "早班金额", "夜班金额", "总金额", "异常"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{daily.map((item) => <tr key={item.date} className="border-t"><td className="p-3">{item.date}</td><td className="p-3">{item.purchase}</td><td className="p-3">{item.morningSold}</td><td className="p-3">{item.nightSold}</td><td className="p-3">{item.totalSold}</td><td className="p-3">¥{money(item.morningAmount)}</td><td className="p-3">¥{money(item.nightAmount)}</td><td className="p-3">¥{money(item.totalAmount)}</td><td className="p-3">{item.abnormalCount ? <Badge tone="red">{item.abnormalCount}</Badge> : <Badge tone="green">正常</Badge>}</td></tr>)}</tbody></table></div></section> : null}

    {tab === "summary" && rows.length > 0 ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">商品总汇</h2><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["分类", "商品", "价格", "月进货", "早班销量", "夜班销量", "月销量", "月销售金额", "月末结存"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{productSummary.map((item) => <tr key={`${item.category}-${item.name}-${item.price}`} className="border-t"><td className="p-3">{item.category}</td><td className="p-3">{item.name}</td><td className="p-3">¥{money(item.price)}</td><td className="p-3">{item.purchase}</td><td className="p-3">{item.morningSold}</td><td className="p-3">{item.nightSold}</td><td className="p-3">{item.totalSold}</td><td className="p-3">¥{money(item.totalAmount)}</td><td className="p-3">{item.endingStock}</td></tr>)}</tbody></table></div></section> : null}

    {tab === "handover" && rows.length > 0 ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">交接和金额异常</h2><p className="mt-1 text-sm text-slate-500">只检查分类规则中勾选“库存核对”的分类。金额核对只提示。</p></div><button onClick={applyAllWritable} className="rounded-xl bg-yellow-700 px-4 py-2 text-sm text-white">准备写回全部库存异常</button></div><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["类型", "分类", "日期", "次日", "商品", "应为", "实际", "差异", "写回位置", "说明", "操作"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{issues.map((item) => { const prepared = item.writeBack ? corrections[item.writeBack.key] : null; return <tr key={item.id} className="border-t"><td className="p-3"><Badge tone={item.type === "金额核对" ? "red" : "yellow"}>{item.type}</Badge></td><td className="p-3">{item.category}</td><td className="p-3">{item.date}</td><td className="p-3">{item.nextDate ?? ""}</td><td className="p-3">{item.name}</td><td className="p-3">{item.expected}</td><td className="p-3">{item.actual}</td><td className="p-3">{item.diff}</td><td className="p-3">{writeBackLocation(item)}</td><td className="p-3">{item.note}</td><td className="p-3">{item.writeBack ? prepared ? <button className="rounded-lg border px-3 py-1 text-xs" onClick={() => undoCorrection(item.writeBack!.key)}>撤销</button> : <button className="rounded-lg bg-slate-900 px-3 py-1 text-xs text-white" onClick={() => applyIssueCorrection(item)}>准备写回</button> : <Badge>只提示</Badge>}</td></tr>; })}</tbody></table></div></section> : null}

    {tab === "inventory" && rows.length > 0 ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">自动库存盘点</h2><p className="mt-1 text-sm text-slate-500">本月做到：{reportProgressDate || "未识别"}。这里自动显示账面月末应剩，不需要手动输入现场实点。</p></div><button onClick={exportAutoInventoryReport} className="rounded-xl border px-4 py-2 text-sm font-medium">导出自动库存盘点表</button></div>{!inventoryRows.length ? <div className="mt-4 rounded-xl bg-yellow-50 p-4 text-sm text-yellow-900">没有可自动盘点的商品。请到“分类规则”里勾选“自动盘点”。</div> : null}<div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["序号", "分类", "商品", "价格", "本月进货", "早班销量", "夜班销量", "总销量", "月末账面应剩", "最新日期", "异常数", "状态"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{inventoryRows.map((item, index) => { const issueCount = productIssueCount.get(productKey(item.category, item.name, item.price)) ?? 0; return <tr key={`${item.category}-${item.name}-${item.price}`} className="border-t"><td className="p-3 text-slate-500">{index + 1}</td><td className="p-3">{item.category}</td><td className="p-3">{item.name}</td><td className="p-3">¥{money(item.price)}</td><td className="p-3">{item.purchase}</td><td className="p-3">{item.morningSold}</td><td className="p-3">{item.nightSold}</td><td className="p-3">{item.totalSold}</td><td className="p-3 font-bold">{item.endingStock}</td><td className="p-3">{item.latestDate}</td><td className="p-3">{issueCount}</td><td className="p-3">{issueCount > 0 ? <Badge tone="red">需核对</Badge> : <Badge tone="green">正常</Badge>}</td></tr>; })}</tbody></table></div></section> : null}

    {tab === "payment" && rows.length > 0 ? <section className="space-y-5"><div className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">收款渠道月统计</h2><div className="mt-3 grid gap-3 md:grid-cols-5">{paymentSummary.map((item) => <div key={item.channel} className="rounded-xl bg-slate-50 p-4"><div className="font-bold">{item.channel}</div><div className="mt-1 text-sm text-slate-600">¥{money(item.amount)}</div></div>)}</div><div className="mt-4 rounded-xl bg-slate-900 p-4 text-white"><div className="text-sm text-slate-300">全部渠道合计</div><div className="mt-1 text-2xl font-bold">¥{money(totalPaymentAmount)}</div></div></div><div className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">每日收款明细</h2><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["日期", ...paymentChannels, "合计"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{paymentDaily.map((item) => <tr key={item.date} className="border-t"><td className="p-3">{item.date}</td>{paymentChannels.map((channel) => <td key={channel} className="p-3">¥{money(item.totals[channel] ?? 0)}</td>)}<td className="p-3 font-bold">¥{money(item.total)}</td></tr>)}<tr className="border-t bg-slate-900 text-white"><td className="p-3 font-bold">渠道合计</td>{paymentSummary.map((item) => <td key={item.channel} className="p-3 font-bold">¥{money(item.amount)}</td>)}<td className="p-3 font-bold">¥{money(totalPaymentAmount)}</td></tr></tbody></table></div></div></section> : null}
  </div></main>;
}
