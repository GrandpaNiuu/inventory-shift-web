"use client";

import { useMemo, useState, type ChangeEvent, type ReactNode } from "react";
import * as XLSX from "xlsx";

type Tab = "overview" | "daily" | "summary" | "handover" | "count" | "payment";
type IssueType = "疑似未扣早班销量" | "早班交夜班" | "次日接班" | "金额核对";
type PaymentShift = "早班" | "夜班" | "合计" | "未知";

type StockRow = {
  date: string;
  sheetName: string;
  rowIndex: number;
  rowNumber: number;
  no: number | string;
  category: string;
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
  morningStockCol: number;
  nightStockCol: number;
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
  category: string;
  name: string;
  price: number;
  purchase: number;
  morningSold: number;
  nightSold: number;
  totalSold: number;
  totalAmount: number;
  endingStock: number;
};

type CategorySummary = {
  category: string;
  totalSold: number;
  totalAmount: number;
  count: number;
};

type PaymentRecord = {
  date: string;
  sheetName: string;
  shift: PaymentShift;
  channel: string;
  amount: number;
};

type PaymentDailySummary = {
  date: string;
  totals: Record<string, number>;
  total: number;
};

type PaymentSummary = {
  channel: string;
  amount: number;
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
  category: string;
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
  category: string;
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

type TableGroup = {
  headerRow: number;
  startCol: number;
  noCol: number;
  nameCol: number;
  priceCol: number;
  purchaseCol: number;
  morningStockCol: number;
  morningSoldCol: number;
  morningAmountCol: number;
  nightStockCol: number;
  nightSoldCol: number;
  nightAmountCol: number;
  totalAmountCol: number;
  defaultCategory: string;
};

type ParsedWorkbook = {
  rows: StockRow[];
  payments: PaymentRecord[];
};

const nav: Array<[Tab, string]> = [
  ["overview", "总览"],
  ["daily", "每日报表"],
  ["summary", "商品总汇"],
  ["handover", "交接核对"],
  ["count", "现场盘点"],
  ["payment", "收款统计"]
];

const STORE_NAME = "SY11";
const DEFAULT_CATEGORY = "杂项";
const CIGARETTE_CATEGORY = "香烟(酒店库存)";

const HEADER_STYLE = {
  font: { bold: true, color: { rgb: "FFFFFF" } },
  alignment: { horizontal: "center", vertical: "center" },
  fill: { fgColor: { rgb: "1F2937" } },
  border: {
    top: { style: "thin", color: { rgb: "CBD5E1" } },
    bottom: { style: "thin", color: { rgb: "CBD5E1" } },
    left: { style: "thin", color: { rgb: "CBD5E1" } },
    right: { style: "thin", color: { rgb: "CBD5E1" } }
  }
};
const TITLE_STYLE = { font: { bold: true, sz: 18 }, alignment: { horizontal: "center", vertical: "center" } };
const META_STYLE = { font: { bold: true }, alignment: { horizontal: "left", vertical: "center" }, fill: { fgColor: { rgb: "F8FAFC" } } };
const CATEGORY_STYLE = { font: { bold: true, color: { rgb: "1E3A8A" } }, fill: { fgColor: { rgb: "DBEAFE" } } };
const SUBTOTAL_STYLE = { font: { bold: true }, fill: { fgColor: { rgb: "FEF3C7" } } };
const TOTAL_STYLE = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "111827" } } };

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

function text(value: unknown) {
  return String(value ?? "").trim();
}

function rowText(row: unknown[] | undefined) {
  return (row ?? []).map((cell) => text(cell)).filter(Boolean).join(" ");
}

function rowKey(row: StockRow, date = row.date) {
  return `${date}__${row.sheetName}__${row.rowNumber}__${row.category}__${row.name}__${row.price}`;
}

function countKey(row: StockRow) {
  return `${row.category}__${row.name}__${row.price}`;
}

function correctionKey(sheetName: string, rowIndex: number, colIndex: number) {
  return `${sheetName}__${rowIndex}__${colIndex}`;
}

function rowCellKey(row: StockRow, colIndex: number) {
  return correctionKey(row.sheetName, row.rowIndex, colIndex);
}

function isMissedMorningDeduction(row: StockRow) {
  return row.hasNightStock && row.morningSold > 0 && Math.abs(row.sameDayStockDiff - row.morningSold) < 0.01;
}

function correctedNightStock(row: StockRow) {
  return row.sameDayExpectedNightStock;
}

function safeFileName(name: string) {
  return name.replace(/\.xlsx?$/i, "").replace(/[\\/:*?\"<>|]/g, "-");
}

function reportDateRange(daily: DailySummary[]) {
  if (!daily.length) return "";
  return daily.length === 1 ? daily[0].date : `${daily[0].date} 至 ${daily[daily.length - 1].date}`;
}

function queryTime() {
  return new Date().toLocaleString("zh-CN", { hour12: false });
}

function isAmountLikeLabel(value: string) {
  return /^(金额|收款金额|合计|总计|早班|夜班|收款渠道|渠道|售卖金额)$/.test(value) || /^\d+(\.\d+)?$/.test(value);
}

function cleanChannelName(value: unknown) {
  const channel = text(value).replace(/\s/g, "");
  return channel && !isAmountLikeLabel(channel) ? channel : "";
}

function inferCategoryFromText(source: string) {
  const value = source.replace(/\s/g, "");
  if (!value) return null;
  if (/打火机|火机|火柴|点烟器/i.test(value)) return "打火机";
  if (/饮料|矿泉水|纯净水|可乐|雪碧|芬达|苏打水|气泡水|红牛|东鹏|脉动|外星人|农夫山泉|怡宝|百岁山|王老吉|加多宝|茶饮|冰红茶|绿茶|乌龙茶|咖啡|牛奶|酸奶|椰汁|果汁|柠檬茶|阿萨姆|海之言|娃哈哈|康师傅.*茶|康师傅.*水|康师傅.*奶茶/i.test(value)) return "饮料";
  if (/杂项|其他|小商品|日用品|百货|纸巾|扑克牌|牙刷|牙膏|剃须|充电器|数据线|雨伞|方便面|泡面|螺蛳粉|零食|口香糖/i.test(value)) return DEFAULT_CATEGORY;
  if (/烟报|香烟|卷烟|烟草|烟品|中华|芙蓉王|利群|玉溪|黄鹤楼|云烟|南京|双喜|红塔山|白沙|娇子|黄金叶|苏烟|泰山|七匹狼|中南海|牡丹|贵烟|真龙|钻石|煊赫门|万宝路|黄山|长白山|延安|兰州|宽窄|荷花|天子|红河|红金龙|金圣|人民大会堂|细支|中支/i.test(value)) return CIGARETTE_CATEGORY;
  return null;
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

function groupDefaultCategory(grid: unknown[][], headerRow: number, startCol: number, endCol: number, sheetName: string) {
  const samples: string[] = [sheetName];
  for (let r = 0; r <= headerRow; r += 1) {
    for (let c = startCol; c <= endCol; c += 1) samples.push(text(grid[r]?.[c]));
  }
  return inferCategoryFromText(samples.join(" ")) ?? DEFAULT_CATEGORY;
}

function findTableGroups(grid: unknown[][], sheetName: string) {
  const groups: TableGroup[] = [];
  const seen = new Set<string>();
  grid.forEach((row, headerRow) => {
    row.forEach((cell, col) => {
      const header = text(cell);
      const priceHeader = text(row[col + 1]);
      if (!/^(商品名称|售卖商品)$/.test(header) || !/^(价格|售价)$/.test(priceHeader)) return;
      const startCol = Math.max(0, col - 1);
      const key = `${headerRow}_${startCol}`;
      if (seen.has(key)) return;
      seen.add(key);
      groups.push({
        headerRow,
        startCol,
        noCol: startCol,
        nameCol: col,
        priceCol: col + 1,
        purchaseCol: col + 2,
        morningStockCol: col + 3,
        morningSoldCol: col + 4,
        morningAmountCol: col + 5,
        nightStockCol: col + 6,
        nightSoldCol: col + 7,
        nightAmountCol: col + 8,
        totalAmountCol: col + 9,
        defaultCategory: groupDefaultCategory(grid, headerRow, startCol, col + 9, sheetName)
      });
    });
  });
  return groups;
}

function isSummaryLikeName(name: string) {
  return /合计|收款|备注|售卖金额|售卖数量|渠道|小计|总计/.test(name);
}

function paymentShiftFromContext(grid: unknown[][], rowIndex: number) {
  const current = rowText(grid[rowIndex]);
  if (/合计售卖金额|总计|合计/.test(current)) return "合计" as PaymentShift;
  if (/早班/.test(current)) return "早班" as PaymentShift;
  if (/夜班/.test(current)) return "夜班" as PaymentShift;
  const context = `${rowText(grid[rowIndex - 2])} ${rowText(grid[rowIndex - 1])} ${current}`;
  if (/合计售卖金额|总计|合计/.test(context)) return "合计" as PaymentShift;
  if (/早班/.test(context)) return "早班" as PaymentShift;
  if (/夜班/.test(context)) return "夜班" as PaymentShift;
  return "未知" as PaymentShift;
}

function parsePaymentRecords(grid: unknown[][], sheetName: string, date: string) {
  const records: PaymentRecord[] = [];
  grid.forEach((row, rowIndex) => {
    row.forEach((cell, col) => {
      if (text(cell) !== "收款渠道") return;
      const shift = paymentShiftFromContext(grid, rowIndex);
      for (let c = col + 1; c < row.length - 1; c += 2) {
        const channel = cleanChannelName(row[c]);
        if (!channel) continue;
        records.push({ date, sheetName, shift, channel, amount: round2(n(row[c + 1])) });
      }
    });
  });
  return records;
}

function parseWorkbook(workbook: XLSX.WorkBook): ParsedWorkbook {
  const parsedRows: StockRow[] = [];
  const payments: PaymentRecord[] = [];
  workbook.SheetNames.forEach((sheetName, sheetIndex) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;
    const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null });
    const date = extractDate(sheetName, grid[0]?.[0], sheetIndex);
    payments.push(...parsePaymentRecords(grid, sheetName, date));
    findTableGroups(grid, sheetName).forEach((group) => {
      let currentCategory = group.defaultCategory;
      for (let i = group.headerRow + 1; i < grid.length; i += 1) {
        const row = grid[i] ?? [];
        const lineText = rowText(row.slice(group.startCol, group.totalAmountCol + 1));
        const lineCategory = inferCategoryFromText(lineText);
        const name = text(row[group.nameCol]);
        const price = n(row[group.priceCol]);
        if (lineCategory && (!name || !price || /分类|类别|品类|小计|合计|库存|销售|报表/.test(lineText))) currentCategory = lineCategory;
        if (!name || !price || isSummaryLikeName(name)) continue;
        const category = inferCategoryFromText(name) ?? currentCategory ?? DEFAULT_CATEGORY;
        const purchaseRaw = row[group.purchaseCol];
        const morningStockRaw = row[group.morningStockCol];
        const morningSoldRaw = row[group.morningSoldCol];
        const morningAmountRaw = row[group.morningAmountCol];
        const nightStockRaw = row[group.nightStockCol];
        const nightSoldRaw = row[group.nightSoldCol];
        const nightAmountRaw = row[group.nightAmountCol];
        const totalAmountRaw = row[group.totalAmountCol];
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
        parsedRows.push({
          date,
          sheetName,
          rowIndex: i,
          rowNumber: i + 1,
          no: row[group.noCol] as number | string,
          category,
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
          hasTotalAmount,
          morningStockCol: group.morningStockCol,
          nightStockCol: group.nightStockCol
        });
      }
    });
  });
  return {
    rows: parsedRows.sort((a, b) => `${a.date}-${a.category}-${a.rowNumber}`.localeCompare(`${b.date}-${b.category}-${b.rowNumber}`)),
    payments: payments.sort((a, b) => `${a.date}-${a.shift}-${a.channel}`.localeCompare(`${b.date}-${b.shift}-${b.channel}`))
  };
}

function hasRowIssue(row: StockRow) {
  return (row.hasNightStock && Math.abs(row.sameDayStockDiff) > 0.01) || Math.abs(row.morningAmountDiff) > 0.01 || Math.abs(row.nightAmountDiff) > 0.01 || Math.abs(row.totalAmountDiff) > 0.01;
}

function columnName(row: StockRow, colIndex: number) {
  const col = `${XLSX.utils.encode_col(colIndex)}列`;
  if (colIndex === row.morningStockCol) return `${col} 早班库存`;
  if (colIndex === row.nightStockCol) return `${col} 夜班库存`;
  return col;
}

function makeWriteBack(row: StockRow, colIndex: number, before: number, after: number): WriteBackTarget {
  const cell = XLSX.utils.encode_cell({ r: row.rowIndex, c: colIndex });
  return { key: rowCellKey(row, colIndex), date: row.date, sheetName: row.sheetName, rowIndex: row.rowIndex, rowNumber: row.rowNumber, colIndex, cell, columnName: columnName(row, colIndex), category: row.category, name: row.name, price: row.price, before, after: round2(after) };
}

function writeCell(sheet: XLSX.WorkSheet, rowIndex: number, colIndex: number, value: number) {
  const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const cell = (sheet[address] ?? {}) as CellWithStyle;
  delete cell.f;
  delete cell.w;
  cell.t = "n";
  cell.v = value;
  sheet[address] = cell;
}

function setCellStyle(sheet: XLSX.WorkSheet, row: number, col: number, style: Record<string, unknown>) {
  const address = XLSX.utils.encode_cell({ r: row, c: col });
  const cell = (sheet[address] ?? { t: "s", v: "" }) as CellWithStyle;
  cell.s = { ...(cell.s ?? {}), ...style };
  sheet[address] = cell;
}

function setNumFormat(sheet: XLSX.WorkSheet, row: number, col: number, format: string) {
  const address = XLSX.utils.encode_cell({ r: row, c: col });
  const cell = sheet[address] as CellWithStyle | undefined;
  if (cell) cell.z = format;
}

function styleRow(sheet: XLSX.WorkSheet, row: number, colCount: number, style: Record<string, unknown>) {
  for (let col = 0; col < colCount; col += 1) setCellStyle(sheet, row, col, style);
}

function writeBackLocation(issue: HandoverIssue) {
  return issue.writeBack ? `${issue.writeBack.sheetName}!${issue.writeBack.cell}（${issue.writeBack.columnName}）` : "不自动写回";
}

function makeEmptyTotals(channels: string[]) {
  return Object.fromEntries(channels.map((channel) => [channel, 0])) as Record<string, number>;
}

function downloadWorkbook(fileName: string, sheets: Record<string, Record<string, unknown>[]>) {
  const workbook = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name, rows]) => XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), name.slice(0, 31)));
  XLSX.writeFile(workbook, fileName);
}

function Card({ title, value, desc, tone = "default" }: { title: string; value: string; desc?: string; tone?: "default" | "danger" }) {
  const cls = tone === "danger" ? "border-red-500 bg-red-600 text-white shadow-sm" : "border bg-white text-slate-900 shadow-sm";
  const titleCls = tone === "danger" ? "text-red-100" : "text-slate-500";
  const descCls = tone === "danger" ? "text-red-100" : "text-slate-500";
  return <div className={`rounded-2xl p-5 ${cls}`}><div className={`text-sm ${titleCls}`}>{title}</div><div className="mt-2 text-2xl font-bold">{value}</div>{desc ? <div className={`mt-1 text-xs ${descCls}`}>{desc}</div> : null}</div>;
}

function Badge({ children, tone = "gray" }: { children: ReactNode; tone?: "gray" | "green" | "red" | "yellow" }) {
  const cls = { gray: "bg-slate-100 text-slate-700", green: "bg-green-100 text-green-700", red: "bg-red-100 text-red-700", yellow: "bg-yellow-100 text-yellow-800" }[tone];
  return <span className={`rounded-full px-2 py-1 text-xs font-medium ${cls}`}>{children}</span>;
}

function StatsCards({ stats }: { stats: IssueStats }) {
  return <div className="grid gap-4 md:grid-cols-5"><Card tone="danger" title="总异常数" value={`${stats.totalIssueCount} 条`} desc="库存 + 金额" /><Card title="库存异常数" value={`${stats.inventoryIssueCount} 条`} /><Card title="金额异常数" value={`${stats.amountIssueCount} 条`} /><Card title="可写回原表数量" value={`${stats.writableCount} 条`} /><Card title="已准备写回数量" value={`${stats.preparedCount} 条`} /></div>;
}

function groupProductsByCategory(productSummary: ProductSummary[]) {
  const map = new Map<string, ProductSummary[]>();
  productSummary.forEach((item) => {
    const list = map.get(item.category) ?? [];
    list.push(item);
    map.set(item.category, list);
  });
  return Array.from(map.entries()).sort(([a], [b]) => a.localeCompare(b));
}

function addWorksheetStyle(sheet: XLSX.WorkSheet, rowCount: number, colCount: number) {
  styleRow(sheet, 0, colCount, TITLE_STYLE);
  styleRow(sheet, 1, colCount, META_STYLE);
  styleRow(sheet, 3, colCount, HEADER_STYLE);
  for (let row = 0; row < rowCount; row += 1) {
    for (let col = 0; col < colCount; col += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = sheet[address] as CellWithStyle | undefined;
      if (cell && (col === 3 || col === 5)) cell.z = "0.00";
    }
  }
}

export default function InventoryPage() {
  const [tab, setTab] = useState<Tab>("overview");
  const [fileName, setFileName] = useState("");
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [rows, setRows] = useState<StockRow[]>([]);
  const [paymentRecords, setPaymentRecords] = useState<PaymentRecord[]>([]);
  const [countInput, setCountInput] = useState<Record<string, string>>({});
  const [corrections, setCorrections] = useState<Record<string, CorrectionRecord>>({});

  async function handleFile(file: File) {
    const buffer = await file.arrayBuffer();
    const nextWorkbook = XLSX.read(buffer, { type: "array", cellDates: true, cellFormula: true, cellNF: true, cellStyles: true });
    const parsed = parseWorkbook(nextWorkbook);
    setWorkbook(nextWorkbook);
    setRows(parsed.rows);
    setPaymentRecords(parsed.payments);
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
      const key = `${row.category}__${row.name}__${row.price}`;
      const item = map.get(key) ?? { category: row.category, name: row.name, price: row.price, purchase: 0, morningSold: 0, nightSold: 0, totalSold: 0, totalAmount: 0, endingStock: 0 };
      item.purchase += row.purchase;
      item.morningSold += row.morningSold;
      item.nightSold += row.nightSold;
      item.totalSold += row.morningSold + row.nightSold;
      item.totalAmount += row.totalAmount;
      item.endingStock = row.endingStock;
      map.set(key, item);
    });
    return Array.from(map.values()).map((item) => ({ ...item, purchase: round2(item.purchase), morningSold: round2(item.morningSold), nightSold: round2(item.nightSold), totalSold: round2(item.totalSold), totalAmount: round2(item.totalAmount), endingStock: round2(item.endingStock) })).sort((a, b) => a.category.localeCompare(b.category) || b.totalAmount - a.totalAmount);
  }, [rows]);

  const categorySummary = useMemo<CategorySummary[]>(() => {
    const map = new Map<string, CategorySummary>();
    productSummary.forEach((item) => {
      const current = map.get(item.category) ?? { category: item.category, totalSold: 0, totalAmount: 0, count: 0 };
      current.totalSold += item.totalSold;
      current.totalAmount += item.totalAmount;
      current.count += 1;
      map.set(item.category, current);
    });
    return Array.from(map.values()).map((item) => ({ ...item, totalSold: round2(item.totalSold), totalAmount: round2(item.totalAmount) })).sort((a, b) => b.totalAmount - a.totalAmount);
  }, [productSummary]);

  const paymentChannels = useMemo(() => {
    const channels: string[] = [];
    paymentRecords.forEach((record) => {
      if (record.channel && !channels.includes(record.channel)) channels.push(record.channel);
    });
    return channels;
  }, [paymentRecords]);

  const paymentDaily = useMemo<PaymentDailySummary[]>(() => {
    const byDate = new Map<string, PaymentRecord[]>();
    paymentRecords.forEach((record) => {
      const list = byDate.get(record.date) ?? [];
      list.push(record);
      byDate.set(record.date, list);
    });
    return Array.from(byDate.entries()).map(([date, records]) => {
      const totalRecords = records.filter((record) => record.shift === "合计");
      const source = totalRecords.length ? totalRecords : records.filter((record) => record.shift !== "合计");
      const totals = makeEmptyTotals(paymentChannels);
      source.forEach((record) => { totals[record.channel] = round2((totals[record.channel] ?? 0) + record.amount); });
      const total = round2(paymentChannels.reduce((sum, channel) => sum + (totals[channel] ?? 0), 0));
      return { date, totals, total };
    }).sort((a, b) => a.date.localeCompare(b.date));
  }, [paymentRecords, paymentChannels]);

  const paymentSummary = useMemo<PaymentSummary[]>(() => {
    const totals = makeEmptyTotals(paymentChannels);
    paymentDaily.forEach((day) => paymentChannels.forEach((channel) => { totals[channel] = round2((totals[channel] ?? 0) + (day.totals[channel] ?? 0)); }));
    return paymentChannels.map((channel) => ({ channel, amount: round2(totals[channel] ?? 0) }));
  }, [paymentDaily, paymentChannels]);

  const handovers = useMemo<HandoverIssue[]>(() => {
    const issues: HandoverIssue[] = [];
    rows.forEach((row) => {
      if (row.hasNightStock && Math.abs(row.sameDayStockDiff) > 0.01) {
        const type: IssueType = isMissedMorningDeduction(row) ? "疑似未扣早班销量" : "早班交夜班";
        const note = type === "疑似未扣早班销量" ? "夜班库存比正确值多出的数量，正好等于早班销量。高度疑似交班时忘记扣早班售卖数量。" : "夜班库存应等于：早班库存 + 进货 - 早班售卖数量。";
        const writeBack = makeWriteBack(row, row.nightStockCol, row.nightStock, correctedNightStock(row));
        issues.push({ id: `${type}__${writeBack.key}`, type, date: row.date, category: row.category, name: row.name, expected: row.sameDayExpectedNightStock, actual: row.nightStock, diff: row.sameDayStockDiff, note, writeBack });
      }
      if (Math.abs(row.morningAmountDiff) > 0.01) issues.push({ id: `金额核对__早班__${rowKey(row)}`, type: "金额核对", date: row.date, category: row.category, name: row.name, expected: round2(row.morningSold * row.price), actual: row.morningAmount, diff: row.morningAmountDiff, note: "早班金额不等于：早班售卖数量 × 价格。金额核对异常只提示，不自动改原表。" });
      if (Math.abs(row.nightAmountDiff) > 0.01) issues.push({ id: `金额核对__夜班__${rowKey(row)}`, type: "金额核对", date: row.date, category: row.category, name: row.name, expected: round2(row.nightSold * row.price), actual: row.nightAmount, diff: row.nightAmountDiff, note: "夜班金额不等于：夜班售卖数量 × 价格。金额核对异常只提示，不自动改原表。" });
      if (Math.abs(row.totalAmountDiff) > 0.01) issues.push({ id: `金额核对__合计__${rowKey(row)}`, type: "金额核对", date: row.date, category: row.category, name: row.name, expected: round2(row.morningAmount + row.nightAmount), actual: row.totalAmount, diff: row.totalAmountDiff, note: "合计金额不等于：早班金额 + 夜班金额。金额核对异常只提示，不自动改原表。" });
    });

    const byProduct = new Map<string, StockRow[]>();
    rows.forEach((row) => {
      const key = `${row.category}__${row.name}__${row.price}`;
      const list = byProduct.get(key) ?? [];
      list.push(row);
      byProduct.set(key, list);
    });
    byProduct.forEach((list) => {
      const sorted = [...list].sort((a, b) => a.date.localeCompare(b.date));
      for (let i = 0; i < sorted.length - 1; i += 1) {
        const current = sorted[i];
        const next = sorted[i + 1];
        if (!current.hasNightStock || !next.hasMorningStock) continue;
        const expected = round2(current.nightStock - current.nightSold);
        const diff = round2(next.morningStock - expected);
        if (Math.abs(diff) > 0.01) {
          const writeBack = makeWriteBack(next, next.morningStockCol, next.morningStock, expected);
          issues.push({ id: `次日接班__${writeBack.key}__${current.date}`, type: "次日接班", date: current.date, nextDate: next.date, category: current.category, name: current.name, expected, actual: next.morningStock, diff, note: "次日早班库存应等于：前一日夜班库存 - 前一日夜班售卖数量。", writeBack });
        }
      }
    });
    return issues.sort((a, b) => `${a.date}-${a.category}-${a.name}-${a.type}`.localeCompare(`${b.date}-${b.category}-${b.name}-${b.type}`));
  }, [rows]);

  const latestDate = daily.length ? daily[daily.length - 1].date : "";
  const latestRows = rows.filter((row) => row.date === latestDate);
  const latestSmokeRows = latestRows.filter((row) => row.category === CIGARETTE_CATEGORY);
  const latestMissedMorningRows = latestSmokeRows.filter(isMissedMorningDeduction);
  const totalAmount = daily.reduce((sum, item) => sum + item.totalAmount, 0);
  const totalSold = daily.reduce((sum, item) => sum + item.totalSold, 0);
  const totalPaymentAmount = paymentSummary.reduce((sum, item) => sum + item.amount, 0);
  const correctionRows = Object.values(corrections).sort((a, b) => `${a.date}-${a.category}-${a.name}-${a.cell}`.localeCompare(`${b.date}-${b.category}-${b.name}-${b.cell}`));
  const stats = useMemo<IssueStats>(() => {
    const inventoryIssueCount = handovers.filter((item) => item.type !== "金额核对").length;
    const amountIssueCount = handovers.filter((item) => item.type === "金额核对").length;
    const writableCount = handovers.filter((item) => item.writeBack).length;
    return { totalIssueCount: handovers.length, inventoryIssueCount, amountIssueCount, writableCount, preparedCount: correctionRows.length };
  }, [handovers, correctionRows.length]);

  function makeCorrection(target: WriteBackTarget, reason: string): CorrectionRecord {
    return { ...target, after: round2(target.after), diff: round2(target.after - target.before), reason, correctedAt: new Date().toLocaleString("zh-CN", { hour12: false }) };
  }

  function setCorrection(target: WriteBackTarget, reason: string, old: Record<string, CorrectionRecord>) {
    return { ...old, [target.key]: makeCorrection(target, reason) };
  }

  function correctedExpectedEndingStock(row: StockRow) {
    const correction = corrections[rowCellKey(row, row.nightStockCol)];
    return correction ? round2(correction.after - row.nightSold) : row.endingStock;
  }

  function addOnsiteCorrection(row: StockRow, actualEndingStock: number) {
    const fixedNightStock = round2(actualEndingStock + row.nightSold);
    const target = makeWriteBack(row, row.nightStockCol, row.nightStock, fixedNightStock);
    setCorrections((old) => setCorrection(target, "按现场实点反推夜班库存，写回夜班库存列", old));
    setCountInput((old) => ({ ...old, [countKey(row)]: String(actualEndingStock) }));
  }

  function applyIssueCorrection(issue: HandoverIssue) {
    if (!issue.writeBack) return alert("金额核对异常只提示，不自动改原表。");
    setCorrections((old) => setCorrection(issue.writeBack as WriteBackTarget, issue.note, old));
  }

  function undoCorrectionByKey(key: string) {
    setCorrections((old) => {
      const copy = { ...old };
      delete copy[key];
      return copy;
    });
  }

  function undoCorrection(row: StockRow) {
    undoCorrectionByKey(rowCellKey(row, row.nightStockCol));
  }

  function applyWritableHandoverCorrections(scope: "all" | "missedMorning") {
    const writable = handovers.filter((issue) => issue.writeBack && (scope === "all" || issue.type === "疑似未扣早班销量") && !corrections[issue.writeBack.key]);
    if (!writable.length) return alert("没有新的库存异常需要准备写回。");
    if (!confirm(`确定准备写回 ${writable.length} 条库存异常吗？金额核对异常不会自动改原表。`)) return;
    setCorrections((old) => writable.reduce((acc, issue) => setCorrection(issue.writeBack as WriteBackTarget, issue.note, acc), old));
  }

  function downloadCorrectedOriginal() {
    if (!workbook) return alert("请先上传原始报表 Excel");
    const correctionList = Object.values(corrections);
    if (!correctionList.length) return alert("还没有纠正内容。请先在交接核对或现场盘点里准备写回。");
    const fixed = XLSX.read(XLSX.write(workbook, { bookType: "xlsx", type: "array" }), { type: "array", cellDates: true, cellFormula: true, cellNF: true, cellStyles: true });
    correctionList.forEach((item) => {
      const sheet = fixed.Sheets[item.sheetName];
      if (sheet) writeCell(sheet, item.rowIndex, item.colIndex, item.after);
    });
    XLSX.writeFile(fixed, `已纠正-${safeFileName(fileName || "手工报表")}.xlsx`);
  }

  function exportSalesSummaryReport() {
    if (!rows.length) return alert("请先上传原始报表 Excel");
    const totalQty = round2(productSummary.reduce((sum, item) => sum + item.totalSold, 0));
    const totalMoney = round2(productSummary.reduce((sum, item) => sum + item.totalAmount, 0));
    const grouped = groupProductsByCategory(productSummary);
    const aoa: Array<Array<string | number>> = [
      ["SY11 小商品销售汇总报表", "", "", "", "", ""],
      [`门店：${STORE_NAME}`, "", `营业日范围：${reportDateRange(daily)}`, "", `最后查询时间：${queryTime()}`, ""],
      ["", "", "", "", "", ""],
      ["序号", "商品分类", "商品名称", "平均价格", "数量", "金额"]
    ];
    const categoryHeaderRows: number[] = [];
    const subtotalRows: number[] = [];
    let index = 1;

    grouped.forEach(([category, products]) => {
      categoryHeaderRows.push(aoa.length);
      aoa.push([`【${category}】`, "", "", "", "", ""]);
      products.forEach((item) => {
        aoa.push([index, item.category, item.name, round2(item.price), round2(item.totalSold), round2(item.totalAmount)]);
        index += 1;
      });
      const categoryQty = round2(products.reduce((sum, item) => sum + item.totalSold, 0));
      const categoryMoney = round2(products.reduce((sum, item) => sum + item.totalAmount, 0));
      subtotalRows.push(aoa.length);
      aoa.push([`${category}小计`, "", "", "", categoryQty, categoryMoney]);
      aoa.push(["", "", "", "", "", ""]);
    });

    const totalRow = aoa.length;
    aoa.push(["全部总计", "", "", "", totalQty, totalMoney]);

    const sheet = XLSX.utils.aoa_to_sheet(aoa);
    sheet["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 1 } },
      { s: { r: 1, c: 2 }, e: { r: 1, c: 3 } },
      { s: { r: 1, c: 4 }, e: { r: 1, c: 5 } },
      ...categoryHeaderRows.map((row) => ({ s: { r: row, c: 0 }, e: { r: row, c: 5 } }))
    ];
    sheet["!cols"] = [{ wch: 10 }, { wch: 20 }, { wch: 34 }, { wch: 12 }, { wch: 12 }, { wch: 14 }];
    sheet["!rows"] = aoa.map((_, row) => ({ hpt: row === 0 ? 26 : row === 3 ? 22 : 19 }));
    addWorksheetStyle(sheet, aoa.length, 6);
    categoryHeaderRows.forEach((row) => styleRow(sheet, row, 6, CATEGORY_STYLE));
    subtotalRows.forEach((row) => styleRow(sheet, row, 6, SUBTOTAL_STYLE));
    styleRow(sheet, totalRow, 6, TOTAL_STYLE);
    for (let row = 4; row <= totalRow; row += 1) {
      setNumFormat(sheet, row, 3, "0.00");
      setNumFormat(sheet, row, 5, "0.00");
    }

    const categoryAoa: Array<Array<string | number>> = [
      ["SY11 分类销售汇总", "", "", ""],
      [`门店：${STORE_NAME}`, `营业日范围：${reportDateRange(daily)}`, `最后查询时间：${queryTime()}`, ""],
      ["", "", "", ""],
      ["序号", "商品分类", "品规数", "销售金额"],
      ...categorySummary.map((item, itemIndex) => [itemIndex + 1, item.category, item.count, round2(item.totalAmount)]),
      ["全部总计", "", categorySummary.reduce((sum, item) => sum + item.count, 0), round2(totalMoney)]
    ];
    const categorySheet = XLSX.utils.aoa_to_sheet(categoryAoa);
    categorySheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
    categorySheet["!cols"] = [{ wch: 10 }, { wch: 22 }, { wch: 12 }, { wch: 14 }];
    styleRow(categorySheet, 0, 4, TITLE_STYLE);
    styleRow(categorySheet, 1, 4, META_STYLE);
    styleRow(categorySheet, 3, 4, HEADER_STYLE);
    styleRow(categorySheet, categoryAoa.length - 1, 4, TOTAL_STYLE);
    for (let row = 4; row < categoryAoa.length; row += 1) setNumFormat(categorySheet, row, 3, "0.00");

    const paymentRows = [
      ...paymentSummary.map((item, itemIndex) => ({ 序号: itemIndex + 1, 收款渠道: item.channel, 月收入: item.amount })),
      { 序号: "总计", 收款渠道: "", 月收入: round2(totalPaymentAmount) }
    ];
    const paymentSheet = XLSX.utils.json_to_sheet(paymentRows);
    paymentSheet["!cols"] = [{ wch: 10 }, { wch: 18 }, { wch: 14 }];

    const reportWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(reportWorkbook, sheet, "销售汇总报表");
    XLSX.utils.book_append_sheet(reportWorkbook, categorySheet, "分类汇总");
    XLSX.utils.book_append_sheet(reportWorkbook, paymentSheet, "收款渠道月统计");
    XLSX.writeFile(reportWorkbook, `${STORE_NAME}-小商品销售汇总报表.xlsx`);
  }

  function exportExceptionReport() {
    if (!rows.length) return alert("请先上传原始报表 Excel");
    const aoa: Array<Array<string | number>> = [
      ["报表异常核对报表", "", "", "", "", "", "", "", "", "", ""],
      [`门店：${STORE_NAME}`, `日期范围：${reportDateRange(daily)}`, `总异常数：${stats.totalIssueCount}`, `库存异常数：${stats.inventoryIssueCount}`, `金额异常数：${stats.amountIssueCount}`, "", "", "", "", "", ""],
      ["", "", "", "", "", "", "", "", "", "", ""],
      ["序号", "异常类型", "商品分类", "日期", "次日", "商品名称", "应为", "实际", "差异", "写回位置", "说明"],
      ...handovers.map((item, itemIndex) => [itemIndex + 1, item.type, item.category, item.date, item.nextDate ?? "", item.name, item.expected, item.actual, item.diff, writeBackLocation(item), item.note]),
      ["总计", `共 ${stats.totalIssueCount} 条`, "", "", "", `库存异常 ${stats.inventoryIssueCount} 条`, `金额异常 ${stats.amountIssueCount} 条`, "", "", `可写回 ${stats.writableCount} 条`, ""]
    ];
    const sheet = XLSX.utils.aoa_to_sheet(aoa);
    const totalRow = aoa.length - 1;
    sheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 10 } }];
    sheet["!cols"] = [{ wch: 8 }, { wch: 18 }, { wch: 20 }, { wch: 12 }, { wch: 12 }, { wch: 28 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 24 }, { wch: 52 }];
    setCellStyle(sheet, 0, 0, TITLE_STYLE);
    styleRow(sheet, 3, 11, HEADER_STYLE);
    styleRow(sheet, totalRow, 11, TOTAL_STYLE);
    const reportWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(reportWorkbook, sheet, "异常核对报表");
    XLSX.writeFile(reportWorkbook, "报表异常核对报表.xlsx");
  }

  function exportOnsiteCount() {
    if (!latestSmokeRows.length) return alert("最新日期没有香烟明细，现场盘点只显示香烟。");
    const countRows = latestSmokeRows.map((row) => {
      const actualText = countInput[countKey(row)] ?? "";
      const actual = actualText === "" ? "" : n(actualText);
      const expected = correctedExpectedEndingStock(row);
      const diff = actual === "" ? "" : round2(Number(actual) - expected);
      const correction = corrections[rowCellKey(row, row.nightStockCol)];
      return { 日期: latestDate, 商品分类: row.category, 商品名称: row.name, 价格: row.price, 原应剩数量: row.endingStock, 当前应剩数量: expected, 现场实点: actual, 差异: diff, 是否已准备写回: correction ? "是" : "否", 写回位置: correction ? `${correction.sheetName}!${correction.cell}` : "" };
    });
    downloadWorkbook(`${latestDate}-香烟现场盘点表.xlsx`, { 香烟现场盘点: countRows, 差异香烟: countRows.filter((row) => row.差异 !== "" && row.差异 !== 0) });
  }

  return <main className="min-h-screen p-4 md:p-8"><div className="mx-auto max-w-7xl">
    <header className="mb-6 rounded-3xl bg-slate-900 p-6 text-white shadow-sm"><p className="text-sm text-slate-300">Inventory Report Reconciliation</p><h1 className="mt-1 text-2xl font-bold md:text-4xl">报表自动核对助手</h1><p className="mt-2 max-w-3xl text-sm text-slate-300">支持多块商品表解析，并自动识别 Excel 里实际出现的收款渠道。</p></header>
    <section className="mb-6 rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">导入手工报表 Excel</h2><p className="mt-1 text-sm text-slate-500">系统会扫描商品表，也会扫描每个“收款渠道”后面的渠道名称，不再固定渠道列表。</p></div><label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-slate-900 px-4 py-3 text-sm font-medium text-white">选择报表文件<input className="hidden" type="file" accept=".xlsx,.xls" onChange={(event: ChangeEvent<HTMLInputElement>) => event.target.files?.[0] && handleFile(event.target.files[0])} /></label></div>{fileName ? <p className="mt-3 text-sm text-slate-600">已导入：<b>{fileName}</b>，共解析 {rows.length} 条商品明细，{categorySummary.length} 个分类，{paymentChannels.length} 个收款渠道。</p> : null}</section>
    <nav className="mb-6 flex flex-wrap gap-2">{nav.map(([key, label]) => <button key={key} onClick={() => setTab(key)} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === key ? "bg-slate-900 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>{label}</button>)}</nav>
    {!rows.length ? <section className="rounded-2xl border bg-white p-6 shadow-sm"><h2 className="text-lg font-bold">等待导入报表</h2><p className="mt-2 text-sm text-slate-600">请先上传手工报表 Excel。</p></section> : null}
    {rows.length > 0 && tab === "overview" ? <section className="space-y-6"><StatsCards stats={stats} /><div className="grid gap-4 md:grid-cols-4"><Card title="日期数量" value={`${daily.length} 天`} desc={`最新日期 ${latestDate}`} /><Card title="商品品规" value={`${productSummary.length} 个`} desc={`${categorySummary.length} 个分类`} /><Card title="总销量" value={`${round2(totalSold)} 件`} /><Card title="总销售金额" value={`¥${money(totalAmount)}`} desc={`收款统计 ¥${money(totalPaymentAmount)}`} /></div>{categorySummary.length ? <div className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">分类汇总</h2><div className="mt-3 grid gap-3 md:grid-cols-4">{categorySummary.map((item) => <div key={item.category} className="rounded-xl bg-slate-50 p-4"><div className="font-bold">{item.category}</div><div className="mt-1 text-sm text-slate-600">品规 {item.count} 个 / 销量 {item.totalSold} / 金额 ¥{money(item.totalAmount)}</div></div>)}</div></div> : null}{paymentChannels.length ? <div className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">收款渠道月统计</h2><div className="mt-3 grid gap-3 md:grid-cols-5">{paymentSummary.map((item) => <div key={item.channel} className="rounded-xl bg-slate-50 p-4"><div className="font-bold">{item.channel}</div><div className="mt-1 text-sm text-slate-600">¥{money(item.amount)}</div></div>)}</div></div> : null}<div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><h2 className="text-lg font-bold">导出</h2><div className="flex flex-wrap gap-2"><button onClick={downloadCorrectedOriginal} className="rounded-xl bg-green-700 px-4 py-2 text-sm font-medium text-white">下载已纠正原表</button><button onClick={exportSalesSummaryReport} className="rounded-xl border px-4 py-2 text-sm font-medium">导出销售汇总报表</button><button onClick={exportExceptionReport} className="rounded-xl border px-4 py-2 text-sm font-medium">导出异常核对报表</button></div></div></div></section> : null}
    {rows.length > 0 && tab === "daily" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">每日报表汇总</h2><p className="mt-1 text-sm text-slate-500">统计所有解析到的商品表。</p><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["日期", "进货", "早班销量", "夜班销量", "总销量", "早班金额", "夜班金额", "总金额", "异常"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{daily.map((item) => <tr key={item.date} className="border-t"><td className="p-3">{item.date}</td><td className="p-3">{item.purchase}</td><td className="p-3">{item.morningSold}</td><td className="p-3">{item.nightSold}</td><td className="p-3">{item.totalSold}</td><td className="p-3">¥{money(item.morningAmount)}</td><td className="p-3">¥{money(item.nightAmount)}</td><td className="p-3">¥{money(item.totalAmount)}</td><td className="p-3">{item.abnormalCount ? <Badge tone="red">{item.abnormalCount} 条</Badge> : <Badge tone="green">正常</Badge>}</td></tr>)}</tbody></table></div></section> : null}
    {rows.length > 0 && tab === "summary" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">商品总汇</h2><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["分类", "商品", "价格", "月进货", "早班销量", "夜班销量", "月销量", "月销售金额", "月末结存"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{productSummary.map((item) => <tr key={`${item.category}-${item.name}-${item.price}`} className="border-t"><td className="p-3">{item.category}</td><td className="p-3">{item.name}</td><td className="p-3">¥{money(item.price)}</td><td className="p-3">{item.purchase}</td><td className="p-3">{item.morningSold}</td><td className="p-3">{item.nightSold}</td><td className="p-3">{item.totalSold}</td><td className="p-3">¥{money(item.totalAmount)}</td><td className="p-3">{item.endingStock}</td></tr>)}</tbody></table></div></section> : null}
    {rows.length > 0 && tab === "payment" ? <section className="space-y-5"><div className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">收款渠道月统计</h2>{!paymentChannels.length ? <p className="mt-2 text-sm text-slate-500">没有识别到收款渠道。请确认表格里有“收款渠道”文字。</p> : <div className="mt-3 grid gap-3 md:grid-cols-5">{paymentSummary.map((item) => <div key={item.channel} className="rounded-xl bg-slate-50 p-4"><div className="font-bold">{item.channel}</div><div className="mt-1 text-sm text-slate-600">¥{money(item.amount)}</div></div>)}</div>}<div className="mt-4 rounded-xl bg-slate-900 p-4 text-white"><div className="text-sm text-slate-300">全部渠道合计</div><div className="mt-1 text-2xl font-bold">¥{money(totalPaymentAmount)}</div></div></div>{paymentChannels.length ? <div className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">每日收款明细</h2><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["日期", ...paymentChannels, "合计"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{paymentDaily.map((item) => <tr key={item.date} className="border-t"><td className="p-3">{item.date}</td>{paymentChannels.map((channel) => <td key={channel} className="p-3">¥{money(item.totals[channel] ?? 0)}</td>)}<td className="p-3 font-bold">¥{money(item.total)}</td></tr>)}<tr className="border-t bg-slate-900 text-white"><td className="p-3 font-bold">渠道合计</td>{paymentSummary.map((item) => <td key={item.channel} className="p-3 font-bold">¥{money(item.amount)}</td>)}<td className="p-3 font-bold">¥{money(totalPaymentAmount)}</td></tr></tbody></table></div></div> : null}</section> : null}
    {rows.length > 0 && tab === "handover" ? <section className="space-y-5"><StatsCards stats={stats} /><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">交接和金额异常</h2><p className="mt-1 text-sm text-slate-500">库存异常可准备写回，金额核对只提示。</p></div><div className="flex flex-wrap gap-2"><button onClick={() => applyWritableHandoverCorrections("all")} className="rounded-xl bg-yellow-700 px-4 py-2 text-sm text-white">准备写回全部库存异常</button><button onClick={() => applyWritableHandoverCorrections("missedMorning")} className="rounded-xl border px-4 py-2 text-sm">只准备疑似未扣</button><button onClick={downloadCorrectedOriginal} className="rounded-xl bg-green-700 px-4 py-2 text-sm text-white">下载已纠正原表</button></div></div><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["类型", "分类", "日期", "次日", "商品", "应为", "实际", "差异", "写回位置", "说明", "操作"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{handovers.map((item) => { const prepared = item.writeBack ? corrections[item.writeBack.key] : null; return <tr key={item.id} className="border-t"><td className="p-3"><Badge tone={item.type === "金额核对" ? "red" : item.type === "疑似未扣早班销量" ? "yellow" : "red"}>{item.type}</Badge></td><td className="p-3">{item.category}</td><td className="p-3">{item.date}</td><td className="p-3">{item.nextDate ?? ""}</td><td className="p-3">{item.name}</td><td className="p-3">{item.expected}</td><td className="p-3">{item.actual}</td><td className="p-3">{item.diff}</td><td className="p-3">{writeBackLocation(item)}</td><td className="p-3">{item.note}</td><td className="p-3">{item.writeBack ? prepared ? <button className="rounded-lg border px-3 py-1 text-xs" onClick={() => undoCorrectionByKey(item.writeBack!.key)}>撤销</button> : <button className="rounded-lg bg-slate-900 px-3 py-1 text-xs text-white" onClick={() => applyIssueCorrection(item)}>准备写回</button> : <Badge>只提示</Badge>}</td></tr>; })}</tbody></table></div>{correctionRows.length ? <div className="mt-5 rounded-xl bg-yellow-50 p-4 text-sm text-yellow-900"><b>已准备写回原表 {correctionRows.length} 条：</b>{correctionRows.slice(0, 20).map((item) => <div key={item.key}>{item.sheetName}!{item.cell} {item.category} / {item.name}（{item.columnName}）：{item.before} → {item.after}</div>)}</div> : null}</div></section> : null}
    {rows.length > 0 && tab === "count" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">香烟现场盘点清单</h2><p className="mt-1 text-sm text-slate-500">现场盘点只显示香烟；其他分类不进入现场盘点。</p></div><button onClick={exportOnsiteCount} className="rounded-xl border px-4 py-2 text-sm font-medium">导出香烟现场盘点明细</button></div>{!latestSmokeRows.length ? <div className="mt-4 rounded-xl bg-slate-50 p-4 text-sm text-slate-600">最新日期没有识别到香烟明细。</div> : null}{latestMissedMorningRows.length ? <div className="mt-4 rounded-xl bg-yellow-50 p-4 text-sm text-yellow-900"><b>最新日期检测到 {latestMissedMorningRows.length} 条香烟疑似交班忘扣早班销量。</b><button className="ml-0 mt-3 rounded-lg bg-yellow-700 px-4 py-2 text-sm text-white md:ml-3 md:mt-0" onClick={() => applyWritableHandoverCorrections("missedMorning")}>准备写回疑似未扣</button></div> : null}<div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["分类", "商品", "价格", "当前应剩", "现场实点", "差异", "状态", "纠正原表"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{latestSmokeRows.map((row) => { const key = rowKey(row); const inputKey = countKey(row); const actualText = countInput[inputKey] ?? ""; const actual = actualText === "" ? null : n(actualText); const expected = correctedExpectedEndingStock(row); const diff = actual === null ? null : round2(actual - expected); const correction = corrections[rowCellKey(row, row.nightStockCol)]; return <tr key={key} className="border-t"><td className="p-3">{row.category}</td><td className="p-3">{row.name}{isMissedMorningDeduction(row) ? <div className="text-xs text-yellow-700">疑似未扣早班销量</div> : null}</td><td className="p-3">¥{money(row.price)}</td><td className="p-3 font-bold">{expected}{correction ? <div className="text-xs font-normal text-slate-500">写回夜班库存：{correction.before} → {correction.after}</div> : null}</td><td className="p-3"><input className="w-28 rounded-xl border p-2" inputMode="decimal" value={actualText} onChange={(event: ChangeEvent<HTMLInputElement>) => setCountInput({ ...countInput, [inputKey]: event.target.value })} /></td><td className="p-3">{diff === null ? "" : diff}</td><td className="p-3">{correction ? <Badge tone="yellow">待写回原表</Badge> : diff === null ? <Badge>未点</Badge> : diff === 0 ? <Badge tone="green">正常</Badge> : <Badge tone="red">有差异</Badge>}</td><td className="p-3">{actual !== null && diff !== 0 ? <button className="rounded-lg bg-slate-900 px-3 py-1 text-xs text-white" onClick={() => addOnsiteCorrection(row, actual)}>按实点纠正夜班库存</button> : null}{correction ? <button className="ml-2 rounded-lg border px-3 py-1 text-xs" onClick={() => undoCorrection(row)}>撤销</button> : null}</td></tr>; })}</tbody></table></div>{correctionRows.length ? <div className="mt-5 rounded-xl bg-yellow-50 p-4 text-sm text-yellow-900"><b>待写回原表 {correctionRows.length} 条。</b><button className="ml-3 rounded-lg bg-green-700 px-4 py-2 text-sm text-white" onClick={downloadCorrectedOriginal}>下载已纠正原表</button></div> : null}</section> : null}
  </div></main>;
}
