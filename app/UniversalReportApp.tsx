"use client";

import { useEffect, useMemo, useState, type ChangeEvent, type ReactNode } from "react";
import * as XLSX from "xlsx";
import ProductReportApp from "./InventoryPageConfigurable";

type MainTab = "smart" | "product" | "templates" | "about";
type Severity = "error" | "warning" | "info";
type TableType = "商品库存/销售报表" | "工资表" | "进货表" | "费用表" | "收款表" | "库存表" | "通用表格";
type FieldRole =
  | "date"
  | "name"
  | "product"
  | "employee"
  | "customer"
  | "supplier"
  | "orderNo"
  | "sku"
  | "category"
  | "quantity"
  | "price"
  | "amount"
  | "totalAmount"
  | "paidAmount"
  | "receivable"
  | "payable"
  | "stockOpening"
  | "stockIn"
  | "stockOut"
  | "stockEnding"
  | "purchase"
  | "sold"
  | "basicSalary"
  | "attendance"
  | "overtime"
  | "commission"
  | "bonus"
  | "deduction"
  | "advance"
  | "grossSalary"
  | "netSalary"
  | "feeType"
  | "paymentChannel"
  | "remark";

type DataCell = string | number;
type DataRow = { rowNumber: number; values: Record<string, DataCell> };
type ColumnAnalysis = { index: number; header: string; role?: FieldRole; confidence: number; numericRatio: number; nonEmptyCount: number; examples: string[] };
type SheetAnalysis = { sheetName: string; type: TableType; confidence: number; headerRow: number; columns: ColumnAnalysis[]; rows: DataRow[]; rules: string[]; anomalies: Anomaly[]; matchedTemplate?: string };
type WorkbookAnalysis = { fileName: string; createdAt: string; sheets: SheetAnalysis[] };
type Anomaly = { severity: Severity; type: string; sheetName: string; rowNumber?: number; field?: string; message: string; expected?: number | string; actual?: number | string };
type SavedTemplate = { id: string; name: string; tableType: TableType; createdAt: string; updatedAt: string; headerRow: number; columns: Array<{ header: string; role?: FieldRole; index: number }> };

const STORAGE_KEY = "universal_report_templates_v2";
const MAIN_TABS: Array<[MainTab, string]> = [["smart", "智能上传"], ["product", "商品报表专用"], ["templates", "模板库"], ["about", "规则说明"]];

const ROLE_LABELS: Record<FieldRole, string> = {
  date: "日期",
  name: "名称",
  product: "商品",
  employee: "员工",
  customer: "客户",
  supplier: "供应商",
  orderNo: "单号",
  sku: "编码/SKU",
  category: "分类",
  quantity: "数量",
  price: "单价",
  amount: "金额",
  totalAmount: "合计金额",
  paidAmount: "实收金额",
  receivable: "应收",
  payable: "应付",
  stockOpening: "期初库存",
  stockIn: "入库",
  stockOut: "出库",
  stockEnding: "结存",
  purchase: "进货",
  sold: "销量",
  basicSalary: "基本工资",
  attendance: "出勤",
  overtime: "加班",
  commission: "提成",
  bonus: "奖金",
  deduction: "扣款",
  advance: "借支",
  grossSalary: "应发工资",
  netSalary: "实发工资",
  feeType: "费用类型",
  paymentChannel: "收款渠道",
  remark: "备注"
};

const ROLE_ALIASES: Record<FieldRole, string[]> = {
  date: ["日期", "时间", "营业日", "日", "date"],
  name: ["名称", "姓名", "名字", "name"],
  product: ["商品", "商品名称", "品名", "产品", "货品", "物品", "项目"],
  employee: ["员工", "员工姓名", "姓名", "人员", "工号", "岗位"],
  customer: ["客户", "客户名称", "会员", "买家"],
  supplier: ["供应商", "厂家", "供货商", "采购商"],
  orderNo: ["单号", "订单号", "流水号", "编号", "票号"],
  sku: ["sku", "货号", "编码", "条码", "商品编码"],
  category: ["分类", "类别", "品类", "类型", "部门"],
  quantity: ["数量", "件数", "个数", "瓶数", "包数", "条数", "qty"],
  price: ["价格", "单价", "售价", "进价", "成本价", "price"],
  amount: ["金额", "小计", "销售额", "收入", "支出", "费用", "amount"],
  totalAmount: ["合计", "总计", "总金额", "总价", "合计金额"],
  paidAmount: ["实收", "实收金额", "到账", "收款金额", "已收"],
  receivable: ["应收", "应收金额", "应收款"],
  payable: ["应付", "应付金额", "应付款"],
  stockOpening: ["期初", "初始库存", "早班库存", "上月结存", "昨日结存"],
  stockIn: ["入库", "进货", "采购", "补货", "收入库"],
  stockOut: ["出库", "销量", "售卖", "销售数量", "消耗"],
  stockEnding: ["结存", "库存", "剩余", "月末库存", "夜班库存", "库存数量"],
  purchase: ["进货", "采购", "补货", "入库"],
  sold: ["销量", "售卖", "销售数量", "卖出", "出库"],
  basicSalary: ["基本工资", "底薪", "月薪", "工资标准"],
  attendance: ["出勤", "出勤天数", "天数", "工时", "考勤"],
  overtime: ["加班", "加班费", "加班工资"],
  commission: ["提成", "业绩提成", "销售提成"],
  bonus: ["奖金", "补贴", "满勤", "津贴", "奖励"],
  deduction: ["扣款", "罚款", "扣除", "社保", "个税"],
  advance: ["借支", "预支", "借款"],
  grossSalary: ["应发", "应发工资", "工资合计"],
  netSalary: ["实发", "实发工资", "到手", "实际发放"],
  feeType: ["费用类型", "费用项目", "支出项目", "科目"],
  paymentChannel: ["渠道", "收款渠道", "支付方式", "付款方式", "平台"],
  remark: ["备注", "说明", "原因", "note"]
};

function norm(value: unknown) { return String(value ?? "").replace(/\s/g, "").toLowerCase(); }
function text(value: unknown) { return String(value ?? "").trim(); }
function isBlank(value: unknown) { return value === null || value === undefined || text(value) === ""; }
function round2(value: number) { return Math.round((value + Number.EPSILON) * 100) / 100; }
function toNumber(value: unknown) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const cleaned = text(value).replace(/,/g, "").replace(/[￥¥元%\s]/g, "");
  if (!cleaned) return 0;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}
function money(value: number) { return round2(value).toFixed(2); }
function colName(index: number) { return XLSX.utils.encode_col(index); }
function uniqueHeader(raw: unknown, index: number, used: Set<string>) {
  const base = text(raw) || `列${colName(index)}`;
  let name = base;
  let suffix = 2;
  while (used.has(name)) {
    name = `${base}_${suffix}`;
    suffix += 1;
  }
  used.add(name);
  return name;
}

function sheetToGrid(sheet: XLSX.WorkSheet) {
  return XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null });
}

function scoreHeaderRow(row: unknown[]) {
  const cells = row.map(text).filter(Boolean);
  if (cells.length < 2) return 0;
  let score = Math.min(cells.length, 10);
  const joined = norm(cells.join(" "));
  Object.values(ROLE_ALIASES).flat().forEach((alias) => {
    if (joined.includes(norm(alias))) score += 4;
  });
  const numericCount = cells.filter((item) => /^-?\d+(\.\d+)?$/.test(item)).length;
  if (numericCount > cells.length / 2) score -= 6;
  if (/合计|总计|小计/.test(joined)) score -= 4;
  return score;
}

function detectHeaderRow(grid: unknown[][]) {
  let best = 0;
  let bestScore = -Infinity;
  const limit = Math.min(grid.length, 40);
  for (let i = 0; i < limit; i += 1) {
    const score = scoreHeaderRow(grid[i] ?? []);
    if (score > bestScore) {
      best = i;
      bestScore = score;
    }
  }
  return best;
}

function inferRole(header: string): { role?: FieldRole; confidence: number } {
  const value = norm(header);
  let bestRole: FieldRole | undefined;
  let bestScore = 0;
  (Object.entries(ROLE_ALIASES) as Array<[FieldRole, string[]]>).forEach(([role, aliases]) => {
    aliases.forEach((alias) => {
      const target = norm(alias);
      let score = 0;
      if (value === target) score = 1;
      else if (value.includes(target)) score = 0.82;
      else if (target.includes(value) && value.length >= 2) score = 0.55;
      if (score > bestScore) {
        bestScore = score;
        bestRole = role;
      }
    });
  });
  return { role: bestScore >= 0.5 ? bestRole : undefined, confidence: bestScore };
}

function analyzeColumns(headers: string[], rows: DataRow[]) {
  return headers.map((header, index): ColumnAnalysis => {
    const values = rows.map((row) => row.values[header]).filter((value) => !isBlank(value));
    const numericCount = values.filter((value) => text(value) !== "" && Number.isFinite(toNumber(value))).length;
    const guess = inferRole(header);
    return { index, header, role: guess.role, confidence: guess.confidence, numericRatio: values.length ? numericCount / values.length : 0, nonEmptyCount: values.length, examples: values.slice(0, 3).map(text) };
  });
}

function buildRows(grid: unknown[][], headerRow: number, headers: string[]) {
  const rows: DataRow[] = [];
  for (let r = headerRow + 1; r < grid.length; r += 1) {
    const raw = grid[r] ?? [];
    const nonEmpty = raw.filter((cell) => !isBlank(cell)).length;
    if (nonEmpty === 0) continue;
    const record: Record<string, DataCell> = {};
    headers.forEach((header, index) => {
      const cell = raw[index];
      record[header] = typeof cell === "number" ? cell : text(cell);
    });
    rows.push({ rowNumber: r + 1, values: record });
  }
  return rows;
}

function roleColumn(columns: ColumnAnalysis[], roles: FieldRole | FieldRole[]) {
  const list = Array.isArray(roles) ? roles : [roles];
  return columns.find((column) => column.role && list.includes(column.role));
}
function valueByRole(row: DataRow, analysis: SheetAnalysis, roles: FieldRole | FieldRole[]) {
  const column = roleColumn(analysis.columns, roles);
  return column ? row.values[column.header] : undefined;
}
function numberByRole(row: DataRow, analysis: SheetAnalysis, roles: FieldRole | FieldRole[]) {
  return toNumber(valueByRole(row, analysis, roles));
}
function roleScore(columns: ColumnAnalysis[], roles: FieldRole[]) {
  return roles.reduce((sum, role) => sum + (columns.some((column) => column.role === role) ? 1 : 0), 0);
}

function detectTableType(columns: ColumnAnalysis[]): { type: TableType; confidence: number } {
  const product = roleScore(columns, ["product", "price", "quantity", "sold", "purchase", "stockEnding", "amount"]);
  const payroll = roleScore(columns, ["employee", "basicSalary", "attendance", "commission", "bonus", "deduction", "grossSalary", "netSalary"]);
  const purchase = roleScore(columns, ["supplier", "product", "quantity", "price", "amount", "payable"]);
  const expense = roleScore(columns, ["date", "feeType", "amount", "remark"]);
  const collection = roleScore(columns, ["date", "paymentChannel", "paidAmount", "amount", "orderNo"]);
  const inventory = roleScore(columns, ["product", "stockOpening", "stockIn", "stockOut", "stockEnding"]);
  const candidates: Array<[TableType, number]> = [
    ["工资表", payroll],
    ["商品库存/销售报表", product],
    ["进货表", purchase],
    ["费用表", expense],
    ["收款表", collection],
    ["库存表", inventory],
    ["通用表格", 1]
  ];
  const sorted = candidates.sort((a, b) => b[1] - a[1]);
  const [type, score] = sorted[0];
  return { type, confidence: Math.min(0.98, Math.max(0.35, score / 7)) };
}

function inferRules(type: TableType, columns: ColumnAnalysis[]) {
  const rules: string[] = [];
  const has = (role: FieldRole) => columns.some((column) => column.role === role);
  if (has("quantity") && has("price") && (has("amount") || has("totalAmount"))) rules.push("金额 = 数量 × 单价");
  if ((has("stockOpening") || has("stockEnding")) && (has("stockIn") || has("purchase")) && (has("stockOut") || has("sold"))) rules.push("结存 = 期初/早班库存 + 入库/进货 - 出库/销量");
  if (has("grossSalary") && has("netSalary") && (has("deduction") || has("advance"))) rules.push("实发工资 = 应发工资 - 扣款 - 借支");
  if (has("basicSalary") && has("grossSalary") && (has("bonus") || has("commission") || has("overtime"))) rules.push("应发工资 = 基本工资 + 奖金 + 提成 + 加班费");
  if ((has("paidAmount") || has("amount")) && has("paymentChannel")) rules.push("按收款渠道汇总金额并检查空值/负数/异常值");
  if (type === "通用表格") rules.push("检查空值、重复记录、负数金额、合计行和数值异常");
  return rules;
}

function primaryValue(row: DataRow, analysis: SheetAnalysis) {
  return text(valueByRole(row, analysis, ["product", "employee", "name", "customer", "supplier", "orderNo", "sku"]));
}

function findAnomalies(analysis: SheetAnalysis): Anomaly[] {
  const out: Anomaly[] = [];
  const amountRole: FieldRole[] = ["amount", "totalAmount", "paidAmount", "payable", "receivable"];
  const seen = new Map<string, number>();
  analysis.rows.forEach((row) => {
    const primary = primaryValue(row, analysis);
    if (!primary) {
      out.push({ severity: "warning", type: "关键字段缺失", sheetName: analysis.sheetName, rowNumber: row.rowNumber, message: "本行没有识别到商品/员工/客户/单号等关键名称。" });
    } else {
      const key = `${primary}__${text(valueByRole(row, analysis, ["date"]))}`;
      if (seen.has(key)) out.push({ severity: "warning", type: "疑似重复", sheetName: analysis.sheetName, rowNumber: row.rowNumber, message: `与第 ${seen.get(key)} 行关键字段重复：${primary}` });
      else seen.set(key, row.rowNumber);
    }

    const qty = numberByRole(row, analysis, ["quantity", "sold", "stockOut"]);
    const price = numberByRole(row, analysis, "price");
    const amount = numberByRole(row, analysis, ["amount", "totalAmount"]);
    if (qty && price && amount && Math.abs(round2(qty * price) - amount) > 0.01) {
      out.push({ severity: "error", type: "金额不一致", sheetName: analysis.sheetName, rowNumber: row.rowNumber, message: "金额不等于数量 × 单价。", expected: round2(qty * price), actual: amount });
    }

    const opening = numberByRole(row, analysis, ["stockOpening"]);
    const stockIn = numberByRole(row, analysis, ["stockIn", "purchase"]);
    const stockOut = numberByRole(row, analysis, ["stockOut", "sold"]);
    const ending = numberByRole(row, analysis, "stockEnding");
    if ((opening || stockIn || stockOut) && ending && Math.abs(round2(opening + stockIn - stockOut) - ending) > 0.01) {
      out.push({ severity: "error", type: "库存不一致", sheetName: analysis.sheetName, rowNumber: row.rowNumber, message: "结存不等于期初/早班库存 + 入库/进货 - 出库/销量。", expected: round2(opening + stockIn - stockOut), actual: ending });
    }

    const gross = numberByRole(row, analysis, "grossSalary");
    const net = numberByRole(row, analysis, "netSalary");
    const deduction = numberByRole(row, analysis, "deduction");
    const advance = numberByRole(row, analysis, "advance");
    if (gross && net && Math.abs(round2(gross - deduction - advance) - net) > 0.01) {
      out.push({ severity: "error", type: "工资不一致", sheetName: analysis.sheetName, rowNumber: row.rowNumber, message: "实发工资不等于应发工资 - 扣款 - 借支。", expected: round2(gross - deduction - advance), actual: net });
    }

    analysis.columns.filter((column) => column.role && amountRole.includes(column.role) && column.numericRatio > 0.6).forEach((column) => {
      const value = toNumber(row.values[column.header]);
      if (value < 0) out.push({ severity: "warning", type: "负数金额", sheetName: analysis.sheetName, rowNumber: row.rowNumber, field: column.header, message: `${column.header} 出现负数。`, actual: value });
    });
  });
  return out;
}

function applyTemplate(analysis: SheetAnalysis, templates: SavedTemplate[]): SheetAnalysis {
  let best: { template: SavedTemplate; score: number } | undefined;
  templates.forEach((template) => {
    const hits = template.columns.filter((templateColumn) => analysis.columns.some((column) => norm(column.header) === norm(templateColumn.header))).length;
    const score = template.columns.length ? hits / template.columns.length : 0;
    if (!best || score > best.score) best = { template, score };
  });
  if (!best || best.score < 0.55) return analysis;
  const nextColumns = analysis.columns.map((column) => {
    const templateColumn = best?.template.columns.find((item) => norm(item.header) === norm(column.header) || item.index === column.index);
    return templateColumn ? { ...column, role: templateColumn.role, confidence: Math.max(column.confidence, 0.95) } : column;
  });
  const next: SheetAnalysis = { ...analysis, columns: nextColumns, matchedTemplate: best.template.name };
  return { ...next, rules: inferRules(next.type, next.columns), anomalies: findAnomalies(next) };
}

function analyzeWorkbook(workbook: XLSX.WorkBook, fileName: string, templates: SavedTemplate[]): WorkbookAnalysis {
  const sheets = workbook.SheetNames.map((sheetName): SheetAnalysis | null => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return null;
    const grid = sheetToGrid(sheet);
    const headerRow = detectHeaderRow(grid);
    const used = new Set<string>();
    const headers = (grid[headerRow] ?? []).map((cell, index) => uniqueHeader(cell, index, used));
    const rows = buildRows(grid, headerRow, headers);
    const columns = analyzeColumns(headers, rows);
    const detected = detectTableType(columns);
    const analysis: SheetAnalysis = { sheetName, type: detected.type, confidence: detected.confidence, headerRow, columns, rows, rules: inferRules(detected.type, columns), anomalies: [], matchedTemplate: undefined };
    const withAnomalies = { ...analysis, anomalies: findAnomalies(analysis) };
    return applyTemplate(withAnomalies, templates);
  }).filter((sheet): sheet is SheetAnalysis => Boolean(sheet));
  return { fileName, createdAt: new Date().toLocaleString("zh-CN", { hour12: false }), sheets };
}

function loadTemplates() {
  if (typeof window === "undefined") return [] as SavedTemplate[];
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) as SavedTemplate[] : [];
  } catch {
    return [];
  }
}
function saveTemplates(templates: SavedTemplate[]) {
  if (typeof window !== "undefined") window.localStorage.setItem(STORAGE_KEY, JSON.stringify(templates));
}
function exportAnalysis(analysis: WorkbookAnalysis) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(analysis.sheets.map((sheet) => ({ Sheet: sheet.sheetName, 类型: sheet.type, 置信度: Math.round(sheet.confidence * 100) + "%", 表头行: sheet.headerRow + 1, 数据行: sheet.rows.length, 异常数: sheet.anomalies.length, 套用模板: sheet.matchedTemplate ?? "" }))), "核对总览");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(analysis.sheets.flatMap((sheet) => sheet.columns.map((column) => ({ Sheet: sheet.sheetName, 列: colName(column.index), 表头: column.header, 识别字段: column.role ? ROLE_LABELS[column.role] : "未识别", 置信度: Math.round(column.confidence * 100) + "%", 非空数量: column.nonEmptyCount, 示例: column.examples.join(" / ") })))), "字段识别");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(analysis.sheets.flatMap((sheet) => sheet.rules.map((rule, index) => ({ Sheet: sheet.sheetName, 序号: index + 1, 自动规则: rule })))), "自动规则");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(analysis.sheets.flatMap((sheet) => sheet.anomalies.map((item, index) => ({ Sheet: sheet.sheetName, 序号: index + 1, 级别: item.severity, 类型: item.type, 行号: item.rowNumber ?? "", 字段: item.field ?? "", 说明: item.message, 应为: item.expected ?? "", 实际: item.actual ?? "" })))), "异常明细");
  XLSX.writeFile(wb, `${analysis.fileName.replace(/\.xlsx?$/i, "")}-智能核对报告.xlsx`);
}

function Card({ title, value, desc, tone = "default" }: { title: string; value: string; desc?: string; tone?: "default" | "danger" }) {
  const cls = tone === "danger" ? "border-red-500 bg-red-600 text-white" : "border bg-white text-slate-900";
  return <div className={`rounded-2xl p-5 shadow-sm ${cls}`}><div className={`text-sm ${tone === "danger" ? "text-red-100" : "text-slate-500"}`}>{title}</div><div className="mt-2 text-2xl font-bold">{value}</div>{desc ? <div className={`mt-1 text-xs ${tone === "danger" ? "text-red-100" : "text-slate-500"}`}>{desc}</div> : null}</div>;
}
function Badge({ children, tone = "gray" }: { children: ReactNode; tone?: "gray" | "green" | "red" | "yellow" }) {
  const cls = { gray: "bg-slate-100 text-slate-700", green: "bg-green-100 text-green-700", red: "bg-red-100 text-red-700", yellow: "bg-yellow-100 text-yellow-800" }[tone];
  return <span className={`rounded-full px-2 py-1 text-xs font-medium ${cls}`}>{children}</span>;
}

export default function UniversalReportApp() {
  const [tab, setTab] = useState<MainTab>("smart");
  const [templates, setTemplates] = useState<SavedTemplate[]>([]);
  const [analysis, setAnalysis] = useState<WorkbookAnalysis | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);

  useEffect(() => {
    setTemplates(loadTemplates());
    document.title = "万能表格自动核对助手";
  }, []);

  async function handleFile(file: File) {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true, cellFormula: true, cellNF: true, cellStyles: true });
    const next = analyzeWorkbook(workbook, file.name, templates);
    setAnalysis(next);
    setActiveSheet(0);
    setTab("smart");
  }

  function saveActiveTemplate() {
    if (!analysis?.sheets[activeSheet]) return;
    const sheet = analysis.sheets[activeSheet];
    const name = window.prompt("模板名称", `${sheet.type}-${sheet.sheetName}`);
    if (!name) return;
    const template: SavedTemplate = { id: `${Date.now()}`, name, tableType: sheet.type, createdAt: new Date().toLocaleString("zh-CN", { hour12: false }), updatedAt: new Date().toLocaleString("zh-CN", { hour12: false }), headerRow: sheet.headerRow, columns: sheet.columns.map((column) => ({ header: column.header, role: column.role, index: column.index })) };
    const next = [template, ...templates];
    setTemplates(next);
    saveTemplates(next);
  }

  function deleteTemplate(id: string) {
    const next = templates.filter((template) => template.id !== id);
    setTemplates(next);
    saveTemplates(next);
  }

  const totalAnomalies = useMemo(() => analysis?.sheets.reduce((sum, sheet) => sum + sheet.anomalies.length, 0) ?? 0, [analysis]);
  const currentSheet = analysis?.sheets[activeSheet];

  return <main className="min-h-screen bg-slate-50 p-4 md:p-8"><div className="mx-auto max-w-7xl">
    <header className="mb-6 rounded-3xl bg-slate-950 p-6 text-white shadow-sm"><p className="text-sm text-slate-300">Universal Spreadsheet Reconciliation Engine</p><h1 className="mt-1 text-2xl font-bold md:text-4xl">万能表格自动核对助手</h1><p className="mt-2 max-w-4xl text-sm text-slate-300">任何 Excel 先进入智能识别：自动判断类型、字段、规则和异常；当前商品报表保留专用核对逻辑。</p></header>
    <nav className="mb-6 flex flex-wrap gap-2">{MAIN_TABS.map(([key, label]) => <button key={key} onClick={() => setTab(key)} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === key ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>{label}</button>)}</nav>

    {tab === "smart" ? <section className="space-y-6">
      <div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">智能上传</h2><p className="mt-1 text-sm text-slate-500">上传任意 Excel。系统会先做通用分析；如果是商品报表，可切换到“商品报表专用”使用库存、销售、收款专用规则。</p></div><label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-slate-950 px-4 py-3 text-sm font-medium text-white">选择 Excel<input className="hidden" type="file" accept=".xlsx,.xls" onChange={(event: ChangeEvent<HTMLInputElement>) => event.target.files?.[0] && handleFile(event.target.files[0])} /></label></div></div>

      {!analysis ? <div className="rounded-2xl border bg-white p-6 shadow-sm"><h2 className="text-lg font-bold">等待上传表格</h2><p className="mt-2 text-sm text-slate-600">第一次见到的新表格，系统会自动推断字段和规则；你可以把识别结果保存成模板。之后同类表格会自动套用模板。</p></div> : null}

      {analysis ? <><div className="grid gap-4 md:grid-cols-4"><Card title="文件" value={analysis.fileName} desc={analysis.createdAt} /><Card title="工作表" value={`${analysis.sheets.length} 个`} /><Card title="异常" value={`${totalAnomalies} 条`} tone={totalAnomalies ? "danger" : "default"} /><Card title="模板" value={`${templates.length} 个`} desc="本地浏览器保存" /></div>
      <div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-wrap items-center justify-between gap-3"><div className="flex flex-wrap gap-2">{analysis.sheets.map((sheet, index) => <button key={sheet.sheetName} onClick={() => setActiveSheet(index)} className={`rounded-xl px-3 py-2 text-sm ${index === activeSheet ? "bg-slate-950 text-white" : "bg-slate-100 text-slate-700"}`}>{sheet.sheetName}</button>)}</div><div className="flex flex-wrap gap-2"><button onClick={saveActiveTemplate} className="rounded-xl border px-4 py-2 text-sm">保存为模板</button><button onClick={() => exportAnalysis(analysis)} className="rounded-xl bg-green-700 px-4 py-2 text-sm text-white">导出智能核对报告</button></div></div></div></> : null}

      {currentSheet ? <div className="space-y-5"><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">{currentSheet.sheetName}</h2><p className="mt-1 text-sm text-slate-500">识别类型：<b>{currentSheet.type}</b> / 置信度 {Math.round(currentSheet.confidence * 100)}% / 表头第 {currentSheet.headerRow + 1} 行 / 数据 {currentSheet.rows.length} 行 {currentSheet.matchedTemplate ? `/ 已套用模板：${currentSheet.matchedTemplate}` : ""}</p></div>{currentSheet.type === "商品库存/销售报表" ? <button onClick={() => setTab("product")} className="rounded-xl bg-slate-950 px-4 py-2 text-sm text-white">进入商品专用核对</button> : null}</div><div className="mt-4 grid gap-3 md:grid-cols-3">{currentSheet.rules.map((rule) => <div key={rule} className="rounded-xl bg-slate-50 p-3 text-sm text-slate-700">{rule}</div>)}</div></div>
      <div className="rounded-2xl border bg-white p-5 shadow-sm"><h3 className="font-bold">字段识别</h3><div className="table-scroll mt-3"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["列", "表头", "识别字段", "置信度", "非空", "示例"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{currentSheet.columns.map((column) => <tr key={column.index} className="border-t"><td className="p-3">{colName(column.index)}</td><td className="p-3">{column.header}</td><td className="p-3">{column.role ? <Badge tone="green">{ROLE_LABELS[column.role]}</Badge> : <Badge>未识别</Badge>}</td><td className="p-3">{Math.round(column.confidence * 100)}%</td><td className="p-3">{column.nonEmptyCount}</td><td className="p-3 text-slate-500">{column.examples.join(" / ")}</td></tr>)}</tbody></table></div></div>
      <div className="rounded-2xl border bg-white p-5 shadow-sm"><h3 className="font-bold">异常明细</h3>{currentSheet.anomalies.length ? <div className="table-scroll mt-3"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["级别", "类型", "行号", "字段", "说明", "应为", "实际"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{currentSheet.anomalies.map((item, index) => <tr key={`${item.type}-${index}`} className="border-t"><td className="p-3"><Badge tone={item.severity === "error" ? "red" : "yellow"}>{item.severity}</Badge></td><td className="p-3">{item.type}</td><td className="p-3">{item.rowNumber ?? ""}</td><td className="p-3">{item.field ?? ""}</td><td className="p-3">{item.message}</td><td className="p-3">{item.expected ?? ""}</td><td className="p-3">{item.actual ?? ""}</td></tr>)}</tbody></table></div> : <p className="mt-3 text-sm text-green-700">当前 Sheet 未发现明显异常。</p>}</div></div> : null}
    </section> : null}

    {tab === "product" ? <section className="space-y-4"><div className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">商品报表专用核对</h2><p className="mt-1 text-sm text-slate-500">这里保留当前商品报表的专用逻辑：自动库存盘点、交接异常、金额核对、收款统计、写回原表。</p></div><ProductReportApp /></section> : null}

    {tab === "templates" ? <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">模板库</h2><p className="mt-1 text-sm text-slate-500">模板保存在当前浏览器。已学习过的表格，后续上传会优先套用字段映射。</p><div className="mt-4 space-y-3">{templates.length ? templates.map((template) => <div key={template.id} className="rounded-xl border p-4"><div className="flex items-center justify-between gap-3"><div><div className="font-bold">{template.name}</div><div className="mt-1 text-sm text-slate-500">{template.tableType} / 字段 {template.columns.length} 个 / 创建 {template.createdAt}</div></div><button onClick={() => deleteTemplate(template.id)} className="rounded-xl border px-4 py-2 text-sm text-red-600">删除</button></div></div>) : <p className="text-sm text-slate-600">还没有模板。先在“智能上传”里识别一个表格，然后保存为模板。</p>}</div></section> : null}

    {tab === "about" ? <section className="rounded-2xl border bg-white p-6 shadow-sm"><h2 className="text-lg font-bold">系统边界</h2><div className="mt-3 space-y-3 text-sm leading-6 text-slate-700"><p>已学习过的模板可以做到接近全自动。第一次见到的表格，系统会自动识别并生成核对建议，但业务规则仍建议你确认一次。</p><p>商品报表是内置专用模块，规则更强；其他表格先走通用识别，包括工资、费用、进货、收款、库存和普通明细表。</p><p>下一步可以继续加入“规则编辑器”，让你在页面里配置：某列 = 某列 + 某列 - 某列，保存后后续自动执行。</p></div></section> : null}
  </div></main>;
}
