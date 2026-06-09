"use client";

import { useMemo, useState, type ChangeEvent, type ReactNode } from "react";
import * as XLSX from "xlsx";
import ProductReportApp from "./InventoryPageConfigurable";
import { classifyGlobalProduct } from "./globalProductCatalog";

type MainTab = "smart" | "product" | "guide";
type ReportTab = "overview" | "rules" | "preview" | "summary" | "issues" | "auto" | "export";
type Severity = "error" | "warning" | "info";
type TableType = "商品销售/库存表" | "工资/考勤表" | "进货采购表" | "费用报销表" | "收款对账表" | "应收应付表" | "订单流水表" | "客户供应商表" | "资产设备表" | "通用业务表";
type Role = "date" | "product" | "employee" | "customer" | "supplier" | "orderNo" | "sku" | "barcode" | "category" | "quantity" | "price" | "cost" | "amount" | "totalAmount" | "paidAmount" | "receivable" | "payable" | "stockOpening" | "stockIn" | "stockOut" | "stockEnding" | "purchase" | "sold" | "basicSalary" | "attendance" | "overtime" | "commission" | "bonus" | "deduction" | "advance" | "grossSalary" | "netSalary" | "feeType" | "paymentChannel" | "department" | "position" | "asset" | "assetNo" | "status" | "remark";

type DataRow = { rowNumber: number; values: Record<string, string | number> };
type ColumnInfo = { index: number; header: string; role?: Role; confidence: number; examples: string[]; nonEmpty: number };
type Issue = { severity: Severity; type: string; rowNumber?: number; object?: string; message: string; expected?: string | number; actual?: string | number };
type SheetReport = { sheetName: string; type: TableType; confidence: number; headerRow: number; rows: DataRow[]; columns: ColumnInfo[]; rules: string[]; issues: Issue[] };
type WorkbookReport = { fileName: string; createdAt: string; sheets: SheetReport[] };

type SummaryRow = Record<string, string | number>;

const MAIN_TABS: Array<[MainTab, string]> = [["smart", "智能上传"], ["product", "商品报表专用"], ["guide", "功能说明"]];
const REPORT_TABS: Array<[ReportTab, string]> = [["overview", "总览"], ["rules", "核对规则"], ["preview", "识别预览"], ["summary", "汇总报表"], ["issues", "核对异常"], ["auto", "自动结果"], ["export", "导出"]];

const ROLE_LABEL: Record<Role, string> = {
  date: "日期",
  product: "商品",
  employee: "员工",
  customer: "客户",
  supplier: "供应商",
  orderNo: "单号",
  sku: "SKU/货号",
  barcode: "条码",
  category: "分类",
  quantity: "数量",
  price: "单价",
  cost: "成本",
  amount: "金额",
  totalAmount: "总金额",
  paidAmount: "实收金额",
  receivable: "应收",
  payable: "应付",
  stockOpening: "期初库存",
  stockIn: "入库",
  stockOut: "出库",
  stockEnding: "结存",
  purchase: "采购/进货",
  sold: "销量",
  basicSalary: "基本工资",
  attendance: "出勤",
  overtime: "加班",
  commission: "提成",
  bonus: "奖金/补贴",
  deduction: "扣款",
  advance: "借支",
  grossSalary: "应发工资",
  netSalary: "实发工资",
  feeType: "费用类型",
  paymentChannel: "收款渠道",
  department: "部门",
  position: "岗位",
  asset: "资产/设备",
  assetNo: "资产编号",
  status: "状态",
  remark: "备注"
};

const ROLE_ALIASES: Record<Role, string[]> = {
  date: ["日期", "时间", "date", "day", "营业日"],
  product: ["商品", "商品名称", "产品", "品名", "货品", "物品", "标题", "product", "item", "goods", "товар", "منتج"],
  employee: ["员工", "员工姓名", "姓名", "人员", "工号", "employee", "staff", "worker"],
  customer: ["客户", "客户名称", "会员", "买家", "customer", "client", "buyer"],
  supplier: ["供应商", "厂家", "供货商", "采购商", "supplier", "vendor"],
  orderNo: ["单号", "订单号", "流水号", "编号", "票号", "order", "orderid"],
  sku: ["sku", "货号", "编码", "商品编码", "артикул"],
  barcode: ["条码", "条形码", "barcode", "ean", "upc"],
  category: ["分类", "类别", "品类", "类型", "category", "type"],
  quantity: ["数量", "件数", "个数", "包数", "条数", "瓶数", "qty", "quantity", "count"],
  price: ["价格", "单价", "售价", "销售价", "price", "unitprice"],
  cost: ["成本", "成本价", "进价", "cost", "purchaseprice"],
  amount: ["金额", "小计", "销售额", "收入", "支出", "费用", "amount", "sum"],
  totalAmount: ["合计", "总计", "总金额", "总价", "合计金额", "total"],
  paidAmount: ["实收", "实收金额", "到账", "收款金额", "已收", "paid", "received"],
  receivable: ["应收", "应收金额", "应收款", "receivable"],
  payable: ["应付", "应付金额", "应付款", "payable"],
  stockOpening: ["期初", "初始库存", "早班库存", "上月结存", "opening"],
  stockIn: ["入库", "进货", "采购", "补货", "stockin", "inbound"],
  stockOut: ["出库", "销量", "售卖", "销售数量", "消耗", "stockout", "outbound"],
  stockEnding: ["结存", "库存", "剩余", "月末库存", "夜班库存", "stock", "ending"],
  purchase: ["进货", "采购", "补货", "purchase", "buy"],
  sold: ["销量", "售卖", "销售数量", "卖出", "sold", "salesqty"],
  basicSalary: ["基本工资", "底薪", "月薪", "工资标准", "salary", "base"],
  attendance: ["出勤", "出勤天数", "天数", "工时", "考勤", "attendance", "hours", "days"],
  overtime: ["加班", "加班费", "overtime"],
  commission: ["提成", "业绩提成", "销售提成", "commission"],
  bonus: ["奖金", "补贴", "满勤", "津贴", "奖励", "bonus", "allowance"],
  deduction: ["扣款", "罚款", "扣除", "社保", "个税", "deduction", "fine"],
  advance: ["借支", "预支", "借款", "advance", "loan"],
  grossSalary: ["应发", "应发工资", "工资合计", "gross"],
  netSalary: ["实发", "实发工资", "到手", "实际发放", "net"],
  feeType: ["费用类型", "费用项目", "支出项目", "科目", "expense", "feetype"],
  paymentChannel: ["渠道", "收款渠道", "支付方式", "付款方式", "平台", "channel", "paymentmethod"],
  department: ["部门", "门店", "组织", "department", "team"],
  position: ["岗位", "职位", "职务", "position", "role"],
  asset: ["资产", "设备", "固定资产", "asset", "device", "equipment"],
  assetNo: ["资产编号", "设备编号", "assetno", "serial"],
  status: ["状态", "进度", "status", "state"],
  remark: ["备注", "说明", "原因", "note", "remark", "comment"]
};

function text(value: unknown) { return String(value ?? "").trim(); }
function norm(value: unknown) { return text(value).replace(/[\s\-_，,。.;；:：/\\|()（）\[\]【】{}<>《》"'“”‘’]+/g, "").toLowerCase(); }
function isBlank(value: unknown) { return value === null || value === undefined || text(value) === ""; }
function toNumber(value: unknown) { if (typeof value === "number") return Number.isFinite(value) ? value : 0; const cleaned = text(value).replace(/,/g, "").replace(/[￥¥元%\s]/g, ""); if (!cleaned) return 0; const parsed = Number(cleaned); return Number.isFinite(parsed) ? parsed : 0; }
function round2(value: number) { return Math.round((value + Number.EPSILON) * 100) / 100; }
function money(value: number) { return round2(value).toFixed(2); }
function colName(index: number) { return XLSX.utils.encode_col(index); }

function inferRole(header: string): { role?: Role; confidence: number } {
  const h = norm(header);
  let role: Role | undefined;
  let confidence = 0;
  (Object.entries(ROLE_ALIASES) as Array<[Role, string[]]>).forEach(([candidate, aliases]) => aliases.forEach((alias) => {
    const a = norm(alias);
    let score = 0;
    if (h === a) score = 1;
    else if (h.includes(a)) score = 0.86;
    else if (a.includes(h) && h.length >= 2) score = 0.55;
    if (score > confidence) { role = candidate; confidence = score; }
  }));
  return { role: confidence >= 0.5 ? role : undefined, confidence };
}
function scoreHeaderRow(row: unknown[]) {
  const joined = norm(row.map(text).filter(Boolean).join(" "));
  let score = row.filter((cell) => !isBlank(cell)).length;
  Object.values(ROLE_ALIASES).flat().forEach((alias) => { if (joined.includes(norm(alias))) score += 2; });
  if (/合计|总计|小计/.test(joined)) score -= 4;
  return score;
}
function detectHeaderRow(grid: unknown[][]) { let best = 0; let bestScore = -Infinity; for (let i = 0; i < Math.min(35, grid.length); i += 1) { const score = scoreHeaderRow(grid[i] ?? []); if (score > bestScore) { best = i; bestScore = score; } } return best; }
function uniqueHeader(raw: unknown, index: number, used: Set<string>) { const base = text(raw) || `列${colName(index)}`; let name = base; let suffix = 2; while (used.has(name)) { name = `${base}_${suffix}`; suffix += 1; } used.add(name); return name; }
function buildRows(grid: unknown[][], headerRow: number, headers: string[]) { const rows: DataRow[] = []; for (let r = headerRow + 1; r < grid.length; r += 1) { const raw = grid[r] ?? []; if (!raw.some((cell) => !isBlank(cell))) continue; const values: Record<string, string | number> = {}; headers.forEach((header, index) => { const cell = raw[index]; values[header] = typeof cell === "number" ? cell : text(cell); }); rows.push({ rowNumber: r + 1, values }); } return rows; }
function analyzeColumns(headers: string[], rows: DataRow[]) { return headers.map((header, index): ColumnInfo => { const values = rows.map((row) => row.values[header]).filter((value) => !isBlank(value)); const guess = inferRole(header); return { index, header, role: guess.role, confidence: guess.confidence, examples: values.slice(0, 3).map(text), nonEmpty: values.length }; }); }
function roleColumn(sheet: SheetReport, roles: Role | Role[]) { const list = Array.isArray(roles) ? roles : [roles]; return sheet.columns.find((column) => column.role && list.includes(column.role)); }
function roleValue(row: DataRow, sheet: SheetReport, roles: Role | Role[]) { const column = roleColumn(sheet, roles); return column ? row.values[column.header] : undefined; }
function roleNumber(row: DataRow, sheet: SheetReport, roles: Role | Role[]) { return toNumber(roleValue(row, sheet, roles)); }
function hasRole(sheet: SheetReport, role: Role) { return sheet.columns.some((column) => column.role === role); }
function countRoles(columns: ColumnInfo[], roles: Role[]) { return roles.reduce((sum, role) => sum + (columns.some((column) => column.role === role) ? 1 : 0), 0); }
function detectType(columns: ColumnInfo[]): { type: TableType; confidence: number } {
  const candidates: Array<[TableType, number]> = [
    ["商品销售/库存表", countRoles(columns, ["product", "sku", "quantity", "price", "amount", "stockEnding", "sold"])],
    ["工资/考勤表", countRoles(columns, ["employee", "basicSalary", "attendance", "grossSalary", "netSalary", "deduction"])],
    ["进货采购表", countRoles(columns, ["supplier", "product", "purchase", "quantity", "price", "payable", "amount"])],
    ["费用报销表", countRoles(columns, ["date", "feeType", "amount", "department", "remark"])],
    ["收款对账表", countRoles(columns, ["date", "paymentChannel", "paidAmount", "amount", "orderNo"])],
    ["应收应付表", countRoles(columns, ["customer", "supplier", "receivable", "payable", "paidAmount"])],
    ["订单流水表", countRoles(columns, ["orderNo", "product", "customer", "quantity", "amount", "status"])],
    ["客户供应商表", countRoles(columns, ["customer", "supplier", "receivable", "payable", "amount"])],
    ["资产设备表", countRoles(columns, ["asset", "assetNo", "department", "status", "amount"])],
    ["通用业务表", 1]
  ];
  const [type, score] = candidates.sort((a, b) => b[1] - a[1])[0];
  return { type, confidence: Math.min(0.98, Math.max(0.35, score / 7)) };
}
function buildRules(sheet: SheetReport) {
  const rules: string[] = [];
  if ((hasRole(sheet, "quantity") || hasRole(sheet, "sold")) && hasRole(sheet, "price") && (hasRole(sheet, "amount") || hasRole(sheet, "totalAmount"))) rules.push("金额 = 数量 × 单价");
  if ((hasRole(sheet, "stockOpening") || hasRole(sheet, "stockEnding")) && (hasRole(sheet, "stockIn") || hasRole(sheet, "purchase")) && (hasRole(sheet, "stockOut") || hasRole(sheet, "sold"))) rules.push("结存 = 期初库存 + 入库/进货 - 出库/销量");
  if (hasRole(sheet, "grossSalary") && hasRole(sheet, "netSalary")) rules.push("实发工资 = 应发工资 - 扣款 - 借支");
  if (hasRole(sheet, "basicSalary") && hasRole(sheet, "grossSalary")) rules.push("应发工资 = 基本工资 + 加班 + 提成 + 奖金");
  if (hasRole(sheet, "receivable") || hasRole(sheet, "payable")) rules.push("应收/应付/实收金额差异核对");
  if (hasRole(sheet, "paymentChannel") || hasRole(sheet, "paidAmount")) rules.push("按收款渠道汇总并检查异常金额");
  rules.push("重复记录、关键字段缺失、负数金额检查");
  return Array.from(new Set(rules));
}
function objectName(row: DataRow, sheet: SheetReport) { return text(roleValue(row, sheet, ["product", "sku", "barcode", "employee", "customer", "supplier", "orderNo", "asset"])); }
function buildIssues(sheet: SheetReport) {
  const issues: Issue[] = [];
  const seen = new Map<string, number>();
  sheet.rows.forEach((row) => {
    const object = objectName(row, sheet);
    if (!object) issues.push({ severity: "warning", type: "关键字段缺失", rowNumber: row.rowNumber, message: "本行缺少可识别的商品、员工、客户、供应商或单号。" });
    else if (seen.has(object)) issues.push({ severity: "warning", type: "疑似重复", rowNumber: row.rowNumber, object, message: `与第 ${seen.get(object)} 行关键对象重复。` });
    else seen.set(object, row.rowNumber);

    const qty = roleNumber(row, sheet, ["quantity", "sold", "stockOut"]);
    const price = roleNumber(row, sheet, "price");
    const amount = roleNumber(row, sheet, ["amount", "totalAmount"]);
    if (qty && price && amount && Math.abs(round2(qty * price) - amount) > 0.01) issues.push({ severity: "error", type: "金额不一致", rowNumber: row.rowNumber, object, message: "金额不等于数量 × 单价。", expected: round2(qty * price), actual: amount });

    const opening = roleNumber(row, sheet, "stockOpening");
    const stockIn = roleNumber(row, sheet, ["stockIn", "purchase"]);
    const stockOut = roleNumber(row, sheet, ["stockOut", "sold"]);
    const ending = roleNumber(row, sheet, "stockEnding");
    if ((opening || stockIn || stockOut) && ending && Math.abs(round2(opening + stockIn - stockOut) - ending) > 0.01) issues.push({ severity: "error", type: "库存不一致", rowNumber: row.rowNumber, object, message: "结存不等于期初库存 + 入库 - 出库。", expected: round2(opening + stockIn - stockOut), actual: ending });
    if (ending < 0) issues.push({ severity: "warning", type: "负库存", rowNumber: row.rowNumber, object, message: "结存库存为负数。", actual: ending });

    const base = roleNumber(row, sheet, "basicSalary");
    const overtime = roleNumber(row, sheet, "overtime");
    const commission = roleNumber(row, sheet, "commission");
    const bonus = roleNumber(row, sheet, "bonus");
    const gross = roleNumber(row, sheet, "grossSalary");
    const deduction = roleNumber(row, sheet, "deduction");
    const advance = roleNumber(row, sheet, "advance");
    const net = roleNumber(row, sheet, "netSalary");
    if (gross && base && Math.abs(round2(base + overtime + commission + bonus) - gross) > 0.01) issues.push({ severity: "error", type: "应发工资不一致", rowNumber: row.rowNumber, object, message: "应发工资不等于基本工资 + 加班 + 提成 + 奖金。", expected: round2(base + overtime + commission + bonus), actual: gross });
    if (gross && net && Math.abs(round2(gross - deduction - advance) - net) > 0.01) issues.push({ severity: "error", type: "实发工资不一致", rowNumber: row.rowNumber, object, message: "实发工资不等于应发工资 - 扣款 - 借支。", expected: round2(gross - deduction - advance), actual: net });
  });
  return issues;
}
function analyzeWorkbook(workbook: XLSX.WorkBook, fileName: string): WorkbookReport {
  const sheets = workbook.SheetNames.map((sheetName): SheetReport | null => {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return null;
    const grid = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, defval: null });
    const headerRow = detectHeaderRow(grid);
    const used = new Set<string>();
    const headers = (grid[headerRow] ?? []).map((cell, index) => uniqueHeader(cell, index, used));
    const rows = buildRows(grid, headerRow, headers);
    const columns = analyzeColumns(headers, rows);
    const detected = detectType(columns);
    const base: SheetReport = { sheetName, type: detected.type, confidence: detected.confidence, headerRow, rows, columns, rules: [], issues: [] };
    const withRules = { ...base, rules: buildRules(base) };
    return { ...withRules, issues: buildIssues(withRules) };
  }).filter((sheet): sheet is SheetReport => Boolean(sheet));
  return { fileName, createdAt: new Date().toLocaleString("zh-CN", { hour12: false }), sheets };
}
function groupBy(sheet: SheetReport, role: Role | Role[], amountRoles: Role | Role[]) {
  const map = new Map<string, { name: string; count: number; quantity: number; amount: number }>();
  sheet.rows.forEach((row) => {
    const name = text(roleValue(row, sheet, role)) || "未识别";
    const old = map.get(name) ?? { name, count: 0, quantity: 0, amount: 0 };
    old.count += 1;
    old.quantity += roleNumber(row, sheet, ["quantity", "sold", "stockOut"]);
    old.amount += roleNumber(row, sheet, amountRoles);
    map.set(name, old);
  });
  return Array.from(map.values()).sort((a, b) => b.amount - a.amount || b.quantity - a.quantity || b.count - a.count);
}
function exportReport(report: WorkbookReport) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(report.sheets.map((sheet) => ({ Sheet: sheet.sheetName, 类型: sheet.type, 置信度: `${Math.round(sheet.confidence * 100)}%`, 表头行: sheet.headerRow + 1, 数据行: sheet.rows.length, 规则数: sheet.rules.length, 异常数: sheet.issues.length }))), "总览");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(report.sheets.flatMap((sheet) => sheet.columns.map((column) => ({ Sheet: sheet.sheetName, 列: colName(column.index), 表头: column.header, 字段: column.role ? ROLE_LABEL[column.role] : "未识别", 置信度: `${Math.round(column.confidence * 100)}%`, 示例: column.examples.join(" / ") })))), "识别预览");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(report.sheets.flatMap((sheet) => sheet.rules.map((rule, index) => ({ Sheet: sheet.sheetName, 序号: index + 1, 核对规则: rule })))), "核对规则");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(report.sheets.flatMap((sheet) => sheet.issues.map((issue, index) => ({ Sheet: sheet.sheetName, 序号: index + 1, 级别: issue.severity, 类型: issue.type, 行号: issue.rowNumber ?? "", 对象: issue.object ?? "", 说明: issue.message, 应为: issue.expected ?? "", 实际: issue.actual ?? "" })))), "核对异常");
  XLSX.writeFile(workbook, `${report.fileName.replace(/\.xlsx?$/i, "")}-结构化核对报告.xlsx`);
}
function Card({ title, value, desc, danger = false }: { title: string; value: string; desc?: string; danger?: boolean }) { return <div className={`rounded-2xl p-5 shadow-sm ${danger ? "bg-red-600 text-white" : "border bg-white text-slate-900"}`}><div className={`text-sm ${danger ? "text-red-100" : "text-slate-500"}`}>{title}</div><div className="mt-2 text-2xl font-bold">{value}</div>{desc ? <div className={`mt-1 text-xs ${danger ? "text-red-100" : "text-slate-500"}`}>{desc}</div> : null}</div>; }
function Badge({ children, tone = "gray" }: { children: ReactNode; tone?: "gray" | "green" | "red" | "yellow" }) { const cls = { gray: "bg-slate-100 text-slate-700", green: "bg-green-100 text-green-700", red: "bg-red-100 text-red-700", yellow: "bg-yellow-100 text-yellow-800" }[tone]; return <span className={`rounded-full px-2 py-1 text-xs font-medium ${cls}`}>{children}</span>; }
function DataTable({ rows, headers }: { rows: SummaryRow[]; headers: string[] }) { return <div className="table-scroll mt-3 max-h-96 overflow-auto"><table className="min-w-full text-sm"><thead className="sticky top-0 bg-slate-50 text-left"><tr>{headers.map((header) => <th key={header} className="p-3">{header}</th>)}</tr></thead><tbody>{rows.map((row, index) => <tr key={index} className="border-t">{headers.map((header) => <td key={header} className="p-3">{row[header]}</td>)}</tr>)}</tbody></table></div>; }
function Section({ title, desc, children }: { title: string; desc?: string; children: ReactNode }) { return <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">{title}</h2>{desc ? <p className="mt-1 text-sm text-slate-500">{desc}</p> : null}{children}</section>; }
function SummaryTables({ sheet }: { sheet: SheetReport }) {
  const productRows = groupBy(sheet, ["product", "sku", "barcode"], ["amount", "totalAmount", "paidAmount"]);
  const employeeRows = groupBy(sheet, "employee", ["netSalary", "grossSalary", "amount"]);
  const supplierRows = groupBy(sheet, "supplier", ["amount", "payable", "totalAmount"]);
  const customerRows = groupBy(sheet, "customer", ["amount", "receivable", "paidAmount"]);
  const channelRows = groupBy(sheet, "paymentChannel", ["paidAmount", "amount", "totalAmount"]);
  const feeRows = groupBy(sheet, "feeType", "amount");
  const categoryRows = groupBy(sheet, "category", ["amount", "totalAmount"]);
  const rows = hasRole(sheet, "product") || hasRole(sheet, "sku") ? productRows : hasRole(sheet, "employee") ? employeeRows : hasRole(sheet, "supplier") ? supplierRows : hasRole(sheet, "customer") ? customerRows : hasRole(sheet, "paymentChannel") ? channelRows : hasRole(sheet, "feeType") ? feeRows : hasRole(sheet, "category") ? categoryRows : [];
  const label = hasRole(sheet, "product") || hasRole(sheet, "sku") ? "商品/SKU" : hasRole(sheet, "employee") ? "员工" : hasRole(sheet, "supplier") ? "供应商" : hasRole(sheet, "customer") ? "客户" : hasRole(sheet, "paymentChannel") ? "收款渠道" : hasRole(sheet, "feeType") ? "费用类型" : hasRole(sheet, "category") ? "分类" : "对象";
  if (!rows.length) return <p className="mt-3 text-sm text-slate-600">当前表没有识别到可汇总的主对象。</p>;
  return <DataTable rows={rows.slice(0, 120).map((item, index) => ({ 序号: index + 1, [label]: item.name, 记录数: item.count, 数量: round2(item.quantity), 金额: `¥${money(item.amount)}` }))} headers={["序号", label, "记录数", "数量", "金额"]} />;
}
function AutoResults({ sheet }: { sheet: SheetReport }) {
  if (hasRole(sheet, "product") || hasRole(sheet, "sku")) {
    const rows = sheet.rows.map((row, index) => { const name = text(roleValue(row, sheet, ["product", "sku", "barcode"])); const matched = classifyGlobalProduct(name); return { 序号: index + 1, 商品: name, 自动分类: matched.category, 置信度: `${Math.round(matched.confidence * 100)}%`, 匹配词: matched.matchedKeyword, 数量: roleNumber(row, sheet, ["quantity", "sold"]), 金额: `¥${money(roleNumber(row, sheet, ["amount", "totalAmount", "paidAmount"]))}` }; });
    return <DataTable rows={rows.slice(0, 150)} headers={["序号", "商品", "自动分类", "置信度", "匹配词", "数量", "金额"]} />;
  }
  if (hasRole(sheet, "stockOpening") || hasRole(sheet, "stockEnding")) {
    const rows = sheet.rows.map((row, index) => ({ 序号: index + 1, 对象: objectName(row, sheet), 期初: roleNumber(row, sheet, "stockOpening"), 入库: roleNumber(row, sheet, ["stockIn", "purchase"]), 出库: roleNumber(row, sheet, ["stockOut", "sold"]), 结存: roleNumber(row, sheet, "stockEnding"), 应为: round2(roleNumber(row, sheet, "stockOpening") + roleNumber(row, sheet, ["stockIn", "purchase"]) - roleNumber(row, sheet, ["stockOut", "sold"])) }));
    return <DataTable rows={rows.slice(0, 150)} headers={["序号", "对象", "期初", "入库", "出库", "结存", "应为"]} />;
  }
  if (hasRole(sheet, "employee")) {
    const rows = sheet.rows.map((row, index) => ({ 序号: index + 1, 员工: text(roleValue(row, sheet, "employee")), 部门: text(roleValue(row, sheet, "department")), 岗位: text(roleValue(row, sheet, "position")), 应发: roleNumber(row, sheet, "grossSalary"), 实发: roleNumber(row, sheet, "netSalary"), 扣款: roleNumber(row, sheet, "deduction"), 借支: roleNumber(row, sheet, "advance") }));
    return <DataTable rows={rows.slice(0, 150)} headers={["序号", "员工", "部门", "岗位", "应发", "实发", "扣款", "借支"]} />;
  }
  return <p className="mt-3 text-sm text-slate-600">当前表先生成基础识别、汇总和异常核对结果。</p>;
}

export default function UniversalStructuredReportApp() {
  const [mainTab, setMainTab] = useState<MainTab>("smart");
  const [reportTab, setReportTab] = useState<ReportTab>("overview");
  const [report, setReport] = useState<WorkbookReport | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);
  async function handleFile(file: File) { const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true, cellFormula: true, cellNF: true, cellStyles: true }); setReport(analyzeWorkbook(workbook, file.name)); setActiveSheet(0); setReportTab("overview"); setMainTab("smart"); }
  const sheet = report?.sheets[activeSheet];
  const totalIssues = report?.sheets.reduce((sum, item) => sum + item.issues.length, 0) ?? 0;
  const totalRows = report?.sheets.reduce((sum, item) => sum + item.rows.length, 0) ?? 0;
  return <main className="min-h-screen p-4 md:p-8"><div className="mx-auto max-w-7xl">
    <header className="mb-6 rounded-3xl bg-slate-950 p-6 text-white shadow-sm"><p className="text-sm text-slate-300">Structured Business Report Reconciliation</p><h1 className="mt-1 text-2xl font-bold md:text-4xl">智能表格自动核对</h1><p className="mt-2 max-w-4xl text-sm leading-6 text-slate-300">上传任意业务表后，系统按商品报表的思路生成统一结构：总览、规则、识别预览、汇总报表、核对异常、自动结果和导出。</p></header>
    <nav className="mb-6 flex flex-wrap gap-2">{MAIN_TABS.map(([key, label]) => <button key={key} onClick={() => setMainTab(key)} className={`rounded-xl px-4 py-2 text-sm font-medium ${mainTab === key ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>{label}</button>)}</nav>
    {mainTab === "smart" ? <section className="space-y-6"><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">导入业务表 Excel</h2><p className="mt-1 text-sm text-slate-500">上传后自动生成类似商品报表的完整核对页面，不再输出零散面板。</p></div><label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-slate-950 px-4 py-3 text-sm font-medium text-white">选择 Excel<input className="hidden" type="file" accept=".xlsx,.xls" onChange={(event: ChangeEvent<HTMLInputElement>) => event.target.files?.[0] && handleFile(event.target.files[0])} /></label></div></div>{!report ? <Section title="等待导入表格" desc="支持商品、库存、工资、进货、费用、收款、应收应付、订单、客户供应商、资产等业务表。上传后会生成统一结构的核对结果。"><div className="mt-3 grid gap-3 md:grid-cols-4">{["总览", "核对规则", "识别预览", "汇总报表", "核对异常", "自动结果", "导出"].map((item) => <div key={item} className="rounded-xl bg-slate-50 p-3 text-sm font-medium text-slate-700">{item}</div>)}</div></Section> : null}{report && sheet ? <><div className="grid gap-4 md:grid-cols-4"><Card title="文件" value={report.fileName} desc={report.createdAt} /><Card title="工作表" value={`${report.sheets.length} 个`} /><Card title="数据行" value={`${totalRows} 行`} /><Card title="异常" value={`${totalIssues} 条`} danger={totalIssues > 0} /></div><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-wrap gap-2">{report.sheets.map((item, index) => <button key={item.sheetName} onClick={() => setActiveSheet(index)} className={`rounded-xl px-3 py-2 text-sm ${index === activeSheet ? "bg-slate-950 text-white" : "bg-slate-100 text-slate-700"}`}>{item.sheetName}</button>)}</div><div className="mt-4 flex flex-wrap gap-2">{REPORT_TABS.map(([key, label]) => <button key={key} onClick={() => setReportTab(key)} className={`rounded-xl px-4 py-2 text-sm font-medium ${reportTab === key ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>{label}</button>)}</div></div>{reportTab === "overview" ? <Section title="总览" desc="按统一口径展示表格类型、识别可信度、数据量、规则和异常。"><div className="mt-4 grid gap-4 md:grid-cols-4"><Card title="识别类型" value={sheet.type} desc={`置信度 ${Math.round(sheet.confidence * 100)}%`} /><Card title="表头位置" value={`第 ${sheet.headerRow + 1} 行`} /><Card title="核对规则" value={`${sheet.rules.length} 条`} /><Card title="异常数量" value={`${sheet.issues.length} 条`} danger={sheet.issues.length > 0} /></div></Section> : null}{reportTab === "rules" ? <Section title="核对规则" desc="系统根据表头字段自动生成可执行的基础核对规则。"><div className="mt-3 space-y-2">{sheet.rules.map((rule, index) => <div key={rule} className="rounded-xl bg-slate-50 p-3 text-sm"><b>{index + 1}.</b> {rule}</div>)}</div></Section> : null}{reportTab === "preview" ? <Section title="识别预览" desc="查看每一列被识别成什么字段，方便判断是否识别正确。"><DataTable rows={sheet.columns.map((column) => ({ 列: colName(column.index), 表头: column.header, 识别字段: column.role ? ROLE_LABEL[column.role] : "未识别", 置信度: `${Math.round(column.confidence * 100)}%`, 非空数: column.nonEmpty, 示例: column.examples.join(" / ") }))} headers={["列", "表头", "识别字段", "置信度", "非空数", "示例"]} /></Section> : null}{reportTab === "summary" ? <Section title="汇总报表" desc="按当前表最主要的对象自动汇总，例如商品、员工、供应商、客户、渠道、费用类型。"><SummaryTables sheet={sheet} /></Section> : null}{reportTab === "issues" ? <Section title="核对异常" desc="所有自动核对发现的问题集中显示。">{sheet.issues.length ? <DataTable rows={sheet.issues.map((issue, index) => ({ 序号: index + 1, 级别: issue.severity, 类型: issue.type, 行号: issue.rowNumber ?? "", 对象: issue.object ?? "", 说明: issue.message, 应为: issue.expected ?? "", 实际: issue.actual ?? "" }))} headers={["序号", "级别", "类型", "行号", "对象", "说明", "应为", "实际"]} /> : <p className="mt-3 text-sm text-green-700">当前表未发现明显异常。</p>}</Section> : null}{reportTab === "auto" ? <Section title="自动结果" desc="根据识别类型生成当前业务表最有用的自动结果。"><AutoResults sheet={sheet} /></Section> : null}{reportTab === "export" ? <Section title="导出" desc="导出结构化核对报告，包含总览、识别预览、核对规则和异常明细。"><button onClick={() => exportReport(report)} className="mt-4 rounded-xl bg-green-700 px-4 py-3 text-sm font-medium text-white">导出结构化核对报告</button></Section> : null}</> : null}</section> : null}
    {mainTab === "product" ? <section className="space-y-4"><Section title="商品报表专用核对" desc="这里保留你当前商品报表的专用逻辑：自动库存盘点、交接异常、金额核对、收款统计、写回原表。"><ProductReportApp /></Section></section> : null}
    {mainTab === "guide" ? <Section title="功能说明" desc="智能上传现在不是零散面板，而是给每种表生成统一的核对流程。"><div className="mt-4 grid gap-3 md:grid-cols-3">{["商品表：销售汇总、分类识别、金额核对", "库存表：期初/入库/出库/结存核对", "工资表：应发、实发、扣款、借支核对", "进货表：供应商汇总、金额核对", "费用表：费用类型汇总、负数异常", "收款表：渠道汇总、实收核对", "应收应付：对象差异核对", "订单表：订单重复、状态、金额核对", "客户供应商：对象汇总与异常"].map((item) => <div key={item} className="rounded-xl bg-slate-50 p-3 text-sm text-slate-700">{item}</div>)}</div></Section> : null}
  </div></main>;
}
