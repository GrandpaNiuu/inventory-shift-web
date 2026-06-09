"use client";

import { useMemo, useState, type ChangeEvent, type ReactNode } from "react";
import * as XLSX from "xlsx";
import ProductReportApp from "./InventoryPageConfigurable";
import { classifyGlobalProduct } from "./globalProductCatalog";

type AppTab = "smart" | "product" | "rules";
type Severity = "error" | "warning" | "info";
type TableType = "商品销售表" | "库存表" | "工资表" | "进货采购表" | "费用表" | "收款表" | "应收应付表" | "订单表" | "客户供应商表" | "资产设备表" | "考勤表" | "通用表格";
type Role =
  | "date"
  | "product"
  | "employee"
  | "customer"
  | "supplier"
  | "orderNo"
  | "sku"
  | "barcode"
  | "category"
  | "quantity"
  | "price"
  | "cost"
  | "amount"
  | "totalAmount"
  | "paidAmount"
  | "receivable"
  | "payable"
  | "tax"
  | "discount"
  | "profit"
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
  | "department"
  | "position"
  | "asset"
  | "assetNo"
  | "status"
  | "remark";

type ColumnInfo = { index: number; header: string; role?: Role; confidence: number; numericRatio: number; examples: string[]; nonEmpty: number };
type DataRow = { rowNumber: number; values: Record<string, string | number> };
type Issue = { severity: Severity; type: string; sheetName: string; rowNumber?: number; field?: string; message: string; expected?: string | number; actual?: string | number };
type SheetAnalysis = { sheetName: string; type: TableType; confidence: number; headerRow: number; columns: ColumnInfo[]; rows: DataRow[]; issues: Issue[]; panels: string[] };
type WorkbookAnalysis = { fileName: string; sheets: SheetAnalysis[]; createdAt: string };

type SummaryRow = Record<string, string | number>;

const TABS: Array<[AppTab, string]> = [["smart", "智能上传"], ["product", "商品报表专用"], ["rules", "面板说明"]];

const ROLE_LABELS: Record<Role, string> = {
  date: "日期",
  product: "商品",
  employee: "员工",
  customer: "客户",
  supplier: "供应商",
  orderNo: "订单/单号",
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
  tax: "税费",
  discount: "折扣",
  profit: "利润",
  stockOpening: "期初库存",
  stockIn: "入库",
  stockOut: "出库",
  stockEnding: "结存",
  purchase: "进货/采购",
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
  date: ["日期", "时间", "营业日", "date", "day", "месяц", "дата", "التاريخ"],
  product: ["商品", "商品名称", "产品", "品名", "货品", "物品", "标题", "item", "product", "goods", "товар", "название", "منتج", "اسم المنتج"],
  employee: ["员工", "员工姓名", "姓名", "人员", "工号", "employee", "staff", "worker", "сотрудник", "الموظف"],
  customer: ["客户", "客户名称", "会员", "买家", "customer", "client", "buyer", "клиент", "عميل"],
  supplier: ["供应商", "厂家", "供货商", "采购商", "supplier", "vendor", "поставщик", "مورد"],
  orderNo: ["单号", "订单号", "流水号", "编号", "票号", "order", "order no", "номер", "طلب"],
  sku: ["sku", "货号", "编码", "商品编码", "Артикул", "رمز"],
  barcode: ["条码", "条形码", "barcode", "ean", "upc", "штрих", "باركود"],
  category: ["分类", "类别", "品类", "类型", "category", "type", "категория", "الفئة"],
  quantity: ["数量", "件数", "个数", "包数", "条数", "瓶数", "qty", "quantity", "count", "количество", "كمية"],
  price: ["价格", "单价", "售价", "销售价", "price", "unit price", "цена", "السعر"],
  cost: ["成本", "成本价", "进价", "cost", "purchase price", "себестоимость", "تكلفة"],
  amount: ["金额", "小计", "销售额", "收入", "支出", "费用", "amount", "sum", "сумма", "المبلغ"],
  totalAmount: ["合计", "总计", "总金额", "总价", "合计金额", "total", "итого", "المجموع"],
  paidAmount: ["实收", "实收金额", "到账", "收款金额", "已收", "paid", "received", "оплачено", "مدفوع"],
  receivable: ["应收", "应收金额", "应收款", "receivable", "ar", "к получению", "مستحق"],
  payable: ["应付", "应付金额", "应付款", "payable", "ap", "к оплате", "دائن"],
  tax: ["税", "税费", "税额", "tax", "vat", "налог", "ضريبة"],
  discount: ["折扣", "优惠", "满减", "discount", "скидка", "خصم"],
  profit: ["利润", "毛利", "profit", "margin", "прибыль", "ربح"],
  stockOpening: ["期初", "初始库存", "早班库存", "上月结存", "昨日结存", "opening", "beginning", "начальный", "مخزون أول"],
  stockIn: ["入库", "进货", "采购", "补货", "收入库", "stock in", "inbound", "приход", "وارد"],
  stockOut: ["出库", "销量", "售卖", "销售数量", "消耗", "stock out", "outbound", "расход", "صادر"],
  stockEnding: ["结存", "库存", "剩余", "月末库存", "夜班库存", "库存数量", "ending", "stock", "остаток", "مخزون"],
  purchase: ["进货", "采购", "补货", "purchase", "buy", "закуп", "شراء"],
  sold: ["销量", "售卖", "销售数量", "卖出", "sold", "sales qty", "продано", "مبيعات"],
  basicSalary: ["基本工资", "底薪", "月薪", "工资标准", "base salary", "salary", "оклад", "راتب"],
  attendance: ["出勤", "出勤天数", "天数", "工时", "考勤", "attendance", "days", "hours", "посещ", "حضور"],
  overtime: ["加班", "加班费", "加班工资", "overtime", "сверхур", "إضافي"],
  commission: ["提成", "业绩提成", "销售提成", "commission", "комиссия", "عمولة"],
  bonus: ["奖金", "补贴", "满勤", "津贴", "奖励", "bonus", "allowance", "премия", "بدل"],
  deduction: ["扣款", "罚款", "扣除", "社保", "个税", "deduction", "fine", "налог", "خصم"],
  advance: ["借支", "预支", "借款", "advance", "loan", "аванс", "سلفة"],
  grossSalary: ["应发", "应发工资", "工资合计", "gross", "gross salary", "начислено", "الإجمالي"],
  netSalary: ["实发", "实发工资", "到手", "实际发放", "net", "net salary", "выплачено", "الصافي"],
  feeType: ["费用类型", "费用项目", "支出项目", "科目", "expense type", "fee type", "расход", "مصروف"],
  paymentChannel: ["渠道", "收款渠道", "支付方式", "付款方式", "平台", "channel", "payment method", "канал", "طريقة الدفع"],
  department: ["部门", "门店", "组织", "department", "team", "отдел", "قسم"],
  position: ["岗位", "职位", "职务", "position", "role", "должность", "منصب"],
  asset: ["资产", "设备", "固定资产", "asset", "device", "equipment", "актив", "معدات"],
  assetNo: ["资产编号", "设备编号", "asset no", "serial", "серийный", "رقم الأصل"],
  status: ["状态", "进度", "status", "state", "статус", "حالة"],
  remark: ["备注", "说明", "原因", "note", "remark", "comment", "примеч", "ملاحظة"]
};

function text(value: unknown) { return String(value ?? "").trim(); }
function norm(value: unknown) { return text(value).replace(/[\s\-_，,。.;；:：/\\|()（）\[\]【】{}<>《》"'“”‘’]+/g, "").toLowerCase(); }
function isBlank(value: unknown) { return value === null || value === undefined || text(value) === ""; }
function toNumber(value: unknown) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const cleaned = text(value).replace(/,/g, "").replace(/[￥¥元%\s]/g, "");
  if (!cleaned) return 0;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}
function round2(value: number) { return Math.round((value + Number.EPSILON) * 100) / 100; }
function money(value: number) { return round2(value).toFixed(2); }
function colName(index: number) { return XLSX.utils.encode_col(index); }
function tableKey(value: string) { return value || "未识别"; }

function inferRole(header: string): { role?: Role; confidence: number } {
  const h = norm(header);
  let role: Role | undefined;
  let confidence = 0;
  (Object.entries(ROLE_ALIASES) as Array<[Role, string[]]>).forEach(([candidate, aliases]) => {
    aliases.forEach((alias) => {
      const a = norm(alias);
      let score = 0;
      if (h === a) score = 1;
      else if (h.includes(a)) score = 0.86;
      else if (a.includes(h) && h.length >= 2) score = 0.55;
      if (score > confidence) {
        confidence = score;
        role = candidate;
      }
    });
  });
  return { role: confidence >= 0.5 ? role : undefined, confidence };
}

function scoreHeaderRow(row: unknown[]) {
  const joined = norm(row.map(text).filter(Boolean).join(" "));
  let score = row.filter((cell) => !isBlank(cell)).length;
  Object.values(ROLE_ALIASES).flat().forEach((alias) => {
    if (joined.includes(norm(alias))) score += 2;
  });
  if (/合计|总计|小计/.test(joined)) score -= 4;
  return score;
}
function detectHeaderRow(grid: unknown[][]) {
  let best = 0;
  let bestScore = -Infinity;
  for (let i = 0; i < Math.min(grid.length, 35); i += 1) {
    const score = scoreHeaderRow(grid[i] ?? []);
    if (score > bestScore) {
      best = i;
      bestScore = score;
    }
  }
  return best;
}
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
function buildRows(grid: unknown[][], headerRow: number, headers: string[]) {
  const rows: DataRow[] = [];
  for (let r = headerRow + 1; r < grid.length; r += 1) {
    const raw = grid[r] ?? [];
    if (!raw.some((cell) => !isBlank(cell))) continue;
    const values: Record<string, string | number> = {};
    headers.forEach((header, index) => {
      const cell = raw[index];
      values[header] = typeof cell === "number" ? cell : text(cell);
    });
    rows.push({ rowNumber: r + 1, values });
  }
  return rows;
}
function analyzeColumns(headers: string[], rows: DataRow[]) {
  return headers.map((header, index): ColumnInfo => {
    const values = rows.map((row) => row.values[header]).filter((value) => !isBlank(value));
    const numeric = values.filter((value) => text(value) !== "" && Number.isFinite(toNumber(value))).length;
    const guessed = inferRole(header);
    return { index, header, role: guessed.role, confidence: guessed.confidence, numericRatio: values.length ? numeric / values.length : 0, examples: values.slice(0, 3).map(text), nonEmpty: values.length };
  });
}
function roleColumn(sheet: SheetAnalysis, roles: Role | Role[]) {
  const list = Array.isArray(roles) ? roles : [roles];
  return sheet.columns.find((column) => column.role && list.includes(column.role));
}
function roleValue(row: DataRow, sheet: SheetAnalysis, roles: Role | Role[]) {
  const column = roleColumn(sheet, roles);
  return column ? row.values[column.header] : undefined;
}
function roleNumber(row: DataRow, sheet: SheetAnalysis, roles: Role | Role[]) {
  return toNumber(roleValue(row, sheet, roles));
}
function hasRole(sheet: SheetAnalysis, role: Role) { return sheet.columns.some((column) => column.role === role); }
function scoreRoles(columns: ColumnInfo[], roles: Role[]) { return roles.reduce((sum, role) => sum + (columns.some((column) => column.role === role) ? 1 : 0), 0); }
function detectTableType(columns: ColumnInfo[]): { type: TableType; confidence: number } {
  const candidates: Array<[TableType, number]> = [
    ["工资表", scoreRoles(columns, ["employee", "basicSalary", "attendance", "commission", "deduction", "grossSalary", "netSalary"])],
    ["商品销售表", scoreRoles(columns, ["product", "sku", "quantity", "price", "amount", "sold"])],
    ["库存表", scoreRoles(columns, ["product", "stockOpening", "stockIn", "stockOut", "stockEnding"])],
    ["进货采购表", scoreRoles(columns, ["supplier", "product", "purchase", "quantity", "price", "payable", "amount"])],
    ["费用表", scoreRoles(columns, ["date", "feeType", "amount", "department", "remark"])],
    ["收款表", scoreRoles(columns, ["date", "paymentChannel", "paidAmount", "amount", "orderNo"])],
    ["应收应付表", scoreRoles(columns, ["customer", "supplier", "receivable", "payable", "paidAmount"])],
    ["订单表", scoreRoles(columns, ["orderNo", "product", "customer", "quantity", "amount", "status"])],
    ["客户供应商表", scoreRoles(columns, ["customer", "supplier", "amount", "receivable", "payable"])],
    ["资产设备表", scoreRoles(columns, ["asset", "assetNo", "department", "status", "amount"])],
    ["考勤表", scoreRoles(columns, ["employee", "attendance", "department", "position"])],
    ["通用表格", 1]
  ];
  const [type, score] = candidates.sort((a, b) => b[1] - a[1])[0];
  return { type, confidence: Math.min(0.98, Math.max(0.35, score / 7)) };
}
function availablePanels(sheet: SheetAnalysis) {
  const panels: string[] = ["异常总览", "数据质量核对"];
  if (hasRole(sheet, "product") || hasRole(sheet, "sku") || hasRole(sheet, "sold")) panels.push("商品销售核对", "SKU/条码核对", "商品分类覆盖");
  if (hasRole(sheet, "stockOpening") || hasRole(sheet, "stockIn") || hasRole(sheet, "stockOut") || hasRole(sheet, "stockEnding")) panels.push("库存核对");
  if (hasRole(sheet, "employee") || hasRole(sheet, "netSalary") || hasRole(sheet, "grossSalary")) panels.push("工资核对", "人事考勤核对");
  if (hasRole(sheet, "supplier") || hasRole(sheet, "purchase") || sheet.type === "进货采购表") panels.push("进货采购核对", "供应商汇总");
  if (hasRole(sheet, "feeType") || sheet.type === "费用表") panels.push("费用核对");
  if (hasRole(sheet, "paymentChannel") || hasRole(sheet, "paidAmount")) panels.push("收款核对");
  if (hasRole(sheet, "receivable") || hasRole(sheet, "payable")) panels.push("应收应付核对");
  if (hasRole(sheet, "orderNo") || sheet.type === "订单表") panels.push("订单核对");
  if (hasRole(sheet, "customer") || hasRole(sheet, "supplier")) panels.push("客户供应商分析");
  if (hasRole(sheet, "asset") || hasRole(sheet, "assetNo")) panels.push("资产设备核对");
  return Array.from(new Set(panels));
}
function buildIssues(sheet: SheetAnalysis) {
  const issues: Issue[] = [];
  const seenKey = new Map<string, number>();
  sheet.rows.forEach((row) => {
    const key = text(roleValue(row, sheet, ["orderNo", "sku", "barcode", "product", "employee", "customer", "supplier"]));
    if (!key) issues.push({ severity: "warning", type: "关键字段缺失", sheetName: sheet.sheetName, rowNumber: row.rowNumber, message: "本行缺少可识别的关键字段。" });
    else if (seenKey.has(key)) issues.push({ severity: "warning", type: "疑似重复", sheetName: sheet.sheetName, rowNumber: row.rowNumber, message: `与第 ${seenKey.get(key)} 行关键字段重复：${key}` });
    else seenKey.set(key, row.rowNumber);

    const qty = roleNumber(row, sheet, ["quantity", "sold", "stockOut"]);
    const price = roleNumber(row, sheet, "price");
    const amount = roleNumber(row, sheet, ["amount", "totalAmount"]);
    if (qty && price && amount && Math.abs(round2(qty * price) - amount) > 0.01) issues.push({ severity: "error", type: "金额不一致", sheetName: sheet.sheetName, rowNumber: row.rowNumber, message: "金额不等于数量 × 单价。", expected: round2(qty * price), actual: amount });

    const opening = roleNumber(row, sheet, "stockOpening");
    const stockIn = roleNumber(row, sheet, ["stockIn", "purchase"]);
    const stockOut = roleNumber(row, sheet, ["stockOut", "sold"]);
    const ending = roleNumber(row, sheet, "stockEnding");
    if ((opening || stockIn || stockOut) && ending && Math.abs(round2(opening + stockIn - stockOut) - ending) > 0.01) issues.push({ severity: "error", type: "库存不一致", sheetName: sheet.sheetName, rowNumber: row.rowNumber, message: "结存不等于期初 + 入库 - 出库。", expected: round2(opening + stockIn - stockOut), actual: ending });
    if (ending < 0) issues.push({ severity: "warning", type: "负库存", sheetName: sheet.sheetName, rowNumber: row.rowNumber, message: "结存库存为负数。", actual: ending });

    const gross = roleNumber(row, sheet, "grossSalary");
    const base = roleNumber(row, sheet, "basicSalary");
    const overtime = roleNumber(row, sheet, "overtime");
    const commission = roleNumber(row, sheet, "commission");
    const bonus = roleNumber(row, sheet, "bonus");
    const deduction = roleNumber(row, sheet, "deduction");
    const advance = roleNumber(row, sheet, "advance");
    const net = roleNumber(row, sheet, "netSalary");
    if (gross && base && Math.abs(round2(base + overtime + commission + bonus) - gross) > 0.01) issues.push({ severity: "error", type: "应发工资不一致", sheetName: sheet.sheetName, rowNumber: row.rowNumber, message: "应发工资不等于基本工资 + 加班 + 提成 + 奖金。", expected: round2(base + overtime + commission + bonus), actual: gross });
    if (gross && net && Math.abs(round2(gross - deduction - advance) - net) > 0.01) issues.push({ severity: "error", type: "实发工资不一致", sheetName: sheet.sheetName, rowNumber: row.rowNumber, message: "实发工资不等于应发工资 - 扣款 - 借支。", expected: round2(gross - deduction - advance), actual: net });

    ["amount", "totalAmount", "paidAmount", "receivable", "payable", "netSalary", "grossSalary"].forEach((role) => {
      const value = roleNumber(row, sheet, role as Role);
      if (value < 0) issues.push({ severity: "warning", type: "负数金额", sheetName: sheet.sheetName, rowNumber: row.rowNumber, field: ROLE_LABELS[role as Role], message: `${ROLE_LABELS[role as Role]} 为负数。`, actual: value });
    });
  });
  return issues;
}
function analyzeWorkbook(workbook: XLSX.WorkBook, fileName: string): WorkbookAnalysis {
  const sheets = workbook.SheetNames.map((sheetName): SheetAnalysis | null => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return null;
    const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null });
    const headerRow = detectHeaderRow(grid);
    const used = new Set<string>();
    const headers = (grid[headerRow] ?? []).map((cell, index) => uniqueHeader(cell, index, used));
    const rows = buildRows(grid, headerRow, headers);
    const columns = analyzeColumns(headers, rows);
    const detected = detectTableType(columns);
    const base: SheetAnalysis = { sheetName, type: detected.type, confidence: detected.confidence, headerRow, columns, rows, issues: [], panels: [] };
    const withPanels = { ...base, panels: availablePanels(base) };
    return { ...withPanels, issues: buildIssues(withPanels) };
  }).filter((sheet): sheet is SheetAnalysis => Boolean(sheet));
  return { fileName, sheets, createdAt: new Date().toLocaleString("zh-CN", { hour12: false }) };
}
function groupSum(sheet: SheetAnalysis, groupRoles: Role | Role[], sumRoles: Role | Role[]) {
  const map = new Map<string, { name: string; count: number; quantity: number; amount: number }>();
  sheet.rows.forEach((row) => {
    const name = tableKey(text(roleValue(row, sheet, groupRoles)));
    const old = map.get(name) ?? { name, count: 0, quantity: 0, amount: 0 };
    old.count += 1;
    old.quantity += roleNumber(row, sheet, ["quantity", "sold", "stockOut"]);
    old.amount += roleNumber(row, sheet, sumRoles);
    map.set(name, old);
  });
  return Array.from(map.values()).sort((a, b) => b.amount - a.amount || b.quantity - a.quantity || b.count - a.count);
}
function exportReport(analysis: WorkbookAnalysis) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(analysis.sheets.map((sheet) => ({ Sheet: sheet.sheetName, 类型: sheet.type, 置信度: Math.round(sheet.confidence * 100) + "%", 表头行: sheet.headerRow + 1, 数据行: sheet.rows.length, 面板数: sheet.panels.length, 异常数: sheet.issues.length }))), "业务总览");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(analysis.sheets.flatMap((sheet) => sheet.columns.map((column) => ({ Sheet: sheet.sheetName, 列: colName(column.index), 表头: column.header, 字段: column.role ? ROLE_LABELS[column.role] : "未识别", 置信度: Math.round(column.confidence * 100) + "%", 示例: column.examples.join(" / ") })))), "字段识别");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(analysis.sheets.flatMap((sheet) => sheet.issues.map((issue, index) => ({ Sheet: sheet.sheetName, 序号: index + 1, 级别: issue.severity, 类型: issue.type, 行号: issue.rowNumber ?? "", 字段: issue.field ?? "", 说明: issue.message, 应为: issue.expected ?? "", 实际: issue.actual ?? "" })))), "异常明细");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(analysis.sheets.flatMap((sheet) => sheet.panels.map((panel) => ({ Sheet: sheet.sheetName, 自动面板: panel })))), "自动面板");
  XLSX.writeFile(wb, `${analysis.fileName.replace(/\.xlsx?$/i, "")}-自动业务核对报告.xlsx`);
}
function Card({ title, value, desc, danger = false }: { title: string; value: string; desc?: string; danger?: boolean }) {
  return <div className={`rounded-2xl p-5 shadow-sm ${danger ? "bg-red-600 text-white" : "border bg-white text-slate-900"}`}><div className={`text-sm ${danger ? "text-red-100" : "text-slate-500"}`}>{title}</div><div className="mt-2 text-2xl font-bold">{value}</div>{desc ? <div className={`mt-1 text-xs ${danger ? "text-red-100" : "text-slate-500"}`}>{desc}</div> : null}</div>;
}
function Badge({ children, tone = "gray" }: { children: ReactNode; tone?: "gray" | "green" | "red" | "yellow" }) {
  const cls = { gray: "bg-slate-100 text-slate-700", green: "bg-green-100 text-green-700", red: "bg-red-100 text-red-700", yellow: "bg-yellow-100 text-yellow-800" }[tone];
  return <span className={`rounded-full px-2 py-1 text-xs font-medium ${cls}`}>{children}</span>;
}
function MiniTable({ rows, headers }: { rows: SummaryRow[]; headers: string[] }) {
  return <div className="table-scroll mt-3 max-h-80 overflow-auto"><table className="min-w-full text-sm"><thead className="sticky top-0 bg-slate-50 text-left"><tr>{headers.map((header) => <th key={header} className="p-3">{header}</th>)}</tr></thead><tbody>{rows.map((row, index) => <tr key={index} className="border-t">{headers.map((header) => <td key={header} className="p-3">{row[header]}</td>)}</tr>)}</tbody></table></div>;
}
function Panel({ title, desc, children }: { title: string; desc?: string; children: ReactNode }) {
  return <section className="rounded-2xl border bg-white p-5 shadow-sm"><h3 className="font-bold">{title}</h3>{desc ? <p className="mt-1 text-sm text-slate-500">{desc}</p> : null}{children}</section>;
}
function BusinessPanels({ sheet }: { sheet: SheetAnalysis }) {
  const productSummary = groupSum(sheet, ["product", "sku", "barcode"], ["amount", "totalAmount"]);
  const supplierSummary = groupSum(sheet, "supplier", ["amount", "payable", "totalAmount"]);
  const customerSummary = groupSum(sheet, "customer", ["amount", "receivable", "paidAmount"]);
  const channelSummary = groupSum(sheet, "paymentChannel", ["paidAmount", "amount", "totalAmount"]);
  const feeSummary = groupSum(sheet, "feeType", "amount");
  const employeeSummary = groupSum(sheet, "employee", ["netSalary", "grossSalary", "amount"]);
  const categorySummary = groupSum(sheet, "category", ["amount", "totalAmount"]);
  const productClassRows = sheet.rows.filter((row) => text(roleValue(row, sheet, ["product", "sku", "barcode"]))).map((row, index) => { const name = text(roleValue(row, sheet, ["product", "sku", "barcode"])); const classified = classifyGlobalProduct(name); return { 序号: index + 1, 商品: name, 自动分类: classified.category, 置信度: `${Math.round(classified.confidence * 100)}%`, 匹配词: classified.matchedKeyword }; });
  return <div className="space-y-5">
    <Panel title="自动生成的业务核对面板" desc="系统会根据字段识别结果自动打开相关面板，不再只停留在表格预览。"><div className="mt-3 flex flex-wrap gap-2">{sheet.panels.map((panel) => <Badge key={panel} tone="green">{panel}</Badge>)}</div></Panel>
    <Panel title="异常总览" desc="汇总所有自动核对发现的问题。">{sheet.issues.length ? <MiniTable rows={sheet.issues.slice(0, 120).map((issue, index) => ({ 序号: index + 1, 级别: issue.severity, 类型: issue.type, 行号: issue.rowNumber ?? "", 说明: issue.message, 应为: issue.expected ?? "", 实际: issue.actual ?? "" }))} headers={["序号", "级别", "类型", "行号", "说明", "应为", "实际"]} /> : <p className="mt-3 text-sm text-green-700">当前 Sheet 暂未发现明显异常。</p>}</Panel>
    {(hasRole(sheet, "product") || hasRole(sheet, "sku") || hasRole(sheet, "sold")) ? <Panel title="商品销售核对面板" desc="按商品/SKU 汇总数量和金额，并检查金额=数量×单价。"><MiniTable rows={productSummary.slice(0, 80).map((item, index) => ({ 序号: index + 1, 商品: item.name, 记录数: item.count, 数量: round2(item.quantity), 金额: `¥${money(item.amount)}` }))} headers={["序号", "商品", "记录数", "数量", "金额"]} /></Panel> : null}
    {(hasRole(sheet, "product") || hasRole(sheet, "sku")) ? <Panel title="商品分类覆盖面板" desc="按全球商品词库识别各种商品，不局限于香烟和饮料。"><MiniTable rows={productClassRows.slice(0, 100)} headers={["序号", "商品", "自动分类", "置信度", "匹配词"]} /></Panel> : null}
    {(hasRole(sheet, "stockOpening") || hasRole(sheet, "stockIn") || hasRole(sheet, "stockOut") || hasRole(sheet, "stockEnding")) ? <Panel title="库存核对面板" desc="检查结存=期初+入库-出库，并提示负库存。"><MiniTable rows={sheet.rows.slice(0, 100).map((row, index) => ({ 序号: index + 1, 名称: text(roleValue(row, sheet, ["product", "sku"])), 期初: roleNumber(row, sheet, "stockOpening"), 入库: roleNumber(row, sheet, ["stockIn", "purchase"]), 出库: roleNumber(row, sheet, ["stockOut", "sold"]), 结存: roleNumber(row, sheet, "stockEnding"), 应为: round2(roleNumber(row, sheet, "stockOpening") + roleNumber(row, sheet, ["stockIn", "purchase"]) - roleNumber(row, sheet, ["stockOut", "sold"])) }))} headers={["序号", "名称", "期初", "入库", "出库", "结存", "应为"]} /></Panel> : null}
    {(hasRole(sheet, "employee") || hasRole(sheet, "netSalary") || hasRole(sheet, "grossSalary")) ? <Panel title="工资核对面板" desc="核对应发、实发、扣款、提成、奖金、借支。"><MiniTable rows={employeeSummary.slice(0, 100).map((item, index) => ({ 序号: index + 1, 员工: item.name, 记录数: item.count, 金额: `¥${money(item.amount)}` }))} headers={["序号", "员工", "记录数", "金额"]} /></Panel> : null}
    {(hasRole(sheet, "employee") || hasRole(sheet, "attendance") || hasRole(sheet, "department")) ? <Panel title="人事考勤核对面板" desc="按员工/部门识别出勤、岗位、部门字段，辅助工资和排班核对。"><MiniTable rows={sheet.rows.slice(0, 100).map((row, index) => ({ 序号: index + 1, 员工: text(roleValue(row, sheet, "employee")), 部门: text(roleValue(row, sheet, "department")), 岗位: text(roleValue(row, sheet, "position")), 出勤: roleNumber(row, sheet, "attendance") }))} headers={["序号", "员工", "部门", "岗位", "出勤"]} /></Panel> : null}
    {(hasRole(sheet, "supplier") || hasRole(sheet, "purchase")) ? <Panel title="进货采购核对面板" desc="按供应商/商品汇总采购数量和金额，检查数量×单价。"><MiniTable rows={supplierSummary.slice(0, 80).map((item, index) => ({ 序号: index + 1, 供应商: item.name, 记录数: item.count, 数量: round2(item.quantity), 金额: `¥${money(item.amount)}` }))} headers={["序号", "供应商", "记录数", "数量", "金额"]} /></Panel> : null}
    {hasRole(sheet, "feeType") ? <Panel title="费用核对面板" desc="按费用类型汇总支出，并检查负数和异常金额。"><MiniTable rows={feeSummary.slice(0, 80).map((item, index) => ({ 序号: index + 1, 费用类型: item.name, 笔数: item.count, 金额: `¥${money(item.amount)}` }))} headers={["序号", "费用类型", "笔数", "金额"]} /></Panel> : null}
    {(hasRole(sheet, "paymentChannel") || hasRole(sheet, "paidAmount")) ? <Panel title="收款核对面板" desc="按收款渠道汇总金额，适合微信、支付宝、现金、平台渠道对账。"><MiniTable rows={channelSummary.slice(0, 80).map((item, index) => ({ 序号: index + 1, 渠道: item.name, 笔数: item.count, 金额: `¥${money(item.amount)}` }))} headers={["序号", "渠道", "笔数", "金额"]} /></Panel> : null}
    {(hasRole(sheet, "receivable") || hasRole(sheet, "payable")) ? <Panel title="应收应付核对面板" desc="检查应收、应付、实收之间的差异。"><MiniTable rows={sheet.rows.slice(0, 100).map((row, index) => ({ 序号: index + 1, 对象: text(roleValue(row, sheet, ["customer", "supplier"])), 应收: roleNumber(row, sheet, "receivable"), 应付: roleNumber(row, sheet, "payable"), 实收: roleNumber(row, sheet, "paidAmount") }))} headers={["序号", "对象", "应收", "应付", "实收"]} /></Panel> : null}
    {(hasRole(sheet, "orderNo") || hasRole(sheet, "status")) ? <Panel title="订单核对面板" desc="检查订单号、状态、客户、金额和重复单号。"><MiniTable rows={sheet.rows.slice(0, 100).map((row, index) => ({ 序号: index + 1, 单号: text(roleValue(row, sheet, "orderNo")), 客户: text(roleValue(row, sheet, "customer")), 状态: text(roleValue(row, sheet, "status")), 金额: `¥${money(roleNumber(row, sheet, ["amount", "totalAmount", "paidAmount"]))}` }))} headers={["序号", "单号", "客户", "状态", "金额"]} /></Panel> : null}
    {(hasRole(sheet, "customer") || hasRole(sheet, "supplier")) ? <Panel title="客户供应商分析面板" desc="按客户或供应商汇总交易笔数和金额。"><MiniTable rows={[...customerSummary, ...supplierSummary].slice(0, 100).map((item, index) => ({ 序号: index + 1, 对象: item.name, 记录数: item.count, 金额: `¥${money(item.amount)}` }))} headers={["序号", "对象", "记录数", "金额"]} /></Panel> : null}
    {(hasRole(sheet, "asset") || hasRole(sheet, "assetNo")) ? <Panel title="资产设备核对面板" desc="识别资产编号、设备名称、部门和状态。"><MiniTable rows={sheet.rows.slice(0, 100).map((row, index) => ({ 序号: index + 1, 资产: text(roleValue(row, sheet, "asset")), 编号: text(roleValue(row, sheet, "assetNo")), 部门: text(roleValue(row, sheet, "department")), 状态: text(roleValue(row, sheet, "status")) }))} headers={["序号", "资产", "编号", "部门", "状态"]} /></Panel> : null}
    {hasRole(sheet, "category") ? <Panel title="分类汇总面板" desc="按表格自带分类字段汇总数量和金额。"><MiniTable rows={categorySummary.slice(0, 80).map((item, index) => ({ 序号: index + 1, 分类: item.name, 记录数: item.count, 数量: round2(item.quantity), 金额: `¥${money(item.amount)}` }))} headers={["序号", "分类", "记录数", "数量", "金额"]} /></Panel> : null}
  </div>;
}

export default function UniversalBusinessReportApp() {
  const [tab, setTab] = useState<AppTab>("smart");
  const [analysis, setAnalysis] = useState<WorkbookAnalysis | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);

  async function handleFile(file: File) {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true, cellFormula: true, cellNF: true, cellStyles: true });
    setAnalysis(analyzeWorkbook(workbook, file.name));
    setActiveSheet(0);
    setTab("smart");
  }

  const currentSheet = analysis?.sheets[activeSheet];
  const totalIssues = analysis?.sheets.reduce((sum, sheet) => sum + sheet.issues.length, 0) ?? 0;
  const totalPanels = analysis?.sheets.reduce((sum, sheet) => sum + sheet.panels.length, 0) ?? 0;

  return <main className="min-h-screen p-4 md:p-8"><div className="mx-auto max-w-7xl">
    <header className="mb-6 rounded-3xl bg-slate-950 p-6 text-white shadow-sm"><p className="text-sm text-slate-300">Automatic Business Reconciliation Panels</p><h1 className="mt-1 text-2xl font-bold md:text-4xl">智能上传与自动业务核对面板</h1><p className="mt-2 max-w-4xl text-sm leading-6 text-slate-300">上传任意 Excel 后，系统不只识别字段，还会自动生成商品、库存、工资、进货、费用、收款、应收应付、订单、客户供应商等业务核对面板。</p></header>
    <nav className="mb-6 flex flex-wrap gap-2">{TABS.map(([key, label]) => <button key={key} onClick={() => setTab(key)} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === key ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>{label}</button>)}</nav>

    {tab === "smart" ? <section className="space-y-6">
      <div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">智能上传</h2><p className="mt-1 text-sm text-slate-500">上传任意 Excel，自动生成广泛业务核对面板，不再只是查看表格结构。</p></div><label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-slate-950 px-4 py-3 text-sm font-medium text-white">选择 Excel<input className="hidden" type="file" accept=".xlsx,.xls" onChange={(event: ChangeEvent<HTMLInputElement>) => event.target.files?.[0] && handleFile(event.target.files[0])} /></label></div></div>
      {!analysis ? <div className="rounded-2xl border bg-white p-6 shadow-sm"><h2 className="text-lg font-bold">等待上传表格</h2><p className="mt-2 text-sm text-slate-600">系统会根据字段自动打开相关面板：商品销售、库存、工资、人事考勤、进货采购、费用、收款、应收应付、订单、客户供应商、资产设备、分类汇总和异常总览。</p></div> : null}
      {analysis ? <><div className="grid gap-4 md:grid-cols-4"><Card title="文件" value={analysis.fileName} desc={analysis.createdAt} /><Card title="工作表" value={`${analysis.sheets.length} 个`} /><Card title="自动面板" value={`${totalPanels} 个`} /><Card title="异常" value={`${totalIssues} 条`} danger={totalIssues > 0} /></div><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-wrap items-center justify-between gap-3"><div className="flex flex-wrap gap-2">{analysis.sheets.map((sheet, index) => <button key={sheet.sheetName} onClick={() => setActiveSheet(index)} className={`rounded-xl px-3 py-2 text-sm ${index === activeSheet ? "bg-slate-950 text-white" : "bg-slate-100 text-slate-700"}`}>{sheet.sheetName}</button>)}</div><button onClick={() => exportReport(analysis)} className="rounded-xl bg-green-700 px-4 py-2 text-sm text-white">导出自动业务核对报告</button></div></div></> : null}
      {currentSheet ? <div className="space-y-5"><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">{currentSheet.sheetName}</h2><p className="mt-1 text-sm text-slate-500">识别类型：<b>{currentSheet.type}</b> / 置信度 {Math.round(currentSheet.confidence * 100)}% / 表头第 {currentSheet.headerRow + 1} 行 / 数据 {currentSheet.rows.length} 行 / 自动面板 {currentSheet.panels.length} 个</p></div>{currentSheet.type === "商品销售表" || currentSheet.type === "库存表" ? <button onClick={() => setTab("product")} className="rounded-xl bg-slate-950 px-4 py-2 text-sm text-white">进入商品报表专用</button> : null}</div><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["列", "表头", "识别字段", "置信度", "示例"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{currentSheet.columns.map((column) => <tr key={column.index} className="border-t"><td className="p-3">{colName(column.index)}</td><td className="p-3">{column.header}</td><td className="p-3">{column.role ? <Badge tone="green">{ROLE_LABELS[column.role]}</Badge> : <Badge>未识别</Badge>}</td><td className="p-3">{Math.round(column.confidence * 100)}%</td><td className="p-3 text-slate-500">{column.examples.join(" / ")}</td></tr>)}</tbody></table></div></div><BusinessPanels sheet={currentSheet} /></div> : null}
    </section> : null}

    {tab === "product" ? <section className="space-y-4"><div className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">商品报表专用核对</h2><p className="mt-1 text-sm text-slate-500">保留你当前商品报表的专用规则：自动库存盘点、交接异常、金额核对、收款统计、写回原表。</p></div><ProductReportApp /></section> : null}

    {tab === "rules" ? <section className="rounded-2xl border bg-white p-6 shadow-sm"><h2 className="text-lg font-bold">自动业务面板覆盖范围</h2><div className="mt-4 grid gap-3 md:grid-cols-3">{["商品销售核对", "商品分类覆盖", "库存核对", "工资核对", "人事考勤核对", "进货采购核对", "供应商汇总", "费用核对", "收款核对", "应收应付核对", "订单核对", "客户供应商分析", "SKU/条码核对", "资产设备核对", "分类汇总", "数据质量核对", "异常总览"].map((item) => <div key={item} className="rounded-xl bg-slate-50 p-3 text-sm font-medium text-slate-700">{item}</div>)}</div><p className="mt-4 text-sm leading-6 text-slate-600">这不是最终上限。后续还能继续加平台类目核对、利润核对、税费核对、物流核对、员工绩效核对、门店对比、月报自动生成等面板。</p></section> : null}
  </div></main>;
}
