"use client";

import { useState, type ChangeEvent, type ReactNode } from "react";
import * as XLSX from "xlsx";
import { classifyGlobalProduct } from "./globalProductCatalog";

type ReportTab = "overview" | "rules" | "preview" | "daily" | "summary" | "handover" | "inventory" | "payment" | "issues" | "auto" | "export";
type ReportKind = "product" | "payroll" | "generic";
type Severity = "error" | "warning" | "info";
type Role = "date" | "product" | "employee" | "department" | "position" | "quantity" | "price" | "amount" | "paidAmount" | "paymentChannel" | "supplier" | "customer" | "orderNo" | "baseSalary" | "expectedDays" | "attendanceDays" | "hourlyRate" | "workHours" | "overtimeHours" | "overtimePay" | "commission" | "bonus" | "allowance" | "deduction" | "advance" | "socialInsurance" | "tax" | "grossSalary" | "netSalary" | "remark";
type PaymentShift = "早班" | "夜班" | "合计" | "未知";

type Issue = { severity: Severity; type: string; sheetName?: string; rowNumber?: number; object?: string; message: string; expected?: string | number; actual?: string | number };
type ColumnInfo = { sheetName: string; index: number; header: string; role?: Role; label: string; confidence: number; examples: string[] };
type LearnedTemplate = { id: string; name: string; kind: ReportKind; createdAt: string; columns: Array<{ header: string; role?: Role; index: number }> };
type ProductAlias = { id: string; alias: string; category: string; createdAt: string };
type ProductRow = { sheetName: string; date: string; rowNumber: number; sourceOrder: number; category: string; name: string; price: number; purchase: number; morningStock: number; morningSold: number; morningAmount: number; nightStock: number; nightSold: number; nightAmount: number; totalAmount: number; endingStock: number; expectedNightStock: number; hasNightStock: boolean };
type ProductSummary = { category: string; name: string; price: number; purchase: number; morningSold: number; nightSold: number; totalSold: number; totalAmount: number; endingStock: number; latestDate: string; sourceOrder: number };
type ProductSheet = { sheetName: string; date: string; tableCount: number; rowCount: number; paymentCount: number };
type PaymentRecord = { date: string; sheetName: string; shift: PaymentShift; channel: string; amount: number };
type ProductReport = { kind: "product"; sheets: ProductSheet[]; rows: ProductRow[]; summary: ProductSummary[]; payments: PaymentRecord[]; issues: Issue[]; columns: ColumnInfo[] };
type PayrollRow = { sheetName: string; rowNumber: number; employee: string; department: string; position: string; baseSalary: number; expectedDays: number; attendanceDays: number; hourlyRate: number; workHours: number; overtimeHours: number; overtimePay: number; commission: number; bonus: number; allowance: number; deduction: number; advance: number; socialInsurance: number; tax: number; grossSalary: number; netSalary: number; remark: string; attendancePay: number; calculatedGross: number; calculatedNet: number; grossDiff: number; netDiff: number };
type PayrollReport = { kind: "payroll"; rows: PayrollRow[]; issues: Issue[]; columns: ColumnInfo[] };
type GenericRow = { sheetName: string; rowNumber: number; values: Record<string, string | number> };
type GenericReport = { kind: "generic"; rows: GenericRow[]; issues: Issue[]; columns: ColumnInfo[]; matchedTemplate?: string };
type SmartReport = { fileName: string; createdAt: string; kind: ReportKind; product?: ProductReport; payroll?: PayrollReport; generic: GenericReport };
type TableRow = Record<string, string | number>;

const TEMPLATE_KEY = "smart_report_templates_v4";
const ALIAS_KEY = "smart_product_aliases_v4";
const ROLE_LABEL: Record<Role, string> = {
  date: "日期", product: "商品", employee: "员工", department: "部门", position: "岗位", quantity: "数量", price: "单价", amount: "金额", paidAmount: "实收金额", paymentChannel: "收款渠道", supplier: "供应商", customer: "客户", orderNo: "单号", baseSalary: "底薪/基本工资", expectedDays: "应出勤", attendanceDays: "实际出勤", hourlyRate: "时薪", workHours: "工时", overtimeHours: "加班小时", overtimePay: "加班费", commission: "提成", bonus: "奖金", allowance: "补贴", deduction: "扣款", advance: "借支", socialInsurance: "社保", tax: "个税", grossSalary: "应发工资", netSalary: "实发工资", remark: "备注"
};
const ROLE_ALIASES: Record<Role, string[]> = {
  date: ["日期", "时间", "营业日", "date", "day"],
  product: ["商品", "商品名称", "售卖商品", "产品", "品名", "货品", "物品"],
  employee: ["员工", "员工姓名", "姓名", "人员", "工号"],
  department: ["部门", "门店", "店铺", "分店", "班组", "组织"],
  position: ["岗位", "职位", "职务", "工种"],
  quantity: ["数量", "件数", "个数", "包数", "瓶数", "销量"],
  price: ["价格", "单价", "售价", "销售价"],
  amount: ["金额", "小计", "销售额", "收入", "支出", "费用", "合计", "总计"],
  paidAmount: ["实收", "实收金额", "到账", "收款", "收款金额", "已收"],
  paymentChannel: ["渠道", "收款渠道", "支付方式", "付款方式", "平台"],
  supplier: ["供应商", "厂家", "供货商"],
  customer: ["客户", "客户名称", "会员", "买家"],
  orderNo: ["单号", "订单号", "流水号", "编号"],
  baseSalary: ["基本工资", "底薪", "月薪", "工资标准", "固定工资"],
  expectedDays: ["应出勤", "应出勤天数", "满勤天数", "标准天数"],
  attendanceDays: ["实际出勤", "出勤", "出勤天数", "上班天数", "考勤"],
  hourlyRate: ["时薪", "小时工资", "每小时"],
  workHours: ["工时", "工作小时", "小时"],
  overtimeHours: ["加班小时", "加班工时"],
  overtimePay: ["加班费", "加班工资"],
  commission: ["提成", "业绩提成", "销售提成"],
  bonus: ["奖金", "满勤", "奖励", "绩效"],
  allowance: ["补贴", "津贴", "餐补", "房补", "交通补贴"],
  deduction: ["扣款", "罚款", "扣除", "缺勤扣款"],
  advance: ["借支", "预支", "借款"],
  socialInsurance: ["社保", "五险", "保险"],
  tax: ["个税", "个人所得税", "税"],
  grossSalary: ["应发", "应发工资", "工资合计", "应付工资"],
  netSalary: ["实发", "实发工资", "到手", "实际发放"],
  remark: ["备注", "说明", "原因"]
};

function text(value: unknown) { return String(value ?? "").trim(); }
function norm(value: unknown) { return text(value).replace(/[\s\-_，,。.;；:：/\\|()（）\[\]【】{}<>《》"'“”‘’]+/g, "").toLowerCase(); }
function isBlank(value: unknown) { return value === null || value === undefined || text(value) === ""; }
function toNumber(value: unknown) { if (typeof value === "number") return Number.isFinite(value) ? value : 0; const cleaned = text(value).replace(/,/g, "").replace(/[￥¥元%\s]/g, ""); if (!cleaned) return 0; const parsed = Number(cleaned); return Number.isFinite(parsed) ? parsed : 0; }
function round2(value: number) { return Math.round((value + Number.EPSILON) * 100) / 100; }
function money(value: number) { return round2(value).toFixed(2); }
function colName(index: number) { return XLSX.utils.encode_col(index); }
function loadList<T>(key: string): T[] { if (typeof window === "undefined") return []; try { const raw = window.localStorage.getItem(key); return raw ? JSON.parse(raw) as T[] : []; } catch { return []; } }
function saveList<T>(key: string, value: T[]) { if (typeof window !== "undefined") window.localStorage.setItem(key, JSON.stringify(value)); }
function rowText(row: unknown[] | undefined) { return (row ?? []).map(text).filter(Boolean).join(" "); }
function isAmountLikeLabel(value: string) { return /^(金额|收款金额|合计|总计|早班|夜班|收款渠道|渠道|售卖金额)$/.test(value) || /^\d+(\.\d+)?$/.test(value); }
function cleanChannelName(value: unknown) { const channel = text(value).replace(/\s/g, ""); return channel && !isAmountLikeLabel(channel) ? channel : ""; }

function classifyWithAliases(name: string, aliases: ProductAlias[]) {
  const value = norm(name);
  const alias = aliases.find((item) => value === norm(item.alias) || (norm(item.alias).length >= 2 && value.includes(norm(item.alias))));
  if (alias) return { category: alias.category, confidence: 0.99, matchedKeyword: alias.alias };
  const match = classifyGlobalProduct(name);
  return { category: match.category === "烟草烟具" ? "香烟(酒店库存)" : match.category, confidence: match.confidence, matchedKeyword: match.matchedKeyword };
}
function inferRole(header: unknown): { role?: Role; label: string; confidence: number } {
  const h = norm(header);
  let best: { role?: Role; label: string; confidence: number } = { label: "未识别", confidence: 0 };
  (Object.entries(ROLE_ALIASES) as Array<[Role, string[]]>).forEach(([role, aliases]) => aliases.forEach((alias) => {
    const a = norm(alias); let score = 0;
    if (h === a) score = 1; else if (h.includes(a)) score = 0.86; else if (a.includes(h) && h.length >= 2) score = 0.55;
    if (score > best.confidence) best = { role, label: ROLE_LABEL[role], confidence: score };
  }));
  return best.confidence >= 0.5 ? best : { label: "未识别", confidence: 0 };
}
function scoreHeaderRow(row: unknown[]) {
  const joined = norm(row.map(text).filter(Boolean).join(" "));
  let score = row.filter((cell) => !isBlank(cell)).length;
  Object.values(ROLE_ALIASES).flat().forEach((alias) => { if (joined.includes(norm(alias))) score += 2; });
  if (/合计|总计|小计/.test(joined)) score -= 5;
  return score;
}
function detectHeaderRow(grid: unknown[][]) { let best = 0; let bestScore = -Infinity; for (let i = 0; i < Math.min(40, grid.length); i += 1) { const score = scoreHeaderRow(grid[i] ?? []); if (score > bestScore) { best = i; bestScore = score; } } return best; }
function buildColumns(sheetName: string, grid: unknown[][], headerRow: number, templates: LearnedTemplate[]) {
  const headers = grid[headerRow] ?? [];
  let columns = headers.map((header, index): ColumnInfo => {
    const guessed = inferRole(header);
    const examples = grid.slice(headerRow + 1, headerRow + 4).map((row) => text(row?.[index])).filter(Boolean);
    return { sheetName, index, header: text(header) || `列${colName(index)}`, role: guessed.role, label: guessed.label, confidence: guessed.confidence, examples };
  });
  let matchedTemplate: string | undefined;
  let bestTemplate: LearnedTemplate | undefined;
  let bestScore = 0;
  templates.forEach((template) => {
    const hits = template.columns.filter((saved) => columns.some((column) => norm(column.header) === norm(saved.header))).length;
    const score = template.columns.length ? hits / template.columns.length : 0;
    if (score > bestScore) { bestScore = score; bestTemplate = template; }
  });
  if (bestTemplate && bestScore >= 0.55) {
    matchedTemplate = bestTemplate.name;
    columns = columns.map((column) => {
      const saved = bestTemplate?.columns.find((item) => norm(item.header) === norm(column.header));
      return saved ? { ...column, role: saved.role, label: saved.role ? ROLE_LABEL[saved.role] : "未识别", confidence: 0.98 } : column;
    });
  }
  return { columns, matchedTemplate };
}
function columnFor(columns: ColumnInfo[], role: Role) { return columns.find((column) => column.role === role); }
function cell(row: unknown[], columns: ColumnInfo[], role: Role) { const column = columnFor(columns, role); return column ? row[column.index] : undefined; }
function str(row: unknown[], columns: ColumnInfo[], role: Role) { return text(cell(row, columns, role)); }
function num(row: unknown[], columns: ColumnInfo[], role: Role) { return toNumber(cell(row, columns, role)); }

function extractDate(sheetName: string, sheetIndex: number) { const day = /^\d{1,2}$/.test(sheetName) ? Number(sheetName) : sheetIndex + 1; const now = new Date(); return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`; }
function findProductGroups(grid: unknown[][]) {
  const groups: Array<{ headerRow: number; nameCol: number; priceCol: number; purchaseCol: number; morningStockCol: number; morningSoldCol: number; morningAmountCol: number; nightStockCol: number; nightSoldCol: number; nightAmountCol: number; totalAmountCol: number }> = [];
  grid.forEach((row, headerRow) => row.forEach((cell, col) => {
    const label = norm(cell); const next = norm(row[col + 1]);
    if ((label === "商品名称" || label === "售卖商品") && (next === "价格" || next === "售价")) groups.push({ headerRow, nameCol: col, priceCol: col + 1, purchaseCol: col + 2, morningStockCol: col + 3, morningSoldCol: col + 4, morningAmountCol: col + 5, nightStockCol: col + 6, nightSoldCol: col + 7, nightAmountCol: col + 8, totalAmountCol: col + 9 });
  }));
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
  const seen = new Set<string>();
  function push(shift: PaymentShift, channel: string, amount: number) {
    if (!channel || !amount) return;
    const key = `${date}_${sheetName}_${shift}_${channel}_${amount}_${records.length}`;
    if (seen.has(key)) return;
    seen.add(key);
    records.push({ date, sheetName, shift, channel, amount: round2(amount) });
  }
  grid.forEach((row, rowIndex) => row.forEach((cell, col) => {
    if (text(cell) !== "收款渠道") return;
    const shift = paymentShiftFromContext(grid, rowIndex);
    const nextRow = grid[rowIndex + 1] ?? [];
    for (let c = col + 1; c < Math.max(row.length, nextRow.length); c += 1) {
      const sameRowChannel = cleanChannelName(row[c]);
      const sameRowAmount = toNumber(row[c + 1]);
      if (sameRowChannel && sameRowAmount) push(shift, sameRowChannel, sameRowAmount);
      const nextRowChannel = cleanChannelName(row[c]);
      const nextRowAmount = toNumber(nextRow[c]);
      if (nextRowChannel && nextRowAmount) push(shift, nextRowChannel, nextRowAmount);
    }
  }));
  return records;
}

function parseProductReport(workbook: XLSX.WorkBook, aliases: ProductAlias[]): ProductReport | undefined {
  const rows: ProductRow[] = [];
  const sheets: ProductSheet[] = [];
  const columns: ColumnInfo[] = [];
  const payments: PaymentRecord[] = [];
  workbook.SheetNames.forEach((sheetName, sheetIndex) => {
    const worksheet = workbook.Sheets[sheetName]; if (!worksheet) return;
    const grid = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, defval: null });
    const groups = findProductGroups(grid);
    const date = extractDate(sheetName, sheetIndex);
    const sheetPayments = parsePaymentRecords(grid, sheetName, date);
    payments.push(...sheetPayments);
    let rowCount = 0;
    groups.forEach((group, groupIndex) => {
      columns.push(...["商品名称", "价格", "进货", "早班库存", "早班售卖", "早班金额", "夜班库存", "夜班售卖", "夜班金额", "总金额"].map((header, offset) => ({ sheetName, index: group.nameCol + offset, header, role: offset === 0 ? "product" as Role : offset === 1 ? "price" as Role : offset === 9 ? "amount" as Role : undefined, label: header, confidence: 1, examples: [] })));
      for (let r = group.headerRow + 1; r < grid.length; r += 1) {
        const line = grid[r] ?? [];
        const name = text(line[group.nameCol]);
        const price = toNumber(line[group.priceCol]);
        if (!name || !price || /合计|总计|小计|备注|收款|金额/.test(name)) continue;
        const purchase = toNumber(line[group.purchaseCol]);
        const morningStock = toNumber(line[group.morningStockCol]);
        const morningSold = toNumber(line[group.morningSoldCol]);
        const morningAmount = toNumber(line[group.morningAmountCol]) || round2(morningSold * price);
        const nightStockRaw = line[group.nightStockCol];
        const nightStock = toNumber(nightStockRaw);
        const nightSold = toNumber(line[group.nightSoldCol]);
        const nightAmount = toNumber(line[group.nightAmountCol]) || round2(nightSold * price);
        const totalAmount = toNumber(line[group.totalAmountCol]) || round2(morningAmount + nightAmount);
        const active = purchase || morningStock || morningSold || nightStock || nightSold || morningAmount || nightAmount || totalAmount;
        if (!active) continue;
        const expectedNightStock = round2(morningStock + purchase - morningSold);
        const endingStock = round2((isBlank(nightStockRaw) ? expectedNightStock : nightStock) - nightSold);
        const matched = classifyWithAliases(name, aliases);
        rows.push({ sheetName, date, rowNumber: r + 1, sourceOrder: sheetIndex * 1_000_000 + groupIndex * 100_000 + r, category: matched.category === "未识别" ? "未分类" : matched.category, name, price, purchase, morningStock, morningSold, morningAmount, nightStock, nightSold, nightAmount, totalAmount, endingStock, expectedNightStock, hasNightStock: !isBlank(nightStockRaw) });
        rowCount += 1;
      }
    });
    if (rowCount > 0 || sheetPayments.length > 0) sheets.push({ sheetName, date, tableCount: groups.length, rowCount, paymentCount: sheetPayments.length });
  });
  if (rows.length < 1) return undefined;
  const issues: Issue[] = [];
  rows.forEach((row) => {
    if (row.hasNightStock && Math.abs(row.nightStock - row.expectedNightStock) > 0.01) issues.push({ severity: "error", type: "早班交夜班", sheetName: row.sheetName, rowNumber: row.rowNumber, object: row.name, message: "夜班库存应等于早班库存 + 进货 - 早班销量。", expected: row.expectedNightStock, actual: row.nightStock });
    if (Math.abs(row.morningAmount - row.morningSold * row.price) > 0.01) issues.push({ severity: "warning", type: "早班金额核对", sheetName: row.sheetName, rowNumber: row.rowNumber, object: row.name, message: "早班金额不等于早班销量 × 价格。", expected: round2(row.morningSold * row.price), actual: row.morningAmount });
    if (Math.abs(row.nightAmount - row.nightSold * row.price) > 0.01) issues.push({ severity: "warning", type: "夜班金额核对", sheetName: row.sheetName, rowNumber: row.rowNumber, object: row.name, message: "夜班金额不等于夜班销量 × 价格。", expected: round2(row.nightSold * row.price), actual: row.nightAmount });
  });
  const map = new Map<string, ProductSummary>();
  rows.forEach((row) => {
    const key = `${row.category}__${row.name}__${row.price}`;
    const item = map.get(key) ?? { category: row.category, name: row.name, price: row.price, purchase: 0, morningSold: 0, nightSold: 0, totalSold: 0, totalAmount: 0, endingStock: 0, latestDate: row.date, sourceOrder: row.sourceOrder };
    item.purchase += row.purchase; item.morningSold += row.morningSold; item.nightSold += row.nightSold; item.totalSold += row.morningSold + row.nightSold; item.totalAmount += row.totalAmount;
    if (row.date >= item.latestDate) { item.latestDate = row.date; item.endingStock = row.endingStock; }
    item.sourceOrder = Math.min(item.sourceOrder, row.sourceOrder);
    map.set(key, item);
  });
  const summary = Array.from(map.values()).map((item) => ({ ...item, purchase: round2(item.purchase), morningSold: round2(item.morningSold), nightSold: round2(item.nightSold), totalSold: round2(item.totalSold), totalAmount: round2(item.totalAmount), endingStock: round2(item.endingStock) })).sort((a, b) => a.category.localeCompare(b.category) || a.sourceOrder - b.sourceOrder);
  return { kind: "product", sheets, rows, summary, payments, issues, columns };
}

function parsePayroll(workbook: XLSX.WorkBook, templates: LearnedTemplate[]): PayrollReport | undefined {
  const rows: PayrollRow[] = [];
  const columns: ColumnInfo[] = [];
  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName]; if (!worksheet) return;
    const grid = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, defval: null });
    const headerRow = detectHeaderRow(grid);
    const built = buildColumns(sheetName, grid, headerRow, templates);
    const sheetColumns = built.columns;
    const payrollHits = sheetColumns.filter((c) => c.role && ["employee", "baseSalary", "attendanceDays", "grossSalary", "netSalary", "deduction", "advance"].includes(c.role)).length;
    if (payrollHits < 2) return;
    columns.push(...sheetColumns);
    for (let r = headerRow + 1; r < grid.length; r += 1) {
      const raw = grid[r] ?? [];
      if (!raw.some((v) => !isBlank(v))) continue;
      const employee = str(raw, sheetColumns, "employee");
      const department = str(raw, sheetColumns, "department");
      const position = str(raw, sheetColumns, "position");
      if (!employee && !department && !position) continue;
      if (/合计|总计|小计/.test(employee)) continue;
      const baseSalary = num(raw, sheetColumns, "baseSalary");
      const expectedDays = num(raw, sheetColumns, "expectedDays");
      const attendanceDays = num(raw, sheetColumns, "attendanceDays");
      const hourlyRate = num(raw, sheetColumns, "hourlyRate");
      const workHours = num(raw, sheetColumns, "workHours");
      const overtimeHours = num(raw, sheetColumns, "overtimeHours");
      const overtimePay = num(raw, sheetColumns, "overtimePay") || round2(overtimeHours * hourlyRate * 1.5);
      const commission = num(raw, sheetColumns, "commission");
      const bonus = num(raw, sheetColumns, "bonus");
      const allowance = num(raw, sheetColumns, "allowance");
      const deduction = num(raw, sheetColumns, "deduction");
      const advance = num(raw, sheetColumns, "advance");
      const socialInsurance = num(raw, sheetColumns, "socialInsurance");
      const tax = num(raw, sheetColumns, "tax");
      const grossSalary = num(raw, sheetColumns, "grossSalary");
      const netSalary = num(raw, sheetColumns, "netSalary");
      const remark = str(raw, sheetColumns, "remark");
      const attendancePay = expectedDays > 0 && attendanceDays > 0 && baseSalary > 0 ? round2(baseSalary / expectedDays * attendanceDays) : workHours > 0 && hourlyRate > 0 ? round2(workHours * hourlyRate) : baseSalary;
      const calculatedGross = round2(attendancePay + overtimePay + commission + bonus + allowance);
      const calculatedNet = round2(calculatedGross - deduction - advance - socialInsurance - tax);
      rows.push({ sheetName, rowNumber: r + 1, employee, department, position, baseSalary, expectedDays, attendanceDays, hourlyRate, workHours, overtimeHours, overtimePay, commission, bonus, allowance, deduction, advance, socialInsurance, tax, grossSalary, netSalary, remark, attendancePay, calculatedGross, calculatedNet, grossDiff: grossSalary ? round2(grossSalary - calculatedGross) : 0, netDiff: netSalary ? round2(netSalary - calculatedNet) : 0 });
    }
  });
  if (rows.length < 1) return undefined;
  const issues: Issue[] = [];
  const seen = new Map<string, number>();
  rows.forEach((row) => {
    if (!row.employee) issues.push({ severity: "error", type: "员工缺失", sheetName: row.sheetName, rowNumber: row.rowNumber, object: row.employee, message: "员工姓名为空。" });
    const key = `${row.employee}__${row.department || ""}`;
    if (row.employee && seen.has(key)) issues.push({ severity: "warning", type: "疑似重复员工", sheetName: row.sheetName, rowNumber: row.rowNumber, object: row.employee, message: `与第 ${seen.get(key)} 行员工重复。` }); else if (row.employee) seen.set(key, row.rowNumber);
    if (row.expectedDays && row.attendanceDays > row.expectedDays) issues.push({ severity: "warning", type: "出勤异常", sheetName: row.sheetName, rowNumber: row.rowNumber, object: row.employee, message: "实际出勤天数大于应出勤天数。", expected: row.expectedDays, actual: row.attendanceDays });
    if (!row.expectedDays && row.attendanceDays > 31) issues.push({ severity: "warning", type: "出勤异常", sheetName: row.sheetName, rowNumber: row.rowNumber, object: row.employee, message: "实际出勤天数超过 31 天。", actual: row.attendanceDays });
    if (row.grossSalary && Math.abs(row.grossDiff) > 0.01) issues.push({ severity: "error", type: "应发工资不一致", sheetName: row.sheetName, rowNumber: row.rowNumber, object: row.employee, message: "应发工资不等于系统计算应发。", expected: row.calculatedGross, actual: row.grossSalary });
    if (row.netSalary && Math.abs(row.netDiff) > 0.01) issues.push({ severity: "error", type: "实发工资不一致", sheetName: row.sheetName, rowNumber: row.rowNumber, object: row.employee, message: "实发工资不等于系统计算实发。", expected: row.calculatedNet, actual: row.netSalary });
    if ((row.deduction || row.advance || row.socialInsurance || row.tax) && !row.remark) issues.push({ severity: "info", type: "扣款无备注", sheetName: row.sheetName, rowNumber: row.rowNumber, object: row.employee, message: "存在扣款/借支/社保/个税，但备注为空。" });
  });
  return { kind: "payroll", rows, issues, columns };
}
function parseGeneric(workbook: XLSX.WorkBook, templates: LearnedTemplate[]): GenericReport {
  const rows: GenericRow[] = [];
  const issues: Issue[] = [];
  const columns: ColumnInfo[] = [];
  let matchedTemplate: string | undefined;
  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName]; if (!worksheet) return;
    const grid = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, defval: null });
    const headerRow = detectHeaderRow(grid);
    const built = buildColumns(sheetName, grid, headerRow, templates);
    if (built.matchedTemplate) matchedTemplate = built.matchedTemplate;
    columns.push(...built.columns);
    const headers = grid[headerRow] ?? [];
    for (let r = headerRow + 1; r < grid.length; r += 1) {
      const raw = grid[r] ?? [];
      if (!raw.some((v) => !isBlank(v))) continue;
      const values: Record<string, string | number> = {};
      headers.forEach((header, index) => { values[text(header) || `列${colName(index)}`] = typeof raw[index] === "number" ? raw[index] as number : text(raw[index]); });
      rows.push({ sheetName, rowNumber: r + 1, values });
    }
  });
  return { kind: "generic", rows, issues, columns, matchedTemplate };
}
function analyzeWorkbook(workbook: XLSX.WorkBook, fileName: string, templates: LearnedTemplate[], aliases: ProductAlias[]): SmartReport { const product = parseProductReport(workbook, aliases); const payroll = parsePayroll(workbook, templates); const generic = parseGeneric(workbook, templates); const kind: ReportKind = product ? "product" : payroll ? "payroll" : "generic"; return { fileName, createdAt: new Date().toLocaleString("zh-CN", { hour12: false }), kind, product, payroll, generic }; }
function activeIssues(report: SmartReport | null) { if (!report) return []; return report.kind === "product" ? report.product?.issues ?? [] : report.kind === "payroll" ? report.payroll?.issues ?? [] : report.generic.issues; }
function activeColumns(report: SmartReport | null) { if (!report) return []; return report.kind === "product" ? report.product?.columns ?? [] : report.kind === "payroll" ? report.payroll?.columns ?? [] : report.generic.columns; }
function activeRowsCount(report: SmartReport | null) { if (!report) return 0; return report.kind === "product" ? report.product?.rows.length ?? 0 : report.kind === "payroll" ? report.payroll?.rows.length ?? 0 : report.generic.rows.length; }
function availableTabs(kind?: ReportKind): Array<[ReportTab, string]> { const base: Array<[ReportTab, string]> = [["overview", "总览"], ["rules", "核对规则"], ["preview", "识别预览"]]; if (kind === "product") return [...base, ["daily", "每日报表"], ["summary", "商品总汇"], ["handover", "交接核对"], ["inventory", "自动库存盘点"], ["payment", "收款统计"], ["issues", "核对异常"], ["auto", "自动结果"], ["export", "导出"]]; if (kind === "payroll") return [...base, ["summary", "工资汇总"], ["issues", "核对异常"], ["auto", "自动结果"], ["export", "导出"]]; return [...base, ["summary", "汇总报表"], ["issues", "核对异常"], ["auto", "自动结果"], ["export", "导出"]]; }
function groupPayroll(rows: PayrollRow[], key: "department" | "position") { const map = new Map<string, { name: string; count: number; gross: number; net: number; issues: number }>(); rows.forEach((row) => { const name = row[key] || "未识别"; const old = map.get(name) ?? { name, count: 0, gross: 0, net: 0, issues: 0 }; old.count += 1; old.gross += row.grossSalary || row.calculatedGross; old.net += row.netSalary || row.calculatedNet; old.issues += Math.abs(row.grossDiff) > 0.01 || Math.abs(row.netDiff) > 0.01 ? 1 : 0; map.set(name, old); }); return Array.from(map.values()).map((item) => ({ ...item, gross: round2(item.gross), net: round2(item.net) })).sort((a, b) => b.net - a.net); }
function paymentBaseRecords(payments: PaymentRecord[]) { const active = payments.filter((item) => item.amount > 0); const nonTotal = active.filter((item) => item.shift !== "合计"); return nonTotal.length ? nonTotal : active; }
function paymentChannelRows(payments: PaymentRecord[]) { const map = new Map<string, { channel: string; amount: number; count: number }>(); paymentBaseRecords(payments).forEach((item) => { const old = map.get(item.channel) ?? { channel: item.channel, amount: 0, count: 0 }; old.amount += item.amount; old.count += 1; map.set(item.channel, old); }); return Array.from(map.values()).map((item, index) => ({ 序号: index + 1, 收款渠道: item.channel, 金额: `¥${money(item.amount)}`, 记录数: item.count })).sort((a, b) => Number(String(b.金额).replace(/[¥,]/g, "")) - Number(String(a.金额).replace(/[¥,]/g, ""))); }
function paymentDailyRows(payments: PaymentRecord[]) { return paymentBaseRecords(payments).map((item, index) => ({ 序号: index + 1, 日期: item.date, Sheet: item.sheetName, 班次: item.shift, 收款渠道: item.channel, 金额: `¥${money(item.amount)}` })); }
function exportSmart(report: SmartReport) { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{ 文件: report.fileName, 类型: report.kind === "product" ? "商品报表" : report.kind === "payroll" ? "工资表" : "通用业务表", 数据行: activeRowsCount(report), 异常数: activeIssues(report).length }]), "总览"); if (report.product) { XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(report.product.summary.map((item, index) => ({ 序号: index + 1, 分类: item.category, 商品: item.name, 价格: item.price, 进货: item.purchase, 早班销量: item.morningSold, 夜班销量: item.nightSold, 总销量: item.totalSold, 销售金额: item.totalAmount, 月末账面应剩: item.endingStock }))), "商品总汇"); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(paymentChannelRows(report.product.payments)), "收款渠道"); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(paymentDailyRows(report.product.payments)), "收款明细"); } if (report.payroll) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(report.payroll.rows.map((row, index) => ({ 序号: index + 1, 员工: row.employee, 部门: row.department, 岗位: row.position, 系统应发: row.calculatedGross, 表格应发: row.grossSalary, 应发差异: row.grossDiff, 系统实发: row.calculatedNet, 表格实发: row.netSalary, 实发差异: row.netDiff }))), "工资明细"); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(activeIssues(report).map((issue, index) => ({ 序号: index + 1, 级别: issue.severity, 类型: issue.type, Sheet: issue.sheetName ?? "", 行号: issue.rowNumber ?? "", 对象: issue.object ?? "", 说明: issue.message, 应为: issue.expected ?? "", 实际: issue.actual ?? "" }))), "核对异常"); XLSX.writeFile(wb, `${report.fileName.replace(/\.xlsx?$/i, "")}-智能核对报告.xlsx`); }

function Card({ title, value, desc, danger = false }: { title: string; value: string; desc?: string; danger?: boolean }) { return <div className={`rounded-2xl p-5 shadow-sm ${danger ? "bg-red-600 text-white" : "border bg-white text-slate-900"}`}><div className={`text-sm ${danger ? "text-red-100" : "text-slate-500"}`}>{title}</div><div className="mt-2 text-2xl font-bold">{value}</div>{desc ? <div className={`mt-1 text-xs ${danger ? "text-red-100" : "text-slate-500"}`}>{desc}</div> : null}</div>; }
function DataTable({ rows, headers, emptyText = "暂无数据。" }: { rows: TableRow[]; headers: string[]; emptyText?: string }) { if (!rows.length) return <p className="mt-3 rounded-xl bg-slate-50 p-4 text-sm text-slate-600">{emptyText}</p>; return <div className="table-scroll mt-3 max-h-96 overflow-auto"><table className="min-w-full text-sm"><thead className="sticky top-0 bg-slate-50 text-left"><tr>{headers.map((header) => <th key={header} className="p-3">{header}</th>)}</tr></thead><tbody>{rows.map((row, index) => <tr key={index} className="border-t">{headers.map((header) => <td key={header} className="p-3">{row[header]}</td>)}</tr>)}</tbody></table></div>; }
function Section({ title, desc, children }: { title: string; desc?: string; children: ReactNode }) { return <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">{title}</h2>{desc ? <p className="mt-1 text-sm text-slate-500">{desc}</p> : null}{children}</section>; }

export default function UniversalSmartReportAppV6() {
  const [reportTab, setReportTab] = useState<ReportTab>("overview");
  const [report, setReport] = useState<SmartReport | null>(null);
  const [templates, setTemplates] = useState<LearnedTemplate[]>(() => loadList<LearnedTemplate>(TEMPLATE_KEY));
  const [aliases, setAliases] = useState<ProductAlias[]>(() => loadList<ProductAlias>(ALIAS_KEY));
  const tabs = availableTabs(report?.kind);
  const issues = activeIssues(report);
  const columns = activeColumns(report);

  async function handleFile(file: File) { const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true, cellFormula: true, cellNF: true, cellStyles: true }); const next = analyzeWorkbook(workbook, file.name, templates, aliases); setReport(next); setReportTab("overview"); }
  function saveTemplate() { if (!report || !columns.length) return alert("请先上传表格。"); const name = prompt("模板名称", `${report.kind}-${report.fileName}`); if (!name) return; const now = new Date().toLocaleString("zh-CN", { hour12: false }); const next = [{ id: `${Date.now()}`, name, kind: report.kind, createdAt: now, columns: columns.map((column) => ({ header: column.header, role: column.role, index: column.index })) }, ...templates]; setTemplates(next); saveList(TEMPLATE_KEY, next); alert("已保存模板。后续同类表格会优先套用。"); }
  function addAlias(aliasText?: string) { const alias = aliasText || prompt("商品别名/关键词"); if (!alias) return; const category = prompt("归入哪个分类？", "食品饮料"); if (!category) return; const now = new Date().toLocaleString("zh-CN", { hour12: false }); const next = [{ id: `${Date.now()}`, alias, category, createdAt: now }, ...aliases.filter((item) => norm(item.alias) !== norm(alias))]; setAliases(next); saveList(ALIAS_KEY, next); alert("已保存商品别名。下次导入会优先识别。"); }

  const paymentRows = report?.product ? paymentChannelRows(report.product.payments) : [];
  const paymentDetailRows = report?.product ? paymentDailyRows(report.product.payments) : [];
  const salesTotal = report?.product?.summary.reduce((s, x) => s + x.totalAmount, 0) ?? 0;
  const paymentTotal = report?.product ? paymentBaseRecords(report.product.payments).reduce((s, x) => s + x.amount, 0) : 0;
  const handoverRows = report?.product?.rows.filter((row) => row.hasNightStock && Math.abs(row.nightStock - row.expectedNightStock) > 0.01).map((row, index) => ({ 序号: index + 1, Sheet: row.sheetName, 商品: row.name, 应为: row.expectedNightStock, 实际: row.nightStock, 差异: round2(row.nightStock - row.expectedNightStock) })) ?? [];

  return <main className="min-h-screen p-4 md:p-8"><div className="mx-auto max-w-7xl space-y-6">
    <div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">导入任意业务表 Excel</h2><p className="mt-1 text-sm text-slate-500">上传后系统自动判断表格类型，并显示该类型该有的功能。</p></div><label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-slate-950 px-4 py-3 text-sm font-medium text-white">选择 Excel<input className="hidden" type="file" accept=".xlsx,.xls" onChange={(event: ChangeEvent<HTMLInputElement>) => event.target.files?.[0] && handleFile(event.target.files[0])} /></label></div></div>
    {!report ? <Section title="等待导入表格" desc="商品表、工资表、进货表、费用表、收款表等都从这里进入。系统会自动选择对应功能。"><div className="mt-3 grid gap-3 md:grid-cols-4">{["总览", "核对规则", "识别预览", "商品/工资/业务汇总", "核对异常", "自动结果", "导出"].map((item) => <div key={item} className="rounded-xl bg-slate-50 p-3 text-sm font-medium text-slate-700">{item}</div>)}</div></Section> : null}
    {report ? <>
      <div className="grid gap-4 md:grid-cols-4"><Card title="文件" value={report.fileName} desc={report.createdAt} /><Card title="识别类型" value={report.kind === "product" ? "商品报表" : report.kind === "payroll" ? "工资表" : "通用业务表"} /><Card title="数据行" value={`${activeRowsCount(report)} 行`} /><Card title="异常" value={`${issues.length} 条`} danger={issues.length > 0} /></div>
      <div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-wrap gap-2">{tabs.map(([key, label]) => <button key={key} onClick={() => setReportTab(key)} className={`rounded-xl px-4 py-2 text-sm font-medium ${reportTab === key ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>{label}</button>)}</div></div>
      {reportTab === "overview" ? <Section title="总览" desc="系统会根据识别类型显示对应功能，而不是让你手动选专用模块。"><div className="mt-4 grid gap-4 md:grid-cols-4"><Card title="类型" value={report.kind === "product" ? "商品报表" : report.kind === "payroll" ? "工资表" : "通用表"} /><Card title="字段数" value={`${columns.length} 个`} /><Card title="模板" value={report.generic.matchedTemplate || "未套用"} /><Card title="异常" value={`${issues.length} 条`} danger={issues.length > 0} /></div></Section> : null}
      {reportTab === "rules" ? <Section title="核对规则" desc="当前表格自动套用的规则。"><div className="mt-3 space-y-2">{(report.kind === "product" ? ["夜班库存 = 早班库存 + 进货 - 早班销量", "月末账面应剩 = 夜班库存 - 夜班销量", "金额 = 数量 × 单价", "收款统计 = 按收款渠道汇总"] : report.kind === "payroll" ? ["出勤工资 = 底薪 ÷ 应出勤 × 实际出勤", "系统应发 = 出勤工资 + 加班费 + 提成 + 奖金 + 补贴", "系统实发 = 系统应发 - 扣款 - 借支 - 社保 - 个税"] : ["字段识别", "重复记录检查", "金额字段汇总"]).map((rule, index) => <div key={rule} className="rounded-xl bg-slate-50 p-3 text-sm"><b>{index + 1}.</b> {rule}</div>)}</div><button onClick={saveTemplate} className="mt-4 rounded-xl bg-slate-950 px-4 py-2 text-sm text-white">保存当前识别为模板</button></Section> : null}
      {reportTab === "preview" ? <Section title="识别预览" desc="查看系统识别了哪些字段。"><DataTable rows={columns.map((column) => ({ Sheet: column.sheetName, 列: colName(column.index), 表头: column.header, 识别字段: column.label, 置信度: `${Math.round(column.confidence * 100)}%`, 示例: column.examples.join(" / ") }))} headers={["Sheet", "列", "表头", "识别字段", "置信度", "示例"]} /></Section> : null}
      {reportTab === "daily" && report.product ? <Section title="每日报表" desc="只显示有有效商品数据或收款记录的 Sheet。"><DataTable rows={report.product.sheets.map((item, index) => ({ 序号: index + 1, Sheet: item.sheetName, 日期: item.date, 商品区: item.tableCount, 商品行: item.rowCount, 收款记录: item.paymentCount }))} headers={["序号", "Sheet", "日期", "商品区", "商品行", "收款记录"]} /></Section> : null}
      {reportTab === "summary" ? <Section title={report.kind === "payroll" ? "工资汇总" : report.kind === "product" ? "商品总汇" : "汇总报表"} desc="根据识别类型自动生成对应汇总。">{report.product ? <DataTable rows={report.product.summary.map((item, index) => ({ 序号: index + 1, 分类: item.category, 商品: item.name, 价格: item.price, 进货: item.purchase, 早班销量: item.morningSold, 夜班销量: item.nightSold, 总销量: item.totalSold, 销售金额: `¥${money(item.totalAmount)}`, 月末账面应剩: item.endingStock }))} headers={["序号", "分类", "商品", "价格", "进货", "早班销量", "夜班销量", "总销量", "销售金额", "月末账面应剩"]} /> : report.payroll ? <DataTable rows={groupPayroll(report.payroll.rows, "department").map((item, index) => ({ 序号: index + 1, 部门: item.name, 人数: item.count, 应发合计: `¥${money(item.gross)}`, 实发合计: `¥${money(item.net)}`, 异常人数: item.issues }))} headers={["序号", "部门", "人数", "应发合计", "实发合计", "异常人数"]} /> : <p className="mt-3 text-sm text-slate-600">通用表已完成结构识别。保存模板后会继续增强该类型汇总。</p>}</Section> : null}
      {reportTab === "handover" && report.product ? <Section title="交接核对" desc="检查早班交夜班库存是否正确。"><DataTable rows={handoverRows} headers={["序号", "Sheet", "商品", "应为", "实际", "差异"]} emptyText="暂无交接异常，库存交接正常。" /></Section> : null}
      {reportTab === "inventory" && report.product ? <Section title="自动库存盘点" desc="自动计算月末账面应剩。"><DataTable rows={report.product.summary.map((item, index) => ({ 序号: index + 1, 分类: item.category, 商品: item.name, 价格: item.price, 总销量: item.totalSold, 销售金额: `¥${money(item.totalAmount)}`, 月末账面应剩: item.endingStock }))} headers={["序号", "分类", "商品", "价格", "总销量", "销售金额", "月末账面应剩"]} /></Section> : null}
      {reportTab === "payment" && report.product ? <Section title="收款统计" desc="按付款/收款渠道统计，并与商品销售金额对照。"><div className="mt-4 grid gap-4 md:grid-cols-4"><Card title="销售金额" value={`¥${money(salesTotal)}`} /><Card title="渠道收款" value={`¥${money(paymentTotal)}`} /><Card title="差异" value={`¥${money(paymentTotal - salesTotal)}`} danger={Math.abs(paymentTotal - salesTotal) > 0.01 && paymentTotal > 0} /><Card title="渠道数" value={`${paymentRows.length} 个`} /></div><div className="mt-5"><h3 className="font-bold text-slate-900">渠道汇总</h3><DataTable rows={paymentRows} headers={["序号", "收款渠道", "金额", "记录数"]} emptyText="没有识别到收款渠道。请确认表格里有“收款渠道”以及对应金额。" /></div><div className="mt-5"><h3 className="font-bold text-slate-900">收款明细</h3><DataTable rows={paymentDetailRows} headers={["序号", "日期", "Sheet", "班次", "收款渠道", "金额"]} emptyText="暂无收款明细。" /></div></Section> : null}
      {reportTab === "issues" ? <Section title="核对异常" desc="所有自动核对发现的问题。">{issues.length ? <DataTable rows={issues.map((issue, index) => ({ 序号: index + 1, 级别: issue.severity, 类型: issue.type, Sheet: issue.sheetName ?? "", 行号: issue.rowNumber ?? "", 对象: issue.object ?? "", 说明: issue.message, 应为: issue.expected ?? "", 实际: issue.actual ?? "" }))} headers={["序号", "级别", "类型", "Sheet", "行号", "对象", "说明", "应为", "实际"]} /> : <p className="mt-3 text-sm text-green-700">当前表未发现明显异常。</p>}</Section> : null}
      {reportTab === "auto" ? <Section title="自动结果" desc="系统根据表格类型生成最终结果。">{report.payroll ? <DataTable rows={report.payroll.rows.map((row, index) => ({ 序号: index + 1, 员工: row.employee, 部门: row.department, 岗位: row.position, 系统应发: row.calculatedGross, 表格应发: row.grossSalary, 应发差异: row.grossDiff, 系统实发: row.calculatedNet, 表格实发: row.netSalary, 实发差异: row.netDiff }))} headers={["序号", "员工", "部门", "岗位", "系统应发", "表格应发", "应发差异", "系统实发", "表格实发", "实发差异"]} /> : report.product ? <><DataTable rows={report.product.summary.map((item, index) => ({ 序号: index + 1, 商品: item.name, 分类: item.category, 月末账面应剩: item.endingStock, 状态: item.category === "未分类" ? "可加入别名库" : "已识别" }))} headers={["序号", "商品", "分类", "月末账面应剩", "状态"]} /><div className="mt-3 flex flex-wrap gap-2">{report.product.summary.filter((item) => item.category === "未分类").slice(0, 20).map((item) => <button key={item.name} onClick={() => addAlias(item.name)} className="rounded-xl border px-3 py-2 text-xs">加入别名：{item.name}</button>)}</div></> : <p className="mt-3 text-sm text-slate-600">通用表已生成基础结构化结果。</p>}</Section> : null}
      {reportTab === "export" ? <Section title="导出" desc="导出当前识别类型对应的核对报告。"><button onClick={() => exportSmart(report)} className="mt-4 rounded-xl bg-green-700 px-4 py-3 text-sm font-medium text-white">导出智能核对报告</button></Section> : null}
    </> : null}
  </div></main>;
}
