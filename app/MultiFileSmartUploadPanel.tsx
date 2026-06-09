"use client";

import { useState, type ChangeEvent } from "react";
import * as XLSX from "xlsx";

type BatchKind = "product" | "payroll" | "generic";
type PaymentShift = "早班" | "夜班" | "合计" | "未知";
type BatchIssue = { fileName: string; storeName: string; type: string; message: string };
type PaymentRecord = { fileName: string; storeName: string; date: string; sheetName: string; shift: PaymentShift; channel: string; amount: number };
type BatchResult = { fileName: string; storeName: string; kind: BatchKind; rowCount: number; primaryAmount: number; secondaryAmount: number; diffAmount: number; amount: number; issueCount: number; channelCount: number; issues: BatchIssue[]; payments: PaymentRecord[] };
type Role = "product" | "price" | "amount" | "paidAmount" | "employee" | "department" | "position" | "baseSalary" | "expectedDays" | "attendanceDays" | "hourlyRate" | "workHours" | "overtimeHours" | "overtimePay" | "commission" | "bonus" | "allowance" | "deduction" | "advance" | "tax" | "socialInsurance" | "grossSalary" | "netSalary" | "remark";

const ROLE_ALIASES: Record<Role, string[]> = {
  product: ["商品", "商品名称", "售卖商品", "产品", "品名", "货品"],
  price: ["价格", "单价", "售价", "销售价"],
  amount: ["金额", "小计", "销售额", "收入", "合计"],
  paidAmount: ["实收", "实收金额", "到账", "收款", "收款金额", "已收"],
  employee: ["员工", "员工姓名", "姓名", "人员", "工号", "employee", "staff", "worker"],
  department: ["部门", "门店", "店铺", "分店", "班组", "组织", "department", "team"],
  position: ["岗位", "职位", "职务", "工种", "position", "role", "job"],
  baseSalary: ["基本工资", "底薪", "月薪", "工资标准", "固定工资", "base", "salary"],
  expectedDays: ["应出勤", "应出勤天数", "满勤天数", "标准天数", "应上班", "expecteddays"],
  attendanceDays: ["实际出勤", "出勤", "出勤天数", "上班天数", "考勤", "attendance", "days"],
  hourlyRate: ["时薪", "小时工资", "每小时", "hourly", "hourlyrate"],
  workHours: ["工时", "工作小时", "小时", "workhours", "hours"],
  overtimeHours: ["加班小时", "加班工时", "overtimehours"],
  overtimePay: ["加班费", "加班工资", "overtime", "overtimepay"],
  commission: ["提成", "业绩提成", "销售提成", "commission"],
  bonus: ["奖金", "满勤", "奖励", "绩效", "bonus", "performance"],
  allowance: ["补贴", "津贴", "餐补", "房补", "交通补贴", "allowance", "subsidy"],
  deduction: ["扣款", "罚款", "扣除", "缺勤扣款", "deduction", "fine"],
  advance: ["借支", "预支", "借款", "advance", "loan"],
  tax: ["个税", "个人所得税", "税", "tax"],
  socialInsurance: ["社保", "五险", "保险", "social", "insurance"],
  grossSalary: ["应发", "应发工资", "工资合计", "应付工资", "gross", "grosssalary"],
  netSalary: ["实发", "实发工资", "到手", "实际发放", "net", "netsalary"],
  remark: ["备注", "说明", "原因", "note", "remark", "comment"]
};

function text(value: unknown) { return String(value ?? "").trim(); }
function norm(value: unknown) { return text(value).replace(/[\s\-_，,。.;；:：/\\|()（）\[\]【】{}<>《》"'“”‘’]+/g, "").toLowerCase(); }
function toNumber(value: unknown) { if (typeof value === "number") return Number.isFinite(value) ? value : 0; const parsed = Number(text(value).replace(/,/g, "").replace(/[￥¥元%\s]/g, "")); return Number.isFinite(parsed) ? parsed : 0; }
function round2(value: number) { return Math.round((value + Number.EPSILON) * 100) / 100; }
function money(value: number) { return round2(value).toFixed(2); }
function rowText(row: unknown[] | undefined) { return (row ?? []).map(text).filter(Boolean).join(" "); }
function storeFromFileName(fileName: string) { return fileName.replace(/\.(xlsx|xls)$/i, "").replace(/\d{4}[-年.]?\d{1,2}[-月.]?\d{0,2}[日]?/g, "").replace(/报表|日报|月报|工资表|库存表|核对表/g, "").replace(/[-_ ]+/g, " ").trim() || fileName.replace(/\.(xlsx|xls)$/i, ""); }
function isAmountLikeLabel(value: string) { return /^(金额|收款金额|合计|总计|早班|夜班|收款渠道|渠道|售卖金额)$/.test(value) || /^\d+(\.\d+)?$/.test(value); }
function cleanChannelName(value: unknown) { const channel = text(value).replace(/\s/g, ""); return channel && !isAmountLikeLabel(channel) ? channel : ""; }

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

function parsePaymentRecords(grid: unknown[][], sheetName: string, date: string, fileName: string, storeName: string) {
  const records: PaymentRecord[] = [];
  grid.forEach((row, rowIndex) => row.forEach((cell, col) => {
    if (text(cell) !== "收款渠道") return;
    const shift = paymentShiftFromContext(grid, rowIndex);
    for (let c = col + 1; c < row.length - 1; c += 2) {
      const channel = cleanChannelName(row[c]);
      const amount = round2(toNumber(row[c + 1]));
      if (channel && amount) records.push({ fileName, storeName, date, sheetName, shift, channel, amount });
    }
  }));
  return records;
}

function paymentSourceForReconcile(payments: PaymentRecord[]) {
  const byFileDate = new Map<string, PaymentRecord[]>();
  payments.forEach((record) => byFileDate.set(`${record.fileName}__${record.date}`, [...(byFileDate.get(`${record.fileName}__${record.date}`) ?? []), record]));
  return Array.from(byFileDate.values()).flatMap((records) => {
    const totalRecords = records.filter((record) => record.shift === "合计");
    return totalRecords.length ? totalRecords : records.filter((record) => record.shift !== "合计");
  });
}

function channelSummary(payments: PaymentRecord[]) {
  const map = new Map<string, { channel: string; amount: number; count: number }>();
  paymentSourceForReconcile(payments).forEach((record) => {
    const old = map.get(record.channel) ?? { channel: record.channel, amount: 0, count: 0 };
    old.amount += record.amount;
    old.count += 1;
    map.set(record.channel, old);
  });
  return Array.from(map.values()).map((item) => ({ ...item, amount: round2(item.amount) })).sort((a, b) => b.amount - a.amount);
}

function scoreHeaderRow(row: unknown[]) { const joined = norm(row.map(text).filter(Boolean).join(" ")); let score = row.filter((cell) => text(cell)).length; Object.values(ROLE_ALIASES).flat().forEach((alias) => { if (joined.includes(norm(alias))) score += 2; }); if (/合计|总计|小计/.test(joined)) score -= 5; return score; }
function detectHeaderRow(grid: unknown[][]) { let best = 0; let bestScore = -Infinity; for (let i = 0; i < Math.min(40, grid.length); i += 1) { const score = scoreHeaderRow(grid[i] ?? []); if (score > bestScore) { best = i; bestScore = score; } } return best; }
function inferRole(header: unknown): Role | undefined { const h = norm(header); let best: { role?: Role; score: number } = { score: 0 }; (Object.entries(ROLE_ALIASES) as Array<[Role, string[]]>).forEach(([role, aliases]) => aliases.forEach((alias) => { const a = norm(alias); let score = 0; if (h === a) score = 1; else if (h.includes(a)) score = 0.86; else if (a.includes(h) && h.length >= 2) score = 0.55; if (score > best.score) best = { role, score }; })); return best.score >= 0.5 ? best.role : undefined; }
function findProductGroups(grid: unknown[][]) { const groups: Array<{ headerRow: number; nameCol: number; priceCol: number; purchaseCol: number; morningStockCol: number; morningSoldCol: number; morningAmountCol: number; nightStockCol: number; nightSoldCol: number; nightAmountCol: number; totalAmountCol: number }> = []; grid.forEach((row, headerRow) => row.forEach((cell, col) => { const label = norm(cell); const next = norm(row[col + 1]); if ((label === "商品名称" || label === "售卖商品") && (next === "价格" || next === "售价")) groups.push({ headerRow, nameCol: col, priceCol: col + 1, purchaseCol: col + 2, morningStockCol: col + 3, morningSoldCol: col + 4, morningAmountCol: col + 5, nightStockCol: col + 6, nightSoldCol: col + 7, nightAmountCol: col + 8, totalAmountCol: col + 9 }); })); return groups; }
function roleIndex(roles: Array<Role | undefined>, role: Role) { return roles.findIndex((item) => item === role); }
function getText(row: unknown[], index: number) { return index >= 0 ? text(row[index]) : ""; }
function getNumber(row: unknown[], index: number) { return index >= 0 ? toNumber(row[index]) : 0; }

function result(fileName: string, storeName: string, kind: BatchKind, rowCount: number, primaryAmount: number, secondaryAmount: number, diffAmount: number, issues: BatchIssue[], payments: PaymentRecord[] = []): BatchResult {
  return { fileName, storeName, kind, rowCount, primaryAmount: round2(primaryAmount), secondaryAmount: round2(secondaryAmount), diffAmount: round2(diffAmount), amount: round2(primaryAmount), issueCount: issues.length, channelCount: channelSummary(payments).length, issues, payments };
}

function parseProductFile(workbook: XLSX.WorkBook, fileName: string, storeName: string): BatchResult | null {
  let rowCount = 0;
  let salesAmount = 0;
  const issues: BatchIssue[] = [];
  const payments: PaymentRecord[] = [];
  workbook.SheetNames.forEach((sheetName, sheetIndex) => {
    const sheet = workbook.Sheets[sheetName]; if (!sheet) return;
    const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null });
    const date = extractDate(sheetName, grid[0]?.[0], sheetIndex);
    payments.push(...parsePaymentRecords(grid, sheetName, date, fileName, storeName));
    findProductGroups(grid).forEach((group) => {
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
        if (!(purchase || morningStock || morningSold || nightStock || nightSold || totalAmount)) continue;
        rowCount += 1;
        salesAmount += totalAmount;
        const expectedNightStock = round2(morningStock + purchase - morningSold);
        if (nightStockRaw !== null && nightStockRaw !== undefined && text(nightStockRaw) !== "" && Math.abs(nightStock - expectedNightStock) > 0.01) {
          issues.push({ fileName, storeName, type: "商品交接异常", message: `${sheetName} 第${r + 1}行 ${name}：夜班库存应为 ${expectedNightStock}，实际 ${nightStock}` });
        }
      }
    });
  });
  if (rowCount < 3) return null;
  const paymentRecords = paymentSourceForReconcile(payments);
  const paymentAmount = round2(paymentRecords.reduce((sum, record) => sum + record.amount, 0));
  return result(fileName, storeName, "product", rowCount, salesAmount, paymentAmount, paymentAmount - salesAmount, issues, payments);
}

function parseStructuredFile(workbook: XLSX.WorkBook, fileName: string, storeName: string): BatchResult {
  let rowCount = 0;
  let totalGross = 0;
  let totalNet = 0;
  let totalGeneric = 0;
  let totalMismatch = 0;
  let payrollScore = 0;
  let genericAmountScore = 0;
  const issues: BatchIssue[] = [];
  const seenEmployees = new Map<string, number>();

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName]; if (!sheet) return;
    const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null });
    if (!grid.length) return;
    const headerRow = detectHeaderRow(grid);
    const headers = grid[headerRow] ?? [];
    const roles = headers.map(inferRole);
    const amountCols = roles.map((role, index) => role === "amount" || role === "paidAmount" ? index : -1).filter((index) => index >= 0);
    const employeeCol = roleIndex(roles, "employee");
    const departmentCol = roleIndex(roles, "department");
    const positionCol = roleIndex(roles, "position");
    const baseCol = roleIndex(roles, "baseSalary");
    const expectedCol = roleIndex(roles, "expectedDays");
    const attendanceCol = roleIndex(roles, "attendanceDays");
    const hourlyCol = roleIndex(roles, "hourlyRate");
    const workHoursCol = roleIndex(roles, "workHours");
    const overtimeHoursCol = roleIndex(roles, "overtimeHours");
    const overtimePayCol = roleIndex(roles, "overtimePay");
    const commissionCol = roleIndex(roles, "commission");
    const bonusCol = roleIndex(roles, "bonus");
    const allowanceCol = roleIndex(roles, "allowance");
    const deductionCol = roleIndex(roles, "deduction");
    const advanceCol = roleIndex(roles, "advance");
    const socialCol = roleIndex(roles, "socialInsurance");
    const taxCol = roleIndex(roles, "tax");
    const grossCol = roleIndex(roles, "grossSalary");
    const netCol = roleIndex(roles, "netSalary");
    const remarkCol = roleIndex(roles, "remark");
    const sheetPayrollHits = [employeeCol, grossCol, netCol, baseCol, attendanceCol, deductionCol, advanceCol].filter((index) => index >= 0).length;
    payrollScore += sheetPayrollHits;
    genericAmountScore += amountCols.length;

    for (let r = headerRow + 1; r < grid.length; r += 1) {
      const row = grid[r] ?? [];
      if (!row.some((value) => text(value))) continue;
      if (sheetPayrollHits >= 2) {
        const employee = getText(row, employeeCol);
        const department = getText(row, departmentCol);
        const position = getText(row, positionCol);
        if (!employee && !department && !position) continue;
        if (/合计|总计|小计/.test(employee)) continue;
        rowCount += 1;
        const baseSalary = getNumber(row, baseCol);
        const expectedDays = getNumber(row, expectedCol);
        const attendanceDays = getNumber(row, attendanceCol);
        const hourlyRate = getNumber(row, hourlyCol);
        const workHours = getNumber(row, workHoursCol);
        const overtimeHours = getNumber(row, overtimeHoursCol);
        const explicitOvertimePay = getNumber(row, overtimePayCol);
        const overtimePay = explicitOvertimePay || round2(overtimeHours * hourlyRate * 1.5);
        const commission = getNumber(row, commissionCol);
        const bonus = getNumber(row, bonusCol);
        const allowance = getNumber(row, allowanceCol);
        const deduction = getNumber(row, deductionCol);
        const advance = getNumber(row, advanceCol);
        const socialInsurance = getNumber(row, socialCol);
        const tax = getNumber(row, taxCol);
        const grossSalary = getNumber(row, grossCol);
        const netSalary = getNumber(row, netCol);
        const remark = getText(row, remarkCol);
        const attendancePay = expectedDays > 0 && attendanceDays > 0 && baseSalary > 0 ? round2(baseSalary / expectedDays * attendanceDays) : workHours > 0 && hourlyRate > 0 ? round2(workHours * hourlyRate) : baseSalary;
        const calculatedGross = round2(attendancePay + overtimePay + commission + bonus + allowance);
        const calculatedNet = round2(calculatedGross - deduction - advance - socialInsurance - tax);
        const grossDiff = grossSalary ? round2(grossSalary - calculatedGross) : 0;
        const netDiff = netSalary ? round2(netSalary - calculatedNet) : 0;
        totalGross += grossSalary || calculatedGross;
        totalNet += netSalary || calculatedNet;
        totalMismatch += Math.abs(grossDiff) + Math.abs(netDiff);
        if (!employee) issues.push({ fileName, storeName, type: "工资表员工缺失", message: `${sheetName} 第${r + 1}行：员工姓名为空` });
        const duplicateKey = `${employee}__${department}`;
        if (employee && seenEmployees.has(duplicateKey)) issues.push({ fileName, storeName, type: "疑似重复员工", message: `${sheetName} 第${r + 1}行 ${employee}：与第 ${seenEmployees.get(duplicateKey)} 行员工重复` });
        else if (employee) seenEmployees.set(duplicateKey, r + 1);
        if (expectedDays && attendanceDays > expectedDays) issues.push({ fileName, storeName, type: "出勤异常", message: `${sheetName} 第${r + 1}行 ${employee}：实际出勤 ${attendanceDays} 大于应出勤 ${expectedDays}` });
        if (!expectedDays && attendanceDays > 31) issues.push({ fileName, storeName, type: "出勤异常", message: `${sheetName} 第${r + 1}行 ${employee}：实际出勤超过 31 天` });
        if (grossSalary && Math.abs(grossDiff) > 0.01) issues.push({ fileName, storeName, type: "应发工资不一致", message: `${sheetName} 第${r + 1}行 ${employee}：系统应发 ${calculatedGross}，表格应发 ${grossSalary}` });
        if (netSalary && Math.abs(netDiff) > 0.01) issues.push({ fileName, storeName, type: "实发工资不一致", message: `${sheetName} 第${r + 1}行 ${employee}：系统实发 ${calculatedNet}，表格实发 ${netSalary}` });
        if (netSalary < 0 || calculatedNet < 0) issues.push({ fileName, storeName, type: "负数工资", message: `${sheetName} 第${r + 1}行 ${employee}：实发工资或系统计算实发为负数` });
        if ((deduction || advance || socialInsurance || tax) && !remark) issues.push({ fileName, storeName, type: "扣款无备注", message: `${sheetName} 第${r + 1}行 ${employee}：存在扣款/借支/社保/个税，但备注为空` });
      } else {
        rowCount += 1;
        amountCols.forEach((index) => totalGeneric += toNumber(row[index]));
      }
    }
  });

  if (payrollScore >= 2) return result(fileName, storeName, "payroll", rowCount, totalGross, totalNet, totalMismatch, issues);
  return result(fileName, storeName, "generic", rowCount, totalGeneric, 0, 0, issues);
}

async function analyzeFile(file: File): Promise<BatchResult> {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const fileName = file.name;
  const storeName = storeFromFileName(fileName);
  return parseProductFile(workbook, fileName, storeName) ?? parseStructuredFile(workbook, fileName, storeName);
}

function exportBatch(results: BatchResult[]) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(results.map((item, index) => ({ 序号: index + 1, 门店: item.storeName, 文件: item.fileName, 类型: item.kind === "product" ? "商品表" : item.kind === "payroll" ? "工资表" : "通用表", 数据行: item.rowCount, 销售或应发: item.primaryAmount, 收款或实发: item.secondaryAmount, 核对差异: item.diffAmount, 渠道数: item.channelCount, 异常数: item.issueCount, 状态: item.issueCount ? "需检查" : "正常" }))), "全部门店总览");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(results.flatMap((item) => item.payments).map((record, index) => ({ 序号: index + 1, 门店: record.storeName, 文件: record.fileName, 日期: record.date, Sheet: record.sheetName, 班次: record.shift, 收款渠道: record.channel, 金额: record.amount }))), "收款明细");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(channelSummary(results.flatMap((item) => item.payments)).map((item, index) => ({ 序号: index + 1, 收款渠道: item.channel, 金额: item.amount, 记录数: item.count }))), "收款渠道汇总");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(results.flatMap((item) => item.issues).map((issue, index) => ({ 序号: index + 1, 门店: issue.storeName, 文件: issue.fileName, 类型: issue.type, 说明: issue.message }))), "异常明细");
  XLSX.writeFile(wb, `多门店批量核对报告-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

export default function MultiFileSmartUploadPanel() {
  const [results, setResults] = useState<BatchResult[]>([]);
  const [loading, setLoading] = useState(false);
  async function handleFiles(event: ChangeEvent<HTMLInputElement>) {
    const files = Array.from(event.target.files ?? []).filter((file) => /\.(xlsx|xls)$/i.test(file.name));
    if (!files.length) return;
    setLoading(true);
    try { setResults(await Promise.all(files.map(analyzeFile))); } finally { setLoading(false); event.target.value = ""; }
  }
  const totalPrimary = results.reduce((sum, item) => sum + item.primaryAmount, 0);
  const totalSecondary = results.reduce((sum, item) => sum + item.secondaryAmount, 0);
  const totalDiff = results.reduce((sum, item) => sum + item.diffAmount, 0);
  const totalRows = results.reduce((sum, item) => sum + item.rowCount, 0);
  const totalIssues = results.reduce((sum, item) => sum + item.issueCount, 0);

  return <section className="mx-auto mb-6 max-w-7xl px-4 md:px-8"><div className="rounded-2xl border bg-white p-5 shadow-sm">
    <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><h2 className="text-lg font-bold text-slate-900">多门店批量核对</h2><label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-slate-950 px-4 py-3 text-sm font-medium text-white">{loading ? "处理中..." : "选择多个 Excel"}<input className="hidden" type="file" accept=".xlsx,.xls" multiple onChange={handleFiles} /></label></div>
    {results.length ? <>
      <div className="mt-5 grid gap-4 md:grid-cols-5"><div className="rounded-xl bg-slate-50 p-4"><div className="text-sm text-slate-500">文件数</div><div className="mt-1 text-2xl font-bold">{results.length}</div></div><div className="rounded-xl bg-slate-50 p-4"><div className="text-sm text-slate-500">数据行</div><div className="mt-1 text-2xl font-bold">{totalRows}</div></div><div className="rounded-xl bg-slate-50 p-4"><div className="text-sm text-slate-500">销售/应发</div><div className="mt-1 text-2xl font-bold">¥{money(totalPrimary)}</div></div><div className="rounded-xl bg-slate-50 p-4"><div className="text-sm text-slate-500">收款/实发</div><div className="mt-1 text-2xl font-bold">¥{money(totalSecondary)}</div></div><div className={`rounded-xl p-4 ${Math.abs(totalDiff) > 0.01 || totalIssues ? "bg-red-600 text-white" : "bg-green-50 text-green-800"}`}><div className="text-sm opacity-80">核对差异</div><div className="mt-1 text-2xl font-bold">¥{money(totalDiff)}</div></div></div>
      <div className="mt-5 overflow-auto rounded-xl border"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["门店", "文件", "类型", "数据行", "销售/应发", "收款/实发", "核对差异", "渠道数", "异常"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{results.map((item) => <tr key={item.fileName} className="border-t"><td className="p-3">{item.storeName}</td><td className="p-3">{item.fileName}</td><td className="p-3">{item.kind === "product" ? "商品表" : item.kind === "payroll" ? "工资表" : "通用表"}</td><td className="p-3">{item.rowCount}</td><td className="p-3">¥{money(item.primaryAmount)}</td><td className="p-3">¥{money(item.secondaryAmount)}</td><td className="p-3">¥{money(item.diffAmount)}</td><td className="p-3">{item.channelCount}</td><td className="p-3">{item.issueCount}</td></tr>)}</tbody></table></div>
      <div className="mt-5 rounded-xl bg-slate-50 p-4"><div className="font-bold text-slate-900">收款渠道汇总</div><div className="mt-3 grid gap-2 md:grid-cols-3">{channelSummary(results.flatMap((item) => item.payments)).map((item) => <div key={item.channel} className="rounded-xl bg-white p-3 text-sm shadow-sm"><div className="font-medium">{item.channel}</div><div className="mt-1 text-lg font-bold">¥{money(item.amount)}</div><div className="text-xs text-slate-500">{item.count} 条记录</div></div>)}</div></div>
      {totalIssues ? <div className="mt-5 rounded-xl bg-red-50 p-4 text-sm text-red-800"><div className="font-bold">异常预览</div><ul className="mt-2 space-y-1">{results.flatMap((item) => item.issues).slice(0, 10).map((issue, index) => <li key={index}>• {issue.storeName}：{issue.message}</li>)}</ul></div> : null}
      <button onClick={() => exportBatch(results)} className="mt-5 rounded-xl bg-green-700 px-4 py-3 text-sm font-medium text-white">导出多门店总报告</button>
    </> : null}
  </div></section>;
}
