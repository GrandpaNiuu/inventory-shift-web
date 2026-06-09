"use client";

import { useMemo, useState, type ChangeEvent, type ReactNode } from "react";
import * as XLSX from "xlsx";

type Tab = "overview" | "rules" | "preview" | "summary" | "issues" | "auto" | "export";
type Severity = "error" | "warning" | "info";
type FieldKey =
  | "employee"
  | "department"
  | "position"
  | "baseSalary"
  | "expectedDays"
  | "attendanceDays"
  | "hourlyRate"
  | "workHours"
  | "overtimeHours"
  | "overtimePay"
  | "commission"
  | "bonus"
  | "allowance"
  | "deduction"
  | "advance"
  | "socialInsurance"
  | "tax"
  | "grossSalary"
  | "netSalary"
  | "remark";

type PayrollColumn = { key: FieldKey; label: string; aliases: string[] };
type ColumnMatch = { index: number; header: string; key?: FieldKey; label: string; confidence: number; examples: string[] };
type PayrollRow = {
  sheetName: string;
  rowNumber: number;
  employee: string;
  department: string;
  position: string;
  baseSalary: number;
  expectedDays: number;
  attendanceDays: number;
  hourlyRate: number;
  workHours: number;
  overtimeHours: number;
  overtimePay: number;
  commission: number;
  bonus: number;
  allowance: number;
  deduction: number;
  advance: number;
  socialInsurance: number;
  tax: number;
  grossSalary: number;
  netSalary: number;
  remark: string;
  calculatedAttendancePay: number;
  calculatedGross: number;
  calculatedNet: number;
  grossDiff: number;
  netDiff: number;
};
type PayrollIssue = { severity: Severity; type: string; sheetName: string; rowNumber: number; employee: string; message: string; expected?: number | string; actual?: number | string };
type SheetPreview = { sheetName: string; headerRow: number; rowCount: number; matchedFields: number };
type DepartmentSummary = { name: string; count: number; gross: number; net: number; issues: number };

type PayrollReport = { fileName: string; rows: PayrollRow[]; issues: PayrollIssue[]; previews: SheetPreview[]; columns: Record<string, ColumnMatch[]> };

const TABS: Array<[Tab, string]> = [["overview", "总览"], ["rules", "核对规则"], ["preview", "识别预览"], ["summary", "工资汇总"], ["issues", "核对异常"], ["auto", "自动结果"], ["export", "导出"]];

const PAYROLL_COLUMNS: PayrollColumn[] = [
  { key: "employee", label: "员工", aliases: ["员工", "员工姓名", "姓名", "人员", "工号", "employee", "staff", "worker"] },
  { key: "department", label: "部门", aliases: ["部门", "门店", "班组", "组织", "department", "team"] },
  { key: "position", label: "岗位", aliases: ["岗位", "职位", "职务", "工种", "position", "role", "job"] },
  { key: "baseSalary", label: "底薪/基本工资", aliases: ["基本工资", "底薪", "月薪", "工资标准", "固定工资", "base", "salary"] },
  { key: "expectedDays", label: "应出勤", aliases: ["应出勤", "应出勤天数", "满勤天数", "标准天数", "应上班", "expecteddays"] },
  { key: "attendanceDays", label: "实际出勤", aliases: ["实际出勤", "出勤", "出勤天数", "上班天数", "考勤", "attendance", "days"] },
  { key: "hourlyRate", label: "时薪", aliases: ["时薪", "小时工资", "每小时", "hourly", "hourlyrate"] },
  { key: "workHours", label: "工时", aliases: ["工时", "工作小时", "小时", "workhours", "hours"] },
  { key: "overtimeHours", label: "加班小时", aliases: ["加班小时", "加班工时", "overtimehours"] },
  { key: "overtimePay", label: "加班费", aliases: ["加班费", "加班工资", "overtime", "overtimepay"] },
  { key: "commission", label: "提成", aliases: ["提成", "业绩提成", "销售提成", "commission"] },
  { key: "bonus", label: "奖金", aliases: ["奖金", "满勤", "奖励", "绩效", "bonus", "performance"] },
  { key: "allowance", label: "补贴", aliases: ["补贴", "津贴", "餐补", "房补", "交通补贴", "allowance", "subsidy"] },
  { key: "deduction", label: "扣款", aliases: ["扣款", "罚款", "扣除", "缺勤扣款", "deduction", "fine"] },
  { key: "advance", label: "借支", aliases: ["借支", "预支", "借款", "advance", "loan"] },
  { key: "socialInsurance", label: "社保", aliases: ["社保", "五险", "保险", "social", "insurance"] },
  { key: "tax", label: "个税", aliases: ["个税", "个人所得税", "税", "tax"] },
  { key: "grossSalary", label: "应发工资", aliases: ["应发", "应发工资", "工资合计", "应付工资", "gross", "grosssalary"] },
  { key: "netSalary", label: "实发工资", aliases: ["实发", "实发工资", "到手", "实际发放", "net", "netsalary"] },
  { key: "remark", label: "备注", aliases: ["备注", "说明", "原因", "note", "remark", "comment"] }
];

function text(value: unknown) { return String(value ?? "").trim(); }
function norm(value: unknown) { return text(value).replace(/[\s\-_，,。.;；:：/\\|()（）\[\]【】{}<>《》"'“”‘’]+/g, "").toLowerCase(); }
function isBlank(value: unknown) { return value === null || value === undefined || text(value) === ""; }
function toNumber(value: unknown) { if (typeof value === "number") return Number.isFinite(value) ? value : 0; const cleaned = text(value).replace(/,/g, "").replace(/[￥¥元%\s]/g, ""); if (!cleaned) return 0; const parsed = Number(cleaned); return Number.isFinite(parsed) ? parsed : 0; }
function round2(value: number) { return Math.round((value + Number.EPSILON) * 100) / 100; }
function money(value: number) { return round2(value).toFixed(2); }
function colName(index: number) { return XLSX.utils.encode_col(index); }

function inferField(header: unknown): { key?: FieldKey; label: string; confidence: number } {
  const h = norm(header);
  let best: { key?: FieldKey; label: string; confidence: number } = { label: "未识别", confidence: 0 };
  PAYROLL_COLUMNS.forEach((field) => {
    field.aliases.forEach((alias) => {
      const a = norm(alias);
      let score = 0;
      if (h === a) score = 1;
      else if (h.includes(a)) score = 0.86;
      else if (a.includes(h) && h.length >= 2) score = 0.55;
      if (score > best.confidence) best = { key: field.key, label: field.label, confidence: score };
    });
  });
  return best.confidence >= 0.5 ? best : { label: "未识别", confidence: 0 };
}

function scoreHeaderRow(row: unknown[]) {
  const joined = norm(row.map(text).filter(Boolean).join(" "));
  let score = row.filter((cell) => !isBlank(cell)).length;
  PAYROLL_COLUMNS.flatMap((item) => item.aliases).forEach((alias) => { if (joined.includes(norm(alias))) score += 3; });
  if (/合计|总计|小计/.test(joined)) score -= 5;
  return score;
}
function detectHeaderRow(grid: unknown[][]) { let best = 0; let bestScore = -Infinity; for (let i = 0; i < Math.min(40, grid.length); i += 1) { const score = scoreHeaderRow(grid[i] ?? []); if (score > bestScore) { best = i; bestScore = score; } } return best; }
function cell(row: unknown[], matches: ColumnMatch[], key: FieldKey) { const match = matches.find((item) => item.key === key); return match ? row[match.index] : undefined; }
function num(row: unknown[], matches: ColumnMatch[], key: FieldKey) { return toNumber(cell(row, matches, key)); }
function str(row: unknown[], matches: ColumnMatch[], key: FieldKey) { return text(cell(row, matches, key)); }
function buildPayrollRow(sheetName: string, rowNumber: number, raw: unknown[], matches: ColumnMatch[]): PayrollRow | null {
  const employee = str(raw, matches, "employee");
  const department = str(raw, matches, "department");
  const position = str(raw, matches, "position");
  if (!employee && !department && !position) return null;
  if (/合计|总计|小计/.test(employee)) return null;
  const baseSalary = num(raw, matches, "baseSalary");
  const expectedDays = num(raw, matches, "expectedDays");
  const attendanceDays = num(raw, matches, "attendanceDays");
  const hourlyRate = num(raw, matches, "hourlyRate");
  const workHours = num(raw, matches, "workHours");
  const overtimeHours = num(raw, matches, "overtimeHours");
  const overtimePay = num(raw, matches, "overtimePay") || round2(overtimeHours * hourlyRate * 1.5);
  const commission = num(raw, matches, "commission");
  const bonus = num(raw, matches, "bonus");
  const allowance = num(raw, matches, "allowance");
  const deduction = num(raw, matches, "deduction");
  const advance = num(raw, matches, "advance");
  const socialInsurance = num(raw, matches, "socialInsurance");
  const tax = num(raw, matches, "tax");
  const grossSalary = num(raw, matches, "grossSalary");
  const netSalary = num(raw, matches, "netSalary");
  const remark = str(raw, matches, "remark");
  const calculatedAttendancePay = expectedDays > 0 && attendanceDays > 0 && baseSalary > 0 ? round2(baseSalary / expectedDays * attendanceDays) : workHours > 0 && hourlyRate > 0 ? round2(workHours * hourlyRate) : baseSalary;
  const calculatedGross = round2(calculatedAttendancePay + overtimePay + commission + bonus + allowance);
  const calculatedNet = round2(calculatedGross - deduction - advance - socialInsurance - tax);
  return { sheetName, rowNumber, employee, department, position, baseSalary, expectedDays, attendanceDays, hourlyRate, workHours, overtimeHours, overtimePay, commission, bonus, allowance, deduction, advance, socialInsurance, tax, grossSalary, netSalary, remark, calculatedAttendancePay, calculatedGross, calculatedNet, grossDiff: grossSalary ? round2(grossSalary - calculatedGross) : 0, netDiff: netSalary ? round2(netSalary - calculatedNet) : 0 };
}
function buildIssues(rows: PayrollRow[]): PayrollIssue[] {
  const issues: PayrollIssue[] = [];
  const seen = new Map<string, number>();
  rows.forEach((row) => {
    if (!row.employee) issues.push({ severity: "error", type: "员工缺失", sheetName: row.sheetName, rowNumber: row.rowNumber, employee: row.employee, message: "员工姓名为空。" });
    const key = `${row.employee}__${row.department || ""}`;
    if (row.employee && seen.has(key)) issues.push({ severity: "warning", type: "疑似重复员工", sheetName: row.sheetName, rowNumber: row.rowNumber, employee: row.employee, message: `与第 ${seen.get(key)} 行员工重复。` });
    else if (row.employee) seen.set(key, row.rowNumber);
    if (row.expectedDays && row.attendanceDays > row.expectedDays) issues.push({ severity: "warning", type: "出勤异常", sheetName: row.sheetName, rowNumber: row.rowNumber, employee: row.employee, message: "实际出勤天数大于应出勤天数。", expected: row.expectedDays, actual: row.attendanceDays });
    if (!row.expectedDays && row.attendanceDays > 31) issues.push({ severity: "warning", type: "出勤异常", sheetName: row.sheetName, rowNumber: row.rowNumber, employee: row.employee, message: "实际出勤天数超过 31 天。", actual: row.attendanceDays });
    if (row.grossSalary && Math.abs(row.grossDiff) > 0.01) issues.push({ severity: "error", type: "应发工资不一致", sheetName: row.sheetName, rowNumber: row.rowNumber, employee: row.employee, message: "应发工资不等于系统计算应发。", expected: row.calculatedGross, actual: row.grossSalary });
    if (row.netSalary && Math.abs(row.netDiff) > 0.01) issues.push({ severity: "error", type: "实发工资不一致", sheetName: row.sheetName, rowNumber: row.rowNumber, employee: row.employee, message: "实发工资不等于系统计算实发。", expected: row.calculatedNet, actual: row.netSalary });
    if (row.netSalary < 0 || row.calculatedNet < 0) issues.push({ severity: "warning", type: "负数工资", sheetName: row.sheetName, rowNumber: row.rowNumber, employee: row.employee, message: "实发工资或系统计算实发为负数。", actual: row.netSalary || row.calculatedNet });
    if ((row.deduction || row.advance || row.socialInsurance || row.tax) && !row.remark) issues.push({ severity: "info", type: "扣款无备注", sheetName: row.sheetName, rowNumber: row.rowNumber, employee: row.employee, message: "存在扣款/借支/社保/个税，但备注为空。" });
  });
  return issues;
}
function parseWorkbook(workbook: XLSX.WorkBook, fileName: string): PayrollReport {
  const rows: PayrollRow[] = [];
  const previews: SheetPreview[] = [];
  const columns: Record<string, ColumnMatch[]> = {};
  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;
    const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null });
    const headerRow = detectHeaderRow(grid);
    const headers = grid[headerRow] ?? [];
    const matches = headers.map((header, index): ColumnMatch => {
      const guessed = inferField(header);
      const examples = grid.slice(headerRow + 1, headerRow + 4).map((row) => text(row?.[index])).filter(Boolean);
      return { index, header: text(header) || `列${colName(index)}`, key: guessed.key, label: guessed.label, confidence: guessed.confidence, examples };
    });
    columns[sheetName] = matches;
    let rowCount = 0;
    for (let r = headerRow + 1; r < grid.length; r += 1) {
      const raw = grid[r] ?? [];
      if (!raw.some((value) => !isBlank(value))) continue;
      const payrollRow = buildPayrollRow(sheetName, r + 1, raw, matches);
      if (!payrollRow) continue;
      rows.push(payrollRow);
      rowCount += 1;
    }
    previews.push({ sheetName, headerRow, rowCount, matchedFields: matches.filter((item) => item.key).length });
  });
  return { fileName, rows, previews, columns, issues: buildIssues(rows) };
}
function groupSummary(rows: PayrollRow[], key: "department" | "position") {
  const map = new Map<string, DepartmentSummary>();
  rows.forEach((row) => {
    const name = row[key] || "未识别";
    const old = map.get(name) ?? { name, count: 0, gross: 0, net: 0, issues: 0 };
    old.count += 1;
    old.gross += row.grossSalary || row.calculatedGross;
    old.net += row.netSalary || row.calculatedNet;
    old.issues += Math.abs(row.grossDiff) > 0.01 || Math.abs(row.netDiff) > 0.01 ? 1 : 0;
    map.set(name, old);
  });
  return Array.from(map.values()).map((item) => ({ ...item, gross: round2(item.gross), net: round2(item.net) })).sort((a, b) => b.net - a.net);
}
function exportPayroll(report: PayrollReport) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{ 文件: report.fileName, 员工数: report.rows.length, 应发合计: round2(report.rows.reduce((s, r) => s + (r.grossSalary || r.calculatedGross), 0)), 实发合计: round2(report.rows.reduce((s, r) => s + (r.netSalary || r.calculatedNet), 0)), 异常数: report.issues.length }]), "工资总览");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(report.rows.map((row, index) => ({ 序号: index + 1, Sheet: row.sheetName, 行号: row.rowNumber, 员工: row.employee, 部门: row.department, 岗位: row.position, 底薪: row.baseSalary, 应出勤: row.expectedDays, 实际出勤: row.attendanceDays, 出勤工资: row.calculatedAttendancePay, 加班费: row.overtimePay, 提成: row.commission, 奖金: row.bonus, 补贴: row.allowance, 扣款: row.deduction, 借支: row.advance, 社保: row.socialInsurance, 个税: row.tax, 系统应发: row.calculatedGross, 表格应发: row.grossSalary, 应发差异: row.grossDiff, 系统实发: row.calculatedNet, 表格实发: row.netSalary, 实发差异: row.netDiff, 备注: row.remark }))), "员工工资明细");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(report.issues.map((issue, index) => ({ 序号: index + 1, 级别: issue.severity, 类型: issue.type, Sheet: issue.sheetName, 行号: issue.rowNumber, 员工: issue.employee, 说明: issue.message, 应为: issue.expected ?? "", 实际: issue.actual ?? "" }))), "异常工资清单");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(groupSummary(report.rows, "department").map((item, index) => ({ 序号: index + 1, 部门: item.name, 人数: item.count, 应发合计: item.gross, 实发合计: item.net, 异常人数: item.issues }))), "部门工资汇总");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(groupSummary(report.rows, "position").map((item, index) => ({ 序号: index + 1, 岗位: item.name, 人数: item.count, 应发合计: item.gross, 实发合计: item.net, 异常人数: item.issues }))), "岗位工资汇总");
  XLSX.writeFile(wb, `${report.fileName.replace(/\.xlsx?$/i, "")}-工资核对报告.xlsx`);
}
function Card({ title, value, desc, danger = false }: { title: string; value: string; desc?: string; danger?: boolean }) { return <div className={`rounded-2xl p-5 shadow-sm ${danger ? "bg-red-600 text-white" : "border bg-white text-slate-900"}`}><div className={`text-sm ${danger ? "text-red-100" : "text-slate-500"}`}>{title}</div><div className="mt-2 text-2xl font-bold">{value}</div>{desc ? <div className={`mt-1 text-xs ${danger ? "text-red-100" : "text-slate-500"}`}>{desc}</div> : null}</div>; }
function Badge({ children, tone = "gray" }: { children: ReactNode; tone?: "gray" | "green" | "red" | "yellow" }) { const cls = { gray: "bg-slate-100 text-slate-700", green: "bg-green-100 text-green-700", red: "bg-red-100 text-red-700", yellow: "bg-yellow-100 text-yellow-800" }[tone]; return <span className={`rounded-full px-2 py-1 text-xs font-medium ${cls}`}>{children}</span>; }
function DataTable({ rows, headers }: { rows: Record<string, string | number>[]; headers: string[] }) { return <div className="table-scroll mt-3 max-h-96 overflow-auto"><table className="min-w-full text-sm"><thead className="sticky top-0 bg-slate-50 text-left"><tr>{headers.map((header) => <th key={header} className="p-3">{header}</th>)}</tr></thead><tbody>{rows.map((row, index) => <tr key={index} className="border-t">{headers.map((header) => <td key={header} className="p-3">{row[header]}</td>)}</tr>)}</tbody></table></div>; }
function Section({ title, desc, children }: { title: string; desc?: string; children: ReactNode }) { return <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">{title}</h2>{desc ? <p className="mt-1 text-sm text-slate-500">{desc}</p> : null}{children}</section>; }

export default function PayrollReportApp() {
  const [tab, setTab] = useState<Tab>("overview");
  const [report, setReport] = useState<PayrollReport | null>(null);
  async function handleFile(file: File) { const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true, cellFormula: true, cellNF: true, cellStyles: true }); setReport(parseWorkbook(workbook, file.name)); setTab("overview"); }
  const rows = report?.rows ?? [];
  const issues = report?.issues ?? [];
  const totalGross = round2(rows.reduce((sum, row) => sum + (row.grossSalary || row.calculatedGross), 0));
  const totalNet = round2(rows.reduce((sum, row) => sum + (row.netSalary || row.calculatedNet), 0));
  const totalDeduction = round2(rows.reduce((sum, row) => sum + row.deduction + row.advance + row.socialInsurance + row.tax, 0));
  return <main className="min-h-screen p-4 md:p-8"><div className="mx-auto max-w-7xl">
    <header className="mb-6 rounded-3xl bg-slate-950 p-6 text-white shadow-sm"><p className="text-sm text-slate-300">Payroll Reconciliation</p><h1 className="mt-1 text-2xl font-bold md:text-4xl">工资表专用核对</h1><p className="mt-2 max-w-4xl text-sm leading-6 text-slate-300">上传工资表后自动识别员工、部门、岗位、底薪、出勤、加班、提成、奖金、补贴、扣款、借支、社保、个税、应发和实发。</p></header>
    <section className="mb-6 rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">导入工资表 Excel</h2><p className="mt-1 text-sm text-slate-500">系统会自动计算应发、实发，并与原表金额对比。</p></div><label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-slate-950 px-4 py-3 text-sm font-medium text-white">选择工资表<input className="hidden" type="file" accept=".xlsx,.xls" onChange={(event: ChangeEvent<HTMLInputElement>) => event.target.files?.[0] && handleFile(event.target.files[0])} /></label></div>{report ? <p className="mt-3 text-sm text-slate-600">已导入：<b>{report.fileName}</b>，识别员工 {rows.length} 行，异常 {issues.length} 条。</p> : null}</section>
    <nav className="mb-6 flex flex-wrap gap-2">{TABS.map(([key, label]) => <button key={key} onClick={() => setTab(key)} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === key ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>{label}</button>)}</nav>
    {!report ? <Section title="等待导入工资表" desc="支持固定月薪、出勤天数、工时、加班、提成、奖金、补贴、扣款、借支、社保、个税等字段。"><div className="mt-3 grid gap-3 md:grid-cols-4">{TABS.map(([, label]) => <div key={label} className="rounded-xl bg-slate-50 p-3 text-sm font-medium text-slate-700">{label}</div>)}</div></Section> : null}
    {report && tab === "overview" ? <Section title="总览" desc="工资表整体核对结果。"><div className="mt-4 grid gap-4 md:grid-cols-4"><Card title="员工行数" value={`${rows.length} 行`} /><Card title="应发合计" value={`¥${money(totalGross)}`} /><Card title="实发合计" value={`¥${money(totalNet)}`} /><Card title="异常数量" value={`${issues.length} 条`} danger={issues.length > 0} /></div><div className="mt-4 grid gap-4 md:grid-cols-3"><Card title="扣款/借支/社保/个税" value={`¥${money(totalDeduction)}`} /><Card title="部门数量" value={`${groupSummary(rows, "department").length} 个`} /><Card title="岗位数量" value={`${groupSummary(rows, "position").length} 个`} /></div></Section> : null}
    {report && tab === "rules" ? <Section title="核对规则" desc="系统当前使用的工资核对规则。"><div className="mt-3 space-y-2">{["出勤工资 = 底薪 ÷ 应出勤天数 × 实际出勤天数", "工时工资 = 工时 × 时薪", "加班费 = 加班小时 × 时薪 × 1.5（表格未填加班费时自动估算）", "系统应发 = 出勤工资/底薪 + 加班费 + 提成 + 奖金 + 补贴", "系统实发 = 系统应发 - 扣款 - 借支 - 社保 - 个税", "核对表格应发与系统应发差异", "核对表格实发与系统实发差异", "检查重复员工、出勤异常、负数工资、扣款无备注"].map((rule, index) => <div key={rule} className="rounded-xl bg-slate-50 p-3 text-sm"><b>{index + 1}.</b> {rule}</div>)}</div></Section> : null}
    {report && tab === "preview" ? <Section title="识别预览" desc="查看工资表每个 Sheet 的字段识别情况。"><DataTable rows={Object.entries(report.columns).flatMap(([sheetName, matches]) => matches.map((match) => ({ Sheet: sheetName, 列: colName(match.index), 表头: match.header, 识别字段: match.label, 置信度: `${Math.round(match.confidence * 100)}%`, 示例: match.examples.join(" / ") })))} headers={["Sheet", "列", "表头", "识别字段", "置信度", "示例"]} /></Section> : null}
    {report && tab === "summary" ? <section className="space-y-5"><Section title="部门工资汇总" desc="按部门统计人数、应发、实发和异常人数。"><DataTable rows={groupSummary(rows, "department").map((item, index) => ({ 序号: index + 1, 部门: item.name, 人数: item.count, 应发合计: `¥${money(item.gross)}`, 实发合计: `¥${money(item.net)}`, 异常人数: item.issues }))} headers={["序号", "部门", "人数", "应发合计", "实发合计", "异常人数"]} /></Section><Section title="岗位工资汇总" desc="按岗位统计工资。"><DataTable rows={groupSummary(rows, "position").map((item, index) => ({ 序号: index + 1, 岗位: item.name, 人数: item.count, 应发合计: `¥${money(item.gross)}`, 实发合计: `¥${money(item.net)}`, 异常人数: item.issues }))} headers={["序号", "岗位", "人数", "应发合计", "实发合计", "异常人数"]} /></Section></section> : null}
    {report && tab === "issues" ? <Section title="核对异常" desc="集中显示工资表异常。">{issues.length ? <DataTable rows={issues.map((issue, index) => ({ 序号: index + 1, 级别: issue.severity, 类型: issue.type, Sheet: issue.sheetName, 行号: issue.rowNumber, 员工: issue.employee, 说明: issue.message, 应为: issue.expected ?? "", 实际: issue.actual ?? "" }))} headers={["序号", "级别", "类型", "Sheet", "行号", "员工", "说明", "应为", "实际"]} /> : <p className="mt-3 text-sm text-green-700">当前工资表没有发现明显异常。</p>}</Section> : null}
    {report && tab === "auto" ? <Section title="自动结果" desc="系统计算出的工资结果和原表对比。"><DataTable rows={rows.map((row, index) => ({ 序号: index + 1, 员工: row.employee, 部门: row.department, 岗位: row.position, 出勤工资: row.calculatedAttendancePay, 加班费: row.overtimePay, 提成: row.commission, 奖金: row.bonus, 补贴: row.allowance, 系统应发: row.calculatedGross, 表格应发: row.grossSalary, 应发差异: row.grossDiff, 扣款: row.deduction, 借支: row.advance, 社保: row.socialInsurance, 个税: row.tax, 系统实发: row.calculatedNet, 表格实发: row.netSalary, 实发差异: row.netDiff }))} headers={["序号", "员工", "部门", "岗位", "出勤工资", "加班费", "提成", "奖金", "补贴", "系统应发", "表格应发", "应发差异", "扣款", "借支", "社保", "个税", "系统实发", "表格实发", "实发差异"]} /></Section> : null}
    {report && tab === "export" ? <Section title="导出" desc="导出工资核对报告，包含总览、员工工资明细、异常清单、部门汇总和岗位汇总。"><button onClick={() => exportPayroll(report)} className="mt-4 rounded-xl bg-green-700 px-4 py-3 text-sm font-medium text-white">导出工资核对报告</button></Section> : null}
  </div></main>;
}
