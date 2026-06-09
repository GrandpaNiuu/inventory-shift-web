"use client";

import { useState, type ChangeEvent } from "react";
import * as XLSX from "xlsx";

type BatchKind = "product" | "payroll" | "generic";
type BatchIssue = { fileName: string; storeName: string; type: string; message: string };
type BatchResult = { fileName: string; storeName: string; kind: BatchKind; rowCount: number; amount: number; issueCount: number; issues: BatchIssue[] };
type Role = "product" | "price" | "amount" | "paidAmount" | "employee" | "department" | "baseSalary" | "attendanceDays" | "grossSalary" | "netSalary" | "deduction" | "advance" | "tax" | "socialInsurance";

const ROLE_ALIASES: Record<Role, string[]> = {
  product: ["商品", "商品名称", "售卖商品", "产品", "品名", "货品"], price: ["价格", "单价", "售价", "销售价"], amount: ["金额", "小计", "销售额", "收入", "合计"], paidAmount: ["实收", "实收金额", "到账", "收款", "收款金额", "已收"], employee: ["员工", "员工姓名", "姓名", "人员", "工号"], department: ["部门", "门店", "店铺", "分店", "班组"], baseSalary: ["基本工资", "底薪", "月薪", "工资标准"], attendanceDays: ["出勤", "出勤天数", "实际出勤", "上班天数"], grossSalary: ["应发", "应发工资", "工资合计", "应付工资"], netSalary: ["实发", "实发工资", "到手", "实际发放"], deduction: ["扣款", "罚款", "扣除", "缺勤扣款"], advance: ["借支", "预支", "借款"], tax: ["个税", "个人所得税", "税"], socialInsurance: ["社保", "五险", "保险"]
};

function text(value: unknown) { return String(value ?? "").trim(); }
function norm(value: unknown) { return text(value).replace(/[\s\-_，,。.;；:：/\\|()（）\[\]【】{}<>《》"'“”‘’]+/g, "").toLowerCase(); }
function toNumber(value: unknown) { if (typeof value === "number") return Number.isFinite(value) ? value : 0; const parsed = Number(text(value).replace(/,/g, "").replace(/[￥¥元%\s]/g, "")); return Number.isFinite(parsed) ? parsed : 0; }
function round2(value: number) { return Math.round((value + Number.EPSILON) * 100) / 100; }
function money(value: number) { return round2(value).toFixed(2); }
function storeFromFileName(fileName: string) { return fileName.replace(/\.(xlsx|xls)$/i, "").replace(/\d{4}[-年.]?\d{1,2}[-月.]?\d{0,2}[日]?/g, "").replace(/报表|日报|月报|工资表|库存表|核对表/g, "").replace(/[-_ ]+/g, " ").trim() || fileName.replace(/\.(xlsx|xls)$/i, ""); }
function scoreHeaderRow(row: unknown[]) { const joined = norm(row.map(text).filter(Boolean).join(" ")); let score = row.filter((cell) => text(cell)).length; Object.values(ROLE_ALIASES).flat().forEach((alias) => { if (joined.includes(norm(alias))) score += 2; }); if (/合计|总计|小计/.test(joined)) score -= 5; return score; }
function detectHeaderRow(grid: unknown[][]) { let best = 0; let bestScore = -Infinity; for (let i = 0; i < Math.min(40, grid.length); i += 1) { const score = scoreHeaderRow(grid[i] ?? []); if (score > bestScore) { best = i; bestScore = score; } } return best; }
function inferRole(header: unknown): Role | undefined { const h = norm(header); let best: { role?: Role; score: number } = { score: 0 }; (Object.entries(ROLE_ALIASES) as Array<[Role, string[]]>).forEach(([role, aliases]) => aliases.forEach((alias) => { const a = norm(alias); let score = 0; if (h === a) score = 1; else if (h.includes(a)) score = 0.86; else if (a.includes(h) && h.length >= 2) score = 0.55; if (score > best.score) best = { role, score }; })); return best.score >= 0.5 ? best.role : undefined; }
function findProductGroups(grid: unknown[][]) { const groups: Array<{ headerRow: number; nameCol: number; priceCol: number; purchaseCol: number; morningStockCol: number; morningSoldCol: number; morningAmountCol: number; nightStockCol: number; nightSoldCol: number; nightAmountCol: number; totalAmountCol: number }> = []; grid.forEach((row, headerRow) => row.forEach((cell, col) => { const label = norm(cell); const next = norm(row[col + 1]); if ((label === "商品名称" || label === "售卖商品") && (next === "价格" || next === "售价")) groups.push({ headerRow, nameCol: col, priceCol: col + 1, purchaseCol: col + 2, morningStockCol: col + 3, morningSoldCol: col + 4, morningAmountCol: col + 5, nightStockCol: col + 6, nightSoldCol: col + 7, nightAmountCol: col + 8, totalAmountCol: col + 9 }); })); return groups; }

function parseProductFile(workbook: XLSX.WorkBook, fileName: string, storeName: string): BatchResult | null {
  let rowCount = 0; let amount = 0; const issues: BatchIssue[] = [];
  workbook.SheetNames.forEach((sheetName) => { const sheet = workbook.Sheets[sheetName]; if (!sheet) return; const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null }); findProductGroups(grid).forEach((group) => { for (let r = group.headerRow + 1; r < grid.length; r += 1) { const line = grid[r] ?? []; const name = text(line[group.nameCol]); const price = toNumber(line[group.priceCol]); if (!name || !price || /合计|总计|小计|备注|收款|金额/.test(name)) continue; const purchase = toNumber(line[group.purchaseCol]); const morningStock = toNumber(line[group.morningStockCol]); const morningSold = toNumber(line[group.morningSoldCol]); const morningAmount = toNumber(line[group.morningAmountCol]) || round2(morningSold * price); const nightStockRaw = line[group.nightStockCol]; const nightStock = toNumber(nightStockRaw); const nightSold = toNumber(line[group.nightSoldCol]); const nightAmount = toNumber(line[group.nightAmountCol]) || round2(nightSold * price); const totalAmount = toNumber(line[group.totalAmountCol]) || round2(morningAmount + nightAmount); if (!(purchase || morningStock || morningSold || nightStock || nightSold || totalAmount)) continue; rowCount += 1; amount += totalAmount; const expectedNightStock = round2(morningStock + purchase - morningSold); if (nightStockRaw !== null && nightStockRaw !== undefined && text(nightStockRaw) !== "" && Math.abs(nightStock - expectedNightStock) > 0.01) issues.push({ fileName, storeName, type: "商品交接异常", message: `${sheetName} 第${r + 1}行 ${name}：夜班库存应为 ${expectedNightStock}，实际 ${nightStock}` }); } }); });
  if (rowCount < 3) return null; return { fileName, storeName, kind: "product", rowCount, amount: round2(amount), issueCount: issues.length, issues };
}

function parseStructuredFile(workbook: XLSX.WorkBook, fileName: string, storeName: string): BatchResult {
  let rowCount = 0; let amount = 0; let payrollScore = 0; let genericAmountScore = 0; const issues: BatchIssue[] = [];
  workbook.SheetNames.forEach((sheetName) => { const sheet = workbook.Sheets[sheetName]; if (!sheet) return; const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null }); if (!grid.length) return; const headerRow = detectHeaderRow(grid); const headers = grid[headerRow] ?? []; const roles = headers.map(inferRole); const employeeCol = roles.findIndex((role) => role === "employee"); const grossCol = roles.findIndex((role) => role === "grossSalary"); const netCol = roles.findIndex((role) => role === "netSalary"); const baseCol = roles.findIndex((role) => role === "baseSalary"); const deductionCol = roles.findIndex((role) => role === "deduction"); const advanceCol = roles.findIndex((role) => role === "advance"); const amountCols = roles.map((role, index) => role === "amount" || role === "paidAmount" ? index : -1).filter((index) => index >= 0); payrollScore += [employeeCol, grossCol, netCol, baseCol, deductionCol, advanceCol].filter((index) => index >= 0).length; genericAmountScore += amountCols.length; for (let r = headerRow + 1; r < grid.length; r += 1) { const row = grid[r] ?? []; if (!row.some((value) => text(value))) continue; rowCount += 1; if (payrollScore >= 2) { const employee = employeeCol >= 0 ? text(row[employeeCol]) : ""; const gross = grossCol >= 0 ? toNumber(row[grossCol]) : 0; const net = netCol >= 0 ? toNumber(row[netCol]) : 0; const deduction = deductionCol >= 0 ? toNumber(row[deductionCol]) : 0; const advance = advanceCol >= 0 ? toNumber(row[advanceCol]) : 0; amount += net || gross; if (!employee) issues.push({ fileName, storeName, type: "工资表员工缺失", message: `${sheetName} 第${r + 1}行：员工姓名为空` }); if (net && gross && net > gross + 0.01) issues.push({ fileName, storeName, type: "工资金额异常", message: `${sheetName} 第${r + 1}行 ${employee}：实发工资大于应发工资` }); if ((deduction || advance) && !net && !gross) issues.push({ fileName, storeName, type: "工资字段缺失", message: `${sheetName} 第${r + 1}行 ${employee}：有扣款或借支，但缺少应发/实发金额` }); } else { amountCols.forEach((index) => amount += toNumber(row[index])); } } });
  const kind: BatchKind = payrollScore >= 2 ? "payroll" : genericAmountScore > 0 ? "generic" : "generic"; return { fileName, storeName, kind, rowCount, amount: round2(amount), issueCount: issues.length, issues };
}

async function analyzeFile(file: File): Promise<BatchResult> { const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true }); const fileName = file.name; const storeName = storeFromFileName(fileName); const product = parseProductFile(workbook, fileName, storeName); return product ?? parseStructuredFile(workbook, fileName, storeName); }
function exportBatch(results: BatchResult[]) { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(results.map((item, index) => ({ 序号: index + 1, 门店: item.storeName, 文件: item.fileName, 类型: item.kind === "product" ? "商品表" : item.kind === "payroll" ? "工资表" : "通用表", 数据行: item.rowCount, 金额合计: item.amount, 异常数: item.issueCount, 状态: item.issueCount ? "需检查" : "正常" }))), "全部门店总览"); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(results.flatMap((item) => item.issues).map((issue, index) => ({ 序号: index + 1, 门店: issue.storeName, 文件: issue.fileName, 类型: issue.type, 说明: issue.message }))), "异常明细"); XLSX.writeFile(wb, `多门店批量核对报告-${new Date().toISOString().slice(0, 10)}.xlsx`); }

export default function MultiFileSmartUploadPanel() {
  const [results, setResults] = useState<BatchResult[]>([]);
  const [loading, setLoading] = useState(false);
  async function handleFiles(event: ChangeEvent<HTMLInputElement>) { const files = Array.from(event.target.files ?? []).filter((file) => /\.(xlsx|xls)$/i.test(file.name)); if (!files.length) return; setLoading(true); try { const next = await Promise.all(files.map(analyzeFile)); setResults(next); } finally { setLoading(false); event.target.value = ""; } }
  const totalAmount = results.reduce((sum, item) => sum + item.amount, 0);
  const totalRows = results.reduce((sum, item) => sum + item.rowCount, 0);
  const totalIssues = results.reduce((sum, item) => sum + item.issueCount, 0);

  return <section className="mx-auto mb-6 max-w-7xl px-4 md:px-8">
    <div className="rounded-2xl border bg-white p-5 shadow-sm">
      <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
        <h2 className="text-lg font-bold text-slate-900">多门店批量核对</h2>
        <label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-slate-950 px-4 py-3 text-sm font-medium text-white">
          {loading ? "处理中..." : "选择多个 Excel"}
          <input className="hidden" type="file" accept=".xlsx,.xls" multiple onChange={handleFiles} />
        </label>
      </div>
      {results.length ? <>
        <div className="mt-5 grid gap-4 md:grid-cols-4">
          <div className="rounded-xl bg-slate-50 p-4"><div className="text-sm text-slate-500">文件数</div><div className="mt-1 text-2xl font-bold">{results.length}</div></div>
          <div className="rounded-xl bg-slate-50 p-4"><div className="text-sm text-slate-500">数据行</div><div className="mt-1 text-2xl font-bold">{totalRows}</div></div>
          <div className="rounded-xl bg-slate-50 p-4"><div className="text-sm text-slate-500">金额合计</div><div className="mt-1 text-2xl font-bold">¥{money(totalAmount)}</div></div>
          <div className={`rounded-xl p-4 ${totalIssues ? "bg-red-600 text-white" : "bg-green-50 text-green-800"}`}><div className="text-sm opacity-80">异常数</div><div className="mt-1 text-2xl font-bold">{totalIssues}</div></div>
        </div>
        <div className="mt-5 overflow-auto rounded-xl border"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["门店", "文件", "类型", "数据行", "金额合计", "异常", "状态"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{results.map((item) => <tr key={item.fileName} className="border-t"><td className="p-3">{item.storeName}</td><td className="p-3">{item.fileName}</td><td className="p-3">{item.kind === "product" ? "商品表" : item.kind === "payroll" ? "工资表" : "通用表"}</td><td className="p-3">{item.rowCount}</td><td className="p-3">¥{money(item.amount)}</td><td className="p-3">{item.issueCount}</td><td className="p-3">{item.issueCount ? "需检查" : "正常"}</td></tr>)}</tbody></table></div>
        {totalIssues ? <div className="mt-5 rounded-xl bg-red-50 p-4 text-sm text-red-800"><div className="font-bold">异常预览</div><ul className="mt-2 space-y-1">{results.flatMap((item) => item.issues).slice(0, 10).map((issue, index) => <li key={index}>• {issue.storeName}：{issue.message}</li>)}</ul></div> : null}
        <button onClick={() => exportBatch(results)} className="mt-5 rounded-xl bg-green-700 px-4 py-3 text-sm font-medium text-white">导出多门店总报告</button>
      </> : null}
    </div>
  </section>;
}
