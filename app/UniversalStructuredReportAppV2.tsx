"use client";

import { useMemo, useState, type ChangeEvent, type ReactNode } from "react";
import * as XLSX from "xlsx";
import ProductReportApp from "./InventoryPageConfigurable";
import { classifyGlobalProduct } from "./globalProductCatalog";

type MainTab = "smart" | "product" | "guide";
type ReportTab = "overview" | "rules" | "preview" | "summary" | "issues" | "auto" | "export";
type Severity = "error" | "warning" | "info";
type GenericType = "商品销售/库存表" | "工资/考勤表" | "进货采购表" | "费用报销表" | "收款对账表" | "应收应付表" | "订单流水表" | "客户供应商表" | "资产设备表" | "通用业务表";
type Role = "date" | "product" | "employee" | "customer" | "supplier" | "orderNo" | "sku" | "barcode" | "category" | "quantity" | "price" | "amount" | "totalAmount" | "paidAmount" | "receivable" | "payable" | "stockOpening" | "stockIn" | "stockOut" | "stockEnding" | "purchase" | "sold" | "basicSalary" | "attendance" | "overtime" | "commission" | "bonus" | "deduction" | "advance" | "grossSalary" | "netSalary" | "feeType" | "paymentChannel" | "department" | "position" | "asset" | "assetNo" | "status" | "remark";

type DataRow = { rowNumber: number; values: Record<string, string | number> };
type ColumnInfo = { index: number; header: string; role?: Role; confidence: number; examples: string[]; nonEmpty: number };
type GenericSheet = { sheetName: string; type: GenericType; confidence: number; headerRow: number; rows: DataRow[]; columns: ColumnInfo[]; rules: string[]; issues: Issue[] };
type Issue = { severity: Severity; type: string; sheetName?: string; rowNumber?: number; object?: string; message: string; expected?: string | number; actual?: string | number };
type WorkbookReport = { fileName: string; createdAt: string; genericSheets: GenericSheet[]; productReport?: ProductReport };
type ProductRow = { sheetName: string; date: string; rowNumber: number; sourceOrder: number; category: string; name: string; price: number; purchase: number; morningStock: number; morningSold: number; morningAmount: number; nightStock: number; nightSold: number; nightAmount: number; totalAmount: number; endingStock: number; expectedNightStock: number; hasNightStock: boolean };
type ProductSheetPreview = { sheetName: string; date: string; tableCount: number; rowCount: number };
type ProductSummary = { category: string; name: string; price: number; purchase: number; morningSold: number; nightSold: number; totalSold: number; totalAmount: number; endingStock: number; latestDate: string; sourceOrder: number };
type ProductReport = { sheets: ProductSheetPreview[]; rows: ProductRow[]; issues: Issue[]; summary: ProductSummary[] };
type SummaryRow = Record<string, string | number>;

const MAIN_TABS: Array<[MainTab, string]> = [["smart", "智能上传"], ["product", "商品报表专用"], ["guide", "功能说明"]];
const REPORT_TABS: Array<[ReportTab, string]> = [["overview", "总览"], ["rules", "核对规则"], ["preview", "识别预览"], ["summary", "汇总报表"], ["issues", "核对异常"], ["auto", "自动结果"], ["export", "导出"]];

const ROLE_LABEL: Record<Role, string> = { date: "日期", product: "商品", employee: "员工", customer: "客户", supplier: "供应商", orderNo: "单号", sku: "SKU/货号", barcode: "条码", category: "分类", quantity: "数量", price: "单价", amount: "金额", totalAmount: "总金额", paidAmount: "实收金额", receivable: "应收", payable: "应付", stockOpening: "期初库存", stockIn: "入库", stockOut: "出库", stockEnding: "结存", purchase: "采购/进货", sold: "销量", basicSalary: "基本工资", attendance: "出勤", overtime: "加班", commission: "提成", bonus: "奖金/补贴", deduction: "扣款", advance: "借支", grossSalary: "应发工资", netSalary: "实发工资", feeType: "费用类型", paymentChannel: "收款渠道", department: "部门", position: "岗位", asset: "资产/设备", assetNo: "资产编号", status: "状态", remark: "备注" };
const ROLE_ALIASES: Record<Role, string[]> = { date: ["日期", "时间", "date", "day", "营业日"], product: ["商品", "商品名称", "售卖商品", "产品", "品名", "货品", "物品", "标题", "product", "item", "goods", "товар", "منتج"], employee: ["员工", "员工姓名", "姓名", "人员", "工号", "employee", "staff", "worker"], customer: ["客户", "客户名称", "会员", "买家", "customer", "client", "buyer"], supplier: ["供应商", "厂家", "供货商", "采购商", "supplier", "vendor"], orderNo: ["单号", "订单号", "流水号", "编号", "票号", "order", "orderid"], sku: ["sku", "货号", "编码", "商品编码", "артикул"], barcode: ["条码", "条形码", "barcode", "ean", "upc"], category: ["分类", "类别", "品类", "类型", "category", "type"], quantity: ["数量", "件数", "个数", "包数", "条数", "瓶数", "qty", "quantity", "count"], price: ["价格", "单价", "售价", "销售价", "price", "unitprice"], amount: ["金额", "小计", "销售额", "收入", "支出", "费用", "amount", "sum"], totalAmount: ["合计", "总计", "总金额", "总价", "合计金额", "total"], paidAmount: ["实收", "实收金额", "到账", "收款金额", "已收", "paid", "received"], receivable: ["应收", "应收金额", "应收款", "receivable"], payable: ["应付", "应付金额", "应付款", "payable"], stockOpening: ["期初", "初始库存", "早班库存", "上月结存", "opening"], stockIn: ["入库", "进货", "采购", "补货", "stockin", "inbound"], stockOut: ["出库", "销量", "售卖", "销售数量", "消耗", "stockout", "outbound"], stockEnding: ["结存", "库存", "剩余", "月末库存", "夜班库存", "stock", "ending"], purchase: ["进货", "采购", "补货", "purchase", "buy"], sold: ["销量", "售卖", "销售数量", "卖出", "sold", "salesqty"], basicSalary: ["基本工资", "底薪", "月薪", "工资标准", "salary", "base"], attendance: ["出勤", "出勤天数", "天数", "工时", "考勤", "attendance", "hours", "days"], overtime: ["加班", "加班费", "overtime"], commission: ["提成", "业绩提成", "销售提成", "commission"], bonus: ["奖金", "补贴", "满勤", "津贴", "奖励", "bonus", "allowance"], deduction: ["扣款", "罚款", "扣除", "社保", "个税", "deduction", "fine"], advance: ["借支", "预支", "借款", "advance", "loan"], grossSalary: ["应发", "应发工资", "工资合计", "gross"], netSalary: ["实发", "实发工资", "到手", "实际发放", "net"], feeType: ["费用类型", "费用项目", "支出项目", "科目", "expense", "feetype"], paymentChannel: ["渠道", "收款渠道", "支付方式", "付款方式", "平台", "channel", "paymentmethod"], department: ["部门", "门店", "组织", "department", "team"], position: ["岗位", "职位", "职务", "position", "role"], asset: ["资产", "设备", "固定资产", "asset", "device", "equipment"], assetNo: ["资产编号", "设备编号", "assetno", "serial"], status: ["状态", "进度", "status", "state"], remark: ["备注", "说明", "原因", "note", "remark", "comment"] };

function text(value: unknown) { return String(value ?? "").trim(); }
function norm(value: unknown) { return text(value).replace(/[\s\-_，,。.;；:：/\\|()（）\[\]【】{}<>《》"'“”‘’]+/g, "").toLowerCase(); }
function isBlank(value: unknown) { return value === null || value === undefined || text(value) === ""; }
function toNumber(value: unknown) { if (typeof value === "number") return Number.isFinite(value) ? value : 0; const cleaned = text(value).replace(/,/g, "").replace(/[￥¥元%\s]/g, ""); if (!cleaned) return 0; const parsed = Number(cleaned); return Number.isFinite(parsed) ? parsed : 0; }
function round2(value: number) { return Math.round((value + Number.EPSILON) * 100) / 100; }
function money(value: number) { return round2(value).toFixed(2); }
function colName(index: number) { return XLSX.utils.encode_col(index); }
function categoryOf(name: string) { const match = classifyGlobalProduct(name); return match.category === "烟草烟具" ? "香烟(酒店库存)" : match.category === "未识别" ? "未分类" : match.category; }
function productKey(category: string, name: string, price: number) { return `${category}__${name}__${price}`; }
function extractDate(sheetName: string, sheetIndex: number) { const day = /^\d{1,2}$/.test(sheetName) ? Number(sheetName) : sheetIndex + 1; const now = new Date(); return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`; }

function findProductGroups(grid: unknown[][]) {
  const groups: Array<{ headerRow: number; nameCol: number; priceCol: number; purchaseCol: number; morningStockCol: number; morningSoldCol: number; morningAmountCol: number; nightStockCol: number; nightSoldCol: number; nightAmountCol: number; totalAmountCol: number }> = [];
  grid.forEach((row, headerRow) => row.forEach((cell, col) => {
    const label = norm(cell);
    const next = norm(row[col + 1]);
    if ((label === "商品名称" || label === "售卖商品") && (next === "价格" || next === "售价")) {
      groups.push({ headerRow, nameCol: col, priceCol: col + 1, purchaseCol: col + 2, morningStockCol: col + 3, morningSoldCol: col + 4, morningAmountCol: col + 5, nightStockCol: col + 6, nightSoldCol: col + 7, nightAmountCol: col + 8, totalAmountCol: col + 9 });
    }
  }));
  return groups;
}
function parseProductReport(workbook: XLSX.WorkBook): ProductReport | undefined {
  const rows: ProductRow[] = [];
  const previews: ProductSheetPreview[] = [];
  workbook.SheetNames.forEach((sheetName, sheetIndex) => {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return;
    const grid = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, defval: null });
    const groups = findProductGroups(grid);
    const date = extractDate(sheetName, sheetIndex);
    let rowCount = 0;
    groups.forEach((group, groupIndex) => {
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
        rows.push({ sheetName, date, rowNumber: r + 1, sourceOrder: sheetIndex * 1_000_000 + groupIndex * 100_000 + r, category: categoryOf(name), name, price, purchase, morningStock, morningSold, morningAmount, nightStock, nightSold, nightAmount, totalAmount, endingStock, expectedNightStock, hasNightStock: !isBlank(nightStockRaw) });
        rowCount += 1;
      }
    });
    previews.push({ sheetName, date, tableCount: groups.length, rowCount });
  });
  if (rows.length < 10) return undefined;
  const issues: Issue[] = [];
  rows.forEach((row) => {
    if (row.hasNightStock && Math.abs(row.nightStock - row.expectedNightStock) > 0.01) issues.push({ severity: "error", type: "早班交夜班", sheetName: row.sheetName, rowNumber: row.rowNumber, object: row.name, message: "夜班库存应等于早班库存 + 进货 - 早班销量。", expected: row.expectedNightStock, actual: row.nightStock });
    if (Math.abs(row.morningAmount - row.morningSold * row.price) > 0.01) issues.push({ severity: "warning", type: "早班金额核对", sheetName: row.sheetName, rowNumber: row.rowNumber, object: row.name, message: "早班金额不等于早班销量 × 价格。", expected: round2(row.morningSold * row.price), actual: row.morningAmount });
    if (Math.abs(row.nightAmount - row.nightSold * row.price) > 0.01) issues.push({ severity: "warning", type: "夜班金额核对", sheetName: row.sheetName, rowNumber: row.rowNumber, object: row.name, message: "夜班金额不等于夜班销量 × 价格。", expected: round2(row.nightSold * row.price), actual: row.nightAmount });
  });
  const summaryMap = new Map<string, ProductSummary>();
  rows.forEach((row) => {
    const key = productKey(row.category, row.name, row.price);
    const item = summaryMap.get(key) ?? { category: row.category, name: row.name, price: row.price, purchase: 0, morningSold: 0, nightSold: 0, totalSold: 0, totalAmount: 0, endingStock: 0, latestDate: row.date, sourceOrder: row.sourceOrder };
    item.purchase += row.purchase; item.morningSold += row.morningSold; item.nightSold += row.nightSold; item.totalSold += row.morningSold + row.nightSold; item.totalAmount += row.totalAmount;
    if (row.date >= item.latestDate) { item.latestDate = row.date; item.endingStock = row.endingStock; }
    item.sourceOrder = Math.min(item.sourceOrder, row.sourceOrder);
    summaryMap.set(key, item);
  });
  const summary = Array.from(summaryMap.values()).map((item) => ({ ...item, purchase: round2(item.purchase), morningSold: round2(item.morningSold), nightSold: round2(item.nightSold), totalSold: round2(item.totalSold), totalAmount: round2(item.totalAmount), endingStock: round2(item.endingStock) })).sort((a, b) => a.category.localeCompare(b.category) || a.sourceOrder - b.sourceOrder);
  return { sheets: previews, rows, issues, summary };
}

function inferRole(header: string): { role?: Role; confidence: number } { const h = norm(header); let role: Role | undefined; let confidence = 0; (Object.entries(ROLE_ALIASES) as Array<[Role, string[]]>).forEach(([candidate, aliases]) => aliases.forEach((alias) => { const a = norm(alias); let score = 0; if (h === a) score = 1; else if (h.includes(a)) score = 0.86; else if (a.includes(h) && h.length >= 2) score = 0.55; if (score > confidence) { role = candidate; confidence = score; } })); return { role: confidence >= 0.5 ? role : undefined, confidence }; }
function detectGeneric(workbook: XLSX.WorkBook): GenericSheet[] { return workbook.SheetNames.map((sheetName): GenericSheet | null => { const worksheet = workbook.Sheets[sheetName]; if (!worksheet) return null; const grid = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, defval: null }); let headerRow = 0; let best = -Infinity; for (let i = 0; i < Math.min(35, grid.length); i += 1) { const row = grid[i] ?? []; const joined = norm(row.map(text).filter(Boolean).join(" ")); const score = row.filter((cell) => !isBlank(cell)).length + Object.values(ROLE_ALIASES).flat().filter((alias) => joined.includes(norm(alias))).length * 2; if (score > best) { best = score; headerRow = i; } } const used = new Set<string>(); const headers = (grid[headerRow] ?? []).map((cell, index) => { const base = text(cell) || `列${colName(index)}`; let name = base; let suffix = 2; while (used.has(name)) { name = `${base}_${suffix}`; suffix += 1; } used.add(name); return name; }); const dataRows: DataRow[] = []; for (let r = headerRow + 1; r < grid.length; r += 1) { const raw = grid[r] ?? []; if (!raw.some((cell) => !isBlank(cell))) continue; const values: Record<string, string | number> = {}; headers.forEach((header, index) => { values[header] = typeof raw[index] === "number" ? raw[index] as number : text(raw[index]); }); dataRows.push({ rowNumber: r + 1, values }); } const columns = headers.map((header, index): ColumnInfo => { const values = dataRows.map((row) => row.values[header]).filter((value) => !isBlank(value)); const guessed = inferRole(header); return { index, header, role: guessed.role, confidence: guessed.confidence, examples: values.slice(0, 3).map(text), nonEmpty: values.length }; }); const type: GenericType = columns.some((c) => c.role === "employee") ? "工资/考勤表" : columns.some((c) => c.role === "supplier") ? "进货采购表" : columns.some((c) => c.role === "paymentChannel") ? "收款对账表" : columns.some((c) => c.role === "product") ? "商品销售/库存表" : "通用业务表"; return { sheetName, type, confidence: 0.55, headerRow, rows: dataRows, columns, rules: ["字段识别、重复记录和基础金额检查"], issues: [] }; }).filter((sheet): sheet is GenericSheet => Boolean(sheet)); }
function analyzeWorkbook(workbook: XLSX.WorkBook, fileName: string): WorkbookReport { return { fileName, createdAt: new Date().toLocaleString("zh-CN", { hour12: false }), genericSheets: detectGeneric(workbook), productReport: parseProductReport(workbook) }; }
function exportWorkbookReport(report: WorkbookReport) { const workbook = XLSX.utils.book_new(); const product = report.productReport; XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([{ 文件: report.fileName, 类型: product ? "商品报表专用格式" : "通用业务表", 工作表: product?.sheets.length ?? report.genericSheets.length, 数据行: product?.rows.length ?? report.genericSheets.reduce((s, x) => s + x.rows.length, 0), 异常: product?.issues.length ?? 0 }]), "总览"); if (product) { XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(product.summary.map((item, index) => ({ 序号: index + 1, 分类: item.category, 商品: item.name, 价格: item.price, 进货: item.purchase, 早班销量: item.morningSold, 夜班销量: item.nightSold, 总销量: item.totalSold, 销售金额: item.totalAmount, 月末账面应剩: item.endingStock }))), "商品总汇"); XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(product.issues.map((item, index) => ({ 序号: index + 1, 类型: item.type, Sheet: item.sheetName, 行号: item.rowNumber, 商品: item.object, 说明: item.message, 应为: item.expected, 实际: item.actual }))), "核对异常"); } XLSX.writeFile(workbook, `${report.fileName.replace(/\.xlsx?$/i, "")}-智能核对报告.xlsx`); }
function Card({ title, value, desc, danger = false }: { title: string; value: string; desc?: string; danger?: boolean }) { return <div className={`rounded-2xl p-5 shadow-sm ${danger ? "bg-red-600 text-white" : "border bg-white text-slate-900"}`}><div className={`text-sm ${danger ? "text-red-100" : "text-slate-500"}`}>{title}</div><div className="mt-2 text-2xl font-bold">{value}</div>{desc ? <div className={`mt-1 text-xs ${danger ? "text-red-100" : "text-slate-500"}`}>{desc}</div> : null}</div>; }
function Badge({ children, tone = "gray" }: { children: ReactNode; tone?: "gray" | "green" | "red" | "yellow" }) { const cls = { gray: "bg-slate-100 text-slate-700", green: "bg-green-100 text-green-700", red: "bg-red-100 text-red-700", yellow: "bg-yellow-100 text-yellow-800" }[tone]; return <span className={`rounded-full px-2 py-1 text-xs font-medium ${cls}`}>{children}</span>; }
function DataTable({ rows, headers }: { rows: SummaryRow[]; headers: string[] }) { return <div className="table-scroll mt-3 max-h-96 overflow-auto"><table className="min-w-full text-sm"><thead className="sticky top-0 bg-slate-50 text-left"><tr>{headers.map((header) => <th key={header} className="p-3">{header}</th>)}</tr></thead><tbody>{rows.map((row, index) => <tr key={index} className="border-t">{headers.map((header) => <td key={header} className="p-3">{row[header]}</td>)}</tr>)}</tbody></table></div>; }
function Section({ title, desc, children }: { title: string; desc?: string; children: ReactNode }) { return <section className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">{title}</h2>{desc ? <p className="mt-1 text-sm text-slate-500">{desc}</p> : null}{children}</section>; }

export default function UniversalStructuredReportAppV2() {
  const [mainTab, setMainTab] = useState<MainTab>("smart");
  const [reportTab, setReportTab] = useState<ReportTab>("overview");
  const [report, setReport] = useState<WorkbookReport | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);
  async function handleFile(file: File) { const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true, cellFormula: true, cellNF: true, cellStyles: true }); setReport(analyzeWorkbook(workbook, file.name)); setReportTab("overview"); setMainTab("smart"); setActiveSheet(0); }
  const product = report?.productReport;
  const sheet = report?.genericSheets[activeSheet];
  const rowsCount = product?.rows.length ?? report?.genericSheets.reduce((sum, item) => sum + item.rows.length, 0) ?? 0;
  const issueCount = product?.issues.length ?? 0;
  return <main className="min-h-screen p-4 md:p-8"><div className="mx-auto max-w-7xl">
    <header className="mb-6 rounded-3xl bg-slate-950 p-6 text-white shadow-sm"><p className="text-sm text-slate-300">Structured Business Report Reconciliation</p><h1 className="mt-1 text-2xl font-bold md:text-4xl">智能表格自动核对</h1><p className="mt-2 max-w-4xl text-sm leading-6 text-slate-300">智能上传会优先识别你当前商品报表的专用规格：多块商品表、早班库存、夜班库存、早晚班销量和金额。</p></header>
    <nav className="mb-6 flex flex-wrap gap-2">{MAIN_TABS.map(([key, label]) => <button key={key} onClick={() => setMainTab(key)} className={`rounded-xl px-4 py-2 text-sm font-medium ${mainTab === key ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>{label}</button>)}</nav>
    {mainTab === "smart" ? <section className="space-y-6"><div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">导入业务表 Excel</h2><p className="mt-1 text-sm text-slate-500">上传后优先识别商品报表专用规格；如果不是商品报表，再按通用业务表处理。</p></div><label className="inline-flex cursor-pointer items-center justify-center rounded-xl bg-slate-950 px-4 py-3 text-sm font-medium text-white">选择 Excel<input className="hidden" type="file" accept=".xlsx,.xls" onChange={(event: ChangeEvent<HTMLInputElement>) => event.target.files?.[0] && handleFile(event.target.files[0])} /></label></div></div>{!report ? <Section title="等待导入表格" desc="支持当前商品报表规格：一张 Sheet 中多个商品区域，每个区域含商品名称/价格/进货/早班库存/早班售卖/夜班库存/夜班售卖/金额。"><div className="mt-3 grid gap-3 md:grid-cols-4">{REPORT_TABS.map(([, label]) => <div key={label} className="rounded-xl bg-slate-50 p-3 text-sm font-medium text-slate-700">{label}</div>)}</div></Section> : null}{report ? <><div className="grid gap-4 md:grid-cols-4"><Card title="文件" value={report.fileName} desc={report.createdAt} /><Card title="识别类型" value={product ? "商品报表专用格式" : (sheet?.type ?? "通用业务表")} desc={product ? "已识别多块商品表结构" : "通用字段识别"} /><Card title="数据行" value={`${rowsCount} 行`} /><Card title="异常" value={`${issueCount} 条`} danger={issueCount > 0} /></div><div className="rounded-2xl border bg-white p-5 shadow-sm">{product ? <div className="flex flex-wrap gap-2">{product.sheets.map((item) => <Badge key={item.sheetName}>{item.sheetName}：{item.tableCount} 个商品区 / {item.rowCount} 行</Badge>)}</div> : <div className="flex flex-wrap gap-2">{report.genericSheets.map((item, index) => <button key={item.sheetName} onClick={() => setActiveSheet(index)} className={`rounded-xl px-3 py-2 text-sm ${index === activeSheet ? "bg-slate-950 text-white" : "bg-slate-100 text-slate-700"}`}>{item.sheetName}</button>)}</div>}<div className="mt-4 flex flex-wrap gap-2">{REPORT_TABS.map(([key, label]) => <button key={key} onClick={() => setReportTab(key)} className={`rounded-xl px-4 py-2 text-sm font-medium ${reportTab === key ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>{label}</button>)}</div></div>{reportTab === "overview" ? <Section title="总览" desc={product ? "已按商品报表专用规格解析，不再只按普通表头识别。" : "按通用业务表结构展示识别结果。"}><div className="mt-4 grid gap-4 md:grid-cols-4"><Card title="工作表" value={`${product?.sheets.length ?? report.genericSheets.length} 个`} /><Card title="商品区/规则" value={product ? `${product.sheets.reduce((s, x) => s + x.tableCount, 0)} 个` : `${sheet?.rules.length ?? 0} 条`} /><Card title="商品品规" value={product ? `${product.summary.length} 个` : "-"} /><Card title="异常数量" value={`${issueCount} 条`} danger={issueCount > 0} /></div></Section> : null}{reportTab === "rules" ? <Section title="核对规则" desc="系统当前执行的核对规则。"><div className="mt-3 space-y-2">{(product ? ["识别同一 Sheet 中多个商品区域", "夜班库存 = 早班库存 + 进货 - 早班销量", "早班金额 = 早班销量 × 价格", "夜班金额 = 夜班销量 × 价格", "月末账面应剩 = 最新夜班库存 - 夜班销量"] : sheet?.rules ?? []).map((rule, index) => <div key={rule} className="rounded-xl bg-slate-50 p-3 text-sm"><b>{index + 1}.</b> {rule}</div>)}</div></Section> : null}{reportTab === "preview" ? <Section title="识别预览" desc="查看识别到的工作表、商品区域和字段。">{product ? <DataTable rows={product.sheets.map((item) => ({ Sheet: item.sheetName, 日期: item.date, 商品区域: item.tableCount, 商品行: item.rowCount }))} headers={["Sheet", "日期", "商品区域", "商品行"]} /> : <DataTable rows={(sheet?.columns ?? []).map((column) => ({ 列: colName(column.index), 表头: column.header, 字段: column.role ? ROLE_LABEL[column.role] : "未识别", 置信度: `${Math.round(column.confidence * 100)}%`, 示例: column.examples.join(" / ") }))} headers={["列", "表头", "字段", "置信度", "示例"]} />}</Section> : null}{reportTab === "summary" ? <Section title="汇总报表" desc="按商品报表规则生成商品总汇。">{product ? <DataTable rows={product.summary.map((item, index) => ({ 序号: index + 1, 分类: item.category, 商品: item.name, 价格: item.price, 进货: item.purchase, 早班销量: item.morningSold, 夜班销量: item.nightSold, 总销量: item.totalSold, 销售金额: `¥${money(item.totalAmount)}`, 月末账面应剩: item.endingStock }))} headers={["序号", "分类", "商品", "价格", "进货", "早班销量", "夜班销量", "总销量", "销售金额", "月末账面应剩"]} /> : <p className="mt-3 text-sm text-slate-600">通用表暂只输出识别预览和异常。</p>}</Section> : null}{reportTab === "issues" ? <Section title="核对异常" desc="商品库存和金额异常集中显示。">{product && product.issues.length ? <DataTable rows={product.issues.map((item, index) => ({ 序号: index + 1, 类型: item.type, Sheet: item.sheetName ?? "", 行号: item.rowNumber ?? "", 商品: item.object ?? "", 说明: item.message, 应为: item.expected ?? "", 实际: item.actual ?? "" }))} headers={["序号", "类型", "Sheet", "行号", "商品", "说明", "应为", "实际"]} /> : <p className="mt-3 text-sm text-green-700">当前没有识别到明显异常。</p>}</Section> : null}{reportTab === "auto" ? <Section title="自动结果" desc="自动生成账面库存盘点结果。">{product ? <DataTable rows={product.summary.map((item, index) => ({ 序号: index + 1, 分类: item.category, 商品: item.name, 最新日期: item.latestDate, 月末账面应剩: item.endingStock, 本月销量: item.totalSold, 状态: product.issues.some((issue) => issue.object === item.name) ? "需核对" : "正常" }))} headers={["序号", "分类", "商品", "最新日期", "月末账面应剩", "本月销量", "状态"]} /> : <p className="mt-3 text-sm text-slate-600">通用业务表没有专用库存盘点结果。</p>}</Section> : null}{reportTab === "export" ? <Section title="导出" desc="导出智能核对报告。"><button onClick={() => exportWorkbookReport(report)} className="mt-4 rounded-xl bg-green-700 px-4 py-3 text-sm font-medium text-white">导出智能核对报告</button></Section> : null}</> : null}</section> : null}
    {mainTab === "product" ? <section className="space-y-4"><Section title="商品报表专用核对" desc="这里保留原商品报表专用页面。"><ProductReportApp /></Section></section> : null}
    {mainTab === "guide" ? <Section title="功能说明" desc="当前智能上传已优先识别商品报表专用规格。"><div className="mt-4 grid gap-3 md:grid-cols-3">{["自动识别多个商品区域", "识别商品名称/售价/进货", "识别早班库存/早班销量", "识别夜班库存/夜班销量", "生成商品总汇", "生成自动库存盘点", "输出交接和金额异常", "可导出核对报告"].map((item) => <div key={item} className="rounded-xl bg-slate-50 p-3 text-sm text-slate-700">{item}</div>)}</div></Section> : null}
  </div></main>;
}
