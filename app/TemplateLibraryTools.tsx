"use client";

import { useEffect, useRef, useState, type ChangeEvent } from "react";
import * as XLSX from "xlsx";

const TEMPLATE_KEY = "smart_report_templates_v4";
const ALIAS_KEY = "smart_product_aliases_v4";

type ReportKind = "product" | "payroll" | "generic";
type Role = "date" | "product" | "employee" | "department" | "position" | "quantity" | "price" | "amount" | "paidAmount" | "paymentChannel" | "supplier" | "customer" | "orderNo" | "baseSalary" | "expectedDays" | "attendanceDays" | "hourlyRate" | "workHours" | "overtimeHours" | "overtimePay" | "commission" | "bonus" | "allowance" | "deduction" | "advance" | "socialInsurance" | "tax" | "grossSalary" | "netSalary" | "remark";

type LearnedTemplate = {
  id: string;
  name: string;
  kind: ReportKind;
  createdAt: string;
  columns: Array<{ header: string; role?: Role; index: number }>;
};

type BackupPayload = {
  version: string;
  exportedAt: string;
  templates: unknown[];
  aliases: unknown[];
};

const ROLE_ALIASES: Record<Role, string[]> = {
  date: ["日期", "时间", "date", "day", "营业日"],
  product: ["商品", "商品名称", "售卖商品", "产品", "品名", "货品", "物品", "标题", "product", "item"],
  employee: ["员工", "员工姓名", "姓名", "人员", "工号", "employee", "staff"],
  department: ["部门", "门店", "班组", "组织", "department", "team"],
  position: ["岗位", "职位", "职务", "工种", "position", "role"],
  quantity: ["数量", "件数", "个数", "包数", "瓶数", "销量", "qty", "quantity"],
  price: ["价格", "单价", "售价", "销售价", "price"],
  amount: ["金额", "小计", "销售额", "收入", "支出", "费用", "合计", "total", "amount"],
  paidAmount: ["实收", "实收金额", "到账", "收款金额", "已收", "paid", "received"],
  paymentChannel: ["渠道", "收款渠道", "支付方式", "付款方式", "平台", "channel"],
  supplier: ["供应商", "厂家", "供货商", "supplier", "vendor"],
  customer: ["客户", "客户名称", "会员", "买家", "customer", "client"],
  orderNo: ["单号", "订单号", "流水号", "编号", "order"],
  baseSalary: ["基本工资", "底薪", "月薪", "工资标准", "固定工资", "base", "salary"],
  expectedDays: ["应出勤", "应出勤天数", "满勤天数", "标准天数"],
  attendanceDays: ["实际出勤", "出勤", "出勤天数", "上班天数", "考勤", "attendance"],
  hourlyRate: ["时薪", "小时工资", "每小时", "hourly"],
  workHours: ["工时", "工作小时", "小时", "workhours"],
  overtimeHours: ["加班小时", "加班工时", "overtimehours"],
  overtimePay: ["加班费", "加班工资", "overtimepay"],
  commission: ["提成", "业绩提成", "销售提成", "commission"],
  bonus: ["奖金", "满勤", "奖励", "绩效", "bonus"],
  allowance: ["补贴", "津贴", "餐补", "房补", "交通补贴", "allowance"],
  deduction: ["扣款", "罚款", "扣除", "缺勤扣款", "deduction"],
  advance: ["借支", "预支", "借款", "advance", "loan"],
  socialInsurance: ["社保", "五险", "保险", "social", "insurance"],
  tax: ["个税", "个人所得税", "税", "tax"],
  grossSalary: ["应发", "应发工资", "工资合计", "应付工资", "gross"],
  netSalary: ["实发", "实发工资", "到手", "实际发放", "net"],
  remark: ["备注", "说明", "原因", "note", "remark"]
};

function text(value: unknown) {
  return String(value ?? "").trim();
}

function norm(value: unknown) {
  return text(value).replace(/[\s\-_，,。.;；:：/\\|()（）\[\]【】{}<>《》"'“”‘’]+/g, "").toLowerCase();
}

function safeReadList<T = unknown>(key: string): T[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = window.localStorage.getItem(key);
    const value = raw ? JSON.parse(raw) : [];
    return Array.isArray(value) ? value as T[] : [];
  } catch {
    return [];
  }
}

function safeWriteList(key: string, value: unknown[]) {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(key, JSON.stringify(value));
}

function itemKey(item: unknown) {
  if (!item || typeof item !== "object") return "";
  const record = item as Record<string, unknown>;
  return String(record.id || record.name || record.alias || "");
}

function mergeList(oldList: unknown[], newList: unknown[]) {
  const result = [...oldList];
  const oldKeys = new Set(result.map(itemKey).filter(Boolean));
  newList.forEach((item) => {
    const key = itemKey(item);
    if (!key || oldKeys.has(key)) return;
    oldKeys.add(key);
    result.unshift(item);
  });
  return result;
}

function inferRole(header: unknown): Role | undefined {
  const h = norm(header);
  let best: { role?: Role; score: number } = { score: 0 };
  (Object.entries(ROLE_ALIASES) as Array<[Role, string[]]>).forEach(([role, aliases]) => {
    aliases.forEach((alias) => {
      const a = norm(alias);
      let score = 0;
      if (h === a) score = 1;
      else if (h.includes(a)) score = 0.86;
      else if (a.includes(h) && h.length >= 2) score = 0.55;
      if (score > best.score) best = { role, score };
    });
  });
  return best.score >= 0.5 ? best.role : undefined;
}

function scoreHeaderRow(row: unknown[]) {
  const joined = norm(row.map(text).filter(Boolean).join(" "));
  let score = row.filter((cell) => text(cell)).length;
  Object.values(ROLE_ALIASES).flat().forEach((alias) => {
    if (joined.includes(norm(alias))) score += 2;
  });
  if (/合计|总计|小计/.test(joined)) score -= 5;
  return score;
}

function detectHeaderRow(grid: unknown[][]) {
  let best = 0;
  let bestScore = -Infinity;
  for (let i = 0; i < Math.min(40, grid.length); i += 1) {
    const score = scoreHeaderRow(grid[i] ?? []);
    if (score > bestScore) {
      best = i;
      bestScore = score;
    }
  }
  return best;
}

function inferKind(columns: Array<{ role?: Role }>): ReportKind {
  const roles = columns.map((column) => column.role).filter(Boolean) as Role[];
  const productHits = roles.filter((role) => ["product", "price", "quantity", "amount"].includes(role)).length;
  const payrollHits = roles.filter((role) => ["employee", "baseSalary", "attendanceDays", "grossSalary", "netSalary", "deduction", "advance"].includes(role)).length;
  if (payrollHits >= 2) return "payroll";
  if (productHits >= 2) return "product";
  return "generic";
}

function buildTemplatesFromWorkbook(workbook: XLSX.WorkBook, fileName: string): LearnedTemplate[] {
  const now = new Date().toLocaleString("zh-CN", { hour12: false });
  const templates: LearnedTemplate[] = [];
  workbook.SheetNames.forEach((sheetName, sheetIndex) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;
    const grid = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null });
    if (!grid.length) return;
    const headerRow = detectHeaderRow(grid);
    const headers = grid[headerRow] ?? [];
    const columns = headers.map((header, index) => ({ header: text(header), role: inferRole(header), index })).filter((column) => column.header);
    if (columns.length < 2) return;
    const knownCount = columns.filter((column) => column.role).length;
    if (knownCount === 0) return;
    templates.push({
      id: `excel-template-${Date.now()}-${sheetIndex}`,
      name: `${fileName.replace(/\.(xlsx|xls)$/i, "")} - ${sheetName}`,
      kind: inferKind(columns),
      createdAt: now,
      columns
    });
  });
  return templates;
}

export default function TemplateLibraryTools() {
  const [templateCount, setTemplateCount] = useState(0);
  const [aliasCount, setAliasCount] = useState(0);
  const jsonInputRef = useRef<HTMLInputElement | null>(null);
  const excelInputRef = useRef<HTMLInputElement | null>(null);

  function refresh() {
    setTemplateCount(safeReadList(TEMPLATE_KEY).length);
    setAliasCount(safeReadList(ALIAS_KEY).length);
  }

  useEffect(() => {
    refresh();
  }, []);

  function exportLibrary() {
    const payload: BackupPayload = {
      version: "smart-template-library-v1",
      exportedAt: new Date().toLocaleString("zh-CN", { hour12: false }),
      templates: safeReadList(TEMPLATE_KEY),
      aliases: safeReadList(ALIAS_KEY)
    };
    const encoded = encodeURIComponent(JSON.stringify(payload, null, 2));
    const link = document.createElement("a");
    link.setAttribute("href", `data:text/json;charset=utf-8,${encoded}`);
    link.setAttribute("download", `智能核对模板库-${new Date().toISOString().slice(0, 10)}.json`);
    link.click();
  }

  async function createTemplateFromExcel(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      alert("从 Excel 创建模板只支持 .xlsx 或 .xls 文件。");
      if (excelInputRef.current) excelInputRef.current.value = "";
      return;
    }
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
      const templates = buildTemplatesFromWorkbook(workbook, file.name);
      if (!templates.length) {
        alert("没有从这个 Excel 中识别到可保存的模板。请确认表格有清晰的表头，例如：姓名、商品名称、金额、实发工资等。");
        return;
      }
      safeWriteList(TEMPLATE_KEY, mergeList(safeReadList(TEMPLATE_KEY), templates));
      refresh();
      alert(`已从 Excel 创建 ${templates.length} 个模板。后续同类表格会优先套用。`);
    } catch {
      alert("创建模板失败。请确认这是正常的 Excel 文件。");
    } finally {
      if (excelInputRef.current) excelInputRef.current.value = "";
    }
  }

  async function importLibrary(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".json")) {
      alert("恢复学习库只支持系统备份出来的 JSON 文件。Excel 表格请点“从 Excel 创建模板”或回到“智能上传”导入。");
      if (jsonInputRef.current) jsonInputRef.current.value = "";
      return;
    }
    try {
      const raw = await file.text();
      const parsed = JSON.parse(raw) as Partial<BackupPayload>;
      const importedTemplates = Array.isArray(parsed.templates) ? parsed.templates : [];
      const importedAliases = Array.isArray(parsed.aliases) ? parsed.aliases : [];
      if (!importedTemplates.length && !importedAliases.length) {
        alert("没有发现可恢复的模板或商品别名。请确认这是系统备份出来的 JSON 学习库文件。");
        return;
      }
      safeWriteList(TEMPLATE_KEY, mergeList(safeReadList(TEMPLATE_KEY), importedTemplates));
      safeWriteList(ALIAS_KEY, mergeList(safeReadList(ALIAS_KEY), importedAliases));
      refresh();
      alert(`恢复完成：模板 ${importedTemplates.length} 条，商品别名 ${importedAliases.length} 条。页面将刷新以立即生效。`);
      window.location.reload();
    } catch {
      alert("恢复失败：请使用系统备份出来的 JSON 学习库文件。");
    } finally {
      if (jsonInputRef.current) jsonInputRef.current.value = "";
    }
  }

  return <section className="mx-auto mb-4 max-w-7xl px-4 md:px-8">
    <div className="rounded-2xl border bg-white p-4 shadow-sm">
      <div className="flex flex-col gap-4">
        <div>
          <h2 className="text-base font-bold text-slate-900">学习库</h2>
          <p className="mt-1 text-sm text-slate-500">不会写 JSON 也没关系。你可以直接从 Excel 创建模板，也可以把学习库备份成 JSON，换设备后再恢复。当前模板 {templateCount} 条，商品别名 {aliasCount} 条。</p>
        </div>
        <div className="grid gap-3 md:grid-cols-3">
          <button onClick={() => excelInputRef.current?.click()} className="rounded-xl bg-slate-950 px-4 py-3 text-sm font-medium text-white">从 Excel 创建模板</button>
          <button onClick={exportLibrary} className="rounded-xl border px-4 py-3 text-sm font-medium text-slate-700 hover:bg-slate-50">备份学习库 JSON</button>
          <button onClick={() => jsonInputRef.current?.click()} className="rounded-xl border px-4 py-3 text-sm font-medium text-slate-700 hover:bg-slate-50">恢复学习库 JSON</button>
          <input ref={excelInputRef} className="hidden" type="file" accept=".xlsx,.xls" onChange={createTemplateFromExcel} />
          <input ref={jsonInputRef} className="hidden" type="file" accept=".json,application/json" onChange={importLibrary} />
        </div>
        <div className="rounded-xl bg-slate-50 p-3 text-sm text-slate-600">
          <b>怎么用：</b>第一次拿到某种表格，点“从 Excel 创建模板”；以后换设备或怕丢失，点“备份学习库 JSON”；有备份文件时，点“恢复学习库 JSON”。
        </div>
      </div>
    </div>
  </section>;
}
