"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

type Shift = "morning" | "evening";
type Tab = "dashboard" | "income" | "monthly" | "products" | "stock";

type ShiftIncome = {
  id: string;
  date: string;
  shift: Shift;
  cash: number;
  wechat: number;
  alipay: number;
  platform: number;
  refund: number;
  expense: number;
  expectedTotal: number;
  actualTotal: number;
  note: string;
  updatedAt: string;
};

type Product = {
  id: string;
  sku: string;
  name: string;
  category: string;
  spec: string;
  systemStock: number;
  warningStock: number;
};

type CountItem = {
  sku: string;
  name: string;
  systemStock: number;
  countedStock: number | null;
  difference: number | null;
};

type CountSession = {
  id: string;
  weekStart: string;
  countDate: string;
  status: "draft" | "confirmed";
  items: CountItem[];
  updatedAt: string;
};

type IncomeForm = {
  date: string;
  shift: Shift;
  cash: string;
  wechat: string;
  alipay: string;
  platform: string;
  refund: string;
  expense: string;
  actualTotal: string;
  note: string;
};

type ProductForm = {
  sku: string;
  name: string;
  category: string;
  spec: string;
  systemStock: string;
  warningStock: string;
};

const STORAGE_KEY = "inventory-shift-web-v1";
const shiftLabel: Record<Shift, string> = { morning: "早班", evening: "晚班" };

const incomeFields: Array<[keyof Omit<IncomeForm, "date" | "shift" | "note">, string]> = [
  ["cash", "现金"],
  ["wechat", "微信"],
  ["alipay", "支付宝"],
  ["platform", "平台收入"],
  ["refund", "退款"],
  ["expense", "支出"],
  ["actualTotal", "实收合计（不填=应收）"]
];

const productFields: Array<[keyof ProductForm, string]> = [
  ["sku", "SKU"],
  ["name", "商品名称"],
  ["category", "类目"],
  ["spec", "规格"],
  ["systemStock", "系统库存"],
  ["warningStock", "预警库存"]
];

function todayISO() {
  const d = new Date();
  d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
  return d.toISOString().slice(0, 10);
}

function thisMonth() {
  return todayISO().slice(0, 7);
}

function thisMonday() {
  const d = new Date();
  const day = d.getDay() || 7;
  d.setDate(d.getDate() - day + 1);
  d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
  return d.toISOString().slice(0, 10);
}

function uid() {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function round2(n: number) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

function money(n: number) {
  return round2(n).toFixed(2);
}

function parseMoney(value: unknown) {
  if (typeof value === "number") return round2(value);
  const text = String(value ?? "").replace(/,/g, "").replace(/[￥¥元\s]/g, "").trim();
  if (!text) return 0;
  const n = Number(text);
  return Number.isFinite(n) ? round2(n) : 0;
}

function normalizeHeader(text: string) {
  return text.toLowerCase().replace(/[\s_*：:（）()\-]/g, "");
}

function pick(row: Record<string, unknown>, names: string[]) {
  const wanted = names.map(normalizeHeader);
  const key = Object.keys(row).find((candidate) => {
    const normalized = normalizeHeader(candidate);
    return wanted.some((name) => normalized.includes(name) || name.includes(normalized));
  });
  return key ? row[key] : "";
}

function normalizeDate(value: unknown) {
  if (value instanceof Date) {
    value.setMinutes(value.getMinutes() - value.getTimezoneOffset());
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === "number") {
    const d = new Date(Math.round((value - 25569) * 86400 * 1000));
    d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
    return d.toISOString().slice(0, 10);
  }
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const full = raw.match(/(20\d{2}|19\d{2})[年/.-]?(\d{1,2})[月/.-]?(\d{1,2})/);
  if (full) return `${full[1]}-${full[2].padStart(2, "0")}-${full[3].padStart(2, "0")}`;
  const short = raw.match(/(\d{1,2})\s*月\s*(\d{1,2})\s*日?/);
  if (short) return `${new Date().getFullYear()}-${short[1].padStart(2, "0")}-${short[2].padStart(2, "0")}`;
  const parsed = new Date(raw.replace(/\//g, "-"));
  if (Number.isNaN(parsed.getTime())) return "";
  parsed.setMinutes(parsed.getMinutes() - parsed.getTimezoneOffset());
  return parsed.toISOString().slice(0, 10);
}

async function readSheet(file: File) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });
}

function getMonthDays(month: string) {
  const [year, monthNumber] = month.split("-").map(Number);
  const days = new Date(year, monthNumber, 0).getDate();
  return Array.from({ length: days }, (_, index) => `${year}-${String(monthNumber).padStart(2, "0")}-${String(index + 1).padStart(2, "0")}`);
}

function expectedTotal(input: Pick<ShiftIncome, "cash" | "wechat" | "alipay" | "platform" | "refund" | "expense">) {
  return round2(input.cash + input.wechat + input.alipay + input.platform - input.refund - input.expense);
}

function upsertIncome(list: ShiftIncome[], incoming: ShiftIncome[]) {
  const map = new Map(list.map((item) => [`${item.date}-${item.shift}`, item]));
  incoming.forEach((item) => map.set(`${item.date}-${item.shift}`, item));
  return Array.from(map.values()).sort((a, b) => `${b.date}-${b.shift}`.localeCompare(`${a.date}-${a.shift}`));
}

function Badge({ children, tone = "gray" }: { children: React.ReactNode; tone?: "gray" | "green" | "red" | "yellow" }) {
  const cls = {
    gray: "bg-slate-100 text-slate-700",
    green: "bg-green-100 text-green-700",
    red: "bg-red-100 text-red-700",
    yellow: "bg-yellow-100 text-yellow-800"
  }[tone];
  return <span className={`rounded-full px-2 py-1 text-xs font-medium ${cls}`}>{children}</span>;
}

function Card({ title, value, desc }: { title: string; value: string; desc?: string }) {
  return (
    <div className="rounded-2xl border bg-white p-5 shadow-sm">
      <div className="text-sm text-slate-500">{title}</div>
      <div className="mt-2 text-2xl font-bold text-slate-900">{value}</div>
      {desc ? <div className="mt-1 text-xs text-slate-500">{desc}</div> : null}
    </div>
  );
}

export default function Home() {
  const [tab, setTab] = useState<Tab>("dashboard");
  const [incomes, setIncomes] = useState<ShiftIncome[]>([]);
  const [products, setProducts] = useState<Product[]>([]);
  const [sessions, setSessions] = useState<CountSession[]>([]);
  const [month, setMonth] = useState(thisMonth());
  const [weekStart, setWeekStart] = useState(thisMonday());
  const [incomeForm, setIncomeForm] = useState<IncomeForm>({ date: todayISO(), shift: "morning", cash: "", wechat: "", alipay: "", platform: "", refund: "", expense: "", actualTotal: "", note: "" });
  const [productForm, setProductForm] = useState<ProductForm>({ sku: "", name: "", category: "", spec: "", systemStock: "", warningStock: "" });

  useEffect(() => {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    try {
      const data = JSON.parse(raw) as { incomes?: ShiftIncome[]; products?: Product[]; sessions?: CountSession[] };
      setIncomes(data.incomes ?? []);
      setProducts(data.products ?? []);
      setSessions(data.sessions ?? []);
    } catch {
      localStorage.removeItem(STORAGE_KEY);
    }
  }, []);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ incomes, products, sessions }));
  }, [incomes, products, sessions]);

  const monthlyRows = useMemo(() => {
    return getMonthDays(month).map((date) => {
      const morning = incomes.find((item) => item.date === date && item.shift === "morning");
      const evening = incomes.find((item) => item.date === date && item.shift === "evening");
      const morningDiff = morning ? round2(morning.actualTotal - morning.expectedTotal) : 0;
      const eveningDiff = evening ? round2(evening.actualTotal - evening.expectedTotal) : 0;
      const issues: string[] = [];
      if (!morning) issues.push("缺早班");
      if (!evening) issues.push("缺晚班");
      if (Math.abs(morningDiff) > 0.01) issues.push("早班不平");
      if (Math.abs(eveningDiff) > 0.01) issues.push("晚班不平");
      return {
        date,
        morningExpected: morning?.expectedTotal ?? 0,
        morningActual: morning?.actualTotal ?? 0,
        morningDiff,
        eveningExpected: evening?.expectedTotal ?? 0,
        eveningActual: evening?.actualTotal ?? 0,
        eveningDiff,
        dayExpected: round2((morning?.expectedTotal ?? 0) + (evening?.expectedTotal ?? 0)),
        dayActual: round2((morning?.actualTotal ?? 0) + (evening?.actualTotal ?? 0)),
        dayDiff: round2(morningDiff + eveningDiff),
        issues
      };
    });
  }, [incomes, month]);

  const monthlySummary = useMemo(() => {
    const rows = incomes.filter((item) => item.date.startsWith(month));
    const sum = (field: "expectedTotal" | "actualTotal", shift?: Shift) => round2(rows.filter((row) => !shift || row.shift === shift).reduce((acc, row) => acc + row[field], 0));
    return {
      morning: sum("expectedTotal", "morning"),
      evening: sum("expectedTotal", "evening"),
      expected: sum("expectedTotal"),
      actual: sum("actualTotal"),
      diff: round2(sum("actualTotal") - sum("expectedTotal")),
      abnormalDays: monthlyRows.filter((row) => row.issues.length > 0).length,
      missingDays: monthlyRows.filter((row) => row.issues.includes("缺早班") || row.issues.includes("缺晚班")).length
    };
  }, [incomes, month, monthlyRows]);

  const activeSession = sessions.find((session) => session.weekStart === weekStart && session.status === "draft");
  const latestSession = sessions[0];
  const lowStockCount = products.filter((item) => item.warningStock > 0 && item.systemStock <= item.warningStock).length;

  function saveIncome() {
    const base = {
      cash: parseMoney(incomeForm.cash),
      wechat: parseMoney(incomeForm.wechat),
      alipay: parseMoney(incomeForm.alipay),
      platform: parseMoney(incomeForm.platform),
      refund: parseMoney(incomeForm.refund),
      expense: parseMoney(incomeForm.expense)
    };
    const expected = expectedTotal(base);
    const actual = incomeForm.actualTotal.trim() ? parseMoney(incomeForm.actualTotal) : expected;
    const record: ShiftIncome = {
      id: `${incomeForm.date}-${incomeForm.shift}`,
      date: incomeForm.date,
      shift: incomeForm.shift,
      ...base,
      expectedTotal: expected,
      actualTotal: actual,
      note: incomeForm.note,
      updatedAt: new Date().toISOString()
    };
    setIncomes((old) => upsertIncome(old, [record]));
    setIncomeForm((old) => ({ ...old, cash: "", wechat: "", alipay: "", platform: "", refund: "", expense: "", actualTotal: "", note: "" }));
  }

  async function importIncome(file: File, shift: Shift) {
    const rows = await readSheet(file);
    const parsed: ShiftIncome[] = [];
    rows.forEach((row) => {
      const date = normalizeDate(pick(row, ["日期", "date", "账期", "营业日期"]));
      if (!date) return;
      const base = {
        cash: parseMoney(pick(row, ["现金", "cash"])),
        wechat: parseMoney(pick(row, ["微信", "wechat", "微信收入"])),
        alipay: parseMoney(pick(row, ["支付宝", "alipay", "支付宝收入"])),
        platform: parseMoney(pick(row, ["平台", "平台收入", "线上", "online"])),
        refund: parseMoney(pick(row, ["退款", "refund"])),
        expense: parseMoney(pick(row, ["支出", "费用", "expense"]))
      };
      const expected = expectedTotal(base);
      const actualRaw = pick(row, ["实收", "实际到账", "到账", "总收入", "合计", "收入合计", "actual"]);
      const actual = String(actualRaw ?? "").trim() ? parseMoney(actualRaw) : expected;
      parsed.push({ id: `${date}-${shift}`, date, shift, ...base, expectedTotal: expected, actualTotal: actual, note: String(pick(row, ["备注", "note"]) ?? ""), updatedAt: new Date().toISOString() });
    });
    setIncomes((old) => upsertIncome(old, parsed));
    alert(`已导入 ${parsed.length} 条${shiftLabel[shift]}记录`);
  }

  function exportMonthly() {
    const workbook = XLSX.utils.book_new();
    const summary = [
      { 项目: "早班合计", 金额: monthlySummary.morning },
      { 项目: "晚班合计", 金额: monthlySummary.evening },
      { 项目: "应收合计", 金额: monthlySummary.expected },
      { 项目: "实收合计", 金额: monthlySummary.actual },
      { 项目: "差异", 金额: monthlySummary.diff },
      { 项目: "异常天数", 金额: monthlySummary.abnormalDays }
    ];
    const daily = monthlyRows.map((row) => ({ 日期: row.date, 早班应收: row.morningExpected, 早班实收: row.morningActual, 早班差异: row.morningDiff, 晚班应收: row.eveningExpected, 晚班实收: row.eveningActual, 晚班差异: row.eveningDiff, 当日应收: row.dayExpected, 当日实收: row.dayActual, 当日差异: row.dayDiff, 核对状态: row.issues.length ? row.issues.join("、") : "正常" }));
    const detail = incomes.filter((item) => item.date.startsWith(month)).sort((a, b) => `${a.date}-${a.shift}`.localeCompare(`${b.date}-${b.shift}`)).map((item) => ({ 日期: item.date, 班次: shiftLabel[item.shift], 现金: item.cash, 微信: item.wechat, 支付宝: item.alipay, 平台收入: item.platform, 退款: item.refund, 支出: item.expense, 应收合计: item.expectedTotal, 实收合计: item.actualTotal, 差异: round2(item.actualTotal - item.expectedTotal), 备注: item.note }));
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summary), "月度汇总");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(daily), "每日合并核对");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(detail), "班次明细");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(daily.filter((row) => row.核对状态 !== "正常")), "异常记录");
    XLSX.writeFile(workbook, `${month}-早晚班合并对账表.xlsx`);
  }

  function saveProduct() {
    const sku = productForm.sku.trim();
    const name = productForm.name.trim();
    if (!sku || !name) return alert("SKU 和商品名称必须填写");
    const record: Product = { id: sku, sku, name, category: productForm.category.trim(), spec: productForm.spec.trim(), systemStock: parseMoney(productForm.systemStock), warningStock: parseMoney(productForm.warningStock) };
    setProducts((old) => [record, ...old.filter((item) => item.sku !== sku)]);
    setProductForm({ sku: "", name: "", category: "", spec: "", systemStock: "", warningStock: "" });
  }

  async function importProducts(file: File) {
    const rows = await readSheet(file);
    const parsed: Product[] = [];
    rows.forEach((row) => {
      const sku = String(pick(row, ["sku", "货号", "编码", "商品编码"]) ?? "").trim();
      const name = String(pick(row, ["商品名称", "名称", "name"]) ?? "").trim();
      if (!sku || !name) return;
      parsed.push({ id: sku, sku, name, category: String(pick(row, ["类目", "分类", "category"]) ?? ""), spec: String(pick(row, ["规格", "型号", "颜色尺码", "spec"]) ?? ""), systemStock: parseMoney(pick(row, ["系统库存", "当前库存", "库存", "stock"])), warningStock: parseMoney(pick(row, ["预警库存", "库存预警", "warning"])) });
    });
    setProducts((old) => {
      const map = new Map(old.map((item) => [item.sku, item]));
      parsed.forEach((item) => map.set(item.sku, item));
      return Array.from(map.values()).sort((a, b) => a.sku.localeCompare(b.sku));
    });
    alert(`已导入 ${parsed.length} 个商品`);
  }

  function createWeeklySession() {
    if (!products.length) return alert("请先添加或导入商品库存");
    const session: CountSession = {
      id: uid(),
      weekStart,
      countDate: todayISO(),
      status: "draft",
      items: products.map((product) => ({ sku: product.sku, name: product.name, systemStock: product.systemStock, countedStock: null, difference: null })),
      updatedAt: new Date().toISOString()
    };
    setSessions((old) => [session, ...old.filter((item) => !(item.weekStart === weekStart && item.status === "draft"))]);
  }

  function updateCount(sku: string, value: string) {
    setSessions((old) => old.map((session) => {
      if (session.id !== activeSession?.id) return session;
      return {
        ...session,
        updatedAt: new Date().toISOString(),
        items: session.items.map((item) => {
          if (item.sku !== sku) return item;
          const countedStock = value === "" ? null : parseMoney(value);
          return { ...item, countedStock, difference: countedStock === null ? null : round2(countedStock - item.systemStock) };
        })
      };
    }));
  }

  function confirmWeeklySession() {
    if (!activeSession) return;
    const uncounted = activeSession.items.filter((item) => item.countedStock === null).length;
    if (uncounted > 0 && !confirm(`还有 ${uncounted} 个商品未盘点，确定确认吗？`)) return;
    setSessions((old) => old.map((session) => session.id === activeSession.id ? { ...session, status: "confirmed", updatedAt: new Date().toISOString() } : session));
  }

  function exportWeeklyCount() {
    const session = activeSession ?? latestSession;
    if (!session) return alert("没有可导出的盘点批次");
    const rows = session.items.map((item) => ({ SKU: item.sku, 商品名称: item.name, 系统库存: item.systemStock, 实盘库存: item.countedStock ?? "", 差异: item.difference ?? "未盘点" }));
    const diff = rows.filter((row) => typeof row.差异 === "number" && row.差异 !== 0);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([{ 周开始日期: session.weekStart, 盘点日期: session.countDate, 状态: session.status === "draft" ? "草稿" : "已确认", 商品数: session.items.length, 差异商品数: diff.length }]), "盘点汇总");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), "SKU明细");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(diff), "差异商品");
    XLSX.writeFile(workbook, `${session.weekStart}-每周库存盘点表.xlsx`);
  }

  const nav: Array<[Tab, string]> = [["dashboard", "首页"], ["income", "账表导入/录入"], ["monthly", "月度合并核对"], ["products", "商品库存"], ["stock", "每周盘点"]];

  return (
    <main className="min-h-screen p-4 md:p-8">
      <div className="mx-auto max-w-7xl">
        <header className="mb-6 rounded-3xl bg-slate-900 p-6 text-white shadow-sm">
          <p className="text-sm text-slate-300">Inventory Shift Web</p>
          <h1 className="mt-1 text-2xl font-bold md:text-4xl">库存盘点与早晚班账表合并助手</h1>
          <p className="mt-2 max-w-3xl text-sm text-slate-300">先完成可上线测试的网页 MVP：早班/晚班账表导入合并、自动核对、月报导出、每周盘点清单和差异表。</p>
        </header>

        <nav className="mb-6 flex flex-wrap gap-2">
          {nav.map(([key, label]) => <button key={key} onClick={() => setTab(key)} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === key ? "bg-slate-900 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>{label}</button>)}
        </nav>

        {tab === "dashboard" && <section className="space-y-6">
          <div className="grid gap-4 md:grid-cols-4">
            <Card title="本月早班收入" value={`¥${money(monthlySummary.morning)}`} />
            <Card title="本月晚班收入" value={`¥${money(monthlySummary.evening)}`} />
            <Card title="本月实收合计" value={`¥${money(monthlySummary.actual)}`} desc={`差异 ¥${money(monthlySummary.diff)}`} />
            <Card title="异常天数" value={`${monthlySummary.abnormalDays} 天`} desc={`缺班次 ${monthlySummary.missingDays} 天`} />
          </div>
          <div className="grid gap-4 md:grid-cols-3">
            <Card title="商品数量" value={`${products.length} 个`} />
            <Card title="低库存预警" value={`${lowStockCount} 个`} />
            <Card title="最近盘点" value={latestSession ? latestSession.weekStart : "未盘点"} desc={latestSession ? (latestSession.status === "confirmed" ? "已确认" : "草稿") : "先生成盘点清单"} />
          </div>
          <div className="rounded-2xl border bg-white p-5 shadow-sm text-sm text-slate-700">
            <b>使用顺序：</b>先导入商品库存 → 导入早班/晚班账表 → 月度合并核对并导出 Excel → 每周生成盘点清单并录入实盘。
          </div>
        </section>}

        {tab === "income" && <section className="grid gap-6 lg:grid-cols-2">
          <div className="rounded-2xl border bg-white p-5 shadow-sm">
            <h2 className="text-lg font-bold">导入早班/晚班账表</h2>
            <p className="mt-1 text-sm text-slate-500">支持 xlsx/csv。建议表头：日期、现金、微信、支付宝、平台收入、退款、支出、实收、备注。</p>
            <div className="mt-4 grid gap-4 md:grid-cols-2">
              <label className="rounded-xl border border-dashed p-4 text-sm"><div className="mb-2 font-medium">导入早班表</div><input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && importIncome(e.target.files[0], "morning")} /></label>
              <label className="rounded-xl border border-dashed p-4 text-sm"><div className="mb-2 font-medium">导入晚班表</div><input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && importIncome(e.target.files[0], "evening")} /></label>
            </div>
          </div>
          <div className="rounded-2xl border bg-white p-5 shadow-sm">
            <h2 className="text-lg font-bold">手工录入/修正</h2>
            <div className="mt-4 grid gap-3 md:grid-cols-2">
              <label className="text-sm">日期<input className="mt-1 w-full rounded-xl border p-2" type="date" value={incomeForm.date} onChange={(e) => setIncomeForm({ ...incomeForm, date: e.target.value })} /></label>
              <label className="text-sm">班次<select className="mt-1 w-full rounded-xl border p-2" value={incomeForm.shift} onChange={(e) => setIncomeForm({ ...incomeForm, shift: e.target.value as Shift })}><option value="morning">早班</option><option value="evening">晚班</option></select></label>
              {incomeFields.map(([key, label]) => <label key={key} className="text-sm">{label}<input className="mt-1 w-full rounded-xl border p-2" inputMode="decimal" value={incomeForm[key]} onChange={(e) => setIncomeForm({ ...incomeForm, [key]: e.target.value })} /></label>)}
              <label className="text-sm md:col-span-2">备注<textarea className="mt-1 w-full rounded-xl border p-2" value={incomeForm.note} onChange={(e) => setIncomeForm({ ...incomeForm, note: e.target.value })} /></label>
            </div>
            <button onClick={saveIncome} className="mt-4 rounded-xl bg-slate-900 px-4 py-2 text-sm text-white">保存/覆盖该日期班次</button>
          </div>
          <div className="rounded-2xl border bg-white p-5 shadow-sm lg:col-span-2">
            <h2 className="text-lg font-bold">最近记录</h2>
            <div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["日期", "班次", "应收", "实收", "差异", "备注"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{incomes.slice(0, 30).map((item) => <tr key={`${item.date}-${item.shift}`} className="border-t"><td className="p-3">{item.date}</td><td className="p-3">{shiftLabel[item.shift]}</td><td className="p-3">¥{money(item.expectedTotal)}</td><td className="p-3">¥{money(item.actualTotal)}</td><td className="p-3">¥{money(item.actualTotal - item.expectedTotal)}</td><td className="p-3">{item.note}</td></tr>)}</tbody></table></div>
          </div>
        </section>}

        {tab === "monthly" && <section className="space-y-6">
          <div className="rounded-2xl border bg-white p-5 shadow-sm">
            <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">月度早晚班合并核对</h2><p className="mt-1 text-sm text-slate-500">按日期自动合并早班和晚班，检查缺班次、应收实收差异。</p></div><div className="flex gap-2"><input type="month" value={month} onChange={(e) => setMonth(e.target.value)} className="rounded-xl border p-2" /><button onClick={exportMonthly} className="rounded-xl bg-slate-900 px-4 py-2 text-sm text-white">导出 Excel</button></div></div>
            <div className="mt-4 grid gap-4 md:grid-cols-5"><Card title="早班" value={`¥${money(monthlySummary.morning)}`} /><Card title="晚班" value={`¥${money(monthlySummary.evening)}`} /><Card title="应收" value={`¥${money(monthlySummary.expected)}`} /><Card title="实收" value={`¥${money(monthlySummary.actual)}`} /><Card title="差异" value={`¥${money(monthlySummary.diff)}`} /></div>
          </div>
          <div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="table-scroll"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["日期", "早班应收", "早班实收", "晚班应收", "晚班实收", "当日实收", "差异", "状态"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{monthlyRows.map((row) => <tr key={row.date} className="border-t"><td className="p-3">{row.date}</td><td className="p-3">¥{money(row.morningExpected)}</td><td className="p-3">¥{money(row.morningActual)}</td><td className="p-3">¥{money(row.eveningExpected)}</td><td className="p-3">¥{money(row.eveningActual)}</td><td className="p-3">¥{money(row.dayActual)}</td><td className="p-3">¥{money(row.dayDiff)}</td><td className="p-3">{row.issues.length ? <Badge tone="red">{row.issues.join("、")}</Badge> : <Badge tone="green">正常</Badge>}</td></tr>)}</tbody></table></div></div>
        </section>}

        {tab === "products" && <section className="grid gap-6 lg:grid-cols-3">
          <div className="rounded-2xl border bg-white p-5 shadow-sm"><h2 className="text-lg font-bold">新增/覆盖商品</h2><div className="mt-4 space-y-3">{productFields.map(([key, label]) => <label key={key} className="block text-sm">{label}<input className="mt-1 w-full rounded-xl border p-2" value={productForm[key]} onChange={(e) => setProductForm({ ...productForm, [key]: e.target.value })} /></label>)}</div><button onClick={saveProduct} className="mt-4 rounded-xl bg-slate-900 px-4 py-2 text-sm text-white">保存商品</button></div>
          <div className="rounded-2xl border bg-white p-5 shadow-sm lg:col-span-2"><h2 className="text-lg font-bold">批量导入商品库存</h2><p className="mt-1 text-sm text-slate-500">表头建议：SKU、商品名称、类目、规格、系统库存、预警库存。</p><input className="mt-4" type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && importProducts(e.target.files[0])} /><div className="table-scroll mt-4"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["SKU", "商品", "类目", "规格", "系统库存", "预警"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{products.map((item) => <tr key={item.sku} className="border-t"><td className="p-3 font-mono">{item.sku}</td><td className="p-3">{item.name}</td><td className="p-3">{item.category}</td><td className="p-3">{item.spec}</td><td className="p-3">{item.systemStock}</td><td className="p-3">{item.warningStock > 0 && item.systemStock <= item.warningStock ? <Badge tone="yellow">低库存</Badge> : item.warningStock}</td></tr>)}</tbody></table></div></div>
        </section>}

        {tab === "stock" && <section className="space-y-6">
          <div className="rounded-2xl border bg-white p-5 shadow-sm"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><h2 className="text-lg font-bold">每周盘点自动化</h2><p className="mt-1 text-sm text-slate-500">一键生成 SKU 盘点清单，填实盘数，自动计算差异并导出。</p></div><div className="flex flex-wrap gap-2"><input type="date" value={weekStart} onChange={(e) => setWeekStart(e.target.value)} className="rounded-xl border p-2" /><button onClick={createWeeklySession} className="rounded-xl bg-slate-900 px-4 py-2 text-sm text-white">生成本周清单</button><button onClick={exportWeeklyCount} className="rounded-xl border px-4 py-2 text-sm">导出盘点表</button></div></div></div>
          <div className="rounded-2xl border bg-white p-5 shadow-sm">{!activeSession ? <p className="text-sm text-slate-500">还没有本周草稿。先导入商品，再生成本周盘点清单。</p> : <><div className="mb-4 flex items-center justify-between"><div><b>盘点周：</b>{activeSession.weekStart} <Badge>{activeSession.status === "draft" ? "草稿" : "已确认"}</Badge></div><button onClick={confirmWeeklySession} className="rounded-xl bg-green-700 px-4 py-2 text-sm text-white">确认本周盘点</button></div><div className="table-scroll"><table className="min-w-full text-sm"><thead className="bg-slate-50 text-left"><tr>{["SKU", "商品", "系统库存", "实盘库存", "差异", "状态"].map((h) => <th key={h} className="p-3">{h}</th>)}</tr></thead><tbody>{activeSession.items.map((item) => <tr key={item.sku} className="border-t"><td className="p-3 font-mono">{item.sku}</td><td className="p-3">{item.name}</td><td className="p-3">{item.systemStock}</td><td className="p-3"><input className="w-28 rounded-xl border p-2" inputMode="numeric" value={item.countedStock ?? ""} onChange={(e) => updateCount(item.sku, e.target.value)} /></td><td className="p-3">{item.difference === null ? "未盘" : item.difference}</td><td className="p-3">{item.difference === null ? <Badge>未盘</Badge> : item.difference === 0 ? <Badge tone="green">正常</Badge> : <Badge tone="red">有差异</Badge>}</td></tr>)}</tbody></table></div></>}</div>
        </section>}
      </div>
    </main>
  );
}
