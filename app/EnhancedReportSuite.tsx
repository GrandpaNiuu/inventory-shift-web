"use client";

import { useEffect, useState } from "react";
import UniversalSmartReportAppV5 from "./UniversalSmartReportAppV5";
import GlobalProductCoverageCenter from "./GlobalProductCoverageCenter";

type SuiteTab = "reconcile" | "settings";

const tabs: Array<[SuiteTab, string, string]> = [
  ["reconcile", "智能核对中心", "上传任意表格，自动显示对应核对功能"],
  ["settings", "设置", "商品识别库、商品词库和未识别商品维护"]
];

export default function EnhancedReportSuite() {
  const [tab, setTab] = useState<SuiteTab>("reconcile");

  useEffect(() => {
    document.title = "全球商品经营数据自动核对系统";
  }, []);

  return <main className="min-h-screen bg-slate-50">
    <div className="mx-auto max-w-7xl px-4 pt-4 md:px-8 md:pt-8">
      <header className="rounded-3xl bg-gradient-to-br from-slate-950 to-slate-800 p-6 text-white shadow-sm">
        <p className="text-sm text-slate-300">Global Product & Spreadsheet Reconciliation Suite</p>
        <h1 className="mt-1 text-2xl font-bold md:text-4xl">全球商品经营数据自动核对系统</h1>
        <p className="mt-2 max-w-4xl text-sm leading-6 text-slate-300">所有业务表统一从智能核对中心导入。商品表自动显示商品报表功能，工资表自动显示工资核对功能，其他表自动进入通用结构化核对。</p>
      </header>
      <nav className="mt-6 grid gap-3 md:grid-cols-2">
        {tabs.map(([key, label, desc]) => <button key={key} onClick={() => setTab(key)} className={`rounded-2xl border p-4 text-left shadow-sm transition ${tab === key ? "border-slate-950 bg-slate-950 text-white" : "bg-white text-slate-800 hover:bg-slate-100"}`}>
          <div className="font-bold">{label}</div>
          <div className={`mt-1 text-sm ${tab === key ? "text-slate-300" : "text-slate-500"}`}>{desc}</div>
        </button>)}
      </nav>
    </div>

    <div className="mt-6">
      {tab === "reconcile" ? <UniversalSmartReportAppV5 /> : null}
      {tab === "settings" ? <div className="mx-auto max-w-7xl px-4 pb-8 md:px-8"><GlobalProductCoverageCenter /></div> : null}
    </div>
  </main>;
}
