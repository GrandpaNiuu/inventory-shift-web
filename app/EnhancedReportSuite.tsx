"use client";

import { useEffect, useState } from "react";
import UniversalStructuredReportAppV3 from "./UniversalStructuredReportAppV3";
import GlobalProductCoverageCenter from "./GlobalProductCoverageCenter";

type SuiteTab = "reconcile" | "settings";

const tabs: Array<[SuiteTab, string, string]> = [
  ["reconcile", "智能核对中心", "支持模板学习、商品别名学习和结构化核对结果"],
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
        <p className="mt-2 max-w-4xl text-sm leading-6 text-slate-300">支持商品报表规格识别、结构化核对结果、当前商品报表专用核对、模板学习和商品别名库。确认一次字段或商品归类后，后续导入会优先套用。</p>
      </header>
      <nav className="mt-6 grid gap-3 md:grid-cols-2">
        {tabs.map(([key, label, desc]) => <button key={key} onClick={() => setTab(key)} className={`rounded-2xl border p-4 text-left shadow-sm transition ${tab === key ? "border-slate-950 bg-slate-950 text-white" : "bg-white text-slate-800 hover:bg-slate-100"}`}>
          <div className="font-bold">{label}</div>
          <div className={`mt-1 text-sm ${tab === key ? "text-slate-300" : "text-slate-500"}`}>{desc}</div>
        </button>)}
      </nav>
    </div>

    <div className="mt-6">
      {tab === "reconcile" ? <UniversalStructuredReportAppV3 /> : null}
      {tab === "settings" ? <div className="mx-auto max-w-7xl px-4 pb-8 md:px-8"><GlobalProductCoverageCenter /></div> : null}
    </div>
  </main>;
}
