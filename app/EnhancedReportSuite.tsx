"use client";

import { useEffect, useState } from "react";
import UniversalStructuredReportApp from "./UniversalStructuredReportApp";
import GlobalProductCoverageCenter from "./GlobalProductCoverageCenter";

type SuiteTab = "universal" | "products";

const tabs: Array<[SuiteTab, string, string]> = [
  ["universal", "万能核对中心", "任意 Excel 自动生成类似商品报表的结构化核对结果"],
  ["products", "全球商品覆盖中心", "识别各种商品分类、覆盖率和未识别商品"]
];

export default function EnhancedReportSuite() {
  const [tab, setTab] = useState<SuiteTab>("universal");

  useEffect(() => {
    document.title = "全球商品经营数据自动核对系统";
  }, []);

  return <main className="min-h-screen bg-slate-50">
    <div className="mx-auto max-w-7xl px-4 pt-4 md:px-8 md:pt-8">
      <header className="rounded-3xl bg-gradient-to-br from-slate-950 to-slate-800 p-6 text-white shadow-sm">
        <p className="text-sm text-slate-300">Global Product & Spreadsheet Reconciliation Suite</p>
        <h1 className="mt-1 text-2xl font-bold md:text-4xl">全球商品经营数据自动核对系统</h1>
        <p className="mt-2 max-w-4xl text-sm leading-6 text-slate-300">支持任意 Excel 智能识别、结构化核对结果、当前商品报表专用核对，以及更广范围的全球商品分类覆盖。系统会逐步通过模板和商品词库学习，提高后续自动化准确率。</p>
      </header>
      <nav className="mt-6 grid gap-3 md:grid-cols-2">
        {tabs.map(([key, label, desc]) => <button key={key} onClick={() => setTab(key)} className={`rounded-2xl border p-4 text-left shadow-sm transition ${tab === key ? "border-slate-950 bg-slate-950 text-white" : "bg-white text-slate-800 hover:bg-slate-100"}`}>
          <div className="font-bold">{label}</div>
          <div className={`mt-1 text-sm ${tab === key ? "text-slate-300" : "text-slate-500"}`}>{desc}</div>
        </button>)}
      </nav>
    </div>

    <div className="mt-6">
      {tab === "universal" ? <UniversalStructuredReportApp /> : null}
      {tab === "products" ? <div className="mx-auto max-w-7xl px-4 pb-8 md:px-8"><GlobalProductCoverageCenter /></div> : null}
    </div>
  </main>;
}
