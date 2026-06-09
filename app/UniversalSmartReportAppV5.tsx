"use client";

import { useState } from "react";
import UniversalSmartReportAppV6 from "./UniversalSmartReportAppV6";
import TemplateLibraryTools from "./TemplateLibraryTools";

type MainTab = "smart" | "learn" | "guide";

export default function UniversalSmartReportAppV5() {
  const [tab, setTab] = useState<MainTab>("smart");

  return <main className="min-h-screen p-4 md:p-8">
    <div className="mx-auto max-w-7xl">
      <header className="mb-6 rounded-3xl bg-slate-950 p-6 text-white shadow-sm">
        <p className="text-sm text-slate-300">Universal Smart Reconciliation Engine</p>
        <h1 className="mt-1 text-2xl font-bold md:text-4xl">智能表格自动核对</h1>
        <p className="mt-2 max-w-4xl text-sm leading-6 text-slate-300">一个入口上传一个或多个 Excel，系统自动识别商品表、工资表和通用业务表，并显示对应核对结果。</p>
      </header>

      <div className="mb-6 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
        <nav className="flex flex-wrap gap-2">
          <button onClick={() => setTab("smart")} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === "smart" ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>智能上传</button>
          <button onClick={() => setTab("learn")} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === "learn" ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>学习库</button>
          <button onClick={() => setTab("guide")} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === "guide" ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>功能说明</button>
        </nav>
      </div>
    </div>

    {tab === "smart" ? <section className="smart-upload-embedded">
      <style>{`.smart-upload-embedded > main { padding: 0 !important; min-height: 0 !important; }`}</style>
      <UniversalSmartReportAppV6 />
    </section> : null}

    {tab === "learn" ? <section className="space-y-6">
      <TemplateLibraryTools />
    </section> : null}

    {tab === "guide" ? <section className="mx-auto max-w-7xl px-4 md:px-8">
      <div className="rounded-2xl border bg-white p-5 shadow-sm">
        <h2 className="text-lg font-bold text-slate-900">功能说明</h2>
        <p className="mt-1 text-sm text-slate-500">同一个上传入口，根据文件数量和表格内容自动显示对应功能。</p>
        <div className="mt-4 grid gap-3 md:grid-cols-3">
          {[
            "一个 Excel：自动判断商品表、工资表或通用表",
            "多个 Excel：可混合上传商品表和工资表，分别出结果",
            "商品表：每日、总汇、交接、库存、收款渠道",
            "工资表：应发、实发、扣款、借支、社保、个税",
            "通用表：字段识别、金额汇总、结构化导出",
            "学习库：从 Excel 建模板、备份和恢复"
          ].map((item) => <div key={item} className="rounded-xl bg-slate-50 p-3 text-sm text-slate-700">{item}</div>)}
        </div>
      </div>
    </section> : null}
  </main>;
}
