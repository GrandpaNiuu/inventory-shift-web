"use client";

import { useState } from "react";
import UniversalSmartReportAppV4 from "./UniversalSmartReportAppV4";
import MultiFileSmartUploadPanel from "./MultiFileSmartUploadPanel";
import TemplateLibraryTools from "./TemplateLibraryTools";

type MainTab = "smart" | "learn" | "guide";

export default function UniversalSmartReportAppV5() {
  const [tab, setTab] = useState<MainTab>("smart");

  return <main className="min-h-screen p-4 md:p-8">
    <div className="mx-auto max-w-7xl">
      <header className="mb-6 rounded-3xl bg-slate-950 p-6 text-white shadow-sm">
        <p className="text-sm text-slate-300">Universal Smart Reconciliation Engine</p>
        <h1 className="mt-1 text-2xl font-bold md:text-4xl">智能表格自动核对</h1>
        <p className="mt-2 max-w-4xl text-sm leading-6 text-slate-300">一个入口上传任意表格。选择多个 Excel 时自动生成多门店总览；选择单个 Excel 时继续做详细核对。</p>
      </header>

      <nav className="mb-6 flex flex-wrap gap-2">
        <button onClick={() => setTab("smart")} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === "smart" ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>智能上传</button>
        <button onClick={() => setTab("learn")} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === "learn" ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>学习库</button>
        <button onClick={() => setTab("guide")} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === "guide" ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>功能说明</button>
      </nav>
    </div>

    {tab === "smart" ? <section className="space-y-6">
      <MultiFileSmartUploadPanel />
      <section className="smart-upload-embedded">
        <style>{`.smart-upload-embedded > main { padding: 0 !important; min-height: 0 !important; } .smart-upload-embedded > main > div > header, .smart-upload-embedded > main > div > nav { display: none !important; }`}</style>
        <UniversalSmartReportAppV4 />
      </section>
    </section> : null}

    {tab === "learn" ? <section className="space-y-6">
      <TemplateLibraryTools />
    </section> : null}

    {tab === "guide" ? <section className="mx-auto max-w-7xl px-4 md:px-8">
      <div className="rounded-2xl border bg-white p-5 shadow-sm">
        <h2 className="text-lg font-bold text-slate-900">功能说明</h2>
        <p className="mt-1 text-sm text-slate-500">同一个智能上传入口，根据选择文件数量和表格内容自动显示对应功能。</p>
        <div className="mt-4 grid gap-3 md:grid-cols-3">
          {[
            "单个 Excel：详细核对和模板学习",
            "多个 Excel：多门店批量总览和总报告",
            "商品表：每日、总汇、交接、库存、收款",
            "工资表：应发、实发、扣款、部门岗位汇总",
            "通用表：字段识别、模板学习、结构化导出",
            "学习库：从 Excel 建模板、备份和恢复"
          ].map((item) => <div key={item} className="rounded-xl bg-slate-50 p-3 text-sm text-slate-700">{item}</div>)}
        </div>
      </div>
    </section> : null}
  </main>;
}
