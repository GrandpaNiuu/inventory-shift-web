"use client";

import { useEffect, useState } from "react";
import UniversalSmartReportAppV4 from "./UniversalSmartReportAppV4";
import TemplateLibraryTools from "./TemplateLibraryTools";

type MainTab = "smart" | "learn" | "guide";

const TEMPLATE_KEY = "smart_report_templates_v4";
const ALIAS_KEY = "smart_product_aliases_v4";

function readCount(key: string) {
  if (typeof window === "undefined") return 0;
  try {
    const raw = window.localStorage.getItem(key);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed.length : 0;
  } catch {
    return 0;
  }
}

export default function UniversalSmartReportAppV5() {
  const [tab, setTab] = useState<MainTab>("smart");
  const [templateCount, setTemplateCount] = useState(0);
  const [aliasCount, setAliasCount] = useState(0);

  function refreshCounts() {
    setTemplateCount(readCount(TEMPLATE_KEY));
    setAliasCount(readCount(ALIAS_KEY));
  }

  useEffect(() => {
    refreshCounts();
  }, [tab]);

  return <main className="min-h-screen p-4 md:p-8">
    <div className="mx-auto max-w-7xl">
      <header className="mb-6 rounded-3xl bg-slate-950 p-6 text-white shadow-sm">
        <p className="text-sm text-slate-300">Universal Smart Reconciliation Engine</p>
        <h1 className="mt-1 text-2xl font-bold md:text-4xl">智能表格自动核对</h1>
        <p className="mt-2 max-w-4xl text-sm leading-6 text-slate-300">一个入口上传任意表格。商品表自动获得商品报表功能，工资表自动获得工资核对功能，其他表进入通用结构化核对。</p>
      </header>

      <nav className="mb-6 flex flex-wrap gap-2">
        <button onClick={() => setTab("smart")} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === "smart" ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>智能上传</button>
        <button onClick={() => setTab("learn")} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === "learn" ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>学习库</button>
        <button onClick={() => setTab("guide")} className={`rounded-xl px-4 py-2 text-sm font-medium ${tab === "guide" ? "bg-slate-950 text-white" : "bg-white text-slate-700 shadow-sm hover:bg-slate-100"}`}>功能说明</button>
      </nav>
    </div>

    {tab === "smart" ? <section className="smart-upload-embedded">
      <style>{`.smart-upload-embedded > main { padding: 0 !important; min-height: 0 !important; } .smart-upload-embedded > main > div > header, .smart-upload-embedded > main > div > nav { display: none !important; }`}</style>
      <UniversalSmartReportAppV4 />
    </section> : null}

    {tab === "learn" ? <section className="space-y-6">
      <TemplateLibraryTools />
      <div className="mx-auto max-w-7xl px-4 md:px-8">
        <section className="rounded-2xl border bg-white p-5 shadow-sm">
          <h2 className="text-lg font-bold text-slate-900">模板学习</h2>
          <p className="mt-1 text-sm text-slate-500">保存字段映射后，下次上传相似表头会优先套用。当前已保存模板 {templateCount} 条。</p>
          <div className="mt-4 rounded-xl bg-slate-50 p-4 text-sm text-slate-700">
            模板不是在这里手动乱填。正确流程是：先到“智能上传”上传一张表 → 系统识别字段 → 在“核对规则”里保存当前识别为模板 → 回到这里导出备份或导入恢复。
          </div>
        </section>

        <section className="mt-6 rounded-2xl border bg-white p-5 shadow-sm">
          <h2 className="text-lg font-bold text-slate-900">商品别名库</h2>
          <p className="mt-1 text-sm text-slate-500">商品确认一次分类后，下次导入会优先识别。当前已保存商品别名 {aliasCount} 条。</p>
          <div className="mt-4 rounded-xl bg-slate-50 p-4 text-sm text-slate-700">
            商品别名可以通过商品表“自动结果”里的未分类商品加入，也可以通过导入模板库恢复之前保存的别名。
          </div>
        </section>
      </div>
    </section> : null}

    {tab === "guide" ? <section className="mx-auto max-w-7xl px-4 md:px-8">
      <div className="rounded-2xl border bg-white p-5 shadow-sm">
        <h2 className="text-lg font-bold text-slate-900">功能说明</h2>
        <p className="mt-1 text-sm text-slate-500">同一个上传入口，根据表格内容自动显示对应功能。</p>
        <div className="mt-4 grid gap-3 md:grid-cols-3">
          {[
            "商品表：每日、总汇、交接、库存、收款",
            "工资表：应发、实发、扣款、部门岗位汇总",
            "通用表：字段识别、模板学习、结构化导出",
            "模板学习：保存字段映射",
            "商品别名：保存商品分类经验",
            "导入导出：备份和恢复学习库"
          ].map((item) => <div key={item} className="rounded-xl bg-slate-50 p-3 text-sm text-slate-700">{item}</div>)}
        </div>
      </div>
    </section> : null}
  </main>;
}
