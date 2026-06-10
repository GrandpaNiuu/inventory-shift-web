"use client";

import { useState } from "react";
import MultiFileSmartUploadPanel from "./MultiFileSmartUploadPanel";
import InventoryPageConfigurable from "./InventoryPageConfigurable";
import PayrollReportApp from "./PayrollReportApp";

type Mode = "auto" | "product" | "payroll";

export default function UniversalSmartReportAppV6() {
  const [mode, setMode] = useState<Mode>("product");

  return <section className="space-y-5">
    <div className="mx-auto max-w-7xl px-4 md:px-8">
      <div className="rounded-2xl border bg-white p-4 shadow-sm">
        <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
          <div>
            <h2 className="text-base font-bold text-slate-900">智能核对入口</h2>
            <p className="mt-1 text-sm text-slate-500">商品详细功能已恢复原自动盘点模块，只修复收款重复统计。</p>
          </div>
          <div className="flex flex-wrap gap-2">
            <button onClick={() => setMode("product")} className={`rounded-xl px-4 py-2 text-sm font-medium ${mode === "product" ? "bg-slate-950 text-white" : "bg-slate-100 text-slate-700"}`}>商品详细功能</button>
            <button onClick={() => setMode("auto")} className={`rounded-xl px-4 py-2 text-sm font-medium ${mode === "auto" ? "bg-slate-950 text-white" : "bg-slate-100 text-slate-700"}`}>自动识别上传</button>
            <button onClick={() => setMode("payroll")} className={`rounded-xl px-4 py-2 text-sm font-medium ${mode === "payroll" ? "bg-slate-950 text-white" : "bg-slate-100 text-slate-700"}`}>工资详细功能</button>
          </div>
        </div>
      </div>
    </div>
    {mode === "product" ? <InventoryPageConfigurable /> : null}
    {mode === "auto" ? <MultiFileSmartUploadPanel /> : null}
    {mode === "payroll" ? <PayrollReportApp /> : null}
  </section>;
}
