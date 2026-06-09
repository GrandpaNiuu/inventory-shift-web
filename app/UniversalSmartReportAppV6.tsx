"use client";

import { useState } from "react";
import InventoryPageConfigurable from "./InventoryPageConfigurable";
import PayrollReportApp from "./PayrollReportApp";

type SingleMode = "product" | "payroll";

export default function UniversalSmartReportAppV6() {
  const [mode, setMode] = useState<SingleMode>("product");

  return <section className="space-y-4">
    <div className="mx-auto max-w-7xl px-4 md:px-8">
      <div className="rounded-2xl border bg-white p-4 shadow-sm">
        <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
          <div>
            <h2 className="text-base font-bold text-slate-900">单表详细核对</h2>
            <p className="mt-1 text-sm text-slate-500">商品表走旧商品模块算法；工资表走完整工资核对算法。</p>
          </div>
          <div className="flex flex-wrap gap-2">
            <button onClick={() => setMode("product")} className={`rounded-xl px-4 py-2 text-sm font-medium ${mode === "product" ? "bg-slate-950 text-white" : "bg-slate-100 text-slate-700"}`}>商品表核对</button>
            <button onClick={() => setMode("payroll")} className={`rounded-xl px-4 py-2 text-sm font-medium ${mode === "payroll" ? "bg-slate-950 text-white" : "bg-slate-100 text-slate-700"}`}>工资表核对</button>
          </div>
        </div>
      </div>
    </div>
    {mode === "product" ? <InventoryPageConfigurable /> : <PayrollReportApp />}
  </section>;
}
