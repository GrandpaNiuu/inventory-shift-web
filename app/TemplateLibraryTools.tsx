"use client";

import { useEffect, useRef, useState, type ChangeEvent } from "react";

const TEMPLATE_KEY = "smart_report_templates_v4";
const ALIAS_KEY = "smart_product_aliases_v4";

type BackupPayload = {
  version: string;
  exportedAt: string;
  templates: unknown[];
  aliases: unknown[];
};

function safeReadList(key: string): unknown[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = window.localStorage.getItem(key);
    const value = raw ? JSON.parse(raw) : [];
    return Array.isArray(value) ? value : [];
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

export default function TemplateLibraryTools() {
  const [templateCount, setTemplateCount] = useState(0);
  const [aliasCount, setAliasCount] = useState(0);
  const inputRef = useRef<HTMLInputElement | null>(null);

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

  async function importLibrary(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const raw = await file.text();
      const parsed = JSON.parse(raw) as Partial<BackupPayload>;
      const importedTemplates = Array.isArray(parsed.templates) ? parsed.templates : [];
      const importedAliases = Array.isArray(parsed.aliases) ? parsed.aliases : [];
      if (!importedTemplates.length && !importedAliases.length) {
        alert("没有发现可导入的模板或商品别名。");
        return;
      }
      safeWriteList(TEMPLATE_KEY, mergeList(safeReadList(TEMPLATE_KEY), importedTemplates));
      safeWriteList(ALIAS_KEY, mergeList(safeReadList(ALIAS_KEY), importedAliases));
      refresh();
      alert(`导入完成：模板 ${importedTemplates.length} 条，商品别名 ${importedAliases.length} 条。页面将刷新以立即生效。`);
      window.location.reload();
    } catch {
      alert("导入失败：请使用系统导出的 JSON 模板库文件。");
    } finally {
      if (inputRef.current) inputRef.current.value = "";
    }
  }

  return <section className="mx-auto mb-4 max-w-7xl px-4 md:px-8">
    <div className="rounded-2xl border bg-white p-4 shadow-sm">
      <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
        <div>
          <h2 className="text-base font-bold text-slate-900">模板库导入 / 导出</h2>
          <p className="mt-1 text-sm text-slate-500">备份或恢复字段模板和商品别名库。当前模板 {templateCount} 条，商品别名 {aliasCount} 条。</p>
        </div>
        <div className="flex flex-wrap gap-2">
          <button onClick={exportLibrary} className="rounded-xl bg-slate-950 px-4 py-2 text-sm font-medium text-white">导出模板库</button>
          <button onClick={() => inputRef.current?.click()} className="rounded-xl border px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50">导入模板库</button>
          <input ref={inputRef} className="hidden" type="file" accept=".json" onChange={importLibrary} />
        </div>
      </div>
    </div>
  </section>;
}
