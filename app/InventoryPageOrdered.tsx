"use client";

import { useEffect } from "react";
import InventoryPageConfigurable from "./InventoryPageConfigurable";

const MERGED_TO_DRINK_KEYWORDS = /打火机|火机|火柴|点烟器|方便面|泡面|螺蛳粉|零食|口香糖|杂项|其他|小商品|日用品|百货|纸巾|扑克牌|牙刷|牙膏|剃须|充电器|数据线|雨伞/;

function patchCategoryMatching() {
  const proto = String.prototype as String & { __inventoryMergedDrinkPatched?: boolean; __inventoryOriginalIncludes?: typeof String.prototype.includes };
  if (proto.__inventoryMergedDrinkPatched) return;
  const originalIncludes = String.prototype.includes;
  proto.__inventoryOriginalIncludes = originalIncludes;
  String.prototype.includes = function patchedIncludes(searchString: string, position?: number) {
    const source = String(this);
    const keyword = String(searchString);
    if (keyword === "饮料" && MERGED_TO_DRINK_KEYWORDS.test(source)) return true;
    return originalIncludes.call(source, searchString, position);
  };
  proto.__inventoryMergedDrinkPatched = true;
}

function hideMergedRuleRows() {
  const inputs = Array.from(document.querySelectorAll("input"));
  inputs.forEach((input) => {
    if (input.value === "打火机" || input.value === "杂项") {
      const row = input.closest(".rounded-xl.border.p-4") as HTMLElement | null;
      if (row) row.style.display = "none";
    }
  });
}

export default function InventoryPageOrdered() {
  patchCategoryMatching();

  useEffect(() => {
    hideMergedRuleRows();
    const observer = new MutationObserver(hideMergedRuleRows);
    observer.observe(document.body, { childList: true, subtree: true });
    return () => observer.disconnect();
  }, []);

  return <InventoryPageConfigurable />;
}
