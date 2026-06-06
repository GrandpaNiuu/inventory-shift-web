import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "烟报自动核对助手",
  description: "导入手工烟报，自动生成每日汇总、月度汇总、交接异常和每周盘点表。"
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
