import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "报表核对助手",
  description: "导入手工报表，自动生成每日汇总、月度汇总、交接异常、现场盘点和收款统计。"
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
