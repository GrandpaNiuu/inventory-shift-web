import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "库存盘点与班次对账助手",
  description: "早班晚班账表合并核对、月度报表导出、每周库存盘点自动化。"
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
