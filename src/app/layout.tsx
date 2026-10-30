import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "로또 6/45 분석 및 실시간 데이터",
  description: "매주 토요일 자동 갱신되는 동행복권 최신 당첨 데이터 분석",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="ko" className="h-full antialiased">
      <body className="min-h-full flex flex-col font-sans">{children}</body>
    </html>
  );
}
