import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "원가표준 변환기",
  description: "엑셀 원가 자료를 지정된 표준 양식으로 자동 변환합니다.",
  icons: {
    icon: "/favicon.svg",
    shortcut: "/favicon.svg",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ko">
      <body>{children}</body>
    </html>
  );
}
