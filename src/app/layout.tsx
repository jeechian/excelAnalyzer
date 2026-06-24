import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Excel Analyzer",
  description: "Upload, explore, and summarize Excel data",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
