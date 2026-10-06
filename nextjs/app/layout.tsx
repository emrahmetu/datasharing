import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Soda Grubu Hazine Yönetim Sistemi",
  description: "Konsolide nakit yönetimi",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="tr">
      <body>{children}</body>
    </html>
  );
}
