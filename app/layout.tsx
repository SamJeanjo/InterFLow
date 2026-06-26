import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Catalog Validator",
  description: "Validate supplier catalog workbooks before submission."
};

export default function RootLayout({
  children
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
