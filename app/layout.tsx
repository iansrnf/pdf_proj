import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "PDF to Word",
  description: "Convert PDF files into editable Word documents.",
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
