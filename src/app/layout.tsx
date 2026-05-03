import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";
import { Toaster } from "@/components/ui/sonner";
import StoreProvider from "./StroreProvider";
const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: "Kiw ✖",
  description: "data",
  icons: {
    icon: [
      {
        media: "(prefers-color-scheme: light)",
        url: "/icon2.png",
      },
      {
        media: "(prefers-color-scheme: dark)",
        url: "/icon2.png",
      },
    ],
  },
};


export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" suppressHydrationWarning>
      <body
        className={`${geistSans.variable} ${geistMono.variable} antialiased`} suppressHydrationWarning
      >

        <Toaster />
        <StoreProvider>{children}</StoreProvider>
      </body>
    </html>
  );
}
