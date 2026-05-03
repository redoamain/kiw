# 🚀 KIW — Production Planning & Monitoring System

Aplikasi untuk **monitoring produksi** dan **perencanaan kebutuhan material** di departemen **PPIC (Production Planning & Inventory Control)**.

---

## 📌 Overview

KIW membantu tim dalam:

* Memantau **Production Order (PO)** secara real-time
* Menghitung kebutuhan material berdasarkan **Bill of Materials (BOM)**
* Mengontrol ketersediaan stok
* Menghindari kekurangan material saat proses produksi

---

## 🎯 Fitur Utama

| Fitur                       | Deskripsi                                 |
| --------------------------- | ----------------------------------------- |
| 📦 **Monitor PO Aktif**     | Menampilkan semua PO yang belum selesai   |
| 🧮 **Perhitungan Material** | Otomatis hitung kebutuhan dari BOM        |
| 📊 **Cek Stok**             | Bandingkan stok tersedia dengan kebutuhan |
| 🔒 **Commit PO**            | Reserve stok untuk produksi               |
| 📥 **Export Laporan**       | Download laporan dalam format Excel       |

---

## 👥 Role Pengguna

| Role           | Akses                              |
| -------------- | ---------------------------------- |
| **PPIC**       | Input, commit, dan uncommit PO     |
| **Produksi**   | Melihat kebutuhan material         |
| **Gudang**     | Melihat stok yang sudah di-reserve |
| **Purchasing** | Melihat material yang perlu dibeli |

---

## 🔄 Alur Proses

```text
PO dibuat
   ↓
PPIC monitoring di dashboard
   ↓
Hitung kebutuhan material (BOM)
   ↓
Cek ketersediaan stok
   ↓
[Jika cukup] → Commit PO (reserve stok)
   ↓
Export laporan untuk produksi
   ↓
PO selesai → otomatis hilang dari monitoring
```

---

## 🚦 Status PO

| Status           | Deskripsi                           |
| ---------------- | ----------------------------------- |
| 🟡 **PO Aktif**  | Belum di-commit / masih berjalan    |
| 🔵 **Committed** | Stok sudah di-reserve               |
| ✅ **Completed**  | PO selesai (tidak ditampilkan lagi) |

---

## 📄 Struktur Laporan Export

File Excel terdiri dari beberapa sheet:

* **PO** → Data Production Order
* **BOM** → Struktur material
* **Per Departemen** → Kebutuhan tiap departemen
* **Rekap** → Ringkasan total kebutuhan
* **Keterangan** → Panduan membaca laporan

---

## ⚙️ Cara Menjalankan

```bash
# Install dependencies
pnpm install

# Jalankan development server
pnpm dev

# Akses aplikasi
http://localhost:3000
```

---

## 🧩 Teknologi

* Next.js
* Shadcn UI
* React
* TypeScript
* Redux Toolkit
* Tailwind CSS

---

## 📈 Tujuan Sistem

* Mengurangi kesalahan perhitungan material
* Meningkatkan efisiensi perencanaan produksi
* Memberikan visibilitas stok secara real-time
* Mempermudah koordinasi antar departemen

---

## 📝 Catatan

Sistem ini dirancang untuk membantu proses **Production Planning & Control** agar lebih cepat, akurat, dan terstruktur.
