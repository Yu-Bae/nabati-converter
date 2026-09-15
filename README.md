# Nabati Converter 🚀

Aplikasi Web Konverter Data Penjualan Sales Import resmi untuk distributor Nabati.

---

## 📖 Panduan Penggunaan Langkah demi Langkah

### 1. Pilih Branch Target Import
- Pilih **Branch Target** pada menu dropdown:
  - **Air Molek** (Branch ID: `1764907397498`)
  - **Ujung Batu** (Branch ID: `1764907108089`)
- *Company ID* otomatis diatur ke `NS6081030004380`.

### 2. Upload / Masukkan File Penjualan
- **Cara 1**: Drag & Drop (Tarik dan lepas) file Excel ke area dropzone.
- **Cara 2**: Klik **Pilih File Excel** untuk mengambil file `.xlsx` / `.xls` dari komputer.
- **Cara 3 (Simulasi)**: Klik **Coba Demo Data** untuk mencoba sistem secara instan.

> 💡 *Sistem otomatis mendeteksi file format **Sebelum di Olah** (Raw Sales) maupun **Sesudah di Olah**.*

### 3. Proses Konversi Data
- Klik tombol **Proses Data**.
- Sistem akan memproses secara otomatis:
  1. Formatting ID & Tanggal (`yyyy-mm-dd`).
  2. Pemecahan kemasan (`SATUAN/ISI` → `SATU`, `AN/`, `ISI`).
  3. Kalkulasi `HARGA PCS` = `JMLBRUTTO / QUANTITY` dan `HARGA CTN` = `HARGA PCS × AN/`.
  4. Konversi otomatis diskon nominal < 100 ke % `(disc / sellingPrice) * 100`.
  5. Penyusunan 43 kolom template import standar Nabati.

### 4. Tinjau Ringkasan & Preview
- Periksa ringkasan pada **Stats Dashboard** (Jenis Input, Jumlah Baris, Branch Target, Diskon Dikonversi).
- Tinjau 10 baris pertama di tabel **Preview Ready-Import Data**. Anda bisa menggunakan kotak pencarian untuk meretrospeksi transaksi tertentu.

### 5. Download Hasil Konversi
Pilih format file download yang dibutuhkan:
- 🟢 **Hasil Olah (.xlsx)**: File olahan sales lengkap dengan rumus Excel (`Q/O` & `Y*M`).
- ⬛ **Import File (.xlsx)**: Template import resmi 43 kolom format Excel.
- ⚪ **CSV Import (.csv)**: Format CSV berpemisah titik koma `;`.
- ⚪ **TXT Import (.txt)**: Format TXT berpemisah pipe `|`.

---

## 🔒 Keamanan & Kerahasiaan Data
100% diproses langsung di dalam browser pengguna (*client-side*). Data penjualan Anda tidak pernah dikirim ke server mana pun.
