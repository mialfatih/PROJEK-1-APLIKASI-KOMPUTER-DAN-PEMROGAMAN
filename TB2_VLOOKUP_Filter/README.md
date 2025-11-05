# TB2 — VLOOKUP, Diskon (berdasarkan NIM), Sorting & Filtering

Tugas ini berfokus pada penggunaan **VLOOKUP**, penerapan **diskon berdasarkan digit terakhir NIM**, serta proses **sort & filter** untuk menghasilkan dua sheet laporan.

---

## ✅ Tujuan Pembelajaran
- Mengambil data **Negara**, **Produk**, **Segmen**, dan **Harga** menggunakan `VLOOKUP`
- Memilih tabel harga berdasarkan **ganjil / genap** digit terakhir NIM
- Menghitung nilai **Penjualan**
- Melakukan **filtering** data karyawan sesuai kriteria
- Membuat worksheet baru berdasarkan hasil filter

---

## 📂 Data & Template yang digunakan
| File | Fungsi |
|------|--------|
| `Template_VLookup_HLookup.xlsx` / `TBI1_KELAS_NIM_NAMA.xlsx` | Template tugas |
| `Employee Data.txt` | Dataset untuk bagian filtering |

---

## 📌 Struktur Kolom (sheet utama)
Baris data dimulai pada **row 2**.

| Kolom | Isi |
|-------|-----|
| A | Kode (format: `(Negara)(Produk)(Segmen)`, contoh `MXAMRG`) |
| B | Negara (hasil VLOOKUP) |
| C | Produk (dipilih dari tabel ganjil/genap) |
| D | Segmen |
| E | Tanggal |
| F | Jumlah Barang |
| G | Harga Barang |
| H | Penjualan = Jumlah × Harga × (1 - Diskon dari NIM) |

---

## 1) Input NIM

Isi **NIM** pada area identitas (misalnya sel di sebelah label "NIM").

> Contoh: `104042699992` → digit terakhir **2** → genap → gunakan tabel harga **Mahasiswa NIM Genap**

Kemudian buat **Named Range** untuk sel ini:

Formulas → Define Name → Name: NIM_SISWA


---

## 2) Buat Named Range (penting sebelum rumus)

| Nama | Range yang dipilih di Sheet Lookup |
|------|-----------------------------------|
| `NegaraTbl` | blok tabel: Kode Negara – Negara (2 kolom) |
| `ProdukGanjilTbl` | blok tabel produk **Mahasiswa NIM Ganjil** (Kode Produk – Produk – Harga) |
| `ProdukGenapTbl` | blok tabel produk **Mahasiswa NIM Genap** (Kode Produk – Produk – Harga) |
| `SegmenTbl` | tabel kode segmen (2 kolom) |

> Pastikan blok **tidak ikut header kosong** dan **kolom pertama adalah Kode Produk**.

---

## 3) Rumus VLOOKUP (tempel di baris 2, lalu tarik ke bawah)

### **B2 — Negara**
```excel
=VLOOKUP(LEFT($A2;2);NegaraTbl;2;FALSE)

## C2 — Produk (cek ganjil/genap dari digit terakhir NIM)

=IF(NIM_SISWA="";""; 
   IF(ISEVEN(--RIGHT(NIM_SISWA;1));
      VLOOKUP(UPPER(TRIM(MID($A2;3;3)));ProdukGenapTbl;2;FALSE);
      VLOOKUP(UPPER(TRIM(MID($A2;3;3)));ProdukGanjilTbl;2;FALSE)
))

## D2 — Segmen
=VLOOKUP(RIGHT($A2;1);SegmenTbl;2;FALSE)

## G2 — Harga Barang
=IF(NIM_SISWA="";""; 
   IF(ISEVEN(--RIGHT(NIM_SISWA;1));
      VLOOKUP(UPPER(TRIM(MID($A2;3;3)));ProdukGenapTbl;3;FALSE);
      VLOOKUP(UPPER(TRIM(MID($A2;3;3)));ProdukGanjilTbl;3;FALSE)
))

4) Hitung Penjualan (H2)
Buat dulu DiskonNIM (sel bantu)

=VALUE(RIGHT(NIM_SISWA;1))/100

Define Name → DiskonNIM
Rumus Penjualan
=F2 * G2 * (1 - DiskonNIM)

5) H82–H85 — Total / Rata-Rata / Terbesar / Terkecil

(Sesuaikan range jika jumlah baris lebih banyak)

H82 → =SUM(H2:H81)
H83 → =AVERAGE(H2:H81)
H84 → =MAX(H2:H81)
H85 → =MIN(H2:H81)

6) Freeze header + Sort berdasarkan tanggal

1. View → Freeze Panes → Freeze Top Row

2. Klik kolom tanggal (E)

3. Data → Sort → Sort Oldest to Newest

4. Saat muncul pilihan:
- ✅ pilih Expand the selection
- ❌ jangan pilih "Continue with current selection"

Jika tanggal tidak berubah urutan, ubah ke format tanggal:
Data → Text to Columns → Finish (tanpa ubahan apa pun)

7) Import file Employee Data
Data → From Text/CSV → pilih Employee Data.txt → Load

8) Filter #1 → sheet baru: “R&D Low Satisfaction”
Di sheet Employee Data:
| Kolom                   | Filter                   |
| ----------------------- | ------------------------ |
| Department              | `Research & Development` |
| EnvironmentSatisfaction | `1`                      |
| JobSatisfaction         | `1`                      |

- Blok hasil + header → Copy
- Buat sheet baru → rename R&D Low Satisfaction
- Paste

9) Filter #2 → sheet baru: “Onsite Campus Fair Rep.”
Clear filter dulu:
Data → Clear
| Kolom      | Filter                              |
| ---------- | ----------------------------------- |
| Department | `Sales`                             |
| JobRole    | `Sales Representative`              |
| Age        | `< 30` (Number Filters → Less Than) |

Copy hasil → sheet baru → rename:
Onsite Campus Fair Rep
