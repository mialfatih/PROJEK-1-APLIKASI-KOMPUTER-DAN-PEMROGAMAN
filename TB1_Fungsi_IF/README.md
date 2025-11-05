# TB1 — Logika Excel (IF, Nested IF, AND, OR)

## 🎯 Tujuan
Mengolah data penjualan dan mengkategorikan hasil dengan fungsi:
- `IF`
- `AND`, `OR`
- Nested `IF`

## 📌 Langkah

### 1. Load CSV

Data → From Text/CSV → Sales_Data.csv → Load


### 2. Isi kolom N–W
| Kolom | Rumus |
|-------|--------|
| N (NIM) | isi manual di N2 lalu tarik ke bawah |
| O (Faktor_NIM) | `=VALUE(RIGHT($N2;3))/100` |
| P (Performance_Score) | `=E2*$O2` |
| Q (Kategori_Performance) | `=IF(P2>=5000;"Excellent";IF(P2>=3000;"Good";IF(P2>=1500;"Average";"Below Average")))` |
| R (Profit_Per_Unit) | `=I2-H2` |
| S (Total_Profit) | `=R2*F2` |
| T (Profit_Category) | `=IF(S2>=1000;"High Profit";IF(S2>=500;"Medium Profit";"Low Profit"))` |
| U (Discount_Category) | (versi aman format angka) `=IF(IF(K2>=1;K2/100;K2)>=0,15;"High Discount";IF(IF(K2>=1;K2/100;K2)>=0,05;"Medium Discount";IF(IF(K2>=1;K2/100;K2)>0;"Low Discount";"No Discount")))` |
| V (Sales_Priority) | logika AND/OR sesuai template |
| W (Bonus_Amount) | nested IF sesuai kategori |

### 3. Simpan file
