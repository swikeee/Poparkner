# Poparkner

**Agent Co Dashboard · Podomoro Park Bandung**: satu tempat untuk koordinasi agent properti, dari MOU kerjasama sampai komisi cair.

## Segmen
| # | Segmen | Isi |
|---|--------|-----|
| 01 | MOU Kerjasama | Upload & review dokumen agency (KTP Direksi, NPWP, NIB, Rekening, Akta, AREBI) |
| 02 | Data Agent | Profil agent + konsumen yang dibawa (minat cluster/tipe, budget, follow-up) |
| 03 | Aktivitas | Heatmap visit/closing, kalender event & dampaknya, program promo, tracker visit |
| 04 | Bank Rekanan | Pengajuan BI checking (SLIK) & simulasi KPR / cash bertahap / cash keras |
| 05 | Pencairan Komisi | Pengajuan nominal → balasan berkas → verifikasi SAD → cair |

## Role
`Agent Co` · `Principal` · `Sales Agent` · `Pihak Bank`. Setiap role hanya melihat data miliknya; di mode Google Sheet, filter akses dijalankan di server.

## Mode
- **Lokal**: buka `index.html`, data disimpan di browser (data dummy).
- **Google Sheet**: backend Apps Script di `backend/Code.gs`, data di Google Sheet, file di Google Drive.

## Setup backend
1. Buat Google Sheet baru → **Extensions → Apps Script** → tempel `backend/Code.gs` → Save.
2. Jalankan fungsi `setup` sekali (izinkan akses). Tab data, tab `Users`, dan folder Drive *Poparkner Files* dibuat otomatis.
3. **Deploy → New deployment → Web app**: Execute as *Me*, Who has access *Anyone*. Salin URL `/exec`.
4. Buka dashboard → **Hubungkan Google Sheet** di layar login → tempel URL → Simpan.
   (Atau isi `DEFAULT_API_URL` di `index.html` agar semua pengguna otomatis terhubung.)

### Kelola user
Tambah baris di tab **Users** (`username`, `password_baru`, `role`, dan `agency` / `agent` / `bank` sesuai role), lalu menu **Poparkner → Proses password user baru**. Password di-hash dan kolom `password_baru` dikosongkan otomatis.

Akun demo setelah `setup`: `agentco` / `agentco123`, `principal.raywhite` / `principal123`, `sales.rina` / `sales123`, `bank.bca` / `bank123`. **Ganti password ini sebelum dipakai sungguhan.**

> Suku bunga bank di simulasi KPR masih data dummy dan belum terhubung ke website bank.
