# Cek BPJS Pengemudi LPP

Situs sederhana untuk koordinator di lokasi klien: masukkan NIK pengemudi PT Layanan Pusaka Prima (LPP), lalu lihat status BPJS Ketenagakerjaan (No. KPJ) atau BPJS Kesehatan (No. kartu JKN, keluarga ditanggung). Setiap hasil menyebut periode data tagihan yang dipakai.

## Cara kerja
- `index.html`: halaman cek (statis, di-hosting GitHub Pages).
- `data/bpjs.json`: data tagihan bulan terbaru, **terenkripsi per pengemudi**. Kunci tiap data diturunkan dari NIK pengemudi itu sendiri (PBKDF2-SHA256, 300.000 iterasi, lalu AES-GCM), jadi isi hanya terbuka bila NIK 16 digit diisi dengan tepat. Nama ditampilkan tersamar.
- `tools/build_data.py`: mengolah dua file Excel unduhan resmi BPJS menjadi `data/bpjs.json`.

## Pembaruan bulanan
1. Simpan file Excel terbaru di Google Drive, folder **Cek BPJS Pengemudi LPP**:
   - `1. BPJS Ketenagakerjaan` → file *Data Periode Tagihan*
   - `2. BPJS Kesehatan` → file *Rincian Tagihan*
2. Tugas terjadwal Claude (Senin pagi) mengambil file terbaru, menjalankan `tools/build_data.py`, lalu meng-commit `data/bpjs.json` baru ke repo ini. GitHub Pages memperbaruinya otomatis.

Manual: `pip install openpyxl cryptography` lalu
`python3 tools/build_data.py TK.xlsx KS.xlsx data/bpjs.json`

## Catatan
- **Jangan pernah commit file Excel mentah** (sudah diblokir oleh `.gitignore`).
- Status "Aktif" berarti tercantum dalam data tagihan periode tersebut, bukan status langsung dari sistem BPJS.
- Peserta BPJS Kesehatan dihubungkan ke NIK lewat Nomor Pegawai (kode L-xxxx) di file BPJS TK, karena file Kesehatan tidak memuat NIK.
