# Rapor Generator — Default Settings Update

Updated default settings for biodata and signature pages based on the requested template-specific values.

## Biodata defaults
- Glow Lower: X 0, Y 0, Font 12
- Glow Middle: X 0, Y 5, Font 12
- Glow Upper: X 0, Y 5, Font 12
- Sunny: X 0, Y 5, Font 12
- Infant: X 0, Y 5, Font 12

## Signature defaults
Each template/page now has its own default date, teacher, principal, and digital signature box coordinates.

The localStorage keys were versioned so the new defaults are applied instead of older saved defaults from previous versions. After the user changes a setting, it remains saved normally.

## Upload foto lintas sistem operasi

- Tombol **Pilih Folder** memakai pemilih folder modern jika tersedia (membutuhkan browser yang mendukung File System Access API dan konteks aman seperti HTTPS). Jika tidak tersedia, tombol beralih ke pemilih file sebagai fallback.
- **Pilih Banyak Foto** adalah opsi lintas browser/OS. Di dialog file, buka folder foto lalu pilih banyak foto sekaligus (misalnya `Ctrl+A` untuk memilih semua gambar yang terlihat).
- Foto JPG/PNG/WebP memakai URL file asli untuk thumbnail agar tidak bergantung pada dukungan konversi canvas yang berbeda antar-browser. HEIC/HEIF dikonversi menjadi PNG untuk thumbnail. Decoder HEIC dimuat saat dibutuhkan dari CDN jsDelivr, sehingga koneksi internet diperlukan untuk konversi HEIC/HEIF.
- Pencocokan otomatis tetap menggunakan nama file foto dan nama siswa. Pastikan nama file foto sesuai dengan nama siswa.



### Dukungan foto HEIC/HEIF

HEIC/HEIF dikonversi di browser memakai `heic-to` (libheif terbaru) yang dimuat dari jsDelivr ketika diperlukan. Koneksi internet diperlukan untuk memuat decoder. Hasil PNG yang sama digunakan untuk thumbnail dan PDF. Tidak semua HEIF yang memakai codec/varian tidak umum dijamin didukung oleh decoder.
