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

- Gunakan **Pilih Folder** untuk memilih satu folder foto jika browser/sistem operasi mendukung pemilihan folder.
- Jika pemilih folder tidak tersedia di Linux atau browser tertentu, gunakan **Pilih Banyak Foto**. Pada dialog file, buka folder foto lalu pilih beberapa foto sekaligus (misalnya `Ctrl+A` untuk memilih semua file yang terlihat).
- Foto HEIC/HEIF dikonversi menjadi PNG untuk thumbnail dan diproses lagi saat dibuat ke PDF. Decoder HEIC dimuat saat dibutuhkan dari CDN jsDelivr, sehingga koneksi internet diperlukan untuk konversi HEIC/HEIF.
- Pencocokan otomatis tetap menggunakan nama file foto dan nama siswa. Pastikan nama file foto sesuai dengan nama siswa.

