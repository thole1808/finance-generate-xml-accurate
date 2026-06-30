# Generate Excel to XML Accurate

Aplikasi desktop untuk Windows yang membaca file Excel dan menghasilkan file XML.

Pilihan `Jenis XML` yang tersedia:

- `SALES_INVOICE`
- `SALES_RECEIPT`
- `OTHER_PAYMENT`

Jika `Jenis XML` dipilih `SALES_INVOICE`, aplikasi menghasilkan format Accurate dengan struktur `NMEXML > TRANSACTIONS > SALESINVOICE > ITEMLINE`.

Jika `Jenis XML` dipilih `SALES_RECEIPT`, aplikasi menghasilkan format Accurate dengan struktur `NMEXML > TRANSACTIONS > CUSTOMERRECEIPT > InvoiceLine`.

Jika `Jenis XML` dipilih `OTHER_PAYMENT`, aplikasi menghasilkan format Accurate dengan struktur `NMEXML > TRANSACTIONS > OTHERPAYMENT > ACCOUNTLINE`.

Aplikasi mendukung file Excel `.xls`, `.xlsx`, `.xlsm`, `.xltx`, dan `.xltm`.

Secara default hasil XML disimpan ke folder `Downloads`. Gunakan tombol `Download XML / Save As` untuk memilih lokasi lain, dan tombol `Buka Folder` untuk membuka lokasi hasil export.

## Cara Jalan di Windows

1. Install Python 3.11 atau 3.12.
2. Double-click:

```text
run_windows.bat
```

Atau jalankan manual dari Command Prompt:

```bash
pip install -r requirements.txt
python app.py
```

## Build Jadi EXE

Double-click:

```text
build_exe_windows.bat
```

Atau jalankan manual:

```bash
pyinstaller --noconsole --onefile --name "Generate Accurate XML" app.py
```

Hasil aplikasi ada di:

```text
dist/Generate Accurate XML.exe
```

File `.exe` ini yang dibagikan ke user Windows. User tidak perlu menerima file `.bat`, source code, atau menjalankan script.

Catatan: build `.exe` sebaiknya dilakukan di Windows.

PyInstaller tidak bisa membuat `.exe` Windows langsung dari macOS secara normal. Jika sedang memakai Mac, gunakan salah satu opsi berikut untuk build Windows:

- Windows VM di Mac, lalu jalankan `build_exe_windows.bat`.
- PC/laptop Windows, lalu jalankan `build_exe_windows.bat`.
- GitHub Actions runner Windows untuk build otomatis.

## Build Windows EXE dari macOS via GitHub Actions

Push project ke GitHub, lalu buka:

```text
Actions > Build Windows EXE > Run workflow
```

Setelah workflow selesai, download artifact:

```text
Generate-Accurate-XML-Windows
```

Di dalam artifact ada file:

```text
Generate Accurate XML.exe
```

## Build Jadi Installer macOS PKG

Jalankan di Mac:

```bash
chmod +x build_macos_pkg.sh
./build_macos_pkg.sh
```

Hasil installer ada di:

```text
release-macos/Generate Accurate XML.pkg
```

File `.pkg` ini yang dibagikan ke user macOS. User tidak perlu menerima script atau source code.

Catatan: build `.pkg` dilakukan di macOS.

Kalau hanya butuh file `.app`, jalankan:

```bash
chmod +x build_macos_app.sh
./build_macos_app.sh
```

Hasilnya:

```text
release-macos/Generate Accurate XML.app
```

## Cara Jalan di macOS / Linux

Jalankan:

```bash
pip install -r requirements.txt
python app.py
```

## Format Excel

Baris pertama dianggap sebagai header kolom. Baris berikutnya dianggap sebagai data.

Contoh:

```text
Kode Customer | Nama Customer | Alamat
C001          | PT Contoh     | Jakarta
```

akan menjadi XML dengan tag sesuai nama header.

## Langkah Berikutnya

Kirim contoh Excel dan contoh XML yang diharapkan dari Accurate supaya mapping XML bisa dibuat persis.
