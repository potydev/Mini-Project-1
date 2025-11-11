# Project Mini - Analisis Kelulusan dan Predikat Wisuda Mahasiswa

## 📖 Deskripsi Project

Program ini digunakan untuk menganalisis data calon wisudawan dari file Excel. Sistem melakukan perhitungan jumlah wisudawan per program studi, klasifikasi grade nilai akademik, penentuan predikat kelulusan, dan visualisasi hasil analisis menggunakan grafik.

## 🎯 Tujuan Project

- Menghitung jumlah wisudawan per program studi
- Klasifikasi grade nilai akademik berdasarkan IPK
- Penentuan predikat kelulusan (cumlaude, sangat memuaskan, memuaskan, dll.)
- Visualisasi hasil analisis menggunakan grafik batang dan pie chart
- Ekspor hasil analisis ke file Excel baru

## 📁 Struktur File

```
mini project 1/
├── Project Mini.py          # File utama program Python
├── Dataa Wisudawannn.xlsx   # File data input (raw data)
├── rekap_wisuda_final.xlsx  # File output hasil analisis
└── README.md                # Dokumentasi project ini
```

## 🛠️ Alur Proses Pembuatan Project

### Phase 1: Persiapan dan Planning (Jobsheet Analysis)
1. **Membaca requirements dari jobsheet** (`mantap.md`)
2. **Memahami struktur data** yang dibutuhkan:
   - NIM, Nama Mahasiswa, Program Studi
   - IPK, Lama Studi (Semester), Tahun Wisuda

### Phase 2: Development - Version 1 (Basic Structure)
1. **Import library yang dibutuhkan**:
   ```python
   import pandas as pd      # Untuk processing data Excel
   import matplotlib.pyplot as plt  # Untuk visualisasi
   ```

2. **Load data dari Excel**:
   ```python
   data = pd.read_excel('Dataa Wisudawannn.xlsx')
   ```

3. **Implementasi logika kalkulasi**:
   - Grade berdasarkan IPK (A, B+, B, C, D)
   - Predikat berdasarkan IPK dan lama studi

### Phase 3: Refactoring - Version 2 (Modular Structure)
1. **Restrukturisasi kode menjadi fungsi-fungsi modular**:
   - `load_and_validate_data()` - Load dan validasi data
   - `calculate_grade()` - Hitung grade IPK
   - `calculate_predikat()` - Hitung predikat kelulusan
   - `process_data()` - Proses data utama
   - `save_results()` - Simpan hasil ke Excel

2. **Menambahkan error handling dan validasi**:
   - Validasi kolom yang diperlukan
   - Validasi range IPK (0-4)
   - Deteksi data duplikat
   - Error handling untuk file tidak ditemukan

### Phase 4: Enhancement - Version 3 (Visualization & UI)
1. **Membuat visualisasi yang informatif**:
   - **Grafik Batang**: Distribusi wisudawan per prodi dengan breakdown cumlaude/non-cumlaude
   - **Grafik Pie Chart**: Distribusi predikat kelulusan
   - **Grafik IPK Comparison**: Rata-rata IPK per program studi

2. **Menambahkan fitur statistik lengkap**:
   - Total wisudawan per prodi
   - Persentase cumlaude per prodi
   - Statistik IPK (tertinggi, terendah, rata-rata)
   - Statistik lama studi

3. **Improvement user interface**:
   - Progress indicators saat processing
   - Statistik ringkasan di console
   - Judul grafik yang jelas ("GRAFIK BATANG", "GRAFIK PIE CHART")

### Phase 5: Final Optimization
1. **Simplifikasi tampilan grafik**:
   - Menghapus total dan persentase di atas batang
   - Menampilkan hanya nilai numerik di dalam segmen batang
   - Menyesuaikan judul grafik sesuai permintaan user

2. **Final testing dan validation**:
   - Testing dengan data real (313 wisudawan)
   - Validasi output Excel sesuai format
   - Verifikasi semua visualisasi berjalan dengan baik

## 📊 Kriteria Kalkulasi

### Grade IPK
| Rentang IPK | Grade |
|-------------|-------|
| 3.75 - 4.00 | A |
| 3.50 - 3.74 | B+ |
| 3.00 - 3.49 | B |
| 2.50 - 2.99 | C |
| < 2.50      | D |

### Predikat Kelulusan
| Kriteria | Predikat |
|----------|----------|
| IPK ≥ 3.75 dan Lama Studi ≤ 8 semester | Cumlaude (Dengan Pujian) |
| IPK ≥ 3.50 dan Lama Studi ≤ 9 semester | Sangat Memuaskan |
| IPK ≥ 3.00 | Memuaskan |
| IPK < 3.00 | Cukup Memuaskan |

## 🚀 Cara Menjalankan Program

1. **Pastikan library terinstall**:
   ```bash
   pip install pandas matplotlib openpyxl
   ```

2. **Jalankan program**:
   ```bash
   python "Project Mini.py"
   ```

3. **Output yang dihasilkan**:
   - Console: Statistik lengkap analisis
   - File Excel: `rekap_wisuda_final.xlsx`
   - Grafik: 3 visualisasi otomatis terbuka

## 📈 Hasil Analisis (Sample Output)

### Statistik Ringkasan:
- **Total Wisudawan**: 313 orang
- **Cumlaude**: 40 orang (12.8%)
- **Program Studi Terbanyak**: D4 TRET (34 orang)
- **Program Studi Cumlaude Terbanyak**: D3 TL (31.8%)

### Visualisasi:
1. **GRAFIK BATANG**: Menunjukkan distribusi wisudawan per prodi
2. **GRAFIK PIE CHART**: Persentase distribusi predikat kelulusan
3. **Grafik IPK**: Perbandingan rata-rata IPK antar prodi

## 🛡️ Error Handling

Program dilengkapi dengan:
- ✅ Validasi keberadaan file input
- ✅ Validasi struktur kolom data
- ✅ Deteksi IPK di luar range valid
- ✅ Deteksi data NIM duplikat
- ✅ Graceful error messages

## 📝 Lesson Learned

1. **Modular structure** membuat kode lebih maintainable
2. **Error handling** penting untuk production-ready code
3. **User feedback** improves usability
4. **Simplicity** is key for effective visualization
5. **Iterative development** leads to better results

## 🎉 Status Project

**Status**: ✅ COMPLETED
**Last Updated**: 2025-01-11
**Version**: 3.0 (Final Optimized)

Program siap digunakan untuk presentasi dan analisis data wisuda!