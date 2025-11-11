import pandas as pd
import matplotlib.pyplot as plt
import os
import sys

# Konfigurasi
CONFIG = {
    'input_file': 'Dataa Wisudawannn.xlsx',
    'output_file': 'rekap_wisuda_final.xlsx',
    'show_plots': True
}

def load_and_validate_data(file_path):
    """
    Memuat data dari file Excel dan melakukan validasi dasar
    """
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File {file_path} tidak ditemukan!")

        data = pd.read_excel(file_path)

        # Validasi kolom yang diperlukan
        required_columns = ['NIM', 'Nama Mahasiswa', 'Program Studi', 'IPK', 'Lama Studi (Semester)', 'Tahun Wisuda']
        missing_columns = [col for col in required_columns if col not in data.columns]
        if missing_columns:
            raise ValueError(f"Kolom yang diperlukan tidak ada: {missing_columns}")

        # Validasi data kosong
        if data.empty:
            raise ValueError("File Excel kosong!")

        # Validasi range IPK
        invalid_ipk = data[(data['IPK'] < 0) | (data['IPK'] > 4)]
        if not invalid_ipk.empty:
            print(f"[WARNING] Terdapat {len(invalid_ipk)} data dengan IPK tidak valid (diluar range 0-4)")

        # Cek duplikasi NIM
        duplicates = data[data.duplicated('NIM', keep=False)]
        if not duplicates.empty:
            print(f"[WARNING] Terdapat {len(duplicates)} data NIM duplikat")

        print(f"[OK] Data berhasil dimuat: {len(data)} baris")
        return data

    except FileNotFoundError as e:
        print(f"[ERROR] {e}")
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] Saat memuat data: {e}")
        sys.exit(1)

def calculate_grade(ipk):
    """
    Menghitung grade berdasarkan IPK
    """
    if ipk >= 3.75:
        return 'A'
    elif ipk >= 3.50:
        return 'B+'
    elif ipk >= 3.00:
        return 'B'
    elif ipk >= 2.50:
        return 'C'
    else:
        return 'D'

def calculate_predikat(row):
    """
    Menghitung predikat kelulusan berdasarkan IPK dan lama studi
    """
    ipk = row['IPK']
    lama_studi = row['Lama Studi (Semester)']

    if ipk >= 3.75 and lama_studi <= 8:
        return 'Cumlaude (Dengan Pujian)'
    elif ipk >= 3.50 and lama_studi <= 9:
        return 'Sangat Memuaskan'
    elif ipk >= 3.00:
        return 'Memuaskan'
    else:
        return 'Cukup Memuaskan'

def process_data(data):
    """
    Memproses data: menambahkan kolom grade dan predikat
    """
    print("[PROCESSING] Memproses data...")

    # Tambahkan kolom Grade
    data['Grade'] = data['IPK'].apply(calculate_grade)

    # Tambahkan kolom Predikat
    data['Predikat'] = data.apply(calculate_predikat, axis=1)

    print("[OK] Data berhasil diproses")
    return data

def save_results(data, output_file):
    """
    Menyimpan hasil ke file Excel baru
    """
    try:
        data.to_excel(output_file, index=False)
        print(f"[OK] File '{output_file}' berhasil dibuat!")
        return True
    except Exception as e:
        print(f"[ERROR] Menyimpan file: {e}")
        return False

def create_bar_chart(data):
    """
    Membuat grafik batang
    """
    print("[CHART] Membuat grafik batang yang informatif...")

    # Menghitung jumlah wisudawan per prodi
    jumlah_per_prodi = data['Program Studi'].value_counts()

    # Menghitung jumlah cumlaude per prodi
    cumlaude_per_prodi = data[data['Predikat'] == 'Cumlaude (Dengan Pujian)']['Program Studi'].value_counts()

    # Menyusun data untuk grafik
    prodi_list = jumlah_per_prodi.index.tolist()
    total_list = jumlah_per_prodi.values.tolist()
    cumlaude_list = [cumlaude_per_prodi.get(prodi, 0) for prodi in prodi_list]
    non_cumlaude_list = [total - cumlaude for total, cumlaude in zip(total_list, cumlaude_list)]

    # Membuat figure dengan ukuran lebih besar
    plt.figure(figsize=(14, 8))

    # Membuat grafik batang bertumpuk dengan warna yang lebih kontras
    bars_non_cumlaude = plt.bar(prodi_list, non_cumlaude_list, color='lightsteelblue',
                               label='Non-Cumlaude', alpha=0.8, edgecolor='navy', linewidth=0.5)
    bars_cumlaude = plt.bar(prodi_list, cumlaude_list, bottom=non_cumlaude_list,
                           color='gold', label='Cumlaude', alpha=0.9, edgecolor='orange', linewidth=0.5)

    # Menambahkan judul dan label yang lebih deskriptif
    plt.title('GRAFIK BATANG',
             fontsize=16, fontweight='bold', pad=20)
    plt.xlabel('Program Studi', fontsize=13, fontweight='bold')
    plt.ylabel('Jumlah Mahasiswa', fontsize=13, fontweight='bold')

    # Memutar label x dan mengatur font
    plt.xticks(rotation=45, ha='right', fontsize=11)
    plt.yticks(fontsize=11)

    # Menambahkan grid untuk kemudahan membaca
    plt.grid(axis='y', alpha=0.3, linestyle='--')

    # Menambahkan nilai numerik di setiap bagian batang
    for i, (non_cum, cumlaude) in enumerate(zip(non_cumlaude_list, cumlaude_list)):
        # Nilai untuk bagian non-cumlaude
        if non_cum > 0:
            plt.text(i, non_cum/2, f'{non_cum}', ha='center', va='center',
                    fontsize=10, fontweight='bold', color='darkblue')

        # Nilai untuk bagian cumlaude
        if cumlaude > 0:
            plt.text(i, non_cum + cumlaude/2, f'{cumlaude}', ha='center', va='center',
                    fontsize=10, fontweight='bold', color='darkred')

    # Menambahkan informasi statistik di bagian bawah
    total_cumlaude = sum(cumlaude_list)
    total_students = sum(total_list)
    overall_cumlaude_pct = (total_cumlaude/total_students)*100 if total_students > 0 else 0

    stats_text = (f'Total Wisudawan: {total_students} orang | '
                 f'Cumlaude: {total_cumlaude} orang ({overall_cumlaude_pct:.1f}%) | '
                 f'Jumlah Program Studi: {len(prodi_list)}')
    plt.figtext(0.5, 0.02, stats_text, ha="center", fontsize=11,
                bbox=dict(boxstyle="round,pad=0.5", facecolor="lightgray", alpha=0.8))

    # Memperbaiki legenda dengan informasi tambahan
    legend_text = [f'Non-Cumlaude ({sum(non_cumlaude_list)} orang)',
                  f'Cumlaude ({total_cumlaude} orang)']
    plt.legend(legend_text, loc='upper right', fontsize=11,
              bbox_to_anchor=(1.0, 1.0), framealpha=0.9)

    # Mengatur batas y-axis agar lebih proporsional
    plt.ylim(0, max(total_list) + 2)

    # Mengatur layout
    plt.tight_layout()

    # Menampilkan grafik
    plt.show()

def create_pie_chart(data):
    """
    Membuat grafik lingkaran distribusi predikat kelulusan
    """
    print("[CHART] Membuat grafik lingkaran...")

    # Menghitung distribusi predikat
    distribusi_predikat = data['Predikat'].value_counts()

    # Membuat figure
    plt.figure(figsize=(8, 8))

    # Membuat label dengan nama predikat dan jumlahnya
    labels_with_counts = [f'{label}\n({count} orang)' for label, count in zip(distribusi_predikat.index, distribusi_predikat.values)]

    # Color scheme yang menarik
    colors = ['#FF9999', '#66B2FF', '#99FF99', '#FFCC99', '#FFD700']

    # Membuat pie chart
    plt.pie(distribusi_predikat.values,
            labels=labels_with_counts,
            autopct='%1.1f%%',
            startangle=90,
            colors=colors[:len(distribusi_predikat)])

    plt.title('GRAFIK PIE CHART', fontsize=14, fontweight='bold')
    plt.tight_layout()
    plt.show()

def create_ipk_comparison_chart(data):
    """
    Membuat grafik perbandingan rata-rata IPK per program studi (tambahan baru)
    """
    print("[CHART] Membuat grafik rata-rata IPK per prodi...")

    # Menghitung rata-rata IPK per prodi
    avg_ipk_per_prodi = data.groupby('Program Studi')['IPK'].mean().sort_values(ascending=False)

    # Membuat figure
    plt.figure(figsize=(12, 6))

    # Membuat grafik batang horizontal
    bars = plt.barh(range(len(avg_ipk_per_prodi)), avg_ipk_per_prodi.values, color='lightgreen')

    # Menambahkan nilai IPK di setiap batang
    for i, (bar, value) in enumerate(zip(bars, avg_ipk_per_prodi.values)):
        plt.text(value + 0.02, bar.get_y() + bar.get_height()/2,
                f'{value:.2f}', ha='left', va='center')

    # Mengatur label
    plt.yticks(range(len(avg_ipk_per_prodi)), avg_ipk_per_prodi.index)
    plt.xlabel('Rata-rata IPK', fontsize=12)
    plt.ylabel('Program Studi', fontsize=12)
    plt.title('Rata-rata IPK per Program Studi', fontsize=14, fontweight='bold')

    # Mengatur range x-axis
    plt.xlim(0, max(avg_ipk_per_prodi.values) + 0.3)

    plt.tight_layout()
    plt.show()

def create_visualizations(data):
    """
    Fungsi utama untuk membuat semua visualisasi
    """
    if not CONFIG['show_plots']:
        return

    create_bar_chart(data)
    create_pie_chart(data)
    create_ipk_comparison_chart(data)

def print_summary_statistics(data):
    """
    Mencetak statistik ringkasan data
    """
    print("\n" + "="*60)
    print("STATISTIK RINGKASAN ANALISIS WISUDA")
    print("="*60)

    print(f"\nTotal Wisudawan: {len(data)} orang")

    # Statistik per program studi
    print(f"\nJumlah Wisudawan per Program Studi:")
    for prodi, count in data['Program Studi'].value_counts().items():
        cumlaude_count = data[(data['Program Studi'] == prodi) &
                             (data['Predikat'] == 'Cumlaude (Dengan Pujian)')].shape[0]
        percentage = (cumlaude_count / count) * 100
        print(f"   • {prodi}: {count} orang (Cumlaude: {cumlaude_count} orang, {percentage:.1f}%)")

    # Statistik predikat
    print(f"\nDistribusi Predikat Kelulusan:")
    for predikat, count in data['Predikat'].value_counts().items():
        percentage = (count / len(data)) * 100
        print(f"   • {predikat}: {count} orang ({percentage:.1f}%)")

    # Statistik IPK
    print(f"\nStatistik IPK:")
    print(f"   • IPK Tertinggi: {data['IPK'].max():.2f}")
    print(f"   • IPK Terendah: {data['IPK'].min():.2f}")
    print(f"   • Rata-rata IPK: {data['IPK'].mean():.2f}")

    # Statistik lama studi
    print(f"\nStatistik Lama Studi:")
    print(f"   • Tercepat: {data['Lama Studi (Semester)'].min()} semester")
    print(f"   • Terlama: {data['Lama Studi (Semester)'].max()} semester")
    print(f"   • Rata-rata: {data['Lama Studi (Semester)'].mean():.1f} semester")

    print("="*60)

def main():
    """
    Fungsi utama program
    """
    print("PROGRAM ANALISIS KELULUSAN WISUDAWAN")
    print("="*50)

    # 1. Load dan validasi data
    data = load_and_validate_data(CONFIG['input_file'])

    # 2. Proses data (tambah grade dan predikat)
    processed_data = process_data(data)

    # 3. Tampilkan hasil awal
    print("\nSample Data Hasil Processing:")
    print(processed_data[['NIM', 'Nama Mahasiswa', 'Program Studi', 'IPK', 'Grade', 'Predikat']].head(10))

    # 4. Simpan hasil
    save_results(processed_data, CONFIG['output_file'])

    # 5. Tampilkan statistik ringkasan
    print_summary_statistics(processed_data)

    # 6. Buat visualisasi
    create_visualizations(processed_data)

    print("\n[OK] Program selesai dijalankan!")

if __name__ == "__main__":
    main()


