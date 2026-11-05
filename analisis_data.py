"""PRAKTIKUM MAHADATA — Analisis Antrean 3 Kantin Sekolah.
Sumber utama: database SQLite (database.db), bukan CSV.
Install: python -m pip install pandas matplotlib openpyxl
Jalankan: python analisis_data.py
"""
from pathlib import Path
import sqlite3
import pandas as pd
import matplotlib.pyplot as plt

BASE = Path(__file__).resolve().parent
DB = BASE / 'database.db'
GRAFIK = BASE / 'grafik'
GRAFIK.mkdir(exist_ok=True)

if not DB.is_file():
    raise FileNotFoundError(f'Database tidak ditemukan: {DB}. Simpan database.db bersama analisis_data.py')

# 1. KONEKSI SQLITE DAN QUERY SQL
with sqlite3.connect(DB) as conn:
    jumlah_data = pd.read_sql_query('SELECT COUNT(*) AS total FROM antrean', conn).iloc[0]['total']
    df = pd.read_sql_query('SELECT * FROM antrean ORDER BY id', conn)
    sql_jam = pd.read_sql_query('''
        SELECT jam, COUNT(*) AS jumlah_pesanan,
               ROUND(AVG(panjang_antrean), 2) AS rata_antrean,
               ROUND(AVG(waktu_tunggu_menit), 2) AS rata_tunggu_menit
        FROM antrean GROUP BY jam ORDER BY jumlah_pesanan DESC, jam ASC LIMIT 10
    ''', conn)
    sql_kantin = pd.read_sql_query('''
        SELECT kantin, COUNT(*) AS jumlah_pesanan,
               ROUND(AVG(panjang_antrean), 2) AS rata_antrean,
               ROUND(AVG(waktu_tunggu_menit), 2) AS rata_tunggu_menit
        FROM antrean GROUP BY kantin ORDER BY rata_antrean DESC
    ''', conn)
    sql_menu = pd.read_sql_query('''
        SELECT menu, COUNT(*) AS jumlah_pesanan
        FROM antrean GROUP BY menu ORDER BY jumlah_pesanan DESC, menu ASC LIMIT 10
    ''', conn)
    sql_kombinasi = pd.read_sql_query('''
        SELECT kantin, jam, COUNT(*) AS jumlah_pesanan,
               ROUND(AVG(panjang_antrean), 2) AS rata_antrean
        FROM antrean GROUP BY kantin, jam
        ORDER BY rata_antrean DESC, jumlah_pesanan DESC LIMIT 10
    ''', conn)

# 2. PANDAS: CEK DATA, PEMBERSIHAN, DAN PENGOLAHAN
print('\n=== DATABASE & PANDAS ===')
print(f'Jumlah baris database: {jumlah_data}')
print('Ukuran data:', df.shape)
print('Kolom:', ', '.join(df.columns))
print('Data kosong per kolom:\n', df.isna().sum())
print('Jumlah ID duplikat:', df.duplicated(subset=['id']).sum())
df = df.drop_duplicates(subset=['id']).copy()
df['tanggal'] = pd.to_datetime(df['tanggal'], errors='coerce')
for kolom in ['jam', 'harga_rp', 'panjang_antrean', 'waktu_tunggu_menit']:
    df[kolom] = pd.to_numeric(df[kolom], errors='coerce')
df = df.dropna(subset=['tanggal', 'jam', 'kantin', 'menu', 'panjang_antrean'])
df = df[(df['panjang_antrean'] >= 0) & (df['waktu_tunggu_menit'] >= 0)].copy()
print('Baris valid setelah pembersihan:', len(df))

# Gunakan Pandas groupby untuk analisis dan grafik.
per_jam = df.groupby('jam', as_index=False).agg(
    jumlah_pesanan=('id', 'count'), rata_antrean=('panjang_antrean', 'mean'))
per_jam = per_jam.sort_values('jam')
per_kantin = df.groupby('kantin', as_index=False).agg(
    jumlah_pesanan=('id', 'count'), rata_antrean=('panjang_antrean', 'mean'),
    rata_tunggu_menit=('waktu_tunggu_menit', 'mean')).sort_values('rata_antrean', ascending=False)
per_menu = df.groupby('menu', as_index=False).agg(jumlah_pesanan=('id', 'count'))
per_menu = per_menu.sort_values(['jumlah_pesanan','menu'], ascending=[False,True]).head(10)

# 3. TAMPILKAN HASIL QUERY SQL TOP 10
for judul, tabel in [('TOP 10 JAM RAMAI (SQL)', sql_jam),
                     ('PERINGKAT 3 KANTIN (SQL)', sql_kantin),
                     ('TOP 10 MENU (SQL)', sql_menu),
                     ('TOP 10 KANTIN + JAM BERDASARKAN ANTREAN (SQL)', sql_kombinasi)]:
    print('\n===', judul, '===')
    print(tabel.to_string(index=False))

# 4. EKSPOR DATA KE CSV DAN EXCEL (bukan sumber utama)
df_export = df.copy()
df_export['tanggal'] = df_export['tanggal'].dt.strftime('%Y-%m-%d')
df_export.to_csv(BASE / 'hasil_export.csv', index=False, encoding='utf-8-sig')
with pd.ExcelWriter(BASE / 'hasil_export.xlsx', engine='openpyxl') as writer:
    df_export.to_excel(writer, sheet_name='Data 1000 Baris', index=False)
    sql_jam.to_excel(writer, sheet_name='Top 10 Jam', index=False)
    sql_kantin.to_excel(writer, sheet_name='Peringkat Kantin', index=False)
    sql_menu.to_excel(writer, sheet_name='Top 10 Menu', index=False)
    sql_kombinasi.to_excel(writer, sheet_name='Top 10 Kantin Jam', index=False)

# 5. GRAFIK DENGAN MATPLOTLIB
plt.rcParams.update({'font.size': 10})
fig, ax = plt.subplots(figsize=(9, 5))
bars = ax.bar([f'{int(j):02d}.00' for j in per_jam['jam']], per_jam['jumlah_pesanan'])
ax.bar_label(bars, padding=3)
ax.set(title='Jumlah Pesanan Menurut Jam', xlabel='Jam', ylabel='Jumlah pesanan (10 hari)')
ax.grid(axis='y', alpha=.2)
fig.tight_layout()
fig.savefig(GRAFIK / 'grafik_1_jam_tersibuk.png', dpi=160)
plt.close(fig)

fig, ax = plt.subplots(figsize=(8, 5))
bars = ax.bar(per_kantin['kantin'], per_kantin['rata_antrean'])
ax.bar_label(bars, labels=[f'{v:.1f}' for v in per_kantin['rata_antrean']], padding=3)
ax.set(title='Rata-rata Panjang Antrean per Kantin', xlabel='Kantin', ylabel='Siswa dalam antrean')
ax.set_ylim(0, per_kantin['rata_antrean'].max() * 1.2)
ax.grid(axis='y', alpha=.2)
fig.tight_layout()
fig.savefig(GRAFIK / 'grafik_2_kantin_paling_antre.png', dpi=160)
plt.close(fig)

fig, ax = plt.subplots(figsize=(10, 6))
menu_plot = per_menu.iloc[::-1]
bars = ax.barh(menu_plot['menu'], menu_plot['jumlah_pesanan'])
ax.bar_label(bars, padding=3)
ax.set(title='Top 10 Menu Paling Banyak Dipesan', xlabel='Jumlah pesanan', ylabel='Menu')
ax.set_xlim(0, menu_plot['jumlah_pesanan'].max() * 1.16)
ax.grid(axis='x', alpha=.2)
fig.tight_layout()
fig.savefig(GRAFIK / 'grafik_3_menu_terlaris.png', dpi=160)
plt.close(fig)

print('\n=== KESIMPULAN ===')
print(f"Jam tersibuk: {int(sql_jam.iloc[0]['jam']):02d}.00 ({int(sql_jam.iloc[0]['jumlah_pesanan'])} pesanan)")
print(f"Kantin antrean terpanjang: {sql_kantin.iloc[0]['kantin']} (rata-rata {sql_kantin.iloc[0]['rata_antrean']:.1f} siswa)")
print(f"Menu terlaris: {sql_menu.iloc[0]['menu']} ({int(sql_menu.iloc[0]['jumlah_pesanan'])} pesanan)")
print('\nSELESAI: hasil_export.csv, hasil_export.xlsx, dan 3 PNG dalam folder grafik/')
