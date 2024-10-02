from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd

# Path ke ChromeDriver
driver_path = r'C:\chromedriver-win64\chromedriver.exe'
service = Service(driver_path)
driver = webdriver.Chrome(service=service)

# Daftar kelas yang akan dicari
kelas_per_tingkat = {
    'Tingkat 1': [f'1IA{str(i).zfill(2)}' for i in range(1, 16)],
    'Tingkat 2': [f'2IA{str(i).zfill(2)}' for i in range(1, 19)],
    'Tingkat 3': [f'3IA{str(i).zfill(2)}' for i in range(1, 21)],
    'Tingkat 4': [f'4IA{str(i).zfill(2)}' for i in range(1, 20)],
}

data_jadwal = []

# Loop untuk setiap tingkat dan kelas
for tingkat, kelas_list in kelas_per_tingkat.items():
    for kelas in kelas_list:
        url = 'https://baak.gunadarma.ac.id/jadwal/cariJadKul?_token=xculcm7MqI3CM9t2I3mPySFzHQ9kBjczuKMZFycb&filter=*.html'
        driver.get(url)

        # Tunggu sampai halaman sepenuhnya dimuat
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'body'))
        )

        # Mengisi search bar dengan nama kelas
        search_box = driver.find_element(By.NAME, 'teks')
        search_box.clear()
        search_box.send_keys(kelas)
        search_box.submit()

        try:
            # Tunggu hingga tabel muncul
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'table-custom'))
            )
            
            # Ambil halaman HTML
            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            
            # Temukan tabel
            tables = soup.find_all('table', {'class': 'table table-custom table-primary table-fixed bordered-table stacktable small-only'})

            # Cek jumlah tabel ditemukan
            print(f"Jumlah tabel ditemukan untuk {kelas}: {len(tables)}")
            
            # Ambil data dari tabel
            for table in tables:
                rows = table.find_all('tr')[1:]  # Lewati header
                for row in rows:
                    cols = row.find_all('td')
                    if len(cols) >= 6:  # Pastikan ada cukup kolom
                        data_row = {
                            'TINGKAT': tingkat,
                            'KELAS': cols[0].text.strip(),
                            'HARI': cols[1].text.strip(),
                            'MATA KULIAH': cols[2].text.strip(),
                            'WAKTU': cols[3].text.strip(),
                            'RUANG': cols[4].text.strip(),
                            'DOSEN': cols[5].text.strip(),
                        }
                        data_jadwal.append(data_row)
        
        except Exception as e:
            print(f"Terjadi kesalahan untuk {kelas}: {e}")
            print(driver.page_source)  # Cetak halaman untuk debugging

# Simpan data jadwal ke dalam DataFrame
df_jadwal = pd.DataFrame(data_jadwal)

# Proses data terstruktur
# Inisialisasi DataFrame untuk jadwal terstruktur
days = ['SENIN', 'SELASA', 'RABU', 'KAMIS', "JUM'AT", 'SABTU']
periods = [str(i) for i in range(1, 11)]  # 1 to 10 periods

# Ambil daftar kelas unik dari DataFrame
classes = df_jadwal['KELAS'].unique()  # Dapatkan semua kelas unik

# Inisialisasi jadwal terstruktur
schedule = pd.DataFrame(index=classes, columns=pd.MultiIndex.from_product([days, periods]))

# Isi jadwal terstruktur berdasarkan data yang sudah diambil
for index, row in df_jadwal.iterrows():
    kelas = row['KELAS']
    hari = row['HARI'].upper()
    
    # Check if 'WAKTU' is not NaN and convert to string
    if pd.notna(row['WAKTU']):
        waktu = str(row['WAKTU']).split('/')  # Split the periods (e.g., 1/2/3)
        ruang = row['RUANG'].strip()  # Hanya menyimpan kode ruang

        # Isi setiap slot waktu untuk mata kuliah dengan kode ruang
        for period in waktu:
            period = period.strip()  # Menghapus spasi di awal/akhir
            if period in periods and hari in days:  # Cek apakah periode dan hari valid
                schedule.loc[kelas, (hari, period)] = ruang  # Simpan hanya kode ruang

# Simpan jadwal terstruktur ke file Excel
schedule.to_excel('jadwal_perkuliahan_terstruktur_tingkat.xlsx')
print("Data berhasil disimpan ke jadwal_perkuliahan_terstruktur_tingkat.xlsx")

# Tutup driver
driver.quit()
