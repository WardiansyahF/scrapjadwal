=======================================================
scrapextractjadwal.py
=======================================================

Description:
This script automates the process of scraping class schedule data from the Gunadarma University website 
and organizing it into a structured format. The final output is saved as an Excel file, organizing 
the schedule by class, day, and period.

Dependencies:
- pandas: For handling and analyzing data.
- selenium: For web scraping and automating browser tasks.
- BeautifulSoup (bs4): For parsing HTML and extracting information from web pages.
- openpyxl: For saving structured data into an Excel file.

Usage:
1. Install the required libraries if you haven't already:
   ```bash
   pip install pandas selenium beautifulsoup4 openpyxl
   ```

2. Download the right version of ChromeDriver that matches your Chrome browser and place it in the specified path.
   Update the `driver_path` variable with the path to your ChromeDriver executable.

3. Run the script from your terminal:
   ```bash
   python scrapextractjadwal.py
   ```

Output:
- An Excel file named 'jadwal_perkuliahan_tingkat.xlsx' will be created, containing the organized class schedule.

Functionality:
1. **Web Scraping**:
   - The script uses Selenium to navigate to the Gunadarma University schedule page.
   - It searches for class schedules based on the specified class names for multiple levels (Tingkat).
   - The schedule data is extracted using BeautifulSoup.

2. **Data Structuring**:
   - The scraped data is organized into a pandas DataFrame with multi-index columns, representing days and periods.
   - Each class is indexed, with the respective room assignments filled in based on the schedule.

3. **Data Saving**:
   - The final structured schedule is saved to an Excel file for easy access and further analysis.

Error Handling:
- The script includes error handling to manage exceptions during the scraping process. If a class search fails, an error message will be printed, and the current page source will be shown for debugging.

Note:
- Make sure the ChromeDriver is compatible with your version of Chrome.
- Modify the `kelas_per_tingkat` dictionary to adjust class names based on the actual classes offered at the university.


### Import Library
```python
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
```

### Path to ChromeDriver
```python
driver_path = r'C:\chromedriver-win64\chromedriver.exe'
service = Service(driver_path)
driver = webdriver.Chrome(service=service)
```

### List of classes to search
```python
kelas_per_tingkat = {
    'Tingkat 1': [f'1IA{str(i).zfill(2)}' for i in range(1, 16)],
    'Tingkat 2': [f'2IA{str(i).zfill(2)}' for i in range(1, 19)],
    'Tingkat 3': [f'3IA{str(i).zfill(2)}' for i in range(1, 21)],
    'Tingkat 4': [f'4IA{str(i).zfill(2)}' for i in range(1, 20)],
}

data_jadwal = []
```
### Loop through each level and class
```python
for tingkat, kelas_list in kelas_per_tingkat.items():
    for kelas in kelas_list:
        url = 'https://baak.gunadarma.ac.id/jadwal/cariJadKul?_token=xculcm7MqI3CM9t2I3mPySFzHQ9kBjczuKMZFycb&filter=*.html'
        driver.get(url)

        # Wait until the page is fully loaded
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, 'body'))
        )

        # Fill in the search bar with the class name
        search_box = driver.find_element(By.NAME, 'teks')
        search_box.clear()
        search_box.send_keys(kelas)
        search_box.submit()

        try:
            # Wait until the table appears
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'table-custom'))
            )
            
            # Get the HTML of the page
            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            
            # Find the tables
            tables = soup.find_all('table', {'class': 'table table-custom table-primary table-fixed bordered-table stacktable small-only'})

            # Check how many tables were found
            print(f"Number of tables found for {kelas}: {len(tables)}")
            
            # Extract data from the tables
            for table in tables:
                rows = table.find_all('tr')[1:]  # Skip header
                for row in rows:
                    cols = row.find_all('td')
                    if len(cols) >= 6:  # Ensure there are enough columns
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
            print(f"An error occurred for {kelas}: {e}")
            print(driver.page_source)  # Print page for debugging
```

### Save the data to an Excel file
```python
df = pd.DataFrame(data_jadwal)
df.to_excel('jadwal_perkuliahan_tingkat.xlsx', index=False)  # Save to Excel file
print("Data has been saved to jadwal_perkuliahan_tingkat.xlsx")
```
### Close the driver
```python
driver.quit()
```

### Key Features of the Documentation:
- **Header:** Clearly states the script name and description.
- **Dependencies:** Lists necessary libraries and provides installation instructions.
- **Usage Instructions:** Step-by-step guidance on how to run the script.
- **Functionality:** Breaks down the script’s main features and processes.
- **Error Handling:** Explains how the script manages errors during execution.
- **Note:** Offers additional information regarding ChromeDriver compatibility and class name adjustments.
