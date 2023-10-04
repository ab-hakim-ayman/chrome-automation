import openpyxl
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import datetime
import time
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# current_day_name = datetime.datetime.now().strftime('%A')
current_day_name = 'Friday'

excel_file = "Excel.xlsx"
workbook = openpyxl.load_workbook(excel_file)

worksheet = workbook[current_day_name]

df = pd.DataFrame(worksheet.values)
print(df)

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

driver.get('https://www.google.com')
driver.maximize_window()
search_box = driver.find_element(By.NAME, 'q')

for row in range(len(df.values)):  # This creates a range from 2 to 12
    keyword = df.at[row, 2]
    search_box.clear()
    if keyword is not None:
        search_box.send_keys(keyword)
        time.sleep(2)
        suggestions = driver.find_elements(By.TAG_NAME, 'li')
        suggestions = suggestions[0:len(suggestions)-2]
        shortest_suggestion = min(suggestions, key=lambda x: len(x.text)).text
        longest_suggestion = max(suggestions, key=lambda x: len(x.text)).text

        print('Shortest Suggestion:', shortest_suggestion)
        print('Longest Suggestion:', longest_suggestion)

        worksheet.cell(row=row+1, column=4, value=longest_suggestion)
        worksheet.cell(row=row+1, column=5, value=shortest_suggestion)

workbook.save(excel_file)
driver.quit()


