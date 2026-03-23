# Import libraries
import os
import time
import pyautogui as py
import sys
from dotenv import load_dotenv as ldv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl

# Load .env content
ldv()

# Fetch the variables
email = os.getenv("email")
password = os.getenv("password")

# Config. Chrome window
chrome_options = Options()
chrome_options.add_argument("--incognito")
chrome_options.add_experimental_option("detach", True)
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
servicey = Service(ChromeDriverManager().install())
browser = webdriver.Chrome(service=servicey, options=chrome_options)
browser.maximize_window()

# ==========================================
# OPEN TABS
# ==========================================
# 1st Email tab
browser.get("https://accounts.google.com/")
WebDriverWait(browser, 30).until(EC.visibility_of_element_located((By.ID, "identifierId"))).send_keys(email)
browser.find_element(By.ID, "identifierNext").click()
WebDriverWait(browser, 30).until(EC.visibility_of_element_located((By.XPATH, "//input[@type='password']"))).send_keys(password + Keys.ENTER)
time.sleep(5)
browser.get("https://mail.google.com/") 

# 2nd Forms Tab
browser.execute_script("window.open('about:blank', '_blank');")
browser.switch_to.window(browser.window_handles[1]) 
browser.get("https://forms.gle/SGy2pnLA9LdQfXeSA")

# ==========================================
# OPEN THE ARCHIVE
# ==========================================
py.hotkey('win', 'e')
time.sleep(1)
py.press(['tab' for _ in range(7)], interval=0.5)
time.sleep(2)
py.typewrite('Base_dados')
time.sleep(2)
py.moveTo(341, 162, duration=0.8)
py.click(x=341, y=162, clicks=1)
py.press('enter')

# ==========================================
# READ EXCEL DATA
# ==========================================
def read_excel_data(file_path, sheet_name, cell_coordinate):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]
    cell_value = sheet[cell_coordinate].value
    return cell_value

# Usage
file_path = r'C:\Users\mateus_s_cavalheiro\Downloads\Base_dados.xlsx'
one_value = read_excel_data(file_path, 'Página1', 'B2')
two_value = read_excel_data(file_path, 'Página1', 'B3')
three_value = read_excel_data(file_path, 'Página1', 'B4')
four_value = read_excel_data(file_path, 'Página1', 'B5')
five_value = read_excel_data(file_path, 'Página1', 'B6')
six_value = read_excel_data(file_path, 'Página1', 'B7')
seven_value = read_excel_data(file_path, 'Página1', 'B8')
eight_value = read_excel_data(file_path, 'Página1', 'B9')
nine_value = read_excel_data(file_path, 'Página1', 'B10')
ten_value = read_excel_data(file_path, 'Página1', 'B11')

# ==========================================
# FORM AUTOMATION (PyAutoGUI)
# ==========================================
# First Entry
time.sleep(3)
py.keyDown('alt')
time.sleep(0.2)
py.press('tab')
time.sleep(0.2)
py.press('tab')
time.sleep(1)
py.keyUp('alt')
time.sleep(2)
py.moveTo(390, 657, duration=0.8)
py.click(x=390, y=657, clicks=1)
time.sleep(2)
py.press('tab')
time.sleep(1)
py.press('enter')
time.sleep(1)
py.press('down')
time.sleep(1)
py.press('enter')
time.sleep(1)
py.press('tab')
py.typewrite(str(one_value))
time.sleep(1)
py.press('tab')
time.sleep(2)
py.moveTo(389, 699, duration=0.8)
py.click(x=389, y=699, clicks=1)
time.sleep(1)
py.press('tab')
time.sleep(1)
py.press('tab')
time.sleep(1)
py.press('enter')
time.sleep(1)

# Second Entry (Loop for the remaining data)
# Grouping the remaining values with their corresponding 'down' presses
remaining_data = [
    (2, two_value), (3, three_value), (4, four_value),
    (5, five_value), (6, six_value), (7, seven_value),
    (8, eight_value), (9, nine_value), (10, ten_value)
]

for downs, value in remaining_data:
    time.sleep(2)
    
    # 1. Press 2 tabs

    py.press('tab', presses=2, interval=1)
    
    # 2. Press down conforming the id
    time.sleep(1)
    py.press('enter')
    time.sleep(1)
    py.press('down', presses=downs, interval=0.2)
    time.sleep(1)
    
    py.press('enter') 
    
    # Go to type box
    time.sleep(1)
    py.press('tab')
    time.sleep(1)
    py.typewrite(str(value))
    
    # 3. Press 2 tabs
    time.sleep(1)   
    py.press('tab', presses=2, interval=1)
    
    # 4. If else logic
    # Convert valu to string
    try:
        numeric_value = float(value)
    except (ValueError, TypeError):
        numeric_value = 0
        
    time.sleep(1)
    if numeric_value >= 40:
        py.moveTo(387, 359, duration=0.8)
        py.click()
    else:
        py.moveTo(391, 399, duration=0.8)
        py.click()