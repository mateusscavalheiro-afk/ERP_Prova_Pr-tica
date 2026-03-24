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
time.sleep(5) # Give the form a moment to fully render              

# ==========================================
# READ EXCEL DATA (Happens in the background!)
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
# FORM AUTOMATION (PyAutoGUI - Keyboard Only)
# ==========================================

# --- GET FOCUS ON THE FIRST DROPDOWN ---
# Using coordinates here to bypass the "Confirm Email" box and focus the form
py.moveTo(390, 657, duration=0.8)
py.click(x=390, y=657, clicks=1)
time.sleep(1)

# --- FIRST ENTRY (Page 1) ---
py.press('tab')
time.sleep(0.5)
py.press('down')
time.sleep(0.5)
py.press('enter')
time.sleep(0.5)

# Tab to type box
py.press('tab')
time.sleep(0.5)
py.typewrite(str(one_value))
time.sleep(0.5)

# Logic for first entry radio button
try:
    num_val_1 = float(one_value)
except (ValueError, TypeError):
    num_val_1 = 0

time.sleep(1)
if num_val_1 >= 40:
    # If yes: one tab, space
    py.press('tab')
    time.sleep(0.5)
    py.press('space')
else:
    # If no: two tabs, space
    py.press('tab', presses=2, interval=0.5)
    time.sleep(0.5)
    py.press('space')

time.sleep(1)

# --- GO TO PAGE 2 (Click Avançar) ---
# Tabbing past "Limpar seleção" to hit "Avançar"
py.press('tab', presses=2, interval=0.5)
time.sleep(0.5)
py.press('enter')
time.sleep(2.5) # Wait for Page 2 to fully load

# ==========================================
# --- REMAINING ENTRIES LOOP (Pages 2 through 10) ---
# ==========================================
remaining_data = [
    (2, two_value), (3, three_value), (4, four_value),
    (5, five_value), (6, six_value), (7, seven_value),
    (8, eight_value), (9, nine_value), (10, ten_value)
]

for downs, value in remaining_data:
    # 1. Tab into the NEW page's dropdown 
    # (Adjust this number if a fresh page requires more tabs to hit the first question)
    py.press('tab', presses=2, interval=0.5)

    # 2. Open dropdown and press down corresponding to the id
    time.sleep(0.5)
    py.press('enter') # Opens dropdown
    time.sleep(0.5)
    py.press('down', presses=downs, interval=0.2)
    time.sleep(0.5)
    py.press('enter') # Confirms dropdown selection
    
    # 3. Tab to type box
    time.sleep(0.5)
    py.press('tab')
    time.sleep(0.5)
    py.typewrite(str(value))
    
    # 4. If else logic (Tab + Space navigation)
    try:
        numeric_value = float(value)
    except (ValueError, TypeError):
        numeric_value = 0
        
    time.sleep(1)
    if numeric_value >= 40:
        # If yes: one tab, space
        py.press('tab')
        time.sleep(0.5)
        py.press('space') 
        py.press('tab')
    else:
        # If no: two tabs, space
        py.press('tab', presses=2, interval=0.5)
        time.sleep(0.5)
        py.press('space')
        
    time.sleep(1)

    # 5. Click Avançar for the loop to load the next page
    py.press('tab', presses=2, interval=0.5)
    time.sleep(0.5)
    py.press('enter') # Clicks Avançar
    time.sleep(2.5) # Wait for the next page to load before restarting loop