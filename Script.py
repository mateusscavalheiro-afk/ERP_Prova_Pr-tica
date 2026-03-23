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
# OPEN THE ARCHIVE
# ==========================================
py.hotkey('win', 'e')
time.sleep(1)
py.press(['tab' for _ in range(7)], interval=0.5)
time.sleep(2)
py.typewrite('Base_dados')
time.sleep(2)
# It is generally safer to hit enter to open rather than specific coordinates, 
# but if this works reliably for your specific screen, we'll leave it.
py.moveTo(341, 162, duration=0.8)
py.click(x=341, y=162, clicks=1)
py.press('enter')
time.sleep(3) # Wait for Excel to open

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
# FORM AUTOMATION (PyAutoGUI - Keyboard Only)
# ==========================================

# Switch back to the Chrome window
py.keyDown('alt')
time.sleep(0.2)
py.press('tab', presses=2, interval=0.2)
time.sleep(1)
py.keyUp('alt')
time.sleep(2)

# Click the VERY FIRST dropdown just to get the form in focus.
# Because the page hasn't scrolled yet, this single coordinate click is safe.
py.moveTo(390, 657, duration=0.8)
py.click(x=390, y=657, clicks=1)
time.sleep(1)

# --- FIRST ENTRY ---
# The dropdown is already open from the click above
py.press('down')
time.sleep(0.5)
py.press('enter')
time.sleep(0.5)

# Tab to type box
py.press('tab')
time.sleep(0.5)
py.typewrite(str(one_value))
time.sleep(0.5)

# Tab to radio buttons
py.press('tab', presses=2, interval=0.5)
time.sleep(1)

# Logic for first entry radio button
try:
    num_val_1 = float(one_value)
except (ValueError, TypeError):
    num_val_1 = 0

if num_val_1 >= 40:
    py.press('space') # Selects the first option
else:
    py.press('down')  # Moves to the second option
    time.sleep(0.5)
    py.press('space') # Selects it

time.sleep(1)

# --- REMAINING ENTRIES LOOP ---
remaining_data = [
    (2, two_value), (3, three_value), (4, four_value),
    (5, five_value), (6, six_value), (7, seven_value),
    (8, eight_value), (9, nine_value), (10, ten_value)
]

for downs, value in remaining_data:
    time.sleep(1.5)

    # 1. Tab to the next dropdown 
    # (Adjust the number of tabs here if it doesn't land exactly on the dropdown)
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
    
    # 4. Tab to radio buttons
    time.sleep(0.5)   
    py.press('tab', presses=2, interval=0.5)
    
    # 5. If else logic (Keyboard navigation)
    try:
        numeric_value = float(value)
    except (ValueError, TypeError):
        numeric_value = 0
        
    time.sleep(1)
    if numeric_value >= 40:
        # Pressing space selects the currently focused radio button
        py.press('space') 
    else:
        # Pressing down moves to the next radio button in the group
        py.press('down')
        time.sleep(0.5)
        py.press('space')