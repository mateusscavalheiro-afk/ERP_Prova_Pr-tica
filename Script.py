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

# OPEN TABS
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

# Open the archieve
py.press("win")
py.typewrite("gerenciador de arquivos")
py.press("enter")
time.sleep(3)
py.press(['tab', 'tab', 'tab', 'tab', 'tab', 'tab', 'tab'], interval=0.5)
time.sleep(2)
py.typewrite('Base_dados')
time.sleep(2)
py.moveTo(341, 162, duration=0.8)
py.click(x=341, y=162, clicks=1)
py.press('enter')

# Read the variables of the spreadsheet
