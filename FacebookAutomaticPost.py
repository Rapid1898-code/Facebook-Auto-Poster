# pyinstaller --onefile --hidden-import pycountry --exclude-module matplotlib FacebookAutomaticPost.py

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import time
import sys, os
from sys import platform
import xlwings as xw
from datetime import datetime, timedelta

def login(id,password):
    email = driver.find_element_by_id("email")
    email.send_keys(id)
    Password = driver.find_element_by_id("pass")
    Password.send_keys(password)
    button = driver.find_element_by_name("login").click()
    pass

# read input excel
FN = "FacebookAutomaticPost.xlsx"
path = os.path.abspath (os.path.dirname (sys.argv[0]))
fn = path + "/" + FN
wb = xw.Book (fn)
ws = wb.sheets["Overview"]
ws2 = wb.sheets["Texte"]
listGroups = ws.range ("B5:B100").value
for idx, cont in enumerate (listGroups):
  if cont == None:
    maxRow = int (idx) + 4
    break
idxRow = 2

# login to facebook
link = "https://www.facebook.com/"
options = Options()
options.add_argument("--disable-infobars")
options.add_argument("start-maximized")
options.add_argument("--disable-extensions")
options.add_argument("--disable-notifications")
options.add_argument("--disable-application-cache")
if ws["B3"].value.upper() not in ["YES","JA","J","Y"]:
    options.add_argument('--headless')
options.add_experimental_option ('excludeSwitches', ['enable-logging'])
path = os.path.abspath (os.path.dirname (sys.argv[0]))
if platform == "win32": cd = '/chromedriver.exe'
elif platform == "linux": cd = '/chromedriver_linux'
elif platform == "darwin": cd = '/chromedriver'
driver = webdriver.Chrome (path + cd, options=options)
driver.get (link)  # Read link
time.sleep (3)
# driver.find_element_by_xpath ("//*[@id="u_0_j_Zi"]").click ()
driver.find_element_by_css_selector("[title*='Alle akzeptieren']").click()
print(f"Press on <Accept All> button...")
time.sleep (3)
login(ws["B1"].value,ws["B2"].value )
print(f"Login To Facebook...")
time.sleep (3)

for idx, stock in enumerate (listGroups):
    if idx == maxRow:
        break
    linkcode = ws["B" + str (idxRow)].value
    textCell = ws["D" + str (idxRow)].value
    dateCell = ws["E" + str (idxRow)].value
    print("Debug: ", linkcode,textCell,dateCell)

    if textCell not in [None,""] and dateCell in [None,""]:
        text = ws2[textCell].value
    else:
        idxRow += 1
        continue

    dt1 = datetime.today ()
    dt1 = datetime.strftime (dt1, "%Y-%m-%d")

    if dateCell in [None,""]:
        try:
            link2 = "https://www.facebook.com/groups" + linkcode
            driver.get (link2)
            print(f"Working on link {link2}...")
            time.sleep (3)
            driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[4]/div/div/div/div/div[1]/div[1]/div/div/div/div[1]/div/div[1]/span').click()
            print(f"Click in Message field...")
            time.sleep (3)
            driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/div/div[4]/div/div/div[1]/div/div[2]/div/div/div/div/div[1]/form/div/div[1]/div/div/div[1]/div[2]/div[1]/div[1]/div[1]/div/div/div/div/div[2]/div/div/div/div')\
            .send_keys(text)
            print(f"Write text to message box...")
            time.sleep (3)
            driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/div/div[4]/div/div/div[1]/div/div[2]/div/div/div/div/div[1]/form/div/div[1]/div/div/div[1]/div[3]/div[2]/div/div/div[1]').click()
            print(f"Send message...")
            time.sleep (10)
            ws["E" + str (idxRow)].value = dt1
            print(f"Message send on Facebook for row {idxRow} site {linkcode} for cell {textCell}")
            idxRow += 1
        except Exception as e:
            print(f"Error while working on row {idxRow} for site {linkcode} - wrote error to column E...")
            print(f"ErrorCode: {e}")
            ws["E" + str (idxRow)].value = "Error"
            idxRow += 1
driver.quit()
