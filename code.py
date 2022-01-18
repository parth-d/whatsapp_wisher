from datetime import date
import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time
import sys


def close_windows():
    windows = browser.window_handles

    for w in windows:
        browser.switch_to.window(w)
        browser.close()
    browser.quit()


def reg_num(stri):
    sto = ""
    for char in str(stri):
        if char.isdigit():
            sto += char
    if len(sto) == 12:
        return sto
    elif len(sto) == 10:
        return "91" + sto
    else:
        return "Invalid Number"


todayname = []
todaynum = []
today = date.today()
d1 = today.strftime("%d/%m")
print("Today's date:", d1)
data = pd.read_excel(r'Patients.xlsx')
name = data['Name'].tolist()
bd = data['Birthday'].tolist()
num = data['Number'].tolist()
for i in range(len(bd)):
    if bd[i] == d1:
        todayname.append(name[i])
        numb = reg_num(num[i])
        if numb == "Invalid Number":
            print("Invalid Number for %s" % name[i])
            sys.exit()
        else:
            todaynum.append(numb)
if len(todayname) == 0:
    print("No birthdays today")
else:
    print("Today's Birthdays:")
    for elem in todayname:
        print(elem)
    print("Opening WhatsApp web to send messages")
    options = webdriver.ChromeOptions()
    options.add_argument(
        # enter the user name as per requirement
        "user-data-dir=C:\\Users\\vaish\\AppData\\Local\\Google\\Chrome\\User Data")
    browser = webdriver.Chrome(executable_path=r'chromedriver.exe', options=options)
    # browser.minimize_window()
    wait = WebDriverWait(browser, 15000)
    for i in range(len(todayname)):
        link = "https://web.whatsapp.com/send?phone={}&text&source&data&app_absent".format(todaynum[i])
        string = "To %s,\n Wishing you many many happy returns of the day!" % todayname[i]
        browser.get(link)
        try:
            x_arg = '//*[@id="main"]/footer/div[1]/div[2]/div/div[2]'
            target = wait.until(ec.presence_of_element_located((By.XPATH, x_arg)))
            target.click()
            print("WhatsApp Web opened")
            print("Sending message to", todayname[i])
            input_box = browser.find_element('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
            for ch in string:
                if ch == "\n":
                    ActionChains(browser).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(
                        Keys.SHIFT).key_up(Keys.BACKSPACE).perform()
                else:
                    input_box.send_keys(ch)
            input_box.send_keys(Keys.ENTER)
            print("Message sent successfully to ", todayname[i])
            time.sleep(3)  # sleep timer for giving time for the message to be sent
        except NoSuchElementException:
            print("Failed to send message")
close_windows()
# input()
