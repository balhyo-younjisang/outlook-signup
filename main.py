import email
from itertools import count
from re import S
from this import d
from tkinter.colorchooser import Chooser
from tokenize import Number
from webbrowser import Chrome
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import time
from random import *
import random
import string
from strgen import StringGenerator
import urllib.request
import pyperclip

################function##################
def clipboard_input(user_xpath, user_input):
        temp_user_input = pyperclip.paste()  # 사용자 클립보드를 따로 저장

        pyperclip.copy(user_input)
        # driver.find_element_by_xpath(user_xpath).click()
        driver.find_element(By.XPATH, user_xpath).click()
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()

        pyperclip.copy(temp_user_input)  # 사용자 클립보드에 저장 된 내용을 다시 가져 옴
        time.sleep(2)


############URL##########################
#chromedriver url
driverURL = Service("chromedriver.exe")
#microsoft outlook URL
msURL = "https://www.microsoft.com/ko-kr/microsoft-365/outlook/email-and-calendar-software-microsoft-outlook"

##############init###########
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option('detach', True)
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
# to maximize the browser window
driver.maximize_window()

###############microsoft outlook mail sign up########################
#Open ms browser
driver.get(msURL)

#Click the signup button
driver.find_element('xpath', '//*[@id="office-Hero5050-e0h0pts"]/section/div[1]/div[1]/div/div/div/div/div[1]/a').click()

time.sleep(5)

#Change to tab
driver.switch_to.window(driver.window_handles[1])

driver.find_element('xpath', '//*[@id="koreaConsentDataCollectionLearnMore"]').click()

time.sleep(5)

driver.find_element('xpath', '/html/body/div[1]/div/div/div[2]/div/div[2]/div[3]/div/div[1]/div[5]/div/div/form/input').click()

time.sleep(5)

driver.find_element('xpath', '//*[@id="koreaConsentDataProvisionLearnMore"]').click()

time.sleep(5)

driver.find_element('xpath', '/html/body/div[1]/div/div/div[2]/div/div[2]/div[3]/div/div[1]/div[5]/div/div/form/input').click()

time.sleep(5)
driver.find_element('xpath', '//*[@id="iSignupAction"]').click()

#Make an info
mail_address = ""
password = ""
length=20
first_name = ""
last_name = ""
name_length = 5

for i in range(length):
    mail_address += str(random.choice(string.ascii_letters))
    password += str(random.choice(string.ascii_letters))

for i in range(name_length):
    first_name += str(random.choice(string.ascii_letters))
    last_name += str(random.choice(string.ascii_letters))

time.sleep(5)
# input_form = driver.find_element('xpath', '//*[@id="MemberName"]')
# input_form.send_keys(mail_address)
clipboard_input('//*[@id="MemberName"]', mail_address)

driver.find_element(By.XPATH, '//*[@id="iSignupAction"]').click()

time.sleep(5)
# driver.find_element(By.XPATH, '//*[@id="PasswordInput"]').send_keys(password)
clipboard_input('//*[@id="PasswordInput"]', password)

driver.find_element(By.XPATH, '//*[@id="iConsentAction"]').click()

time.sleep(5)
# driver.find_element(By.XPATH, '//*[@id="LastName"]').send_keys(last_name)
clipboard_input('//*[@id="LastName"]', last_name)
# driver.find_element(By.XPATH, '//*[@id="FirstName"]').send_keys(first_name)
clipboard_input('//*[@id="FirstName"]', first_name)
driver.find_element(By.XPATH, '//*[@id="iSignupAction"]').click()
time.sleep(5)
# driver.find_element(By.XPATH, '//*[@id="BirthYear"]').send_keys(randint(1980, 1999))
clipboard_input('//*[@id="BirthYear"]', randint(1980, 1999))
selectMonth = Select(driver.find_element(By.XPATH, '//*[@id="BirthMonth"]'))
selectMonth.select_by_index(randint(1, 12))
time.sleep(1)
selectDay = Select(driver.find_element(By.XPATH, '//*[@id="BirthDay"]'))
selectDay.select_by_index(randint(1, 31))
#browser close
# driver.close()