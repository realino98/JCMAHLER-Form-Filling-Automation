from webbrowser import Chrome
from numpy import real, true_divide
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import pandas as pd
import pyautogui
import requests
import time
import os
from sys import exit


#Global Variables
os.system("cls")
print("Initializing")
root = os.getcwd()
PATH = os.getcwd() + '/webdriver/chromedriver.exe'
option = webdriver.ChromeOptions()
option.add_experimental_option('excludeSwitches', ['enable-logging'])
driver_service = Service(executable_path=PATH)
# driver_service = Service(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=driver_service, service_log_path=None, options=option)
# driver.set_window_size(1920, 1080, driver.window_handles[0])
driver.minimize_window()

URL = "https://intranet.hlb.global"
DELAY_TIME = 60

def defaultCsv():
    csv_file = root + '/data/data.csv'
    df = pd.read_csv(csv_file)
    return df

def defaultExcel():
    excel = pd.ExcelFile("data/data.xlsx")
    df = excel.parse("Main")
    return df

def loading():
    for i in range(4):
        print('.', end='')
        time.sleep(0.3)
    print()
def cls():
    os.system("cls")

def executeFormFill(df, data_no):
    driver.maximize_window()
    driver.get(URL)
    driver.find_element(By.NAME, "firstname").send_keys(df["firstname"][data_no])
    driver.find_element(By.NAME, "lastname").send_keys(df["lastname"][data_no])
    driver.find_element(By.NAME, "email").send_keys(df["email"][data_no])
    #company
    
    driver.find_element(By.CSS_SELECTOR, "#wpcf7-cf7sg-form-intranet-registrations_copy > div:nth-child(5) > div > div > div > span > span > span.selection > span").click()
    driver.find_element(By.CSS_SELECTOR, "body > span > span > span.select2-search.select2-search--dropdown > input").send_keys(df["company"][data_no].replace("&amp;", "&"), Keys.ENTER)
    #partner
    if df["partner"][data_no] == "NO":
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div/article/div/div/div[2]/div/div/div/div/div/div/form/div[6]/div/div/div/div/span/span/span[2]/label").click()
    else:
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div/article/div/div/div[2]/div/div/div/div/div/div/form/div[6]/div/div/div/div/span/span/span[1]/label").click()
        pass
    #interest
    # driver.find_element(By.XPATH, "/html/body/div[1]/div/div/article/div/div/div[2]/div/div/div/div/div/div/form/div[7]/div/div/div/span/span/span[1]/span/ul/li/input").click()
    driver.find_element(By.XPATH, "/html/body/div[1]/div/div/article/div/div/div[2]/div/div/div/div/div/div/form/div[7]/div/div/div/span/span/span[1]/span/ul/li/input").send_keys(df["interest"][data_no], Keys.ENTER)

    driver.find_element(By.CSS_SELECTOR, "#registerbutton").click()
    # driver.minimize_window()
    for i in range(DELAY_TIME):
        cls()
        print("Data no", data_no, "OK")
        print("waiting for next data in", DELAY_TIME-i, "seconds")
        time.sleep(1)

def menu():
    cls()
    print("Jcmahler - Automation")
    print("Menu:")
    print("1. Help")
    print("2. View Data")
    print("3. Execute")
    print("4. Refresh Data")
    print("0. Exit")

def mainMenu():
    df = defaultExcel() #from data folder
    menu()
    option = input('Choose option: ')
    while option != "0":
        if option == "1":
            cls()
            print("Read the Readme.md")
            cont = input("Press any key to continue...")
            
        elif option == "2":
            if df is not None:
                print(df)
                loading()
                cont = input("Press any key to continue...")
            else:
                print('no data')
                loading()
        elif option == "3":
            cls()
            print("Fetching Data", end='')
            loading()
            if df is not None:
                data_available = True
            else:
                data_available = False

            if data_available:
                cls()
                for i in range(len(df)):
                    executeFormFill(df, i)
            else:
                cls()
                print("Data not available")
            loading()
        elif option == "4":
            cls()
            print("Refreshing Data", end='')
            df = defaultExcel()
            loading()
            print("Data Refreshed")
            cont = input("Press any key to continue...")
        elif option == "0":
            print("Thank you for using the program")
            loading()
            break
        else:
            cls()
            print("invalid input")
            loading()
        menu()
        option = input('Choose option: ')
    cls()
    print("Thank you for using the program")
    loading()
mainMenu()

