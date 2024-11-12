from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import Web_login
import datetime
import time
import requests
import fitz

today = datetime.date.today()
formatted_date = today.strftime("%m/%d/%Y")

driver = Web_login.lcp_login()
MultCPR = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='mainform']/table/tbody/tr/td[3]/p[2]/a/span")))
driver.execute_script("arguments[0].scrollIntoView(true);", MultCPR)
MultCPR.click()   
print("Clicked the CPR link successfully!")
select_list = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='Contractor_chosen']/a")))
driver.execute_script("arguments[0].scrollIntoView(true);", select_list)
select_list.click()   
search_box = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='Contractor_chosen']/div/div/input")))

search_text = "Elevator Repair Service, Inc."  # Replace with the actual text you want to search for

search_box.send_keys(search_text)
search_box.send_keys(Keys.RETURN)
To_Date = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ToDate']")))
To_Date.send_keys(formatted_date)
To_Date.send_keys(Keys.RETURN)
From_Date = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='FromDate']")))
From_Date.send_keys('01/01/2020')
From_Date.send_keys(Keys.RETURN)    
run_MultCPR = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='runButton']")))
run_MultCPR.click()



try:
    download_controls = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "downloads")))
    download_controls.click()
    print("Clicked downloads.")
except Exception as e:
    print(f"An occurred: {e}")



input("Press Enter to close the browser...")
driver.quit()


















