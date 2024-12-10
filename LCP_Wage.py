import os
import shutil
import time
import pyautogui
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import log_in
import datetime
import re
import pandas as pd


def delete_all_in_folder(folder_path):
    if not os.path.exists(folder_path):
        return  
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.remove(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f"Failed to delete {file_path}: {e}")


def perform_cpr_report_search(driver, search_text, formatted_date, output_directory, from_date="01/01/2020"):
   
    original_window = driver.current_window_handle  

    try:
        select_list = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='Contractor_chosen']/a")))
        driver.execute_script("arguments[0].scrollIntoView(true);", select_list)
        select_list.click()
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='Contractor_chosen']/div/div/input")))
        search_box.send_keys(search_text)
        search_box.send_keys(Keys.RETURN)
        to_date = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='ToDate']")))
        to_date.clear()
        to_date.send_keys(formatted_date)
        to_date.send_keys(Keys.RETURN)
        from_date_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='FromDate']")))
        from_date_field.clear()
        from_date_field.send_keys(from_date)
        from_date_field.send_keys(Keys.RETURN)
        run_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='runButton']")))
        run_button.click()
        time.sleep(5)
        print("Clicked the runButton successfully!")
        time.sleep(5)
        try:
            original_window = driver.current_window_handle
            print(f"Original window handle: {original_window}")
        except Exception as e:
            print(f"Error while getting the current window handle: {e}")      
        time.sleep(5)       
        safe_search_text = re.sub(r'[<>:"/\\|?*]', '_', search_text)
        filename = f"PerformingReport_{safe_search_text}.pdf"
        file_path = os.path.join(output_directory, filename)  
        os.makedirs(output_directory, exist_ok=True)
        print("output_directory successfully3!")    
        time.sleep(15)
        if len(driver.window_handles) == 3:
            print("There are exactly 3 tabs open.")
            for handle in driver.window_handles:
                driver.switch_to.window(handle) 
                current_url = driver.current_url
                print(f"Checking tab with URL: {current_url}")
                if current_url.endswith("MultipleCprsReport/PerformingReport"):
                    print(f"Found the tab with URL ending in 'MultipleCprsReport/PerformingReport': {current_url}")
                    break  
        else:
            print(f"There are {len(driver.window_handles)} tabs open.")
            if len(driver.window_handles) > 1:
                driver.switch_to.window(driver.window_handles[1])
                driver.close()
                print("Closed the second tab.")
                driver.switch_to.window(driver.window_handles[0])
                run_MultCPR = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='runButton']")))
                run_MultCPR.click()
                print("Clicked on the 'Run' button for MultiCPR.")
                time.sleep(15)
                for handle in driver.window_handles:
                    driver.switch_to.window(handle) 
                    current_url = driver.current_url
                    print(f"Checking tab with URL: {current_url}")
                    if current_url.endswith("MultipleCprsReport/PerformingReport"):
                        print(f"Found the tab with URL ending in 'MultipleCprsReport/PerformingReport': {current_url}")
                        break    


           
        driver.implicitly_wait(5)
        for _ in range(4):
            pyautogui.hotkey("shift", "tab")
            time.sleep(0.2)
        pyautogui.press("enter")
        time.sleep(1)
        pyautogui.hotkey("ctrl", "a")
        pyautogui.write(file_path)
        time.sleep(1)
        pyautogui.press("enter")
        print(f"Report saved successfully: {file_path}")


    except TimeoutException as e:
        print(f"Timeout occurred: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Close all windows except the original
        for handle in driver.window_handles:
            if handle != original_window:
                driver.switch_to.window(handle)
                driver.close()
        driver.switch_to.window(original_window)


if __name__ == "__main__":
    # Cleanup the folder
    folder_path = r"C:\Users\dsamu\dsamllc.net\dsamllc.net - Documents\FIS Project Documents\Testing_Jeff\Wages"
    delete_all_in_folder(folder_path)

    # Login to LCP Tracker
    driver = log_in.lcp_login()
    MultCPR = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='mainform']/table/tbody/tr/td[3]/p[2]/a/span")))
    driver.execute_script("arguments[0].scrollIntoView(true);", MultCPR)
    MultCPR.click()  
    print("Clicked the CPR link successfully!")
   
    formatted_date = datetime.date.today().strftime("%m/%d/%Y")
    print("F. time successfully!")
   
    # Path to the Excel file
    file_path = r"C:\Users\dsamu\dsamllc.net\dsamllc.net - Documents\FIS Project Documents\1Power Bi\PayrollInfo.xlsx"

    # Read the specific sheet and column
    sheet_name = "Payroll Details (2)"
    column_name = "Contractor Name"
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    contractors = df[column_name].drop_duplicates().dropna().tolist()
   
   
    for contractor in contractors:
        perform_cpr_report_search(driver, contractor, formatted_date, folder_path)
        print(f"File saved: {contractor}")

    print("All reports have been processed. You may now close the browser.")
    input("Press Enter to exit...")
    driver.quit()


    # Read contractor names from the DataFrame
    sheet_name = "Payroll Details (2)"
    column_name = "Contractor Name"
    
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    contractors = df[column_name].drop_duplicates().dropna().tolist()
    

    # Perform report searches for each contractor
    for contractor in contractors:
        safe_contractor_name = re.sub(r'[<>:"/\\|?*]', '_', contractor)
        pdf_path = rf"C:\Users\dsamu\dsamllc.net\dsamllc.net - Documents\FIS Project Documents\Testing_Jeff\Wages\PerformingReport_{safe_contractor_name}.pdf"
        print(f"Processing report for contractor: {contractor}")




