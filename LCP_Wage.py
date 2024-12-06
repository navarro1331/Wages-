from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
import Web_login
import datetime
import log_in
import time
import pyautogui
import win32api
import win32con
import keyboard
import os
import shutil


def delete_all_in_folder(folder_path):
    # Verify the folder exists
    if not os.path.exists(folder_path):
        print(f"The folder '{folder_path}' does not exist.")
        return
    
    # Loop through all items in the folder
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        
        # Check if it is a file or a folder and delete accordingly
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.remove(file_path)  # Delete the file
                print(f"Deleted file: {file_path}")
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)  # Delete the directory and its contents
                print(f"Deleted directory: {file_path}")
        except Exception as e:
            print(f"Failed to delete {file_path}. Reason: {e}")


def switch_to_window(driver, keywords, timeout=30):

    try:
        time.sleep(5)
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            current_url = driver.current_url
            if any(keyword in current_url for keyword in keywords):
                print(f"Found window with URL containing one of {keywords}: {current_url}")
                WebDriverWait(driver, timeout).until(
                    EC.presence_of_element_located(('xpath', '//*[@id="toolbar"]')))
                print(f"Page loaded successfully with element ")
                return True

    except TimeoutException:
        print(f"Timeout: Could not load a page with XPath 'viewer' within {timeout} seconds.")
    
    print(f"No window found with a URL containing any of {keywords}")
    return False





folder_path = r"C:\Users\dsamu\dsamllc.net\dsamllc.net - Documents\FIS Project Documents\Testing_Jeff\Wages"  # Replace with the path to your folder
delete_all_in_folder(folder_path)



today = datetime.date.today()
formatted_date = today.strftime("%m/%d/%Y")

driver = log_in.lcp_login()
MultCPR = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='mainform']/table/tbody/tr/td[3]/p[2]/a/span")))
driver.execute_script("arguments[0].scrollIntoView(true);", MultCPR)
MultCPR.click()   
print("Clicked the CPR link successfully!")


def perform_cpr_report_search(driver, search_text, formatted_date, output_directory):
    """
    Perform CPR report search and download the report.

    Args:
        driver: Selenium WebDriver instance.
        search_text: The contractor's name to search for.
        formatted_date: The formatted date to set in the "To Date" field.
        output_directory: Directory path to save the report.
        from_date: Start date for the report (default: '01/01/2020').
    """
    try:
        # Locate and interact with the contractor dropdown
        select_list = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='Contractor_chosen']/a"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", select_list)
        select_list.click()

        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='Contractor_chosen']/div/div/input"))
        )
        search_box.send_keys(search_text)
        search_box.send_keys(Keys.RETURN)


        to_date = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='ToDate']")))
        to_date.clear()
        to_date.send_keys(formatted_date)
        to_date.send_keys(Keys.RETURN)

        from_date_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[@id='FromDate']")))
        from_date_field.clear()
        from_date_field.send_keys("01/01/2020")
        from_date_field.send_keys(Keys.RETURN)

        # Run the report
        run_button = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='runButton']"))
        )
        run_button.click()


        time.sleep(0.5)

        # Wait for the report to generate and switch to the correct tab
        # Usage Example
        # Usage Example
        try:
            # Check for PerformingReport or NonPerformingReport and ensure the viewer is loaded
            if not switch_to_window(driver, ["PerformingReport", "NonPerformingReport"],  timeout=30):
                print("Could not find the desired report tab or the page did not fully load.")
            else:
                print("Successfully switched to the desired report tab.")
        except Exception as e:
            print(f"An error occurred: {e}")





        for _ in range(4):
            keyboard.send("shift+tab")
            time.sleep(0.1)
        keyboard.send("enter")
        keyboard.send("tab")
        time.sleep(0.5)
        filename = f"PerformingReport_{search_text}"
        keyboard.write(filename)
        time.sleep(0.2)
        for _ in range(7):
            keyboard.send("shift+tab")
            time.sleep(0.1)
        keyboard.send("enter")
        keyboard.write(output_directory)
        time.sleep(0.2)
        keyboard.send("enter")
        for _ in range(5):
            keyboard.send("shift+tab")
            time.sleep(0.1)
        keyboard.send("enter")
        time.sleep(0.5)
        handles = driver.window_handles
        print(f"Available window handles: {handles}")
        if len(handles) > 2:
            # Scenario 1: More than 2 windows
            print("More than 2 windows detected. Closing excess windows.")
            driver.switch_to.window(handles[2])  # Switch to the third tab if it exists
            driver.close()
            if len(handles) > 1:  # Recheck handles after closing
                driver.switch_to.window(handles[1])  # Switch to the second tab if it exists
                driver.close()
            if len(handles) > 0:  # Recheck handles after closing
                driver.switch_to.window(handles[0])  # Switch back to the first tab if it exists

        elif len(handles) > 1:
            # Scenario 2: Exactly 2 windows
            print("Exactly 2 windows detected. Closing one and interacting with the remaining one.")
            driver.switch_to.window(handles[1])  # Switch to the second tab
            driver.close()
            driver.switch_to.window(handles[0])  # Switch back to the first tab
            # Wait for the 'Run' button to be available and click it
            run_button = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//*[@id='runButton']")))
            run_button.click()
            time.sleep(15)
            if not switch_to_window(driver, "PerformingReport"):
                print("Could not find the 'PerformingReport' tab.")
                return
            for _ in range(4):
                keyboard.send("shift+tab")
                time.sleep(0.1)
            keyboard.send("enter")
            keyboard.send("tab")
            time.sleep(0.5)
            filename = f"PerformingReport_{search_text}"
            keyboard.write(filename)
            time.sleep(0.2)
            for _ in range(7):
                keyboard.send("shift+tab")
                time.sleep(0.1)
            keyboard.send("enter")
            keyboard.write(output_directory)
            time.sleep(0.2)
            keyboard.send("enter")
            for _ in range(5):
                keyboard.send("shift+tab")
                time.sleep(0.1)
            keyboard.send("enter")
            time.sleep(0.5)
        else: print("Only one window detected. No additional actions required.")

    except Exception as e:
        print(f"An error occurred: {e}")








# Define parameters
search_text = "Elevator Repair Service, Inc."
output_directory = r"C:\Users\dsamu\dsamllc.net\dsamllc.net - Documents\FIS Project Documents\Testing_Jeff\Wages"

# Call the function
perform_cpr_report_search(driver, search_text, formatted_date, output_directory)


time.sleep(1)


# Define parameters
search_text = "DIFFCO, LLC"
output_directory = r"C:\Users\dsamu\dsamllc.net\dsamllc.net - Documents\FIS Project Documents\Testing_Jeff\Wages"

# Call the function
perform_cpr_report_search(driver, search_text, formatted_date, output_directory)
time.sleep(1)

# Define parameters
search_text = "TAG Companies"
output_directory = r"C:\Users\dsamu\dsamllc.net\dsamllc.net - Documents\FIS Project Documents\Testing_Jeff\Wages"

# Call the function
perform_cpr_report_search(driver, search_text, formatted_date, output_directory)
time.sleep(1)

input("Press Enter to close the browser...")
driver.quit()



















