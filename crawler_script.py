import os
import sys
import json
import time
import getpass
import pandas as pd
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- 1. USER CONFIGURATION ---
# ! YOU MUST UPDATE THESE VALUES !

# --- URLs ---
# The page where you log in (Microsoft sign-in)
LOGIN_URL = "https://login.microsoftonline.com/78aac226-2f03-4b4d-9037-b46d56c55210/oauth2/v2.0/authorize?response_type=code&client_id=c728a481-39c5-4916-86a8-0c8bc65d42b6&state=NX5VeTNTSFhPZENoeTduQ1FYY3pITmxKN1ZvVDJYeHJmLXZmQ2FpYllNY1hi%3B%252Fresults%252Fpeople&redirect_uri=https%3A%2F%2Fauthdirectory.utoronto.ca%2Findex.html&scope=openid%20profile%20email%20offline_access&code_challenge=6zw-jjMe-HaB-zhQsM0mEqWnVKv1AWxp2wp982T_cTE&code_challenge_method=S256&nonce=NX5VeTNTSFhPZENoeTduQ1FYY3pITmxKN1ZvVDJYeHJmLXZmQ2FpYllNY1hi&sso_reload=true"
# The page you land on *after* login, where you can start searching
SEARCH_PAGE_URL = "https://authdirectory.utoronto.ca/results/people"

# --- Selectors for Multi-Stage Login ---
# NO LONGER NEEDED FOR MANUAL LOGIN
# MS_EMAIL_FIELD_ID = "i0116"
# MS_EMAIL_NEXT_BUTTON_ID = "idSIButton9"
# UOFT_USERNAME_FIELD_ID = "username"
# UOFT_PASSWORD_FIELD_ID = "password"
# UOFT_LOGIN_BUTTON_ID = "login-btn"
# DUO_IFRAME_ID = "duo_iframe" 
# DUO_TRUST_DEVICE_BUTTON_ID = "trust-browser-button"
# MS_STAY_SIGNED_IN_BUTTON_ID = "idSIButton9" 

# --- Selectors for the search page (from your HTML) ---
DEPARTMENT_FIELD_NAME = "department" # (name="department")
SEARCH_BUTTON_XPATH = "//button[text()='Search']"
EXPORT_CSV_BUTTON_XPATH = "//button[contains(text(), 'Export Results as CSV')]"

# --- Data ---
# Add all the departments you want to query, exactly as they need to be typed
if len(sys.argv) > 1:
    try:
        DEPARTMENT_LIST = json.loads(sys.argv[1])
    except Exception as e:
        print("Error loading department list from arguments:", e)
        DEPARTMENT_LIST = []
else:
	DEPARTMENT_LIST = []

# --- File Paths ---
# This script will create a 'downloads' folder in the same directory to store the files
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
FINAL_EXCEL_FILE = os.path.join(DOWNLOAD_DIR, "UofT_Staff_Report.xlsx")
# The default name of the file downloaded from the website (from your screenshot)
DEFAULT_DOWNLOAD_FILENAME = "People_Results.csv" 

# ----------------------------------
# END OF CONFIGURATION
# ----------------------------------

def setup_driver(download_path):
    """Configures and returns a Selenium Chrome webdriver."""
    if not os.path.exists(download_path):
        os.makedirs(download_path)
        print(f"Created download directory: {download_path}")

    chrome_options = webdriver.ChromeOptions()
    # Set the default download directory for this browser session
    prefs = {"download.default_directory": download_path}
    chrome_options.add_experimental_option("prefs", prefs)
    
    print("Setting up Chrome driver...")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    # Use a longer implicit wait for this complex login
    driver.implicitly_wait(10) 
    print("Driver setup complete.")
    return driver

# --- login_to_site function is removed ---

def download_department_csv(driver, department_name, download_dir):
    """Searches for a department, downloads the CSV, and renames it."""
    try:
        print(f"  Navigating to search page for '{department_name}'...")
        # Ensure we are on the search page before every search
        driver.get(SEARCH_PAGE_URL) 

        # Wait for the department field to be ready
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, DEPARTMENT_FIELD_NAME))
        )

        # --- Clean up old download ---
        # Delete any existing 'results.csv' to ensure we get a new one
        default_file_path = os.path.join(download_dir, DEFAULT_DOWNLOAD_FILENAME)
        if os.path.exists(default_file_path):
            os.remove(default_file_path)
            # print(f"  Removed old '{DEFAULT_DOWNLOAD_FILENAME}'.") # Quieter log

        # --- Fill and submit search form ---
        print(f"  Searching for '{department_name}'...")
        # Clear the field first in case it has old data
        dept_input = driver.find_element(By.NAME, DEPARTMENT_FIELD_NAME)
        dept_input.clear()
        dept_input.send_keys(department_name)
        
        # Wait a brief moment for any autosuggestions to appear (if any)
        time.sleep(0.5) # This helps ensure the page processes the input
        
        # Send the Enter key to submit the search, as requested
        dept_input.send_keys(Keys.ENTER)
        
        #driver.find_element(By.XPATH, SEARCH_BUTTON_XPATH).click() # Removed this click

        # --- Click download button ---
        # Wait for the "Export Results as CSV" button to appear
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, EXPORT_CSV_BUTTON_XPATH))
        )
        print(f"  Clicking download link...")
        driver.find_element(By.XPATH, EXPORT_CSV_BUTTON_XPATH).click()

        # --- Wait for download to complete ---
        print(f"  Waiting for download...")
        timeout = 30  # 30 seconds
        start_time = time.time()
        while not os.path.exists(default_file_path):
            time.sleep(0.5)
            if time.time() - start_time > timeout:
                raise Exception(f"Download timed out for '{department_name}'")
        
        print(f"  Download complete.")

        # --- Rename the file ---
        new_file_path = os.path.join(download_dir, f"{department_name}.csv")
        if os.path.exists(new_file_path):
            os.remove(new_file_path) # Remove old file if it exists
        
        os.rename(default_file_path, new_file_path)
        print(f"  Renamed file to '{department_name}.csv'.")
        return new_file_path

    except Exception as e:
        print(f"\n--- !! FAILED to download for '{department_name}' !! ---")
        print(f"Error: {e}")
        print("Please check your selectors and make sure the page structure hasn't changed.")
        return None

def combine_csvs_to_excel(csv_files_dict, excel_path):
    """Finds all CSVs in the download directory and combines them into one sheet."""
    
    print(f"\nLooking for CSV files in '{DOWNLOAD_DIR}'...")
    
    # Use glob to find all .csv files in the download directory
    # This finds files like "Faculty of Law.csv", "Department of Physics.csv", etc.
    csv_file_list = glob.glob(os.path.join(DOWNLOAD_DIR, "*.csv"))
    
    if not csv_file_list:
        print("No CSV files found in the download directory. Exiting.")
        return

    print(f"Found {len(csv_file_list)} CSV files. Combining them...")
    all_dataframes = []
    
    try:
        for file_path in csv_file_list:
            # Get the department name from the file name
            # e.g., "C:\...\downloads\Faculty of Law.csv" -> "Faculty of Law"
            department_name = os.path.basename(file_path).replace('.csv', '')
            try:
                df = pd.read_csv(file_path)
                # Add a new column to identify the source department
                df['Source Department'] = department_name
                all_dataframes.append(df)
                print(f"  Loaded data for '{department_name}'")
            except Exception as e:
                print(f"  --- !! FAILED to read or process '{file_path}' !! ---")
                print(f"  Error: {e}")
        
        if not all_dataframes:
            print("No data was loaded from CSVs. Excel file not created.")
            return

        # Combine all individual dataframes into one large dataframe
        combined_df = pd.concat(all_dataframes, ignore_index=True)
        
        # Save the combined dataframe to a single sheet in an Excel file
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            combined_df.to_excel(writer, sheet_name='All Staff', index=False)
        
        print("\n--- SUCCESS! ---")
        print(f"Excel file with all staff on one sheet has been created: {excel_path}")
    
    except Exception as e:
        print(f"\n--- !! FAILED to create Excel file !! ---")
        print(f"Error: {e}")

def main():
    """Main function to run the crawler."""
    
    # Credentials no longer needed at the start
    # email = input("Enter your UofT Email (e.g., firstname.lastname@utoronto.ca): ")
    # username = input("Enter your UofT Username (e.g., utorid): ")
    # password = getpass.getpass("Enter your password: ")
    print("Starting the directory crawler...")

    driver = setup_driver(DOWNLOAD_DIR)
    downloaded_files = {} # To store { "Dept Name": "path/to/file.csv" }

    try:
        # --- Step 1: Manual Log in ---
        driver.get(LOGIN_URL)
        print("\n--- !! ACTION REQUIRED !! ---")
        print("Please log in manually in the browser window that just opened.")
        print("Log in with your email, password, and Duo 2-Factor push.")
        print("The script will automatically take over once you land on the directory search page...")
        
        try:
            # Wait for the user to successfully log in by watching for the search page URL
            # Set a long timeout (e.g., 5 minutes = 300 seconds) to give you time
            WebDriverWait(driver, 300).until(
                EC.url_contains(SEARCH_PAGE_URL)
            )
            print("\n--- LOGIN SUCCESSFUL! ---")
            print("Taking over and starting automated downloads...")

        except Exception as e:
            print("\n--- !! LOGIN TIMED OUT !! ---")
            print("Script did not detect the search page URL after 5 minutes. Exiting.")
            driver.quit()
            exit()


        # --- Step 2: Loop and download ---
        print(f"\nStarting to download data for {len(DEPARTMENT_LIST)} departments...")
        for dept in DEPARTMENT_LIST:
            file_path = download_department_csv(driver, dept, DOWNLOAD_DIR)
            if file_path:
                downloaded_files[dept] = file_path

        print("\nAll downloads finished.")
    
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    
    finally:
        # --- Step 3: Clean up ---
        print("Closing the browser.")
        #driver.quit()

    # --- Step 4: Combine files ---
    combine_csvs_to_excel(downloaded_files, FINAL_EXCEL_FILE)

if __name__ == "__main__":
    main()



