import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Hide the tkinter root window
Tk().withdraw()

# Ask user to select an Excel file
file_path = askopenfilename(title="Select Excel file with names")

# Load the Excel file
df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()  # Remove any leading/trailing spaces

# Launch Chrome using Selenium
driver = webdriver.Chrome()  # Make sure ChromeDriver is installed and in PATH

# Go through each row in Excel
for index, row in df.iterrows():
    first_name = row["First Name"]
    last_name = row["Last Name"]

    # Open the web form (test site)
    driver.get("https://www.selenium.dev/selenium/web/web-form.html")
    time.sleep(1)

    # Find input fields on the page
    input_elements = driver.find_elements(By.TAG_NAME, "input")

    if len(input_elements) >= 2:
        input_elements[0].clear()
        input_elements[0].send_keys(first_name)

        input_elements[1].clear()
        input_elements[1].send_keys(last_name)

        print(f"Filled: {first_name} {last_name}")
    else:
        print("Could not find the required input fields")

    time.sleep(2)  # Wait 2 seconds to see the result

driver.quit()
