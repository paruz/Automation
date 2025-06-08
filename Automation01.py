# install with pip: selenium, pandas, openpyxl
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from tkinter import filedialog, Tk

# Step 1: Load Excel file
def load_excel():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    df = pd.read_excel(file_path)
    return df

# Step 2: Fill website fields using Selenium
def fill_form(name):
    driver = webdriver.Chrome()  # Or use webdriver.Firefox(), with geckodriver
    driver.get("https://example.com/form")

    # Wait and fill
    name_field = driver.find_element(By.ID, "name")  # Adjust to real HTML
    name_field.send_keys(name)

    # You can add more fields here...
    # driver.find_element(...).send_keys(...)

    # driver.find_element(By.ID, "submit").click()  # optional

# Main logic
if __name__ == "__main__":
    df = load_excel()
    for index, row in df.iterrows():
        fill_form(row['Name'])  # Match column name in Excel
