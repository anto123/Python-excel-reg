import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl
def load_data_from_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    rows = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        rows.append(row)
    return rows

def fill_field(field_element, field_value):
    try:
        field_element.clear()
        field_element.send_keys(str(field_value))
        print(f"Filled field with value: {field_value}")
    except Exception as e:
        print(f"Failed to fill field: {e}")

def click_element(element):
    try:
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        time.sleep(1)
        element.click()
        print("Clicked the element.")
    except Exception as e:
        print(f"Failed to click element: {e}")

def process_form_data(driver, rows, field_id_mapping):
    for row in rows:
        field_name = row[0]
        field_value = row[1]

        print(f"Processing: {field_name} -> {field_value}")
        field_id = field_id_mapping.get(field_name)

        if not field_id:
            print(f"No mapping found for '{field_name}'. Skipping.")
            continue

        try:
            # Click logic
            if isinstance(field_value, str) and field_value.lower() == "click":
                try:
                    element = driver.find_element(By.ID, field_id)
                except:
                    element = driver.find_element(By.XPATH, f"//button[@id='{field_id}']")
                click_element(element)

            else:
                # Input fill logic
                element = driver.find_element(By.ID, field_id)
                fill_field(element, field_value)

        except Exception as e:
            print(f"Error processing '{field_name}': {e}")

def print_xpath_locators(driver):
    try:
        form = driver.find_element(By.ID, "register-form")
        elements = form.find_elements(By.XPATH, ".//input[@id] | .//select[@id] | .//button[@id] | .//textarea[@id]")
        
        print("\n=== XPath Locators ===")
        for elem in elements:
            tag = elem.tag_name
            elem_id = elem.get_attribute("id")
            if elem_id:
                print(f"//{tag}[@id='{elem_id}']")
    except Exception as e:
        print(f"Failed to extract XPath locators: {e}")

driver = webdriver.Chrome()
driver.get("https://sys-unilevel.epixelmlmsystem.com/en/register/")
time.sleep(3)
driver.maximize_window()

# Load Excel data
excel_file_path = '/home/antonyraj.m/Desktop/python/datadetails.xlsx'
rows = load_data_from_excel(excel_file_path)

field_id_mapping = {
    "First Name": "id_first_name",
    "Last Name": "id_last_name",
    "Username": "id_username",
    "Email address": "id_email",
    "Sponsor": "id_sponsor",
    "Subdomain": "id_subdomain",
    "Phone Number": "id_phone_number",
    "DASSTE Time": "id_dasstetime",
    "Password": "id_password1",
    "Password Confirm": "id_password2",
    "Terms and Conditions": "agree-terms-conditions",  # checkbox
    "addnew-member": "addnew-member"  # submit button
    
}

# Fill the form and click necessary fields
process_form_data(driver, rows, field_id_mapping)

# Print XPath locators for all elements in the form
print_xpath_locators(driver)

# Wait and close browser
time.sleep(5)
driver.quit()
