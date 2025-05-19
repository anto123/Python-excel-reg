import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    ElementClickInterceptedException,
)

# Load data from Excel
def load_data_from_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    rows = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:
            rows.append((row[0].strip(), str(row[1]).strip()))
    return rows

# Wait helper
def wait_for_element(driver, by, locator, timeout=10):
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, locator)))
    except TimeoutException:
        return None

# Fill normal fields
def fill_field(element, value):
    current_val = element.get_attribute("value").strip()
    if current_val:
        print(f"Field already filled: {current_val}")
        return
    element.clear()
    element.send_keys(value)
    print(f"Filled field with value: {value}")

# Dropdown select
def select_dropdown(element, value):
    select = Select(element)
    select.select_by_visible_text(value)
    print(f"Selected dropdown value: {value}")

# Generic click
def click_element(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, element.get_attribute("id"))))
    time.sleep(0.5)
    element.click()
    print("Clicked the element.")

# Click checkbox using ID
def click_checkbox(driver, checkbox_id):
    checkbox = wait_for_element(driver, By.ID, checkbox_id, timeout=5)
    if checkbox and not checkbox.is_selected():
        checkbox.click()
        print("Clicked checkbox.")
    else:
        print("Checkbox already selected or not found.")

# Accept popup
def click_accept_popup(driver):
    try:
        accept_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//button[text()='Accept']"))
        )
        accept_button.click()
        print("Clicked Accept button.")
        time.sleep(2)
    except TimeoutException:
        print("No Accept popup.")

# Extract IDs on enrollment page
def extract_ids_from_form(driver):
    try:
        form = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//form[@method='POST']"))
        )
        elements = form.find_elements(By.XPATH, ".//*[@id]")
        print("\n=== XPath Locators in Enrollment Page ===")
        id_mapping = {}
        for elem in elements:
            tag = elem.tag_name
            elem_id = elem.get_attribute("id")
            if elem_id:
                print(f"//{tag}[@id='{elem_id}']")
                id_mapping[elem_id] = elem_id

        try:
            submit_elem = driver.find_element(
                By.XPATH,
                "//div[contains(@class, 'card-footer-action')]//input[@type='submit']"
            )
            submit_value = submit_elem.get_attribute("value").strip()
            print(f"//input[@type='submit'] in .card-footer-action => Value: '{submit_value}'")
            id_mapping["submit_button_xpath"] = "//div[contains(@class, 'card-footer-action')]//input[@type='submit']"
        except NoSuchElementException:
            print("Submit button not found.")

        return id_mapping

    except Exception as e:
        print(f"Failed to extract enrollment form elements: {e}")
        return {}

# Main form processor
def process_form_data(driver, rows, field_id_mapping):
    for field_name, field_value in rows:
        print(f"Processing: {field_name} -> {field_value}")
        field_id = field_id_mapping.get(field_name)
        if not field_id:
            print(f"No mapping found for '{field_name}'. Skipping.")
            continue

        if field_name == "Enrollment Package":
            try:
                pkg_elem = wait_for_element(driver, By.ID, field_value)
                if pkg_elem:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", pkg_elem)
                    time.sleep(0.5)
                    pkg_elem.click()
                    print(f"Clicked enrollment package: {field_value}")
                else:
                    print(f"Enrollment package with ID '{field_value}' not found.")
            except Exception as e:
                print(f"Error clicking enrollment package: {e}")
            continue

        if field_name == "Proceed":
            try:
                btn = wait_for_element(driver, By.XPATH, field_value)
                if btn:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                    time.sleep(0.5)
                    btn.click()
                    print("Clicked 'Proceed To Checkout' button.")
                else:
                    print("Proceed button not found.")
            except Exception as e:
                print(f"Error clicking proceed button: {e}")
            continue

        elem = wait_for_element(driver, By.ID, field_id, timeout=5)
        if elem is None:
            print(f"Element with ID '{field_id}' not found.")
            continue

        tag = elem.tag_name.lower()
        input_type = elem.get_attribute("type") or ""

        if field_value.lower() == "click":
            if tag == "input" and input_type == "checkbox":
                click_checkbox(driver, field_id)
            else:
                click_element(driver, elem)
        elif tag == "select":
            select_dropdown(elem, field_value)
        elif input_type == "date":
            driver.execute_script("arguments[0].value = arguments[1];", elem, field_value)
            print(f"Date filled: {field_value}")
        else:
            fill_field(elem, field_value)

# ==== MAIN SCRIPT ====
driver = webdriver.Chrome()
driver.get("https://preprod-binary.epixel.link/en/register/")
driver.maximize_window()

click_accept_popup(driver)

# Excel path
excel_file = '/home/antonyraj.m/Desktop/python/datadetails.xlsx'
rows = load_data_from_excel(excel_file)

# Step 1: Signup mappings
signup_mapping = {
    "Sponsor": "id_sponsor",
    "First Name": "id_first_name",
    "Last Name": "id_last_name",
    "Username": "id_username",
    "Email address": "id_email",
    "Date of Birth": "id_dasstetime",
    "Gender": "id_gender",
    "Subdomain": "id_subdomain",
    "Phone Number": "id_phone_number",
    "Password": "id_password1",
    "Password Confirm": "id_password2",
    "Terms and Conditions": "agree-terms-conditions",
    "addnew-member": "addnew-member"
}

# Step 2: Fill signup form
process_form_data(driver, rows, signup_mapping)

# Step 3: Wait for enrollment page
print("Waiting for Enrollment page to load...")
time.sleep(6)

# Step 4: Extract enrollment field IDs
enrollment_mapping = extract_ids_from_form(driver)

# Step 5: Add Excel-friendly keys to mapping
enrollment_mapping["Enrollment Package"] = "placeholder"  # Value taken from Excel instead
enrollment_mapping["Proceed"] = enrollment_mapping.get("submit_button_xpath")

# Step 6: Process enrollment form
process_form_data(driver, rows, enrollment_mapping)

print("Waiting before closing browser...")
time.sleep(10)
driver.quit()
print("Automation complete.")
