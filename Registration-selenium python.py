import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.common.exceptions import (
    NoSuchElementException, TimeoutException, ElementClickInterceptedException
)

def load_data_from_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    return [(str(row[0]).strip(), str(row[1]).strip()) for row in sheet.iter_rows(min_row=2, values_only=True) if row[0] and row[1]]

def wait_for_element(driver, by, locator, timeout=10):
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, locator)))
    except TimeoutException:
        return None

def fill_field(element, value):
    current_val = element.get_attribute("value").strip()
    if current_val:
        print(f"Field already filled: {current_val}")
        return
    element.clear()
    element.send_keys(value)
    print(f"Filled field with value: {value}")

def select_dropdown(element, value):
    try:
        Select(element).select_by_visible_text(value)
        print(f"Selected dropdown value: {value}")
    except Exception as e:
        print(f"Dropdown selection failed: {e}")

def click_element(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, element.get_attribute("id"))))
    time.sleep(0.5)
    element.click()
    print("Clicked the element.")

def click_checkbox(driver, checkbox_id):
    checkbox = wait_for_element(driver, By.ID, checkbox_id, timeout=5)
    if checkbox and not checkbox.is_selected():
        checkbox.click()
        print("Clicked checkbox.")
    else:
        print("Checkbox already selected or not found.")

def click_accept_popup(driver):
    try:
        accept_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[text()='Accept']"))
        )
        accept_button.click()
        print("Clicked Accept button.")
        time.sleep(1.5)
    except TimeoutException:
        print("No Accept popup.")

def print_xpath_locators(driver):
    try:
        form = driver.find_element(By.ID, "register-form")
        elements = form.find_elements(By.XPATH, 
            ".//input[@id] | .//select[@id] | .//button[@id] | .//button[@type] | .//textarea[@id]"
        )
        print("\n=== XPath Locators (Signup Page) ===")
        for elem in elements:
            elem_id = elem.get_attribute("id")
            tag = elem.tag_name
            if elem_id:
                print(f"//{tag}[@id='{elem_id}']")
            elif tag == "button":
                btn_type = elem.get_attribute("type")
                if btn_type:
                    print(f"//button[@type='{btn_type}']")
    except Exception as e:
        print(f"Failed to print locators: {e}")

def extract_ids_from_form(driver):
    try:
        form = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//form[@method='POST']")))
        elements = form.find_elements(By.XPATH, ".//*[@id]")
        id_mapping = {elem.get_attribute("id"): elem.get_attribute("id") for elem in elements if elem.get_attribute("id")}
        try:
            submit_elem = form.find_element(By.XPATH, ".//input[@type='submit' and not(@disabled)]")
            id_mapping["submit_button_xpath"] = "//form[@method='POST']//input[@type='submit' and not(@disabled)]"
        except NoSuchElementException:
            print("Submit button not found.")
        return id_mapping
    except Exception as e:
        print(f"Failed to extract form elements: {e}")
        return {}

def process_form_data(driver, rows, field_id_mapping):
    for field_name, field_value in rows:
        if field_name == "URL":
            continue  # URL is already handled
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
            except Exception as e:
                print(f"Error clicking enrollment package: {e}")
            continue

        if field_name == "Proceed":
            try:
                submit_xpath = field_id_mapping.get("submit_button_xpath")
                if not submit_xpath:
                    print("No XPath found for Proceed button.")
                    continue
                btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, submit_xpath)))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                time.sleep(0.5)
                btn.click()
                print("Clicked 'Proceed To Checkout' button.")
            except Exception as e:
                print(f"Error clicking Proceed: {e}")
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

# ===== MAIN SCRIPT =====

excel_file = '/home/antonyraj.m/Desktop/python/datadetails.xlsx'
rows = load_data_from_excel(excel_file)

# Get URL from Excel
url = next((val for key, val in rows if key == "URL"), None)
if not url:
    raise Exception("URL not found in Excel. Please provide a URL in the first row.")

driver = webdriver.Chrome()
driver.get(url)
driver.maximize_window()
click_accept_popup(driver)

print_xpath_locators(driver)

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
}

# Fill signup form
process_form_data(driver, rows, signup_mapping)

# Step 1: Click the Sign Up button
try:
    sign_up_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and @label='Sign Up']"))
    )
    sign_up_btn.click()
    print("Clicked 'Sign Up' button.")
except TimeoutException:
    print("Sign Up button not found.")
    driver.quit()

# Step 2: Wait for the popup to possibly appear
time.sleep(3)  # Wait for popup animation/load if any

# Step 3: Check if token input popup appears
try:
    token_input = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.ID, "user_token"))
    )
    if token_input.is_displayed():
        print("Token popup appeared. Entering token...")
        token_input.clear()
        token_input.send_keys("422602")

        # Click the token confirm button
        try:
            confirm_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//button[@type='submit' and contains(text(), 'Sign Up') and @onclick='update_token(this);return false;']"
                ))
            )
            confirm_btn.click()
            print("Submitted token and continuing.")
        except TimeoutException:
            print("Confirm button after token entry not found.")
    else:
        print("Token input field is not visible. Continuing with next step.")
except TimeoutException:
    print("No token popup appeared. Continuing with next step.")


# Wait for enrollment page to load
print("Waiting for Enrollment page...")
time.sleep(6)

# Extract and fill enrollment form
enrollment_mapping = extract_ids_from_form(driver)
enrollment_mapping["Enrollment Package"] = "placeholder"
enrollment_mapping["Proceed"] = "click"
process_form_data(driver, rows, enrollment_mapping)

# Fill billing address manually
print("\nFilling Billing Address...")
driver.find_element(By.NAME, "billing-customer_address_name_line").send_keys("ears")
driver.find_element(By.NAME, "billing-customer_address_premise").send_keys("ears")
driver.find_element(By.NAME, "billing-customer_address_locality").send_keys("ears")
driver.find_element(By.NAME, "billing-customer_address_postal_code").send_keys("895456")

# Country selection
country = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[@id='select2-id_billing-customer_address_country-container']")))
driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", country)
country.click()
india = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//li[text()='India']")))
india.click()
print("Selected country: India")

# State selection
billing_state = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[@aria-labelledby='select2-id_billing-customer_address_state-container']")))
driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", billing_state)
billing_state.click()
state_drop = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//li[text()='Andhra Pradesh']")))
state_drop.click()
print("Selected state: Andhra Pradesh")

# Click checkout
checkout = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@name='checkout']")))
checkout.click()
print("Clicked checkout.")

# Handle payment form
try:
    Test_payment = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//form[contains(@class,'payment-form-default')]//input[@value='Proceed to Make Payment']"))
    )
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", Test_payment)
    Test_payment.click()
    print("Clicked 'Proceed to Make Payment'.")
except TimeoutException:
    print("Error: 'Proceed to Make Payment' button not found.")
    driver.save_screenshot("payment_button_not_found.png")
    raise

# Confirm dropdown and finish
select_element = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//select[@id='id_status']"))
)
Select(select_element).select_by_visible_text("Confirmed")

finish_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Finish']")))
finish_button.click()

print("Waiting before closing browser...")
time.sleep(10)
driver.quit()
print("Automation complete. Registration finished.")
