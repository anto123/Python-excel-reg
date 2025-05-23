import time
import openpyxl
import random
import string
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoSuchElementException, JavascriptException

# === RANDOM GENERATORS ===
EMAIL_PREFIXES = ['epsminuanish', 'epsshyn', 'epskrish', 'epssanjana', 'epssreya', 'epsgoku', 'epsannjoe', 'epsmithuna', 'epsmithuna', 'epsnrk', 'epsakhila', 'epsananden', 'epsanand.m', 
'epssreetha', 'epssdhshone', 'epssdhshtwo','epssarathcr','epssreelakshmi', 
'epsvivekks', 'epssuhail', 'epssuhu', 'epsswathy', 'epsdileepk','epsprabi', 
'epsnadhiya', 'epsishamk', 'epsrx100', 'epsabdulrahim', 'epsrahul1', 'epsarjunk', 
'epsrayees', 'epsbineesha', 'epsshinu','epsriyasac', 'epssaif', 'epsrenu', 
'epssujith','epsvvp','epsgokul','epsdivyavarghese','epsjassim','epslincy', 
'epsunni','epsharooq','epsmshruthi','epsjithinkp','epssudhi','epsfasil', 
'epsaiswaryar','epsdanish','epsasaz','epsanjjuuzz','epsnaseeb','epsijaz']

def generate_random_string(length=6):
    return ''.join(random.choices(string.ascii_lowercase, k=length))

def generate_random_username():
    return generate_random_string(5) + str(random.randint(100, 999))

def generate_random_email():
    prefix = random.choice(EMAIL_PREFIXES)
    return f"{prefix}@gmail.com"

def generate_random_phone():
    # 10-digit phone starting with 6-9
    first_digit = random.choice("6789")
    other_digits = ''.join(random.choices(string.digits, k=9))
    return first_digit + other_digits

# === EXCEL DATA LOADING ===
def load_data_from_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    data = []
    for row in sheet.iter_rows(min_row=2, max_col=2, values_only=True):
        key = row[0]
        val = row[1]
        if key is not None and val is not None:
            data.append((str(key).strip(), str(val).strip()))
    return data

# === WAIT UTILITY ===
def wait_for_element(driver, by, locator, timeout=10):
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, locator)))
    except TimeoutException:
        return None

# === FORM FILLING HELPERS ===
def fill_field(element, value):
    try:
        element.clear()
        element.send_keys(value)
        print(f"Filled field with value: {value}")
    except Exception as e:
        print(f"Error filling field: {e}")

def select_dropdown(element, value):
    try:
        Select(element).select_by_visible_text(value)
        print(f"Selected dropdown value: {value}")
    except Exception as e:
        print(f"Dropdown selection failed: {e}")

def click_element(driver, element):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.5)
        element.click()
        print("Clicked element.")
    except Exception as e:
        print(f"Click failed: {e}")

def click_checkbox(driver, checkbox_id):
    checkbox = wait_for_element(driver, By.ID, checkbox_id)
    if checkbox and not checkbox.is_selected():
        checkbox.click()
        print("Checkbox clicked.")

def click_accept_popup(driver):
    try:
        accept_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[text()='Accept']"))
        )
        accept_btn.click()
        print("Accepted cookies popup.")
    except TimeoutException:
        print("No cookie popup detected.")

def print_xpath_locators(driver):
    try:
        form = driver.find_element(By.ID, "register-form")
        elements = form.find_elements(By.XPATH, 
            ".//input[@id] | .//select[@id] | .//button[@id] | .//textarea[@id]"
        )
        print("\n--- XPath Locators (Signup Form) ---")
        for elem in elements:
            elem_id = elem.get_attribute("id")
            tag = elem.tag_name
            if elem_id:
                print(f"//{tag}[@id='{elem_id}']")
    except Exception as e:
        print(f"Locator print failed: {e}")

def click_change_button_if_present(driver):
    try:
        change_btn = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, "//button[text()='Change']"))
        )
        change_btn.click()
        print("Clicked 'Change' button.")
        time.sleep(1)
        return True
    except TimeoutException:
        return False

def handle_sponsor_field(driver, sponsor_id, sponsor_value):
    # Check for comma-separated sponsor names
    sponsor_raw = str(sponsor_value).strip()
    if ',' in sponsor_raw:
        sponsor_list = [s.strip() for s in sponsor_raw.split(',') if s.strip()]
        sponsor_value = random.choice(sponsor_list)
        print(f"Randomly selected sponsor: {sponsor_value}")

    elem = wait_for_element(driver, By.ID, sponsor_id)
    if not elem:
        print("Sponsor field not found.")
        return
    if click_change_button_if_present(driver):
        try:
            elem = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, sponsor_id)))
            elem.clear()
            print("Cleared sponsor field for change.")
        except Exception:
            pass
    fill_field(elem, sponsor_value)

def extract_ids_from_form(driver):
    try:
        form = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//form[@method='POST']")))
        elements = form.find_elements(By.XPATH, ".//*[@id]")
        id_mapping = {elem.get_attribute("id"): elem.get_attribute("id") for elem in elements if elem.get_attribute("id")}
        # Try to find submit button xpath
        try:
            submit_elem = form.find_element(By.XPATH, ".//input[@type='submit' and not(@disabled)]")
            id_mapping["submit_button_xpath"] = "//form[@method='POST']//input[@type='submit' and not(@disabled)]"
            submit_outer_html=submit_elem.get_attribute("outerHTML")
            print("submit prooceed button", submit_outer_html)
        except Exception:
            pass
        return id_mapping
    except Exception as e:
        print(f"Failed to extract form IDs: {e}")
        return {}

def process_form_data(driver, rows, field_id_mapping, generated_data={}):
    for field_name, field_value in rows:
        if field_name == "URL":
            continue  # skip URL field

        # Override with generated data if available
        if field_name in generated_data:
            field_value = generated_data[field_name]
            print(f"Using generated value for {field_name}: {field_value}")

        field_id = field_id_mapping.get(field_name)
        if not field_id:
            print(f"Skipping field '{field_name}' as no matching element ID found.")
            continue

        # Sponsor special handling
        if field_name == "Sponsor":
            handle_sponsor_field(driver, field_id, field_value)
            continue

        # Enrollment Package special handling
        if field_name == "Enrollment Package":
            pkg_elem = wait_for_element(driver, By.ID, field_value)
            if pkg_elem:
                click_element(driver, pkg_elem)
                print(f"Selected enrollment package: {field_value}")
            else:
                print(f"Enrollment package with ID '{field_value}' not found.")
            continue

        # Proceed special handling - click submit
        if field_name == "Proceed":
            submit_xpath = field_id_mapping.get("submit_button_xpath")
            if submit_xpath:
                try:
                    submit_btn = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, submit_xpath))
                    )
                    click_element(driver, submit_btn)
                    print("Clicked 'Proceed' submit button.")
                except TimeoutException:
                    print("Proceed submit button not clickable/found.")
            else:
                print("Proceed submit button XPath not found in mapping.")
            continue

        # Fill normal fields
        elem = wait_for_element(driver, By.ID, field_id)
        if not elem:
            print(f"Element with ID '{field_id}' not found for field '{field_name}'.")
            continue

        tag = elem.tag_name.lower()
        input_type = elem.get_attribute("type") or ""

        # Handle different field types
        if field_value.lower() == "click":
            # For checkboxes or clickable buttons
            if tag == "input" and input_type == "checkbox":
                click_checkbox(driver, field_id)
            else:
                click_element(driver, elem)
        elif tag == "select":
            select_dropdown(elem, field_value)
        elif input_type == "date":
            # Date input - set value via JS
            driver.execute_script("arguments[0].value = arguments[1];", elem, field_value)
            print(f"Set date field '{field_name}' to '{field_value}'")
        else:
            fill_field(elem, field_value)

# === MAIN SCRIPT START ===
def main():
    excel_file = './newdata.xlsx'
    rows = load_data_from_excel(excel_file)

    # Extract URL from Excel data
    url = next((val for key, val in rows if key == "URL"), None)
    if not url:
        raise Exception("URL is missing in Excel file")

    # Generate random data for specific fields
    generated_data = {
        "First Name": generate_random_string(6),
        "Last Name": generate_random_string(6),
        "Username": generate_random_username(),
        # "Subdomain": generate_random_username(),
        "Email address": generate_random_email(),
        "Phone Number": generate_random_phone(),
    }
    print("Generated Random Data:", generated_data)

    # Setup webdriver
    driver = webdriver.Chrome()
    driver.get(url)
    driver.maximize_window()

    time.sleep(5)
    driver.find_element(By.XPATH, "//input[@id='id_first_name']").send_keys(generated_data['First Name'])
    driver.find_element(By.XPATH, "//input[@id='id_last_name']").send_keys(generated_data['Last Name'])
    driver.find_element(By.XPATH, "//input[@id='id_username']").send_keys(generated_data['Username'])
    driver.find_element(By.XPATH, "//input[@id='id_email']").send_keys(generated_data['Email address'])
    driver.find_element(By.XPATH, "//input[@id='id_phone_number']").send_keys(generated_data['Phone Number'])
    #driver.find_element(By.XPATH, "//input[@id='id_subdomain']").send_keys(generated_data['Subdomain'])
   
    click_accept_popup(driver)
    
    # Print all XPath locators on Signup form
    print_xpath_locators(driver)

    # Map signup form field names to element IDs
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
    process_form_data(driver, rows, signup_mapping, generated_data)

    # Click Sign Up button
    try:
        sign_up_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and (text()='Sign Up' or @label='Sign Up')]"))
        )
        sign_up_btn.click()
        print("Clicked Sign Up button.")
    except TimeoutException:
        print("Sign Up button not found or clickable.")
        driver.quit()
        return

    # Handle token popup (if any)
    time.sleep(3)
    try:
        token_input = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "user_token")))
        if token_input.is_displayed():
            token_input.clear()
            token_input.send_keys("422602")  # Example token, consider externalizing
            confirm_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(text(),'Sign Up') and @onclick='update_token(this);return false;']"))
            )
            confirm_btn.click()
            print("Token submitted.")
    except TimeoutException:
        print("No token popup detected.")

    # Wait for enrollment page to load
    time.sleep(6)

    # Extract enrollment form IDs dynamically
    enrollment_mapping = extract_ids_from_form(driver)
    # Add custom keys for Enrollment Package and Proceed button handling
    enrollment_mapping["Enrollment Package"] = "placeholder"  # Will be replaced by actual package ID from Excel data
    enrollment_mapping["Proceed"] = "click"

    # Process enrollment form
    process_form_data(driver, rows, enrollment_mapping)

    print("Filling Billing Information...")
    try:
        # === Billing Info Filling ===
        driver.find_element(By.XPATH, "//input[@name='billing-customer_address_name_line']").send_keys("ears")
        driver.find_element(By.XPATH, "//input[@name='billing-customer_address_premise']").send_keys("ears")
        driver.find_element(By.XPATH, "//input[@name='billing-customer_address_locality']").send_keys("ears")
        driver.find_element(By.XPATH, "//input[@name='billing-customer_address_postal_code']").send_keys("895456")

        # Country selection
        country = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@id='select2-id_billing-customer_address_country-container']"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", country)
        country.click()

        india = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[text()='India']"))
        )
        india.click()
        print("Selected country: India")

        # State selection
        billing_state = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@aria-labelledby='select2-id_billing-customer_address_state-container']"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", billing_state)
        billing_state.click()

        state_drop = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[text()='Andhra Pradesh']"))
        )
        state_drop.click()
        print("Selected state: Andhra Pradesh")
        
        try:
            address_checkbox = driver.find_element(By.ID, "styled-checkbox-1")
            print("Checkbox ID:", address_checkbox.get_attribute("id"))

            # Scroll into view
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", address_checkbox)

            # Small wait to ensure scroll animation is done
            time.sleep(1)

            # Click via JavaScript
            driver.execute_script("arguments[0].click();", address_checkbox)
            print("Clicked the checkbox via JavaScript.")

        except NoSuchElementException:
            print("Checkbox not found — continuing...")

        except JavascriptException as e:
            print("JavaScript click failed:", str(e))
                      
        # Checkout
        checkout = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@name='checkout']"))
        )
        checkout.click()
        print("Clicked checkout.")
    except TimeoutException:
        print("no problem")
        
        # === Payment Process === Test payment
    handle_test_payment(driver)
def handle_test_payment(driver):
        try:
            print("successssssssss")
            Test_payment = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//form[contains(@class,'payment-form-default')]//input[@value='Proceed to Make Payment']"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", Test_payment)
            time.sleep(1)
            Test_payment.click()
            print("Clicked 'Proceed to Make Payment'.")

        except TimeoutException:
            print("Error: 'Proceed to Make Payment' button not found.")
            driver.save_screenshot("payment_button_not_found.png")
            raise
        #  === Final Step: Confirm and Finish ===
        try:
            select_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//select[@id='id_status']"))
            )
            Select(select_element).select_by_visible_text("Confirmed")
            print("Selected status: Confirmed")

            finish_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[text()='Finish']"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", finish_button)
            finish_button.click()
            print("Clicked Finish.")

        except TimeoutException:
            print("Error: Final confirmation step failed.")
            driver.save_screenshot("confirmation_failed.png")
            raise

        # Wait before closing
        print("Waiting before closing browser...")
        input("script finished")
        time.sleep(5)
        
# def handle_stripe_payment(driver):
#     # Strip payment full work:
    
#     try:
#         driver.find_element(By.XPATH, "//label[@for='stripeintent_subscription']").click()
        
#         stripe = WebDriverWait(driver, 20).until(
#             EC.element_to_be_clickable((By.XPATH, "//form[contains(@class,'payment-form-stripe')]//input[@value='Proceed to Make Payment']"))
#         )
#         driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", stripe)
#         time.sleep(1)
#         stripe.click()
#         print("Clicked 'Proceed to Make Payment'.")

#         # --- NEW: Switch to Stripe iframe for card input ---
#         WebDriverWait(driver, 10).until(
#             EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe[name^='__privateStripeFrame']"))
#         )
#         print("Switched to Stripe iframe.")

#         # Locate and fill the card number input
#         card_number = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.XPATH, "//input[@name='cardnumber']"))
#         )
#         cvc_number = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.XPATH, "//input[@name='cvc']"))
#         )
#         exp_number = WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.XPATH, "//input[@name='exp-date']"))
#         )
#         card_number.click()
#         card_number.send_keys("4242424242424242")
#         time.sleep(1)
#         exp_number.send_keys("0835")  # Expiry MMYY
#         time.sleep(5)
#         cvc_number.send_keys("123")   # CVC
#         print("Card details entered.")

#         # Switch back to main content
#         driver.switch_to.default_content()

#         # Click Pay Now / Submit button
#         pay_button = WebDriverWait(driver, 10).until(
#             EC.element_to_be_clickable((By.ID, "stripe-submit-button"))
#         )
#         driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", pay_button)
#         pay_button.click()
#         print("Clicked Pay Now button.")
#         print("Registration procss completed for stripe")
#         input("Script finished. Press Enter to close browser...")
#     except TimeoutException as e:
#         print("Timeout error:", str(e))
#         driver.save_screenshot("stripe_payment_error.png")
#         driver.switch_to.default_content()

#     except Exception as e:
#         print("Unexpected error:", str(e))
#         driver.save_screenshot("unexpected_stripe_error.png")
#         driver.switch_to.default_content()

if __name__ == "__main__":
    main()