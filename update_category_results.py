import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.firefox import GeckoDriverManager
from docx import Document
import os

DOC_PATH = 'Company Rankings.docx'

# Load existing document
if os.path.exists(DOC_PATH):
    print(f"Loading existing document: {DOC_PATH}")
    doc = Document(DOC_PATH)
else:
    print(f"Error: Document {DOC_PATH} not found.")
    exit()

def auto_select_dropdown(wait, input_field_id, field_name, user_inputs):
    retries = 3
    while retries > 0:
        user_input = input(f"Enter desired {field_name} (leave blank for none): ").strip()
        user_inputs[field_name] = user_input if user_input else ""
        try:
            if not user_input:
                print(f"No input provided for {field_name}; skipping selection.\n")
                return True

            try:
                selector = f".css-n9qnu9 input#{input_field_id}"
                input_field = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
            except TimeoutException:
                print(f"Container-based selector not found for {field_name}. Trying fallback by ID.")
                input_field = wait.until(EC.element_to_be_clickable((By.ID, input_field_id)))

            input_field.click()
            time.sleep(1)
            input_field.clear()
            input_field.send_keys(user_input)
            time.sleep(2)

            try:
                options = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[role='option']")))
            except TimeoutException:
                print(f"Timeout waiting for dropdown options for {field_name}; using input as fallback.")
                user_inputs[field_name] = user_input.upper()
                return True

            found_option = None
            for option in options:
                if user_input.lower() in option.text.lower():
                    found_option = option
                    break

            if found_option:
                print(f"Clicking option: '{found_option.text}' for {field_name}")
                found_option.click()
                time.sleep(2)
                selected_value = user_input.upper()
                print(f"Selected '{selected_value}' for {field_name} (based on user input).\n")
                user_inputs[field_name] = selected_value
                return True
            else:
                print(f"No matching option found for {field_name} with input '{user_input}'; using input anyway.")
                user_inputs[field_name] = user_input.upper()
                return True

        except Exception as e:
            print(f"Error selecting {field_name}: {e}")
            retries -= 1
            if retries > 0:
                print(f"Retrying ({retries} attempts left)...")
                time.sleep(2)
            else:
                print(f"Max retries reached; using input '{user_input.upper()}' as fallback.")
                user_inputs[field_name] = user_input.upper()
                return True

        retry = input(f"Try again for {field_name}? (y/n): ").strip().lower()
        if retry != 'y':
            print(f"Aborting selection for {field_name}.\n")
            return False

def update_category_with_results(doc, level_2, level_3, num_results):
    found_level_2 = False
    num_text = str(num_results if num_results is not None else 0)

    for para in doc.paragraphs:
        if para.style.name == 'Heading 2':
            if para.text.strip() == level_2:
                found_level_2 = True
            else:
                found_level_2 = False  # Reset when encountering a new Heading 2
        elif found_level_2 and para.style.name == 'Heading 3':
            current_text = para.text.strip()
            if current_text:  # Skip empty paragraphs
                base_text = current_text.rsplit(' ', 1)[0] if current_text.split()[-1].isdigit() else current_text
                if base_text == level_3:
                    para.text = f"{level_3} {num_text}"
                    print(f"Updated '{level_3}' to '{para.text}' under '{level_2}'")
                    return
            else:
                print(f"Skipping empty Heading 3 paragraph under '{level_2}'")

    print(f"Error: Category '{level_3}' not found under '{level_2}'. No update performed.")

# Setup Selenium
options = Options()
options.headless = True
driver_path = GeckoDriverManager().install()
if not os.access(driver_path, os.X_OK):
    print(f"Fixing permissions for {driver_path}")
    os.chmod(driver_path, 0o755)
service = Service(executable_path=driver_path, port=4444, log_path="geckodriver.log")
driver = webdriver.Firefox(service=service, options=options)
URL = "https://www.fundssniper.com/en/search-funds"

# Main loop
while True:
    user_inputs = {}
    try:
        driver.get(URL)
        wait = WebDriverWait(driver, 15)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".grow")))

        fields = {
            "fund-country": "react-select-5-input",
            "sector-1": "react-select-6-input",
            "sector-2": "react-select-7-input",
            "continent to invest in": "react-select-8-input",
            "country to invest in": "react-select-9-input",
            "size of investment": "react-select-10-input",
            "currency": "react-select-11-input",
            "min/maj": "react-select-12-input",
            "fund-type": "react-select-13-input",
            "fund-specialty": "react-select-14-input"
        }

        for field_name, field_id in fields.items():
            success = auto_select_dropdown(wait, field_id, field_name, user_inputs)
            if not success:
                print(f"Skipping {field_name} selection.\n")

        time.sleep(2)

        submit_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-variant='primary']")))
        submit_button.click()
        print("Search button clicked.")
        time.sleep(4)

        try:
            results_element = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "p.font-roboto.font-semibold.text-subtitle1.m-0.undefined")
            ))
            results_text = results_element.text.strip()  # e.g., "31 funds matching"
            num_results = int(''.join(filter(str.isdigit, results_text)))  # Extract 31
            print(f"Number of Results from website for this query: {num_results}")
        except TimeoutException:
            num_results = 0
            print("Could not find number of results on the page.")

        # Determine the category
        sector1 = user_inputs.get("sector-1", "Unknown").upper()
        sector2 = user_inputs.get("sector-2", "Unknown").upper()
        fund_type = user_inputs.get("fund-type", "Unknown").upper()
        fund_specialty = user_inputs.get("fund-specialty", "Unknown").upper()
        country = user_inputs.get("fund-country", "")
        us_category = f"US-FUNDS-{sector1}-{sector2}-{fund_type}-{fund_specialty}"
        all_category = f"{sector1}-{sector2}-{fund_type}-{fund_specialty}"

        if country == "":
            update_category_with_results(doc, "ALL FUNDS", all_category, num_results)
        elif country.upper() == "UNITED STATES":
            update_category_with_results(doc, "US FUNDS", us_category, num_results)
        else:
            print(f"No update performed: fund-country '{country}' is neither blank nor 'UNITED STATES'.")

        doc.save(DOC_PATH)
        print(f"Updated document saved: {DOC_PATH}")

        continue_prompt = input("\nWould you like to enter new values and update again? (y/n): ").strip().lower()
        if continue_prompt != 'y':
            print("Exiting program.")
            break

    except Exception as e:
        print(f"An error occurred: {e}")
        continue_prompt = input("An error occurred. Would you like to try again? (y/n): ").strip().lower()
        if continue_prompt != 'y':
            print("Exiting program.")
            break

driver.quit()
doc.save(DOC_PATH)
print(f"Program terminated. Final document saved: {DOC_PATH}")