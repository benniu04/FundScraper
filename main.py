import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from webdriver_manager.firefox import GeckoDriverManager
from docx import Document
import os

DOC_PATH = 'potential_investors.docx'

# Load existing document or create a new one
if os.path.exists(DOC_PATH):
    print(f"Loading existing document: {DOC_PATH}")
    doc = Document(DOC_PATH)
else:
    print(f"Creating new document: {DOC_PATH}")
    doc = Document()
    doc.add_heading('POTENTIAL INVESTORS', level=1)

existing_headings = {}
current_level_2 = None
current_level_3 = None
for paragraph in doc.paragraphs:
    if paragraph.style.name.startswith('Heading'):
        level = int(paragraph.style.name.split()[1])
        if level == 2:
            current_level_2 = paragraph.text
            existing_headings[current_level_2] = {}
        elif level == 3 and current_level_2:
            current_level_3 = paragraph.text
            existing_headings[current_level_2][current_level_3] = []


def auto_select_dropdown(wait, input_field_id, field_name, user_inputs):
    retries = 3
    while retries > 0:
        user_input = input(f"Enter desired {field_name}: ").strip()

        if not user_input:
            user_inputs[field_name] = "Unknown"
            return True
        try:
            # Locate the input field
            try:
                selector = f".css-n9qnu9 input#{input_field_id}"
                input_field = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
            except TimeoutException:
                print(f"Container-based selector not found for {field_name}. Trying fallback by ID.")
                input_field = wait.until(EC.element_to_be_clickable((By.ID, input_field_id)))

            input_field.click()
            time.sleep(1)

            # Enter user input
            input_field.clear()
            input_field.send_keys(user_input)
            time.sleep(2)

            try:
                options = wait.until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[role='option']"))
                )
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
                time.sleep(2)  # Wait for page to stabilize after click
                selected_value = user_input.upper()
                print(f"Selected '{selected_value}' for {field_name} (based on user input).\n")
                user_inputs[field_name] = selected_value
                return True
            else:
                print(f"No matching option found for {field_name} with input '{user_input}'; using input anyway.")
                user_inputs[field_name] = user_input.upper()
                return True

        except StaleElementReferenceException:
            print(f"Stale element encountered for {field_name}; retrying...")
            retries -= 1
            time.sleep(2)
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

def insert_under_heading(doc, level_2, level_3, company_name, collected_data):
    found_level_2 = False
    found_level_3 = False
    insert_position = len(doc.paragraphs)  # Default to end of document

    for i, para in enumerate(doc.paragraphs):
        if para.style.name == 'Heading 2' and para.text == level_2:
            found_level_2 = True
        elif found_level_2 and para.style.name == 'Heading 3' and para.text == level_3:
            found_level_3 = True
            insert_position = i + 1  # Position after the level 3 heading
        elif found_level_2 and para.style.name == 'Heading 2':  # Another level 2 heading encountered
            break  # Stop if we hit the next level 2 heading

    # Insert at the determined position
    if not found_level_2:
        doc.add_heading(level_2, level=2)
        doc.add_heading(level_3, level=3)
        doc.add_heading(company_name, level=4)
        for item in collected_data:
            doc.add_paragraph(item, style='List Bullet')
    elif not found_level_3:
        # Find the last paragraph under the level 2 heading
        for i, para in enumerate(doc.paragraphs):
            if para.style.name == 'Heading 2' and para.text == level_2:
                insert_position = i + 1
            elif para.style.name == 'Heading 2' and i > insert_position:
                break
        doc.paragraphs[insert_position - 1]._element.addnext(doc.add_heading(level_3, level=3)._element)
        doc.add_heading(company_name, level=4)
        for item in collected_data:
            doc.add_paragraph(item, style='List Bullet')
    else:
        # Insert under existing level 3 heading
        doc.paragraphs[insert_position - 1]._element.addnext(doc.add_heading(company_name, level=4)._element)
        for item in collected_data:
            doc.add_paragraph(item, style='List Bullet')

# Setup Selenium (outside the loop)
options = Options()
options.headless = True
options.binary_location = "/Applications/Firefox.app/Contents/MacOS/firefox"

print("Installing latest GeckoDriver...")
driver_path = GeckoDriverManager().install()
print(f"GeckoDriver path: {driver_path}")
print(f"GeckoDriver version: {os.popen(f'{driver_path} --version').read()}")

if not os.access(driver_path, os.X_OK):
    print(f"Fixing permissions for {driver_path}")
    os.chmod(driver_path, 0o755)

service = Service(executable_path=driver_path, port=4444, log_path="geckodriver.log")
try:
    driver = webdriver.Firefox(service=service, options=options)
except Exception as e:
    print(f"Failed to start Firefox driver: {e}")
    raise

URL = "https://www.fundssniper.com/en/search-funds"

# Main loop
while True:
    user_inputs = {}  # Reset user inputs for each iteration
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

        try:
            submit_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-variant='primary']")))
            submit_button.click()
            print("Search button clicked.")
        except TimeoutException:
            print("Search button not found or not clickable.")

        time.sleep(4)

        view_buttons = driver.find_elements(By.XPATH, "//button[normalize-space(text())='View']")
        visible_buttons = [btn for btn in view_buttons if btn.is_displayed()]
        print(f"Found {len(visible_buttons)} visible view buttons.")

        original_window = driver.current_window_handle

        while True:
            view_buttons = driver.find_elements(
                By.XPATH,
                "//button[normalize-space(text())='View' and not(ancestor::div[contains(@class, 'absolute') and "
                "contains(@class, 'z-20') and contains(@class, 'flex') and contains(@class, 'justify-center')])]"
            )
            visible_buttons = [btn for btn in view_buttons if btn.is_displayed()]

            if not visible_buttons:
                print("No more non-blurred view buttons found.")
                break

            view_button = visible_buttons[0]
            print("Clicking a non-blurred view button...")
            driver.execute_script("arguments[0].scrollIntoView(true);", view_button)
            time.sleep(1)

            old_windows = driver.window_handles.copy()
            driver.execute_script("arguments[0].click();", view_button)

            try:
                wait_long = WebDriverWait(driver, 10)
                wait_long.until(lambda d: len(d.window_handles) > len(old_windows))
                new_window = next(handle for handle in driver.window_handles if handle not in old_windows)
                driver.switch_to.window(new_window)
                print("New window detected.")
            except TimeoutException:
                print("No new window detected; detail page loaded in the same tab.")

            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1")))
            except TimeoutException:
                print("Detail page did not load in time.")

            detail_html = driver.page_source
            detail_soup = BeautifulSoup(detail_html, 'html.parser')

            specific_h4_company_name = detail_soup.find('h4', class_="font-kodchasan font-light text-h8 undefined")
            specific_h4_about = detail_soup.find('h4', class_="font-kodchasan font-medium leading-[120%] text-h13 "
                                                              "xl:text-h4 mb-[1rem]")
            specific_p_about = detail_soup.find('p', class_="font-roboto font-regular text-subtitle3 m-0 undefined")

            main_container = detail_soup.find("div", class_="2xl:max-w-[1264px] 2xl:mx-auto xl:mx-[50px] lg:mx-[44px] "
                                                            "md:mx-[40px] sm:mx-[16px] mx-[8px]")
            if main_container:
                main_tags = main_container.select_one("p.font-roboto.font-semibold.text-subtitle1")
                main_info = main_container.select("p.font-roboto.font-semibold.text-subtitle2.m-0.undefined")
                main_title = main_tags.get_text(strip=True) if main_tags else ""
                main_texts = ", ".join(tag.get_text(strip=True) for tag in main_info)

            containers = detail_soup.find_all("div", class_="flex flex-col items-center justify-between lg:flex-row")
            grouped_info = []
            for container in containers:
                header_tag = container.select_one("p.font-roboto.font-semibold.text-subtitle1")
                info_tags = container.select("p.font-roboto.font-semibold.text-subtitle2")
                if header_tag:
                    header_text = header_tag.get_text(strip=True)
                    info_texts = ", ".join(tag.get_text(strip=True) for tag in info_tags)
                    grouped_info.append(f"{header_text}: {info_texts}")

            has_team_section = detail_soup.find(
                lambda tag: tag.name in ['h2', 'h3', 'div'] and 'team' in tag.text.lower())
            team_texts = []
            team_linkedin_info = ""
            if has_team_section:
                print("Team section detected; scraping team info.")
                specific_p_team = detail_soup.find_all('p',
                                                       class_="font-roboto font-regular text-subtitle3 m-0 mb-[1rem]") or []
                team_texts = [p.get_text(strip=True) for p in specific_p_team if p.get_text(strip=True)]
                try:
                    team_linkedin_element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//a[contains(., 'LinkedIn') and @href]"))
                    )
                    team_linkedin_info = team_linkedin_element.get_attribute("href")
                except TimeoutException:
                    print("Team LinkedIn link not found within 10 seconds.")
                all_team_info = team_texts + [team_linkedin_info] if team_linkedin_info else team_texts
                full_team = ", ".join(all_team_info) if all_team_info and any(all_team_info) else ""
            else:
                print("No team section found; skipping team info scraping.")
                full_team = ""

            address_parts = detail_soup.find_all("p",
                                                 class_="font-roboto font-regular text-subtitle3 m-0 text-center") or []
            address_parts = [p for p in address_parts if p.get_text(strip=True) and "ADDRESS" not in p.get_text()]
            full_address = ", ".join(p.get_text(strip=True) for p in address_parts)

            try:
                linkedin_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//a[contains(., 'LinkedIn')]"))
                )
                link_href = linkedin_element.get_attribute("href")
            except TimeoutException:
                print("Main LinkedIn link not found; setting as empty.")
                link_href = ""

            try:
                website_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//a[contains(., 'WEBSITE')]"))
                )
                website_href = website_element.get_attribute("href")
            except TimeoutException:
                print("Website link not found; setting as empty.")
                website_href = ""

            email_label, call_label = None, None
            for p in detail_soup.find_all("p"):
                text = p.get_text(strip=True)
                if "EMAIL" in text.upper() and email_label is None:
                    email_label = p
                if "CALL" in text.upper() and call_label is None:
                    call_label = p
                if email_label and call_label:
                    break

            email_address = ""
            if email_label:
                email_a = email_label.find_next_sibling("a")
                if email_a:
                    email_p = email_a.find("p")
                    email_address = email_p.get_text(strip=True) if email_p else ""

            call_address = ""
            if call_label:
                call_element = call_label.find_next_sibling("a") or call_label.find_next_sibling("p")
                call_address = call_element.get_text(strip=True) if call_element else ""

            collected_data = []
            if specific_h4_company_name:
                collected_data.append("Company: " + specific_h4_company_name.get_text(strip=True))
            if specific_h4_about:
                collected_data.append("About: " + specific_h4_about.get_text(separator=" ", strip=True))
            if specific_p_about:
                collected_data.append("Description: " + specific_p_about.get_text(strip=True))
            if main_title:
                collected_data.append("Main Info: " + f"{main_title}: {main_texts}")
            if grouped_info:
                collected_data.extend(grouped_info)
            if full_team.strip():
                collected_data.append("Team: " + full_team)
            if full_address:
                collected_data.append("Address: " + full_address)
            if link_href:
                collected_data.append("LinkedIn URL: " + link_href)
            if website_href:
                collected_data.append("Website URL: " + website_href)
            if email_address:
                collected_data.append("Email: " + email_address)
            if call_address:
                collected_data.append("Call: " + call_address)

            country = user_inputs.get("country to invest in", "Unknown").upper()
            sector = user_inputs.get("sector-1", "Unknown")
            fund_type = user_inputs.get("fund-type", "Unknown").upper()
            fund_specialty = user_inputs.get("fund-specialty", "Unknown").upper()

            fund_types = ["VC", "PE", "LBO", "OPPORTUNIST"]
            fund_specialties = ["INCUBATION", "SEED", "EARLY"]

            fund_type = fund_type if fund_type in fund_types else "Unknown"
            fund_specialty = fund_specialty if fund_specialty in fund_specialties else "Unknown"

            us_category = f"US FUNDS-{sector}-{fund_type}-{fund_specialty}"
            all_category = f"{sector}-{fund_type}-{fund_specialty}"

            company_name = specific_h4_company_name.get_text(
                strip=True) if specific_h4_company_name else "Unknown Company"

            if country == "US":
                insert_under_heading(doc, "US FUNDS", us_category, company_name, collected_data)

                # Insert under ALL FUNDS
            insert_under_heading(doc, "ALL FUNDS", all_category, company_name, collected_data)

            print(f"Added {company_name} to categories: {us_category if country == 'US' else ''} {all_category}")

            if len(driver.window_handles) > len(old_windows):
                driver.close()
                driver.switch_to.window(old_windows[0])
            else:
                driver.back()

            driver.execute_script("arguments[0].remove();", view_button)
            time.sleep(3)

        doc.save(DOC_PATH)
        print(f"Updated document saved: {DOC_PATH}")

        # Ask user if they want to continue
        continue_prompt = input(
            "\nScraping complete. Would you like to enter new values and scrape again? (y/n): ").strip().lower()
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
