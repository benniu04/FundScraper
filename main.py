import time
from bs4 import BeautifulSoup
from docx import Document
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.firefox import GeckoDriverManager

import os

DOC_PATH = 'test.docx'

# Load existing document or create a new one
if os.path.exists(DOC_PATH):
    print(f"Loading existing document: {DOC_PATH}")
    doc = Document(DOC_PATH)
else:
    print(f"Creating new document: {DOC_PATH}")
    doc = Document()
    doc.add_heading('POTENTIAL INVESTORS', level=1)

all_companies = set()
existing_headings = {}
current_level_2 = None
current_level_3 = None

for paragraph in doc.paragraphs:
    if paragraph.style.name.startswith('Heading'):
        level = int(paragraph.style.name.split()[1])
        if level == 2:
            current_level_2 = paragraph.text.strip()
            if current_level_2 not in existing_headings:
                existing_headings[current_level_2] = {}
        elif level == 3 and current_level_2:
            current_level_3 = paragraph.text.split(' (')[0].strip()
            if current_level_3 not in existing_headings[current_level_2]:
                existing_headings[current_level_2][current_level_3] = set()
        elif level == 4 and current_level_2 and current_level_3:
            company_name = paragraph.text.strip()
            existing_headings[current_level_2][current_level_3].add(company_name)
            all_companies.add(company_name)


def auto_select_dropdown(wait, input_field_id, field_name, user_inputs):
    retries = 3
    while retries > 0:
        print(f"\n=== Starting selection for {field_name} (ID: {input_field_id}, Retry: {4 - retries}/3) ===")
        user_input = input(f"Enter desired {field_name}: ").strip()
        user_inputs[field_name] = user_input if user_input else "Unknown"
        print(f"User entered: '{user_input}'")

        if not user_input:
            print(f"No input provided for {field_name}; using 'Unknown'.\n")
            return True

        try:
            selector = f"input#{input_field_id}"
            print(f"Locating {field_name} with selector: {selector}")
            input_field = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, selector)),
                message=f"Failed to locate clickable input field for {field_name} with ID {input_field_id}"
            )
            print(f"Found input field for {field_name}. Current value: '{input_field.get_attribute('value')}'")

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_field)
            driver.execute_script("arguments[0].focus();", input_field)
            time.sleep(1)
            input_field.click()
            print(f"Clicked {field_name}. Value after click: '{input_field.get_attribute('value')}'")
            time.sleep(1)

            input_field.clear()
            print(f"Cleared {field_name}. Value after clear: '{input_field.get_attribute('value')}'")
            time.sleep(1)

            input_field.send_keys(user_input)
            print(f"Sent '{user_input}' to {field_name}. Value after send_keys: '{input_field.get_attribute('value')}'")
            time.sleep(2)

            try:
                options = wait.until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[role='option']")),
                    message=f"Dropdown options not found for {field_name}"
                )
                print(f"Dropdown options for {field_name}: {[opt.text for opt in options]}")
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
                    user_inputs[field_name] = selected_value
                    print(f"Selected '{selected_value}' for {field_name}.\n")
                    return True
                else:
                    print(f"No matching option found for {field_name} with input '{user_input}'.")
                    retries -= 1
            except TimeoutException:
                print(f"Timeout waiting for dropdown options for {field_name}.")
                retries -= 1

            if retries > 0:
                print(f"Retrying ({retries} attempts left)...")
                time.sleep(2)
            else:
                print(f"Max retries reached; using '{user_input.upper()}' as fallback.")
                user_inputs[field_name] = user_input.upper()
                return True

        except TimeoutException as e:
            print(f"Timeout error: {e}")
            retries -= 1
            if retries > 0:
                print(f"Retrying ({retries} attempts left)...")
                time.sleep(2)
            else:
                print(f"Max retries reached; using '{user_input.upper()}' as fallback.")
                user_inputs[field_name] = user_input.upper()
                return True
        except Exception as e:
            print(f"Unexpected error for {field_name}: {e}")
            retries -= 1
            if retries > 0:
                print(f"Retrying ({retries} attempts left)...")
                time.sleep(2)
            else:
                print(f"Max retries reached; using '{user_input.upper()}' as fallback.")
                user_inputs[field_name] = user_input.upper()
                return True


def insert_under_heading(doc, level_2, level_3, company_name, collected_data, num_results=None):
    found_level_2 = False
    found_level_3 = False
    insert_position = None

    if level_2 not in existing_headings:
        existing_headings[level_2] = {}
    if level_3 not in existing_headings[level_2]:
        existing_headings[level_2][level_3] = set()

    first_in_category = company_name not in existing_headings[level_2][level_3]

    company_exists = company_name in all_companies
    header_text = f"{level_3} ({num_results if num_results is not None else 0})"

    for i, para in enumerate(doc.paragraphs):
        if para.style.name == 'Heading 2':
            if para.text.strip() == level_2:
                found_level_2 = True
                insert_position = i + 1
            elif found_level_2:
                break
        elif found_level_2 and para.style.name == 'Heading 3' and para.text.split(' (')[0].strip() == level_3:
            found_level_3 = True
            para.text = header_text
            insert_position = i + 1

    if not found_level_2:
        doc.add_heading(level_2, level=2)
        doc.add_heading(header_text, level=3)
        new_company = doc.add_heading(company_name, level=4)
        if not company_exists:
            for item in reversed(collected_data):
                doc.add_paragraph(item, style='List Bullet')
        else:
            print(f"Company '{company_name}' already exists in document; appended name only.")
        existing_headings[level_2][level_3].add(company_name)
        all_companies.add(company_name)
    elif not found_level_3:
        new_heading = doc.add_heading(header_text, level=3)
        doc.paragraphs[insert_position - 1]._element.addnext(new_heading._element)
        new_company = doc.add_heading(company_name, level=4)
        new_heading._element.addnext(new_company._element)
        if not company_exists:
            for item in reversed(collected_data):
                new_company._element.addnext(doc.add_paragraph(item, style='List Bullet')._element)
        else:
            print(f"Company '{company_name}' already exists in document; appended name only.")
        existing_headings[level_2][level_3].add(company_name)
        all_companies.add(company_name)
    else:
        new_company = doc.add_heading(company_name, level=4)
        doc.paragraphs[insert_position - 1]._element.addnext(new_company._element)
        if not company_exists:
            for item in reversed(collected_data):
                new_company._element.addnext(doc.add_paragraph(item, style='List Bullet')._element)
            existing_headings[level_2][level_3].add(company_name)
        else:
            print(f"Company '{company_name}' already exists in document; appended name only.")
        existing_headings[level_2][level_3].add(company_name)
        all_companies.add(company_name)


# Setup Selenium
options = Options()
options.headless = True
driver_path = GeckoDriverManager().install()

if not os.access(driver_path, os.X_OK):
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
                (By.CSS_SELECTOR, "p.font-roboto.font-semibold.text-subtitle1.m-0.undefined")))
            results_text = results_element.text.strip()
            num_results = int(''.join(filter(str.isdigit, results_text)))
            print(f"Number of Results from website for this query: {num_results}")
        except TimeoutException:
            num_results = 0
            print("Could not find number of results on the page.")

        sector = user_inputs.get("sector-1", "Unknown")
        fund_type = user_inputs.get("fund-type", "Unknown").upper()
        fund_specialty = user_inputs.get("fund-specialty", "Unknown").upper()
        category = f"{sector}-{fund_type}-{fund_specialty}"

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

            # Company LinkedIn URL
            company_linkedin_url = ""
            try:
                company_linkedin_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        "//a[contains(., 'LinkedIn') and not(ancestor::div[contains(., 'team') or contains(., 'Team')])]"
                    ))
                )
                company_linkedin_url = company_linkedin_element.get_attribute("href")
                print(f"Found company LinkedIn URL: {company_linkedin_url}")
            except TimeoutException:
                print("No company LinkedIn link found outside team section.")

            # Team section scraping
            has_team_section = detail_soup.find(
                lambda tag: tag.name in ['h2', 'h3', 'div'] and 'team' in tag.text.lower())
            team_info = []
            if has_team_section:
                print("Team section detected; scraping team info.")
                specific_p_team = detail_soup.find_all(
                    'p',
                    class_="font-roboto font-regular text-subtitle3 m-0 mb-[1rem]"
                ) or []
                team_texts = [p.get_text(strip=True) for p in specific_p_team if p.get_text(strip=True)]

                try:
                    team_linkedin_elements = WebDriverWait(driver, 5).until(
                        EC.presence_of_all_elements_located((
                            By.XPATH,
                            "//div[contains(., 'team') or contains(., 'Team')]//a[contains(., 'LinkedIn') and @href]"
                        ))
                    )
                    print(f"Found {len(team_linkedin_elements)} team LinkedIn links")
                except TimeoutException:
                    print("No team LinkedIn links found within 5 seconds.")
                    team_linkedin_elements = []

                for i, team_member in enumerate(team_texts):
                    team_linkedin_url = ""
                    if i < len(team_linkedin_elements):
                        try:
                            team_linkedin_url = team_linkedin_elements[i].get_attribute("href")
                            if team_linkedin_url == company_linkedin_url:
                                team_linkedin_url = ""
                                print(f"Skipping team LinkedIn URL that matches company URL for {team_member}")
                        except Exception as e:
                            print(f"Error getting team LinkedIn URL for {team_member}: {e}")

                    team_entry = f"{team_member}: {team_linkedin_url}" if team_linkedin_url else team_member
                    team_info.append(team_entry)

                full_team = ", ".join(team_info) if team_info else "No team members found"
            else:
                print("No team section found; skipping team info scraping.")
                full_team = ""

            address_parts = detail_soup.find_all(
                "p",
                class_="font-roboto font-regular text-subtitle3 m-0 text-center"
            ) or []
            address_parts = [p for p in address_parts if p.get_text(strip=True) and "ADDRESS" not in p.get_text()]
            full_address = ", ".join(p.get_text(strip=True) for p in address_parts)

            website_href = ""
            try:
                website_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//a[contains(., 'WEBSITE')]"))
                )
                website_href = website_element.get_attribute("href")
            except TimeoutException:
                print("Website link not found; setting as empty.")

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
            if company_linkedin_url:
                collected_data.append("Company LinkedIn URL: " + company_linkedin_url)
            if website_href:
                collected_data.append("Website URL: " + website_href)
            if email_address:
                collected_data.append("Email: " + email_address)
            if call_address:
                collected_data.append("Call: " + call_address)

            country = user_inputs.get("fund-country", "Unknown").upper()
            sector1 = user_inputs.get("sector-1", "Unknown")
            sector2 = user_inputs.get("sector-2", "Unknown")
            fund_type = user_inputs.get("fund-type", "Unknown").upper()
            fund_specialty = user_inputs.get("fund-specialty", "Unknown").upper()

            fund_types = ["VC", "PE", "LBO", "OPPORTUNIST"]
            fund_specialties = ["INCUBATION", "SEED", "EARLY", "GROWTH"]

            fund_type = fund_type if fund_type in fund_types else "Unknown"
            fund_specialty = fund_specialty if fund_specialty in fund_specialties else "Unknown"

            us_category = f"US-FUNDS-{sector1}-{sector2}-{fund_type}-{fund_specialty}"
            all_category = f"{sector1}-{sector2}-{fund_type}-{fund_specialty}"

            company_name = specific_h4_company_name.get_text(
                strip=True) if specific_h4_company_name else "Unknown Company"

            if country == "US":
                insert_under_heading(doc, "US FUNDS", us_category, company_name, collected_data, num_results)
            insert_under_heading(doc, "ALL FUNDS", all_category, company_name, collected_data, num_results)

            print(f"Processed {company_name} for categories: {us_category if country == 'US' else ''} {all_category}")

            if len(driver.window_handles) > len(old_windows):
                driver.close()
                driver.switch_to.window(old_windows[0])
            else:
                driver.back()

            driver.execute_script("arguments[0].remove();", view_button)
            time.sleep(2)

        doc.save(DOC_PATH)
        print(f"Updated document saved: {DOC_PATH}")

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