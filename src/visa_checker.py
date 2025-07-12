from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException
import ctypes
from datetime import datetime
from dotenv import load_dotenv
from time import strptime
import calendar
import logging
import requests
import time
import sys
import os

# --- CONFIG ---
load_dotenv()
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
SCHEDULE_ID = os.getenv("SCHEDULE_ID")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID=os.getenv("TELEGRAM_CHAT_ID")
CHROME_DRIVER_PATH=os.getenv("CHROME_DRIVER_PATH")


# -------- URLs --------
LOGIN_URL = 'https://ais.usvisa-info.com/en-ca/niv/users/sign_in'  # Change 'en-ca' based on your country
APPOINTMENT_URL = f'https://ais.usvisa-info.com/en-ca/niv/schedule/{SCHEDULE_ID}/appointment'


# --- Setup Selenium ---
options = webdriver.ChromeOptions()
service = Service(executable_path=CHROME_DRIVER_PATH)
driver = webdriver.Chrome(service=service, options=options)


# -------- Logging Setup --------
log_date = datetime.now().strftime("%Y%m%d")
log_filename = f"visa_checker_{log_date}.log"
logging.basicConfig(filename=log_filename, level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s',)


# -------- Check Session --------
def check_if_session_expired():
    current_url = driver.current_url
    if "sign_in" in current_url or "login" in current_url:
        logging.warning("Session expired or redirected to login.")
        send_telegram_alert("🔐 Session expired. Re-logging in.")
        login()


# -------- Send Telegram Alert --------
def send_telegram_alert(message):
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        payload = {
            "chat_id": TELEGRAM_CHAT_ID,
            "text": message
        }
        response = requests.post(url, data=payload)
        if response.status_code == 200:
            logging.info("Telegram alert sent.")
        else:
            logging.warning(f"Telegram error: {response.text}")
    except Exception as e:
        logging.error(f"Telegram exception: {e}")


# -------- Login Session --------
def login():
    driver.get(LOGIN_URL)
    time.sleep(2) # Increase wait time to ensure page loads completely
    try:
        # Try to find elements with explicit waits
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)
        email_input = driver.find_element(By.ID, "user_email")
        password_input = driver.find_element(By.ID, "user_password")
        email_input.send_keys(EMAIL)
        password_input.send_keys(PASSWORD)
        checkbox_div = driver.find_element(By.CSS_SELECTOR, "div.icheckbox")
        driver.execute_script("arguments[0].click();", checkbox_div)
        driver.find_element(By.NAME, "commit").click() # Find the login button and click it
        send_telegram_alert("🔓 Login Successful")
        time.sleep(2)
    except Exception as e:
        logging.error(f"Error during login: {e}")
        send_telegram_alert("⚠️ Error during login.")
        logging.error("Page source:", driver.page_source[:500])  # Print first 500 chars for debugging


# -------Select Toronto from dropdown-------
def select_toronto_location():
    location_dropdown = driver.find_element(By.ID, 'appointments_consulate_appointment_facility_id')
    for option in location_dropdown.find_elements(By.TAG_NAME, 'option'):
        if "Toronto" in option.text:
            option.click()
            break
    time.sleep(1)            # future - see if this can be reduced / removed


# --------Click to open the date picker-----
def open_calendar():
    date_input = driver.find_element(By.ID, 'appointments_consulate_appointment_date')
    date_input.click()
    time.sleep(0.5)


# ----Select Time Slot and Confirm --------------
def select_time_slot_and_confirm(earliest, timeout=8, poll_frequency=0.5):
    """Swiftly selects the first valid time slot and confirms the appointment using the 'Reschedule' button."""

    start_time = time.time()
    try:
        # Step 1: Wait for the dropdown to appear
        #WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, "appointments_consulate_appointment_time")))
        WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, "appointments_consulate_appointment_time"))) 
        # Step 2: Keep checking for real options
        while time.time() - start_time < timeout:
            try:
                time_dropdown_element = driver.find_element(By.ID, "appointments_consulate_appointment_time")
                time_dropdown_element.click()
                dropdown = Select(time_dropdown_element)
                options = dropdown.options

                if len(options) > 0:
                    dropdown.select_by_index(1)
                    selected_time = options[1].text
                    logging.info(f"Time slot selected: {selected_time}")

                    # Step 3: Find the 'Reschedule' button (priority)
                    confirm_button = None
                    button_texts = ['Reschedule', 'Confirm', 'Submit']  # Prioritized

                    for text in button_texts:
                        buttons = driver.find_elements(By.XPATH, f"//input[@value='{text}'] | //button[contains(text(), '{text}')]")
                        if buttons:
                            confirm_button = buttons[0]
                            break

                    if confirm_button:
                        confirm_button.click()
                        logging.info(f"Clicked confirmation button: {text}")
                        appt_time = selected_time.split(":")
                        earliest_datetime = earliest.replace(hour= int(appt_time[0]), minute= int(appt_time[1]))
                        return earliest_datetime
                    else:
                        logging.warning("⚠️ Could not find confirmation button.")
                        return None

            except (NoSuchElementException, StaleElementReferenceException) as e:
                # DOM likely still loading; retry
                # filename = datetime.now().strftime("screenshot_%Y%m%d_%H%M%S.png")
                # driver.execute_script("window.scrollBy(0, 300);")   # Scroll down by 300 pixels
                #time.sleep(1)
                # driver.save_screenshot(filename)
                logging.error(f"Error selecting time slot: {e}")
                send_telegram_alert("⚠️ Error selecting time slot")
                

            time.sleep(poll_frequency)

        logging.warning("⚠️ Time slot selection timed out.")
        filename = datetime.now().strftime("screenshot_%Y%m%d_%H%M%S.png")
        driver.execute_script("window.scrollBy(0, 300);")   # Scroll down by 300 pixels
        #time.sleep(1)
        driver.save_screenshot(filename)
        send_telegram_alert("⚠️ Time slot selection timed out.")
        return None

    except TimeoutException:
        logging.error("⚠️ Time dropdown failed to load.")
        return None 
        

# --------Fetch Earliest Avail date-----
def get_earliest_available_date():
    valid_months = ["September", "October", "November", "December", "January", "February"]
    valid_year = [2025, 2026]
    earliest = None
    earliest_datetime = None

    for _ in range(7):  # Max 4 months ahead
        #time.sleep(1)
        
        # Get current month and year shown
        month_elem = driver.find_element(By.CLASS_NAME, 'ui-datepicker-month').text
        year_elem = driver.find_element(By.CLASS_NAME, 'ui-datepicker-year').text
        prev_month = month_elem.strip()
        current_year = int(year_elem.strip())
        actual_month = strptime(prev_month,'%B').tm_mon + 1
        if actual_month != 13:
            current_month = calendar.month_name[actual_month]
        else:
            current_month = "January"
            current_year = int(year_elem.strip()) + 1


        # --- Skip if year ≠ 2025 or month not in valid list ---
        if current_year not in valid_year or current_month not in valid_months:
            logging.info(f"Skipping: {current_month} {current_year}")
            next_button = driver.find_element(By.CSS_SELECTOR, ".ui-datepicker-next")
            next_button.click()
            continue


        # Find all clickable (available) date elements
        logging.info(f"Checking: {current_month} {current_year}")
        dates = driver.find_elements(By.CSS_SELECTOR, 'td > a.ui-state-default')
        if dates:
            first_day = int(dates[0].text)
            full_date = datetime.strptime(f"{first_day} {current_month} {current_year}", "%d %B %Y")
            earliest = full_date
            
            #  Select the date by clicking
            try:
                dates[0].click()
                logging.info(f"Selected earliest date: {full_date.strftime('%B %d, %Y')}")
                # time.sleep(1)
            except Exception as e:
                logging.error(f"Error selecting date: {e}")
                break


            # Select the first time slot 
            earliest_datetime = select_time_slot_and_confirm(earliest)
            send_telegram_alert(f"Selected earliest date: {full_date.strftime('%B %d, %Y')}")
            break  # ✅ Exit loop after date/time confirmed
            

        
        # Click next month
        next_button = driver.find_element(By.CSS_SELECTOR, ".ui-datepicker-next")
        next_button.click()

    return earliest_datetime

def check_visa_availability(retry_delay = 60):
    attempt = 1
    busy_count = 0

    while True:
        driver.get(APPOINTMENT_URL)
        time.sleep(1)

        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        logging.info(f"Attempt {attempt}: Checking appointment page...")


        # --- Select Toronto ---
        try:
            select_toronto_location()
            logging.info("Selected Toronto in dropdown.")
        except (NoSuchElementException) as e:
            logging.warning(f"Webpage did not load during retry")
            send_telegram_alert("⚠️ Webpage did not load during retry")
            attempt += 1
            time.sleep(15)
            check_if_session_expired()
            continue
        except Exception as e:
            logging.error(f"Failed to select Toronto: {e}")
            send_telegram_alert("⚠️ Failed to select Toronto in dropdown")
            attempt += 1
            time.sleep(2)
            check_if_session_expired()      # future - this is not req
            continue 

        # --- open calendar widget ---
        try:
            open_calendar()
            logging.info("Opened calendar widget.")
            busy_count = 0  # Reset on success
        except Exception as e:
            logging.error(f"Failed to open calendar / System Busy")
            if busy_count > 4:
                msg = "🌙 System busy for 50 mins. Exiting and Hibernating machine."
                send_telegram_alert(msg)
                logging.critical(msg)
                time.sleep(5)
                os.system("shutdown /h")  # Windows Hibernate in 5 seconds
                break  # Exit program loop
            attempt += 1
            busy_count += 1
            time.sleep(retry_delay * 10) #10min wait when System is Busy
            continue

        
        # --- Get and select earliest available appointment ---
        earliest_date = get_earliest_available_date()
        if earliest_date:
            msg = f"✅ Visa slot confirmed for {earliest_date.strftime('%B %d, %Y')} at {earliest_date.strftime('%H:%M')}"
            send_telegram_alert(msg)
            logging.info(msg)
            logging.info("Exiting program after successful booking.")
            sys.exit(0)  # Exit loop
        else:
            logging.info("No available dates in Sept to Feb")
            if attempt % 30 == 0:
                attempt_msg = f"🔄 Attempt #{attempt}: No Earliest dates Avail. Still checking for visa slots..."
                send_telegram_alert(attempt_msg)
                logging.info(attempt_msg)

        # 🔄 Retry loop
        logging.info("Waiting before next retry...\n")
        attempt += 1
        time.sleep(retry_delay)

        
#------- Main Loop --------
try:
    login()
    check_visa_availability()
except KeyboardInterrupt:
    logging.warning("Stopped by user.")
finally:
    driver.quit()