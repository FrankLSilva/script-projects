import os
from easygui import passwordbox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# LinkedIn Keys
email_key = "session_key"
password_key = "session_password"
login_key = "cursor-pointer"
show_all_key = "discovery-templates-vertical-list__footer"
jobrole_key = "jobs-search-box__text-input"
location_key = "jobs-search-box__text-input"
simple_apply_key = "search-reusables__filter-pill-button"
search_key = "jobs-search-box__submit-button"
jobcard_key = "job-card-container--clickable"
easy_key = "jobs-apply-button"

def linkedin_credentials():
    print("\nLinkedIn EASY APPLY Automation.\n"
          "Follow all the steps below!\n"
          "\n--------------------------------")

    jobrole = input("Jobrole: ")
    email = input("Email: ")
    password = passwordbox("LinkedIn Password:")

    print("--------------------------------\n"
          "Login in... plese wait!")

    os.environ['LINKEDIN_JOBROLE'] = jobrole
    os.environ['LINKEDIN_EMAIL'] = email
    os.environ['LINKEDIN_PASSWORD'] = password

# Aplication Start
linkedin_credentials()
url = "https://www.linkedin.com/jobs"
driver = webdriver.Chrome()
driver.get(url)

# Credential Fill & Login
email = driver.find_element(By.ID, email_key)
email.send_keys(os.environ.get('LINKEDIN_EMAIL'))
password = driver.find_element(By.ID, password_key)
password.send_keys(os.environ.get('LINKEDIN_PASSWORD'))

# Login
login = driver.find_element(By.CLASS_NAME, login_key)
login.click()

# Security Check
input("After security check, tap ENTER here to continue!\n"
      "--------------------------------")

# Job start fill
show_all_bt = driver.find_element(By.CLASS_NAME, show_all_key)
show_all_bt.click()

driver.implicitly_wait(5)
job_field = driver.find_element(By.CLASS_NAME, jobrole_key)
job_field.send_keys(os.environ.get('LINKEDIN_JOBROLE'))

search_call = driver.find_element(By.CLASS_NAME, search_key)
search_call.click()

input("Please determine:\n"
      "- Location\n"
      "- Filters (optional)\n"
      "- Easy Apply (mandatory!!!)"
      "\n--------------------------------\n"
      "Press ENTER here to start the automation!\n"
      "--------------------------------\n")

# Automation
start_at = True
while start_at:
    # Scroll to the bottom to load more content
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CLASS_NAME, jobcard_key)))

    job_list = driver.find_elements(By.CLASS_NAME, jobcard_key)
    call = 0

    for job in job_list:
        driver.execute_script("arguments[0].scrollIntoView();", job)
        driver.implicitly_wait(2)
        job.click()
        call += 1
        print(f"Application {call}")

        easy_button = driver.find_element(By.CLASS_NAME, easy_key)
        easy_button.click()
        input("test")

    at_command = input("\nStart Automation (Y/N)?: ").lower()
    if at_command != "y":
        start_at = False

# Aplication close
input("\n"
      "Tap ENTER here to stop the application!\n"
      "Or just close it through the Window!")

driver.close()