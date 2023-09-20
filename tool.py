from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
import re
import win32com.client

import json
import pandas as pd
import time

def readExcel():
    # Read the Excel file
    df = pd.read_excel('data2.xlsx')
    # Extract the plan IDs from column F
    plan_ids = df['PlanID'].dropna().tolist()
    return plan_ids


def send_outlook_email(to_email, cc_emails):
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to_email
    if cc_emails:
        mail.CC = "; ".join(cc_emails)
    mail.Subject = 'Action Needed: Quarterly Invoice - 401(k) Payment'
    # Set the email body
    mail.Body = """Good morning,

I hope you are well! I am reaching out concerning your most recent quarterly bill.

....

I’ll be happy to answer any questions!
    """
    mail.Display(True)  # This will display the email, remove this line to send the email directly

plan_ids = readExcel()

# Initialize the Selenium web driver (assuming you're using Chrome)
#driver = webdriver.Chrome()

# Connect to the existing Chrome session
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

driver = webdriver.Chrome(options=chrome_options)

wait = WebDriverWait(driver, 20)  # Increase to 20 seconds

# Wait for the page to load

# Loop through each plan ID and perform lookup
for plan_id in plan_ids:
    driver.get(f"https://crm-link.com/search?searchword={plan_id}")
  

   
    # Wait for the search results to load
    wait.until(EC.presence_of_element_located((By.ID, "gsearchDiv")))

    # Wait for the link to be clickable
    wait.until(EC.element_to_be_clickable((By.XPATH, '//a[contains(@href, "/crm/org761441520/EntityInfo.do")]')))

    # Find the link and click it
    link_to_click = driver.find_element(By.XPATH, '//a[contains(@href, "/crm/org761441520/EntityInfo.do")]')
    link_to_click.click()


    # Wait for the new page to load
    time.sleep(5)  # Adjust as needed
    
    # Retrieve all email addresses from mailto links
    all_links = driver.find_elements(By.TAG_NAME,"a")
    email_addresses = []
    # Find the section element
    section_elements = driver.find_elements(By.ID, 'RelatedListCommonDiv')
    section_element = section_elements[1]

    # Scroll to the section element
    driver.execute_script("arguments[0].scrollIntoView();", section_element)
    wait.until(EC.presence_of_element_located((By.ID, "RelatedList_4980552000195072001")))
    # Wait until elements are located
    # Fetch all iframes in the page
    iframe = driver.find_element(By.ID,"RelatedList_4980552000195072001")

    #
    iframe_name = iframe.get_attribute('name')
    driver.switch_to.frame(iframe)
    time.sleep(5)
    section_text = driver.find_element(By.TAG_NAME,"body").text
    
    # Use regex to find all email addresses in that text
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    email_addresses = re.findall(email_pattern, section_text)

    
    to_email = email_addresses[0]
    cc_emails = email_addresses[1:]
    print(to_email, cc_emails)
    send_outlook_email(to_email, cc_emails)

    


# Close the web browser
driver.quit()
