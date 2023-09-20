from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
import re
import win32com.client
import tkinter as tk
from tkinter import filedialog
import ttkbootstrap as ttk
import subprocess

import json
import pandas as pd
import time

text_padding = 5
main = ttk.Window(themename='yeti')
main.title("Catherine's Emailer")
windowcolor = tk.StringVar()
windowcolor.set('yeti')
style = ttk.Style()

def readExcel():
    # Read the Excel file
    df = pd.read_excel('data.xlsx')
    # Extract the plan IDs from column F
    plan_ids = df['PlanID'].dropna().tolist()
    return plan_ids


def send_outlook_email(to_email, cc_emails, subject, body):
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to_email
    if cc_emails:
        mail.CC = "; ".join(cc_emails)
    mail.Subject = subject
    # Set the email body
    mail.Body = body
    mail.Display(True)  # This will display the email, remove this line to send the email directly

def run_scraper(subject, body, clientlist):
    if len(clientlist) > 1:
        plan_ids = clientlist
        print(plan_ids)
    else: 
        print("No list input, looking for Excel sheet")
        plan_ids = readExcel()
    options_button.config(state=ttk.DISABLED)
    subprocess.run(r'"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\Users\crainey\Downloads\temp"', shell=True)
    # Initialize the Selenium web driver (assuming you're using Chrome)
    driver = webdriver.Chrome()

    # Connect to the existing Chrome session
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

    
    driver = webdriver.Chrome(options=chrome_options)

    wait = WebDriverWait(driver, 20)  # Increase to 20 seconds

    # Wait for the page to load
    # Loop through each plan ID and perform lookup
    for plan_id in plan_ids:
        driver.get(f"https://crm.zoho.com/crm/org761441520/search?searchword={plan_id}")
    

    
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
        send_outlook_email(to_email, cc_emails, subject, body)

        


    # Close the web browser
    driver.quit()
    options_button.config(state=ttk.NORMAL)
    
left_frame = ttk.Frame(main )
left_frame.pack(side=ttk.LEFT, padx=10)
client_list_label = ttk.Label(left_frame,text="Client List:")
client_list_label.pack(padx=text_padding, pady = text_padding)
client_list = ttk.Text(left_frame,width=100)
client_list.pack(padx=text_padding, pady = text_padding)

right_frame = ttk.Frame(main)
right_frame.pack(side=ttk.LEFT, padx=10)
email_subject_label = ttk.Label(right_frame, text="Email Subject Line:")
email_subject_label.pack(padx=text_padding, pady = text_padding)
email_subject = ttk.Entry(right_frame,width=100)
email_subject.pack(padx=text_padding, pady = text_padding)

email_body_label = ttk.Label(right_frame,text="Email Body:")
email_body_label.pack(padx=text_padding, pady = text_padding)
email_body = ttk.Text(right_frame,width=100)
email_body.pack(padx=text_padding, pady = text_padding)

options_button = ttk.Button(main,text = 'Start Emailer', command= lambda: run_scraper(email_subject.get(), email_body.get(1.0, ttk.END), client_list.get(1.0,ttk.END),))
options_button.pack(padx=text_padding, pady=text_padding)

main.mainloop()
