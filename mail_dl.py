#pip install selenium

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import os
import shutil

# Replace with the path to your Chrome WebDriver executable
chrome_driver_path = '/path/to/chromedriver'

# Replace with your Outlook Web login credentials
username = 'your_username'
password = 'your_password'

# Initialize the Chrome WebDriver
driver = webdriver.Chrome(chrome_driver_path)

# Step 1: Log in to Outlook Web
driver.get('https://outlook.office365.com/')  # Adjust the URL as needed
driver.find_element_by_id('i0116').send_keys(username)
driver.find_element_by_id('i0118').send_keys(password)
driver.find_element_by_id('i0118').send_keys(Keys.RETURN)

# Add a delay to ensure the login is complete (you can adjust the time)
time.sleep(5000)

# Step 2: Navigate to the "Universities" category
category_name = "Universities"  # Replace with the actual category name
category_selector = driver.find_element_by_xpath(f"//span[contains(text(), '{category_name}')]")
category_selector.click()

# Step 3: Iterate through emails and download .xlsx attachments
email_elements = driver.find_elements_by_css_selector(".email")  # Adjust the selector as needed

for email_element in email_elements:
    email_element.click()  # Click on the email to open it
    email_subject = driver.find_element_by_css_selector(".email-subject").text
    
    # Locate and download .xlsx attachments from the opened email
    attachment_elements = driver.find_elements_by_css_selector(".attachment")  # Adjust the selector as needed
    
    for attachment_element in attachment_elements:
        attachment_name = attachment_element.text
        if attachment_name.endswith(".xlsx"):
            attachment_element.click()  # Click to download the attachment
            
            # Specify the download location (you may need to adjust this)
            download_location = "/path/to/download/folder/"  # Replace with your desired folder
            
            # Save the attachment with the email subject as the filename
            attachment_filename = f"{email_subject}.xlsx"
            download_path = os.path.join(download_location, attachment_filename)
            
            # You may need to handle file download dialogs here
            # For example, using Selenium's WebDriverWait
            
            # Move the downloaded file to the specified location
            shutil.move(attachment_filename, download_path)

# Step 4: Save the email as a PDF
print_menu = driver.find_element_by_css_selector(".print-menu")  # Adjust the selector as needed
print_menu.click()

# Select the "Save as PDF" option
save_as_pdf_option = driver.find_element_by_xpath("//span[contains(text(), 'Save as PDF')]")
save_as_pdf_option.click()

# Specify the PDF save location (you may need to interact with file dialogs)
# Confirm the print action

# Step 5: Repeat for other emails

# Step 6: Logout and close the browser
driver.find_element_by_id('id_l').click()  # Click on the logout button
driver.quit()