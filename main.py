from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from dotenv import load_dotenv
import pandas as pd
import time
import os

load_dotenv()

userDir = os.getenv('USER_DIR')
profile = os.getenv('PROFILE')
sheetName = os.getenv('SHEET_NAME')
website = os.getenv('WEBSITE')
companyCode = os.getenv('COMPANY_CODE')

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument(f"user-data-dir={userDir}")
chrome_options.add_argument(f"profile-directory={profile}")

service = Service(executable_path="chromedriver.exe")
driver = webdriver.Chrome(service=service,options=chrome_options)
 
try: 
    driver.get(website)

    driver.find_element(By.CSS_SELECTOR,"#zl-myapps > div.ea-app-container > div:nth-child(4) > div > a").click() #Books
    driver.implicitly_wait(20)

    organisation ="#ember11" if companyCode=="UK" else "#ember14"

    # driver.find_element(By.CSS_SELECTOR,"#ember14 > div > div > div.col-lg-4 > div > button").click() #Organisation
    driver.find_element(By.CSS_SELECTOR,f"{organisation} > div > div > div.col-lg-4 > div > button").click() #Organisation
    driver.implicitly_wait(20)

    driver.find_element(By.XPATH,"/html/body/div[6]/div[5]/div[3]/nav/ul/li[6]/div/div/h3/button").click() #Purchase
    driver.implicitly_wait(20)

    driver.find_element(By.XPATH,"//a[text()='Vendors']").click() #Vender
    driver.implicitly_wait(20)

    driver.find_element(By.XPATH,"//button[text()='New']").click() #New
    driver.implicitly_wait(20)

    vendors = pd.read_excel('vendors.xlsx',sheet_name=sheetName, dtype={'Employee Bank Account Number': str})

    for index,row in vendors.iterrows():
        # print(row)
        driver.find_element(By.XPATH,"//input[@placeholder='Salutation']").click()
        driver.implicitly_wait(5)

        driver.find_element(By.XPATH,"//div[text()='Mr.']").click()
        driver.find_element(By.XPATH,"//input[@placeholder='First Name']").send_keys(f"{row['First Name']}") #First Name
        driver.find_element(By.XPATH,"//input[@placeholder='Last Name']").send_keys(f"{row['Last Name']}") #Last Name
        driver.find_element(By.XPATH,"/html/body/div[6]/div[5]/div[3]/main/div/div/form/div[2]/div[2]/div/div/div[3]/div/div/div/div/div/div[1]/div/input").send_keys(f"Mr. {row['Employee Name']}") # Display name
        driver.find_element(By.XPATH,"//label[text()='Company Name']/following-sibling::div[@class='col-lg-6']/input[@class='ember-text-field ember-view form-control' and @type='text']").send_keys(f"{row['Employee Name']}") # comapay name
        driver.find_element(By.XPATH,"//input[@placeholder='Mobile']").send_keys(f"{row['Mobile/WhatsApp']}") # Mobile

        if pd.notna(row["Employee Email ID"]) and row["Employee Email ID"].strip() != "":
            driver.find_element(By.XPATH,"//span[@class='form-icon icon-left']/following-sibling::input[@aria-label='Email Address']").send_keys(f"{row['Employee Email ID']}") # email id

        if companyCode != "UK":
            driver.find_element(By.XPATH,"//span[text()='Select a GST treatment']").click() # GST
            driver.implicitly_wait(5)

            gst_serach=driver.switch_to.active_element
            gst_serach.send_keys('Unregistered Business')
        
            # gst_aria_activedescendant = gst_serach.get_attribute("aria-activedescendant")
            # gst_aria_controls = gst_serach.get_attribute("aria-controls")
            # gst_aria_label = gst_serach.get_attribute("aria-label")
            # gst_autocorrect = gst_serach.get_attribute("autocorrect")
            # gst_autocomplete = gst_serach.get_attribute("autocomplete")
            # gst_spellcheck = gst_serach.get_attribute("spellcheck")
            # gst_placeholder = gst_serach.get_attribute("placeholder")
            # gst_aria_describedby = gst_serach.get_attribute("aria-describedby")
            # gst_class = gst_serach.get_attribute("class")
            # gst = driver.find_element(By.XPATH,f"//input[@aria-activedescendant={gst_aria_activedescendant} and @aria-controls={gst_aria_controls} and @aria-label={gst_aria_label} and @autocorrect={gst_autocorrect} and @autocomplete={gst_autocomplete} and @spellcheck={gst_spellcheck} and @placeholder={gst_placeholder} and @aria-describedby={gst_aria_describedby} and @class={gst_class}]")
            # print(gst.get_attribute("outerHTML"))
            driver.implicitly_wait(5)
            driver.switch_to.active_element.send_keys(Keys.RETURN)
            driver.find_element(By.XPATH,"//input[@aria-label='Search']").send_keys(Keys.RETURN)
            driver.implicitly_wait(5)

        if pd.notna(row['PAN']) and row['PAN'].strip() != "":
            driver.find_element(By.XPATH,"/html/body/div[6]/div[5]/div[3]/main/div/div/form/div[3]/div[2]/div[1]/div/div/div/div[3]/div/div/input").send_keys(f"{row['PAN']}") #PAN

        # Address Fields
        if pd.notna(row['Location']) and row['Location'].strip() != "":
            driver.find_element(By.XPATH,"//div[text()='Address']").click()
            driver.implicitly_wait(5)

            driver.find_element(By.XPATH,"/html/body/div[6]/div[5]/div[3]/main/div/div/form/div[3]/div[2]/div[2]/div/div/div/div[1]/fieldset[1]/div/div/input").send_keys(f"Mr. {row['Employee Name']}") # Attention
            driver.find_element(By.XPATH,"/html/body/div[6]/div[5]/div[3]/main/div/div/form/div[3]/div[2]/div[2]/div/div/div/div[1]/fieldset[2]/div/div/div/div/div[1]/div/div/span").click()
            driver.implicitly_wait(5)

            driver.switch_to.active_element.send_keys(f"{row['Location']}") # Region
            driver.implicitly_wait(5)
            driver.switch_to.active_element.send_keys(Keys.RETURN)

            driver.find_element(By.XPATH,"//textarea[@placeholder='Street 1'][1]").send_keys(f"{row['Employee Address']}") #street 1
            driver.implicitly_wait(5)

            driver.find_element(By.XPATH,"/html/body/div[6]/div[5]/div[3]/main/div/div/form/div[3]/div[2]/div[2]/div/div/div/div[1]/fieldset[5]/div/div/input").send_keys(f"{row['city']}") #city
            driver.find_element(By.XPATH,"//input[@placeholder='Select or type to add' and @aria-label='zb.common.state'][1]").click() # state
            driver.implicitly_wait(5)

            driver.switch_to.active_element.send_keys(f"{row['state']}") #state
            driver.implicitly_wait(5)
            driver.switch_to.active_element.send_keys(Keys.RETURN)

            driver.find_element(By.XPATH,"/html/body/div[6]/div[5]/div[3]/main/div/div/form/div[3]/div[2]/div[2]/div/div/div/div[1]/fieldset[7]/div/div/input").send_keys(f"{row['pincode']}") #pincode

            driver.find_element(By.XPATH,"//a[text()='Copy billing address']").click()
            driver.implicitly_wait(5)

        # Bank Details
        if pd.notna(row['Bank Name']) and row['Bank Name'].strip() != "":
            driver.find_element(By.XPATH,"//div[text()='Bank Details']").click()
            driver.implicitly_wait(5)

            driver.find_element(By.XPATH,"//a[text()='Add Bank Account']").click()
            driver.implicitly_wait(5)


            driver.find_element(By.XPATH,"//label[text()='Account Holder Name']/following-sibling::div/input").send_keys(f"{row['Bank Account Holder Full Name']}") # Holder Name
            driver.find_element(By.XPATH,"//label[text()='Bank Name']/following-sibling::div/input").send_keys(f"{row['Bank Name']}") # Bank Name
            driver.find_element(By.XPATH,"//label[text()='Account Number']/following-sibling::div/input").send_keys(f"{row['Employee Bank Account Number']}") # Account Number
            driver.find_element(By.XPATH,"//label[text()='Re-enter Account Number']/following-sibling::div/input").send_keys(f"{row['Employee Bank Account Number']}") # Re Account Number
            driver.find_element(By.XPATH,"//label[text()='IFSC']/following-sibling::div//input").send_keys(f"{row['Bank Sort Code/IFSC Code']}") # IFSC

        # time.sleep(30) # Testing purpose
        driver.find_element(By.XPATH,"//button[@type='submit' and text()='Save']").click() #Save
        driver.implicitly_wait(10)

        if index < len(vendors)-1:
            driver.find_element(By.XPATH,"//div[@class='btn-toolbar float-end ']/button[@class='btn btn-primary']").click()
            driver.implicitly_wait(10)

        # if index < 1: # Testing purpose
        #     break

    time.sleep(20)
except:
  print("Something went wrong")
finally:
    driver.quit()