import os
print(os.getcwd())
os.chdir("C:/Users/dojst/OneDrive/Documents/Synengco/webFill")

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
# imports for waiting for elements to appear
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd

validStreets = ['Alley', 'Arcade', 'Avenue', 'Boulevard', 'Bypass', 'Circuit', 'Close', 'Corner', 'Court', 'Crescent', 
    'Cul-de-sac', 'Drive', 'Esplanade', 'Green', 'Grove', 'Highway', 'Junction', 'Lane', 'Link', 'Mews', 'Parade', 'Place', 
    'Ridge', 'Road', 'Square', 'Street', 'Terrace', 'Ally', 'Arc', 'Ave', 'Bvd', 'Bypa', 'Cct', 'Cl', 'Crn', 'Ct', 'Cres', 
    'Cds', 'Dr', 'Esp', 'Grn', 'Gr', 'Hwy', 'Jnc', 'Lane', 'Link', 'Mews', 'Pde', 'Pl', 'Rdge', 'Rd', 'Sq', 'St', 'Tce']

states = ['ACT', 'NSW', 'NT', 'QLD', 'SA', 'TAS', 'VIC', 'WA', 'Australian Capital Territory', 'Queensland', 'Victoria',
    'New South Wales', 'Northern Territory', 'South Australia', 'Tasmania', 'Western Australia']

print("This is a script to automatically login to zoho via a microsoft account.")
email = "d.jones3@uqconnect.edu.au" # input("Enter Email: ")
mspassword = "Om*cron27" #input("Enter Password: ")
driver = webdriver.Chrome()
driver.get("https://accounts.zoho.com.au/signin?servicename=ZohoHome&signupurl=https://www.zoho.com/signup.html")

#microsoftLogin = driver.find_element_by_class_name("fed_div azure_fed show_fed")
wait = WebDriverWait(driver, 10)
microsoftLogin = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".fed_div.azure_fed.show_fed")))
microsoftLogin.click()

#Enter email
element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".form-control.ltr_override")))
element.send_keys(email)

#Press next
nextButton = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".btn.btn-block.btn-primary")))
nextButton.click()

#enter password
passwdEnter = wait.until(EC.visibility_of_element_located((By.ID, "i0118")))
passwdEnter.send_keys(mspassword)

#sign in
signInButton = wait.until(EC.visibility_of_element_located((By.ID, "idSIButton9")))
signInButton.click()

yesButton = wait.until(EC.visibility_of_element_located((By.ID, "idSIButton9")))
yesButton.click()

driver.get("https://crm.zoho.com.au/crm/org7000126776/tab/Leads/create")
#driver.close()

xls = pd.ExcelFile("CamCard_Contacts_XLS.xls")
sheet = xls.parse()
firstName = sheet['First Name']
lastName = sheet['Last Name']
company = sheet['Company1']
jobTitle = sheet['Job Title1']
telephone1 = sheet['Tel1']
telephone2 = sheet['Tel2']
telephone3 = sheet['Tel(Other)']
email1 = sheet['Email1']
address1 = sheet['Address1']
webPage = sheet['Web Page']
fax1 = sheet['Fax1']
address1 = sheet['Address1']

def getPhoneAndMobile(telephone1, telephone2, telephone3):
    telephones = [telephone1, telephone2, telephone3]
    numbers = {"phone" : "", "mobile" : ""}
    for telephone in telephones:
        telephone = str(telephone)
        if len(telephone) < 10:
            if len(telephone) != 0:
                print("short telephone number?")
            continue
        if telephone[0] == '+':
            telephone = telephone[1:]
        if telephone[0:2] == '61':
            telephone = '0' + telephone[2:]
        if telephone[1] == '7':
            numbers['phone'] = telephone
        if telephone[1] == '4':
            numbers['mobile'] = telephone
    return numbers

def getFax(faxNumber):
    fax = str(faxNumber)
    if len(fax) < 10:
        if len(fax) != 0:
            print("short fax?")
        return ""
    if fax[0] == '+':
        fax = fax[1:]
    if fax[0:2] == '61':
        fax = '0' + fax[2:]
    return fax
    

def getWebsite(webPage):
    if len(webPage) == 0:
        return ""
    if (webPage[0:10] == 'Home Page:'):
        webPage = webPage[10:]
    return webPage

possibleStreetTypes = []

def getAddressDict(address):
    resultDict = {
        "street" : "",
        "city" : "",
        "state" : "",
        "postCode" : "",
        "country" : "Australia"
    }
    if len(address) == 0:
        return address
    streetTypeIndex = -1
    for i in range(len(address)):
        for j in validStreets:
            if address[i:i + len(j) + 1].lower() == " " + j.lower():
                streetTypeIndex = i + len(j) + 1
                break
        if streetTypeIndex != -1:
            break
    resultDict['street'] = address[0:streetTypeIndex]

    stateIndex = -1
    stateLength = 0
    for i in range(streetTypeIndex, len(address)):
        for j in states:
            if (address[i:i + len(j) + 1].lower() == " " + j.lower()) or (
                address[i:i + len(j) + 1].lower() == "," + j.lower()):
                stateIndex = i + len(j) + 1
                stateLength = len(j)
                break
        if stateIndex != -1:
            break
    resultDict['state'] = address[stateIndex - stateLength : stateIndex]
    resultDict['city'] = address[streetTypeIndex : stateIndex - stateLength].strip(' ,')

    postCodeIndex = -1
    for i in range(stateIndex, len(address) - 3):
        if (address[i].isdigit() and address[i + 1].isdigit() and
            address[i + 2].isdigit() and address[i + 3].isdigit()):
            postCodeIndex = i
    resultDict['postCode'] = address[postCodeIndex:postCodeIndex + 4]
    
    return resultDict


for i in range(len(firstName)):
    input("Press enter to add lead: ")

    firstNameField = wait.until(EC.visibility_of_element_located((By.ID, "Crm_Leads_FIRSTNAME")))
    firstNameField.send_keys(firstName[i])

    lastNameField = wait.until(EC.visibility_of_element_located((By.ID, "Crm_Leads_LASTNAME")))
    lastNameField.send_keys(lastName[i])

    companyField = wait.until(EC.visibility_of_element_located((By.ID, "Crm_Leads_COMPANY")))
    companyField.send_keys(company[i])

    jobField = wait.until(EC.visibility_of_element_located((By.ID, "Crm_Leads_DESIGNATION")))
    jobField.send_keys(jobTitle[i])

    emailField = wait.until(EC.visibility_of_element_located((By.ID, "Crm_Leads_EMAIL")))
    emailField.send_keys(email1[i])

    numbers = getPhoneAndMobile(telephone1[i], telephone2[i], telephone3[i])
    phoneField = wait.until(EC.visibility_of_element_located((By.ID, "Crm_Leads_PHONE")))
    mobileField = wait.until(EC.visibility_of_element_located((By.ID, "Crm_Leads_MOBILE")))

    phoneField.send_keys(numbers.get('phone'))
    mobileField.send_keys(numbers.get('mobile'))

    websiteField = wait.until(EC.visibility_of_element_located((By.ID, "Crm_Leads_WEBSITE")))
    webPageName = getWebsite(webPage[i])
    websiteField.send_keys(webPageName)

    faxField = wait.until(EC.visibility_of_element_located((By.ID, "Crm_Leads_FAX")))
    faxField.send_keys(getFax(fax1[i]))

    streetField = wait.until(EC.visibility_of_element_located((By.ID, "Crm_Leads_LANE")))
    cityField = wait.until(EC.visibility_of_element_located((By.ID, "Crm_Leads_CITY")))
    stateField = wait.until(EC.visibility_of_element_located((By.ID, "Crm_Leads_STATE")))
    postCodeField = wait.until(EC.visibility_of_element_located((By.ID, "Crm_Leads_CODE")))
    countryField = wait.until(EC.visibility_of_element_located((By.ID, "Crm_Leads_COUNTRY")))

    addressDict = getAddressDict(address1[i])

    streetField.send_keys(addressDict.get('street'))
    cityField.send_keys(addressDict.get('city'))
    stateField.send_keys(addressDict.get('state'))
    postCodeField.send_keys(addressDict.get('postCode'))
    countryField.send_keys(addressDict.get('country'))    


