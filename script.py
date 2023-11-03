# -*- coding: utf-8 -*-
"""
Created on Fri Jan  7 14:22:09 2023

@author: fedig
"""
import os

import urllib3
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC, wait
import time
import pandas as pd
import json
pd.options.mode.chained_assignment = None  # default='warn'
import numpy as np
import string
import tkinter as tk
from pathlib import Path
import os


def progressBarLength(excelFile):
    members = pd.read_excel(excelFile)
    compt = 0
    try:
        for i in members["Status"]:
            if i != "Done":
                compt += 1
        return compt
    except KeyError:
        tk.messagebox.showerror('Error', "'Status' Column doesn't EXIST !!")


def getStatusIndex(excelFile):
    letters = list(string.ascii_uppercase)
    members = pd.read_excel(excelFile)
    compt = 0
    for i in members.columns:
        compt += 1

    return letters[compt - 1]


def addStatusColumn(excelFile):
    members = pd.read_excel(excelFile)
    for i in members.columns:
        if i.upper() == "STATUS":
            return 0
    members["Status"] = "Undone"
    members.to_excel(excelFile)

def addCartColumn(excelFile):
    members = pd.read_excel(excelFile)
    for i in members.columns:
        if i.upper() == "CART NUMBER":
            return 0
    members["Cart Number"] = ""
    members.to_excel(excelFile)

def addAmountColumn(excelFile):
    members = pd.read_excel(excelFile)
    for i in members.columns:
        if i.upper() == "AMOUNT(USD)":
            return 0
    members["Amount(USD)"] = ""
    members.to_excel(excelFile)




def splitExcel(excelFile, nbreOfExcelFiles):
    addStatusColumn(excelFile)
    members = pd.read_excel(excelFile)
    split = np.array_split(members, nbreOfExcelFiles)
    fileName = excelFile.split("/")[len(excelFile.split("/")) - 1]

    path = ""

    for i in range(len(excelFile.split("/")) - 2):
        path += excelFile.split("/")[i]

    for i in range(len(split)):
        split[i].to_excel(fileName[:len(fileName) - 4] + str(i) + '.xlsx')
        # Delete 1st column
        '''
        book = openpyxl.load_workbook(path+fileName[:len(fileName)-4]+str(i)+'.xlsx')
        sheet = book.active
        sheet.delete_cols(1)
        book.save(path+fileName[:len(fileName)-4]+str(i)+'.xlsx')
        '''


def checkMembership(membership, memberships, driver, url):
    if membership in memberships:
        driver.get(url)
        try:
            if membership == "PES":
                membershipSelectElement = driver.find_element(By.ID, "subscription-media-type")
                membershipSelect = Select(membershipSelectElement)
                membershipSelect.select_by_index(1)
            # addItemBtn = driver.find_element(By.ID, "addItems")
            addItemBtn = driver.find_element(By.XPATH, "//div[6]/div/div/div/input")
            addItemBtn.click()
        except Exception:
            print("error")
            pass


def saveInvoice(driver,email):
    driver.get("https://www.ieee.org/cart/publish/viewPaymentByMail.html")
    driver.execute_script('window.print();')
    time.sleep(2)
    downloads_path = str(Path.home() / "Downloads")
    new_download_directory = os.getcwd() + "\\invoices"
    try:
        os.mkdir(new_download_directory)
    except OSError as error:
        pass
        #print(error)

    try:
        os.rename(downloads_path + "\\viewPaymentByMail.html.pdf", new_download_directory + "\\"+email+".pdf")
    except FileExistsError:
        print("Overriding existing file ...")
        os.remove(new_download_directory + "\\"+email+".pdf")
        os.rename(downloads_path + "\\viewPaymentByMail.html.pdf", new_download_directory + "\\"+email+".pdf")







def proceedToPayment(driver, cardNumber, cardYear, cardMonth, cardOwner):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    continueBtnToCartPayment = driver.find_element(By.XPATH, "//div[6]/input")
    # time.sleep(10)
    continueBtnToCartPayment.click()
    # time.sleep(10)
    continueBtnToCartPayment1 = driver.find_element(By.ID, "proceed-to-checkout-btn-sec")

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    continueBtnToCartPayment1.click()

    creditCardRadio = driver.find_element(By.ID, "op-ccpay")
    creditCardRadio.click()

    billingAddressRadio = driver.find_element(By.ID, "sameas_mailing_address")
    billingAddressRadio.click()
    # time.sleep(5)

    creditCardTypeSelectElement = driver.find_element(By.ID, "choose-card-type")
    creditCardSelect = Select(creditCardTypeSelectElement)
    creditCardSelect.select_by_visible_text("MasterCard")

    creditCardNumber = driver.find_element(By.ID, "credit-card-number")
    creditCardNumber.send_keys(cardNumber)

    creditCardYearSelectElement = driver.find_element(By.ID, "expiration-year")
    creditCardYearSelect = Select(creditCardYearSelectElement)
    creditCardYearSelect.select_by_visible_text(cardYear)

    creditCardMonthSelectElement = driver.find_element(By.ID, "month")
    creditCardMonthSelect = Select(creditCardMonthSelectElement)
    creditCardMonthSelect.select_by_visible_text(cardMonth)

    nameOncard = driver.find_element(By.ID, "name-on-card")
    nameOncard.send_keys(cardOwner)
    time.sleep(7)
    verifyCardBtn = driver.find_element(By.XPATH, "//div[10]/div/input")
    verifyCardBtn.click()
    time.sleep(10)

    continuePaymentBtn = driver.find_element(By.XPATH, "//div[14]/div/input")
    continuePaymentBtn.click()
    termsConditionsCheckbox1 = driver.find_element(By.ID, "terms-conditions")
    termsConditionsCheckbox2 = driver.find_element(By.ID, "membership-terms-conditions")

    termsConditionsCheckbox1.click()
    termsConditionsCheckbox2.click()

    closeTermsBtn = driver.find_element(By.ID, "ibpModalClose")
    closeTermsBtn.click()

    ####finishPaymentBtn = driver.find_element(By.ID, "terms-conditions")


def setUpAccountNoPayment(email, password, memberships, memberName, memberID):
    chrome_options = webdriver.ChromeOptions()
    settings = {"recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}],
                "selectedDestinationId": "Save as PDF", "version": 2}
    prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings)}
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--kiosk-printing')
    # os.chdir("D:/IEEEMembershipAutomation/")
    service = Service(executable_path="/chromedriver")
    # driver = webdriver.Chrome(service=service)
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    driver.implicitly_wait(10)

    driver.execute_script('document.getElementsByTagName("html")[0].style.scrollBehavior = "auto"')

    driver.get("https://www.ieee.org/")


    # cookiesBtn = driver.find_element(By.LINK_TEXT, "Accept & Close")
    # cookiesBtn.click()

    signInLink = driver.find_element(By.LINK_TEXT, "Sign In")

    signInLink.click()

    emailTextbox = driver.find_element(by=By.NAME, value="pf.username")
    passwordTextbox = driver.find_element(by=By.NAME, value="pf.pass")
    signInBtn = driver.find_element(By.ID, "modalWindowRegisterSignInBtn")

    emailTextbox.send_keys(email)
    passwordTextbox.send_keys(password)
    signInBtn.click()

    driver.get("https://www.ieee.org/membership-application/join.html?grade=Student")

    driver.find_element(By.NAME, "customer.addresses[0].line1").clear()
    driver.find_element(By.NAME, "customer.addresses[0].city").clear()
    driver.find_element(By.NAME, "customer.addresses[0].postalCode").clear()

    '''
    countrySelectElement = driver.find_element(By.ID, "country")
    countrySelect = Select(countrySelectElement)
    countrySelect.select_by_visible_text('Tunisia')
    '''
    addressTextbox = driver.find_element(By.NAME, "customer.addresses[0].line1")
    cityTextbox = driver.find_element(By.NAME, "customer.addresses[0].city")
    zipTextbox = driver.find_element(By.NAME, "customer.addresses[0].postalCode")
    continueBtn = driver.find_element(by=By.XPATH, value="//div/div[6]/input")
    addressType = driver.find_element(By.XPATH, "//div[3]/input")
    #addressType = driver.find_element(By.XPATH, "//div[4]/input")
    provinceSelectElement = driver.find_element(By.ID, "province")
    provinceSelect = Select(provinceSelectElement)

    cartContent = driver.find_element(By.CLASS_NAME, "mc-section").text
    # print(cartContent)
    cartNumber = driver.find_element(By.XPATH, "//div[3]/div/div/div/div/div/div/div[5]").text.split("cart number")[1]
    print(cartNumber)



    if addressType.is_selected():
        pass
    else:
        addressType.click()

    addressTextbox.send_keys("ESPRIT, Pole Technologique - El Ghazala")
    cityTextbox.send_keys("El Ghazala")
    zipTextbox.send_keys("2083")
    provinceSelect.select_by_visible_text('Ariana')
    try:
        continueBtn.click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "addressHeader")))
        #print("Address Text found on the webpage.")
        continueBtn.click()
    except:
        #print("Address Text not found on the webpage.")
        pass


    # enajem nsala7ha b try catch student w professional ken catcha error ibadel
    try:
        undergradRadio = driver.find_element(By.ID, "studentStatusUndergraduate Student")
    except Exception:
        time.sleep(7)
        studentSelectElement = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "careerPhase")))
        studentSelect = Select(studentSelectElement)
        studentSelect.select_by_index(1)
        time.sleep(5)
        undergradRadio = driver.find_element(By.ID, "studentStatusUndergraduate Student")

    directoryCheckbox = driver.find_element(By.ID, "memberdir-options1")
    whyJoin1Checkbox = driver.find_element(By.ID, "TechnicallyCurrent")
    whyJoin2Checkbox = driver.find_element(By.ID, "CareerOpurtunities")
    whyJoin3Checkbox = driver.find_element(By.ID, "ExpandProfessionalNetwork")
    whyJoin4Checkbox = driver.find_element(By.ID, "ConnectToLocalActivities")
    #continue2Btn = driver.find_element(By.XPATH, value="//div[6]/div/div[2]/div/input")
    continue2Btn = driver.find_element(By.XPATH, value="//div/div[6]/div/div[2]/div/input")

    stateSelectElement = driver.find_element(By.ID, "state")
    stateSelect = Select(stateSelectElement)
    stateSelect.select_by_visible_text('Ariana')

    academicSelectElement = driver.find_element(By.ID, "stud-academic-program")
    academicSelect = Select(academicSelectElement)
    academicSelect.select_by_visible_text("Computer Science")

    gradMSelectElement = driver.find_element(By.ID, "estimated-grad-month")
    gradMSelect = Select(gradMSelectElement)
    gradMSelect.select_by_visible_text('June')

    gradYSelectElement = driver.find_element(By.ID, "estimated-grad-year")
    gradYSelect = Select(gradYSelectElement)
    gradYSelect.select_by_visible_text('2028')

    studySelectElement = driver.find_element(By.ID, "stud-current-study")
    studySelect = Select(studySelectElement)
    studySelect.select_by_visible_text("Computer Sciences and Information Technologies")

    techSelectElement = driver.find_element(By.ID, "stud-technical-focus")
    techSelect = Select(techSelectElement)
    techSelect.select_by_visible_text("Computing and Processing (Hardware/Software)")

    uniSearchBtn = driver.find_element(By.ID, "searchUniversity")
    uniSearchBtn.click()

    uniFilterTextbox = driver.find_element(By.ID, "universityFilter")
    uniFilterTextbox.send_keys("esprit")

    #espritLinktext = driver.find_element(By.LINK_TEXT,"Private Higher School of Engineering and Technology (ESPRIT) (ESPRIT)")
    espritLinktext = driver.find_element(By.LINK_TEXT,"ESPRIT")
    espritLinktext.click()
    undergradRadio.click()
    directoryCheckbox.click()

    degreeSelectElement = driver.find_element(By.ID, "stud-degree-pursued")
    degreeSelect = Select(degreeSelectElement)
    degreeSelect.select_by_visible_text("Engineer")

    if whyJoin1Checkbox.is_selected():
        pass
    else:
        whyJoin1Checkbox.click()

    if whyJoin2Checkbox.is_selected():
        pass
    else:
        whyJoin2Checkbox.click()

    if whyJoin3Checkbox.is_selected():
        pass
    else:
        whyJoin3Checkbox.click()

    if whyJoin4Checkbox.is_selected():
        pass
    else:
        whyJoin4Checkbox.click()

    accSelectElement = driver.find_element(By.ID, "stud-prog-accredited")
    accSelect = Select(accSelectElement)
    accSelect.select_by_value("Yes")

    driver.find_element(By.ID, "referring-mem-name").clear()
    driver.find_element(By.ID, "referring-mem-number").clear()
    referringName = driver.find_element(By.ID, "referring-mem-name")
    referringID = driver.find_element(By.ID, "referring-mem-number")

    referringName.send_keys(memberName)
    referringID.send_keys(memberID)

    # 96246048

    try:
        continue2Btn.click()
    except:
        pass

    time.sleep(1)

    try:
        cookiesBtn = driver.find_element(By.LINK_TEXT, "Accept & Close")
        cookiesBtn.click()
    except:
        pass

    try:
        if "Computer" not in cartContent:
            checkMembership("CS", memberships, driver,
                            "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMC016&searchResults=Y")
        if "Computational" not in cartContent:
            checkMembership("CIS", memberships, driver,
                            "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMCIS011&searchResults=Y")
        if "Power" not in cartContent:
            checkMembership("PES", memberships, driver,
                            "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMPE031&searchResults=Y")
        if "Industry" not in cartContent:
            checkMembership("IAS", memberships, driver,
                            "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMIA034&searchResults=Y")
        if "Microwave" not in cartContent:
            checkMembership("MTTS", memberships, driver,
                            "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMMTT017&searchResults=Y")
        if "Industrial" not in cartContent:
            checkMembership("IES", memberships, driver,
                            "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMIE013&searchResults=Y")
        if "Robotics" not in cartContent:
            checkMembership("RAS", memberships, driver,
                            "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMRA024&searchResults=Y")
        if "SIGHT" not in cartContent:
            checkMembership("SIGHT", memberships, driver,
                            "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMSIGHT&searchResults=Y")
        if "Women" not in cartContent:
            checkMembership("WIE", memberships, driver,
                            "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMWIE050&searchResults=Y")
        if "Aerospace" not in cartContent:
            checkMembership("AESS", memberships, driver,
                            "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMAES010&searchResults=Y")
    except Exception as e:
        print(e)
        print("Exception Error !")

    time.sleep(2)


    if(driver.current_url != "https://www.ieee.org/membership-catalog/index.html?N=0"):
        driver.get("https://www.ieee.org/membership-catalog/index.html?N=0")
        amount = driver.find_element(By.CLASS_NAME, "mc-checkout").text.split("*")[1]
    else:
        amount = driver.find_element(By.CLASS_NAME, "mc-checkout").text.split("*")[1]
    #print(amount)
    saveInvoice(driver, email)

    return cartNumber, amount
    driver.quit()





def setUpAccountWithPayment(email, password, memberships, memberName, memberID, cardNumber, cardYear, cardMonth,
                            cardOwner):
    # os.chdir("D:/IEEEMembershipAutomation/")
    service = Service(executable_path="/chromedriver")
    # driver = webdriver.Chrome(service=service)
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.implicitly_wait(10)
    driver.execute_script('document.getElementsByTagName("html")[0].style.scrollBehavior = "auto"')

    driver.get("https://www.ieee.org/")

    cookiesBtn = driver.find_element(By.LINK_TEXT, "Accept & Close")
    signInLink = driver.find_element(By.LINK_TEXT, "Sign In")
    cookiesBtn.click()
    signInLink.click()

    emailTextbox = driver.find_element(by=By.NAME, value="pf.username")
    passwordTextbox = driver.find_element(by=By.NAME, value="pf.pass")
    signInBtn = driver.find_element(By.ID, "modalWindowRegisterSignInBtn")

    emailTextbox.send_keys(email)
    passwordTextbox.send_keys(password)
    signInBtn.click()

    driver.get("https://www.ieee.org/membership-application/join.html?grade=Student")

    driver.find_element(By.NAME, "customer.addresses[0].line1").clear()
    driver.find_element(By.NAME, "customer.addresses[0].city").clear()
    driver.find_element(By.NAME, "customer.addresses[0].postalCode").clear()

    '''
    countrySelectElement = driver.find_element(By.ID, "country")
    countrySelect = Select(countrySelectElement)
    countrySelect.select_by_visible_text('Tunisia')
    '''
    addressTextbox = driver.find_element(By.NAME, "customer.addresses[0].line1")
    cityTextbox = driver.find_element(By.NAME, "customer.addresses[0].city")
    zipTextbox = driver.find_element(By.NAME, "customer.addresses[0].postalCode")
    continueBtn = driver.find_element(by=By.XPATH, value="//div/div[6]/input")
    addressType = driver.find_element(By.XPATH, "//div[3]/input")
    provinceSelectElement = driver.find_element(By.ID, "province")
    provinceSelect = Select(provinceSelectElement)

    cartContent = driver.find_element(By.CLASS_NAME, "mc-section").text
    # print(cartContent)
    if addressType.is_selected():
        pass
    else:
        addressType.click()

    addressTextbox.send_keys("ESPRIT, Pole Technologique - El Ghazala")
    cityTextbox.send_keys("El Ghazala")
    zipTextbox.send_keys("2083")
    provinceSelect.select_by_visible_text('Ariana')

    continueBtn.click()

    # enajem nsala7ha b try catch student w professional ken catcha error ibadel
    try:
        undergradRadio = driver.find_element(By.ID, "studentStatusUndergraduate Student")
    except Exception:
        time.sleep(7)
        studentSelectElement = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "careerPhase")))
        studentSelect = Select(studentSelectElement)
        studentSelect.select_by_index(1)
        time.sleep(5)
        undergradRadio = driver.find_element(By.ID, "studentStatusUndergraduate Student")

    directoryCheckbox = driver.find_element(By.ID, "memberdir-options1")
    whyJoin1Checkbox = driver.find_element(By.ID, "TechnicallyCurrent")
    whyJoin2Checkbox = driver.find_element(By.ID, "CareerOpurtunities")
    whyJoin3Checkbox = driver.find_element(By.ID, "ExpandProfessionalNetwork")
    whyJoin4Checkbox = driver.find_element(By.ID, "ConnectToLocalActivities")
    continue2Btn = driver.find_element(By.XPATH, "//div[5]/div/div[2]/div/input")

    stateSelectElement = driver.find_element(By.ID, "state")
    stateSelect = Select(stateSelectElement)
    stateSelect.select_by_visible_text('Ariana')

    academicSelectElement = driver.find_element(By.ID, "stud-academic-program")
    academicSelect = Select(academicSelectElement)
    academicSelect.select_by_visible_text("Computer Science")

    gradMSelectElement = driver.find_element(By.ID, "estimated-grad-month")
    gradMSelect = Select(gradMSelectElement)
    gradMSelect.select_by_visible_text('June')

    gradYSelectElement = driver.find_element(By.ID, "estimated-grad-year")
    gradYSelect = Select(gradYSelectElement)
    gradYSelect.select_by_visible_text('2028')

    studySelectElement = driver.find_element(By.ID, "stud-current-study")
    studySelect = Select(studySelectElement)
    studySelect.select_by_visible_text("Computer Sciences and Information Technologies")

    techSelectElement = driver.find_element(By.ID, "stud-technical-focus")
    techSelect = Select(techSelectElement)
    techSelect.select_by_visible_text("Computing and Processing (Hardware/Software)")

    uniSearchBtn = driver.find_element(By.ID, "searchUniversity")
    uniSearchBtn.click()

    uniFilterTextbox = driver.find_element(By.ID, "universityFilter")
    uniFilterTextbox.send_keys("esprit")

    espritLinktext = driver.find_element(By.LINK_TEXT,
                                         "Private Higher School of Engineering and Technology (ESPRIT) (ESPRIT)")
    espritLinktext.click()
    undergradRadio.click()
    directoryCheckbox.click()

    degreeSelectElement = driver.find_element(By.ID, "stud-degree-pursued")
    degreeSelect = Select(degreeSelectElement)
    degreeSelect.select_by_visible_text("Engineer")

    if whyJoin1Checkbox.is_selected():
        pass
    else:
        whyJoin1Checkbox.click()

    if whyJoin2Checkbox.is_selected():
        pass
    else:
        whyJoin2Checkbox.click()

    if whyJoin3Checkbox.is_selected():
        pass
    else:
        whyJoin3Checkbox.click()

    if whyJoin4Checkbox.is_selected():
        pass
    else:
        whyJoin4Checkbox.click()

    accSelectElement = driver.find_element(By.ID, "stud-prog-accredited")
    accSelect = Select(accSelectElement)
    accSelect.select_by_value("Yes")

    driver.find_element(By.ID, "referring-mem-name").clear()
    driver.find_element(By.ID, "referring-mem-number").clear()
    referringName = driver.find_element(By.ID, "referring-mem-name")
    referringID = driver.find_element(By.ID, "referring-mem-number")

    referringName.send_keys(memberName)
    referringID.send_keys(memberID)

    # 96246048

    continue2Btn.click()
    time.sleep(1)

    if "Computer" not in cartContent:
        checkMembership("CS", memberships, driver,
                        "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMC016&searchResults=Y")
    if "Computational" not in cartContent:
        checkMembership("CIS", memberships, driver,
                        "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMCIS011&searchResults=Y")
    if "Power" not in cartContent:
        checkMembership("PES", memberships, driver,
                        "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMPE031&searchResults=Y")
    if "Industry" not in cartContent:
        checkMembership("IAS", memberships, driver,
                        "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMIA034&searchResults=Y")
    if "Microwave" not in cartContent:
        checkMembership("MTTS", memberships, driver,
                        "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMMTT017&searchResults=Y")
    if "Industrial" not in cartContent:
        checkMembership("IES", memberships, driver,
                        "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMIE013&searchResults=Y")
    if "Robotics" not in cartContent:
        checkMembership("RAS", memberships, driver,
                        "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMRA024&searchResults=Y")
    if "SIGHT" not in cartContent:
        checkMembership("SIGHT", memberships, driver,
                        "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMSIGHT&searchResults=Y")
    if "Women" not in cartContent:
        checkMembership("WIE", memberships, driver,
                        "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMWIE050&searchResults=Y")
    if "Aerospace" not in cartContent:
        checkMembership("AESS", memberships, driver,
                        "https://www.ieee.org/membership-catalog/productdetail/showProductDetailPage.html?product=MEMAES010&searchResults=Y")

    time.sleep(2)
    '''
    driver.get("https://www.ieee.org/membership-catalog/index.html?N=0")
    time.sleep(2)
    '''
    # proceed to payment
    proceedToPayment(driver, cardNumber, cardYear, cardMonth, cardOwner)

    driver.quit()


def mainNoPayment(excelFile, memberName, memberID):
    # nbrOfUndoneAccounts = progressBarLength(excelFile)
    addStatusColumn(excelFile)
    # colLetter=getStatusIndex(excelFile)
    addCartColumn(excelFile)
    addAmountColumn(excelFile)
    members = pd.read_excel(excelFile)


    for ind in members.index:
        try:
            if members["Status"][ind] != "Done":
                cartNumber, amount = setUpAccountNoPayment(members["EmailAddress"][ind], members["Password"][ind],
                                      members["Memberships"][ind], memberName, memberID)
                members["Cart Number"][ind] = cartNumber
                members["Amount(USD)"][ind] = amount
                members["Status"][ind] = "Done"
                members.to_excel(excelFile)

            else:
                pass
        except urllib3.exceptions.NewConnectionError as e:
            print(f"Connection error: {e}")
        except Exception as e:
            print(f"An error occurred: {e}")
            members["Status"][ind] = "Undone"
            members.to_excel(excelFile)

    tk.messagebox.showinfo("Done", "- Done !")
    print("FINISHED !")


def mainWithPayment(excelFile, memberName, memberID, cardNumber, cardYear, cardMonth, cardOwner):
    addStatusColumn(excelFile)
    # colLetter=getStatusIndex(excelFile)
    members = pd.read_excel(excelFile)

    # setUpAccount(members["EmailAddress"][0],members["Password"][0],members["Memberships"][0],memberName,memberID,cardNumber,cardYear,cardMonth,cardOwner)

    '''
    for ind in members.index:
        if members["Status"][ind]!="Done":
            setUpAccountWithPayment(members["EmailAddress"][ind],members["Password"][ind],members["Memberships"][ind],memberName,memberID,cardNumber,cardYear,cardMonth,cardOwner)
    '''

    for ind in members.index:
        try:
            if members["Status"][ind] != "Done":
                setUpAccountWithPayment(members["EmailAddress"][ind], members["Password"][ind],
                                        members["Memberships"][ind], memberName, memberID, cardNumber, cardYear,
                                        cardMonth, cardOwner)
                members["Status"][ind] = "Done"
                members.to_excel(excelFile)
            else:
                pass
        except Exception:
            members["Status"][ind] = "Undone"
            members.to_excel(excelFile)

    tk.messagebox.showinfo("Done", "- Done !")
    print("FINISHED !")

# mainWithPayment("[Automation] IEEE Membership (Responses).xlsx","memberName","96246048","5359563923720646","2030","June","Salem")
