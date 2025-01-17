####### Project Developed to make -- Omnipay Reference tab changes -- for MS Operation ( Vendor Management) Zankar Patil / Gaurav Kamat ####################
####### Deveoped by Surya/Bala #######


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common import alert
import xlrd
from xlrd import sheet
from xlrd.sheet import Sheet
import time
from selenium.common.exceptions import TimeoutException
import os
import logging

logging.basicConfig()
logging.basicConfig(filename='app.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')
logging.warning('This will get logged to a file')


print("====Execution Started====")
working_dir = os.path.dirname(os.path.abspath(__file__))
browser = webdriver.Chrome("C:\\PY_Reff_Tab_Change\\MainProject\\chromedriver.exe")

    #executable_path= working_dir + "\\Chrome_Browse\\chromedriver.exe")
browser.implicitly_wait(20)
browser.maximize_window()
browser.get('https://www2.omnipaygroup.com/ramtool')   
#browser.get('https://ramu17.omnipaytest.com/ramtool')  

#----------------------------------------------------------------------------------
# Below Code is to read User Name and Password from Notepad
#----------------------------------------------------------------------------------

#filepath = working_dir + "C:\\PY_Reff_Tab_Change\\Login.txt"
#filepath = "./Login.txt"
filepath ="C:\\PY_Reff_Tab_Change\\MainProject\\Login.txt"
combo = open(filepath,"r")  
line = list(filepath)    
for line in combo :        
    temp = line.strip().split(":") or any
    email = temp[0]       
    pass1 = temp[1]   
    browser.find_element_by_id("69").send_keys(email)
    browser.find_element_by_id("76").send_keys(pass1)
# Click on Login Button
element=browser.find_element_by_xpath('/html/body/div/div[2]/form/div[2]/div/input[1]').click()

#----------------------------------------------------------------------------------------------------------------------
#Merchant Administration
WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@href = '#Merchant Administration']"))) 
browser.find_element_by_xpath("//a[@href = '#Merchant Administration']").click()
#Merchant Maintenance
#WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@href* = '#Merchant Maintenance']"))) 
#browser.find_element_by_xpath("//a[@href* = 'Merchant Maintenance']").click()
WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space() = 'Merchant Maintenance']"))) 
browser.find_element_by_xpath("//a[normalize-space() = 'Merchant Maintenance']").click()
#Maintain Merchant Details
WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='MERCH_MAINTAIN_DETAILS']"))) 
browser.find_element_by_css_selector("a[href*='MERCH_MAINTAIN_DETAILS']").click()
#----------------------------------------------------------------------------------------------------------------------

# read the excel file 
#workbook = xlrd.open_workbook(working_dir + "/Test_Cases.xls")

workbook = xlrd.open_workbook("C:\\PY_Reff_Tab_Change\\MainProject\\Test_Cases.xls")
sheet = workbook.sheet_by_name("InputData")

rowCount = sheet.nrows
colCount = sheet.ncols
print(("Total MID", rowCount-1))
for curr_row in range(1, rowCount):
    SearchField_MID_values = sheet.cell_value(curr_row, 0)
    SearchField_Ref_values = sheet.cell_value(curr_row, 1)
    print("Processing Case # " , SearchField_MID_values)

    # 1] Click on Enter Merchant Number [ Button ] - Element - ID [merchbutton-button]
    time.sleep(0.10)
    browser.find_element_by_id("merchbutton-button").click()

    # 2] Enter Merchant Number in Text Box - Element - ID [id_40A]
    browser.find_element_by_id("id_40A").click()
    time.sleep(0.10)
    browser.find_element_by_id("id_40A").clear()
    time.sleep(0.10)
    #----------------------------------------------------------------------------------------------------------------------        
    #----------------------------------------------------------------------------------------------------------------------

    # here i am joining element with excel cell
    SearchField = browser.find_element_by_id("id_40A")
    SearchField.send_keys(SearchField_MID_values)
    time.sleep(0.10)
    #----------------------------------------------------------------------- 

    # Click on Change Button
    element=browser.find_element_by_id("changeMerchBtn").click()#changeMerchBtn
    print("click on change button")
    time.sleep(0.10)
    # Click on References Tab
    element=browser.find_element_by_link_text("References").click()
    print("click on References tab")
    #============================================================================
    tbl=browser.find_element_by_xpath("//form/div/div[5]/div/div[3]/table/tbody[2]")
    for row in tbl.find_elements_by_xpath("./tr"):
        icon=row.find_elements_by_tag_name("td")[0]
        refftype=row.find_elements_by_tag_name("td")[1]
        reffvalue=row.find_elements_by_tag_name("td")[2]
        if (refftype.text == "MID Type"):
            icon.click()
            edit_element = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, "Edit")))
            edit_element.click()
            time.sleep(0.10)
           
            while True:
                time.sleep(0.10)
                ref_element = browser.find_element_by_id("ID48Bbb")
                print("Ref Value1:",ref_element.get_attribute('value').strip().upper())
                print("Ref Value2:",reffvalue.text.strip().upper())
                if ref_element.get_attribute('value').strip().upper() == reffvalue.text.strip().upper() or ref_element.get_attribute('value').strip().upper() != "":
                    break
            ref_element.clear()
            time.sleep(0.10)
            ref_element.send_keys(SearchField_Ref_values)
            time.sleep(0.10)
            browser.find_element_by_xpath("//button[@class ='update']").click()
            time.sleep(0.10)
            print("MID Type Updated for", SearchField_MID_values)
            break
  
    #-----------------------------------------------------------------------    
print("========== Completed ==========")
