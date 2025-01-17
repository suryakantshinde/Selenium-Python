from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.webdriver.common import alert
import xlrd
from xlrd import sheet
from xlrd.sheet import Sheet
import time
from selenium.common.exceptions import TimeoutException

browser = webdriver.Chrome(
    executable_path="C://001_Python_Development//PY_CPV_Group_Automation//Chrome_Browse//chromedriver.exe")
browser.implicitly_wait(60)
browser.maximize_window()
browser.get('https://www2.omnipaygroup.com/ramtool')    

#----------------------------------------------------------------------------------
# Below Code is to read User Name and Password from Notepad
#----------------------------------------------------------------------------------

filepath = "C://001_Python_Development//PY_CPV_Group_Automation//Login.txt"
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
WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@href = '#yui-gen22']"))) 
browser.find_element_by_xpath("//a[@href = '#yui-gen22']").click()
#Maintain Merchant Details
WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href*='MERCH_MAINTAIN_DETAILS']"))) 
browser.find_element_by_css_selector("a[href*='MERCH_MAINTAIN_DETAILS']").click()
#----------------------------------------------------------------------------------------------------------------------

# 1] Click on Enter Merchant Number [ Button ] - Element - ID [merchbutton-button]
time.sleep(0.2)
browser.find_element_by_id("merchbutton-button").click()

# 2] Enter Merchant Number in Text Box - Element - ID [id_40A]
browser.find_element_by_id("id_40A").click()
#----------------------------------------------------------------------------------------------------------------------        
#----------------------------------------------------------------------------------------------------------------------
# read the excel file 
workbook = xlrd.open_workbook("C://001_Python_Development//PY_CPV_Group_Automation//Test_Cases.xls")
sheet = workbook.sheet_by_name("InputData")

rowCount = sheet.nrows
colCount = sheet.ncols

for curr_row in range(1, rowCount):
    SearchField_MID_values = sheet.cell_value(curr_row, 0)
    print("Processing Case # " , SearchField_MID_values)
    browser.switch_to.default_content()
   
    # here i am joining element with excel cell
    SearchField = browser.find_element_by_id("id_40A")
    SearchField.send_keys(SearchField_MID_values)
    time.sleep(0.5)
    #----------------------------------------------------------------------- 

    # Click on Change Button
    element=browser.find_element_by_id("changeMerchBtn").click()
    print("click on change button")
    time.sleep(0.10)
    # Click on Property Tab
    element=browser.find_element_by_link_text("Properties").click()
    print("click on property tab")
    #click on Add and Update Properties Button
    element=browser.find_element_by_id("btnAdd").click()
    print("click on Add and Update Properties Button")
    # Click on Check Box CPV
    time.sleep(0.10)
    element=browser.find_element_by_name('SEL_009').click()
    time.sleep(0.5)
    #ID205aby_009
    SearchField_values = sheet.cell_value(curr_row, 1)
    print("Processing Case # " , SearchField_values)
    time.sleep(0.2)  
    # here i am joining element with excel cell
    #SearchField = browser.find_element_by_id("ID205aby_009")
    #SearchField.send_keys("Test123@gmal.com")
    EM = browser.find_element_by_id("ID205aby_009").get_attribute('value')
    if EM.strip() == "": 
        time.sleep(0.2)
        SearchField = browser.find_element_by_id("ID205aby_009")
        SearchField.send_keys(SearchField_values)
    else:    
        SearchField = browser.find_element_by_id("ID205aby_009")
        SearchField.send_keys(',',SearchField_values) # adding comma between two email ID's
        time.sleep(0.2)
    SC = "C://001_Python_Development//PY_CPV_Group_Automation//Screenshot" + SearchField_MID_values + '.png' # Saving image file with MID number
    browser.save_screenshot(SC)
    #Click on Add Button yui-gen32-button
    element=browser.find_element_by_id("yui-gen32-button").click()
    
    
    #-----------------------------------------------------------------------    
print("========== Completed ==========")
