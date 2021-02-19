from selenium import webdriver
from selenium.webdriver.common.by import By
from EmployeeRegistrationPage.verifyEmployeeDetail import assertMethod
from EmployeeRegistrationPage.verifyEmployeeDetail import exceptionHandler
from EmployeeRegistrationPage.readexcel import readexcel
from EmployeeRegistrationPage.setBrowser import setBrowser

############## Setting up Browser
driver = setBrowser()

################# Opening a URL
driver.get("https://rpmsoftware.com/hiring/2020/integration-test/form-edit.html")

############# setting up employee Values
#setEmployeeDetails(driver)
readexcel(driver)
driver.implicitly_wait(5)
#############verify Submit button
submitBtn=driver.find_element(By.XPATH,"//*[@id='FormEditPanel']/div[18]/button")
userurl= "https://rpmsoftware.com/hiring/2020/integration-test/form.html#"
if submitBtn.is_enabled():
    driver.get("https://rpmsoftware.com/hiring/2020/integration-test/form.html#")

###################### Checking the url
try:
    assert userurl == driver.current_url
except:
    exceptionHandler("Url Error", "Url is not correct", "Please provide correct url")


###################### Asserting Page header
empName = "Isabel Britt"
try:
    header= driver.find_element(By.XPATH, "/html/body/div/h1").text
    assert header  == empName
    print(header)

except:
    exceptionHandler("Employee Error", "Details belong to other Employee", "Please provide details for Employee - Isabel Britt")

############# asserting Employee Values

assertMethod(driver, userurl,empName)


############## Checking Map
#driver.implicitly_wait(30)
driver.find_element(By.XPATH,"//*[@id='Field.500_25:ValueContainer']/a").click()
window_after = driver.window_handles[1]
driver.switch_to.window(window_after)
titlename = driver.title
print(titlename)
try:
    assert titlename.endswith("Google Maps")
except:
    exceptionHandler("Google Map Error", "unable to open Map",
                     "Clicking Map , should navigate to Google Maps  ")

driver.implicitly_wait(6)


############### Close Browser
driver.quit()

