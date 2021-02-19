from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.firefox.service import Service


def setBrowser():

    ############# Enter the browser
    browsername = input ("Please enter your browser [chrome ; firefox ; edge  ]")
   # browsername = "chrome"
    if browsername.lower() == "chrome":
        driver = webdriver.Chrome(ChromeDriverManager().install())
    elif browsername.lower() == "firefox":
        #driver = webdriver.firefox(executable_path=GeckoDriverManager().install())
        ff = Service("../resources/geckodriver.exe")
        driver = webdriver.Firefox(service=ff)
        driver.get("https://rpmsoftware.com/hiring/2020/integration-test/form-edit.html")
    elif browsername.lower() == "edge":
        driver = webdriver.Edge(EdgeChromiumDriverManager().install())
    else:
        print("Please pass the correct browser name : "+ browsername)
        raise Exception('Browser is not found')
    driver.implicitly_wait(5)
    return driver

