from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from time import sleep
import xlwt
import xlrd

loc = ("C:\\Users\\hp\\Desktop\\python\\Automation\\example.xls")
driver = webdriver.Chrome()
driver.maximize_window()
driver.get("https://web.whatsapp.com/")
sleep(10)

def call(mobile):
    try:
        # driver.get("https://web.whatsapp.com/send?phone=919414457521&text=Automation bangya%2C%20World!")
        s= "https://web.whatsapp.com/send?phone=91"+str(mobile)
        driver.get(s)     
        sleep(3)
        driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[1]/div/div[2]').send_keys(" ***NEW YEAR OFFER*** from jeet sikarwar")
        act=ActionChains(driver)
        act.send_keys(Keys.ENTER).perform()
    except Exception:
        pass
    
wb = xlrd.open_workbook(loc)

sheet = wb.sheet_by_index(0)
print(sheet.nrows)
print(sheet.ncols)

for i in range(0,sheet.nrows):
    name=sheet.cell_value(i,0)
    n=int(name)
    call(n)
