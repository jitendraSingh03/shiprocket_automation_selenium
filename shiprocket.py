from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import xlrd


# import pandas as pd
import time

# df = pd.read_excel('autobot.xls')
web = webdriver.Chrome(executable_path = 'C:\\Users\\hp\\Desktop\\python\\Automation\\chromedriver.exe')
loc = ("C:\\Users\\hp\\Desktop\\python\\Automation\\example1.xls")
web.get('https://app.shiprocket.in/addorder/?redirect_url=&customer_id=')
time.sleep(6)

email=web.find_element_by_name('emailid')
email.send_keys('ABC@gmail.com')	# Enter your email
password=web.find_element_by_name('password')
password.send_keys('abc@123')	#Enter your password
web.find_element_by_xpath('/html/body/app-root/app-login/div[2]/div[2]/div/div[1]/form/div[4]/div/button').click()
time.sleep(5)

def call(name,phone,price,email,address,pin,or_id):
    try:
        first_name_input = web.find_element_by_id('form_shipping_first_name')
        first_name_input.send_keys(name)
        time.sleep(2)
        phone_input = web.find_element_by_id('form_shipping_mobile')
        phone_input.send_keys(int(phone))
        time.sleep(2)
        email_input = web.find_element_by_id('form_shipping_email')
        email_input.send_keys(email)
        time.sleep(2)
        shipping_address = web.find_element_by_name('shipping_address')
        shipping_address.send_keys(address)
        time.sleep(2)
        
        shipping_pincode = web.find_element_by_name('shipping_pincode')
        shipping_pincode.send_keys(int(pin))
        time.sleep(5)
        # web.find_element_by_xpath('//*[@id="accordiongroup-14-3728-panel"]/div/div/div[5]/button').click()

        # -------------------------------------
        form_order_id = web.find_element_by_id('form_order_id')
        form_order_id.send_keys(or_id)
        

        productName = web.find_element_by_name('order_items.0.name')
        productName.send_keys('LifeCare24*7 care')
        
        form_enter_sku = web.find_element_by_id('form_enter_sku')
        form_enter_sku.send_keys('LC24')
        time.sleep(2)
        form_enter_qty = web.find_element_by_id('form_enter_qty')
        form_enter_qty.send_keys('1')
        time.sleep(2)

        form_enter_price = web.find_element_by_id('form_enter_price')
        form_enter_price.send_keys(int(price))
        
        time.sleep(10)
        # -------------------------------------------
        weight = web.find_element_by_name('weight')
        weight.send_keys(1)
        length = web.find_element_by_name('length')
        length.send_keys(15)
        breadth = web.find_element_by_name('breadth')
        breadth.send_keys(10)
        height = web.find_element_by_name('height')
        height.send_keys(6)
        time.sleep(5)
    except Exception:
        pass 

    #----------------------------------------------
  
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
print(sheet.nrows)
print(sheet.ncols)

for i in range(0,sheet.nrows):
    or_id=sheet.cell_value(i,0)
    name=sheet.cell_value(i,1)
    phone=sheet.cell_value(i,2)
    email=sheet.cell_value(i,3)
    address=sheet.cell_value(i,4)
    pin=sheet.cell_value(i,5)
    price=sheet.cell_value(i,6)
    web.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 't')
    web.get('https://app.shiprocket.in/addorder/?redirect_url=&customer_id=')
    call(name,phone,price,email,address,pin,or_id)

