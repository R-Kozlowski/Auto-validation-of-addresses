import os, bs4, time, xlrd, wx
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
      
#waiting modules
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
      
###prepared waiting module
##wait = WebDriverWait(browser, 10)
##element = wait.until(EC.element_to_be_clickable((By.ID, 'navigate')))
     
#changing the place where the program can find excel file with shipment numbers
#os.chdir("D:\\PROGRAMOWANIE\\PYTHON\\automat do walidacji RABEN")
      
first_cancel = 0
second_cancel = 0
 
 
#open the browser
url = 'company webside'
browser = webdriver.Chrome()
browser.get(url)
      
shipments = pd.read_excel('shipments.xlsx', 'Sheet0', usecols=['File number'])
lista = shipments['File number'].tolist()
      
#login window
app = wx.App()
#first: login
dlg = wx.TextEntryDialog(None, "Login", caption="Logowanie")
if dlg.ShowModal() == wx.ID_OK:
    user = dlg.GetValue()
#second: password
dlg = wx.TextEntryDialog(None, "Password", caption="Logowanie")
if dlg.ShowModal() == wx.ID_OK:
    password = dlg.GetValue()
 
 
login = browser.find_element_by_xpath('//*[@id="username"]')
login.send_keys(str(user))
haslo = browser.find_element_by_xpath('//*[@id="main"]/div/span[1]/form/p/input')
haslo.send_keys(str(password))
haslo.submit()
time.sleep(10)
     
#open shipments overview
orders_management = browser.find_element_by_xpath('//*[@id="navigate"]/div/div[2]/div[1]/a/table/tbody/tr/td[3]/span')
orders_management.click()
time.sleep(2)
shipments_overwiew = browser.find_element_by_xpath('//*[@id="navigate"]/div/div[2]/div[2]/div/div/div[1]/table/tbody/tr/td[2]/span')
#shipments_overwiew = browser.find_element_by_link_text('Shipment overview')
shipments_overwiew.click()
time.sleep(3)
     
pole_przesylki = browser.find_element_by_xpath('//*[@id="ic_maintab"]/div[2]/div/div/div/div[2]/div/div/div[1]/div/div/table/tbody/tr/td/div/div[22]/div/table/tbody/tr/td[2]/input')
      
def street_changer(street_city):
    if street_city != None:
        street_city = street_city.split('-')
        street_city = ''.join(street_city)
        street_city = street_city.split(' ')
        street_city = ''.join(street_city)  
        street_city = str(street_city).lower()
        street_city = street_city.replace('straße', 'street')
        street_city = street_city.replace('str.', 'street')
        street_city = street_city.replace('strasse', 'street')
        street_city = street_city.replace('strase', 'street')
        street_city = street_city.replace('ü', 'u')
        street_city = street_city.replace('ä', 'a')
        street_city = street_city.replace('ö', 'o')
        street_city = street_city.replace('ß', 's')
             
        return street_city
    else:
        street_city = ""
        return street_city
      
def address_veryfication(confirmation_button):
          
    left_street = browser.find_element_by_xpath('//*[@id="data2"]/div[12]/div[10]/div/div/div[1]/div/div/div/div[1]/div/div/table/tbody/tr/td/div/div[6]/div/table/tbody/tr/td[2]/input')
    left_street = left_street.get_attribute('value')
    print('---I located the street: ' + str(left_street))
    left_street = street_changer(left_street)
    print('---I changed this street to: ' + left_street)
          
    right_street = browser.find_element_by_xpath('//*[@id="data2"]/div[12]/div[10]/div/div/div[3]/div/div/div/div/div/div[1]/div/table/tbody/tr[1]/td[4]/div')
    right_street = right_street.text
    print('---I located the street to confirm: ' + str(right_street))
    right_street = street_changer(right_street)
    print('---I changed this street to: ' + right_street)
          
    left_postal_code = browser.find_element_by_xpath('//*[@id="data2"]/div[12]/div[10]/div/div/div[1]/div/div/div/div[1]/div/div/table/tbody/tr/td/div/div[9]/div/table/tbody/tr/td[2]/input')
    left_postal_code = left_postal_code.get_attribute('value')
    print('---I located the postal code: ' + str(left_postal_code))
     
    right_postal_code = browser.find_element_by_xpath('//*[@id="data2"]/div[12]/div[10]/div/div/div[3]/div/div/div/div/div/div[1]/div/table/tbody/tr[1]/td[6]/div')
    right_postal_code = right_postal_code.text
    print('---I located the postal code to confirm: ' + str(right_postal_code))
      
    left_city = browser.find_element_by_xpath('//*[@id="data2"]/div[12]/div[10]/div/div/div[1]/div/div/div/div[1]/div/div/table/tbody/tr/td/div/div[10]/div/table/tbody/tr/td[2]/input')
    left_city = left_city.get_attribute('value')
    print('---I located the city: ' + str(left_city))
    left_city = street_changer(left_city)
    print('---I changed this city to: ' + left_city)
          
    right_city = browser.find_element_by_xpath('//*[@id="data2"]/div[12]/div[10]/div/div/div[3]/div/div/div/div/div/div[1]/div/table/tbody/tr[1]/td[7]/div')
    right_city = right_city.text
    print('---I located the city to confirm: ' + str(right_city))
    right_city = street_changer(right_city)
    print('---I changed this city to: ' + right_city)
      
    #porównanie ulic
    if left_street==right_street and left_postal_code==right_postal_code and left_city==right_city:
        mark_line = browser.find_element_by_xpath('//*[@id="data2"]/div[12]/div[10]/div/div/div[3]/div/div/div/div/div/div[1]/div/table/tbody/tr[1]/td[7]/div')
        mark_line.click()
        time.sleep(5)
        confirmation_button.click()
        print('I confirmed window')
        ##confirmation_button.submit()
    else: #zamykam okno bez potwierdzenia jeśli nie zgadzają się elementy
        try:
            close_window = browser.find_element_by_xpath('//*[@id="data2"]/div[12]/div[1]/div[2]/div[1]')
            close_window.click()
            print('Data are not the same. I closed the window')
        except:
            left_postal_code.send_keys(Keys.ESCAPE)
            print('Data are not the same. I closed window')     
      
for numer_przesyłki in lista:
     pole_przesylki.send_keys(numer_przesylki)
     pole_przesylki.send_keys(Keys.ENTER)
     time.sleep(5)
     pierwsza_przesylka = browser.find_element_by_xpath('//*[@id="ic_maintab"]/div[2]/div/div/div/div[2]/div/div/div[3]/div/div/table/tbody/tr[1]/td[1]/div')
     actionsChains = ActionChains(browser)
     actionsChains.context_click(pierwsza_przesylka).perform()
     time.sleep(2)
     pierwsza_pozycja_menu = browser.find_element_by_xpath('//*[@id="data2"]/div[11]/div/div/div[1]/a/table/tbody/tr/td[2]')
     pierwsza_pozycja_menu.click()
     time.sleep(5)
     
     #first message
     try:
         #browser.find_element_by_xpath('//*[@value="Cancel"]')==True:
         first_ok = browser.find_element_by_xpath('//*[@id="data2"]/div[12]/div[10]/div/div/div/div[2]/div/input[2]').click()
         #first_ok.submit()   
         #first_ok.click()
         print('I refused the first confirmation window. Reason: "No data found"')
         time.sleep(10)
         first_cancel = 1
     except:
         print("I couldn't located the first message")
         time.sleep(10)
     
     #first confirmation window
     if first_cancel == 0:
         try:
             #browser.find_element_by_xpath('//*[@value="Confirm"]')==True:
             confirm = browser.find_element_by_xpath('//*[@value="Confirm"]')
             print('I located the first confirmation window')
             address_veryfication(confirm)
             time.sleep(10)
         except:
             print("I couldn't located the first confirmation")
             time.sleep(10)
      
     #second message
     try:
         #browser.find_element_by_xpath('//*[@value="Canlcel"]')==True:
         second_ok = browser.find_element_by_xpath('//*[@id="data2"]/div[12]/div[10]/div/div/div/div[2]/div/input[2]').click()
         #second_ok.submit()   
         #second_ok.click()
         print('I refused the second confirmation window. Reason: "No data found"')
         second_cancel = 1
         time.sleep(10)
 
     except:
         print("I couldn't located the second message")
         time.sleep(10)
      
     #second confirmation window
     if second_cancel == 0:
         try:
             #browser.find_element_by_xpath('//*[@value="Confirm"]')==True:
             confirm = browser.find_element_by_xpath('//*[@value="Confirm"]')
             print('I located the second confirmation window')
             address_veryfication(confirm)
             time.sleep(10)
    except:
        print("I couldn't located the second confirmation")
        time.sleep(10)
          
     pole_przesylki = browser.find_element_by_xpath('//*[@id="ic_maintab"]/div[2]/div/div/div/div[2]/div/div/div[1]/div/div/table/tbody/tr/td/div/div[22]/div/table/tbody/tr/td[2]/input')
       
     #after all, I can remove last shipment and start with next one
     pole_przesylki.clear()
     first_cancel = 0
     second_cancel = 0

browser.quit()
