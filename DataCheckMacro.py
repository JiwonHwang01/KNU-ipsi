from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import warnings
from time import sleep
import math

warnings.simplefilter(action='ignore', category=FutureWarning)

driver_path = input("ChromeDriver Path : ")
url = input("TestServer Url : ")
file = input("DataFile path : ")
id_row = input("수험번호 열의 열 이름 : ")
name_row = input("이름 열의 열 이름 : ")
birth_row = input("생년월일 열의 열 이름 : ")
habbul_row = input("판정여부 열의 열 이름 : ")

driver = webdriver.Chrome(executable_path = driver_path)
driver.implicitly_wait(10)
driver.get(url)
sleep(1)
write = pd.ExcelWriter(file, mode = 'a', engine='openpyxl', if_sheet_exists='overlay')

dataset = pd.read_excel(file)
Id = dataset[id_row]
Name = dataset[name_row]
Birth = dataset[birth_row]
habbul = dataset[birth_row]
tf_check = dataset['체크']

pass_check = [False] * data_num
data = []

i = 0
while True:
    if i == data_num: 
        break
    if not math.isnan(tf_check[i]):
        data.append(True)
        print(i)
        i += 1
        continue

    try:
        birth_data = str(Birth[i]).replace(".","")
        driver.find_element(By.CSS_SELECTOR, '#id_APPLYNO').send_keys(str(Id[i]))
        driver.find_element(By.CSS_SELECTOR, '#id_BIRTHDATE6').send_keys(birth_data)
        driver.find_element(By.CSS_SELECTOR, '#id_NAMEKOR').send_keys(str(Name[i]))
        driver.find_element(By.CSS_SELECTOR, '#wrapcontents > div.red > form > div > div > p > a').click()
        driver.implicitly_wait(30)
        driver.find_element(By.CSS_SELECTOR, '#wrapcontents > div.red > form:nth-child(3) > div.container > div:nth-child(1) > table > tbody > tr:nth-child(5) > td')
        driver.implicitly_wait(10)

        text2 = WebDriverWait(driver,30,2).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#wrapcontents > div.red > form:nth-child(3) > div.container > div:nth-child(1) > table > tbody > tr:nth-child(5) > td'))).text
        
        print(i)
        print('이름 :', Name[i])
        print('합불 :',habbul[i])
        print('확인 :', text2)
        print(habbul[i] == text2)

        if habbul[i] != '합격' and habbul[i] != '불합격':
            text2 = '후보' + text2[12:-1]

        data.append(habbul[i] == text2)
        driver.get(url)
        driver.implicitly_wait(20)
        sleep(1)

        i += 1
        
    except Exception as e:
        print(e)
        break

df = pd.DataFrame(data, columns = ['체크'])
df.to_excel(
    write,
    startcol=22,
    startrow=0,
    index=False
)
write.save()
write.close()