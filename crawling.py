# selenium의 webdriver를 사용하기 위한 import
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# selenium으로 키를 조작하기 위한 import
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.by import By


# 크롬드라이버 실행
driver = webdriver.Chrome() 

#크롬 드라이버에 url 주소 넣고 실행
driver.get('https://devye.tistory.com/104#google_vignette')

# 페이지가 완전히 로딩되도록 3초동안 기다림
time.sleep(5)

elements  = driver.find_elements(By.TAG_NAME,"p")
for element in elements:
    print(element.text)

