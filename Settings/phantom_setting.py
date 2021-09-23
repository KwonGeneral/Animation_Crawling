
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

# PhantomJS 드라이버 설정
pha = 'C:/dev_files/pha/bin/pha.exe'
driver = webdriver.PhantomJS(pha)

# 크롤링할 사이트 호출
# driver.get("https://www.python.org/")

