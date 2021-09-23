
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

# 크롬드라이버 설정
chd = 'C:/dev_files/chd/chd.exe'

# driver = webdriver.Chrome(chd)

options = webdriver.ChromeOptions()
# options.add_argument("window-size=1920x1080")
# options.add_argument("disable-gpu")
# options.add_argument("User-Agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.82")

driver = webdriver.Chrome(chd, options=options)
# 크롤링할 사이트 호출
# driver.get("https://kwonputer.com")
# driver.get("https://www.python.org/")

# 셀레니움은 웹테스트를 위한 프레임워크로 다음과 같은 방식으로 웹테스트를 자동으로 진행함
# 아래의 내용이 없으면 프로그램 종료
# assert "찾을 내용" in driver.title

