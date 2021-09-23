
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

chd = 'C:/dev_files/chd/chd.exe'

# 헤들리스 설정
options = webdriver.ChromeOptions()
options.add_argument("headless")

# 웹페이지 사이즈 설정
options.add_argument("window-size=1920x1080")

# 그래픽 사용 안함
options.add_argument("disable-gpu")

# 인증 정보
# options.add_argument("User-Agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_4) AppleWebKit/537.36 (KHTML, "
#                      "like Gecko) Chrome/93.0.4577.82")

# 언어 설정 : 한국어
options.add_argument("lang=ko_KR")

driver = webdriver.Chrome(chd, options=options)

# 크롤링할 사이트 호출
# driver.get("https://www.python.org/")
