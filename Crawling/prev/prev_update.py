from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

# https://anime.onnada.com/

chd = 'C:/dev_files/chd/chd.exe'
options = webdriver.ChromeOptions()
options.add_argument("headless")
options.add_argument("window-size=1920x1080")
options.add_argument("disable-gpu")
options.add_argument("User-Agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_4) AppleWebKit/537.36 (KHTML, "
                     "like Gecko) Chrome/93.0.4577.82")
options.add_argument("lang=ko_KR")
driver = webdriver.Chrome(chd, options=options)

driver.get("https://anime.onnada.com/")

find_date_xpath = "/html/body/div/div/div/div/ul/li/a"

find_date = driver.find_elements_by_xpath(find_date_xpath)

for a in find_date:
    print(a.text)
