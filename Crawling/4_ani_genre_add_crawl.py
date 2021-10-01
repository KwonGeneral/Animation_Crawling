
import time

from openpyxl import load_workbook
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException, \
    InvalidArgumentException, WebDriverException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver

chd = 'C:/dev_files/chd/chd.exe'
options = webdriver.ChromeOptions()
options.add_argument("headless")
options.add_argument("window-size=1920x1080")
options.add_argument("disable-gpu")
options.add_argument("User-Agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_4) AppleWebKit/537.36 (KHTML, "
                     "like Gecko) Chrome/93.0.4577.82")
driver = webdriver.Chrome(chd, options=options)
driver.get("https://www.chuing.net/db/search.php?cdbsearch=asdfasdf")

search_xpath = "//*[@id='SjestForm']/div/div[1]/input"
result_a_xpath = "/html/body/div[5]/div[2]/div/div[2]/div/div[4]/div[1]"

work_book = load_workbook("ani_detail.xlsx")
sheet = work_book['Sheet']

for no in range(2, len(sheet["A"]) + 1):
    try:
        search_tag = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, search_xpath)))
        search_tag.clear()

        search_tag.send_keys(sheet["A" + str(no)].value)
        search_tag.send_keys(Keys.ENTER)

        result_a = driver.find_elements_by_xpath(result_a_xpath)
        result_total = []
        sheet_b = sheet["B" + str(no)].value

        if len(result_a) > 0:
            for a in result_a:
                for kkk in str(a.text).split("장르: ")[1].split(","):
                    result_total.append(kkk)
        if len(result_total) > 0:
            if sheet_b is not None and sheet_b != "" and sheet_b != "None":
                sheet["B" + str(no)].value = sheet_b + ", " + ", ".join(result_total)
            else:
                sheet["B" + str(no)].value = ", ".join(result_total)

        print("[ ", str(no), " ] ", sheet["A" + str(no)].value, "[ ", sheet["B" + str(no)].value, " ]")
        time.sleep(3)

    except StaleElementReferenceException:
        print("StaleElementReferenceException")
    except InvalidArgumentException:
        print("InvalidArgumentException")
    except TimeoutException:
        print("TimeoutException")
    except NoSuchElementException:
        print("NoSuchElementException")
    except WebDriverException:
        print("WebDriverException")
    except IndexError:
        print("IndexError")

work_book.save("ani_detail.xlsx")
driver.quit()

