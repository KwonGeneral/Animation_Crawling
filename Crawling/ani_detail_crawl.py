import shutil
import time
from urllib.request import urlretrieve

from openpyxl.styles import Font
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException, \
    InvalidArgumentException, WebDriverException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook, Workbook
import pandas as pd

# from Settings.selenium_setting import *
from selenium import webdriver


def namuwiki_crawl(workbook, file_name, url, sheet):
    # 크롬드라이버 설정
    if url == "" or url is None:
        return

    chd = 'C:/dev_files/chd/chd.exe'
    driver = webdriver.Chrome(chd)
    driver.get(url)

    # 만화 정보
    cartoon_a_genre = "//*[@id='app']/div/div[2]/article/div[3]/div/div/div/div/table/tbody/tr/td[2]/div/a"

    cartoon_b_genre = "//*[@id='app']/div/div[2]/article/div[3]/div/div/div/div/table/tbody/tr[4]/td[2]/div/a"

    cartoon_c_genre = "//*[@id='app']/div/div[2]/article/div[3]/div/div/div/div/table/tbody/tr[3]/td[2]/div/a"

    cartoon_check_a_genre = "//*[@id='app']/div/div[2]/article/div[3]/div/div/div/div/table/tbody/tr/td[" \
                            "1]/div/strong/span"

    # 애니메이션 정보
    ani_a_genre = "//*[@id='app']/div/div[2]/article/div[3]/div/div/div/div/" \
                  "table/tbody/tr/td/div/span/div/dl/dd/div/table/tbody/tr[1]/td[2]/div/a"

    ani_b_genre = "//*[@id='app']/div/div[2]/article/div[3]/div/div/div/div/" \
                  "table/tbody/tr[3]/td/div/span/div/dl/dd/div/table/tbody/tr[1]/td[2]/div"

    ani_check_a_genre = "//*[@id='app']/div/div[2]/article/div[3]/div/div/div/div/" \
                        "table/tbody/tr[3]/td/div/span/div/dl/dd/div/table/tbody/tr[1]/td[1]/div/strong/span"

    ani_check_b_genre = "//*[@id='app']/div/div[2]/article/div[3]/div/div/div/div/" \
                        "table/tbody/tr[3]/td/div/span/div/dl/dd/div/table/tbody/tr[1]/td[2]/div/strong/span"

    total_genre = ""

    try:
        text = driver.find_elements_by_xpath(cartoon_a_genre)
        check_genre = ""
        try:
            check_genre = driver.find_element_by_xpath(cartoon_check_a_genre).text
            if "작품" in check_genre:
                check_genre = "작품"
                text = driver.find_elements_by_xpath(cartoon_b_genre)

        except NoSuchElementException:
            pass

        if len(text) > 0:
            if check_genre == "장르":
                try:
                    text2 = driver.find_elements_by_xpath(cartoon_c_genre)

                    if len(text2) > 0:
                        if check_genre == "장르":
                            count = 0
                            for zz in text2:
                                if count == 0:
                                    count += 1
                                    total_genre = zz.text
                                else:
                                    total_genre = total_genre + ", " + zz.text

                    else:
                        text2 = driver.find_elements_by_xpath(cartoon_b_genre)
                        count = 0
                        for aa in text2:
                            if count == 0:
                                count += 1
                                total_genre = aa.text
                            else:
                                total_genre = total_genre + ", " + aa.text

                except NoSuchElementException:
                    text = driver.find_elements_by_xpath(cartoon_b_genre)

                    count = 0
                    for aa in text:
                        if count == 0:
                            count += 1
                            total_genre = aa.text
                        else:
                            total_genre = total_genre + ", " + aa.text
            else:
                try:
                    text2 = driver.find_elements_by_xpath(cartoon_c_genre)

                    if len(text2) > 0:
                        if check_genre == "장르":
                            count = 0
                            for zz in text2:
                                if count == 0:
                                    count += 1
                                    total_genre = zz.text
                                else:
                                    total_genre = total_genre + ", " + zz.text
                        else:
                            text = driver.find_elements_by_xpath(cartoon_c_genre)

                            count = 0
                            for aa in text:
                                if count == 0:
                                    count += 1
                                    total_genre = aa.text
                                else:
                                    total_genre = total_genre + ", " + aa.text

                    else:
                        text2 = driver.find_elements_by_xpath(cartoon_b_genre)
                        count = 0
                        for aa in text2:
                            if count == 0:
                                count += 1
                                total_genre = aa.text
                            else:
                                total_genre = total_genre + ", " + aa.text

                except NoSuchElementException:
                    pass
        else:
            try:
                btn = "//*[@id='app']/div/div[2]/article/div[3]/div[2]/div/div/div/table/tbody/tr/td"
                btn2 = "//*[@id='app']/div/div[2]/article/div[3]/div[3]/div/div/div/table/tbody/tr/td"
                btn3 = "//*[@id='app']/div/div[2]/article/div[3]/div[4]/div/div/div/table/tbody/tr/td"

                btn_text = driver.find_elements_by_xpath(btn)
                btn2_text = driver.find_elements_by_xpath(btn2)
                btn3_text = driver.find_elements_by_xpath(btn3)

                for nnn in btn_text:
                    if "작품" in nnn.text:
                        nnn.click()

                for qdq in btn2_text:
                    if "작품" in qdq.text:
                        qdq.click()

                for wdw in btn3_text:
                    if "작품" in wdw.text:
                        wdw.click()

                time.sleep(1)

                text = driver.find_elements_by_xpath(ani_a_genre)

                check_genre = ""

                try:
                    check_genre = driver.find_element_by_xpath(ani_check_a_genre).text
                except NoSuchElementException:
                    pass

                if len(text) > 0:
                    if check_genre == "장르":
                        count = 0
                        for cc in text:
                            if count == 0:
                                count += 1
                                total_genre = cc.text
                            else:
                                total_genre = total_genre + ", " + cc.text
                else:
                    text = driver.find_elements_by_xpath(ani_b_genre)

                    check_genre = ""

                    try:
                        check_genre = driver.find_element_by_xpath(ani_check_b_genre).text
                    except NoSuchElementException:
                        pass

                    if len(text) > 0:
                        if check_genre == "장르":
                            count = 0
                            for cc in text:
                                if count == 0:
                                    count += 1
                                    total_genre = cc.text
                                else:
                                    total_genre = total_genre + ", " + cc.text

            except TimeoutException:
                print("TimeoutException")
            except NoSuchElementException:
                print("NoSuchElementException")
    except StaleElementReferenceException:
        print("StaleElementReferenceException")
        driver.quit()
        return total_genre
    except InvalidArgumentException:
        print("InvalidArgumentException")
        driver.quit()
        return total_genre
    except WebDriverException:
        print("WebDriverException")
        driver.quit()
        return total_genre

    print("total_genre : ", total_genre)
    driver.quit()

    sheet.value = total_genre
    workbook.save(file_name)
    return total_genre
