import shutil
import time

from openpyxl.styles import Font
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook, Workbook
import pandas as pd

# from Settings.selenium_setting import *
from selenium import webdriver


def chd_active(url, xpath):
    # 크롬드라이버 설정
    chd = 'C:/dev_files/chd/chd.exe'
    driver = webdriver.Chrome(chd)
    driver.get(url)
    text = driver.find_element_by_xpath(xpath).text
    time.sleep(3)
    driver.quit()
    print(text)
    return text


# 엑셀파일 쓰기
work_book = load_workbook("ani_a.xlsx")
sheet = work_book['Sheet']
title = "//*[@id='app']/div/div[2]/article/div[3]/div[2]/div/div/div[1]/table/tbody/tr[3]/td[2]/div/a"

chd_active(sheet["B1"].value, title)

time.sleep(3)

chd_active(sheet["B2"].value, title)

# title = driver.find_element_by_xpath("//*[@id='app']/div/div[2]/article/div[3]/div[2]/div/div/div[1]/table/tbody/tr[3]/td[2]/div/a")

