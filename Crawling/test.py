#
# import shutil
# import time
# from urllib.request import urlretrieve
#
# from openpyxl.styles import Font
# from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException, \
#     InvalidArgumentException, WebDriverException
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.support.wait import WebDriverWait
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support import expected_conditions as EC
# from openpyxl import load_workbook, Workbook
# import pandas as pd
#
# # from Settings.selenium_setting import *
# from selenium import webdriver
#
# chd = 'C:/dev_files/chd/chd.exe'
#
# # 헤들리스 설정
# # options = webdriver.ChromeOptions()
# # options.add_argument("headless")
# #
# # # 웹페이지 사이즈 설정
# # options.add_argument("window-size=1920x1080")
# #
# # # 그래픽 사용 안함
# # options.add_argument("disable-gpu")
# #
# # # 인증 정보
# # options.add_argument("User-Agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_4) AppleWebKit/537.36 (KHTML, "
# #                      "like Gecko) Chrome/93.0.4577.82")
# #
# # # 언어 설정 : 한국어
# # options.add_argument("lang=ko_KR")
# #
# # driver = webdriver.Chrome(chd, options=options)
# driver = webdriver.Chrome(chd)
#
# driver.get("https://www.google.com/search?q=.")
#
# first_search = "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input"
# search = "//*[@id='tsf']/div[1]/div[1]/div[2]/div/div[2]/input"
#
# # 꿈을 먹는 메리
# result_a_xpath = "/html/body/div[7]/div/div[6]/div/div/div/div[1]/div/div[1]/g-scrolling-carousel/" \
#                  "div[1]/div/div/a/div/div/div[2]/div"
# # 가난뱅이 신이
# result_b_xpath = "//*[@id='rso']/div[1]/div/div[1]/div[1]/div[1]/div/div[1]/div/div[1]/div[2]/div/div[1]"
# # 가면의 메이드가이
# result_c_xpath = "//*[@id='rso']/div[1]/div/div[1]/div/div[1]/div/div/div/div/div[1]/div/div/table/tbody/tr[2]/td[2]"
# result_c_xpath_check = "//*[@id='rso']/div[1]/div/div[1]/div/div[1]/div/div/div/div/div[1]/div/div/table/tbody/" \
#                        "tr[2]/td[1]/b"
# # 가시나무 왕
# result_d_xpath = "/html/body/div[7]/div/div[6]/div/div/div/div[1]/div/div[1]/g-scrolling-carousel/div[1]/div/div/a/" \
#                  "div/div/div/div"
#
# xlsx_list = ["ani_h.xlsx", "ani_i.xlsx", "ani_j.xlsx", "ani_k.xlsx", "ani_l.xlsx", "ani_m.xlsx", "ani_n.xlsx",
#              "ani_o.xlsx", "ani_p.xlsx"]
#
# for xl in xlsx_list:
#     # xl = xlsx_list[1]
#     work_book = load_workbook(xl)
#     sheet = work_book['Sheet']
#     for no in range(2, len(sheet["A"]) + 1):
#         try:
#             search_tag = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, search)))
#             search_tag.clear()
#
#             search_tag.send_keys(sheet["A" + str(no)].value + " 장르")
#             search_tag.send_keys(Keys.ENTER)
#
#             result_a = driver.find_elements_by_xpath(result_a_xpath)
#             result_b = driver.find_elements_by_xpath(result_b_xpath)
#             result_c = driver.find_elements_by_xpath(result_c_xpath)
#             result_d = driver.find_elements_by_xpath(result_d_xpath)
#
#             result_total = ""
#
#             if len(result_a) > 0:
#                 for a in result_a:
#                     result_total = result_total + a.text + ", "
#             elif len(result_b) > 0:
#                 for b in result_b:
#                     if "http" in b.text:
#                         for c in result_c:
#                             result_total = result_total + c.text + ", "
#                     else:
#                         result_total = result_total + b.text + ", "
#             elif len(result_c) > 0:
#                 for c in result_c:
#                     result_total = result_total + c.text + ", "
#             elif len(result_d) > 0:
#                 for d in result_d:
#                     result_total = result_total + d.text + ", "
#
#             print(" 제목 : [ ", str(no), " ] ", sheet["A" + str(no)].value)
#
#             if sheet["B" + str(no)].value is not None and sheet["B" + str(no)].value != "":
#                 if result_total.replace(",", "").replace(" ", "") != "":
#                     result_total = result_total.rstrip(", ")
#                     result_total = ", " + result_total
#                     if result_total != ", ":
#                         print("result_total 11 : ", result_total)
#                         sheet["B" + str(no)].value = sheet["B" + str(no)].value + result_total
#             else:
#                 # print("sheet value 2 : ", sheet["B" + str(no)].value)
#                 result_total = result_total.rstrip(", ").lstrip(",").lstrip(" ")
#                 print("result_total 22 : ", result_total)
#                 if sheet["B" + str(no)].value is None:
#                     sheet["B" + str(no)].value = ""
#                 sheet["B" + str(no)].value = result_total
#
#             work_book.save(xl)
#
#         except StaleElementReferenceException:
#             print("StaleElementReferenceException")
#         except InvalidArgumentException:
#             print("InvalidArgumentException")
#         except TimeoutException:
#             print("TimeoutException")
#             driver.quit()
#             # options = webdriver.ChromeOptions()
#             # options.add_argument("headless")
#             # options.add_argument("disable-gpu")
#             # options.add_argument("User-Agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_4) AppleWebKit/537.36 "
#             #                      "(KHTML, like Gecko) Chrome/93.0.4577.82")
#             # driver = webdriver.Chrome(chd, options=options)
#             driver = webdriver.Chrome(chd)
#             driver.get("https://www.google.com/search?q=.")
#         except NoSuchElementException:
#             print("NoSuchElementException")
#         except WebDriverException:
#             print("WebDriverException")
#
# driver.quit()
#
# # first_search_tag = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, first_search)))
# # first_search_tag.clear()
# # first_search_tag.send_keys(sheet["A9"].value + " 장르")
# # first_search_tag.send_keys(Keys.ENTER)
# #
# # result_a = driver.find_elements_by_xpath(result_a_xpath)
# # result_b = driver.find_elements_by_xpath(result_b_xpath)
# # result_c = driver.find_elements_by_xpath(result_c_xpath)
# # result_d = driver.find_elements_by_xpath(result_d_xpath)
# #
# # if len(result_a) > 0:
# #     for a in result_a:
# #         print("result_a : ", a.text)
# # elif len(result_b) > 0:
# #     for b in result_b:
# #         if "http" in b.text:
# #             for c in result_c:
# #                 print("result_c : ", c.text)
# #         else:
# #             print("result_b : ", b.text)
# # elif len(result_c) > 0:
# #     for c in result_c:
# #         print("result_c : ", c.text)
# # elif len(result_d) > 0:
# #     for d in result_d:
# #         print("result_d : ", d.text)
# #
# # time.sleep(2)
