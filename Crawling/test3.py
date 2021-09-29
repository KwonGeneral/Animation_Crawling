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
# # driver = webdriver.Chrome(chd)
# # driver.get("https://www.chuing.net/db/search.php?cdbsearch=asdfasdf")
#
# options = webdriver.ChromeOptions()
# options.add_argument("headless")
# options.add_argument("window-size=1920x1080")
# options.add_argument("disable-gpu")
# options.add_argument("User-Agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_4) AppleWebKit/537.36 (KHTML, "
#                      "like Gecko) Chrome/93.0.4577.82")
#
# driver = webdriver.Chrome(chd, options=options)
# driver.get("https://www.chuing.net/db/search.php?cdbsearch=asdfasdf")
#
# # xlsx_list = ["ani_a.xlsx", "ani_b.xlsx", "ani_c.xlsx", "ani_d.xlsx", "ani_e.xlsx", "ani_f.xlsx", "ani_g.xlsx",
# #              "ani_h.xlsx", "ani_i.xlsx", "ani_j.xlsx", "ani_k.xlsx", "ani_l.xlsx", "ani_m.xlsx", "ani_n.xlsx",
# #              "ani_o.xlsx", "ani_p.xlsx"]
#
# xlsx_list = ["ani_c.xlsx", "ani_d.xlsx", "ani_e.xlsx", "ani_f.xlsx", "ani_g.xlsx",
#              "ani_h.xlsx", "ani_i.xlsx", "ani_j.xlsx", "ani_k.xlsx", "ani_l.xlsx", "ani_m.xlsx", "ani_n.xlsx",
#              "ani_o.xlsx", "ani_p.xlsx"]
#
# # xl = xlsx_list[1]
#
# search_xpath = "//*[@id='SjestForm']/div/div[1]/input"
# result_a_xpath = "/html/body/div[5]/div[2]/div/div[2]/div/div[4]/div[1]"
# # result_a_xpath = "/html/body/div[5]/div[2]/div/div[2]/div/div[4]/div[1]/text()[2]"
#
# for xl in xlsx_list:
#     work_book = load_workbook(xl)
#     sheet = work_book['Sheet']
#
#     for no in range(2, len(sheet["A"]) + 1):
#         try:
#             search_tag = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, search_xpath)))
#             search_tag.clear()
#
#             search_tag.send_keys(sheet["A" + str(no)].value)
#             search_tag.send_keys(Keys.ENTER)
#
#             result_a = driver.find_elements_by_xpath(result_a_xpath)
#
#             result_total = ""
#
#             if len(result_a) > 0:
#                 for a in result_a:
#                     result_total = result_total + str(a.text).split("장르: ")[1].replace(",", ", ") + ", "
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
#                 result_total = result_total.rstrip(", ").lstrip(",").lstrip(" ")
#                 print("result_total 22 : ", result_total)
#                 if sheet["B" + str(no)].value is None:
#                     sheet["B" + str(no)].value = ""
#                 sheet["B" + str(no)].value = result_total
#
#             work_book.save(xl)
#
#             time.sleep(3)
#
#         except StaleElementReferenceException:
#             print("StaleElementReferenceException")
#         except InvalidArgumentException:
#             print("InvalidArgumentException")
#         except TimeoutException:
#             print("TimeoutException")
#             time.sleep(30)
#         except NoSuchElementException:
#             print("NoSuchElementException")
#         except WebDriverException:
#             print("WebDriverException")
#         except IndexError:
#             print("IndexError")
#
# driver.quit()
#
