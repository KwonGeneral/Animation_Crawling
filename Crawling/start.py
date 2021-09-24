

# 엑셀파일
from openpyxl import load_workbook
from selenium.common.exceptions import InvalidArgumentException, WebDriverException

from Crawling.ani_detail_crawl import namuwiki_crawl


# work_book = load_workbook("ani_k.xlsx")
# sheet = work_book['Sheet']
# for no in range(64, len(sheet["A"]) + 1):
#     try:
#         namuwiki_crawl(work_book, "ani_k.xlsx", sheet["H"+str(no)].value, sheet["B"+str(no)])
#     except InvalidArgumentException:
#         print("InvalidArgumentException2")
#     except WebDriverException:
#         print("WebDriverException")

print("\n 1 : 나무위키 크롤링\n")
print(" 2 : 구글 크롤링\n")
select = input(" 선택 : ")

# 나무위키 크롤링
if select == "1":
    xlsx_list = ["ani_l.xlsx", "ani_m.xlsx", "ani_n.xlsx", "ani_o.xlsx", "ani_p.xlsx"]

    for xl in xlsx_list:
        work_book = load_workbook(xl)
        sheet = work_book['Sheet']
        for no in range(2, len(sheet["A"]) + 1):
            try:
                namuwiki_crawl(work_book, xl, sheet["H"+str(no)].value, sheet["B"+str(no)])
            except InvalidArgumentException:
                print("InvalidArgumentException2")

# 구글 크롤링
if select == "2":
    pass
