from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
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

# 엑셀파일 생성
save_book = Workbook()
# 시트 입력
save_sheet = save_book.active
save_sheet.column_dimensions["A"].width = 30
save_sheet.column_dimensions["B"].width = 55

save_sheet['A1'] = "분기"
save_sheet['A1'].font = Font(name="나눔고딕", color="000000", bold=True)
save_sheet['A1'].alignment = Alignment(horizontal='center')
save_sheet['B1'] = "링크"
save_sheet['B1'].font = Font(name="나눔고딕", color="000000", bold=True)
save_sheet['B1'].alignment = Alignment(horizontal='center')

find_date_xpath = "/html/body/div/div/div/div/ul/li/a"

find_date = driver.find_elements_by_xpath(find_date_xpath)

temp_update = []
for index, da in enumerate(find_date):
    save_sheet["A"+str(index + 2)].value = da.text
    save_sheet["B"+str(index + 2)].value = da.get_attribute("href")
    temp_update.append(da.text)

update_file = open("update_log.txt", 'r+')  # r : read, w : write, a : append
update_file.write("\n".join(temp_update))
update_file.close()

save_book.save("ani_find_date.xlsx")
driver.quit()
