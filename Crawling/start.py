

# 엑셀파일
from openpyxl import load_workbook

from Crawling.ani_detail_crawl import namuwiki_crawl

work_book = load_workbook("ani_b.xlsx")
sheet = work_book['Sheet']

# 나무위키 크롤링 1)
for no in range(2, len(sheet["A"]) + 1):
    namuwiki_crawl(work_book, "ani_b.xlsx", sheet["H"+str(no)].value, sheet["B"+str(no)])

# 나무위키 크롤링 2)
# for no in range(1, len(sheet["A"]) + 1):
#     if sheet["B"+str(no)].value is None:
#         print("제목 : ", sheet["A"+str(no)].value)
#         namuwiki_crawl(work_book, "ani_a.xlsx", sheet["H" + str(no)].value, sheet["B" + str(no)])

# namuwiki_crawl(work_book, "ani_a.xlsx", sheet["H86"].value, sheet["B86"])

# 구글 크롤링

# 건강 전라계 수영부 우미쇼 38
# 게이트 - 자위대. 그의 땅에서, 이처럼 싸우며 55
# 고식 73
# 골프천재 탄도 86
# 괴짜가족 107
# 귀를 기울이면 109
# 구구레 코쿠리상 = 나와라! 코쿠리상 114
# 구인사가 119
# 구명전사 제노사이버 117
# 귀신전 123
# 그래플러 바키 140
# 금색의 갓슈 151
# 기분 나쁜 모노노케안 168

#글라스립 153
#그러나 죄인은 용과 춤춘다