
from openpyxl import Workbook, load_workbook


work_book = load_workbook("ani_detail.xlsx")
sheet = work_book['Sheet']

check_list = []

for no in range(2, len(sheet["A"]) + 1):
    if sheet["B"+str(no)].value is not None:
        for kk in sheet["B"+str(no)].value.split(", "):
            check_list.append(kk)

check_list = list(set(check_list))

print(check_list)

# ['마녀', '스포츠', '마법', '퇴마', '추리', '뱀파이어', '천사', '코미디', '유령', 'TS', '먼치킨', '연애', '범죄', '기타',
# '메이드', '모험', '버추얼 리얼리티', 'SF', '하렘', '게임', '치유물', '시대', '집사', '닌자', '전쟁', '성우', '악마', '능력',
# '드라마', '이세계', '요괴', '귀신', '괴물', '영웅', '일상', '동물', '신', '사무라이', '쇼타', '배틀', '전투', '학원', '드래곤',
# '공포', '좀비', '메카닉', '변신', '군', '판타지', 'BL', '음악', '마법소녀', '정령', '아이돌', '액션', '로리', '로맨스',
# '소꿉친구', '요리', '멘붕', '미스테리', '카페', '부활동', '마왕', '아동물', '재판', '백합']
