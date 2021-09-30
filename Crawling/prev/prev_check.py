from openpyxl import load_workbook

work_book = load_workbook("prev_total.xlsx")
sheet = work_book['Sheet']

blank_data = []
for no in range(2, len(sheet["A"])):
    if sheet["B"+str(no)].value == "" or sheet["B"+str(no)].value is None:
        blank_data.append("[ " + str(no) + " ] " + sheet["A"+str(no)].value)

for i in blank_data:
    print(i)

