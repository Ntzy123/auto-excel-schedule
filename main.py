from openpyxl import Workbook

wb = Workbook()
ws = wb.active
lst = [
  ["倪昌飞","1,2,3,4"],
  ["李海波","5,6,7,8"]
]
ws['A1'] = lst[0][0]
ws['B1'] = lst[0][1]
ws['A2'] = lst[1][0]
ws['B2'] = lst[1][1]
wb.save("output.xlsx")
