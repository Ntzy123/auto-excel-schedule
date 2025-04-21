import os
from openpyxl import Workbook


# schedule类(排班表)
class Schedule:
    name = []
    date = []

    # 分割有效信息
    def _split(self,str):
        str = str.replace("\n","")
        name = str.split("=")
        date = name[1].split(",")
        return name[0],date
    
    # 读取文件内容
    def load(self,filename):
        with open(filename,'r',encoding='utf-8') as file:
            for line in file.readlines():
                if "#" in line :
                    continue
                data = self._split(line)
                self.name.append(data[0])
                self.date.append(data[1])



sch = Schedule()
sch.load('1.txt')
wb = Workbook()
ws = wb.active
ws.append(["星期",])
for name in sch.name:
    
ws['A1'] = lst[0][0]
ws['B1'] = lst[0][1]
ws['A2'] = lst[1][0]
ws['B2'] = lst[1][1]
wb.save("output.xlsx")
os.system("echo '已导出 output.xlsx'")


#print(sch.name)
#print()
#print(sch.date)