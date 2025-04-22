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
    def loadfile(self,filename):
        with open(filename,'r',encoding='utf-8') as file:
            for line in file.readlines():
                if "#" in line :
                    continue
                data = self._split(line)
                self.name.append(data[0])
                self.date.append(data[1])



sch = Schedule()
sch.loadfile('1.txt')
wb = Workbook()
ws = wb.active
for i, name in enumerate(sch.name,start=1):
    ws.cell(row=i,column=1,value=name)
    ws.cell(row=i,column=2,value=str(sch.date[i-1]))

wb.save("output.xlsx")
print("已导出 output.xlsx")


#print(sch.name)
#print()
#print(sch.date)