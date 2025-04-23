import os
import platform
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText

# schedule类(排班表)
class Schedule:
    name = []
    date = []

    # 格式化文本信息
    def _format(self,str):
        str = str.replace(" ","")
        str = str.replace("\t","")
        str = str.replace("[","")
        str = str.replace("]","")
        return str
    
    # 分割有效信息（需要重写）
    def _split(self,str):
        str = str.replace("\n","")
        name = str.split("=")
        date = name[1].split(",")
        return name[0],date
    
    # 读取文件内容（暂时没用，需要重写）
    def loadfile(self,workbook,filename):
        with open(filename,'r',encoding='utf-8') as file:
            for line in file.readlines():
                if "#" in line :
                    continue
                data = self._split(line)
                self.name.append(data[0])
                self.date.append(data[1])
    
    # 导出关键信息txt
    def exportfile(self,workbook,filename):
        # 初始化
        wb = workbook
        ws = workbook.active
        
        # 处理数据
        with open(filename,"w",encoding="utf-8") as file:
            file.write("# 张三=1,2,3,4   :家中有事需要连休\n")
            for row_index in range(4,20):
                if ws.cell(row=row_index,column=1).value == None:
                    break
                self.name.append(ws.cell(row=row_index,column=1).value)
                for column_index in range(2,33): #处理休息日
                    if ws.cell(row=row_index,column=column_index).value == None:
                        break
                    elif ws.cell(row=row_index,column=column_index).value == "休":
                        self.date.append([])
                        self.date[row_index-4].append(ws.cell(row=3,column=column_index).value)
                remark = ws.cell(row=row_index,column=33).value
                if remark == None:
                    remark = ""
                else:
                    remark = "      :" + remark
                text = f"{self.name[row_index-4]}={self.date[row_index-4]}{remark}\n"
                text = text.replace("[","")
                text = text.replace("]","")
                file.write(text)
        
        # 检测系统并打开txt
        if platform.system() == "Windows":
            os.startfile(filename)
                
    
    # 将txt内容导入排班表
    def importfile(self,workbook,filename):
        pass


# 获取年月
def fetch_year_month():
    year_now = datetime.now().year
    month_now = datetime.now().month
    
    while True:
        try:
            year = input(f"请输入当前年份（默认为{year_now}年）：")
            if year == "":
                year = year_now
                break
            year = int(year)
            if 0<=year<=3000:
                break
            else:
                print("请输入正确的年份！")

        except ValueError:
            print("请输入正确的年份！")

    while True:
        try:
            month = input(f"请输入当前月份（默认为{month_now}月）：")
            if month == "":
                month = month_now
                break
            month = int(month)
            if 1<=month<=12:
                break
            else:
                print("请输入正确的月份！")

        except ValueError:
            print("请输入正确的月份！")

    if month == 12:
        title = (f"{year}年{month}月21日至{year+1}年1月20日首钢一期安全员考勤排班表")
    else:
        title = (f"{year}年{month}月21日至{year}年{month+1}月20日首钢一期安全员考勤排班表")
    return title

#高亮关键字
def hiidle_text(title,red_text,sz=17,rFont="黑体"):
    rich_text = CellRichText()
    start_index = title.find(red_text)
    end_index = start_index + len(red_text)
    rich_text.append(TextBlock(InlineFont(color="000000",sz=20,rFont="微软雅黑"), title[:start_index]))
    rich_text.append(TextBlock(InlineFont(color="C00000",sz=20,rFont="微软雅黑"), red_text))
    rich_text.append(TextBlock(InlineFont(color="000000",sz=20,rFont="微软雅黑"), title[end_index:]))
    return rich_text


if __name__ == "__main__":
    # 初始化
    sch = Schedule()
    wb = load_workbook("排班.xlsx")
    ws = wb.active
    
    # 修改表头
    title = fetch_year_month()
    red_text = "首钢一期安全员"
    rich_text = hiidle_text(title,red_text,20,"黑体")

    # 获取列表关键信息导出txt
    sch.exportfile(wb,"content.txt")

    #ws['A1'].value = rich_text
    #wb.save("output.xlsx")

    #for i, name in enumerate(sch.name,start=1):
        #ws.cell(row=i,column=1,value=name)
        #ws.cell(row=i,column=2,value=str(sch.date[i-1]))

    #wb.save("output.xlsx")
    #print("已导出 output.xlsx")