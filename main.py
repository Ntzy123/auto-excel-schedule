from openpyxl import load_workbook
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText

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
    
    # 读取文件内容（暂时没用）
    def loadfile(self,filename):
        with open(filename,'r',encoding='utf-8') as file:
            for line in file.readlines():
                if "#" in line :
                    continue
                data = self._split(line)
                self.name.append(data[0])
                self.date.append(data[1])
    
    def exportfile(self):
        pass



if __name__ == "__main__":
    # 初始化
    sch = Schedule()
    wb = load_workbook("排班.xlsx")
    ws = wb.active
    
    # 获取月份
    while True:
        try:
            month = int(input("请输入当前月份："))
            if 1<=month<=12:
                break
            else:
                print("请输入正确的月份！")

        except ValueError:
            print("请输入正确的月份！")

    prefix = (f"2025年{month}月21日至2025年{month+1}月20日")
    red_text = "首钢一期安全员"
    suffix_text = "考勤排班表"

    rich_text = CellRichText()
    rich_text.append(TextBlock(InlineFont(color="000000",sz=20,rFont="微软雅黑"), prefix))
    rich_text.append(TextBlock(InlineFont(color="C00000",sz=20,rFont="微软雅黑"), red_text))
    rich_text.append(TextBlock(InlineFont(color="000000",sz=20,rFont="微软雅黑"), suffix_text))
    ws['A1'].value = rich_text
    wb.save("output.xlsx")

    #for i, name in enumerate(sch.name,start=1):
        #ws.cell(row=i,column=1,value=name)
        #ws.cell(row=i,column=2,value=str(sch.date[i-1]))

    #wb.save("output.xlsx")
    #print("已导出 output.xlsx")