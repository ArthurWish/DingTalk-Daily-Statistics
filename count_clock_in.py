import xlrd
import xlsxwriter

class clock_in:
    wb = None
    
    def __init__(self, file, workbook) -> None:
        self.wb = xlrd.open_workbook(file)
        self.workbook = xlsxwriter.Workbook(workbook)
        self.worksheet = self.workbook.add_worksheet()
        
    def read_member(self):
        member_list = []
        sheet = self.wb.sheet_by_index(0)
        for i in range(1, sheet.nrows):
            member = sheet.cell_value(i, 1)
            if member not in member_list:
                member_list.append(member)
                
        for i in range(len(member_list)):
            self.worksheet.write(i+1, 0, member_list[i])
        return member_list
    
    def read_month(self):
        month_list = []
        sheet = self.wb.sheet_by_index(0)
        for i in range(1, sheet.nrows):
            time = sheet.cell_value(i, 3)
            # month = re.split('年|月|日', time)
            month = time[6:8]
            if month not in month_list:
                month_list.append(month)
        self.worksheet.write(0, 0, "姓名")
    
        for i in range(1, len(month_list)+1):
            self.worksheet.write(0, i, month_list[i-1])
        
        return month_list
    def count(self):
        month_list = self.read_month()
        member_list = self.read_member()
        
        for m in range(len(month_list)):
            count_list = [0] * len(member_list)
            sheet = self.wb.sheet_by_index(0)
            for index, member in enumerate(member_list):
                for i in range(sheet.nrows):
                    name = sheet.cell_value(i, 1)
                    month = sheet.cell_value(i, 3)[6:8]
                    if member == name and month == month_list[m]:
                        count_list[index] += 1
            for c in range(len(count_list)):
                self.worksheet.write(c+1, m+1, count_list[c])
            count_list = [0] * len(member_list)
            
    
    def main(self):
        cl.read_month()
        cl.read_member()
        cl.count()
        self.workbook.close()
        
cl = clock_in(file='日志报表20220402162634718.xls', workbook='01.xlsx')
cl.main()
        