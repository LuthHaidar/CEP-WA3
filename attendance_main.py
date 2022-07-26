from openpyxl import Workbook

class Attendance_Entry:
    def __init__(self, name, level, date):
        self.name = name
        self.level = level
        self.date = date
    
    def write_to_excel(self): #writes yes in the correct cell according to name and date
        pass
Luth = Attendance_Entry('Luth', '3C', '26/7/2022')
Luth.write_to_excel()

