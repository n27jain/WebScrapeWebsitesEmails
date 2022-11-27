import xlsxwriter
from Item import Item
import pandas as pd



class ExcelMaker:
    def __init__(self, fname, items):
        self.fname  = fname
        self.items = items
    def makeExcel(self):
        if self.items:
            print("here again : ", self.fname)
            workbook = xlsxwriter.Workbook("app/"+self.fname + ".xlsx")
            worksheet = workbook.add_worksheet()
            worksheet.write(0, 0, "Name")
            worksheet.write(0, 1, "Address")
            worksheet.write(0, 2, "Website")
            worksheet.write(0, 3, "Email")
            i = 1
            for item in self.items:
                worksheet.write(i + 1, 0, item.name )
                worksheet.write(i + 1, 1, item.addr )
                worksheet.write(i + 1, 2, item.web )
                worksheet.write(i + 1, 3, item.email  )
                i += 1   
            workbook.close()
            return "app/"+ self.fileName + ".xlsx" # save file to machine
        else:
            return None
    def makeExcelByEmails(self, emails):

        print("here again : ", self.fname)
        workbook = xlsxwriter.Workbook("app/"+self.fname + ".xlsx")
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, "Name")
        worksheet.write(0, 1, "Email")
        i = 1
        current = None
        for email in emails:
            print("type: ", type(email))
            print("value: ", email)
            if email[0] == current:
                worksheet.write(i + 1, 1, email[1])
                i += 1
            else:
                current = email[0]
                worksheet.write(i + 1, 0, email[0])
                worksheet.write(i + 1, 1, email[1])
        workbook.close()
        return
            
    def readExcel(self):
        df = pd.read_excel(self.fname)
        size = df.Name.size
        items= [] 
        for i in range(size):
            row = df.loc[i].values
            item = Item(row[0], row[1] ,row[2], row[3], row[4])
            # item.printItem()
            items.append(item)
        self.items = items
    

        
