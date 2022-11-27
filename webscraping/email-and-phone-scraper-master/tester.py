from ExcelMaker import ExcelMaker


emails = []
emails+= [None,None]
emails+= ["namanjain", None]
emails+= []
emails+= ["namanjain", "1@gmail.com"]
emails+= ["namanjain", "2@gmail.com"]
emails+= ["namanjain", "3@gmail.com"]

emails+= ["radjain", "4@gmail.com"]

emails+= ["dadjain", "5@gmail.com"]

for email in emails:
    print(email)

maker = ExcelMaker(fname="haha", items=None)
maker.makeExcelByEmails(emails=emails,savefileName= "Savefile")