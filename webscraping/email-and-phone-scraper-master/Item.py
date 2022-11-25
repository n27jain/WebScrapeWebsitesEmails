

class Item:
    def __init__(self, Name , Addr , Web , Email , PhoneNum ):
        self.Name = Name
        self.Addr = Addr
        self.Web = Web
        self.Email = Email
        self.Phone = PhoneNum
    def printItem(self):
        print( "printing Item details : ",
            "Name  :", self.Name,
            "Addr  :",self.Addr,
            "Web  :",self.Web,
            "Email  :",self.Email,
            "PhoneNum :", self.Phone
        )
    