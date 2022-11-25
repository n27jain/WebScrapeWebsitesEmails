#!/usr/bin/env python3

'''
Webpage scraper by Cameron Cobb                
Scrapes for phone and emails and places in     
spreadsheet.                                   
Python 3 required                              
'''

import re
from urllib.request import urlopen, Request
import os
from datetime import datetime

from openpyxl import Workbook
from bs4 import BeautifulSoup

from Item import Item
from ExcelMaker import ExcelMaker

save_excel = True #Change to "True" to save email into Excel

# book = Workbook()
# sheet = book.active


def start_scrape(page, name_the_file):

    print("\n\nWebpage is currently being scrapped... please wait...")
       
    scrape = BeautifulSoup(page, 'html.parser')
    scrape = scrape.get_text()
    
    phone_numbers = set(re.findall(r"((?:\d{3}|\(\d{3}\))?(?:\s|-|\.)?\d{3}(?:\s|-|\.)\d{4})", scrape)) #"set" removes duplicates
    emails = set(re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,3}", scrape))

    nodupnumber = len(list(phone_numbers))
    nodupemail = len(list(emails))

    dupnumber = len(list(re.findall(r"((?:\d{3}|\(\d{3}\))?(?:\s|-|\.)?\d{3}(?:\s|-|\.)\d{4})", scrape))) 
    dupemail = len(list(re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,3}", scrape)))

    number_of_dup_number = int(dupnumber) - int(nodupnumber) 
    number_of_dup_email = int(dupemail) - int(nodupemail)

    email_list = list(emails)
    phone_numbers_list = list(phone_numbers)
    print("FOUND /n")
    print("emails:", emails)
    print("phone_numbers:", phone_numbers)

   
    if len(phone_numbers) == 0:
        print("No phone number(s) found.")

        print("-----------------------------\n")
    else:
        count = 1
        for item in phone_numbers:
            print("Phone number #" + str(count) + ': ' + item)
            count += 1

    print("-----------------------------\n")

    if len(emails) == 0:
        print("No email address(es) found.")
        print("-----------------------------\n")
    else:
        count = 1
        for item in emails:
            print('Email address #' + str(count) + ': ' + item)
            count += 1

    # if save_excel:
    #     for row in zip(email_list):
    #         sheet.append(row)
    #     excel_file = (name_the_file + ".xlsx")
    #     book.save(excel_file) 
       
  

    # if save_excel:
    #     print("\n\nData has been stored inside of an Excel spreadsheet named: "
    #           + excel_file + " in this directory: " + os.getcwd())
    #     mod_time = os.stat(excel_file).st_mtime
    #     print("\nCompleted at: " + str(datetime.fromtimestamp(mod_time)))
    #     print("\nSize of file: " + str(os.stat(excel_file).st_size) + " KB")
    
    return email_list, phone_numbers_list

def startScrapeFromFile():

    maker = ExcelMaker("California.xlsx", None)
    maker.readExcel()
    emails = []
    phone_nums = []

    for item in maker.items:
        if item.Web != None:
            url = item.Web
            print("url used: ", url)
            print(type(url))
            try:
                print("trying")
                if url:
                    page = urlopen(url)
                    item.email, item.phoneNum = start_scrape(page)
                else:
                    item.email, item.phoneNum = None, None
            except:
                try:
                    print("trying harder")
                    hdr = {'User-Agent': 'Mozilla/5.0'}

                    if url:
                        req = Request(url, headers=hdr)
                        page = urlopen(req)
                        item.email, item.phoneNum = start_scrape(page)
                    else:
                        emails.append(None)
                        phone_nums.append(None)
                except:
                    print("trying even harder")
                    item.email, item.phoneNum = None, None
        # print("itemName: ",item.Web, item.Email)
        # item.printItem()


    # if save_excel:
    #     name_the_file = input("emails")
    # try:
    #     page = urlopen(url ) 
    #     start_scrape(page)
    # except:
    #     hdr = {'User-Agent': 'Mozilla/5.0'}
    #     req = Request(url, headers=hdr)
    #     page = urlopen(req)
    #     start_scrape(page, "emails")


startScrapeFromFile()
