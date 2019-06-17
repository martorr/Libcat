import openpyxl
import requests
from bs4 import BeautifulSoup

# bs4 is used to parse the XML file from the Search API

# Open and load workbook/worksheet
workbookName = 'template.xlsx'
workbook = openpyxl.load_workbook(workbookName)
worksheet = workbook['Sheet1']

row = 2
title = ""
author = ""
edition = ""
response = 'y'
while response == 'y':
    # requests info from OCLC WorldCat Search API
    book_ISBN = input("Book ISBN?: ")
    website_url = requests.get('http://classify.oclc.org/classify2/Classify?isbn=' + book_ISBN + '&summary=true').text
    soup = BeautifulSoup(website_url, 'lxml')

    # Find elements in XML file and output the attribute value within <work> tag
    for element in soup.find_all("classify"):
        for stat in element.find_all("work"):
            title = stat['title']
            print(title)
            author = stat['author']
            print(author)
            edition = stat['editions']
            print(edition)

    # TEST ISBN's
    # 978-0-205-25949-6
    # 978-1-2851990-2-3
    # 1782197567

    # input information from API


    course_num = input("Course number?: ")

    worksheet.cell(row=row, column=1).value = course_num
    worksheet.cell(row=row, column=2).value = title
    worksheet.cell(row=row, column=3).value = author
    worksheet.cell(row=row, column=4).value = edition
    worksheet.cell(row=row, column=5).value = book_ISBN

    response = input("Continue? (Y/N): ")
    if response == "N":
        break
    row += 1


workbook.save(workbookName)
workbook.close

