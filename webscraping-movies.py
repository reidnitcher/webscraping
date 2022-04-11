
from os import readlink
from platform import release
from re import M
from sre_constants import GROUPREF_EXISTS
from tkinter import font
from turtle import onkey
from urllib.request import urlopen
from bs4 import BeautifulSoup
import openpyxl as xl
from openpyxl.styles import Font

from excel1 import MySheet
import openpyxl as xl
from openpyxl.styles import Font





#webpage = 'https://www.boxofficemojo.com/weekend/chart/'
webpage = 'https://www.boxofficemojo.com/year/2022/'

page = urlopen(webpage)			

soup = BeautifulSoup(page, 'html.parser')

title = soup.title
wb = xl.Workbook()

MySheet = wb.active
wb.create_sheet(index=1, title = 'MOVIE SHEET')

MySheet.title = 'Movie Gross'

MySheet.title = 'First Sheet'
print(title.text)
movie_table = soup.find('table')

rows = movie_table.findAll('tr')
fontobject = Font(name='Times New Roman', size=24, italic=True, bold=True)


MySheet['A1'] = 'No.'
MySheet['A1'].font = fontobject

MySheet['B1'] = 'Movie Title'
MySheet['B1'].font = fontobject

MySheet['C1'] = 'Release Date'
MySheet['C1'].font = fontobject

MySheet['D1'] = 'Gross'
MySheet['D1'].font = fontobject

MySheet['E1'] = 'Total Gross'
MySheet['E1'].font = fontobject

MySheet['F1'] = "% of total gross"
MySheet['F1'].font = fontobject

rows = movie_table.findAll('tr')


for x in range(1,6):
    td = rows[x].findAll('td')
    ranking = td[0].text
    title = td[1].text
    gross = int(td[5].text.replace(",","").replace("$",""))
    total_gross = int(td[7].text.replace(",","").replace("$",""))
    release_date = td[8].text

    percent_gross = round((gross/total_gross)*100,2)

    MySheet['A' + str(x+1)] = ranking
    MySheet['B' + str(x+1)] = title
    MySheet['C' + str(x+1)] = release_date
    MySheet['D' + str(x+1)] = gross
    MySheet['E' + str(x+1)] = total_gross
    MySheet['F' + str(x+1)] = str(percent_gross) + '%'


    
MySheet.column_dimensions['A'].width = 5
MySheet.column_dimensions['B'].width = 30
MySheet.column_dimensions['C'].width = 25
MySheet.column_dimensions['D'].width = 16
MySheet.column_dimensions['E'].width = 20
MySheet.column_dimensions['F'].width = 26

header_font = Font(size=16, bold=True)

for cell in MySheet[1:1]:
    cell.font = header_font

wb.save('BoxOfficeReport.xlsx')

##
##
##

