import xlwt
from tempfile import TemporaryFile
import string
 
 
def write_excel(content):
    book = xlwt.Workbook()
    style = xlwt.easyxf(
        'pattern: pattern solid, fore_colour light_blue; ' 
        'font: colour white, bold True;')
    sheet = book.add_sheet("The first sheet")
    row_position = 0
    sheet.write(row_position, 0, "Item in the list", style=style)
    row_position += 1
    for item in content:
        sheet.write(row_position, 0, item)
        row_position += 1
    name = "excel_example.xls"
    book.save(name)
    book.save(TemporaryFile())
    print "done" 

write_excel(['item1', 'item2', 'item3'])