import xlwt
import xlrd
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
    print "Wrote to excel_example.xls successfully"

def read_excel():
    ''' Tutorial link at 
    http://www.youlikeprogramming.com/2012/03/examples-reading-excel-xls-documents-using-pythons-xlrd/ '''
    workbook = xlrd.open_workbook('my_workbook.xls')
    worksheet = workbook.sheet_by_name('Sheet1')
    num_rows = worksheet.nrows - 1
    num_cells = worksheet.ncols - 1
    print "num rows:", num_rows
    print "num cells:", num_cells
    curr_row = -1
    while curr_row < num_rows:
        curr_row += 1
        row = worksheet.row(curr_row)
        print 'Row:', curr_row
        curr_cell = -1
        while curr_cell < num_cells:
            curr_cell += 1
            # Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
            cell_type = worksheet.cell_type(curr_row, curr_cell)
            cell_value = worksheet.cell_value(curr_row, curr_cell)
            print ' ', cell_type, ':', cell_value

write_excel(['item1', 'item2', 'item3'])
read_excel()