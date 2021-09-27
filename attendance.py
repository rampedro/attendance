import xlrd, xlwt
from xlutils.copy import copy
from termcolor import colored

read_book = xlrd.open_workbook("./Lab_004.xlsx") #Make Readable Copy
write_book = copy(read_book) #Make Writeable Copy

write_sheet1 = write_book.get_sheet(0) #Get sheet 1 in writeable copy
#print(read_book.cell(1,1).value)



import xlrd

# Give the location of the file

# To open Workbook
loc = "./Lab_004.xlsx"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# For row 0 and column 0
print(sheet.cell_value(0, 0))



while True:
    ins = input()
    for item in range(0,41):
        idd = sheet.cell_value(item,0)
        if str(idd) == str(ins):

            #print("yes")
            write_sheet1.write(item,6, 'Present') #Write 'test' to cell (1, 11)
            print(colored("Succefully attended",'green'))

            write_book.save("./lab3-attendance-done.xlsx") #Save the newly written copy. Enter the same as the old path to write over

            #print("no")
        
#write_sheet2 = write_book.get_sheet(2) #Get sheet 2 in writeable copy
#write_sheet2.write(3, 14, '135') #Write '135' to cell (3, 14)

