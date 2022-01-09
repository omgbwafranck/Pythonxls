import xlwt
import xlrd
from xlutils.copy import copy

# load the excel file
rb = xlrd.open_workbook('/Users/air/Documents/employeedata(Original).xls')

# copy the content of excel file
wb = copy(rb)

# open the first sheet
w_sheet = wb.get_sheet(0)

# Let's implement row and column number

# row number = 1 column number = 1
w_sheet.write(1,1, 'Karl@handsinhands.org')

# row number = 2 column = 1
w_sheet.write(2,1, 'Mike@handsinhands.org')

# row number = 3 column number = 1
w_sheet.write(3,1, 'Audrey@handsinhands.org')

# row number = 4 column number = 1
w_sheet.write(4,1, 'Franck@handsinhands.org')

# row number = 5 column number = 1
w_sheet.write(5,1, 'Lucianne@handsinhands.org')

# row number = 6 column number = 1
w_sheet.write(6,1, 'Grace@handsinhands.org')

# row number = 7 column number = 1
w_sheet.write(7,1, 'Lyse@handsinhands.org')

# row number = 8 column number = 1
w_sheet.write(8,1, 'Arthur@handsinhands.org')

# row number = 9 column number = 1
w_sheet.write(9,1, 'Vicky@handsinhands.org')

# row number = 10 column number = 1
w_sheet.write(10,1, 'Leandra@handsinhands.org')

# row number = 11 column number = 1
w_sheet.write(11,1, 'Claire@handsinhands.org')

# row number = 12 column number = 1
w_sheet.write(12,1, 'Simon@handsinhands.org')

# row number = 13 column number = 1
w_sheet.write(13,1, 'Hilary@handsinhands.org')

# row number = 14 column number = 1
w_sheet.write(14,1, 'Aurel@handsinhands.org')

# row number = 15 column number = 1
w_sheet.write(15,1, 'Eric@handsinhands.org')

# row number = 16 column number = 1
w_sheet.write(16,1, 'Jean@handsinhands.org')

# row number = 17 column number = 1
w_sheet.write(17,1, 'Giselle@handsinhands.org')

# row number = 18 column number = 1
w_sheet.write(18,1, 'Ariol@handsinhands.org')

# row number = 19 column number = 1
w_sheet.write(19,1, 'Megane@handsinhands.org')

# row number = 20 column number = 1
w_sheet.write(20,1, 'Celeste@handsinhands.org')

# row number = 21 column number = 1
w_sheet.write(21,1, 'Doriante@handsinhands.org')

# row number = 22 column number = 1
w_sheet.write(22,1, 'Claude@handsinhands.org')

# row number = 23 column number = 1
w_sheet.write(23,1, 'Michelle@handsinhands.org')

# row number = 24 column number = 1
w_sheet.write(24,1, 'Jessica@handsinhands.org')

# row number = 25 column number = 1
w_sheet.write(25,1, 'Arnold@handsinhands.org')

# row number = 26 column number = 1
w_sheet.write(26,1, 'Shawn@handsinhands.org')

# row number = 27 column number = 1
w_sheet.write(27,1, 'James@handsinhands.org')

# row number = 28 column number = 1
w_sheet.write(28,1, 'Larry@handsinhands.org')

# row number = 29 column number = 1
w_sheet.write(29,1, 'Loic@handsinhands.org')

# saving the file
wb.save('employeedata.xls')
