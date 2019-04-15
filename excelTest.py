##############################################################################
#
# A simple example of some of the features of the XlsxWriter Python module.
#
# Copyright 2013-2019, John McNamara, jmcnamara@cpan.org
#
import xlsxwriter


# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 20)
worksheet.set_column('B:A', 20)
worksheet.set_column('C:A', 20)
worksheet.set_column('D:A', 20)
worksheet.set_column('E:A', 20)
worksheet.set_column('F:A', 20)
worksheet.set_column('G:A', 20)


# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Write some simple text.
worksheet.write('A1', 'No.')
worksheet.write('B1', 'First Name')
worksheet.write('C1', 'Last Name')
worksheet.write('D1', 'Phone Num.')
worksheet.write('E1', 'DoB')
worksheet.write('F1', 'Height')
worksheet.write('G1', 'Weight')

# Text with formatting.
worksheet.write('A2', 'World', bold)

# Write some numbers, with row/column notation.
worksheet.write(2, 0, 123)
worksheet.write(3, 0, 123.456)
worksheet.write(2, 1, 456)


# Insert an image.
#worksheet.insert_image('B5', 'logo.png')

column = 0
printCount = 0
info = ["Owen", "Lee", "82883343", "05/10/1997", "174"]
info2 = ["Hello", "Test", "2"]
infoC = info + info2

arrayLength = len(infoC)

while (printCount < arrayLength):

    toPrint = infoC[printCount]

    worksheet.write(0, column, toPrint)

    printCount += 1
    column += 1

workbook.close()