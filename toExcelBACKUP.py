import xlsxwriter # Imports libraries required to export to excel

count = 0
count2 = 0
loopCount = 0
firstArray = []
lastArray = []
contactArray = []
birthArray = []
heightArray = []
weightArray = []

f = open("demofile.txt", "r")
doc = f.read()
docArray = doc.split()

totalLength = len(docArray)        
#print(totalLength)

# FUNCTIONS

def firstName():
    global count
    global count2

    if (docArray[count] == "First"):
        loopCount = count + 2

        while (count2 < 1):
            firstArray.append(docArray[loopCount])
            count2 += 1
            loopCount += 1
    
    count2 = 0
    loopCount = 0

def lastName():
    global count
    global count2

    if (docArray[count] == "Last"):
        loopCount = count + 2

        while (count2 < 1):
            lastArray.append(docArray[loopCount])
            count2 += 1
            loopCount += 1

    count2 = 0
    loopCount = 0

def contactNumber():
    global count
    global count2

    if (docArray[count] == "Contact"):
        loopCount = count + 3

        while (count2 < 1):
            contactArray.append(docArray[loopCount])
            count2 += 1
            loopCount += 1

    count2 = 0
    loopCount = 0

def dateOfBirth():
    global count
    global count2

    if (docArray[count] == "Date"):
        loopCount = count + 3

        while (count2 < 1):
            birthArray.append(docArray[loopCount])
            count2 += 1
            loopCount += 1

    count2 = 0
    loopCount = 0

def height():
    global count
    global count2

    if (docArray[count] == "Height"):
        loopCount = count + 2

        while (count2 < 1):
            heightArray.append(docArray[loopCount])
            count2 += 1
            loopCount += 1

    count2 = 0
    loopCount = 0

def weight():
    global count
    global count2

    if (docArray[count] == "Weight"):
        loopCount = count + 2

        while (count2 < 1):
            weightArray.append(docArray[loopCount])
            count2 += 1
            loopCount += 1

    count2 = 0
    loopCount = 0

# MAIN LOOP

while (count < totalLength):

    firstName()
    lastName()
    contactNumber()
    dateOfBirth()
    height()
    weight()

    print(count)
    count += 1


# Excel export functions

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

# Write some text.
worksheet.write('A1', 'No.')
worksheet.write('B1', 'First Name')
worksheet.write('C1', 'Last Name')
worksheet.write('D1', 'Phone Num.')
worksheet.write('E1', 'DoB')
worksheet.write('F1', 'Height')
worksheet.write('G1', 'Weight')

row = 1
column = 1
printCount = 0
combinedInfo = firstArray + lastArray + contactArray + birthArray + heightArray # Makes for easier output to excel
testCount = 0
combinedInfoLength = len(combinedInfo)

while (printCount < combinedInfoLength):

    toPrint = combinedInfo[printCount]

    # Row, Column, Item
    worksheet.write(row, column, toPrint)

    printCount += 1
    column += 1

workbook.close()

# Test outputs
print (' '.join(firstArray))
print (' '.join(lastArray))
print (' '.join(contactArray))
print (' '.join(birthArray))
print (' '.join(heightArray))
