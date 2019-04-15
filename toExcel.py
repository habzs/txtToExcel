#!/usr/bin/env python3

import xlsxwriter # Imports libraries required to export to excel
import glob

count = 0
count2 = 0
loopCount = 0
fileCount = 0
toPrintExcelInfoLength = 0

firstArray = []
lastArray = []
contactArray = []
birthArray = []
heightArray = []
weightArray = []
combinedInfo = []
toPrintExcelInfo = []

# Initiates excel workbook/sheet
workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})

cell_format = workbook.add_format()
cell_format.set_align('center')

# Widen the first column to make the text clearer.
#worksheet.set_column('A:A', 20)
worksheet.set_column('B:A', 20)
worksheet.set_column('C:A', 20)
worksheet.set_column('D:A', 20)
worksheet.set_column('E:A', 20)
worksheet.set_column('F:A', 20)
worksheet.set_column('G:A', 20)

# Write some text.
worksheet.write('A1', 'No.', bold)
worksheet.write('B1', 'First Name', bold)
worksheet.write('C1', 'Last Name', bold)
worksheet.write('D1', 'Phone Num.', bold)
worksheet.write('E1', 'DoB', bold)
worksheet.write('F1', 'Height', bold)
worksheet.write('G1', 'Weight', bold)

# Opens folders and checks for the number of files to be opened and places them into an array
dirname = "./testfiles"
gDirname = glob.glob(dirname+'/*')
numOfFiles = len(gDirname)

# FUNCTIONS
def firstName():
    global count
    global count2

    if (docToArray[count] == "First"):
        loopCount = count + 2

        while (count2 < 1):
            firstArray.append(docToArray[loopCount])
            count2 += 1
            loopCount += 1
    
    count2 = 0
    loopCount = 0

def lastName():
    global count
    global count2

    if (docToArray[count] == "Last"):
        loopCount = count + 2

        while (count2 < 1):
            lastArray.append(docToArray[loopCount])
            count2 += 1
            loopCount += 1

    count2 = 0
    loopCount = 0

def contactNumber():
    global count
    global count2

    if (docToArray[count] == "Contact"):
        loopCount = count + 3

        while (count2 < 1):
            contactArray.append(docToArray[loopCount])
            count2 += 1
            loopCount += 1

    count2 = 0
    loopCount = 0

def dateOfBirth():
    global count
    global count2

    if (docToArray[count] == "Date"):
        loopCount = count + 3

        while (count2 < 1):
            birthArray.append(docToArray[loopCount])
            count2 += 1
            loopCount += 1

    count2 = 0
    loopCount = 0

def height():
    global count
    global count2

    if (docToArray[count] == "Height"):
        loopCount = count + 2

        while (count2 < 1):
            heightArray.append(docToArray[loopCount])
            count2 += 1
            loopCount += 1

    count2 = 0
    loopCount = 0

def weight():
    global count
    global count2

    if (docToArray[count] == "Weight"):
        loopCount = count + 2

        while (count2 < 1):
            weightArray.append(docToArray[loopCount])
            count2 += 1
            loopCount += 1

    count2 = 0
    loopCount = 0


def combinedUsersInfo():
    global toPrintExcelInfoLength

    combinedInfo = firstArray + lastArray + contactArray + birthArray + heightArray + weightArray

    toPrintExcelInfo.insert(fileCount, combinedInfo)
    toPrintExcelInfoLength = len(toPrintExcelInfo[0])
    print("DEBUG toPrintExcelArray: ", toPrintExcelInfo)
    print("DEGUB toPrintExcelArrayLength: ", toPrintExcelInfoLength)

def resetVars():
    global firstArray
    global lastArray
    global contactArray
    global birthArray
    global heightArray
    global weightArray
    global count

    firstArray = []
    lastArray = []
    contactArray = []
    birthArray = []
    heightArray = []
    weightArray = []
    count = 0

# --- END FUNCTION --- 

# MAIN LOOP
while (fileCount < numOfFiles): # CHANGE THIS TO NUMBER OF FILES!!! 
    currentFile = gDirname[fileCount]
    print("DEBUG Current file selected: ", currentFile)
    f = open(currentFile)
    doc = f.read()
    docToArray = doc.split()

    totalDocLength = len(docToArray)
    print("DEBUG Total Length: ", totalDocLength)
    #print("DEBUG Output: ", docToArray)

    # MAIN LOOP
    while (count < totalDocLength):

        firstName()
        lastName()
        contactNumber()
        dateOfBirth()
        height()
        weight()

        #print(count)
        count += 1

    combinedUsersInfo()
    fileCount += 1

    # Resets arrays and counters to prepare it for the next text file
    resetVars()


# Program to write arrays to excel sheet
row = 1
column = 1
currentUser = 0
printCount = 0

while (currentUser < numOfFiles):
    while (printCount < toPrintExcelInfoLength):
        print("currentUser: ", currentUser)
        print("printCount: ", printCount)

        toPrint = toPrintExcelInfo[currentUser][printCount]
        print ("Debug toOUTPUTTOEXCEL: ", toPrint)

        worksheet.write(row, column, toPrint)
        worksheet.write(row, 0, currentUser+1)


        printCount += 1
        column += 1
    
    column = 1
    printCount = 0
    currentUser += 1
    row += 1



workbook.close()

# Test outputs
"""
print (' '.join(firstArray))
print (' '.join(lastArray))
print (' '.join(contactArray))
print (' '.join(birthArray))
print (' '.join(heightArray))
"""