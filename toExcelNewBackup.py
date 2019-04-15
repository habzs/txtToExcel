# import xlsxwriter # Imports libraries required to export to excel
import glob

count = 0
count2 = 0
loopCount = 0
fileCount = 0

firstArray = []
lastArray = []
contactArray = []
birthArray = []
heightArray = []
weightArray = []
combinedInfo = []
toPrintExcelInfo = []


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

    combinedInfo = firstArray + lastArray + contactArray + birthArray + heightArray

    toPrintExcelInfo.insert(fileCount, combinedInfo)
    print("DEBUG toPrintExcelArray: ", toPrintExcelInfo)

def resetVars():
    global firstArray
    global lastArray
    global contactArray
    global birthArray
    global heightArray
    global count

    firstArray = []
    lastArray = []
    contactArray = []
    birthArray = []
    heightArray = []
    count = 0

# --- END FUNCTION --- 

# MAIN LOOP
while (fileCount < 4): # CHANGE THIS TO NUMBER OF FILES!!! 
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


# Test outputs
"""
print (' '.join(firstArray))
print (' '.join(lastArray))
print (' '.join(contactArray))
print (' '.join(birthArray))
print (' '.join(heightArray))
"""