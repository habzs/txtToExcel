import glob
fileCount = 0

dirname = "./testfiles"

gDirname = glob.glob(dirname+'/*')

length = len(gDirname)
print ("items: ", gDirname)
print ("length: ", length)

while (fileCount < length):
    toPrint = gDirname[fileCount]
    print("Current file selected: ", toPrint)

    fileCount += 1