

testArray = [["hazim", "ng", "174"], ["owen", "lee", "299"], ["jie", "xing", "300"]]

testArray2 = ["David", "choo", "778"]

testArray.insert(3, testArray2)
print("TestArray: ", testArray)

numberOfPeople = len(testArray)
numberOfInfo = 3

personCount = 0
infoCount = 0
count = 0

while (personCount < numberOfPeople):

    while (infoCount < numberOfInfo):
        print(testArray[personCount][infoCount])
        infoCount += 1

    infoCount = 0
    personCount += 1
 

