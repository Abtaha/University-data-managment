from pyexcel_ods import get_data
import json
from collections import defaultdict

files = ['Data1', 'Data2', 'Data3', 'Data4', 'IELTS', 'Interview']

stData = defaultdict(lambda:[])

for file in files:
    fileData = json.dumps(get_data(f"Data Task/{file}.ods"))
    fileData = json.loads(fileData)
    fileData = fileData['Sheet1']

    length = len([x for x in fileData if x != []]) - 3

    for i in range(3, length):
        individual = fileData[i]
        stData[individual[0]].append([file, individual[1:]])


result = defaultdict(lambda:[])
for data in stData.items():
    for examData in data[1]:
        if examData[0] == "Data1":
            for i in range(len(examData[1])):
                examData[1][i] = examData[1][i] / 100
            
            examData[1][0] = examData[1][0] * 2
            examData[1][1] = examData[1][1] * 2
            examData[1][4] = examData[1][4] * 2
            
            result[data[0]].append(((sum(examData[1]) / len(examData[1])) * 0.1))
        
        elif examData[0] == "Data2":
            for i in range(len(examData[1])):
                examData[1][i] = examData[1][i] / 100
            
            examData[1][0] = examData[1][0] * 2
            examData[1][1] = examData[1][1] * 2
            examData[1][4] = examData[1][4] * 2
            
            result[data[0]].append(((sum(examData[1]) / len(examData[1])) * 0.1))
        
        elif examData[0] == "Data3":
            for i in range(len(examData[1])):
                examData[1][i] = examData[1][i] / 100
            
            examData[1][0] = examData[1][0] * 2
            examData[1][1] = examData[1][1] * 2
            examData[1][4] = examData[1][4] * 2
            
            result[data[0]].append(((sum(examData[1]) / len(examData[1])) * 0.1))
        
        elif examData[0] == "Data4":
            for i in range(len(examData[1])):
                examData[1][i] = examData[1][i] / 100
            
            examData[1][0] = examData[1][0] * 2
            examData[1][1] = examData[1][1] * 2
            examData[1][4] = examData[1][4] * 2
            
            result[data[0]].append(((sum(examData[1]) / len(examData[1])) * 0.1))
        
        elif examData[0] == "IELTS":
            for i in range(len(examData[1])):
                examData[1][i] = examData[1][i] / 9
            
            result[data[0]].append(((sum(examData[1]) / len(examData[1])) * 0.3))
        
        elif examData[0] == "Interview":
            for i in range(len(examData[1])):
                examData[1][i] = examData[1][i] / 10
            
            result[data[0]].append(((sum(examData[1]) / len(examData[1])) * 0.3))

totalResult = []
for res in result.items():
    totalResult.append([res[0], sum(res[1])])

totalResult.sort(key = lambda x: x[1], reverse=True)


f = open("results.txt", "w+")

for line in totalResult:
    f.write("Name: " + line[0] + "\r\n" + "Score: " + str((line[1] * 100)) + " %" + "\r\n\r\n")

f.close()