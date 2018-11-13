import pandas as pd
from openpyxl import load_workbook

wb = load_workbook('./data.xlsx')
sheetArr =wb.sheetnames

banDevi = wb[sheetArr[4]]
print('Working Sheet', wb[sheetArr[4]])
print(banDevi['C114'].value)

class Student:
    def __init__(self, name, age, grade, gender):
        self.name = name
        self.age = age
        self.grade = grade
        self.gender = gender

nameCount = 2
ageCount = 5
baseCount = 6
students = []
i = 0
while banDevi['C'+str(nameCount)].value is not None:
    if banDevi['D'+str(ageCount)].value is None:
        gender = 'M'
    else:
        gender = 'F'
    student = Student(banDevi['C'+str(nameCount)].value,banDevi['C'+str(ageCount)].value,banDevi['B'+str(ageCount)].value,gender)
    students.append(student)
    nameCount += 8
    ageCount += 8
    startChar = 'F'
    for j in range(5):
        #for litnum Nepali
        print("timeline: ", banDevi[startChar+str(baseCount)].value)
        print("value try", startChar, banDevi[startChar+str(baseCount)].value)
        if banDevi[startChar+str(baseCount)].value is not None:
            litNepali = j+1
            print("lit marks", litNepali, j)
            break
        startChar = chr(ord(startChar)+1)


    baseCount += 8
    print("hello", students[i].name, students[i].age, students[i].grade, students[i].gender)
    i += 1
