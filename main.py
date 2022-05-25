from itertools import count
import pandas
import os

#path =  os.path.realpath(__file__)
directoryList = os.listdir('.')

print('Which file would you like to use?')

excelDocs = []

for item in directoryList:
    if item.find('Course Grid UPDATED.xlsx') != -1:
        continue
    elif item.find('.xlsx') != -1:
        excelDocs.append(item)


if len(excelDocs) < 1:
    print('Please move the course report to my folder!')
    exit()
elif len(excelDocs) > 1:
    print('Which of the following files would you like to use?')
    for item in directoryList:
        if item.count('.xlsx'):
            print(item)
else:
    print('Using ')
    for item in directoryList: 
        if item.count('*.xlsx'):
            print(item)