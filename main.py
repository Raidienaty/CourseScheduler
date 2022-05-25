import pandas
import os

COURSE_GRID_TEMPLATE = 'Course Grid UPDATED.xlsx'

#path =  os.path.realpath(__file__)
directoryList = os.listdir('.')

print('Which file would you like to use?')

excelDocs = []

for item in directoryList:
    if item.find(COURSE_GRID_TEMPLATE) != -1:
        continue
    elif item.find('.xlsx') != -1:
        excelDocs.append(item)

selectedDoc = ''

if len(excelDocs) < 1:
    print('Please move the course report to my folder!')
    exit()
elif len(excelDocs) > 1:
    print('Which of the following files would you like to use?')
    for item in excelDocs:
        print(item)
else:
    print('Using: ')
    for item in excelDocs: 
        print(item)

    selectedDoc = excelDocs[0]

courseTemplate = pandas.read_excel(COURSE_GRID_TEMPLATE)
courseReport = pandas.read_excel(selectedDoc)

print(courseReport.head(10))
print()
print(courseTemplate.head(10))

