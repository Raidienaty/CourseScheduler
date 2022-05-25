import pandas
import os
import re

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

courses = {'AV': [], 'CM': [], 'CS': [], 'EG': [], 'IT': []}

for course in courseReport['Course Name']:
    if re.findall(r'\b(AV|AT|AM)', str(course), re.I):
        courses['AV'].append(str(course))
    elif re.findall(r'\b(CM|CSM)', course, re.I):
        courses['CM'].append(str(course))
    elif re.findall(r'\b(CS)', course, re.I):
        courses['CS'].append(str(course))
    elif re.findall(r'\b(EG|EE)', course, re.I):
        courses['EG'].append(str(course))
    elif re.findall(r'IT', course, re.I):
        courses['IT'].append(str(course))

print(courses['AV'])


# for column in range(courseTemplate.columns):
#     if column == 'TIME':
#         continue
#     for row in range(courseTemplate.TIME):

