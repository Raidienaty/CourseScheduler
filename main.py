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

courses = {'AV': pandas.DataFrame(columns=courseReport.columns), 
            'CM': pandas.DataFrame(columns=courseReport.columns), 
            'CS': pandas.DataFrame(columns=courseReport.columns), 
            'EG': pandas.DataFrame(columns=courseReport.columns), 
            'IT': pandas.DataFrame(columns=courseReport.columns)}

for key, course in courseReport.iterrows():
    # print(courseReport.iloc[key])
    if re.findall(r'\b(AV|AT|AM)', str(courseReport.loc[key]['Course Name']), re.I):
        courses['AV'] = courses['AV'].append(courseReport.iloc[[key]])
    elif re.findall(r'\b(CM|CSM)', str(course), re.I):
        courses['CM'] = courses['CM'].append(courseReport.iloc[[key]])
    elif re.findall(r'\b(CS)', str(course), re.I):
        courses['CS'] = courses['CS'].append(courseReport.iloc[[key]])
    elif re.findall(r'\b(EG|EE)', str(course), re.I):
        courses['EG'] = courses['EG'].append(courseReport.iloc[[key]])
    elif re.findall(r'IT', str(course), re.I):
        courses['IT'] = courses['IT'].append(courseReport.iloc[[key]])




# for column in range(courseTemplate.columns):
#     if column == 'TIME':
#         continue
#     for row in range(courseTemplate.TIME):

