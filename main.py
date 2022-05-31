import pandas
import os
import re
import xlsxwriter

COURSE_GRID_TEMPLATE = 'Course Grid UPDATED.xlsx'
COURSE_GRID_POSITIONS = [
    {'1': 0, '3': 1, '5': 2, '7': 3, '9': 4, '11': 5, '20': 6, '22': 7},
    {'2': 0, '4': 1, '13': 2, '14': 3, '10': 4, '21': 6, '23': 7},
    {'1': 0, '3': 1, '6': 2, '8': 3, '12': 5, '20': 6, '22': 7},
    {'13': 0, '14': 1, '5': 2, '7': 3, '9': 4, '11': 5, '21': 6, '23': 7},
    {'2': 0, '4': 1, '6': 2, '8': 3, '10': 4, '12': 5}
]

def get_col_widths(dataframe):
    # First we find the maximum length of the index column   
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]


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

courseSchedules = {
    'AV': pandas.DataFrame(courseTemplate, columns=courseTemplate.columns),
    'CM': pandas.DataFrame(courseTemplate, columns=courseTemplate.columns),
    'CS': pandas.DataFrame(courseTemplate, columns=courseTemplate.columns),
    'EG': pandas.DataFrame(courseTemplate, columns=courseTemplate.columns),
    'IT': pandas.DataFrame(courseTemplate, columns=courseTemplate.columns)
}

for major in courses:
    for course in courses[major].iterrows():
        period = course[1]['Period']

        schedule = {
            'MONDAY': course[1]['M'], 
            'TUESDAY': course[1]['T'], 
            'WEDNESDAY': course[1]['W'], 
            'THURSDAY': course[1]['TH'], 
            'FRIDAY': course[1]['F']
        }

        for day in schedule:
            courseSchedules[major][day] = courseSchedules[major][day].apply(str)

            if schedule[day] == 'Y':
                position = COURSE_GRID_POSITIONS[list(schedule).index(day)][str(int(period))]

                courseSchedules[major][day][position] += '\n' + str(course[1]['Section Name']) + '-' +  str(course[1]['Instructor Last Name']) + '-' + str(course[1]['Meeting Building']) + '-' + str(course[1]['Meeting Room'])

    # writer = pandas.ExcelWriter('Schedules/' + major + '_course_schedule.xlsx', engine='xlsxwriter')
    # courseSchedules[major].to_excel(writer, sheet_name=major, index=False)

    # workbook = writer.book
    # worksheet = writer.sheets[major]

    # writer.save()

    courseSchedules[major].to_excel('Schedules/' + major + '_course_schedule.xlsx', sheet_name=major, index=False)

    # format = workbook.add_format({'text_wrap': True, 'align': 'top left'})
    # worksheet.set_column(0, len(courseSchedules[major].columns) - 1, cell_format=format)

    # courseSchedules[major].loc[]



# for column in range(courseTemplate.columns):
#     if column == 'TIME':
#         continue
#     for row in range(courseTemplate.TIME):

