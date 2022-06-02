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

def findDocument():
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

        return excelDocs[0]

def filterForClasses(courseReport, courses):
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

    return courses

def checkTimeValid(start, end):
    validTimes = [['8:00AM', '9:15AM'], ['9:30AM', '10:45AM'], ['11:00AM', '12:15PM'], ['12:30PM', '1:45PM'], ['2:00PM', '3:15PM'], ['3:30PM', '4:45PM'], ['5:00PM', '6:15PM'], ['6:30PM', '7:45PM'], ['8:00PM', '9:15PM']]

    for times in validTimes:
        if (start.find(times[0]) != -1 and end.find(times[1]) != -1):
            return True

    return False

# def filterTimes(courseReport, courses):
#     #Start Time
#     #End Time

#     for major in courses:
#         for course in courses[major].iterrows():


def main():
    selectedDoc = findDocument()

    courseTemplate = pandas.read_excel(COURSE_GRID_TEMPLATE)
    courseReport = pandas.read_excel(selectedDoc)

    courses = {
        'AV': pandas.DataFrame(columns=courseReport.columns), 
        'CM': pandas.DataFrame(columns=courseReport.columns), 
        'CS': pandas.DataFrame(columns=courseReport.columns), 
        'EG': pandas.DataFrame(columns=courseReport.columns), 
        'IT': pandas.DataFrame(columns=courseReport.columns)
    }

    courseSchedules = {
        'AV': pandas.DataFrame(courseTemplate.values, columns=courseTemplate.columns),
        'CM': pandas.DataFrame(courseTemplate.values, columns=courseTemplate.columns),
        'CS': pandas.DataFrame(courseTemplate.values, columns=courseTemplate.columns),
        'EG': pandas.DataFrame(courseTemplate.values, columns=courseTemplate.columns),
        'IT': pandas.DataFrame(courseTemplate.values, columns=courseTemplate.columns)
    }

    courses = filterForClasses(courseReport, courses)

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
                tempCourse = []
                courseSchedules[major][day] = courseSchedules[major][day].apply(str)

                start = course[1]['Start Time']
                end = course[1]['End Time']
                
                if (type(start) != str or type(end) != str):
                    continue

                if schedule[day] == 'Y':
                    if not checkTimeValid(start, end):
                        dictList = list(COURSE_GRID_POSITIONS[list(schedule).index(day)])
                        periodIndex = dictList.index(str(int(period))) + 1

                        tempCourse = [course[1]['Instructor Last Name'], course[1]['Person Web Email Address'], course[1]['Section Name'], dictList[periodIndex], start, end, course[1]['Meeting Building'], course[1]['Meeting Room'], course[1]['M'], course[1]['T'], course[1]['W'], course[1]['TH'], course[1]['F']]

                    if tempCourse != []:
                        position = COURSE_GRID_POSITIONS[list(schedule).index(day)][str(int(tempCourse[3]))]

                        sectionName = str(tempCourse[2]) if str(tempCourse[2]) != 'nan' else 'TBD'
                        instructorLastName = str(tempCourse[0]) if str(tempCourse[0]) != 'nan' else 'TBD'
                        meetingBuilding = str(tempCourse[6]) if str(tempCourse[6]) != 'nan' else 'TBD'
                        meetingRoom = str(tempCourse[7]) if str(tempCourse[7]) != 'nan' else 'TBD'

                        courseSchedules[major][day][position] += '\n' + sectionName + '-' + instructorLastName + '-' + meetingBuilding + '-' + meetingRoom
                    
                    position = COURSE_GRID_POSITIONS[list(schedule).index(day)][str(int(period))]
                    courseSchedules[major][day][position] += '\n' + str(course[1]['Section Name']) + '-' +  str(course[1]['Instructor Last Name']) + '-' + str(course[1]['Meeting Building']) + '-' + str(course[1]['Meeting Room'])

        courseSchedules[major].to_excel('Schedules/' + major + '_course_schedule.xlsx', sheet_name=major, index=False)

    # filterTimes(courseReport, courses)

main()