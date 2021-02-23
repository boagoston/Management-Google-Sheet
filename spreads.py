import gspread
import math
import time
from oauth2client.service_account import ServiceAccountCredentials

# init variable values
val = []
header = 3
absent_column = 2
p1_column = 3
p2_column = 4
p3_column = 5
situation_column = 6
exam_grade_column = 7
total_grades = 3
total_classes = 60
approved_grade = 70
disapproved_grade = 50


def presence_calc(total, total_student):  # calculates frequency of student and set result if frequency under 25% of total classes

    if total_student > (total * 0.25):
        val.append('Reprovado por Falta')
        val.append(0)


def average_grade(p1, p2, p3): #calculates average grade and set a result based in average result
    average = 0
    exam_final_grade = 0
    if 'Reprovado por Falta' not in val:
        average = (p1 + p2 + p3) / total_grades
        if average >= approved_grade:
            val.append('Aprovado')
            val.append(0)
        elif (average >= disapproved_grade) and (average < approved_grade):
            val.append('Exame Final')
            exam_final_grade = (100 - average)
            val.append(math.ceil(exam_final_grade))
        elif (average < disapproved_grade):
            val.append('Reprovado por Nota')
            val.append(0)


scope = ['https://spreadsheets.google.com/feeds']                                                                       #set default scope
credentials = ServiceAccountCredentials.from_json_keyfile_name('desafiotunts-305619-6e09c376a6a5.json', scope)          #set credentials to authenticate (open json file in path)
gc = gspread.authorize(credentials)                                                                                     # authenticate
wks = gc.open_by_key('1ypmxio4q5EoJz1xQ0WGCNHi35UEWllwg9rpfFf4a1uE')                                                    #open plan
worksheet = wks.get_worksheet(0)                                                                                        #select first page of plan

length_column = len(worksheet.row_values(header))
max_rows = len(worksheet.get_all_values())                                                                              # this is a list of all data and the length is equal to the number of rows # including header row if it exists in data set
x = header + 1
y = situation_column

while x <= max_rows:                                                                                                    #init call functions and send necessary values
    val = worksheet.row_values(x)
    presence_calc(total_classes, int(val[absent_column]))
    average_grade(int(val[p1_column]), int(val[p2_column]), int(val[p3_column]))
    time.sleep(1)                                                                                                       #this line was necessary because Google Sheets API has a write requests per minute per user equal 60 for free user
    while y <= length_column:                                                                                           #update plan
        worksheet.update_cell(x, y, val[y - 1])
        
        y += 1
    x += 1
    y = situation_column
