import gspread
import math
import time
import logging
from oauth2client.service_account import ServiceAccountCredentials

# init variable values
student_info = []
header = 3
name_column = 1
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
logging.basicConfig(filename='example.log',level=logging.INFO, filemode='w')

def presence_calc(total, total_student):  # calculates frequency of student and set result if frequency under 25% of total classes

    if total_student > (total * 0.25):
        student_info.append('Reprovado por Falta')
        student_info.append(0)


def average_grade(p1, p2, p3): #calculates average grade and set a result based in average result
    average = 0
    exam_final_grade = 0
    if 'Reprovado por Falta' not in student_info:
        average = (p1 + p2 + p3) / total_grades
        if average >= approved_grade:
            student_info.append('Aprovado')
            student_info.append(0)
        elif (average >= disapproved_grade) and (average < approved_grade):
            student_info.append('Exame Final')
            exam_final_grade = (100 - average)
            student_info.append(math.ceil(exam_final_grade))
        elif (average < disapproved_grade):
            student_info.append('Reprovado por Nota')
            student_info.append(0)

logging.info('Programm started')
scope = ['https://spreadsheets.google.com/feeds']                                                                       #set default scope
credentials = ServiceAccountCredentials.from_json_keyfile_name('desafiotunts-305619-6e09c376a6a5.json', scope)          #set credentials to authenticate (open json file in path)
gc = gspread.authorize(credentials)                                                                                     # authenticate
wks = gc.open_by_key('1ypmxio4q5EoJz1xQ0WGCNHi35UEWllwg9rpfFf4a1uE')                                                    #open plan
worksheet = wks.get_worksheet(0)                                                                                        #select first page of plan
logging.info('Plan loaded ')

length_column = len(worksheet.row_values(header))
max_rows = len(worksheet.get_all_values())                                                                              # this is a list of all data and the length is equal to the number of rows # including header row if it exists in data set
x = header + 1
y = situation_column

while x <= max_rows:                                                                                                    #init call functions and send necessary values
    student_info = worksheet.row_values(x)
    presence_calc(total_classes, int(student_info[absent_column]))
    average_grade(int(student_info[p1_column]), int(student_info[p2_column]), int(student_info[p3_column]))
    logging.info(student_info[name_column] + ' info load in list')
    time.sleep(2)                                                                                                       #this line was necessary because Google Sheets API has a write requests per minute per user equal 60 for free user
    while y <= length_column:                                                                                           #update plan
        worksheet.update_cell(x, y, student_info[y - 1])
        y += 1
    logging.info(student_info[name_column] + ' updated in Google Sheets Plan')                
    logging.info(student_info)
    x += 1
    y = situation_column
