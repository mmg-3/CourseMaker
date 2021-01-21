import xlsxwriter
import os
import copy
from functools import reduce


def write_courses(courses, worksheet, heading_cell_format):
	worksheet.set_column(0, 0, 60)
	worksheet.set_column(1, 2, 20)
	worksheet.set_column(4, 4, 15)
	worksheet.set_column(5, 5, 10)
	worksheet.write(0, 0, 'Course Name', heading_cell_format)
	worksheet.write(0, 1, 'Professor', heading_cell_format)
	worksheet.write(0, 2, 'Specializations', heading_cell_format)
	worksheet.write(0, 3, 'Credits', heading_cell_format)
	worksheet.write(0, 4, 'Group', heading_cell_format)
	worksheet.write(0, 5, 'Selected', heading_cell_format)
	for index, course in enumerate(courses, 1):
		worksheet.write(index, 0, course['name'])
		worksheet.write(index, 1, course['professor'])
		specializations = ', '.join([spec['letter'] for spec in course['specializations']])
		worksheet.write(index, 2, specializations)
		worksheet.write_number(index, 3, int(course['credits']))
		worksheet.write(index, 4, course['group'])
		worksheet.write_boolean(index, 5, False)


def write_summary(worksheet, specializations):
	worksheet.write(1, 1, 'Total credits')
	worksheet.write(1, 2, '=SUMIF(Courses!F2:F1000; TRUE(); Courses!D2:D1000)')
	worksheet.write(2, 1, 'Core course credits')
	worksheet.write(2, 2, '=SUMIFS(Courses!D2:D1000; Courses!E2:E1000; "Group 1"; Courses!F2:F1000; TRUE())')
	worksheet.write(3, 1, 'Remaining core credits')
	worksheet.write(3, 2, '=MIN(30; MAX(0; 30-C3))')
	worksheet.write(4, 1, 'Intership')
	worksheet.write(4, 2, '=IF(Courses!F2; "Done"; "Not done")')
	worksheet.write(5, 1, 'Semester project')
	worksheet.write(5, 2, '=IF(Courses!F75; "Done"; "Not done")')
	worksheet.write(6, 1, 'SHS')
	worksheet.write(6, 2, '=IF(AND(Courses!F76,Courses!F77), "Done", "Not done")')
	
	worksheet.merge_range('F2:I2', 'Specialisations')
	worksheet.write(2, 5, 'Name')
	worksheet.write(2, 6, 'Letter')
	worksheet.write(2, 7, 'Credits')
	worksheet.write(2, 8, 'Remaining credits')
	for index, (letter, name) in enumerate(specializations):
		index += 3
		worksheet.write(index, 5, name)
		worksheet.write(index, 6, letter)
		worksheet.write(index, 7, f'=SUMIFS(Courses!D2:D1000; Courses!C2:C1000; "*{letter}*"; Courses!F2:F1000; TRUE())')
		worksheet.write(index, 8, f'=MAX(0; MIN(30; 30 - H{index + 1}))')

def create_excel_file(filename, courses):
	try:
		os.remove(filename)
	except:
		pass
	workbook = xlsxwriter.Workbook(filename)
	heading_cell_format = workbook.add_format({'bold': True, 'font_size': 15, 'border': True})

	specializations = set(reduce(lambda a, b: a + b, map(lambda x: list(map(lambda d: tuple(d.values()), x['specializations'])), courses), []))

	write_courses(courses, workbook.add_worksheet('Courses'), heading_cell_format)
	write_summary(workbook.add_worksheet('Summary'), specializations)

	workbook.close()
