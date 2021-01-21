from excel_creator import create_excel_file
from course_getter import get_courses
import sys

DEFAULT_OUTPUT_FILE = 'courses.xlsx'

if __name__ == "__main__":
	output_file = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_OUTPUT_FILE
	COURSES_URL = 'https://edu.epfl.ch/studyplan/en/master/computer-science'
	create_excel_file(output_file, get_courses(COURSES_URL))
