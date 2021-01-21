import requests
from bs4 import BeautifulSoup


def first(iterable):
	try:
		return next(iter(iterable))
	except:
		return None

def extract_course_info(soup, groups):
	group_name, group_val = 'Unknown', -1
	for g_name, g_line in groups.items():
		if g_line > group_val and g_line <= soup.sourceline:
			group_name, group_val = g_name, g_line
	group_name = group_name.strip()

	course_name = first(soup.select('div.cours-name a')).text.strip()
	credits = first(soup.select('div.credit-time')).text.strip()
	specializations = soup.select('div.specialisation img')
	specialization_data = [
		{
			'letter': specialization['src'][-5].upper().strip(),
			'name': specialization['title'].strip()
		}\
		for specialization in specializations
	]
	professor_link = first(soup.select('div.enseignement-name a'))
	professor = professor_link.text.strip() if professor_link is not None else ''
	return {
		'name': course_name,
		'credits': credits,
		'specializations': specialization_data,
		'professor': professor,
		'group': group_name
	}


def get_courses(url):
	#COURSES_URL = "https://edu.epfl.ch/studyplan/en/master/computer-science"
	page = BeautifulSoup(requests.get(url).text, 'html.parser')


	course_containers = page.select('#content div.line-down div.line')
	groups = {soup.text: soup.sourceline for soup in page.select('#content h4')}
	return [extract_course_info(course, groups) for course in course_containers]

