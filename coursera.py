from lxml import etree
import requests
from bs4 import BeautifulSoup
import json
from openpyxl import Workbook


def get_courses_list(file_path):
    count_of_courses = 20
    namespace = {'sitemap': 'http://www.sitemaps.org/schemas/sitemap/0.9'}
    tree = etree.parse(file_path)
    courses = tree.getroot()
    links_of_course = []
    for course in courses.findall('sitemap:url', namespace):
        if (len(links_of_course) < count_of_courses):
            links_of_course.append(course.find('sitemap:loc', namespace).text)
        else:
            break
    return links_of_course


def get_course_info(course_slug):
    course_info = {}
    try:
        r = requests.get(course_slug)
        soup = BeautifulSoup(r.content, "lxml")
    except requests.exceptions.ConnectionError:
        get_course_info(course_slug)
        return
    try:
        for span in soup.find_all('div', "ratings-text bt3-hidden-xs")[0]:
            course_info['rating'] = span.text
            break
    except IndexError:
        course_info['rating'] = ''
    try:
        course_info['title'] = soup.find('div', "title display-3-text").string
    except AttributeError:
        course_info['title'] = ''
    try:
        course_info['weeks'] = soup.find_all(
            'div', "week-heading body-2-text"
        )[-1].string
    except IndexError:
        course_info['weeks'] = ''
    try:
        course_info['date'] = json.loads(
            soup.find('div', "rc-CourseGoogleSchemaMarkup").script.text
        )['hasCourseInstance'][0]['startDate']
        course_info['language'] = json.loads(
            soup.find('div', "rc-CourseGoogleSchemaMarkup").script.text
        )['hasCourseInstance'][0]['inLanguage']
    except (AttributeError, KeyError):
        course_info['date'] = ''
        course_info['language'] = ''
    course_info['url'] = course_slug
    return course_info


def output_courses_info_to_xlsx(filepath, courses_info):
    wb = Workbook()
    ws = wb.active
    for info in courses_info:
        if info is not None:
            ws.append(
                [info['url'],
                 info['title'],
                 info['language'],
                 info['date'],
                 info['weeks'],
                 info['rating']]
            )
    wb.save(filepath)

if __name__ == '__main__':
    file_to_parse = 'courses.xml'
    file_to_save = 'courses.xlsx'
    courses_information = []
    courses_links = get_courses_list(file_to_parse)
    for course in courses_links:
        courses_information.append(get_course_info(course))
    output_courses_info_to_xlsx(file_to_save, courses_information)
