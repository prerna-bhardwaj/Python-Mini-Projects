"""
Scrapes relevant courses from online sites like Udemy, Coursera, Eduonix, Udacity, etc.
Saves scraped data in an excel sheet in readable format. 

"""

from selenium import webdriver
import pandas as pd
import xlsxwriter
import urllib.request
import time
import io
import os

currDir = os.getcwd()
PATH=os.path.join(currDir, "chromedriver")
course = ''
coursera_url = 'https://www.coursera.org/search?query={course}'
udemy_url = 'https://www.udemy.com/courses/search/?src=ukw&q={course}'


def scrapeUdemy(course, driver):
	driver.get(udemy_url.format(course=course))
	print("Sraping data from Udemy ...")
	time.sleep(5)
	udemy_df = pd.DataFrame(columns=['Image', 'Title', 'Instructors', 'Description', 'Course Time', 'Rating', 'Rated by', 'Price', 'Link'])

	count = 0
	cards = driver.find_elements_by_class_name('course-card--container--1QM2W')
	links = driver.find_elements_by_class_name('browse-course-card--link--3KIkQ')
	for index, card in enumerate(cards):
		image = card.find_elements_by_class_name('browse-course-card--image--35hYN')
		title = card.find_elements_by_class_name('course-card--course-title--vVEjC')
		instructors = card.find_elements_by_class_name('course-card--instructor-list--nH1OC')
		description = card.find_elements_by_class_name('course-card--course-headline--2DAqq')
		courseTime = card.find_elements_by_class_name('course-card--course-meta-info--2jTzN')
		rating = card.find_elements_by_class_name('star-rating--rating-number--2o8YM')
		price = card.find_elements_by_class_name('course-card--discount-price--1bQ5Q')
		rated_by = card.find_elements_by_class_name('course-card--reviews-text--1yloi')
		image_src = image[0].get_attribute('src')
		image_ext = image_src.split('.')[-1]
		if image_ext in ['jpg', 'jpeg', 'png']:
			image_name = 'udemy_{index}.{ext}'.format(index=count, ext=image_ext)
			urllib.request.urlretrieve(image_src, image_name)
		else:
			image_name='dummy.jpg'

		try:
			print(title[0].text)
			obj = {
				'Image': image_name,
				'Title': title[0].text,
				'Instructors': instructors[0].text,
				'Description': description[0].text,
				'Course Time': courseTime[0].text.replace('\n', ' | '),
				'Rating': rating[0].text,
				'Rated by': rated_by[0].text[1:-1],
				'Price': price[0].text.split('\n')[1],
				'Link': links[index].get_attribute('href')
			}
		except:
			continue
		count += 1
		udemy_df = udemy_df.append(obj, ignore_index=True)

	print("Data fetched from Udemy !!")
	return udemy_df


def scrapeCourse(course, driver):
	driver.get(coursera_url.format(course=course))							
	print("\nSraping data from Coursera ...")
	time.sleep(3)							
	coursera_data = pd.DataFrame(columns=['Image', 'Title', 'Creator', 'Type', 'Enrolled', 'Difficulty', 'Rating', 'Rated by', 'Link'])							
								
	cards = driver.find_elements_by_class_name('ais-InfiniteHits-item')							
	print(len(cards))							
	for count, card in enumerate(cards):
		title = card.find_elements_by_class_name('headline-1-text')						
		creator = card.find_elements_by_class_name('partner-name')						
		courseType = card.find_elements_by_class_name('_jen3vs')						
		difficulty = card.find_elements_by_class_name('difficulty')						
		enrolled = card.find_elements_by_class_name('enrollment-number')						
		rated_by = card.find_elements_by_class_name('ratings-count')						
		rating = card.find_elements_by_class_name('ratings-text')				
		# link = getCourseLink(title[0].text, courseType[0].text)

		image = card.find_elements_by_class_name('image-wrapper')						
		# print(image[0].find_elements_by_tag_name('img')[0].get_attribute('src'))
		image_link = image[0].find_elements_by_tag_name('img')[0].get_attribute('src')	
		image_ext = ''
		if 'jpg' in image_link:
			image_ext = 'jpg'
		elif 'png' in image_link:
			image_ext = 'png'
		elif 'jpeg' in image_link:
			image_ext = 'jpeg'
		
		if image_ext in ['jpg', 'jpeg', 'png']:
			image_name = 'coursera_{index}.{ext}'.format(index=count, ext=image_ext)
			urllib.request.urlretrieve(image_link, image_name)
		else:
			image_name='dummy.jpg'
		time.sleep(1)
		
		if len(enrolled) == 0:		enrolled = '0'
		else:						enrolled = enrolled[0].text
		
		if len(rating) == 0:		rating = 'Not rated yet'
		else:						rating = rating[0].text

		if len(rated_by) == 0:		rated_by = '0'
		else:						rated_by = rated_by[0].text[1:-1]

		print(title[0].text)
		obj = {
			'Image': image_name,
			'Title': title[0].text,
			'Creator': creator[0].text,
			'Type': courseType[0].text,
			'Enrolled': enrolled,
			'Difficulty': difficulty[0].text,
			'Rating': rating,
			'Rated by': rated_by,
			'Link': getCourseLink(title[0].text, courseType[0].text)
		}
		coursera_data = coursera_data.append(obj, ignore_index=True)
	
	print("Data fetched from Coursera !!")
	return coursera_data

def getCourseLink(name, typeOfCourse):
	name = name.lower().replace(' ', '-')
	COURSE_TYPE = {
		'SPECIALIZATION': 'specializations',
		'COURSE': 'course',
		'PROFESSIONAL CERTIFICATE': 'professional-certificates',
		'GUIDED PROJECT': 'projects'
	}
	typeUrl = COURSE_TYPE[typeOfCourse]
	baseUrl = 'https://www.coursera.org/{courseType}/{courseName}'
	url = baseUrl.format(courseType=typeUrl, courseName=name)
	return url



def saveDataToExcel(udemy_data, coursera_data):
	workbook = xlsxwriter.Workbook("Course Finder.xlsx")
	global_format = workbook.add_format()
	global_format.set_font_name('Fira Sans')

	# Creating standard formats for entire worksheet
	title_format = workbook.add_format({'bold': 1, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 13, 'bg_color': 'E9A777'})
	title_format.set_shrink()
	title_format.set_text_wrap()
	heading_format = workbook.add_format({'bold': 1, 'valign': 'vcenter'})
	heading_format.set_shrink()
	heading_format.set_text_wrap()
	cell_format = workbook.add_format({'valign': 'vcenter'})
	cell_format.set_shrink()
	cell_format.set_text_wrap()
	image_format = workbook.add_format({'border': 1})
	# # Adding sheet for Udemy courses
	udemy_sheet = workbook.add_worksheet('Udemy')

	udemy_sheet.set_column(1, 20, 12)
	udemy_sheet.set_column(4, 4, 17)
	udemy_sheet.set_column(0, 0, 5)
	udemy_sheet.set_default_row(20)

	# Iterating over udemy courses and adding them as cards in excel.
	time.sleep(3)
	row = 1
	print(udemy_data)
	for index, card in udemy_data.iterrows():
		# print(index, card)
		udemy_sheet.write(row, 0, index+1, cell_format)
		udemy_sheet.merge_range(row, 4, row, 9, card['Title'], title_format)
		udemy_sheet.write(row + 1, 4, 'Created By : ', heading_format)
		udemy_sheet.write(row + 3, 4, 'Description : ', heading_format)
		udemy_sheet.write(row + 5, 4, 'Duration : ', heading_format)
		udemy_sheet.write(row + 6, 4, 'Rating : ', heading_format)
		udemy_sheet.write(row + 7, 4, 'Price : ', heading_format)
		udemy_sheet.write(row + 7, 7, 'Rated by : ', heading_format)
		udemy_sheet.write(row + 8, 4, 'Link : ', heading_format)
		udemy_sheet.merge_range(row, 1, row + 8, 3, '', image_format)

		image_name = card['Image']
		image_file = open(image_name, 'rb')
		image_data = io.BytesIO(image_file.read())
		image_file.close()
		udemy_sheet.insert_image(row, 1,  image_name, {'image_data': image_data, 'x_scale': 0.7, 'y_scale': 1, 'x_offset': 45, 'y_offset': 40})
    
		udemy_sheet.merge_range(row + 1, 5, row + 2, 9, card['Instructors'], cell_format)
		udemy_sheet.merge_range(row + 3, 5, row + 4, 9, card['Description'], cell_format)
		udemy_sheet.merge_range(row + 5, 5, row + 5, 9, card['Course Time'], cell_format)
		udemy_sheet.write(row + 6, 5, card['Rating'], cell_format)
		udemy_sheet.write(row + 7, 5, card['Price'], cell_format)
		udemy_sheet.merge_range(row + 7, 8, row + 7, 9, card['Rated by'], cell_format)
		udemy_sheet.merge_range(row + 8, 5, row + 8, 9, '', cell_format)
		udemy_sheet.write_url(row+8, 5, card['Link'], cell_format, string='Go to course')

		row += 11

	# Adding sheet for Coursera courses
	era_sheet = workbook.add_worksheet('Coursera')

	era_sheet.set_column(1, 20, 12)
	era_sheet.set_column(4, 4, 17)
	era_sheet.set_column(0, 0, 5)
	era_sheet.set_default_row(20)

	# Iterating over udemy courses and adding them as cards in excel.
	row = 1
	print(coursera_data)
	time.sleep(3)
	for index, card in coursera_data.iterrows():
		# print(index, card)
		era_sheet.write(row, 0, index+1, cell_format)
		era_sheet.merge_range(row, 3, row, 8, card['Title'], title_format)
		era_sheet.write(row + 1, 3, 'Created by : ', heading_format)
		era_sheet.write(row + 2, 3, 'Type : ', heading_format)
		era_sheet.write(row + 3, 3, 'Enrolled : ', heading_format)
		era_sheet.write(row + 4, 3, 'Level : ' , heading_format)
		era_sheet.write(row + 5, 3, 'Link : ', heading_format)
		era_sheet.write(row + 3, 6, 'Rating : ', heading_format)
		era_sheet.write(row + 4, 6, 'Rated by : ', heading_format)
		era_sheet.merge_range(row, 1, row + 5, 2, '', image_format)

		image_name = card['Image']
		image_file = open(image_name, 'rb')
		image_data = io.BytesIO(image_file.read())
		image_file.close()
		era_sheet.insert_image(row, 1,  image_name, {'image_data': image_data, 'x_scale': 0.7, 'y_scale': 0.7, 'x_offset': 25, 'y_offset': 10})
    
		era_sheet.merge_range(row + 1, 4, row + 1, 8, card['Creator'], cell_format)
		era_sheet.merge_range(row + 2, 4, row + 2, 5, card['Type'], cell_format)
		era_sheet.merge_range(row + 3, 4, row + 3, 5, card['Enrolled'], cell_format)
		era_sheet.merge_range(row + 4, 4, row + 4, 5, card['Difficulty'], cell_format)
		era_sheet.write(row + 3, 7, card['Rating'], cell_format)
		era_sheet.write(row + 4, 7, card['Rated by'], cell_format)
		era_sheet.write_url(row + 5, 4, card['Link'], cell_format, string='Go to course')

		row += 9
	

	workbook.close()


def deleteImages():
	print('Deleting image files.')
	currDir = os.getcwd()
	for fname in os.listdir(currDir):
		if fname.startswith("udemy") or fname.startswith('coursera'):
			os.remove(os.path.join(currDir, fname))


def main():
	course = input("\nWhat do you want to learn ?? ")
	print("\n{course} is a great choice !!\n".format(course=course))
	course = course.replace(" ", "+")
	driver = webdriver.Chrome(PATH)
	udemy_data = scrapeUdemy(course, driver)
	coursera_data = scrapeCourse(course, driver)
	saveDataToExcel(udemy_data, coursera_data)
	deleteImages()
	driver.close()
	

if __name__ == '__main__':
	main()
