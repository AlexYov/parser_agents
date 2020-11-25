from selenium import webdriver
from openpyxl import Workbook

browser = webdriver.Firefox()

work_book = Workbook()
work_book.active
work_sheet = work_book.create_sheet('агенства', 0)
work_sheet.append(['Название', 'Почта', 'Телефон', 'Сайт'])

browser.get('https://1ibn.ru/reestr?rgr_region=%D0%9E%D1%80%D0%B5%D0%BD%D0%B1%D1%83%D1%80%D0%B3%D1%81%D0%BA%D0%B0%D1%8F+%D0%BE%D0%B1%D0%BB%D0%B0%D1%81%D1%82%D1%8C&term=&page=7')
class_names = browser.find_elements_by_class_name('agent-block')

links = [class_name.find_element_by_tag_name('a').get_attribute('href') for class_name in class_names]

dict_data = {}

for link in links:
    browser.get(link)
    name  = browser.find_element_by_xpath('/html/body/main/div/div[1]/div[1]/div/div/div/div[2]/div[1]/div/h1').text
    rows = browser.find_elements_by_class_name('col-md-8')
    data = {name:{}}
    dict_data.update(data)
    for row in rows:
        strings = row.find_elements_by_class_name('row')
        for string in strings:
            if 'Тел' in string.text:
                phone = string.text.split('Телефон')[1]
                data[name].update({'phone':phone})
                
            if 'Email' in string.text:
                email = string.text.split('Email')[1]
                data[name].update({'email':email})
 
            if 'Веб-сайт' in string.text:
                site = string.text.split('Веб-сайт')[1]
                data[name].update({'site':site})  

row = 1

for name in dict_data.keys():
    row += 1
    work_sheet['A' + str(row)] = name 
    work_sheet['B' + str(row)] = dict_data[name].setdefault('email')
    work_sheet['C' + str(row)] = dict_data[name].setdefault('phone')
    work_sheet['D' + str(row)] = dict_data[name].setdefault('site')
    
work_book.save('агенства Оренбургская обл_7.xlsx')