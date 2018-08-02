import os
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import csv
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup as BS
import xlsxwriter
import time



chromedriver = "/Users/jross128/Desktop/temp/python/chromedriver"
os.environ["webdriver.chrome.driver"] = chromedriver
driver = webdriver.Chrome(chromedriver)
driver.get("https://www.treasurydirect.gov/BC/SBCPrice")

pathName = os.getcwd()
file = open(os.path.join(pathName, "bonds.csv"), "rU")
reader = csv.reader(file, delimiter=',')
for column in reader:
	select = Select(driver.find_element_by_name('Denomination'))
	number = driver.find_element_by_name('SerialNumber')
	date = driver.find_element_by_name('IssueDate')
	enter = driver.find_element_by_name('btnAdd.x')
	number.send_keys(column[0])
	date.send_keys(column[2])
	select.select_by_visible_text(column[1])
	enter.click()

viewAll = driver.find_element_by_name('btnAll.x')
viewAll.click()
time.sleep(5)
table=[]
for tr in driver.find_elements_by_xpath('//table[@class="bnddata"]//tr'):
	tds = tr.find_elements_by_tag_name('td')
	if tds: 
		table.append([td.text for td in tds])
	else:
		ths = tr.find_elements_by_tag_name('th')
		if ths:
			table.append([th.text for th in ths])

workbook = xlsxwriter.Workbook('bonds.xlsx')
worksheet = workbook.add_worksheet()

col = 0

for row, data in enumerate(table):
    worksheet.write_row(row, col, data)

workbook.close()
driver.close()
