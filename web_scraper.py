import xlrd, xlwt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import sys
import os
import boto
import boto.s3.connection
from boto.s3.key import Key
import json
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.mime.application import MIMEApplication
from pyvirtualdisplay import Display

error_str = ''

def error_mail(message):
	username = 'amazonwebscraping@10i.co.in'
	pwd = '10iComm3rc3'
	msg = MIMEMultipart()
	msg['From'] = 'amazonwebscraping@10i.co.in'
	msg['Subject'] = "ERROR while scraping amazon business website"
	body = message
	msg.attach(MIMEText(body, 'plain'))
	server = smtplib.SMTP('smtp.gmail.com:587')
	server.ehlo()
	server.starttls()
	server.login(username,pwd)
	text = msg.as_string()
	server.sendmail('amazonwebscraping@10i.co.in', ['badal@10i.co.in','kishan@10i.co.in'], text)
	server.quit()
	sys.exit()


def xls_mail(file_name):
	username = 'amazonwebscraping@10i.co.in'
	pwd = '10iComm3rc3'
	msg = MIMEMultipart()
	msg['From'] = 'amazonwebscraping@10i.co.in'
	msg['Subject'] = "Scraped data"
	body = 'Data scraped on ' + time.strftime("%d-%m-%Y")
	msg.attach(MIMEText(body, 'plain'))
	part = MIMEApplication(open(file_name,"rb").read())
	part.add_header('Content-Disposition', 'attachment', filename=file_name)
	msg.attach(part)
	server = smtplib.SMTP('smtp.gmail.com:587')
	server.ehlo()
	server.starttls()
	server.login(username,pwd)
	text = msg.as_string()
	server.sendmail('amazonwebscraping@10i.co.in', ['badal@10i.co.in','kishan@10i.co.in'], text)
	server.quit()


if(len(sys.argv) > 1):
	try:
		book = xlrd.open_workbook(sys.argv[1])
	except Exception, e:
		error_str = error_str + str(e)
		error_mail(error_str)
else:
	try:
		book = xlrd.open_workbook("Amazon SKUs.xlsx")
	except Exception, e:
		error_str = error_str + str(e)
		error_mail(error_str)

names = book.sheet_names()

try:
	
	driver = webdriver.Chrome()
	driver.get('https://www.amazonbusiness.in/login?ref_=nb_si')
	time.sleep(1)
	search = driver.find_element_by_name('email')
	search.send_keys('deepa.tallur123@gmail.com')
	search = driver.find_element_by_name('password')
	search.send_keys('Deepasai@123')
	search.send_keys(Keys.RETURN)
	time.sleep(1)
except Exception, e:
		error_str = error_str + str(e)
		error_mail(error_str)


wbook = xlwt.Workbook()
font = xlwt.Font()
font.name = 'Calibri'
font.height = 11*0x14
style = xlwt.XFStyle()
style.font = font


for i in range(book.nsheets):
	sheet = book.sheet_by_index(i)
	sheet_length = sheet.nrows
	wsheet = wbook.add_sheet(names[i])
	wsheet.col(1).width = 256*75
	wsheet.col(2).width = 256*25
	wsheet.col(3).width = 256*25
	wsheet.col(4).width = 256*25
	wsheet.col(5).width = 256*25
	wsheet.col(6).width = 256*25
	wsheet.col(7).width = 256*25
	wsheet.col(8).width = 256*25
	wsheet.col(9).width = 256*25
	wsheet.col(10).width= 256*40
	wsheet.row(0).height_mismatch= True
	wsheet.row(0).height = 400
	wsheet.write(0, 0, 'S. No.', style)
	wsheet.write(0, 1, 'Product Name on Amazon', style)
	wsheet.write(0, 2, 'productASIN', style)
	wsheet.write(0, 3, 'MRP', style)
	wsheet.write(0, 4, 'Price Incl. VAT', style)
	wsheet.write(0, 5, 'Price Excl. VAT', style)
	wsheet.write(0, 6, 'Margin', style)
	wsheet.write(0, 7, 'Availability', style)
	wsheet.write(0, 8, 'Product Weight', style)
	wsheet.write(0, 9, 'Shipping Weight', style)
	wsheet.write(0,10, 'Quantity', style)

	for x in range(0, sheet_length):
		try:
			skuasin = sheet.row_values(x)[0]
			search = driver.find_element_by_name('keywords')
			search.clear()
			search.send_keys(skuasin)
			search.send_keys(Keys.RETURN)
			time.sleep(1)
			res = driver.find_elements_by_class_name('a-size-large')
			wsheet.row(x+1).height_mismatch = True
			wsheet.row(x+1).height = 350
			wsheet.write(x+1, 0, x+1)
			wsheet.write(x+1, 1, res[0].text, style)
			res = driver.find_elements_by_class_name('a-size-medium')
			wsheet.write(x+1, 2, skuasin, style)
		except Exception, e:
			error_str = error_str + "\n" + str(e) + "Row Number :" + str(x+1) + "\n"
			#error_mail(error_str)
		present = (len(res)==3)
		if(present):
			try:
				wsheet.write(x+1, 7, 'Out of Stock', style)
				wsheet.write(x+1, 4, 'N/A', style)
				wsheet.write(x+1, 5, 'N/A', style)
				wsheet.write(x+1, 6, 'N/A', style)
				wsheet.write(x+1,10, 'N/A', style)
				res = driver.find_elements_by_class_name('a-size-base')
				val = res[1].text
				wsheet.write(x+1, 3, 'N/A', style)
			except Exception, e:
				error_str = error_str + "\n" + str(e) + "Row Number :" + str(x+1) + "\n"
				#error_mail(error_str)
		else:
			try:
				wsheet.write(x+1, 5, res[2].text, style)
				wsheet.write(x+1, 4, res[3].text, style)
				wsheet.write(x+1, 7, res[4].text, style)
				res = driver.find_elements_by_class_name('a-size-base')
				val = res[1].text
				ind = val.find('(')
				wsheet.write(x+1, 3, val[4:ind], style)
				wsheet.write(x+1, 6, val[ind+1:len(val)-7], style)
				res = driver.find_element_by_name('quantity')
				res.clear()
				res.send_keys('999')
				#time.sleep(1)
				res = driver.find_element_by_class_name('a-alert-content')
				wsheet.write(x+1, 10, res.text, style)
			except Exception, e:
				error_str = error_str + "\n" + str(e) + "Row Number :" + str(x+1) + "\n"
				#error_mail(error_str)
		res =  driver.find_elements_by_class_name('a-keyvalue')
		res = res[0].text.split("\n")
		try:
			if(len(res) == 3):
				wsheet.write(x+1, 8, res[1][15:len(res[1])], style)
				wsheet.write(x+1, 9, res[2][16:len(res[2])], style)
			elif(len(res)==2):
				wsheet.write(x+1, 9, res[1][15:len(res[1])], style)
				wsheet.write(x+1, 8, 'N/A', style)
			else :
				wsheet.write(x+1, 8, 'N/A', style)
				wsheet.write(x+1, 9, 'N/A', style)
		except Exception, e:
			error_str = error_str + "\n" + str(e) + "Row Number :" + str(x+1) + "\n"
			#error_mail(error_str)

try:
	filename = "Scraped_on_" + time.strftime('%d%m%Y')
	filename +=".xls"
	wbook.save(filename)
	driver.quit()
except Exception, e:
		error_str = error_str + str(e)
		error_mail(error_str)
if(len(error_str)==0):
	book = xlrd.open_workbook(filename)
	names = book.sheet_names()
	json_list = []
	for i in range(book.nsheets):
		sheet = book.sheet_by_index(i)
		sheet_length = sheet.nrows
		for x in range(1, sheet_length):
			row = sheet.row_values(x)
			data = {
				"dt" : time.strftime('%H:%M:%S').encode('ascii', 'ignore'),
				"region" : "karnataka",
				"name" : row[1].encode('ascii', 'ignore'),
				"asin" : row[2].encode('ascii', 'ignore'),
				"MRP" : row[3].encode('ascii', 'ignore'),
				"PriceWithVAT" : row[4].encode('ascii', 'ignore'),
				"PriceWithoutVAT" : row[5].encode('ascii', 'ignore'),
				"Margin" : row[6].encode('ascii', 'ignore'),
				"Availability" : row[7].encode('ascii', 'ignore'),
				"ProductWt" : row[8].encode('ascii', 'ignore'),
				"ShippingWt" : row[9].encode('ascii', 'ignore')
			}	
			json_list.append(data)


	date = time.strftime("%Y%m%d")

	try:
		with open(date+'.json', 'w') as fp:
			json.dump(json_list, fp)

		conn = boto.s3.connect_to_region('us-west-1')
		bucket = conn.get_bucket('10i-webscraping')
		k = Key(bucket)
		UPLOADED_FILENAME = date+'.json'
		k.key = UPLOADED_FILENAME
		k.set_contents_from_filename(date+'.json')
	except Exception, e:
			error_str = error_str + str(e)
			error_mail(error_str)
else:
	xls_mail(filename)
	error_mail(error_str)


try:
	xls_mail(filename)
except Exception, e:
		error_str = error_str + str(e)
		error_mail(error_str)


os.remove(filename)
sys.exit()