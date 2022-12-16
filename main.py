import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook, load_workbook
import time
import xlsxwriter

print("**Welcome**")
print("1. output all")
print("2. quit")
userIn = input("Enter your desired function: ")
fileName = input("Enter xlsx output file name:")
options = Options()
options.headless = False
driver = webdriver.Chrome(options=options)
mUrl = "https://jobapscloud.com/Alameda/auditor/classspecs.asp"
wb = xlsxwriter.Workbook(str(fileName)+".xlsx")
worksheet = wb.add_worksheet("firstSheet")
valueList = ['Job Title', 'Job #', 'Bargaining Unit', 'Min Hourly', 'Max Hourly', 'Description',
		'Qualification']

for i in range(0, 6):
	worksheet.write(0, i, str(valueList[i]))

if userIn == '1':
	count = 2
	count2 = 1
	count3 = 3
	while True:
		driver.get(mUrl)
		try: 
			jUrl = driver.find_element(by=By.XPATH, value = "/html/body/div[1]/div[1]/div[2]/div[3]/div[2]/form/fieldset/ul/li["+str(count3)+"]/span/a").get_attribute('href')
		except:
			if(count2 < 26):
				count2 = count2 + 1
				print("letter#" + str(count2))
				count3 = count3 + 3
				continue
			else:
				print("DONE!")
				break

		driver.get(jUrl)
		title = driver.find_element(by=By.XPATH, value ="/html/body/div/div/h2").get_attribute("textContent")
		codes = driver.find_element(by=By.XPATH, value ="/html/body/div/div/p[3]").get_attribute("textContent")
		
		try:
			qual = driver.find_element(by=By.XPATH, value ="/html/body/div/table/tbody/tr[6]/td/div").get_attribute("textContent")
			desc = driver.find_element(by=By.XPATH, value ="//*[@id='hrscontent']/table/tbody/tr[2]/td/div").get_attribute("textContent")
								
		except:
			driver.get(mUrl)
			worksheet.write((count-1), 0, title)
			count = count + 1
			count3 = count3 + 1
			continue

		try:
			jCode = title.split("(")
			bUnit = codes[17:len(codes)].split("-")
			code3 = bUnit[1].split("$")
			hourly = bUnit[1].split(")")
			hourly2 = bUnit[2].split(" ")
			unionCode = bUnit[0]
			jobID = code3[0]
			hourRate1 = hourly[1]
			hourRate2 = hourly2[0]
			jobCode = jCode[1]
			worksheet.write((count-1), 0, str(jCode[0]))
		except:
			driver.get(mUrl)
			worksheet.write((count-1), 0, title)
			count = count + 1
			count3 = count3 + 1
			continue
		worksheet.write((count-1), 1, str(jobCode))
		worksheet.write((count-1), 2, str(unionCode) + str(jobID))
		worksheet.write((count-1), 3, str(hourRate1))
		worksheet.write((count-1), 4, str(hourRate2))
		worksheet.write((count-1), 5, str(qual))
		worksheet.write((count-1), 6, str(desc))
		count3 = count3 + 1
		count = count + 1

elif userIn == 2:
	quit()

else:
	print("error incorrect option chosen")

wb.close()
driver.quit()