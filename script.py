from selenium import webdriver
import pandas as pd
import time
import re
from selenium.webdriver.common.keys import Keys


path = '/Users/EC9WU4/Downloads/Python/Book2.xlsx'

website = ''
username = ''
password = ''


driver = webdriver.Chrome()



driver = webdriver.Chrome(executable_path='/Users/EC9WU4/Downloads/chromedriver 3')

driver.get(website)

driver.find_element_by_xpath('//input[@type="text"]').send_keys(username)

driver.find_element_by_xpath('//input[@type="password"]').send_keys(password)
driver.find_element_by_xpath('//input[@type="password"]').send_keys(Keys.RETURN)

df = pd.read_excel(path)

for index, row in df.iterrows():
    driver.get(website)
    inpu = driver.find_element_by_xpath('//input[@type="text"]')

    # place your column name
    inpu.send_keys(row['ORDER_NUMBER'])
    driver.find_element_by_xpath('button[@value="Submit"]').click()
    time.sleep(1)
    all_emails = []
    while True:
        page = driver.page_source
        emails = re.findall('[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+',page)
        all_emails.extend(emails)
        try:
            driver.find_element_by_xpath('//a[text()="Next"]').click()
            time.sleep(1)
        except:
            break


    all_emails = list(set(all_emails))
    all_emails = ' , '.join(all_emails)
    # change this to second column
    row['customer email address'] = all_emails


writer = pd.ExcelWriter(path, engine='xlsxwriter')
df.to_excel(writer,sheet_name = 'Data', index=False)
writer.save()

driver.quit()
