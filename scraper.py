from selenium import webdriver
from time import sleep
import openpyxl

url = 'https://portal.moreover.com/index.html#lists'

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--log-level=3')
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')  # required when running as root user
browser = webdriver.Chrome(executable_path='/usr/bin/chromedriver', chrome_options=chrome_options)

print('navigating browser to {}'.format(url))
browser.get(url)

print('logging in')
username = browser.find_element_by_name('j_username')
username.send_keys('xxxxx')
password = browser.find_element_by_name('j_password')
password.send_keys('xxxxx')
submit = browser.find_element_by_xpath('//*[@id="loginSubmit"]')
sleep(1)
submit.click()
sleep(30)

print('setting filters to blog and online')
feed_filter = browser.find_element_by_css_selector('#feedClass_ms')
feed_filter.click()
sleep(1)
blog_checkbox = browser.find_element_by_css_selector('#ui-multiselect-1-feedClass-option-3')
blog_checkbox.click()
sleep(1)
online_checkbox = browser.find_element_by_css_selector('#ui-multiselect-1-feedClass-option-1')
online_checkbox.click()
sleep(1)
feed_filter.click()
sleep(5)

print('opening source_list.xlsx')
wb = openpyxl.load_workbook('source_list.xlsx')
print('loading sheet')
sheet1 = wb['Sheet1']
sheet1['G1'] = 'Source Name'
sheet1['H1'] = 'url'
sheet1['I1'] = 'Language'
print('starting scrape...')
print()

total = 1
for i in range(2, 21239):  # cell range is B2 - B21238
    cell = sheet1['B{}'.format(i)].value

    '''Populate search field'''
    search_text = browser.find_element_by_xpath('//*[@id="search_text"]')
    search_text.send_keys(cell)
    sleep(1)
    search_btn = browser.find_element_by_xpath('//*[@id="search_button"]')
    search_btn.click()
    sleep(5)

    '''Obtain the Title and Language'''
    try:
        Title = browser.find_element_by_xpath('/html/body/div[1]/div[3]/div/main/div/div[2]/div/div/table/tbody/tr/td[2]/p').text
        Language = browser.find_element_by_xpath('/html/body/div[1]/div[3]/div/main/div/div[2]/div/div/table/tbody/tr/td[6]').text
        sheet1['G{}'.format(i)] = Title
        sheet1['H{}'.format(i)] = cell
        sheet1['I{}'.format(i)] = Language
        print(str(total), cell, 'Title:', Title, 'Language:', Language)
        total += 1
    except:
        Title = 'not found'
        Language = 'not found'
        sheet1['G{}'.format(i)] = Title
        sheet1['H{}'.format(i)] = cell
        sheet1['I{}'.format(i)] = Language
        print(str(total), cell, 'Title:', Title, 'Language:', Language)
        total += 1
    finally:
        clear_text = browser.find_element_by_xpath('//*[@id="search_text"]').clear()
        sleep(5)

wb.save('source_list_updated.xlsx')
print()
print('finished scrape')
