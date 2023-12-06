from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import time, argparse, xlsxwriter, re

username = "YOUR_EMAIL_ADDRESS"
password = "YOUR_PASSWORD"

url = "https://prosperousuniverse.com/auth/login/"
excel_file = "item.xlsx"
item_categories = [
    "Agricultural Product",
    "Alloy",
    "Chemical",
    "Construction Material",
    "Construction Part",
    "Construction Prefab",
    "Consumable (basic)",
    "Consumable (luxury)",
    "Drone",
    "Electronic Device",
    "Electronic Part",
    "Electronic Piece",
    "Electronic System",
    "Element",
    "Energy System",
    "Fuel",
    "Gas",
    "Liquid",
    "Medical Equipment",
    "Metal",
    "Mineral",
    "Ore",
    "Plastic",
    "Ship Engine",
    "Ship Kit",
    "Ship Part",
    "Ship Shield",
    "Software Component",
    "Software System",
    "Software Tool",
    "Textile",
    "Unit Prefab",
    "Utility"
]

def initialize_browser():
    options = webdriver.ChromeOptions()
    browser = webdriver.Chrome(options=options)
    return browser

def open_url(browser):
    browser.get(url)

def login(browser):
    username_input = browser.find_element(By.NAME, "login")
    password_input = browser.find_element(By.NAME, "password")
    login_button = browser.find_element(By.XPATH, "//button[@type='submit']")
    username_input.send_keys(username)
    password_input.send_keys(password)
    login_button.click()

def click_play(browser):
    play_button_xpath = "//a[@class='btn--primary btn--large' and @href='https://apex.prosperousuniverse.com' and contains(text(), 'Play!')]"
    play_button = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located((By.XPATH, play_button_xpath))
    )
    play_button.click()

def process_table(table, category, worksheet):
    rows = table.find('tbody').find_all('tr')

    items = []
    for row in rows:
        tds = row.find_all('td')
        name = tds[1].find_all('span')[1].get_text(strip=True).split('.')[0]
        high_price_spans = tds[3].find_all('span')
        high_price = 100000000 if len(high_price_spans) <= 1 else high_price_spans[0].find('span').find('span').get_text(strip=True).replace(',', '')

        low_price_spans = tds[4].find_all('span')
        low_price = 0 if len(low_price_spans) <= 1 else low_price_spans[0].find('span').find('span').get_text(strip=True).replace(',', '')

        supply = tds[5].find_all('span')[0].find('span').get_text(strip=True).replace(',', '')
        demand = tds[5].find('span', class_='BrokerList__subLine___GYIC_zD type__type-small___pMQhMQO').find('span').find('span').get_text(strip=True).replace(',', '')

        items.append((name, category, high_price, low_price, supply, demand))
    return items

def parse_html_for_tables(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    tables = soup.find_all('table', class_='BrokerList__table___ylgeiyg')

    workbook = xlsxwriter.Workbook(excel_file)
    worksheet = workbook.add_worksheet()
    current_row = 0

    for i in range(len(tables)):
        category = item_categories[i]
        items = process_table(tables[i], category, worksheet)

        for item in items :
            worksheet.write(current_row, 0, item[0])
            worksheet.write(current_row, 1, item[1])
            worksheet.write(current_row, 2, item[2])
            worksheet.write(current_row, 3, item[3])
            worksheet.write(current_row, 4, item[4])
            worksheet.write(current_row, 5, item[5])
            current_row += 1

    workbook.close()

def main():
    try:
        browser = initialize_browser()
        open_url(browser)
        login(browser)
        click_play(browser)
        time.sleep(10)
        html_content = browser.page_source
        parse_html_for_tables(html_content)
    finally:
        browser.quit()
        pass

if __name__ == "__main__":
    main()