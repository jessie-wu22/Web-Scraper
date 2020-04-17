from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from multiprocessing import Pool

import requests
from google.cloud import vision
import os

from bs4 import BeautifulSoup

import time
import datetime
import pandas as pd

df = pd.DataFrame()  # pandas creates a data frame structure (to later import into excel spreadsheet)

timeout = 20

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "project1.json" # connects google cloud account to API
browser = webdriver.Chrome(executable_path="/usr/local/bin/chromedriver")
base_url = "https://www.webwinkelkeur.nl/ledenlijst/?fbclid=IwAR1hQFPToBS-3ZDltAb5gk43IqPqqIIV4fYFKw8Skp_3ihIk3ktnUKUcUrU&sort=rating"
browser.get(base_url)

try:  # program will terminate after 20s of page loading
    WebDriverWait(browser, timeout).until(EC.visibility_of_element_located((By.XPATH, "//*[@id='webshops']")))
except TimeoutException:
    print("Timed out waiting for page to load")
    browser.quit()

# creates lists to store information retrieved for each webshop
URLs = []
name_list = []
member_since_list = []
website_url_list = []
address_list = []
phone_number_list = []
chamberofcommerce_list = []
vat_list = []
page_links = []
email_list = []
email_image_urls = []

# runs through all webpages to collect webshop page URLs on Webwinkelkeur
def compile_items_on_page():
    time.sleep(0.5)  # waits 0.5s to ensure webpage loaded properly
    for i in range(0, 355):  # iterates through 365 pages
        time.sleep(0.5)
        elements_list = browser.find_elements_by_css_selector("div.col-sm-6 a.text-body")  # finds all cards on each page

        for element in elements_list:  # extracts links for each card
            URLs.append(element.get_attribute("href"))

        try:
            next_page = browser.find_element_by_xpath("//*[@id='nextWebshopsPage']")  # finds next page button
            next_page.click()  # clicks on button at bottom of page to navigate to next page
        except Exception as e:
            pass

compile_items_on_page()


# gets all the required information - webshop name, membership date, website, address, phone number, KvK, VAT
def extract(url):
    page = requests.get(url)
    soup = BeautifulSoup(page.text, 'html.parser')

    try:
        name = soup.find("span", class_="value").get_text()  # finds name by pointing to corresponding HTML tag
        name_list.append(name)
    except Exception as e:
        name_list.append("N/A")
        pass
    try:
        member_since = soup.select("body > div > main > header > div > aside > ul > li > dl > dd")[0].text.strip()
        year = member_since.split(" ")
        try:
            int(year[-1])
        except ValueError:
            year[-1] = "N/A"
        member_since_list.append(year[-1])
    except Exception as e:
        member_since_list.append("N/A")
        pass
    try:
        website_url = soup.select("#badge-1 > div > div > dl > dd:nth-child(3) > a")[0].text.strip()
        website_url_list.append(website_url)
    except Exception as e:
        website_url_list.append("N/A")
        pass
    try:
        address = soup.select("#badge-1 > div > div > dl > dd:nth-child(6)")[0].text.strip()
        address_list.append(address)
    except Exception as e:
        address_list.append("N/A")
        pass
    try:
        phone_number = soup.select("#badge-1 > div > div > dl > dd:nth-child(12)")[0].text.strip()
        phone_number_list.append(phone_number)
    except Exception as e:
        phone_number_list.append("N/A")
        pass
    try:
        chamberofcommerce = soup.select("#badge-1 > div > div > dl > dd:nth-child(15)")[0].text.strip()
        chamberofcommerce_list.append(chamberofcommerce)
    except Exception as e:
        chamberofcommerce_list.append("N/A")
        pass
    try:
        vat = soup.select("#badge-1 > div > div > dl > dd:nth-child(18)")[0].text.strip()
        vat_list.append(vat)
    except Exception as e:
        vat_list.append("N/A")
        pass
    try:
        email_image = soup.select("#badge-1 > div > div > dl > dd:nth-child(9) > img")[0]['src']
        email_image_urls.append(email_image)
        client = vision.ImageAnnotatorClient()
        image = vision.types.Image()
        image.source.image_uri = email_image
        response = client.text_detection(image=image)
        texts = response.text_annotations
        email_list.append(texts[0].description)
    except Exception as e:
        email.list.append("N/A")
        pass


with Pool(10) as p: # concurrency
    p.map(extract, URLs)


def create_spreadsheet():
    # creating excel sheet from data extracted
    for i in range(len(URLs)):
        try:
            df.loc[i, 'Page URL'] = URLs[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'Webshop Name'] = name_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'Member Since (Year)'] = member_since_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'Website URL'] = website_url_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'Address'] = address_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'Phone Number'] = phone_number_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'Chamber of Commerce Number'] = chamberofcommerce_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'VAT Number'] = vat_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'Email'] = email_list[i]
        except Exception as e:
            pass
        try:
            df.loc[i, 'Email Picture URLs'] = email_image_urls[i]
        except Exception as e:
            pass


create_spreadsheet()
df.to_excel('complete.xlsx')
