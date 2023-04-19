# coding: utf8
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import undetected_chromedriver as uc
import time
import os
import re
from datetime import datetime
import pandas as pd
import warnings
import sys
import xlsxwriter
from multiprocessing import freeze_support
import calendar 
import shutil

warnings.filterwarnings('ignore')
NoneType = type(None)

def initialize_bot():

    # Setting up chrome driver for the bot
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--headless')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # installing the chrome driver
    driver_path = ChromeDriverManager().install()
    chrome_service = ChromeService(driver_path)
    # configuring the driver
    driver = webdriver.Chrome(options=chrome_options, service=chrome_service)
    ver = int(driver.capabilities['chrome']['chromedriverVersion'].split('.')[0])
    driver.quit()
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--lang=en")
    chrome_options.add_argument('--headless=new')
    chrome_options.add_argument("--incognito")
    # disable location prompts & disable images loading
    #prefs = {"profile.default_content_setting_values.geolocation": 2, "profile.managed_default_content_settings.images": 1}  
    chrome_options.page_load_strategy = 'normal'
    #chrome_options.add_experimental_option("prefs", prefs)
    driver = uc.Chrome(version_main = ver, options=chrome_options) 
    driver.set_window_size(1920, 1080)
    driver.maximize_window()
    # page timeout in ms
    driver.set_page_load_timeout(20000)

    return driver

def initialize_output():

    stamp = datetime.now().strftime("%d_%m_%Y_%H_%M")
    path = os.getcwd() + '\\Scraped_Data\\' + stamp
    if os.path.exists(path):
        #os.remove(path)
        shutil.rmtree(path) 
    os.makedirs(path)

    file1 = f'Google_Maps_Output_{stamp}.xlsx'

    # Windws and Linux slashes
    if os.getcwd().find('/') != -1:
        output1 = path.replace('\\', '/') + "/" + file1
    else:
        output1 = path + "\\" + file1  

    # Create an new Excel file and add a worksheet.
    workbook1 = xlsxwriter.Workbook(output1)
    workbook1.add_worksheet()
    workbook1.close()    

    return output1

def get_inputs():
 
    print('Processing The Input Sheet ...')
    # assuming the inputs to be in the same script directory
    path = os.getcwd()
    if '\\' in path:
        path += '\\Inputs.xlsx'
    else:
        path += '/Inputs.xlsx'

    if not os.path.isfile(path):
        print('Error: Missing the settings file "Inputs.xlsx"')
        input('Press any key to exit')
        sys.exit(1)
    try:
        keywords  = []
        limit = 0
        df = pd.read_excel(path)
        cols  = df.columns
        for col in cols:
            df[col] = df[col].astype(str)

        inds = df.index
        for ind in inds:
            row = df.iloc[ind]
            name, loc = '', ''
            for col in cols:
                if row[col] == 'nan': continue
                elif col == 'Name':
                    name = row[col]
                elif col == 'Location':
                    loc = row[col]
                elif col == 'Number of results':
                    limit = float(row[col])

            if name != '' or loc != '':
                keywords.append(name + '; ' + loc)
    except:
        print('Error: Failed to process the settings sheet')
        input('Press any key to exit')
        sys.exit(1)


    return keywords, limit

def scrape_Google_Maps(driver, keywords, output1, limit):

    days = ['Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday', 'Monday', 'Tuesday']
    df = pd.DataFrame()
    n = len(keywords)
    for i, key in enumerate(keywords):
        try:
            keyword_name = key.split(';')[0].strip()
            keyword_loc = key.split(';')[-1].strip()
            print('-'*75)
            print(f'Scraping the info for keyword {i+1}/{n}')
            print('-'*75)
            driver.get('https://www.google.com/maps/?hl=en')
            search = wait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@id='searchboxinput']")))
            search.clear()
            search.send_keys(key)
            search.send_keys(Keys.ENTER)
            time.sleep(2)

            card = False
            res = []
            try:
                # multiple results
                div = wait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class='m6QErb DxyBCb kA9KIf dS8AEf ecceSd']")))[-1]
                while True:
                    height1 = div.get_attribute('scrollHeight')
                    for _ in range(10):
                        div.send_keys(Keys.PAGE_DOWN)
                        time.sleep(0.5)
                    time.sleep(4)
                    height2 = div.get_attribute('scrollHeight')
                    if height1 == height2: 
                        break

                res = wait(driver, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[class='hfpxzc']")))
                nres = len(res)
                card = True
            except:
                # single result
                nres = 1

            # applying the limit
            if limit > 0 and limit < nres:
                nres = int(limit)

            for j in range(nres):
                details = {}
                details['Keyword Name'] = keyword_name
                details['Keyword Location'] =  keyword_loc
                print(f'Scraping the info for result {j+1}/{nres}')
                try:
                    button = res[j]
                    driver.execute_script("arguments[0].click();", button)
                    time.sleep(4)
                except:
                    pass

                # result name
                name = ''
                try:
                    name = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1[class='DUwDvf fontHeadlineLarge']"))).get_attribute('textContent')
                except:
                    pass
                # location details
                add, website, tel, plus,  = '', '', '', ''
                try:
                    divs = wait(driver, 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class*='RcCsl fVHpi w4vB1d NOE9ve M0S7ae AG25L']")))
                    for div in divs:
                        try:
                            button = wait(div, 5).until(EC.presence_of_element_located((By.TAG_NAME, "button")))
                            try:
                                a = wait(div, 2).until(EC.presence_of_element_located((By.TAG_NAME, "a")))
                                if website == '' or website == None:
                                    website = a.get_attribute('href')
                            except:
                                pass

                            if button is NoneType or button is None: 
                                continue

                            label = button.get_attribute('aria-label')
                            if label is NoneType or label is None:
                                continue
                            if 'Plus code:' in label and plus == '':
                                plus = label.split(':')[-1].strip()
                            elif 'Address:' in label and add == '':
                                add = label.split(':')[-1].strip()
                            elif 'Phone:' in label and tel == '':
                                tel = label.split(':')[-1].strip()
                        except:
                            pass
                    if plus == '':
                        divs = wait(driver, 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class='Io6YTe fontBodyMedium']")))
                        for div in divs:
                            text = div.get_attribute('textContent')
                            if '+' in text and text[0] != '+' and 'LGBTQ' not in text:
                                plus = text
                                break
                except:
                    pass

                details['Result Name'] = name
                details['Google Maps URL'] = driver.current_url
                details['Address'] = add
                details['Website'] = website
                details['Telephone'] = tel
                details['Google Location'] = plus

                # rating and number of reviews
                rating, nrevs = '', ''
                try:
                    rating = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='fontDisplayLarge']"))).get_attribute('textContent')
                    nrevs = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[class='HHrUdb fontTitleSmall rqjGif']"))).get_attribute('textContent').split()[0]
                except:
                    pass

                details['Rating'] = rating
                details['Number of Reviews'] = nrevs

                # image
                img = ''
                try:
                    button = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[class='aoRNLd kn2E5e NMjTrf lvtCsd']")))
                    img = wait(button, 5).until(EC.presence_of_element_located((By.TAG_NAME, "img"))).get_attribute('src')
                except:
                    pass

                details['Image'] = img

                # opening hours
                try:
                    if card:
                        try:
                            buttons = wait(driver, 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "button[class='CsEnBe']")))
                            for button in buttons:
                                if 'See more hours' in button.get_attribute('textContent'):
                                    driver.execute_script("arguments[0].click();", button)
                                    time.sleep(1)
                                    break
                        except:
                            pass
                    table = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "table[class*='eK4R0e fontBodyMedium']")))
                    trs = wait(table, 5).until(EC.presence_of_all_elements_located((By.TAG_NAME, "tr")))
                    for tr in trs:
                        text = tr.get_attribute('textContent')
                        for day in days:
                            if day in text:
                                details[f'{day} Working Hours'] = text.replace(day, '').strip().upper().replace('M', 'M\n')

                    if card:
                        try:
                            button = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[class='VfPpkd-icon-LgbsSe yHy1rc eT1oJ mN1ivc']")))
                            driver.execute_script("arguments[0].click();", button)
                            time.sleep(1)
                        except:
                            pass
                except:
                    pass

                # popular times
                for _ in range(6):
                    try:
                        day = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='goog-inline-block goog-menu-button-caption']"))).get_attribute('textContent')
                        div = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='g2BVhd eoFzo']")))
                        divs = wait(div, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class*='dpoVLd']")))
                        max_per, min_per = 0, 1000
                        max_time, min_time = '', ''
                        for div in divs:
                            try:
                                text = div.get_attribute('aria-label').strip()
                                if  text[:2] == '0%': continue
                                per = int(text.split('%')[0].strip())
                                visit_time = text.split('at')[-1].strip()[:-1].upper()
                                max_per = max(per, max_per)
                                if max_per == per:
                                    max_time = visit_time 
                                min_per = min(per, min_per)
                                if min_per == per:
                                    min_time = visit_time
                            except:
                                pass
                        if max_per != 0:
                            details[f'{day} Max Visit %'] = max_per
                            details[f'{day} Max Visit Time'] = max_time
                        if min_per != 1000:
                            details[f'{day} Min Visit %'] = min_per
                            details[f'{day} Min Visit Time'] = min_time
                        button = wait(driver, 2).until(EC.presence_of_element_located((By.XPATH, "//button[@aria-label='Go to the next day']")))
                        driver.execute_script("arguments[0].click();", button)
                        time.sleep(1)
                    except Exception as err:
                        pass

                # restaurant features
                try:
                    divs = wait(driver, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class='LTs0Rc']")))
                    for div in divs:
                        try:
                            feature = div.get_attribute('textContent').split('·')[-1].strip()
                            img = wait(div, 2).until(EC.presence_of_element_located((By.TAG_NAME, "img"))).get_attribute('src')
                            if 'ic_done' in img:
                                details[feature] = 'Yes'
                            else:
                                details[feature] = 'No'
                        except:
                            pass
                except:
                    pass
                # appending the output to the datafame       
                df = df.append([details.copy()])
        except Exception as err:
            print(f'The below error occurred while scraping: {name}')
            print(str(err))

    # outputting the scraped data 
    df.to_excel(output1, index=False)

def main():

    print('Initializing The Bot ...')
    print('-'*75)
    freeze_support()
    start = time.time()
    keywords, limit = get_inputs()
    output1 = initialize_output()
    try:
        driver = initialize_bot()
    except Exception as err:
        print('Failed to initialize the Chrome driver due to the following error:\n')
        print(str(err))
        print('-'*75)
        input('Press any key to exit.')
        sys.exit()

    try:
        scrape_Google_Maps(driver, keywords, output1, limit)
    except Exception as err: 
        print(f'Warning: the below error occurred:\n {err}')
        driver.quit()
        time.sleep(5)
        driver = initialize_bot()

    driver.quit()
    print('-'*75)
    elapsed_time = round(((time.time() - start)/60), 2)
    input(f'Process is completed in {elapsed_time} mins, Press any key to exit.')

if __name__ == '__main__':

    main()
