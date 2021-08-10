from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException, UnexpectedAlertPresentException
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep
import timeit as t
import pandas as pd

driver = webdriver.Firefox()
path = 'sts.xlsx'
table = pd.read_excel(path, sheet_name='work')
delay = 9

# Подключение к univer

driver.get('https://univer.kaznu.kz/user/login')
sleep(0.5)
usrname = 'shayakhmetov.kairat'
pwd = 'QwertyASD!@3'
login = driver.find_element_by_id('id_login')
login.send_keys(usrname)
password = driver.find_element_by_id('id_pass')
password.send_keys(pwd)
password.submit()
questions = []
# ------

# Переход в учебный отдел
try:
    myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[2]/a[3]/div[2]')))
except TimeoutException:
    print("Loading took too much time!")
driver.find_element_by_xpath('/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[2]/a[3]/div[2]').click()
try:
    myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[1]/tbody/tr/td[2]/table[1]/tbody/tr[1]/td[2]/ul/li[1]/a')))
except TimeoutException:
    print("Loading took too much time!")
driver.find_element_by_xpath('/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[1]/tbody/tr/td[2]/table[1]/tbody/tr[1]/td[2]/ul/li[1]/a').click()


for i in range(100):
    # Динамические переменные из таблицы
    try:
        name = table.loc[i]['name']
        chair = table.loc[i]['chair']
        email = table.loc[i]['email']
        predmet = table.loc[i]['predmet']
        faculty = table.loc[i]['faculty']
        student_count = table.loc[i]['student_count']
        split_name = name.split(' ')
        sname = split_name[0]
    except:
        print('Таблица закончилась')
        break
    # --------
    sleep(0.5)
    # Поиск преподователя
    while True:
        try:
            select = Select(driver.find_element_by_id('facultyId'))
            select.select_by_visible_text(faculty)
            # select = Select(driver.find_element_by_id('chairId'))
            # select.select_by_visible_text(chair)
            break
        except:
            continue
    driver.find_element_by_id('sname').send_keys(sname)
    driver.find_element_by_xpath('/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[5]/td[2]/input').click()
    sleep(0.5)
    driver.find_element_by_xpath(f'//*[contains(text(),\'{name}\')]').click()
    driver.find_element_by_id('questioners_btn').click()
    sleep(1)
    driver.find_element_by_id('/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[2]/tbody/tr[2]/td[2]/a[1]').click()
    # -------
    sleep(2)
    block_number = 3
    try:
        while True:
            sub_name = str(driver.find_element_by_xpath(
                f'/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[{block_number}]/tbody/tr[3]/td[2]/div/table/tbody/tr[2]/td[2]').text)
            sub_number = str(driver.find_element_by_xpath(
                f'/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[{block_number}]/tbody/tr[3]/td[2]/div/table/tbody/tr[2]/td[7]').text)
            if sub_name == predmet and sub_number == student_count:
                driver.find_element_by_xpath(
                    f'/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[{block_number}]/tbody/tr[2]/td[2]/a').click()
                break
            block_number += 1
    except:
        driver.get('https://univer.kaznu.kz/depart/learn/teachers/1')
        continue
    questions_count_1 = int(str(driver.find_element_by_xpath(
        '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[2]').text))
    questions_count_2 = int(str(driver.find_element_by_xpath(
        '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[2]').text))
    questions_count_3 = int(str(driver.find_element_by_xpath(
        '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td[2]').text))
    try:
        questions_balls_1 = int(str(driver.find_element_by_xpath(
            '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[3]').text))
        questions_balls_2 = int(str(driver.find_element_by_xpath(
            '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[3]').text))
        questions_balls_3 = int(str(driver.find_element_by_xpath(
            '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td[3]').text))
    except:
        driver.get('https://univer.kaznu.kz/depart/learn/teachers/1')
        continue
    total_questions_count = int(str(driver.find_element_by_xpath(
        '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[4]/td[2]/b').text))
    testname = str(driver.find_element_by_xpath(
        '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[2]/tbody/tr[1]/td[2]').text)
    test_name = testname.split('/')[-1]
    # -----------
    # /html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[4]/tbody/tr[2]/td[2]/table/tbody/tr[8]/td[2]/p
    # /html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[4]/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[2]
    sleep(1.3)
    questions_count_1 = int(str(driver.find_element_by_xpath(
        '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[2]').text))
    questions_count_2 = int(str(driver.find_element_by_xpath(
        '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[2]').text))
    questions_count_3 = int(str(driver.find_element_by_xpath(
        '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td[2]').text))
    try:
        questions_balls_1 = int(str(driver.find_element_by_xpath(
            '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[3]').text))
        questions_balls_2 = int(str(driver.find_element_by_xpath(
            '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[3]').text))
        questions_balls_3 = int(str(driver.find_element_by_xpath(
            '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[3]/td[3]').text))
    except:
        driver.get('https://univer.kaznu.kz/depart/learn/teachers/1')
        continue
    total_questions_count = int(str(driver.find_element_by_xpath(
        '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[3]/tbody/tr[2]/td[2]/table/tbody/tr[4]/td[2]/b').text))
    testname = str(driver.find_element_by_xpath(
        '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[2]/tbody/tr[1]/td[2]').text)
    test_name = testname.split('/')[-1]
    # -----------
    # /html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[4]/tbody/tr[2]/td[2]/table/tbody/tr[8]/td[2]/p
    # /html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[4]/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[2]
    # Сохранение вопросов
    question_extract = 1
    for i in range(1, (total_questions_count) + 1):
        exstract = driver.find_element_by_xpath(
            f'/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[4]/tbody/tr[2]/td[2]/table/tbody/tr[{question_extract}]/td[2]').text
        questions.append(exstract)
        if i == 40:
            try:
                driver.find_element_by_xpath(
                    '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[4]/tbody/tr[3]/td[2]/a[2]').click()
            except:
                break
            question_extract = 0
            sleep(1)
        if i == 80:
            driver.find_element_by_xpath(
                '/html/body/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table[4]/tbody/tr[3]/td[2]/a[3]').click()
            question_extract = 0
            sleep(1)
        question_extract += 1



pd.DataFrame(questions).to_excel()

