from email import encoders
from email.mime.base import MIMEBase
from email.mime.message import MIMEMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

from selenium import webdriver
from time import sleep
import smtplib
import openpyxl
from openpyxl.styles import numbers, Alignment
from pathlib import Path
from selenium.common.exceptions import NoSuchElementException


def get_data():
    driver = webdriver.Chrome()
    driver.get("https://www.moex.com/")
    # Меню
    sleep(1)
    driver.find_element("xpath", "/html/body/div[1]/div[2]/div/header/div[1]/div/div/div/div[3]/button").click()
    # Срочный рынок
    sleep(1)
    driver.find_element("xpath", "/html/body/div[1]/div[2]/div[2]/header/div[3]/div[2]/div/div/div/ul/li[2]/a").click()
    # Согласен
    sleep(1)
    driver.find_element("xpath", "/html/body/div[1]/div/div/div/div/div[1]/div/a[1]").click()
    # Индикативные курсы
    sleep(1)
    driver.find_element("xpath", "/html/body/div[2]/div[4]/div/div/div[2]/div[1]/div/div[2]/div/div/div[2]").click()
    sleep(2)
    driver.find_element("xpath", "/html[1]/body[1]/div[2]/div[4]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[3]/div[18]/div[1]/a[1]/span[1]").click()
    # USD-RUB
    sleep(1)
    driver.find_element("xpath", "/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/form/div[1]/div[1]/span").click()
    sleep(1)
    driver.find_element("xpath", "/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[1]/div[18]/div/div[1]/a").click()
    sleep(2)


    date_raw = driver.find_element("xpath", "/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/strong").text
    ws = openpyxl.Workbook()
    wb = ws.active
    ws.active.title = "Sheet"
    # Сбор и вставка данных USD-RUB
    wb['A1'] = "Дата USD-RUB"
    wb['B1'] = "Курс USD-RUB"
    wb['C1'] = "Время USD-RUB"
    row = 1
    while True:
        try:
            value_date = driver.find_element("xpath", f"/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/div[2]/div[3]/table/tbody/tr[{row}]/td[1]").text
            value_rate = driver.find_element("xpath", f"/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/div[2]/div[3]/table/tbody/tr[{row}]/td[4]").text
            value_time = driver.find_element("xpath", f"/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/div[2]/div[3]/table/tbody/tr[{row}]/td[5]").text
        except NoSuchElementException:
            break
        wb[f'A{row + 1}'] = value_date
        wb[f'B{row + 1}'] = value_rate.replace('.', ',')
        wb[f'B{row + 1}'].number_format = '#,##0.00 "₽"'
        wb[f'C{row + 1}'] = value_time
        row += 1
    row_amount = row
    # JPY-RUB
    sleep(1)
    driver.find_element("xpath", "/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/form/div[1]/div[1]/span").click()
    sleep(1)
    driver.find_element("xpath", "/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[1]/div[8]/div/div[1]/a").click()
    sleep(2)

    # Сбор и вставка данных JPY/RUB
    wb['D1'] = "Дата JPY/RUB"
    wb['E1'] = "Курс JPY/RUB"
    wb['F1'] = "Время JPY/RUB"
    row = 1
    while True:
        try:
            value_date = driver.find_element("xpath", f"/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/div[2]/div[3]/table/tbody/tr[{row}]/td[1]").text
            value_rate = driver.find_element("xpath", f"/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/div[2]/div[3]/table/tbody/tr[{row}]/td[4]").text
            value_time = driver.find_element("xpath", f"/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/div[2]/div[3]/table/tbody/tr[{row}]/td[5]").text
        except NoSuchElementException:
            break
        wb[f'D{row + 1}'] = value_date
        wb[f'E{row + 1}'] = value_rate.replace('.', ',')
        wb[f'E{row + 1}'].number_format = '#,##0.00 "₽"'
        wb[f'F{row + 1}'] = value_time
        row += 1
    # Изменение Excel Файла
    wb['G1'] = "Результат"
    for row in range(1, row_amount):
        wb[f'G{row + 1}'] = f"=B{row + 1}/E{row+1}"
        for column in ["A", "B", "C", "D", "E", "F", "G"]:
            wb[f'{column}{row + 1}'].alignment = Alignment(horizontal="justify")


    ws.save(f"{Path.home()}/Downloads/{date_raw}.xlsx")
    ws.close()

    date_raw = "Данные с 11.09.2024 по 11.10.2024"
    row_amount = 24

    if row_amount % 10 == 1:
        msg_dop = "строка"
    elif row_amount % 10 < 5:
        msg_dop = "строки"
    else:
        msg_dop = "строк"

    file = f"{Path.home()}/Downloads/{date_raw}.xlsx"
    username = "fly.rev1ng@gmail.com"
    password = "liuu crya qyur nowd"
    send_from = "fly.rev1ng@gmail.com"
    send_to = 'AASkoybeda@greenatom.ru, NSGruzdev@Greenatom.ru'
    msg = MIMEMultipart()
    msg.attach(MIMEText(f"Файл Excel содержит {row_amount} {msg_dop}"))
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = f"Файл Excel {date_raw}"
    server = smtplib.SMTP('smtp.gmail.com')
    port = '587'
    fp = open(file, 'rb')
    part = MIMEBase('application', 'vnd.ms-excel')
    part.set_payload(fp.read())
    fp.close()
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=f"{date_raw}.xlsx")
    msg.attach(part)
    smtp = smtplib.SMTP('smtp.gmail.com')
    smtp.ehlo()
    smtp.starttls()
    smtp.login(username, password)
    smtp.sendmail(send_from, send_to.split(','), msg.as_string())
    smtp.quit()

get_data()