from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

from selenium import webdriver
from time import sleep
import smtplib
import openpyxl
from openpyxl.styles import Alignment
from pathlib import Path
from selenium.common.exceptions import NoSuchElementException

row_amount = 0
date_raw = ""
def get_data():
    # Получение данных с сайта и внесение данных в файл Excel


    driver = webdriver.Chrome()
    # Открытие сайта (пункт 1)
    driver.get("https://www.moex.com/")
    # Выбор элемента Меню (пункт 2)
    sleep(1)
    driver.find_element("xpath", "/html/body/div[1]/div[2]/div/header/div[1]/div/div/div/div[3]/button").click()
    # Выбор элемента Срочный рынок (пункт 2)
    sleep(1)
    driver.find_element("xpath", "/html/body/div[1]/div[2]/div[2]/header/div[3]/div[2]/div/div/div/ul/li[2]/a").click()
    # Выбор элемента Согласен (пункт 2)
    sleep(1)
    driver.find_element("xpath", "/html/body/div[1]/div/div/div/div/div[1]/div/a[1]").click()
    # Выбор элемента Индикативные курсы (пункт 2)
    sleep(1)
    driver.find_element("xpath", "/html/body/div[2]/div[4]/div/div/div[2]/div[1]/div/div[2]/div/div/div[2]").click()
    sleep(2)
    driver.find_element("xpath", "/html[1]/body[1]/div[2]/div[4]/div[1]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div[1]/div[3]/div[18]/div[1]/a[1]/span[1]").click()
    # Выбор курса валют USD-RUB (пункт 3-4)
    sleep(1)
    driver.find_element("xpath", "/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/form/div[1]/div[1]/span").click()
    sleep(1)
    driver.find_element("xpath", "/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[1]/div[18]/div/div[1]/a").click()
    sleep(2)

    # Открытие Excel файла
    global date_raw
    date_raw = driver.find_element("xpath", "/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/strong").text
    ws = openpyxl.Workbook()
    wb = ws.active
    ws.active.title = "Sheet"
    # Сбор и вставка данных USD-RUB (пункт 5)
    wb['A1'] = "Дата USD-RUB"
    wb['B1'] = "Курс USD-RUB"
    wb['C1'] = "Время USD-RUB"
    row = 1
    while True:
        try:
            value_date = driver.find_element("xpath", f"/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/div[2]/div[3]/table/tbody/tr[{row}]/td[1]").text
            value_rate = driver.find_element("xpath", f"/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/div[2]/div[3]/table/tbody/tr[{row}]/td[4]").text
            value_time = driver.find_element("xpath", f"/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/div[2]/div[3]/table/tbody/tr[{row}]/td[5]").text
        except NoSuchElementException: # Проверка наличия элемента для остановки считывания данных
            break
        wb[f'A{row + 1}'] = value_date
        wb[f'B{row + 1}'] = value_rate.replace('.', ',')
        wb[f'B{row + 1}'].number_format = '#,##0.00 "₽"' # Изменение формата ячейки
        wb[f'C{row + 1}'] = value_time
        row += 1
    global row_amount
    row_amount = row # Количество строк в файле
    # Выбор курса валют JPY-RUB (пункт 6)
    sleep(1)
    driver.find_element("xpath", "/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/form/div[1]/div[1]/span").click()
    sleep(1)
    driver.find_element("xpath", "/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[1]/div[8]/div/div[1]/a").click()
    sleep(2)

    # Сбор и вставка данных JPY/RUB (пункт 7)
    wb['D1'] = "Дата JPY/RUB"
    wb['E1'] = "Курс JPY/RUB"
    wb['F1'] = "Время JPY/RUB"
    row = 1
    while True:
        try:
            value_date = driver.find_element("xpath", f"/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/div[2]/div[3]/table/tbody/tr[{row}]/td[1]").text
            value_rate = driver.find_element("xpath", f"/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/div[2]/div[3]/table/tbody/tr[{row}]/td[4]").text
            value_time = driver.find_element("xpath", f"/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/div[2]/div[3]/table/tbody/tr[{row}]/td[5]").text
        except NoSuchElementException: # Проверка наличия элемента для остановки считывания данных
            break
        wb[f'D{row + 1}'] = value_date
        wb[f'E{row + 1}'] = value_rate.replace('.', ',')
        wb[f'E{row + 1}'].number_format = '#,##0.00 "₽"' # Изменение формата ячейки
        wb[f'F{row + 1}'] = value_time
        row += 1
    # Запись формул для подсчёта результата и выравнивание ячеек по ширине
    wb['G1'] = "Результат"
    for row in range(1, row_amount):
        wb[f'G{row + 1}'] = f"=B{row + 1}/E{row+1}"
        for column in ["A", "B", "C", "D", "E", "F", "G"]:
            wb[f'{column}{row + 1}'].alignment = Alignment(horizontal="justify")


    ws.save(f"{Path.home()}/Downloads/{date_raw}.xlsx") # Сохрание файла
    ws.close()
def send_message():
    # Отправка письма (пункт 12-13)

    # Обработка верного склонения для содержимого письма
    if row_amount % 10 == 1:
        msg_dop = "строка"
    elif row_amount % 10 < 5:
        msg_dop = "строки"
    else:
        msg_dop = "строк"

    file = f"{Path.home()}/Downloads/{date_raw}.xlsx"
    username = "fly.rev1ng@gmail.com" # Логин
    password = "liuu crya qyur nowd" # Пароль придожения
    send_from = "fly.rev1ng@gmail.com" # Почта отправителя
    send_to = 'AASkoybeda@greenatom.ru, NSGruzdev@Greenatom.ru' # Почты получателей
    msg = MIMEMultipart()
    msg.attach(MIMEText(f"Файл Excel содержит {row_amount} {msg_dop}")) # Добавление текста письма
    msg['From'] = send_from # Добавление отправителя
    # msg['To'] = send_to # Добавление получателей
    msg['To'] = "inbo-07-21@sumirea.ru" # Добавление получателя(для отладки)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = f"Файл Excel {date_raw}" # Добавление темы
    fp = open(file, 'rb')
    part = MIMEBase('application', 'vnd.ms-excel')
    part.set_payload(fp.read())
    fp.close()
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=f"{date_raw}.xlsx")
    msg.attach(part) # Прикрепление файла Excel
    smtp = smtplib.SMTP('smtp.gmail.com')
    smtp.ehlo()
    smtp.starttls()
    smtp.login(username, password) # Вход программы в почту
    smtp.sendmail(send_from, send_to.split(','), msg.as_string()) # Отправка сообщения
    smtp.quit()

def main():
    get_data()
    send_message()
main()