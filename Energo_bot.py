import requests
from bs4 import BeautifulSoup as bs
from docx import Document
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import date
import regex as re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

def connect_to_site_create_DF_crt_and_del_empty_raws():
    global df_tables_1
    global len_df_tables

    url = 'https://LINK'

    r = requests.get(url)

    soup_ing = bs(r.content, 'lxml')
    soup_ing = soup_ing.find('article', 'post medium')

    index = 0
    len_df_tables = 1
    df_tables = []

    for i in soup_ing.find_all('a'):
        
        print('!',i.get('href'))

        link_href = str(i.get('href')[27:-5])
        print(i.get('href'), '------------------------------------')
        energo_link = 'https://LINK' + i.get('href')
        energo_r = requests.get(energo_link)

                    
        with open(f'PATH{link_href}', 'wb+') as f:
            f.write(energo_r.content)

        file_name = f'PATH{link_href}'
        document = Document(f'{file_name}')
        tables = document.tables
        
        for table in tables:
            df = [['' for _ in range(len(table.columns))] for _ in range(len(table.rows))]
            for i, row in enumerate(table.rows):
                desired_row_index = i
                Flag = False
                for j, cell in enumerate(row.cells):
                    if j == 0 and cell.text.strip()[:2].isdigit():
                        Flag = True
                    if Flag:
                        desired_column_index = j
                        desired_cell_test = table.cell(desired_row_index, desired_column_index)
                        df[i][j] = desired_cell_test.text.strip()
                    else:
                        continue 
            df_tables.extend(df)
        
        index += 1
        if index == 2:
            break

    df_tables_1 = pd.DataFrame(df_tables)

def del_empty_lines_dont_relevant_inform_sort(df_tables_1):

    #Удаление пустых строк, неактуальной информации и сортировка ↓
    dropped_var = []

    cur_year_two_num = date.today().strftime('%y')
    cur_month = date.today().strftime('%m')
    cur_day = date.today().strftime('%d')

    for j in range(len(df_tables_1[0].index)):
        date_regex = re.findall(r'\d{2}\.\d{2}\.\d{4}|\d{2}\.\d{2}\.\d{2}', df_tables_1[0].values[j])
        date_regex = '.'.join(date_regex)

        date_regex_year = date_regex[-2:]
        date_regex_moth = date_regex[3:5]
        date_regex_day = date_regex[:2]
        
        if len(df_tables_1[0].values[j]) <= 5:
            dropped_var.append(j)

        if cur_year_two_num > date_regex_year:
            dropped_var.append(j)

        elif cur_month > date_regex_moth:
            dropped_var.append(j)
        if cur_day > date_regex_day:
            dropped_var.append(j)

    df_tables_1 = df_tables_1.drop(index = dropped_var)
    df_tables_1 = df_tables_1.sort_values(by = 0)

    return df_tables_1
    #Удаление пустых строк, неактуальной информации и сортировка ↑

def auth_to_google_sheets():
    global energobot 
    #Подключение и авторизации к google sheets ↓
    credentials_path = 'PATH.json'

    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
    client = gspread.authorize(credentials)

    energobot = client.open('Название').worksheet('Название')
    #Подключение и авторизации к google sheets ↑

def empty_rows():
    empty_rows_list = [[' ', ' ', ' ', ' ', ' ', ' ', ' ' , ' '] for _ in range(100)]
    empty_rows_list_pd = pd.DataFrame(empty_rows_list)
    return empty_rows_list_pd

def find_arrow_up_and_down(df_tables_1):
    global strela_down_id
    global strela_up_id
    global len_df_tables

    len_df_tables = len(df_tables_1.values.tolist())
    #Поиск символов "↓" и "↑" как начало строки для записи ↓

    if energobot.find('↓', in_column = 1):
        strela_down = energobot.find('↓', in_column = 1)
        strela_down_id = int('%d' % (strela_down.row))
        
    else:    
        energobot.update(f'A{25}:H{25}', [['↓', '-------', '-------', '-------', 'Ниже пишет Бот, Бот пишет ниже', '-------', '-------', '↓']])
        strela_down = energobot.find('↓', in_column = 1)
        strela_down_id = int('%d' % (strela_down.row))

    empty_rows_1 = empty_rows()

    energobot.update(f'A{strela_down_id + 1}', empty_rows_1.values.tolist())
    energobot.update(f'A{strela_down_id + 1}', df_tables_1.values.tolist())

    energobot.update(f'A{strela_down_id + len_df_tables + 1}:H{strela_down_id + len_df_tables + 1}', \
                    [['↑', '-------', '-------', '-------', 'Выше пишет Бот, Бот пишет Выше', '-------', '-------', '↑']])

    strela_up = energobot.find('↑', in_column = 1)
    strela_up_id = int('%d' % (strela_up.row))
    #Поиск символов "↓" и "↑" как начало строки для записи ↑

    #Поиск символа "↓" как начало строки для записи ↓
    if len(df_tables_1.index) == 0:
        energobot.update(f'A{strela_down_id + 1}:H{strela_down_id + 1}', [['', '', '', '', 'Нет актуальных отключений', '', '', '']])
        energobot.update(f'A{strela_down_id + 2}:H{strela_down_id + 2}', [['↑', '-------', '-------', '-------', 'Выше пишет Бот, Бот пишет Выше', '-------', '-------', '↑']])
        strela_up = energobot.find('↑')
        strela_up_id = int('%d' % (strela_up.row))
        energobot.delete_rows(strela_down_id + 2, strela_up_id - 1)
    #Поиск символа "↑" как начало строки для записи ↑

def write_to_sheets_who_added_an_antry():
    #Записывание информации из docx-файлов и добавление кто сделал запись(БОТ) ↓
    bot_site_date = [['Бот', 'Сайт', date.today().strftime('%d.%m.%y')] for _ in range(len(df_tables_1))]
    bot_site_date_pd = pd.DataFrame(bot_site_date)

    energobot.update(f'f{strela_down_id + 1}', bot_site_date_pd.values.tolist())
    #Записывание информации из docx-файлов и добавление кто сделал запись(БОТ) ↑

def formatting():
    #стрелка вниз
    energobot.format(f'A{strela_down_id}:H{strela_down_id}', 
    {"backgroundColor": {
        "red": 20,
        "green": 70,
        "blue": 200
        }, "horizontalAlignment": "CENTER", 
        "textFormat": {
        "foregroundColor": {
            "red": 255,
            "green": 255,
            "blue": 255
        },
        "fontSize": 11,
        }})	

    #стрелка вверх
    energobot.format(f'A{strela_up_id}:H{strela_up_id}', 
    {"backgroundColor": {
        "red": 20,
        "green": 70,
        "blue": 200
        }, "horizontalAlignment": "CENTER", 
        "textFormat": {
        "foregroundColor": {
            "red": 255,
            "green": 255,
            "blue": 255
        },
        }})

    #текст между стрелками
    energobot.format(f'A{strela_down_id + 1}:H{strela_up_id - 1}', 
    {"backgroundColor": {
        "red": 1,
        "green": 1,
        "blue": 1
        }, "horizontalAlignment": "left",
            "verticalAlignment": "top",  
        "textFormat": {
        "foregroundColor": {
            "red": 255,
            "green": 255,
            "blue": 255
        },
        }})

def write_to_file_and_send_to_email():

    # Настройки почтового сервера и учетной записи
    smtp_server = "domain_name"
    smtp_port = 587
    smtp_username = "email" 
    smtp_password = "pass"
    sender_email = "email"

    # Создание сообщения
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = 'email'
    msg['Subject'] = 'Отчет в формате XLSX'

    # Открытие и добавление файла XLSX в сообщение
    xlsx_file_path = 'PATH'
    with open(xlsx_file_path, 'rb') as xlsx_file:
        xlsx_attachment = MIMEApplication(xlsx_file.read(), _subtype="xlsx")
        xlsx_attachment.add_header('content-disposition', 'attachment', filename=xlsx_file_path)
        msg.attach(xlsx_attachment)

    # Инициализация SMTP-сервера и отправка сообщения
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(sender_email, msg['To'], msg.as_string())
        server.quit()
        return print("Сообщение успешно отправлено!")
    except Exception as e:
        return print(f"Ошибка при отправке сообщения: {str(e)}")

connect_to_site_create_DF_crt_and_del_empty_raws()
df_tables_1 = del_empty_lines_dont_relevant_inform_sort(df_tables_1) 
auth_to_google_sheets()
find_arrow_up_and_down(df_tables_1)
write_to_sheets_who_added_an_antry()
formatting()
write_to_file_and_send_to_email()