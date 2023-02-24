import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from test import *


def send_email(message):
    mail = 'tesmtpy@gmail.com'
    password = 'bwsgjqcrpjdryfss'
    recipient = input('Введите почту получателя: ')
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()

    # количество строк Excel в правильном склонении
    # проверка на количество
    def count_filled_rows(filename, sheetname):
        workbook = openpyxl.load_workbook(filename)
        sheet = workbook[sheetname]
        count = 0
        for row in sheet.iter_rows():
            if any(cell.value for cell in row):
                count += 1
        return count

    count = count_filled_rows('example.xlsx', 'Sheet')

    # Склонение
    def count_strings(count):
        if count % 10 == 1 and count != 11:
            return f"{count} строка"
        elif 2 <= count % 10 <= 4 and (count < 10 or count > 20):
            return f"{count} строки"
        else:
            return f"{count} строк"

    skl = count_strings(count - 2)

    # Оформление письма + вложение файла
    try:
        server.login(mail, password)
        msg = MIMEMultipart()
        msg['Subject'] = 'Тестовое задание от Дегтярева В.А.'
        msg.attach(MIMEText(
            'Отчет за предыдущий месяц: ' + prev_month_name + ', по курсам валют USD/RUB и JPY/RUB\nСодержит: ' + skl))
        with open('example.xlsx', 'rb') as f:
            file = MIMEApplication(f.read(), _subtype='xlsx')
            file.add_header('content-disposition', 'attachment', filename='example.xlsx')
            msg.attach(file)
        server.sendmail(mail, recipient, msg.as_string())
        return 'Сообщение успешно отправлено!\nЕсли письма нет, проверьте папку СПАМ!'
    except Exception as _ex:
        return f'{_ex}\nПожалуйста, проверьте ваш логин или пароль!'


def main():
    print(send_email('t'))


if __name__ == '__main__':
    main()
