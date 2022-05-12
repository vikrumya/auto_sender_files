import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from platform import python_version
import openpyxl
from email.mime.application import MIMEApplication
from openpyxl import load_workbook
from watchdog.observers import Observer
import time
from watchdog.events import FileSystemEventHandler
def send(filepath):
    wb = load_workbook('база_рассылки.xlsx')
    sheet_ranges = wb['база']
    column_a = sheet_ranges['A']
    email = []
    for i in range(len(column_a)):
        if column_a[i].value != None:
            print(column_a[i].value)
            email.append(column_a[i].value)
    print(email)
    email.pop(0)
    print(email)
    input()
    for mail in email:
        server = 'smtp.*********.ru'
        user = '*********'
        password = '*********'
        print(mail)
        recipients = mail
        sender = '*********@email.ru'
        subject = 'Информационное письмо'
        text = '*********'
        html = '<html><head></head><body><p>' + text + '</p></body></html>'
        msg = MIMEMultipart()

        message = html
        #sender = "dor.freza@mail.ru"
        # setup the parameters of the message
        #password = "your_password"
        msg['From'] = sender
        msg['To'] = mail
        msg['Subject'] = "Информационное письмо от *********"
        part_html = MIMEText(html, 'html')
        msg.attach(part_html)
        with open(filepath, "rb") as f:
            attach = MIMEApplication(f.read())
            attach.add_header('Content-Disposition', 'attachment', filename=str(filepath))
        msg.attach(attach)
        server = smtplib.SMTP('smtp.mail.ru: 587')

        server.starttls()

        # Login Credentials for sending the mail
        server.login(msg['From'], password)

        # send the message via the server.
        server.sendmail(msg['From'], msg['To'], msg.as_string())

        server.quit()
        print("successfully sent email to %s:" % (msg['To']))

def get_file(filepath):
    filepath = filepath
    print(filepath)
    send(filepath)


class EventHandler(FileSystemEventHandler):
    # вызывается на событие создания файла или директории
    def on_created(self, event):
        print(event.event_type, event.src_path)
        get_file(event.src_path)

    # вызывается на событие модификации файла или директории
    def on_modified(self, event):
        print(event.event_type, event.src_path)

    # вызывается на событие удаления файла или директории
    def on_deleted(self, event):
        print(event.event_type, event.src_path)

    # вызывается на событие перемещения\переименования файла или директории
    def on_moved(self, event):
        print(event.event_type, event.src_path, event.dest_path)


if __name__ == "__main__":
    path = r"C:\test"  # отслеживаемая директория с нужным файлом
    event_handler = EventHandler()
    observer = Observer()
    observer.schedule(event_handler, path, recursive=True)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()




