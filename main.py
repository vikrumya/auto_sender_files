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
# from watchdog.events import LoggingEventHandler
from watchdog.events import FileSystemEventHandler
def send(filepath):
    wb = load_workbook('база_рассылки.xlsx')
    """ЗАполнение комбобокса организациями"""
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
        server = 'smtp.mail.ru'
        user = 'dor.freza@mail.ru'
        password = 'be5vZsKMnW1IvLagx4ro'
        print(mail)
        recipients = mail
        sender = 'dor.freza@mail.ru'
        subject = 'Информационное письмо'
        text = 'Оповещаем Вас, как нашего постоянного клиента, что из-за подорожания стоимости горюче-смазочных материалов, запасных частей и услуг по техническому обслуживанию техники, мы были вынуждены увеличить стоимость наших транспортных услуг с 01.08.2021 года.<br>' \
               '- Доставка дорожной фрезы WIRTGEN2000, WIRTGEN200, CATERPILLAR PM60 в черте города до объекта работ и обратно составит <b>36 000</b> (тридцать шесть тысяч) рублей, в том числе НДС (20%).<br>' \
               '- Доставка дорожной фрезы WIRTGEN1000F, WIRTGEN100F в черте города до объекта работ и обратно составит <b>30 000</b> (тридцать тысяч) рублей, в том числе НДС (20%).<br>' \
               '- Дежурство тягача с тралом (ожидание окончания работ на  одном объекте работ и перевозка на следующий объект)  составит <b>30 000</b> (тридцать тысяч) рублей, в том числе НДС (20%).<br>' \
               '<br>Просим отнестись к этой вынужденной мере с пониманием.'
        html = '<html><head></head><body><p>' + text + '</p></body></html>'
        msg = MIMEMultipart()

        message = html
        #sender = "dor.freza@mail.ru"
        # setup the parameters of the message
        #password = "your_password"
        msg['From'] = sender
        msg['To'] = mail
        msg['Subject'] = "Информационное письмо от ГлавАвтоснаб"
        # basename = os.path.basename(filepath)
        # filesize = os.path.getsize(filepath)
        # part_file = MIMEBase('application', 'octet-stream; name="{}"'.format(basename))
        # part_file.set_payload(open(filepath, "rb").read())
        # part_file.add_header('Content-Description', basename)
        # part_file.add_header('Content-Disposition', 'attachment; filename="{}"; size={}'.format(basename, filesize))
        # encoders.encode_base64(part_file)
        #add in the message body
        #msg.attach(MIMEText(message, 'plain'))
        #part_text = MIMEText(text, 'plain')
        part_html = MIMEText(html, 'html')
        #
        #msg.attach(part_text)
        msg.attach(part_html)
        with open(filepath, "rb") as f:
         #attach = email.mime.application.MIMEApplication(f.read(),_subtype="pdf")
            attach = MIMEApplication(f.read())
            attach.add_header('Content-Disposition', 'attachment', filename=str(filepath))
        msg.attach(attach)
        #msg.attach(part_file)
        # # msg.attach(part_file)
        # # create server
        server = smtplib.SMTP('smtp.mail.ru: 587')

        server.starttls()

        # Login Credentials for sending the mail
        server.login(msg['From'], password)

        # send the message via the server.
        server.sendmail(msg['From'], msg['To'], msg.as_string())

        server.quit()
        print("successfully sent email to %s:" % (msg['To']))
        # filepath = "fish.png"
        # basename = os.path.basename(filepath)
        # filesize = os.path.getsize(filepath)
        #
        # msg = MIMEMultipart('alternative')
        # msg['Subject'] = subject
        # msg['From'] = 'ГлавАвтоснаб <' + sender + '>'
        # msg['To'] = ', '.join(recipients)
        # msg['Reply-To'] = sender
        # msg['Return-Path'] = sender
        # msg['X-Mailer'] = 'Python/' + (python_version())
        #
        # part_text = MIMEText(text, 'plain')
        # part_html = MIMEText(html, 'html')
        # part_file = MIMEBase('application', 'octet-stream; name="{}"'.format(basename))
        # part_file.set_payload(open(filepath, "rb").read())
        # part_file.add_header('Content-Description', basename)
        # part_file.add_header('Content-Disposition', 'attachment; filename="{}"; size={}'.format(basename, filesize))
        # encoders.encode_base64(part_file)
        #
        # msg.attach(part_text)
        # msg.attach(part_html)
        # msg.attach(part_file)
        #
        # mail = smtplib.SMTP_SSL(server)
        # mail.login(user, password)
        # mail.sendmail(sender, recipients, msg.as_string())
        # mail.quit()
        # print('Сообщения на e-mail',recipients, ' отправлено!')

def get_file(filepath):
    filepath = filepath
    print(filepath)
    send(filepath)


class EventHandler(FileSystemEventHandler):
    # вызывается на событие создания файла или директории
    def on_created(self, event):
        print(event.event_type, event.src_path)
        #f = open(event.src_path, "w")
        #f.write('ХУЙ')
        #f.close()
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
    # logging.basicConfig(level=logging.INFO,
    #                    format='%(asctime)s - %(message)s',
    #                    datefmt='%Y-%m-%d %H:%M:%S')

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




