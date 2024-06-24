
#@title
## IMPORTANDO BIBLIOTECAS QUE SERÃO USADAS NA MACRO
import pandas as pd
import time
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from mimetypes import guess_type
from google.colab import drive

## IMPORTANDO PASTAS E ARQUIVOS DO DRIVE
drive.mount('/content/gdrive')

print('Início - ' + datetime.today().strftime('%H:%M'))
start_time = time.time()

## LENDO LISTAS DE TRANSPORTADORAS E E-MAILS (sempre deixar atualizado conforme a Skywalker)
transportadoras = pd.read_csv('/content/gdrive/Shareddrives/MM - Logística - TRANSPORTES/SERVIDOR/0-Bases_Nova/99. Python/Automações/to_send.csv', sep=';')
print('leu lista de transportadoras')
print(transportadoras)

import smtplib

#@title Informar E-mail
your_email = input('Enter your gmail: ')
print(f'You entered {your_email}')

#@title Informar Senha
your_password = input('Enter your password gmail: ')
print(f'You entered {your_password}')

## SERVIDOR QUE ENVIARÁ OS E-MAILS
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(your_email, your_password)
email_list = pd.read_csv('/content/gdrive/Shareddrives/MM - Logística - TRANSPORTES/SERVIDOR/0-Bases_Nova/99. Python/Automações/to_send.csv', sep = ';')
print(email_list)

## LENDO A LISTA DE E-MAILS
email_list['email'] = email_list['Lista Emails'].astype(str)
email_list['cc'] = email_list['cc'].astype('string')

print(email_list)

## LENDO A LISTA DE E-MAILS
names = email_list['Analista']
emails = email_list['email']
transportadora = email_list['Transp LM']
link = email_list['LINK SHEETS']
ccs = email_list['cc']
print('getting the names and the emails')

print('aqui ele começa o laço de enviar os emails')

for i in range(len(emails)):

        msg = MIMEMultipart()
        name = names[i]
        email = emails[i]
        assunto = transportadora[i]
        anexo = link[i]
        cc = ccs[i]
        if pd.notna(email):
            # the message to be emailed
            msg['Subject'] = assunto + ' - Base de Devolução e Devolução de Reversa'
            msg['Cc'] = cc
            msg['To'] = email
            ## CORPO DO E-MAIL
            body = '''
Olá, prezados, tudo bem? Espero que sim!

....
                    '''
            today = datetime.today().strftime('%m%d')
            #today='0630'
            html = anexo
            msg.attach(MIMEText(body, 'plain'))
            try:
                msg.attach(MIMEText(html, 'html'))
            except AttributeError:
                pass

            ## ENVIANDO EMAIL
            print('enviando e-mails')

            filename1 = f'/content/gdrive/Shareddrives/MM - Logística - TRANSPORTES/SERVIDOR/01. Perdas Operacionais 2021/Devoluções/Backups/Backup diario Devolucoes/Perdas_{assunto}_{today}.xlsx'
            filename2 = f'/content/gdrive/Shareddrives/MM - Logística - TRANSPORTES/SERVIDOR/01. Perdas Operacionais 2021/Devoluções/Backups/Backup diario Devolucoes/Perdas_{assunto}_{today}.pdf'
            print('pegando arquivos da base de devolução')
            try:
                attachment1 = open(filename1, "rb")
                attachment2 = open(filename2, "rb")
            except FileNotFoundError:
                pass
            print('abrindo arquivos da base de devolução')

            p1 = MIMEBase('application', 'octet-stream')
            p2 = MIMEBase('application', 'octet-stream')


            p1.set_payload((attachment1).read())
            p2.set_payload((attachment2).read())
            print('carregando no anexo arquivos da base de devolução')

            encoders.encode_base64(p1)
            encoders.encode_base64(p2)


            p1.add_header('Content-Disposition', "attachment1", filename=filename1, sep=';')
            p2.add_header('Content-Disposition', "attachment2", filename=filename2, sep=';')


            msg.attach(p1)
            msg.attach(p2)


            s = smtplib.SMTP('smtp.gmail.com', 587)
            s.starttls()
            #s.login(your_email, "dzepkxcqublsiuaf")
            s.login(your_email, your_password)
            text = str(msg)
            # try:
            #     text = msg.as_string()
            # except ValueError:
            #     pass
            s.sendmail(your_email, email.split(",") + cc.split(";"), text)
            s.quit()

            print('E-mail para ' + transportadora[i].__str__() + ' foi enviado')

# close the smtp server
server.close()

print('Fim - ' + datetime.today().strftime('%H:%M'))
