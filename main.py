from __future__ import print_function
import os.path
import base64
import qrcode
import uuid
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from io import BytesIO
from PySimpleGUI import PySimpleGUI as sg

sg.theme('Black')
layout = [

    [sg.Text('GERAR QRCODE SAECOMP', justification='center', size=(52, 1), font=('Any', 18, 'bold'), text_color='white', background_color='green')],
    
    [sg.Text('Email Saecomp:', size=(50, 1), font=('Any', 12)), sg.Input(key='emailsaecomp')],
    [sg.Text('Assunto do Email:', size=(50, 1), font=('Any', 12)), sg.Input(key='assunto')],
    [sg.Text('Mensagem do email:', size=(50, 1), font=('Any', 12)), sg.Input(key='mensagem')],
    [sg.Text('Nome da aba da planilha:', size=(50, 1), font=('Any', 12)), sg.Input(key='aba')],
    [sg.Text('Coluna que deve ser verificada (ex.: A):', size=(50, 1), font=('Any', 12)), sg.Input(key='coluna')],
    [sg.Text('Coluna que ficará o código (ex.: B):', size=(50, 1), font=('Any', 12)), sg.Input(key='colunaqrcode')],
    [sg.Text('Coluna que contém o email do comprador (ex.: C):', size=(50, 1), font=('Any', 12)), sg.Input(key='colunaemail')],
    [sg.Text('Código da planilha (docs.google.com/spreadsheets/d/ "código"/):', size=(50, 1), font=('Any', 12)), sg.Input(key='codigo')],
    [sg.Button('OK', size=(10, 1))]
]

janela = sg.Window('SAECOMP', layout)


def create_message_with_attachment(sender, to, subject, message_text, file_data, filename):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    msg = MIMEText(message_text)
    message.attach(msg)

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(file_data)
    encoders.encode_base64(part)
    part.add_header(
        'Content-Disposition',
        f'attachment; filename={filename}',
    )
    message.attach(part)

    raw = base64.urlsafe_b64encode(message.as_bytes())
    raw = raw.decode()
    return {'raw': raw}

def send_message(service, user_id, message):
    try:
        message = service.users().messages().send(userId=user_id, body=message).execute()
        return message
    except Exception as error:
        sg.popup(f'Um erro ocorreu: {error}')
        return None

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/gmail.send']


def main():
    while True:
        eventos, valores = janela.Read()
        if eventos == sg.WIN_CLOSED:
            break
        if eventos == 'OK':

            janela.hide()

            email_saecomp = valores['emailsaecomp']
            assunto = valores['assunto']
            mensagem = valores['mensagem']
            range_name = f'{valores['aba']}!{(valores['coluna']).upper()}:{(valores['coluna']).upper()}'  
            colunaemail = (valores['colunaemail']).upper()


            creds = None
            if os.path.exists('token.json'):
                creds = Credentials.from_authorized_user_file('token.json', SCOPES)

            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                
                with open('token.json', 'w') as token:
                    token.write(creds.to_json())

            service_sheets = build('sheets', 'v4', credentials=creds)

            spreadsheet_id = valores['codigo']

            sheet = service_sheets.spreadsheets()
            result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
            values = result.get('values', [])

            service_gmail = build('gmail', 'v1', credentials=creds)

            for i, row in enumerate(values):
                if row and row[0] != 'Conferiu Banco?' and row[0] != '':
                    unique_code = str(uuid.uuid4())
                    data = unique_code

                    qr = qrcode.QRCode(
                        version=1,
                        error_correction=qrcode.constants.ERROR_CORRECT_L,
                        box_size=10,
                        border=4,
                    )

                    qr.add_data(data)
                    qr.make(fit=True)

                    buf = BytesIO()
                    img = qr.make_image(fill_color="black", back_color="white")
                    img.save(buf, format='PNG')
                    buf.seek(0)

                    #cadastrar o qrcode na planilha
                    celula_range = f'{valores['aba']}!{valores['colunaqrcode']}{i + 1}'
                    receber = [[data]]
                    update_result = sheet.values().update(
                        spreadsheetId=spreadsheet_id,
                        range=celula_range,
                        valueInputOption="RAW",
                        body={"values": receber}).execute()
                
                    #pegar email da planilha
                    email_range = '{}!{}{}'.format(valores['aba'], colunaemail, i+1)
                    email_result = sheet.values().get(spreadsheetId=spreadsheet_id,range=email_range).execute()
                    if 'values' in email_result:
                        emailrecebedor = email_result['values'][0][0]

                    message = create_message_with_attachment(

                        email_saecomp, emailrecebedor, assunto, mensagem, buf.read(), 'qrcode.png'
                    )
                    send_message(service_gmail, 'me', message)

if __name__ == '__main__':
    main()
