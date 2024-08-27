# https://mailtrap.io/blog/outlook-smtp/
# https://www.youtube.com/watch?v=CgvOCXs8Xec

import warnings
warnings.filterwarnings('ignore')

import os
import creds
import smtplib
import requests
import pandas as pd
from bs4 import BeautifulSoup
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def data_get(cidade = 'jundiai'):
    nome_cidade = cidade
    link_cidade = f"http://servicos.cptec.inpe.br/XML/listaCidades?city={nome_cidade}"
    response_cidade = requests.get(link_cidade)
    xml_cidade = BeautifulSoup(response_cidade.text, features = 'lxml')
    id_cidade = xml_cidade('id')[0].text

    this_week = requests.get(f"http://servicos.cptec.inpe.br/XML/cidade/7dias/{id_cidade}/previsao.xml")
    next_week = requests.get(f"http://servicos.cptec.inpe.br/XML/cidade/{id_cidade}/estendida.xml")

    previsoes_this_week = BeautifulSoup(this_week.text, features = 'lxml').find_all('previsao')
    previsoes_next_week = BeautifulSoup(next_week.text, features = 'lxml').find_all('previsao')
    previsoes = previsoes_this_week + previsoes_next_week

    resultados = []
    for dia in previsoes:
        resultados.append([dia('dia')[0].text, dia('maxima')[0].text, dia('minima')[0].text])
    df_results = pd.DataFrame(resultados, columns=['date', 'max', 'min'])
    df_results.to_csv(r'C:\Users\piato\Desktop\previsao.csv', index=False)
    print("Cidade escolhida: ", nome_cidade.capitalize(), '\n\n', df_results)

    return df_results


def email_new(df):
    message = MIMEMultipart()
    message['Subject'] = "New Data from Today"
    message['From'] = creds.sender
    message['To'] = creds.recipient

    html = MIMEText(df.to_html(index=False), "html")
    message.attach(html)
    
    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.starttls()
        server.login(creds.sender, creds.password)
        server.sendmail(creds.sender, creds.recipient, message.as_string())

    print('E-mail enviado com sucesso!')


if __name__ == '__main__':
    data = data_get()
    email_new(data)