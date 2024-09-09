# https://mailtrap.io/blog/outlook-smtp/
# https://www.youtube.com/watch?v=CgvOCXs8Xec
# https://mailtrap.io/blog/python-send-email/
# https://medium.com/@tempmailwithpassword/automating-email-attachments-in-outlook-with-python-a07224047434

import warnings
warnings.filterwarnings('ignore')

import time
import requests
import pandas as pd
import win32com.client
from bs4 import BeautifulSoup
import plotly.graph_objects as go


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
        resultados.append([dia('dia')[0].text, dia('maxima')[0].text, dia('minima')[0].text, dia('tempo')[0].text.strip()])
    df_results = pd.DataFrame(resultados, columns=['date', 'max', 'min', 'tempo'])
    df_results['date'] = pd.to_datetime(df_results['date'])
    df_results = df_results.astype({'max':'int', 'min':'int'})
    df_results.to_csv(r'C:\Users\piato\OneDrive\Área de Trabalho\previsao.csv', index=False)
    print("Cidade escolhida: ", nome_cidade.capitalize(), '\n\n', df_results)
    return df_results

def send_email():
    sender = "bruno.ipynb@outlook.com"
    recipient = "piatobio@gmail.com"
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = "Previsão do tempo para a semana."
    mail.Body = f"{time.asctime()}: Previsão do tempo para a semana.\nAs siglas das condições do tempo podem ser encontradas em: http://servicos.cptec.inpe.br/XML/#condicoes-tempo"
    attachments = [r'C:\Users\piato\OneDrive\Área de Trabalho\previsao.csv', r'C:\Users\piato\OneDrive\Área de Trabalho\fig1.png']
    for attachment in attachments:
        mail.Attachments.Add(attachment)
    mail.Send()

def draw_graph(data):
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=data['date'], y=data['max'],
                        mode='lines+markers',
                        name='max'))
    fig.add_trace(go.Scatter(x=data['date'], y=data['min'],
                        mode='lines+markers',
                        name='min'))
    fig.update_layout(template="plotly_dark", title="Temperatura máx e min")
    fig.show()

    fig.write_image(r"C:\Users\piato\OneDrive\Área de Trabalho\fig1.png")

if __name__ == "__main__":
    data = data_get()
    draw_graph(data)
    send_email()