import pandas as pd
import datetime
import yfinance as yf
from matplotlib import pyplot as plt
import mplcyberpunk
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.mime.audio import MIMEAudio
from email import encoders
from IPython.display import display
import mimetypes
import os
#import win32com.client as win32 #para enviar email pelo outlook 

#Pegando dados do Yahoo Finance
codigos_de_negociacao = ["^BVSP", "BRL=X"]  
hoje = datetime.datetime.now()
um_ano_atras = hoje - datetime.timedelta(days = 365)
dados_mercado = yf.download(codigos_de_negociacao, um_ano_atras, hoje)
#display(dados_mercado)

#Tratando os dados
dados_fechamento = dados_mercado['Adj Close']
dados_fechamento.columns = ['dolar', 'ibovespa']
dados_fechamento = dados_fechamento.dropna()
dados_fechamento

#Manipulando os dados 
dados_anuais = dados_fechamento.resample("Y").last()
dados_mensais = dados_fechamento.resample("M").last()
dados_anuais

#Calculando rentabilidade 
retorno_anual = dados_anuais.pct_change().dropna()
retorno_mensal = dados_mensais.pct_change().dropna()
retorno_diario = dados_fechamento.pct_change().dropna()
retorno_diario

retorno_diario_dolar = retorno_diario.iloc[-1, 0]
retorno_diario_ibov = retorno_diario.iloc[-1, 1]
retorno_mensal_dolar = retorno_mensal.iloc[-1, 0]
retorno_mensal_ibov = retorno_mensal.iloc[-1, 1]
retorno_anual_dolar = retorno_anual.iloc[-1, 0]
retorno_anual_ibov = retorno_anual.iloc[-1, 1]
print(retorno_anual_dolar)
#display(retorno_anual)

retorno_diario_dolar = round((retorno_diario_dolar * 100), 2)
retorno_diario_ibov = round((retorno_diario_ibov * 100), 2)
retorno_mensal_dolar = round((retorno_mensal_dolar * 100), 2)
retorno_mensal_ibov = round((retorno_mensal_ibov * 100), 2) 
retorno_anual_dolar = round((retorno_anual_dolar * 100), 2)
retorno_anual_ibov = round((retorno_anual_ibov * 100), 2)

#Graficos de performace 
plt.style.use("cyberpunk")
dados_fechamento.plot(y = "ibovespa", use_index = True, legend = False)
plt.title("Ibovespa")
plt.savefig('ibovespa.png', dpi = 300)
plt.show()

plt.style.use("cyberpunk")
dados_fechamento.plot(y = "dolar", use_index = True, legend = False)
plt.title("Dolar")
plt.savefig('dolar.png', dpi = 300)
plt.show()

#Enviar Email 
def enviaemail(emails): 
  host = 'smtp.gmail.com'
  port = '587'
  login = 'email@gmail.com'
  senha = 'exukgesaegmqfaoa' #gerar senha de aplicativo no gerenciamento de conta do google
  
  server = smtplib.SMTP(host,port)
  server.ehlo()
  server.starttls()
  server.login(login,senha)
  
  corpo = f'''Prezado diretor, segue o relatório diário:

  Bolsa:

  No ano o Ibovespa está tendo uma rentabilidade de {retorno_anual_ibov}%, 
  enquanto no mês a rentabilidade é de {retorno_mensal_ibov}%.

  No último dia útil, o fechamento do Ibovespa foi de {retorno_diario_ibov}%.

  Dólar:

  No ano o Dólar está tendo uma rentabilidade de {retorno_anual_dolar}%, 
  enquanto no mês a rentabilidade é de {retorno_mensal_dolar}%.

  No último dia útil, o fechamento do Dólar foi de {retorno_diario_dolar}%.

  Abs,

  O melhor estagiário do mundo

  '''
  
  email_msg = MIMEMultipart()
  email_msg['From'] =  'email@gmail.com' #Quem está mandando o email
  email_msg['To'] = emails #Substitui um email fixo pelo vindo na função.
  email_msg['Subject'] = 'Relatório Diário'
  email_msg.attach(MIMEText(corpo,'plain'))
  
  def adiciona_anexo(email_msg, filename):
    if not os.path.isfile(filename):
        return
    ctype, encoding = mimetypes.guess_type(filename)
    
    if ctype is None or encoding is not None:
        ctype = '/mnt/1b699805-d31a-4111-9978-67c97cff7271/Programação/Varos Bootcamp/FechamentoMercado'
    maintype, subtype = ctype.split('/', 1)

    if maintype == 'text':
        with open(filename) as f:
            mime = MIMEText(f.read(), _subtype=subtype)
    elif maintype == 'image':
        with open(filename, 'rb') as f:
            mime = MIMEImage(f.read(), _subtype=subtype)
    elif maintype == 'audio':
        with open(filename, 'rb') as f:
            mime = MIMEAudio(f.read(), _subtype=subtype)
    else:
        with open(filename, 'rb') as f:
            mime = MIMEBase(maintype, subtype)
            mime.set_payload(f.read())

        encoders.encode_base64(mime)

    mime.add_header('Content-Disposition', 'attachment', filename=filename)
    email_msg.attach(mime)
    
  adiciona_anexo(email_msg, 'ibovespa.png')
  adiciona_anexo(email_msg, 'dolar.png')
   
  server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
  server.quit()
    
emails = ['email@gmail.com']

for i in emails:
     enviaemail(i)