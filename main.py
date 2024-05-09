#BIBLIOTECA PANDAS PARA MANIPULAR ARQUIVOS DO EXCEL
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

#BIBLIOTECA PANDAS PARA LER O ARQUIVO XLSX E TAMBÉM INDICA O NOME DA NOSSA PLANILHA EXCEL
#VOCE DEVE INDICAR O ARQUIVO EXCEL QUE CONTEM A LISTA COM NOME E EMAIL DOS CLIENTES
clientes = pd.read_excel('.\clientes.xlsx')

#PERCORRE CADA LINHA DE CLIENTES NA PLANILHA
for index, cliente in clientes.iterrows():

#PARA CADA LINHA DE CLIENTE É CRIADA UMA MSG DE TEXTO    
    msg = MIMEMultipart()
    msg['subject'] = 'Voce recebeu um email de fulano' #ASSUNTO DO EMAIL
    msg['From'] = 'Digite seu email aqui' #EMAIL QUE VAI ENVIAR
    msg['To'] = cliente['email'] #BUSCA NA PLANILHA O EMAIL QUE IRÁ RECEBER A MSG NA ORDEM DA LISTA
    message = f"Olá {cliente['nome']}, Olá tudo bem com você?" #NOME DO CLIENTE E MENSAGEM DO EMAIL
    msg.attach(MIMEText(message, 'plain'))

    server = smtplib.SMTP('smtp.gmail.com', port=587) #SERVIDOR DO GMAIL PARA CONEXÃO E ENVIO DE EMAILS
    server.starttls() #STARTA O SERVIDOR PARA SE CONECTAR COM A PORTA 587
    server.login('Digite seu email aqui', 'sua senha aqui') #GMAIL E SENHA PARA CONECTAR NO EMAIL 
    server.sendmail(msg['From'], msg['To'], msg.as_string()) #DISPARA O EMAIL E MOSTRA QUEM ESTÁ ENVIANDO E QUEM ESTÁ RECEBENDO
    server.quit() #PARA DESCONECTAR DO SERVIDOR



    

