#bibliotecas usadas
import win32com.client as win32
import pandas as pd

#usar uma planilha
#!!pode ser de qualquer tipo!!
candidatos = pd.read_csv('Pasta.csv')
x = candidatos.loc[:,"email"].values #pega a parte dos email
y = candidatos.loc[:,"nome"].values #pega os nomes

#criar a integração
outlook = win32.Dispatch('outlook.application')

#criar um email no outlook
#lembrando que não a conta, essa você tem que criar sozinho.
email = outlook.CreateItem(0)

#configurações do e-mail.
#condição de repetição usada para enviar mais de um email.
j=0
for i in x:
    email = outlook.CreateItem(0)
    email.to = i
    email.Subject = "Aqui vai o título do email."
    email.HTMLBody =  f''' 
    Olá {y[j]} '''#Aqui vai os nomes certinhos
    email.Send()
    j += 1
print("Enviado") #criado para ser prova de que o envio ocorreu corretamente.
