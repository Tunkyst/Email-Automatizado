#usar uma planilha
#!!pode ser de qualquer tipo!!
candidatos = pd.read_csv('Pasta.csv')
x = candidatos.loc[:,"email"].values


#criar a integração
outlook = win32.Dispatch('outlook.application')

#criar um email no outlook
#lembrando que não a conta, essa você tem que criar sozinho.
email = outlook.CreateItem(0)

#configurações do e-mail.
#condição de repetição usada para enviar mais de um email.
for i in x:
    email = outlook.CreateItem(0)
    email.to = i
    email.Subject = "Aqui vai o título do email."
    email.HTMLBody = """ 
    Aqui vai o texto do email.
    Caso já possua um email pronto: 
        -Inspecione a página que está o email.
        -Procure a parte que engloba o texto.
        -Copie e cole aqui.
        """
    email.Send()
print("Enviado") #criado para ser prova de que o envio ocorreu corretamente.