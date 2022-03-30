# Biblioteca que faz interação com o outlook
import win32com.client as win32

# Criando integração do python com outlook
outlook = win32.Dispatch('outlook.application')

# Criando email
email = outlook.CreateItem(0)

# Configurar as informações do e-mail
email.To = "dev.MarceloP@hotmail.com; marcelo.prodrigues@hotmail.com"
email.Subject = "Email Python"
email.HTMLBody = f"""
<p> Olá email enviado automaticamente pelo pyhton</p>
<p>Não responda esse email </p>
""" 

# Anexando arquivos
anexo = "C://Users/USUARIO/Desktop/meme.png"
email.Attachments.Add(anexo)

# Enviando email
email.Send()

print("Email enviado!")