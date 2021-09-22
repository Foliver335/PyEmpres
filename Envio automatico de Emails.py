
# Criação de um Auto Reporter


import win32com.client as win32

Outlook = win32.Dispatch('outlook.application')

# criação de um e-mail

email = Outlook.CreateItem(0)

# COnfiguração do e-mail de destino
# Destino
email.To = "emaildedestino@outlook.com"
# Assunto
email.Subject = " Email automatico"
# Corpo do Email
email.HTMLBody = '''
<p>Olá, esse é o seu codigo Python de teste sobre faturamento de empresas. segue em anexo o relatorio em xml.</p> br 
<p> Att.: Seu gerente </p>
'''

# Seu e-mail 
email.Send = ""
print('E-mail enviado')
