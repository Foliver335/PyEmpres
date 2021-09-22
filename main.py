
# Criação de um Auto Reporter


import win32com.client as win32

Outlook = win32.Dispatch('outlook.application')

# criação de um e-mail

email = Outlook.CreateItem(0)

# COnfiguração do SEU email
# Destino
email.To = "Felipe-oliver2015@outlook.com"
# Assunto
email.Subject = " Email automatico"
# Corpo do Email
email.HTMLBody = '''
<p>Olá, esse é o seu codigo Python de teste sobre faturamento de empresas. segue em anexo o relatorio em xml.</p> br 
<p> Att.: Seu gerente </p>
'''

# Rementente
email.Send = ""
print('E-mail enviado')
