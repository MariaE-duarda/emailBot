import win32com.client as win32
import termcolor
from termcolor import colored 

# integração com o outlook 
outlook = win32.Dispatch('outlook.application')
# criar um e-mail 
email = outlook.CreateItem(0)
# informações do e-mail 
email.To = 'eduardafreire115@gmail.com'
email.Subject = 'E-mail automático'
email.HTMLBody = """
    <h2>Olá, Eduarda!</h2> 
    <p>Esse é um e-mail automático do Python.</p> </br> 

    <b><p>Att, Maria Eduarda.</p></b>
"""

email.Send()
print('E-mail enviado com',colored('sucesso.', 'green'))