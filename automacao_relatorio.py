# Primeiro vamos trabalhar a parte lógica, o que precisa ser feito:
# 1 - Acessar e baixar tabela de dados no Drive da "empresa"
# 2 - Importar essa base de dados e calcular faturamento e ticket médio
# 3 - Escrever o relatorio e enviar por email.

# Primeiro vamos importar, as já instaladas, bibliotecas e modulos que vamos usar
import pyautogui as auto 
import time
import pyperclip
import pandas as pd
import numpy
import openpyxl

#Precisamos que as interações do nosso sistema esperem um tempo para os outros sistemas estarem funcionando, como abrir um navegador, uma pasta, etc. 
# Para isso vamos estabelecer a espera de 1 segundo enntre as interações.
auto.PAUSE = 2


# Agora vamos abrir o Chrome e acessar o drive da "empresa"
auto.hotkey("winleft")
auto.write("chrome")
auto.press("enter")
time.sleep(1)
auto.click(x=170, y=48)
pyperclip.copy("https://drive.google.com/drive/u/0/folders/1oxwP9llmtp2PQz58w6Fc1CKsaUYPilph")
auto.hotkey("ctrl", "v")
auto.press("enter")
time.sleep(2)
# Vamos baixar a tabela. É bom colocar um timer de 2 secs pra dar tempo do navegador abrir os menus.
auto.click(x=342, y=490)
time.sleep(2)
auto.click(x=1158, y=190)
time.sleep(2)
auto.click(x=1011, y=600)
# vamos esperar 5 segundos para dar tempo do arquivo baixar
time.sleep(5)


# passo 1 concluido! Agora vamos ao passo 2:

# nossa tabela foi baixada, agora precisamos informar nosso sitema onde ela esta apra ele poder importa-la.
tabela = pd.read_excel(r"C:\Users\dougg\Downloads\Vendas - Dez.xlsx")

# Agora calculamos o que queremos apresentar, no caso o faturamento e o ticket médio.
faturamento = tabela["Valor Final"].sum()
ticket = (faturamento/len(tabela))

# Vamos fazer nosso sistema abrir o email para redigi-lo
# vamos aumentar um pouco o tempo de espera no nosso sistema
auto.PAUSE = 3
# Estabelecer as variáveis destinatário do nosso email, o assunto do email, e o corpo do email.
destinatario = ("douggr@gmail.com")
assunto = ("Relatório vendas")
corpo = (f""" Prezados, bom dia
O faturamento total de ontem foi de R${faturamento:,.2f}.
Nosso ticket médio foi de R${ticket:,.2f}.

Obrigado, """)

# a automação em sí de onde e quando clickar
time.sleep(3)
auto.hotkey("winleft")
auto.write("chrome")
auto.press("enter")
auto.click(x=1186, y=131)
time.sleep(2)
auto.click(x=43, y=210)
time.sleep(5)
pyperclip.copy(destinatario)
auto.hotkey("ctrl", "v")
auto.press("tab")
pyperclip.copy(assunto)
auto.hotkey("ctrl", "v")
auto.press("tab")
pyperclip.copy(corpo)
auto.hotkey("ctrl", "v")
auto.press("tab")
auto.press("enter")
auto.hotkey("altleft", "f4")
time.sleep(2)
auto.hotkey("altleft", "f4")
