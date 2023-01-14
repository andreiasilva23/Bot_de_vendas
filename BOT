import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

# pyautogui.click -> clicar
# pyautogui.press -> apertar 1 tecla
# pyautogui.hotkey -> conjunto de teclas
# pyautogui.write -> escreve um texto

# Passo 1: Entrar no sistema da empresa (no nosso caso é o link do drive)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)

# Passo 2: Navegar no sistema e encontrar a base de vendas (entrar na pasta exportar)
pyautogui.click(x=1023, y=710, clicks=2)
time.sleep(2)

# Passo 3: Fazer o download da base de vendas
pyautogui.click(x=1100, y=904) # clicar no arquivo
pyautogui.click(x=3288, y=411) # clicar nos 3 pontinhos
pyautogui.click(x=2716, y=1523) # clicar no fazer download
time.sleep(5) # esperar o download acabar

# Passo 4: Importar a base de vendas pro Python
import pandas as pd

tabela = pd.read_excel(r"C:\Users\joaol\Downloads\Vendas - Dez.xlsx")
display(tabela)

7089 rows × 7 columns


# Passo 5: Calcular os indicadores da empresa
faturamento = tabela["Valor Final"].sum()
print(faturamento)
quantidade = tabela["Quantidade"].sum()
print(quantidade)


# Passo 6: Enviar um e-mail para a diretoria com os indicadores de venda

# abrir aba
pyautogui.hotkey("ctrl", "t")

# entrar no link do email - https://mail.google.com/mail/u/0/#inbox
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# clicar no botão escrever


pyautogui.click(x=240, y=415)

# preencher as informações do e-mail
pyautogui.write("pythonimpressionador@gmail.com")
pyautogui.press("tab") # selecionar o email

pyautogui.press("tab") # pular para o campo de assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")

pyautogui.press("tab") # pular para o campo de corpo do email

texto = f"""
Prezados,
Segue relatório de vendas.
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}
Qualquer dúvida estou à disposição.
Att.,
Lira do Python
"""

# formatação dos números (moeda, dinheiro)

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# enviar o e-mail
pyautogui.hotkey("ctrl", "enter")
Footer
© 2023 GitHub, Inc.
Footer navigation
Terms
Privacy
Security
