import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE =1

print(pyautogui.position())

pyautogui.hotkey('win','s')
pyautogui.write('edge')
pyautogui.press('enter')
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

time.sleep(8)

pyautogui.click(x=447, y=327, clicks = 2)
time.sleep(5)

pyautogui.click(x=447, y=327)
pyautogui.click(x=1119, y=205)
time.sleep(2)

pyautogui.click(x=949, y=718)
time.sleep(8)

pyautogui.click(x=1166, y=151)
pyautogui.click(x=1133, y=191)

tabela = pd.read_excel(r'C:\Users\Antonio\Downloads\Vendas - Dez.xlsx')
quantidade = tabela['Quantidade'].sum()
faturamento = tabela['Valor Final'].sum()



pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl', 'shift', 'l')
time.sleep(8)

pyautogui.click(x=103, y=189)
time.sleep(3)
pyautogui.write(r'dialmeida542@gmail.com')
pyautogui.press('tab', presses = 2)
pyperclip.copy("Relatório")
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
pyautogui.write('''
Relatório de vendas:
--------------------
Total de itens vendidos: {}
Valor arrecadado: R$ {:,.2f}
'''.format(quantidade, faturamento))

pyautogui.hotkey('ctrl', 'enter')
