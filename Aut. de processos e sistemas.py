from numpy import quantile
import pandas
from time import sleep
import pyautogui
import pyperclip
pyautogui.PAUSE = 1
print("kdfksdfdsafsd")

pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

sleep(5)

pyautogui.click(x=518, y=379, clicks=2)

sleep(2)
pyautogui.click(x=518, y=379)
pyautogui.click(x=1614, y=236)
pyautogui.click(x=1264, y=875)

tabela = pandas.read_excel("Vendas - Dez.xlsx")

print(tabela)

quantidade = tabela["Quantidade"].sum()

faturamento =  tabela["Valor Final"].sum()



pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
sleep(2)

pyautogui.click(x=70, y=215)
pyperclip.copy("marcjuniu02@gmail.com")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

pyautogui.click(x=1311, y=1024)
pyperclip.copy("Assunto aqui")
pyautogui.hotkey("ctrl", "v")

pyautogui.click(x=1352, y=577)
txt = f'''A quantidade foi: R${quantidade}
O faturamento foi: R${faturamento}
'''
pyperclip.copy(txt)
pyautogui.hotkey("ctrl", "v")

pyautogui.click(x=1311, y=1024)
