
import time
import pyautogui
import os
import pygetwindow
import cv2
import FUNCOES as funcao


funcao.ACESSO_SAP()



user = "B357509"
password ="140595P@p"
bat = 'C:\\temp\\executar_python.bat'
logon = 'SAP Logon 770 - \\Remote'
transacao = 'SAP - \\Remote' 
nome_processo = 'wfica32.exe'

screenWidth, screenHeight = pyautogui.size()
screenWidth, screenHeight = (1920, 1080)

print("\n","Inicialização do SAP concluida",)

print("\n","Executando entrada no SAP NORDESTE(06.01)")
time.sleep(5)

pyautogui.click(60, 80)
time.sleep(2)

pyautogui.doubleClick(60, 80)
time.sleep(2)

pyautogui.hotkey("super","up")
pyautogui.click(100, 50)
pyautogui.write("SM37", interval=0)
time.sleep(1)
pyautogui.click(200, 200)
pyautogui.write(user, interval=0)
time.sleep(1)
pyautogui.hotkey("tab")
pyautogui.write(password, interval=0)
time.sleep(1)
pyautogui.click(25, 50)


