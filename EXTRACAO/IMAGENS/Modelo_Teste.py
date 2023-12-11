import os
import subprocess
import psutil
import time
import pyautogui
import os
import pygetwindow as window
import pyscreeze


user = "B357509"
password ="140595L@l"
login = "SAP Logon 770 - \\Remote"
transacao = "SAP - \\Remote" 
nome_processo = "wfica32.exe"


def processo_rodando(nome_processo):
    for process in psutil.process_iter(attrs=['pid', 'name']):
        if process.info['name'] == nome_processo:
            return True
        return False

SAP_BAT = "C:\\temp\\SAP.bat"
bat = "C:\\temp\\executar_python.bat" 
if not processo_rodando(nome_processo):
    if os.path.exists(SAP_BAT):
        os.remove(SAP_BAT)
    with open(SAP_BAT, "w") as bat_file:
        bat_file.write('"C:\\Program Files (x86)\\Citrix\\ICA Client\\SelfServicePlugin\\SelfService.exe" -launch -reg "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\neoenergia-ae963f53@@XD02.ASF FPM DEV SAP_s-2" -startmenuShortcut')
        print("Novo arquivo .bat criado com sucesso!")
    subprocess.call(SAP_BAT, shell=True)

    if os.path.exists(SAP_BAT):
        os.remove(SAP_BAT)
else:
    os.system(bat)

#os.system(bat)

time.sleep(30)
pyautogui.hotkey('super','up')

LOGIN_SAP = window.getWindowsWithTitle('SAP Logon 770 - \\Remote')
LOGIN_TRANSACAO = window.getWindowsWithTitle(transacao)
LOGIN_SAP



screenWidth, screenHeight = pyautogui.size()
screenWidth, screenHeight = (1920, 1080)

print("\n","Inicialização do SAP concluida",)

print("\n","Executando entrada no SAP NORDESTE(06.01)")
pyscreeze.locateAllOnScreen('AVANCAR.png', region=(0,0, 200, 400), confidence=0.7)
pyautogui.click
time.sleep(2)

pyautogui.click(1860, 40)
time.sleep(2)
pyautogui.write("06.01")
time.sleep(2)

pyautogui.doubleClick(60, 80)
time.sleep(2)

window = window.getWindowsWithTitle(transacao)
#window.activate()

pyautogui.hotkey('super','up')
pyautogui.click(100, 50)
pyautogui.write('SM37', interval=0)
pyautogui.click(200, 200)
pyautogui.write(user, interval=0)
pyautogui.hotkey('tab')
pyautogui.write(password, interval=0)
time.sleep(1)
pyautogui.click(25, 50)
time.sleep(1)
pyautogui.doubleClick(200, 175)
time.sleep(1)
pyautogui.doubleClick(200, 200)
pyautogui.moveTo(200, 205, duration=1)
pyautogui.doubleClick(200, 205)
pyautogui.click(25, 50)




