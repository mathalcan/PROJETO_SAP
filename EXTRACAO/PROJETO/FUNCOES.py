# Bibliotes Utiladas 
import pyautogui
import psutil
import subprocess
import os
import time
import schedule
import keyboard



# Variaveis de acesso
user = "B357509"
password ="140595L@l"

# Localização dos objetos a ser procurado
AVANCAR_IMG = 'C:\\Users\\B357509\\Desktop\\EXTRACAO\\IMAGENS\\AVANCAR.PNG'
ACESSO_SAP_IMG = 'C:\\Users\\B357509\\Desktop\\EXTRACAO\\IMAGENS\\SAP.png'
ACESSO_SAP2_IMG = 'C:\\Users\\B357509\\Desktop\\EXTRACAO\\IMAGENS\\SAP2.png'
SAP_REMOTE_IMG = 'C:\\Users\\B357509\\Desktop\\EXTRACAO\\IMAGENS\\SAP_Remote.png'
SAP_NORDESTE_IMG = 'C:\\Users\\B357509\\Desktop\\EXTRACAO\\IMAGENS\\NORDESTE_6.01.png'
SAP_MAXIMIZAR_IMG = 'C:\\Users\\B357509\\Desktop\\EXTRACAO\\IMAGENS\\MAXIMIZAR_SAP.png'
SAP_FILTRO_IMG = 'C:\\Users\\B357509\\Desktop\\EXTRACAO\\IMAGENS\\FILTRO.png'
SAP_BAT = "C:\\temp\\SAP.bat"

bat = "C:\\temp\\executar_python.bat"
nome_processo = "SAP"

#def AGENDADOR():
      
      
def criar_bat():
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
            return True
        return False

def ACESSO_SAP():
        time.sleep(5)
        try:
            screen_width, screen_height = pyautogui.size()
            loc_barra_de_tarefas = (0, screen_height - 40, screen_width, 40)
    
            target_location = pyautogui.locateOnScreen(ACESSO_SAP_IMG, region=loc_barra_de_tarefas)
    
            if target_location:
                target_center_x = target_location[0] + target_location[2] // 2
                target_center_y = target_location[1] + target_location[3] // 2
                pyautogui.click(target_center_x, target_center_y)
                time.sleep(5)
                pyautogui.hotkey('super','up')
                time.sleep(5)
                ACESSO_SAP_NORDESTE()



            else:
                print(f"Icone SAP {ACESSO_SAP_IMG} não encontrado na barra de tarefas.")
                time.sleep(1)
                print(f"Iniciou-se uma nova busca. {ACESSO_SAP2_IMG} ")
                time.sleep(1)
                #Caso não seja encontrar o icone de acesso ao SAP na barra de tarefas se inicia um nova busca com um novo padrão de imagem.
                ACESSO_SAP2()
                time.sleep(1)
                pyautogui.hotkey('super','8')
                time.sleep(1)
                pyautogui.hotkey('super','up')
                time.sleep(5)
                pyautogui.click(1860, 40)
                time.sleep(2)
                pyautogui.write("06.01")
                time.sleep(2)
                pyautogui.press('enter')
                keyboard.press(['enter','enter'])






                
                
            return False
        except Exception as e:
            print(f"Erro: {e}")
            return False
        
def ACESSO_SAP2():
        try:
            screen_width, screen_height = pyautogui.size()
            loc_barra_de_tarefas = (0,0, screen_height - 40, screen_width, 40)
    
            target_location = pyautogui.locateOnScreen(ACESSO_SAP2_IMG, region=loc_barra_de_tarefas)
    
            if target_location:
                target_center_x = target_location[0] + target_location[2] // 2
                target_center_y = target_location[1] + target_location[3] // 2
                pyautogui.click(target_center_x, target_center_y)
            else:
                print(f"Icone SAP {ACESSO_SAP2_IMG} não encontrado na barra de tarefas.")
            return False
        except Exception as e:
            print(f"Erro: {e}")
            return False
        



def ACESSO_SAP_NORDESTE():
        try:      
                pyautogui.hotkey('super','up')
        except TypeError:
                time.sleep(1)
        return False


def AVANCAR():
        try:
                # Localizar a posição da imagem na tela
                target_location = pyautogui.locateOnScreen(AVANCAR_IMG)
        
                # Calcular o centro da imagem
                target_center_x = target_location[0] + target_location[2] // 2
                target_center_y = target_location[1] + target_location[3] // 2
        
                # Mover o cursor para o centro da imagem e clicar
                pyautogui.click(target_center_x, target_center_y)
                return True
        except TypeError:
                print("Imagem não encontrada na tela.")
                ACESSO_SAP2
            
        return False
    

    
def SAP_REMOTE():
        try:

                target_location = pyautogui.locateOnScreen(SAP_REMOTE_IMG)    
                target_center_x = target_location[0] + target_location[2] // 2
                target_center_y = target_location[1] + target_location[3] // 2
        
                pyautogui.click(target_center_x, target_center_y)
                return True
        except TypeError:
                print("Imagem não encontrada na tela.")
        return False

def SAP_NORDESTE():
        try:

                target_location = pyautogui.locateOnScreen(SAP_NORDESTE_IMG)    
                target_center_x = target_location[0] + target_location[2] // 2
                target_center_y = target_location[1] + target_location[3] // 2
        
                pyautogui.doubleClick(target_center_x, target_center_y)
                return True
        except TypeError:
                print("Imagem não encontrada na tela.")
        return False

def MAXIMIZAR_SAP():
        try:

                target_location = pyautogui.locateOnScreen(SAP_MAXIMIZAR_IMG)    
                target_center_x = target_location[0] + target_location[2] // 2
                target_center_y = target_location[1] + target_location[3] // 2
        
                pyautogui.click(target_center_x, target_center_y)
                return True
        except TypeError:
                print("Imagem não encontrada na tela.")
        return False

def FILTRO():
        try:

                target_location = pyautogui.locateOnScreen(SAP_FILTRO_IMG)    
                target_center_x = target_location[0] + target_location[2] // 2
                target_center_y = target_location[1] + target_location[3] // 2
        
                pyautogui.click(target_center_x, target_center_y)
                return True
        except TypeError:
                print("Imagem não encontrada na tela.")
        return False

def processo_rodando(nome_processo):
        nome_processo = "SAP"

        for process in psutil.process_iter(attrs=['pid', 'name']):
            if process.info['name'] == nome_processo:
                return True
            return False


#start
#criar_bat()
time.sleep(30)
ACESSO_SAP()
time.sleep(5)








      





        

    






    