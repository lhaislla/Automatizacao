import pyautogui
import time

def manipular_sap_apos_login():
    # Manipular o SAP após o login
    session = None

    # Encontra a janela ativa do SAP
    while session is None:
        session = pyautogui.getWindowsWithTitle('SAP Easy Access')[0]  # Altere o título conforme necessário
        if not session.isActive:
            session = None

    # Maximizar a janela do SAP
    session.maximize()

    # Expande e seleciona os nós no SAP GUI
    pyautogui.press('f5')  # Atualizar
    pyautogui.write("F00002")
    pyautogui.press('enter')
    pyautogui.write("F00007")
    pyautogui.press('enter')

    # Preenche os campos no SAP GUI
    pyautogui.write("a133")
    pyautogui.press('tab')
    pyautogui.write("10")
    pyautogui.press('tab')
    pyautogui.write("2023")
    pyautogui.press('tab')
    pyautogui.write("/LDINIZ")

    # Pressiona botões no SAP GUI
    pyautogui.hotkey('ctrl', 'shift', 'f5')  # Atualizar novamente
    pyautogui.hotkey('ctrl', 'shift', 'f2')  # Exportar para Excel
    pyautogui.press('enter')
    pyautogui.write("custo.xlsx")
    pyautogui.press('enter')

    # Tempo do SAP completar a exportação 
    time.sleep(10)  # Esperar 10 segundos, por exemplo
    

manipular_sap_apos_login()