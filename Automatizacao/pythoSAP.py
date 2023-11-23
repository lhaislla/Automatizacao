import win32com.client
import subprocess
import sys
import time
from  tkinter import *
from tkinter import messagebox

#Script de login ao SAP
class SapGui(object): 
    #Aponta o caminho e a descrição para executar o SAP
    def __init__(self):
        self.path ='C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe'
        subprocess.Popen(self.path)
        
        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = self.SapGuiAuto.GetScriptingEngine
        
        self.connection = application.OpenConnection("09. ALPA R3 Produção", True)
        time.sleep(3)
        self.session= self.connection.Children(0)
        self.session.findById("wnd[0]").maximize
        
    def sapLogin(self):
        #Indica as infromações de login
        try:
            self.session.findById("wnd[0]/urs/txtRYST-MANDT").text = "702"
            self.session.findById("wnd[0]/urs/txtRYST-BNAME").text = "L133168737"
            self.session.findById("wnd[0]/urs/txtRYST-BCODET").text = "Karina1976*"
            self.session.findById("wnd[0]/urs/txtRYST-LANGU").text = "PT"
            self.session.findById("wnd[0]").sendVKey(0)
            
        except:
            print(sys.exec_info()[0])
        #menssagebox.showinfo("showinfo", "Login realizado") 

if __name__ == '__main__':
    window = Tk()
    window.geometry("200x50")
    botao = Button(window, text = "Lgin SAP", command=lambda : SapGui().sapLogin)
    botao.pack()
    mainloop()
    
     
# Cria uma instância da classe SapGui
#sap_gui_instance = SapGui()
# login no SAP GUI
#sap_gui_instance.sapLogin()
        
        