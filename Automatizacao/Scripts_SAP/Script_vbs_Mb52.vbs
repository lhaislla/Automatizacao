' Crie um objeto SAPGUI
Set SapGuiAuto = GetObject("SAPGUI")

' Verifica se já está conectado
If IsObject(Application) Then
    ' Já está conectado, não é necessário fazer login novamente
    WScript.Quit
End If

' Conecta-se ao sistema SAP
Set application = SapGuiAuto.GetScriptingEngine

' Crie um objeto Connection
Set connection = application.OpenConnection("09. ALPA R3 Produção", True)

' Armazene as credenciais de acesso
connection.User = "L133168737"
connection.Password = "Karina1976*"

' Conecte-se à primeira sessão
Set session = connection.Children(0)

' Realize as ações no SAP
session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "mb52"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = "A133"
session.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = "M33"
session.FindById("wnd[0]/usr/ctxtP_VARI").Text = "/ANDRELG"
session.FindById("wnd[0]/usr/ctxtP_VARI").SetFocus
session.FindById("wnd[0]/usr/ctxtP_VARI").CaretPosition = 8
session.FindById("wnd[0]/tbar[1]/btn[8]").Press
session.FindById("wnd[0]/tbar[1]/btn[45]").Press
session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\lcrodrigues\Documents"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "mb52.xls"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 8
session.FindById("wnd[1]/tbar[0]/btn[0]").Press

' Feche a conexão
connection.Close

' Saia do script
WScript.Quit
