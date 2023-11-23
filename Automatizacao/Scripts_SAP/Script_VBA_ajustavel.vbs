Sub SAP_acesso()
    Dim SapGuiAuto As Object
    Dim application As Object
    Dim connection As Object
    Dim session As Object

    ' Verifica se já está conectado
    If IsObject(Application) Then
        ' Já está conectado, não é necessário fazer login novamente
        Exit Sub
    End If

    ' Cria um objeto SAPGUI
    Set SapGuiAuto = GetObject("SAPGUI")

    ' Conecta-se ao sistema SAP
    Set application = SapGuiAuto.GetScriptingEngine

    ' Cria um objeto Connection
    Set connection = application.OpenConnection("09. ALPA R3 Produção", True)

    ' Armazena as credenciais de acesso
    connection.User = "L133168737"
    connection.Password = "Karina1976*"

    ' Conectar-se à primeira sessão
    Set session = connection.Children(0)
   
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "mb52"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "A133"
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "M33"
    session.findById("wnd[0]/usr/ctxtP_VARI").text = "/ANDRELG"
    session.findById("wnd[0]/usr/ctxtP_VARI").setFocus
    session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 8
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[45]").press
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\lcrodrigues\Documents"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "mb52.xls"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[0]").press
End Sub