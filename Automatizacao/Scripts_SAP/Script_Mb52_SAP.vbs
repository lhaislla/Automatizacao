Sub SAP_acesso()
    Dim SapGuiAuto As Object
    Dim application As Object
    Dim connection As Object
    Dim session As Object

    ' Verifica se já está conectado
    If Not IsObject(Application) Then
        ' Cria um objeto SAPGUI
        Set SapGuiAuto = GetObject("SAPGUI")

        ' Criação da conexão
        Set application = SapGuiAuto.GetScriptingEngine
        Set connection = application.Children(0)

        ' Armazena as credenciais de acesso
        connection.User = "L133168737"
        connection.Password = "Karina1976*"

        ' Conecta-se ao sistema SAP
        connection.Connect ("09. ALPA R3 Produção")
    Else
        ' Já está conectado, não é necessário fazer login novamente
        Set application = Application
        Set connection = application.Children(0)
    End If

    ' Obtenha a primeira sessão
    Set session = connection.Children(0)

    ' Executar a transação MB52
    session.findById("wnd[0]").Maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mb52"
    session.findById("wnd[0]").SendVKey 0
    session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = "A133"
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = "M33"
    session.findById("wnd[0]/usr/ctxtP_VARI").Text = "/ANDRELG"
    session.findById("wnd[0]/usr/ctxtP_VARI").SetFocus
    session.findById("wnd[0]/usr/ctxtP_VARI").CaretPosition = 8
    session.findById("wnd[0]/tbar[1]/btn[8]").Press
    session.findById("wnd[0]/tbar[1]/btn[45]").Press
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").Press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\lcrodrigues\Documents"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "mb52.xls"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[0]").Press

    ' Fechar a sessão (opcional)
    ' session.findById("wnd[0]").Close

    ' Encerrar a conexão (opcional)
    ' connection.CloseConnection

    ' Liberar objetos
    Set session = Nothing
    Set connection = Nothing
    Set application = Nothing
    Set SapGuiAuto = Nothing
End Sub
