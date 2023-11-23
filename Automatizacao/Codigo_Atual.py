Sub Fechamento_custo()
    ' Desativa os avisos do Excel para evitar pop-ups de confirmação
    Application.DisplayAlerts = False
    
    ' Caminhos dos arquivos
    Dim caminhoCusto As String
    Dim caminhoNovoArquivo As String
    
    caminhoCusto = "C:\Users\lcrodrigues\Documents\custo\custo.xls"
    caminhoNovoArquivo = "C:\Users\lcrodrigues\Documents\custo\novo.xls"
    
    ' Verifica se o arquivo custo.xls já existe e exclui se existir
    If Dir(caminhoCusto) <> "" Then
        Kill caminhoCusto
    End If
    
    ' Verifica se o arquivo novo.xls já existe e exclui se existir
    If Dir(caminhoNovoArquivo) <> "" Then
        Kill caminhoNovoArquivo
    End If
    
    ' Continua com o restante do código para baixar os novos arquivos do SAP
    Call abrir_sap
    Application.Wait Now + TimeValue("00:00:05")
    Call executar_sap
    
    ' Restaura os avisos do Excel
    Application.DisplayAlerts = True
    
    If Err.Number = 0 Then
        MsgBox "PROCESSAMENTO FINALIZADO"
        ' Verifica se o arquivo novo.xls foi criado na pasta .\custo
        If Dir(caminhoNovoArquivo) <> "" Then
            ' Se o arquivo novo.xls existe, chame o script de tratamento
            Call ScriptTratamento
            Application.Wait Now + TimeValue("00:00:10") ' Aguarda o processamento do script de tratamento
            ' Verifica se o arquivo novo.xls foi criado após o tratamento
            If Dir(caminhoNovoArquivo) <> "" Then
                ' Se o arquivo novo.xls existe, chama o script de salvarrelatorio
                Call ScriptSalvarRelatorio
                MsgBox "Relatório salvo com sucesso!"
            Else
                MsgBox "Erro ao processar o script de tratamento."
            End If
        Else
            MsgBox "Erro: novo.xls não foi criado após a execução do SAP."
        End If
    Else
        MsgBox "Erro: " & Err.Description
    End If
    
    On Error GoTo 0
End Sub

Sub abrir_sap()

    On Error Resume Next
    
    Dim objshell As Object
    Dim objapp As Object

    Set objshell = CreateObject("WScript.Shell")
    Set objapp = objshell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
    Application.Wait Now + TimeValue("00:00:07")
    AppActivate "SAP logon Pad 740"
    Application.Wait Now + TimeValue("00:00:05")

    Application.SendKeys "PRD", True
    Application.Wait Now + TimeValue("00:00:03")
    Application.SendKeys "~", True
    Application.Wait Now + TimeValue("00:00:07")
    
    If Err.Number <> 0 Then
        MsgBox "Erro ao abrir o SAP: " & Err.Description
    End If
    
    On Error GoTo 0

End Sub

Sub executar_sap()

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Menu")
    Dim mes As Variant
    Dim ano As Variant
    
    mes = ws.Range("H7").Value
    ano = ws.Range("H9").Value
    
    
    Dim SapGuiAuto As Object
    Dim Application As Object
    Dim Connection As Object
    Dim session As Object
    
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Application = SapGuiAuto.GetScriptingEngine
    Set Connection = Application.Children(0)
    Set session = Connection.Children(0)

    session.findById("wnd[0]").Maximize
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").DoubleClickNode "F00007"
    session.findById("wnd[0]/usr/ctxtP_WERKS").Text = "a133"
    session.findById("wnd[0]/usr/txtP_POPER").Text = mes
    session.findById("wnd[0]/usr/txtP_BDATJ").Text = ano
    session.findById("wnd[0]/usr/ctxtP_VARIAN").Text = "/LDINIZ"
    session.findById("wnd[0]/usr/ctxtP_VARIAN").SetFocus
    session.findById("wnd[0]/usr/ctxtP_VARIAN").caretPosition = 7
    session.findById("wnd[0]/tbar[1]/btn[8]").Press
    session.findById("wnd[0]/tbar[0]/btn[0]").Press
    session.findById("wnd[0]/tbar[1]/btn[45]").Press
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").Press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\lcrodrigues\Documents\custo\"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "custo.xls"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
    session.findById("wnd[1]/tbar[0]/btn[0]").Press
    
    If Err.Number <> 0 Then
        MsgBox "Erro ao executar o SAP: " & Err.Description
    End If
    
    On Error GoTo 0

End Sub

'//Script de tratamento de dados
Public Sub ScriptTratamento()
    Dim caminhoPython As String
    Dim caminhoScript As String
    ' Especifica o caminho do interpretador Python e o caminho do  arquivo Python
    caminhoPython = "C:\Program Files\Python311\python.exe" ' Caminho do interpretador Python
    caminhoScript = "C:\Users\lcrodrigues\Documents\custo\Script_custo.py" ' Caminho do arquivo Python
     ' Usa a função Shell para chamar o arquivo Python
    VBA.Shell caminhoPython & " " & caminhoScript
End Sub



'// Script pra salvar na planilha de fechamento
Public Sub ScriptSalvarRelatorio()
    Dim caminhoPython As String
    Dim caminhoScript As String
   ' Especifica o caminho do interpretador Python e o caminho do  arquivo Python
    caminhoPython = "C:\Program Files\Python311\python.exe" ' Caminho do interpretador Python
    caminhoScript = "C:\Users\lcrodrigues\Documents\Automatizacao\Script_fechamento.py" ' Caminho do arquivo Python
    ' Usa a função Shell para chamar o arquivo Python
    VBA.Shell caminhoPython & " " & caminhoScript
End Sub
