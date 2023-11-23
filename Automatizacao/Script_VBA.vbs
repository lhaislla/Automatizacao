'Exemplos de chamda de arquivo pelo CMD
'Public Sub Teste()
   'VBA.Shell "cmd.exe /c  ""C:\Users\lcrodrigues\Documents\script_teste.vbs "" "
'End Sub


'Public Sub Teste()
'   VBA.Shell "cmd.exe /c  ""C:\Users\lcrodrigues\Documents\script_teste.vbs "" "
'End Sub

'Public Sub ScripLogin()
    'Dim caminhoPython As String
    'Dim caminhoScript As String
    ' Especifica o caminho do interpretador Python e o caminho do  arquivo Python
    'caminhoPython = "C:\Program Files\Python311\python.exe" ' Caminho do interpretador Python
    'caminhoScript = "C:\Users\lcrodrigues\Documents\Automatizacao\integracao_sap.py" ' Caminho do arquivo Python
    ' Usa a função Shell para chamar o arquivo Python
    'VBA.Shell caminhoPython & " " & caminhoScript
'End Sub

'Funções VBA chamando os scripts Python

Sub Macro_PrimariaComando()

    Dim objshell As Object
    Dim objapp As Object
    
    Call abrir_sap
    Application.Wait Now + TimeValue("00:00:05")
    Call executar_sap
        
    MsgBox "PROCESSAMENTO FINALIZADO"

End Sub

Sub abrir_sap()

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

End Sub



'//Script de macro SAP a serem utilizados na planilha 
Public Sub executar_sap()
    Dim caminhoPython As String
    Dim caminhoScript As String
    ' Especifica o caminho do interpretador Python e o caminho do  arquivo Python
    caminhoPython = "C:\Program Files\Python311\python.exe" ' Caminho do interpretador Python
    caminhoScript = "C:\Users\lcrodrigues\Documents\Automatizacao\integracao_sap.py" ' Caminho do arquivo Python
    ' Usa a função Shell para chamar o arquivo Python
    VBA.Shell caminhoPython & " " & caminhoScript
End Sub


'//Script de tratamento de dados 
Public Sub ScriptTratamento()
    Dim caminhoPython As String
    Dim caminhoScript As String 
    ' Especifica o caminho do interpretador Python e o caminho do  arquivo Python
    caminhoPython = "C:\Program Files\Python311\python.exe" ' Caminho do interpretador Python
    caminhoScript = "C:\Users\lcrodrigues\Documents\Script_custo.py" ' Caminho do arquivo Python
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