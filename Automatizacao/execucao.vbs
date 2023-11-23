'Script VBA para execução do relatório de fechamento de custo
Public Sub ExecutarScriptsCondicionalmente()
    ' Verifica se o arquivo custo.xlsx existe na pasta ./custo
    If Dir("C:\Users\lcrodrigues\Documents\Automatizacao\custo\custo.xlsx") <> "" Then
        ' Executa ScriptLogin()
        Call ScriptLogin
        
        ' Espera até que o ScriptLogin() seja concluído
        Application.Wait Now + TimeValue("00:00:10") ' Aguarda 10 segundos
        
        ' Verifica se o arquivo novo.xlsx existe na pasta ./custo
        If Dir("C:\Users\lcrodrigues\Documents\Automatizacao\custo\novo.xlsx") <> "" Then
            ' Executa ScriptSAP()
            Call ScriptSAP
            
            ' Espera até que o ScriptSAP() seja concluído
            Application.Wait Now + TimeValue("00:00:10") ' Aguarda 10 segundos
            
            ' Executa ScriptTratamento()
            Call ScriptTratamento
            
            ' Espera até que o ScriptTratamento() seja concluído
            Application.Wait Now + TimeValue("00:00:10") ' Aguarda 10 segundos
            
            ' Executa ScriptSalvarRelatorio()
            Call ScriptSalvarRelatorio
        Else
            ' Se o arquivo novo.xlsx não existir, exibe uma mensagem ou executa outras ações conforme necessário
            MsgBox "O arquivo novo.xlsx não foi encontrado na pasta ./custo."
        End If
    Else
        ' Se o arquivo custo.xlsx não existir, exibe uma mensagem ou executa outras ações conforme necessário
        MsgBox "O arquivo custo.xlsx não foi encontrado na pasta ./custo."
    End If
End Sub

'Funções VBA chamando os scripts Python
Public Sub ScripLogin()
    Dim caminhoPython As String
    Dim caminhoScript As String
    ' Especifica o caminho do interpretador Python e o caminho do  arquivo Python
    caminhoPython = "C:\Program Files\Python311\python.exe" ' Caminho do interpretador Python
    caminhoScript = "C:\Users\lcrodrigues\Documents\Automatizacao\integracao_sap.py" ' Caminho do arquivo Python
    ' Usa a função Shell para chamar o arquivo Python
    VBA.Shell caminhoPython & " " & caminhoScript
End Sub

Public Sub ScriptSAP()
    Dim caminhoPython As String
    Dim caminhoScript As String
    ' Especifica o caminho do interpretador Python e o caminho do  arquivo Python
    caminhoPython = "C:\Program Files\Python311\python.exe" ' Caminho do interpretador Python
    caminhoScript = "C:\Users\lcrodrigues\Documents\Automatizacao\integracao_sap.py" ' Caminho do arquivo Python
    ' Usa a função Shell para chamar o arquivo Python
    VBA.Shell caminhoPython & " " & caminhoScript
End Sub

Public Sub ScriptTratamento()
    Dim caminhoPython As String
    Dim caminhoScript As String 
    ' Especifica o caminho do interpretador Python e o caminho do  arquivo Python
    caminhoPython = "C:\Program Files\Python311\python.exe" ' Caminho do interpretador Python
    caminhoScript = "C:\Users\lcrodrigues\Documents\Script_custo.py" ' Caminho do arquivo Python
     ' Usa a função Shell para chamar o arquivo Python
    VBA.Shell caminhoPython & " " & caminhoScript
End Sub
Public Sub ScriptSalvarRelatorio()
    Dim caminhoPython As String
    Dim caminhoScript As String
   ' Especifica o caminho do interpretador Python e o caminho do  arquivo Python
    caminhoPython = "C:\Program Files\Python311\python.exe" ' Caminho do interpretador Python
    caminhoScript = "C:\Users\lcrodrigues\Documents\Automatizacao\Script_fechamento.py" ' Caminho do arquivo Python
    ' Usa a função Shell para chamar o arquivo Python
    VBA.Shell caminhoPython & " " & caminhoScript
End Sub