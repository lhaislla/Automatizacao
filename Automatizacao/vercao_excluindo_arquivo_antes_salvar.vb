
Sub Fechamento_custo()

    ' Caminho do arquivo
    Dim caminhoArquivo As String
    caminhoArquivo = "C:\Users\lcrodrigues\Documents\custo\custo.xls"

    ' Verificar se o arquivo existe
    If Dir(caminhoArquivo) <> "" Then
        ' Se o arquivo existe, exclua-o
        Kill caminhoArquivo
    End If

    ' Continuar com o restante do script
    On Error Resume Next
    Call abrir_sap
    Application.Wait Now + TimeValue("00:00:05")
    Call executar_sap
        
    If Err.Number = 0 Then
        MsgBox "PROCESSAMENTO FINALIZADO"
    Else
        MsgBox "Erro: " & Err.Description
    End If
    
    On Error GoTo 0

End Sub