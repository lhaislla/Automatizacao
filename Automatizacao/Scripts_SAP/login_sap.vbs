Sub ExecutarScriptSAP()
    Dim SapGuiAuto As Object
    Dim SAPApplication As Object
    Dim SAPConnection As Object
    Dim SAPSession As Object

    ' Inicializa a automação do SAP GUI
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApplication = SapGuiAuto.GetScriptingEngine
    Set SAPConnection = SAPApplication.Children(0)
    Set SAPSession = SAPConnection.Children(0)

End Sub