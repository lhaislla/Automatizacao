Sub SAP_Login()

Dim SapGui
Dim Applic
Dim connection
Dim Session
Dim WSMShell

Shell "C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe", vbNormalFocus

Set WSHShell = CreateObject("WScript.Shell")

Do Until WSHShell.AppActive("SAP Logon")
	Application.Wait Now + TimeValue("0:01:00")
Loop 
Set WSHShell = Nothing

Set SapGui = GetObject("SAPGUI")

Set Applic = SapGui.GetScriptingEngine

Set connection = Applic.OpenConnection("09. ALPA R3 Produção", True)

Set session = connection.Childre(0)
session.findById("wnd[0]").maximize

'DAdos para login no sistema'

session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "800" 'client do sistema'
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "Usuário do sistema" usuario'
session.findById("wnd[0]/usr/txtRSYST-BCODE").Text = "Senha"
session.findById("wnd[0]/usr/txtRSYST-LANGU").Text ="PT" 'idioma do sistema'

session.findById("wnd[0]").sendVKey 0

'sair do código'

Set session = Nothing
Application.wait Now TimeValue("0:00:05")
connection.CloseSession ("ses[0]")
Set session = Nothing

End Sub