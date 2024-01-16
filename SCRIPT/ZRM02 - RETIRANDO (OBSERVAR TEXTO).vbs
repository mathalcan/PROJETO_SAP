For i = 1 To 30
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell 0,""
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
session.findById("wnd[0]/tbar[1]/btn[23]").press
session.findById("wnd[1]/usr/btn[0]").press
session.findById("wnd[1]/usr/cntlO_JUSTIF1/shell").text = "POR MOTIVO DE ERRO DE COMPLEMENTAÇÃO DA FASE DE PAGAMENTO DE TODAS AS NOTAS NESTE ITEM." + vbCr + ""
session.findById("wnd[1]/usr/cntlO_JUSTIF1/shell").setSelectionIndexes 87,87
session.findById("wnd[1]/usr/txt").text = "0"
session.findById("wnd[1]/usr/txt").setFocus
session.findById("wnd[1]/usr/txt").caretPosition = 1
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[2]/tbar[0]/btn[0]").press
Next