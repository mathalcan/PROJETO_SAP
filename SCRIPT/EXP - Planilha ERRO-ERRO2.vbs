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
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[1,0]").select
session.findById("wnd[1]/usr/sub/2/sub/2/1/rad[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxt[0]").setFocus
session.findById("wnd[1]/usr/ctxt[0]").caretPosition = 0
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[1]/usr/ctxt[0]").text = "C:\Users\E815847\Desktop\"
session.findById("wnd[1]/usr/ctxt[1]").text = "PLANILHA ERRO.XLS"
session.findById("wnd[1]/usr/ctxt[1]").setFocus
session.findById("wnd[1]/usr/ctxt[1]").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
