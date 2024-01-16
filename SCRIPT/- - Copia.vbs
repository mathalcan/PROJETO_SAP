
For i = 1 To 13
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
session.findById("wnd[0]/tbar[1]/btn[7]").press
session.findById("wnd[0]/usr/ctxt[2]").text = ""
session.findById("wnd[0]/usr/ctxt[3]").text = ""
session.findById("wnd[0]/usr/ctxt[10]").text = "93024629"
session.findById("wnd[0]/usr/ctxt[10]").setFocus
session.findById("wnd[0]/usr/ctxt[10]").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblZCRSEC_CLICK_CHANGE_DATATB_CRT_MATERIAL").getAbsoluteRow(0).selected = true
session.findById("wnd[0]/usr/tblZCRSEC_CLICK_CHANGE_DATATB_CRT_MATERIAL/txt[0,0]").setFocus
session.findById("wnd[0]/usr/tblZCRSEC_CLICK_CHANGE_DATATB_CRT_MATERIAL/txt[0,0]").caretPosition = 0
session.findById("wnd[0]/usr/btn[2]").press
session.findById("wnd[0]/usr/tblZCRSEC_CLICK_CHANGE_DATATB_CRT_MATERIAL/txt[0,0]").text = "0"
session.findById("wnd[0]/usr/tblZCRSEC_CLICK_CHANGE_DATATB_CRT_MATERIAL/txt[1,0]").text = "1"
session.findById("wnd[0]/usr/tblZCRSEC_CLICK_CHANGE_DATATB_CRT_MATERIAL/cmb[2,0]").key = "R"
session.findById("wnd[0]/usr/tblZCRSEC_CLICK_CHANGE_DATATB_CRT_MATERIAL/cmb[2,0]").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
Next