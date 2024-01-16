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
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00033"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00033"
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell 5,"TEXT"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "5"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/btn[2]").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssub/1/2/tblSAPLALDBSINGLE/ctxt[1,0]").text = "CB"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssub/1/2/tblSAPLALDBSINGLE/ctxt[1,0]").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxt[42]").text = "/soli-os"
session.findById("wnd[0]/usr/ctxt[42]").setFocus
session.findById("wnd[0]/usr/ctxt[42]").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 17,"ZZTEL_PARC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 2
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 10
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 21
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 34
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 49
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 67
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 82
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 91
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 99
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 109
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 118
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 126
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 135
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 144
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 153
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 161
