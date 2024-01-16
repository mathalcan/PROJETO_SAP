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
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell 9,"TEXT"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").firstVisibleRow = 5
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "9"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/btn[1]").press



' Caixa de texto para notas
Dim notas
notas = InputBox("Digite as notas separadas por vírgula", "Notas")


Dim listaNotas
listaNotas = Split(notas, ",")



' Preencher notas
For i = 0 To UBound(listaNotas)
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssub/1/2/tblSAPLALDBSINGLE/ctxt[1," & i & "]").text = Trim(listaNotas(i))
Next


session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxt[129]").text = "/soli-os"

session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 1,"ZZNOME_PARC"
session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[2]").select
session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[2]").select

session.findById("wnd[1]/usr/ctxt[1]").text = "PROG.os"
session.findById("wnd[1]/usr/ctxt[1]").caretPosition = 7
session.findById("wnd[1]/tbar[0]/btn[11]").press
