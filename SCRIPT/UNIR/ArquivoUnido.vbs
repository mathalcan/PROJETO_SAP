'Script gerado por SOLI OS em 26/11/2023 20:57:09
'Desligamento BAIXA DE MEDIDOR MONO COM LEITURA
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
session.findById("wnd[0]/usr/ctxt").text = "4900892179"
session.findById("wnd[0]/usr/ctxt").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").caretPosition = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 2
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 4
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 5
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 6
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 7
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 8
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 9
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 10
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 11
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 12
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 13
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 14
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 15
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").text = "Grava  o: N o - Medidor Interno"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/sub/1/btn[2]").press
session.findById("wnd[1]/usr/btn[0]").press
session.findById("wnd[1]/usr/sub/1/rad[1,0]").select
session.findById("wnd[1]/usr/sub/1/rad[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/btn[10]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,1]").text = "EX04"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,2]").text = "DCRE"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,3]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,3]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").select
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[3]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[1]").text = "25112023"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[2]").text = "110623"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[4]").text = "111303"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[6]").text = "93024467"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").key = "2"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_FLCAD/txt[2,0]").text = "5689"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/ctxt[0,0]").text = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/txt[2,0]").text = "1"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").key = "R"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "O script foi conclu do.", 10, "SOLI OS - Cosern", 64


'Script gerado por SOLI OS em 26/11/2023 20:57:10
'Desligamento BAIXA DE MEDIDOR E RAMAL MONO COM LEITURA	
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
session.findById("wnd[0]/usr/ctxt").text = "4900896939"
session.findById("wnd[0]/usr/ctxt").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").caretPosition = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 2
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 4
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 5
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 6
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 7
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 8
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 9
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 10
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 11
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 12
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 13
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 14
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 15
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").text = "Grava  o: Sim - Gravado"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/sub/1/btn[2]").press
session.findById("wnd[1]/usr/btn[0]").press
session.findById("wnd[1]/usr/sub/1/rad[1,0]").select
session.findById("wnd[1]/usr/sub/1/rad[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/btn[10]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,1]").text = "EX04"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,2]").text = "EX06"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,3]").text = "DCRE"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[3]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[1]").text = "25112023"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[2]").text = "111739"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[4]").text = "120134"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[6]").text = "93024467"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").key = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_FLCAD/txt[2,0]").text = "00000"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/ctxt[0,0]").text = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/txt[2,0]").text = "1"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").key = "R"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "O script foi conclu do.", 10, "SOLI OS - Cosern", 64


'Script gerado por SOLI OS em 26/11/2023 20:57:08
'Desligamento BAIXA DE MEDIDOR E RAMAL MONO COM LEITURA	
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
session.findById("wnd[0]/usr/ctxt").text = "4900896940"
session.findById("wnd[0]/usr/ctxt").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").caretPosition = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 2
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 4
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 5
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 6
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 7
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 8
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 9
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 10
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 11
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 12
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 13
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 14
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 15
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").text = "Grava  o: Sim - Gravado"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/sub/1/btn[2]").press
session.findById("wnd[1]/usr/btn[0]").press
session.findById("wnd[1]/usr/sub/1/rad[1,0]").select
session.findById("wnd[1]/usr/sub/1/rad[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/btn[10]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,1]").text = "EX04"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,2]").text = "EX06"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,3]").text = "DCRE"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[3]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[1]").text = "25112023"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[2]").text = "081315"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[4]").text = "083803"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[6]").text = "93024467"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").key = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_FLCAD/txt[2,0]").text = "1556"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/ctxt[0,0]").text = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/txt[2,0]").text = "1"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").key = "R"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "O script foi conclu do.", 10, "SOLI OS - Cosern", 64


'Script gerado por SOLI OS em 26/11/2023 20:57:09
'Desligamento BAIXA DE MEDIDOR E RAMAL MONO COM LEITURA	
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
session.findById("wnd[0]/usr/ctxt").text = "4900896943"
session.findById("wnd[0]/usr/ctxt").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").caretPosition = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 2
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 4
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 5
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 6
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 7
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 8
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 9
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 10
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 11
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 12
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 13
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 14
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 15
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").text = "Grava  o: N o -  rea de Risco"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/sub/1/btn[2]").press
session.findById("wnd[1]/usr/btn[0]").press
session.findById("wnd[1]/usr/sub/1/rad[1,0]").select
session.findById("wnd[1]/usr/sub/1/rad[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/btn[10]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,1]").text = "EX04"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,2]").text = "EX06"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,3]").text = "DCRE"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[3]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[1]").text = "25112023"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[2]").text = "094553"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[4]").text = "095845"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[6]").text = "93024467"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").key = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_FLCAD/txt[2,0]").text = "2046"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/ctxt[0,0]").text = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/txt[2,0]").text = "1"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").key = "R"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "O script foi conclu do.", 10, "SOLI OS - Cosern", 64


'Script gerado por SOLI OS em 26/11/2023 20:57:12
'Desligamento BAIXA DE MEDIDOR E RAMAL MONO SEM LEITURA	COM JUSTIFICATIVA
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
session.findById("wnd[0]/usr/ctxt").text = "4900896945"
session.findById("wnd[0]/usr/ctxt").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").caretPosition = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 2
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 4
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 5
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 6
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 7
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 8
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 9
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 10
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 11
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 12
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 13
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 14
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 15
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").text = "Grava  o: Sim - Gravado"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/sub/1/btn[2]").press
session.findById("wnd[1]/usr/btn[0]").press
session.findById("wnd[1]/usr/sub/1/rad[1,0]").select
session.findById("wnd[1]/usr/sub/1/rad[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/btn[10]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,1]").text = "EX04"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,2]").text = "EX06"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,3]").text = "DCRE"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[3]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[1]").text = "25112023"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[2]").text = "153026"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[4]").text = "153256"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[6]").text = "93024047"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").key = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/rad[1]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/rad[1]").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/cmb").key = "03"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/ctxt[0,0]").text = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/txt[2,0]").text = "1"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").key = "R"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "O script foi conclu do.", 10, "SOLI OS - Cosern", 64


'Script gerado por SOLI OS em 26/11/2023 20:57:16
'Desligamento BAIXA DE MEDIDOR E RAMAL MONO COM LEITURA	
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
session.findById("wnd[0]/usr/ctxt").text = "4900896946"
session.findById("wnd[0]/usr/ctxt").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").caretPosition = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 2
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 4
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 5
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 6
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 7
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 8
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 9
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 10
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 11
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 12
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 13
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 14
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 15
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").text = "Grava  o: Sim - Gravado"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/sub/1/btn[2]").press
session.findById("wnd[1]/usr/btn[0]").press
session.findById("wnd[1]/usr/sub/1/rad[1,0]").select
session.findById("wnd[1]/usr/sub/1/rad[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/btn[10]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,1]").text = "EX04"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,2]").text = "EX06"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,3]").text = "DCRE"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[3]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[1]").text = "25112023"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[2]").text = "153556"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[4]").text = "160332"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[6]").text = "93024047"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").key = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_FLCAD/txt[2,0]").text = "2805"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/ctxt[0,0]").text = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/txt[2,0]").text = "1"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").key = "R"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "O script foi conclu do.", 10, "SOLI OS - Cosern", 64


'Script gerado por SOLI OS em 26/11/2023 20:57:09
'Desligamento BAIXA DE MEDIDOR E RAMAL MONO SEM LEITURA	COM JUSTIFICATIVA
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
session.findById("wnd[0]/usr/ctxt").text = "4900896947"
session.findById("wnd[0]/usr/ctxt").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").caretPosition = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 2
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 4
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 5
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 6
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 7
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 8
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 9
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 10
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 11
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 12
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 13
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 14
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 15
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").text = "Grava  o: Sim - Gravado"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/sub/1/btn[2]").press
session.findById("wnd[1]/usr/btn[0]").press
session.findById("wnd[1]/usr/sub/1/rad[1,0]").select
session.findById("wnd[1]/usr/sub/1/rad[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/btn[10]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,1]").text = "EX04"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,2]").text = "EX06"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,3]").text = "DCRE"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[3]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[1]").text = "25112023"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[2]").text = "100221"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[4]").text = "103449"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[6]").text = "93024467"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").key = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/rad[1]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/rad[1]").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/cmb").key = "03"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/ctxt[0,0]").text = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/txt[2,0]").text = "1"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").key = "R"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "O script foi conclu do.", 10, "SOLI OS - Cosern", 64


'Script gerado por SOLI OS em 26/11/2023 20:57:08
'Desligamento BAIXA DE MEDIDOR E RAMAL MONO SEM LEITURA	COM JUSTIFICATIVA
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
session.findById("wnd[0]/usr/ctxt").text = "4900896952"
session.findById("wnd[0]/usr/ctxt").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").caretPosition = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 2
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 4
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 5
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 6
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 7
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 8
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 9
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 10
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 11
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 12
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 13
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 14
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 15
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").text = "Grava  o: Sim - Gravado"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/sub/1/btn[2]").press
session.findById("wnd[1]/usr/btn[0]").press
session.findById("wnd[1]/usr/sub/1/rad[1,0]").select
session.findById("wnd[1]/usr/sub/1/rad[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/btn[10]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,1]").text = "EX04"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,2]").text = "EX06"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,3]").text = "DCRE"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[3]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[1]").text = "25112023"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[2]").text = "155003"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[4]").text = "162226"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[6]").text = "93024859"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").key = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/rad[1]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/rad[1]").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/cmb").key = "03"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/ctxt[0,0]").text = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/txt[2,0]").text = "1"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").key = "R"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "O script foi conclu do.", 10, "SOLI OS - Cosern", 64


'Script gerado por SOLI OS em 26/11/2023 20:57:09
'Desligamento BAIXA DE MEDIDOR E RAMAL MONO COM LEITURA	
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
session.findById("wnd[0]/usr/ctxt").text = "4900898106"
session.findById("wnd[0]/usr/ctxt").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").caretPosition = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 2
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 4
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 5
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 6
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 7
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 8
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 9
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 10
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 11
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 12
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 13
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 14
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 15
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").text = "Grava  o: N o -  rea de Risco"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/sub/1/btn[2]").press
session.findById("wnd[1]/usr/btn[0]").press
session.findById("wnd[1]/usr/sub/1/rad[1,0]").select
session.findById("wnd[1]/usr/sub/1/rad[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/btn[10]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,1]").text = "EX04"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,2]").text = "EX06"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,3]").text = "DCRE"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[3]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[1]").text = "25112023"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[2]").text = "091613"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[4]").text = "093325"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[6]").text = "93024467"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").key = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_FLCAD/txt[2,0]").text = "3326"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/ctxt[0,0]").text = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/txt[2,0]").text = "1"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").key = "R"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "O script foi conclu do.", 10, "SOLI OS - Cosern", 64


'Script gerado por SOLI OS em 26/11/2023 20:57:09
'Desligamento BAIXA DE MEDIDOR MONO COM LEITURA
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
session.findById("wnd[0]/usr/ctxt").text = "4900898112"
session.findById("wnd[0]/usr/ctxt").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").caretPosition = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 2
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 4
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 5
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 6
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 7
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 8
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 9
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 10
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 11
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 12
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 13
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 14
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 15
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").text = "Grava  o: Sim - Gravado"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/sub/1/btn[2]").press
session.findById("wnd[1]/usr/btn[0]").press
session.findById("wnd[1]/usr/sub/1/rad[1,0]").select
session.findById("wnd[1]/usr/sub/1/rad[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/btn[10]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,1]").text = "EX04"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,2]").text = "DCRE"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,3]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,3]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").select
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[3]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[1]").text = "25112023"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[2]").text = "162345"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[4]").text = "162950"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[6]").text = "93024859"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").key = "2"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_FLCAD/txt[2,0]").text = "2438"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/ctxt[0,0]").text = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/txt[2,0]").text = "1"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").key = "R"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "O script foi conclu do.", 10, "SOLI OS - Cosern", 64


'Script gerado por SOLI OS em 26/11/2023 20:57:09
'Desligamento BAIXA DE MEDIDOR E RAMAL MONO COM LEITURA	
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
session.findById("wnd[0]/usr/ctxt").text = "4900898436"
session.findById("wnd[0]/usr/ctxt").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").caretPosition = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 2
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 4
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 5
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 6
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 7
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 8
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 9
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 10
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 11
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 12
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 13
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 14
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 15
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").text = "Grava  o: N o -  rea de Risco"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/sub/1/btn[2]").press
session.findById("wnd[1]/usr/btn[0]").press
session.findById("wnd[1]/usr/sub/1/rad[1,0]").select
session.findById("wnd[1]/usr/sub/1/rad[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/btn[10]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,1]").text = "EX04"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,2]").text = "EX06"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,3]").text = "DCRE"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[3]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[1]").text = "25112023"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[2]").text = "083939"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[4]").text = "085510"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[6]").text = "93024467"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").key = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_FLCAD/txt[2,0]").text = "5210"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/ctxt[0,0]").text = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/txt[2,0]").text = "1"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").key = "R"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "O script foi conclu do.", 10, "SOLI OS - Cosern", 64


'Script gerado por SOLI OS em 26/11/2023 20:57:10
'Desligamento BAIXA DE MEDIDOR E RAMAL MONO COM LEITURA	
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
session.findById("wnd[0]/usr/ctxt").text = "4900898555"
session.findById("wnd[0]/usr/ctxt").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").caretPosition = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 1
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 2
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 4
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 5
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 6
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 7
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 8
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 9
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 10
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 11
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 12
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 13
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 14
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT").verticalScrollbar.position = 15
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").text = "Grava  o: N o -  rea de Risco"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssub/2/3/sub/2/3/1/sub/2/3/1/1/tblSAPLIQS0TEXT/txt[0,0]").setFocus
session.findById("wnd[0]/usr/sub/1/btn[2]").press
session.findById("wnd[1]/usr/btn[0]").press
session.findById("wnd[1]/usr/sub/1/rad[1,0]").select
session.findById("wnd[1]/usr/sub/1/rad[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/btn[10]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,1]").text = "EX04"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,2]").text = "EX06"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,3]").text = "DCRE"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssub/2/3/tblSAPLIQS0AKTIONEN_VIEWER/ctxt[2,4]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[3]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[1]").text = "25112023"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[2]").text = "161036"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[4]").text = "161158"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/ctxt[6]").text = "93024467"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").key = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/cmb").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC1/ssub/2/3/1/1/1/1/2/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC3/ssub/2/3/1/1/1/1/2/chk[0]").selected = true
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC4/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_FLCAD/txt[2,0]").text = "1777"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/ctxt[0,0]").text = "0"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/txt[2,0]").text = "1"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").key = "R"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB03/ssub/2/3/sub/2/3/1/sub/2/3/1/1/sub/2/3/1/1/1/tabsTABSTRIP/tabpTS_FC6/ssub/2/3/1/1/1/1/2/tblSAPLXQQMTC_MAT/cmb[4,0]").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "O script foi conclu do.", 10, "SOLI OS - Cosern", 64


