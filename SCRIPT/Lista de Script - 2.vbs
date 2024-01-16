Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Caminho da pasta dos scripts
scriptFolderPath = "D:\Users\CLS485467\Desktop\SCRIPT"

Set objFolder = objFSO.GetFolder(scriptFolderPath)
scriptFiles = Array()

For Each objFile In objFolder.Files
    If LCase(objFSO.GetExtensionName(objFile.Name)) = "vbs" Then
        ReDim Preserve scriptFiles(UBound(scriptFiles) + 1)
        scriptFiles(UBound(scriptFiles)) = objFile.Path
    End If
Next

If UBound(scriptFiles) >= 0 Then
    scriptNames = ""
    For i = LBound(scriptFiles) To UBound(scriptFiles)
        scriptNames = scriptNames & (i + 1) & ". " & objFSO.GetBaseName(scriptFiles(i)) & vbCrLf
    Next
    
    result = InputBox("Selecione o número do script que deseja executar:" & vbCrLf & vbCrLf & scriptNames & vbCrLf & vbCrLf & "Digite o número do script:", "Executar Script")
    
    If result <> "" Then
        selectedScriptIndex = CInt(result) - 1
        If selectedScriptIndex >= 0 And selectedScriptIndex <= UBound(scriptFiles) Then
            objShell.Run """" & scriptFiles(selectedScriptIndex) & """", 1, False
        Else
            MsgBox "Número de script inválido.", vbExclamation, "Executar Script"
        End If
    End If
Else
    MsgBox "Nenhum script encontrado na pasta.", vbExclamation, "Executar Script"
End If
