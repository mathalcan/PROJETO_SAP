Option Explicit

Dim objFSO, objFolder, objFile, strOutputFile, objTextStream

' Diretório onde estão os arquivos VBS
Const strDiretorio = "C:\Users\E815847\Desktop\SCRIPT\UNIR\"

' Nome do arquivo de saída
Const strNomeSaida = "ArquivoUnido.vbs"

' Crie uma instância do objeto FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Crie uma referência ao diretório
Set objFolder = objFSO.GetFolder(strDiretorio)

' Crie o arquivo de saída
Set objTextStream = objFSO.CreateTextFile(strDiretorio & strNomeSaida, True)

' Itere sobre todos os arquivos VBS no diretório
For Each objFile In objFolder.Files
    If LCase(objFSO.GetExtensionName(objFile.Path)) = "vbs" Then
        ' Abra o arquivo VBS e leia seu conteúdo
        Dim objInputFile, strContent
        Set objInputFile = objFSO.OpenTextFile(objFile.Path, 1) ' 1 = leitura
        strContent = objInputFile.ReadAll
        objInputFile.Close
        
        ' Escreva o conteúdo no arquivo de saída
        objTextStream.WriteLine strContent
    End If
Next

' Feche o arquivo de saída
objTextStream.Close

' Libere os objetos
Set objTextStream = Nothing
Set objFile = Nothing
Set objFolder = Nothing
Set objFSO = Nothing

WScript.Echo "Arquivos VBS foram unidos com sucesso em " & strDiretorio & strNomeSaida
