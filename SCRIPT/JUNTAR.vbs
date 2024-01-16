Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\Users\userconsolis\Downloads\SCRIPT\UNIR")

' Nome do arquivo de saída (padrão)
strOutputFile = "C:\Users\userconsolis\Downloads\SCRIPT\UNIR\ArquivoUnido.vbs"

' Crie o arquivo de saída ou sobrescreva se já existir
Set objOutputFile = objFSO.CreateTextFile(strOutputFile, True)

' Loop através de todos os arquivos na pasta
For Each objFile In objFolder.Files
    ' Verifique se o arquivo não é o arquivo de saída
    If objFile.Path <> strOutputFile Then
        ' Abra o arquivo de entrada para leitura
        Set objInputFile = objFSO.OpenTextFile(objFile.Path, 1)
        
        ' Leia o conteúdo do arquivo de entrada e escreva no arquivo de saída
        strContent = objInputFile.ReadAll
        objOutputFile.WriteLine strContent
        
        ' Feche o arquivo de entrada
        objInputFile.Close
    End If
Next

' Feche o arquivo de saída
objOutputFile.Close

WScript.Echo "Arquivos unidos com sucesso em " & strOutputFile
