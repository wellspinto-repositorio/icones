' CreateSiciShortcut.vbs
' Rotina para criar atalho "Sici Imobilis" no Desktop com ícone remoto
' Auto-exclui após a execução

Option Explicit

Dim objShell, chromePath, desktopPath, shortcut, urlParams, chromeURL, iconURL, iconPath
Set objShell = WScript.CreateObject("WScript.Shell")

' Configurações
urlParams = "--app=""https://passeli.sici-imobilis.com.br/sici/"" --start-maximized --disable-extensions --disable-popup-blocking --force-device-scale-factor=1"
chromeURL = "https://www.google.com/chrome/"
iconURL = "https://raw.githubusercontent.com/wellspinto-repositorio/sici_original/refs/heads/main/imobilis.ico?token=GHSAT0AAAAAADDOV4RBW27H6OXQCEVBXSZW2A3F5AQ" ' SUBSTITUA pelo URL real do ícone
iconPath = objShell.ExpandEnvironmentStrings("%TEMP%") & "\sici_icon.ico"

' Verificar se o Chrome está instalado
On Error Resume Next
chromePath = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\")
If chromePath = "" Then
    chromePath = objShell.RegRead("HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\")
End If
On Error GoTo 0

If chromePath = "" Then
    ' Chrome não encontrado - oferecer instalação
    Dim answer
    answer = MsgBox("O Google Chrome não foi encontrado no seu sistema." & vbCrLf & _
                    "Deseja abrir o site para download e instalação?", _
                    vbQuestion + vbYesNo, "Chrome não instalado")
    
    If answer = vbYes Then
        objShell.Run chromeURL
    End If
    WScript.Quit
End If

' Baixar o ícone remoto
On Error Resume Next
DownloadFile iconURL, iconPath
If Err.Number <> 0 Then
    iconPath = chromePath & ",0" ' Usar ícone do Chrome se falhar o download
End If
On Error GoTo 0

' Criar atalho no Desktop
desktopPath = objShell.SpecialFolders("Desktop")
Set shortcut = objShell.CreateShortcut(desktopPath & "\Sici Imobilis.lnk")

With shortcut
    .TargetPath = chromePath
    .Arguments = urlParams
    .WindowStyle = 1 ' Normal window
    .IconLocation = iconPath
    .Description = "Acesso rápido ao Sici Imobilis"
    .Save
End With

' Auto-exclusão
DeleteFile WScript.ScriptFullName

Sub DownloadFile(url, path)
    Dim xhr, adoStream
    Set xhr = CreateObject("MSXML2.XMLHTTP")
    xhr.Open "GET", url, False
    xhr.Send
    
    If xhr.Status = 200 Then
        Set adoStream = CreateObject("ADODB.Stream")
        adoStream.Open
        adoStream.Type = 1 ' Binary
        adoStream.Write xhr.responseBody
        adoStream.SaveToFile path, 2 ' Overwrite
        adoStream.Close
    Else
        Err.Raise 1, "DownloadFile", "Falha ao baixar o arquivo"
    End If
End Sub

Sub DeleteFile(path)
    On Error Resume Next
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFile path, True
    Set fso = Nothing
    On Error GoTo 0
End Sub
