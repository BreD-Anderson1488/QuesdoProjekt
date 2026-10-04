Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

Dim appDir
appDir = "D:\MY_PROJEKTS\MY_PROJEKTS\QuesdoProjekt\www"

Dim electronCmd
electronCmd = appDir & "\node_modules\.bin\electron.cmd"

If Not fso.FileExists(electronCmd) Then
    MsgBox "Electron is not installed!" & vbCrLf & _
           "Run: npm install --save-dev electron", _
           vbCritical, "QuestDo"
    WScript.Quit
End If

sh.CurrentDirectory = appDir

sh.Run Chr(34) & electronCmd & Chr(34) & " " & Chr(34) & appDir & Chr(34), 1, False