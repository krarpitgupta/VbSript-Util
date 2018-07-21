Unzip "C:\peeyush\notepad\zip\a.zip" , "C:\peeyush\notepad\zip"
msgbox "done"
Sub Unzip(sSource, sTargetDir)
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    if not oFSO.FolderExists(sTargetDir) then oFSO.CreateFolder(sTargetDir)
    Set oShell = CreateObject("Shell.Application")
    Set oSource = oShell.NameSpace(sSource).Items()
    Set oTarget = oShell.NameSpace(sTargetDir)
    oTarget.CopyHere oSource, 256
End Sub

'16 to overwrite file