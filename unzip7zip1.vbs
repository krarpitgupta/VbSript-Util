Function Zip(sFile,sArchiveName)
  'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com

  Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")
  Set oShell = WScript.CreateObject("Wscript.Shell")

  '--------Find Working Directory--------
  aScriptFilename = Split(Wscript.ScriptFullName, "\")
  sScriptFilename = aScriptFileName(Ubound(aScriptFilename))
  sWorkingDirectory = Replace(Wscript.ScriptFullName, sScriptFilename, "")
  '--------------------------------------

  '-------Ensure we can find 7za.exe------
  If oFSO.FileExists(sWorkingDirectory & "\" & "7za.exe") Then
    s7zLocation = ""
  ElseIf oFSO.FileExists("C:\Program Files\7-Zip\7za.exe") Then
    s7zLocation = "C:\Program Files\7-Zip\"
  Else
    Zip = "Error: Couldn't find 7za.exe"
    Exit Function
  End If
  '--------------------------------------

  oShell.Run """" & s7zLocation & "7za.exe"" a -tzip -y """ & sArchiveName & """ " _
  & sFile, 0, True   

  If oFSO.FileExists(sArchiveName) Then
    Zip = 1
  Else
    Zip = "Error: Archive Creation Failed."
  End If
End Function

Function UnZip(sArchiveName,sLocation)
  'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com
debug
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oShell = CreateObject("Wscript.Shell")

  '--------Find Working Directory--------
  
'  aScriptFilename = Split(Wscript.ScriptFullName, "\")
'  sScriptFilename = aScriptFileName(Ubound(aScriptFilename))
'  sWorkingDirectory = Replace(Wscript.ScriptFullName, sScriptFilename, "")
  '--------------------------------------

  '-------Ensure we can find 7za.exe------
'  If oFSO.FileExists("C:\Program Files\7-Zip" & "\" & "7za.exe") Then
'    s7zLocation = ""
'  ElseIf oFSO.FileExists("C:\Program Files\7-Zip\7za.exe") Then
'    s7zLocation = "C:\Program Files\7-Zip\"
'  Else
'    UnZip = "Error: Couldn't find 7za.exe"
'    Exit Function
'  End If

   If oFSO.FileExists("C:\Program Files\7-Zip\7z.exe") Then
    s7zLocation = "C:\Program Files\7-Zip\"
  Else
    UnZip = "Error: Couldn't find 7za.exe"
    Exit Function
  End If




  '--------------------------------------

  '-Ensure we can find archive to uncompress-
  If Not oFSO.FileExists(sArchiveName) Then
    UnZip = "Error: File Not Found."
    Exit Function
  End If
  '--------------------------------------

  oShell.Run """" & s7zLocation & """ x -y -o""" & sLocation & """ """ & sArchiveName & """",,True
  UnZip = 1
End Function


Call UnZip("C:\test\TagIdFile.7z","C:\test\")
msgbox "done"



