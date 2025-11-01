' ****************************************************************
' Funktion:     Kopiert die Release-EXE an ihren Bestimmugsort.
' Autor:				Wilhelm Happe, 2012
' ****************************************************************
Option Explicit
On Error Resume Next
Dim objWS, objFSO
Dim strDatei, strQuellOrdner, strQuellPfad, strZielOrdner, strZielPfad
Const OverwriteExisting = True
Set objWS = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Einzige Vorgabe: Bitte anpassen!
strZielOrdner = "E:\tdHiDrive\Software\CSharp"
'Pfad des Skripts aus dessen Pfadnamen isolieren
strQuellOrdner = objFSO.GetParentFolderName(Wscript.ScriptFullName)
'Voraussetzung ist, dass Ordnername und ExeName übereinstimmen!
strDatei = objFSO.GetFolder(strQuellOrdner).Name & ".exe"
strQuellOrdner = strQuellOrdner & "\bin\Release\"
strQuellPfad = strQuellOrdner & strDatei

If Not objFSO.FolderExists(strZielOrdner) Then
  MsgBox "Der Zielpfad" & strZielOrdner & "wurde nicht gefunden!" & vbCr & _
	  "Die Pfadangaben im Skript sind möglichweise falsch.", vbCritical, WScript.ScriptName
  Wscript.Quit
End If

objFSO.CopyFile strQuellPfad, strZielOrdner & "\", OverwriteExisting
If Err <> 0 Then
  MsgBox "Fehler beim Kopieren der Datei '" & strDatei & "':" & vbCr & _
	  Err.Description, vbCritical, WScript.ScriptName
  Else
    MsgBox  "Die Datei wurde erfolgreich kopiert:" & vbCr & _
      strQuellPfad, vbInformation, WScript.ScriptName
  End If
Err.Clear

strZielPfad = strZielOrdner &  "\" & strDatei
objWS.Run strZielPfad, 1, False
If Err <> 0 Then
  MsgBox "Fehler beim Ausführen der Datei:" & vbCr & strZielPfad & _
    vbCr & Err.Description, vbCritical, WScript.ScriptName
End If

Wscript.Quit
