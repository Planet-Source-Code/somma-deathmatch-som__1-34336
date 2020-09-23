Attribute VB_Name = "modFileStuff"
Option Explicit

Public Function FileExists(FileName As String) As Boolean
'Checks if a file exists. There *has* to be
'a better way of doing this...

Dim CheckThis As String
On Error Resume Next

CheckThis = Dir(FileName)
If CheckThis = "" Then
    FileExists = False
Else
    FileExists = True
End If

End Function

