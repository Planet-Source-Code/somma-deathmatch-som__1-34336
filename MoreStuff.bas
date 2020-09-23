Attribute VB_Name = "modMoreStuff"
Option Explicit
Dim Free As Byte

Type Moves
    Right As Single
    Down As Single
End Type


'Public Function MoveHowMuch(ByVal MoveDir As String, ByVal Degrees As Integer) As Moves
Public Sub MoveHowMuch(ByVal UpDown As String, ByVal Degrees As Single, ByRef X As Single, ByRef Y As Single)

'This sub returns the x and y distance that will be moved in one unit given the
'direction and the angle. These maths functions are now used instead of the
'commented-out crap below so that objects can now move at any angle,
'not just any angle divisible by 45

Y = Hor * Cos(DegToRad(ConvDeg(360 - Degrees)))
X = Hor * Sin(DegToRad(ConvDeg(360 - Degrees)))

'Up is just the opposite of down, so reverse directions
If UpDown = "up" Then
    X = X * -1
    Y = Y * -1
End If

'This is the perfect example of why Case-bashing is usually not the
'best way to solve a problem. A few simple maths functions can do
'what the select case below does, and more.

'Stuff i dont use anymore. Enjoy my commented-out code!
'Dim MoveRight As Integer
'Dim MoveDown As Integer
'Dim NewHor As Integer
'Dim NewDiag As Integer
'
'Dim UpDown As Integer

'Select Case Degrees
'    Case 0
'        MoveHowMuch.Right = 0
'        MoveHowMuch.Down = Hor
'    Case 45
'        MoveHowMuch.Right = -Diag
'        MoveHowMuch.Down = Diag
'    Case 90
'        MoveHowMuch.Right = -Hor
'        MoveHowMuch.Down = 0
'    Case 135
'        MoveHowMuch.Right = -Diag
'        MoveHowMuch.Down = -Diag
'    Case 180
'        MoveHowMuch.Right = 0
'        MoveHowMuch.Down = -Hor
'    Case 225
'        MoveHowMuch.Right = Diag
'        MoveHowMuch.Down = -Diag
'    Case 270
'        MoveHowMuch.Right = Hor
'        MoveHowMuch.Down = 0
'    Case 315
'        MoveHowMuch.Right = Diag
'        MoveHowMuch.Down = Diag
'End Select
'
'If MoveDir = "up" Then
'    MoveHowMuch.Down = MoveHowMuch.Down * -1
'    MoveHowMuch.Right = MoveHowMuch.Right * -1
'End If



End Sub

Public Function WhoseKey(Code As Integer) As Integer

'Finds which player uses the key.

'If Code = vbKeyLeft Or Code = vbKeyRight Or Code = vbKeyUp Or Code = vbKeyDown Or Code = 96 Or Code = 97 Or Code = 98 Then
'    WhoseKey = 1
'ElseIf Code = Asc("A") Or Code = Asc("D") Or Code = Asc("W") Or Code = Asc("S") Or Code = Asc("G") Or Code = Asc("T") Or Code = Asc("Y") Then
'    WhoseKey = 0
'Else
'    WhoseKey = -1
'End If

Dim i As Integer
Dim n As Integer

WhoseKey = -1
For i = 0 To UBound(Face)
    With Face(i)
        For n = 0 To UBound(.PKeys)
            If Code = .PKeys(n) Then
                WhoseKey = i
            End If
        Next n
    End With
Next i


End Function

Public Function FindKeyWord(ByVal Words As String, Optional Seperator As String = " ") As String

'This commented out bit shows a very complicated way of doing something very simple
'It's obvious I wrote it a while ago :)
'Dim i As Integer
'
'If seperator does not exist, assumes that Words is just
'the keyword with no data
'If InStr(Words, Seperator) <> 0 Then
'    For i = 0 To Len(Words)
'        If InStr(Left(Words, i), Seperator) > 0 Then
'            FindKeyWord = Trim(Left(Words, i - 1))
'            Exit For
'        End If
'    Next i
'Else
'    FindKeyWord = Words
'End If

'Now *this* is how it should be done
Dim StartHere As Integer

StartHere = InStr(Words, Seperator)
If StartHere <> 0 Then
    FindKeyWord = Trim(Left(Words, StartHere - 1))
Else
    FindKeyWord = Trim(Words)
End If
    
End Function


'Public Function FindNewSet(ByVal Words As String, Keyword As String) As String
'
'If Words <> "" Then
'    FindNewSet = Right(Words, Len(Words) - Len(Keyword) - 1)
'End If
'
'End Function
Public Function FindNewSet(Words As String, Optional Seperator As String = " ") As String

'If seperator does not exist, or words are nothing, then
'findnewset is nothing
If Words <> "" And InStr(Words, Seperator) <> 0 Then
    FindNewSet = Trim(Right(Words, Len(Words) - InStr(Words, Seperator)))
End If

End Function


Public Function SetIndMain(ByVal Desc As String) As Byte

Dim i As Byte
Dim FoundWord As Boolean

For i = LBound(BSettings) To UBound(BSettings)
    If BSettings(i).SetDesc = UCase(Desc) Then
        SetIndMain = i
        FoundWord = True
        Exit For
    End If
Next i

If FoundWord = False Then SetIndMain = 255

End Function

Public Function SetInd(ByVal Keyword As String, ByRef WhatArray() As SettingsBetter) As Byte

Dim i As Byte
Dim FoundWord As Boolean

For i = LBound(WhatArray) To UBound(WhatArray)
    If WhatArray(i).SetDesc = Keyword Then
        SetInd = i
        FoundWord = True
        Exit For
    End If
Next i

If FoundWord = False Then SetInd = 255

End Function

Public Function ShotNameFind(ByVal ShotIndex As Byte, WhichList() As ShotSettings) As String
Dim NameStr As String

'ShotNameFind = ShotSets(ShotIndex).Setting(SetInd("SHOT", ShotSets(ShotIndex).Setting)).SettingData
'ShotNameFind = WhichList(ShotIndex).Setting(SetInd(NameStr, WhichList(ShotIndex).Setting)).SettingData
ShotNameFind = WhichList(ShotIndex).Setting(0).SettingData

End Function

Public Function ShotIndex(ShotName As String, WhichList() As ShotSettings) As Integer

Dim n As Integer
For n = LBound(WhichList) To UBound(WhichList)
    If ShotNameFind(n, WhichList) = ShotName Then
        ShotIndex = n
        Exit For
    End If
Next n

End Function

Public Function ShotData(ByVal ShotWhat As Integer, Data As String) As String

With ShotSets(ShotWhat)
    ShotData = .Setting(SetInd(Data, .Setting)).SettingData
End With

End Function


Public Function FindStepRight(ByVal Keyword As String) As Byte
If IsNumeric(Right(Keyword, 1)) Then
    Dim StepRight As Byte
    Dim n As Integer
    For n = 1 To 3
        If IsNumeric(Right(Keyword, n)) = False Then
            'Must be numeric when n=1, because this loop
            'is in an if block which deciedes this
            FindStepRight = n - 1
            Exit For
        End If
    Next n
End If

End Function

Public Sub GetKeys()

'Gets the user-defined keys from the registry.
Dim Temp(1) As String
Dim i As Integer
Temp(0) = GetSetting(AppName, SetSec, "Player1Keys", "87,83,65,68,71,84,89")
Temp(1) = GetSetting(AppName, SetSec, "Player2Keys", "38,40,37,39,96,97,98")

For i = 0 To UBound(Face(0).PKeys)
    Face(0).PKeys(i) = FindBetweenCommas(Temp(0), i + 1, AComma)
    Face(1).PKeys(i) = FindBetweenCommas(Temp(1), i + 1, AComma)
Next i


End Sub

Public Sub ChSettings(ByRef WhichArray() As SettingsBetter, ByRef ListNum As String, ByRef InputBoxes As Object)

'Changes settings given the array settings are stored in,
'the list number of what needs changing(if any) e.g SHOT1
'and the control array of text boxes which new the
'settings will be read from

'Dim Free As Byte
'Dim SetsStore As String
'Dim Keyword As String
'Dim NewSetting As String
'Dim Temp As String
'
'
'Free = FreeFile
'Open data file
'Open "data.pos" For Input As #Free

'Do Until EOF(Free)
'    'Inputs one line at a time
'    Input #Free, Temp
'
'    'Finds keyword of the line
'    Keyword = FindKeyWord(Temp)
'    'Checks to see if it is part of a list
'    If IsNumeric(Right(Keyword, 1)) Then
'        If Keyword = Left(Keyword, Len(Keyword) - FindStepRight(Keyword)) & ListNum Then
'            Keyword = Left(Keyword, Len(Keyword) - FindStepRight(Keyword))
'        End If
'    End If
'
'    If SetInd(Keyword, WhichArray) <> 255 And Keyword <> "" Then
'        SetsStore = SetsStore & Keyword & ListNum & " " & InputBoxes(SetInd(Keyword, WhichArray)).Text & Chr(13)
'    Else
'        SetsStore = SetsStore & Temp & Chr(13)
'    End If
'Loop
'Close #Free
'
''txtTemp.Text = SetsStore
'Free = FreeFile
'
'Open "data.pos" For Output As #Free
'
'Print #Free, Trim(SetsStore)
'
'Close #Free

Dim i As Integer
For i = LBound(WhichArray) To UBound(WhichArray)
    WhichArray(i).SettingData = InputBoxes(i).Text
Next i

Call SaveFileSets
End Sub

Public Sub SaveFileSets()
Dim i As Integer
Dim n As Integer
Free = FreeFile

Open SetFile For Output As #Free

Print #Free, "'Please do not change anything unless you know what you're doing!"
For i = LBound(BSettings) To UBound(BSettings)
    Print #Free, WriteSet(BSettings(i))
Next i

'Leave a space
Print #Free, ""

Print #Free, "STARTSHOTS"
Print #Free, ""
Call WriteList(ShotSets)
Print #Free, ""

Print #Free, "STARTPOWERS"
Print #Free, ""
Call WriteList(PowerUps)
Print #Free, ""

Print #Free, "STARTCLASSES"
Print #Free, ""
Call WriteList(Classes)

Close #Free

'This needs to be done because if the 'chance' setting
'of a powerup gets changed, the sum of all the chances
'would have changed. THIS IS WHAT HAS BEEN CAUSING ALL
'THOSE OVERFLOW ERRORS
Call frmPositions.SetUpPower
End Sub

Public Sub WriteList(ByRef WhichList() As ShotSettings)

'This writes shotsetting-style settings into the data file.
'Includes settings for powerups, weapons, and classes
Dim i As Integer
Dim n As Integer
For i = LBound(WhichList) To UBound(WhichList)
    With WhichList(i)
        For n = LBound(.Setting) To UBound(.Setting)
            Print #Free, WriteSet(.Setting(n), i)
        Next n
    End With
    Print #Free, ""
Next i

End Sub

Public Function WriteSet(WhatSettings As SettingsBetter, Optional ListNum As Integer = -1) As String

'Gets the setting's description and data and puts it in a string

Dim Num As String
If ListNum >= 0 Then
    Num = ListNum
End If
WriteSet = WhatSettings.SetDesc & Num & " " & WhatSettings.SettingData

End Function


Public Sub NumsOnlyUpdate(WhichBoxes As Object, WhichArray() As NumsOnly)

''NumsOnly' stuff. See frmoptions for details
Dim i As Byte
For i = WhichBoxes.LBound To WhichBoxes.UBound
    WhichArray(i).OldText = WhichBoxes(i).Text
    WhichArray(i).SelPosition = WhichBoxes(i).SelStart
Next i

End Sub

Public Sub NumsOnlyCheck(ByRef WhichBoxes As Object, Index As Integer, WhichArray() As NumsOnly, WhichSettings() As SettingsBetter)

If IsNumeric(WhichSettings(Index).SettingData) Then
    Dim NotNumber As Boolean
    
    If IsNumeric(WhichBoxes(Index).Text) Then
'        If CStr(Int(WhichBoxes(Index).Text)) = WhichBoxes(Index).Text Then
            Call NumsOnlyUpdate(WhichBoxes, WhichArray)
'        Else
'            NotNumber = True
'        End If
    Else
        NotNumber = True
    End If
    
    If NotNumber = True Then
        Beep
        WhichBoxes(Index).Text = WhichArray(Index).OldText
        WhichBoxes(Index).SelStart = WhichArray(Index).SelPosition
    End If
End If

End Sub

Public Function FindBetweenCommas(ByVal Words As String, ByVal CommaNumber As Byte, Optional Seperator As String = ";") As String

'Just a whole lot of string manipulations
'Simply finds a string in between two semicolons
'(used to be commas, hence the name)

'Words should be in the format "something;more stuff;even more stuff",
'where each data string is seperated by the seperator (default semicolon)
Dim i As Integer
Dim StartHere As Integer

If Words <> "" Then

    'This stops the last data to get cut off because there is no separator after it
    Words = Words & Seperator

    StartHere = FindStartHere(Words, CommaNumber, Seperator)
    
    'Loads the string in between seprators
    For i = StartHere To Len(Words)
        If Mid(Words, i, 1) = Seperator Or i = Len(Words) Then
            FindBetweenCommas = Trim(Mid(Words, StartHere, i - StartHere))
            Exit For
        End If
    Next i
End If

End Function

Public Function FindStartHere(ByVal Words As String, ByVal Target As Integer, Sepa As String) As Integer

'Used by FindBetweenCommas to find where to start reading a data from a 3;4;32
'style string
Dim Count As Integer
Dim i As Integer


For i = 1 To Len(Words)
    If Mid(Words, i, 1) = Sepa Then
        Count = Count + 1
    End If
    
    'In 1;2;3;4 for example, to find data 2, the program stops at the 1st
    'semicolon. Hence target-1.
    If Count = Target - 1 Then
        If Count = 0 Then
            FindStartHere = i
        Else
            FindStartHere = i + 1
        End If
        Exit For
    End If
        
    If i = Len(Words) Then
        FindStartHere = Len(Words)
    End If
    
Next i

End Function

Public Function HasLength(ByRef WhatArray() As SettingsBetter) As Boolean
On Error Resume Next
Dim a As Integer
a = UBound(WhatArray)
If Err > 0 Then
    HasLength = False
Else
    HasLength = True
End If

End Function

Public Sub SetUpListSets(WhatSetting() As ShotSettings, ByVal Index As Integer)

If UBound(WhatSetting) < Index Then
    ReDim Preserve WhatSetting(Index)
End If
If HasLength(WhatSetting(Index).Setting) = False Then
    ReDim WhatSetting(Index).Setting(0)
End If

End Sub

Public Sub LoadSets(WhatSetting() As SettingsBetter, ByVal ListName As String, ByVal NewSet As String)

If HasLength(WhatSetting) = False Then
    ReDim WhatSetting(0)
End If
If WhatSetting(0).SetDesc <> "" Then
    ReDim Preserve WhatSetting(UBound(WhatSetting) + 1)
End If
WhatSetting(UBound(WhatSetting)).SetDesc = ListName
WhatSetting(UBound(WhatSetting)).SettingData = NewSet

End Sub

Public Sub SetErr(ErrDes As String, Optional ErrType As Byte)

'Make sure it wont repeat itself if the error occurs more than once
'in a timer event, such as a shot moving 3 units. Also, if there is
'an error during loading, there will probably be another error later
'on as well.
Dim WhatCap As String
If Not BadData Then
    
    Dim ErrMsg As String
    
    'Selects which message to show. There are only 2 now, but there could be
    'more in the future, which is why errtype is not a boolean.
    Select Case ErrType
        Case 0
            ErrMsg = "There is a problem with your settings file." & Chr(13) & "You probably entered invalid data in the Options dialog, such as text for a numerical setting." & Chr(13) & "Open" & SetFile & " or the Options dialog to correct any invalid data."
            WhatCap = SBrack & "Bad Settings" & EBrack
        Case 1
            ErrMsg = "A bitmap that Positions needs is missing. Make sure all neccessary pictures are in their appropriate folders." & Chr(13) & "Go to 'Change Skins' and make sure the file paths are correct."
            WhatCap = SBrack & "Missing Picture" & EBrack
    End Select

    'A universal error message for invalid settings
    MsgBox "The error is: " & ErrDes & Chr(13) & Chr(13) & ErrMsg & Chr(13) & Chr(13) & "The program must be restarted for new settings to take affect.", vbExclamation
    frmPositions.lblPaused.Caption = WhatCap
    
    BadData = True
    
End If

End Sub

Sub Wait(Seconds As Single, Optional Freeze As Boolean = False)

'Waits for a while. Used, for example, in the ending sequence where
'the winner is displayed.
Dim StartTime As Single
StartTime = Timer

Do Until Timer >= StartTime + Seconds
    If Not Freeze Then
        DoEvents
    End If
Loop


End Sub
