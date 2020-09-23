Attribute VB_Name = "modServeStuff"
Option Explicit

'My attempt at network play. It worked, but just too damn slowly even over
'a LAN, so I gave up. I just dont have enough experience yet, i guess

Public Const AComma = ","
Public Const ASemi = ";"

Public UnsentData As Boolean

Public NewData As String

Public OpReady As Boolean
Public OpChosen As Boolean

Public UsingShot As Integer

Dim SendThis As String

Public ErrorStop As Boolean

Public AlreadySent() As Boolean


Public Sub ServerSend()

If Not ErrorStop Then
    Dim i As Integer
    Dim n As Integer


    Dim CurShots As String

    For i = 0 To UBound(Shooting)
        With Shooting(i)
            If .What >= 0 Then
                Call MoreData("S" & " " & .What & AComma & .ShotX & AComma & .ShotY)
            End If
        End With
    Next i

    For i = LBound(PowerUpNow) To UBound(PowerUpNow)
        With PowerUpNow(i)
            If .What >= 0 Then
                Call MoreData("POW" & " " & .What & AComma & .X & AComma & .Y)
            End If
        End With
    Next i

    Call MoreData("ARM" & " " & Face(0).PlusArmour.TimeLeft & AComma & Face(1).PlusArmour.TimeLeft)
    Call MoreData("DAM" & " " & Face(0).PlusDam.TimeLeft & AComma & Face(1).PlusDam.TimeLeft)

    With frmGameType.LotsaSocks
        Call .SendStuff(SendThis)
        SendThis = ""
        If .WhatError <> "" Then
            MsgBox "The connection has been lost."
            ErrorStop = True
            Exit Sub
        End If
    End With

End If

End Sub

Sub MoreData(AddThis As String)

SendThis = SendThis & AddThis & ASemi

End Sub

Public Sub DataArrive(Data As String)

Dim i As Integer
Dim n As Integer
Dim DataDone As Boolean


'If GameType = 2 Then

Dim DataGroup As String
Dim Keyword As String
Dim NewData As String
Dim XVal As Single
Dim YVal As Single


If GameType = 2 Then
    'Covers up shots, faces, etc
    With frmPositions
        Call .CoverStuff

        For n = LBound(PowerUpNow) To UBound(PowerUpNow)
            If PowerUpNow(n).What >= 0 Then
                Call .CoverPowers(n)
            End If
        Next n

        Call .UpdateKaboom
        Call .UpdateTrails


        For n = 0 To UBound(Shooting)
            'If Shooting(i).What >= 0 Then
                Call frmPositions.ShotReset(n)
            'End If
        Next n

        For n = LBound(PowerUpNow) To UBound(PowerUpNow)
            Call .PowerUpReset(n)
        Next n

    End With
End If


i = 1
Do
    DataGroup = FindBetweenCommas(Data, i)
    Keyword = FindKeyWord(DataGroup)
    NewData = FindNewSet(DataGroup)

    If IsNumeric(Right(Keyword, 1)) Then

        Dim ListName As String
        Dim LInd As Integer
        Dim NumNums As Byte

        NumNums = FindStepRight(Keyword)

        LInd = Right(Keyword, NumNums)
        ListName = Left(Keyword, Len(Keyword) - NumNums)

        Select Case ListName

            Case "POS"
                With Face(LInd)
                    .X = FindBetweenCommas(NewData, 1, AComma)
                    .Y = FindBetweenCommas(NewData, 2, AComma)
                End With

            Case "DEG"
                Face(LInd).Degrees = NewData
            Case "HURT"
                Face(LInd).LastHurt = NewData

        End Select

    Else

        Select Case Keyword
            Case "HP"
                Face(0).HP = FindBetweenCommas(NewData, 1, AComma)
                Face(1).HP = FindBetweenCommas(NewData, 2, AComma)
            Case "ARM"
                Face(0).PlusArmour.TimeLeft = FindBetweenCommas(NewData, 1, AComma)
                Face(1).PlusArmour.TimeLeft = FindBetweenCommas(NewData, 2, AComma)

            Case "DAM"
                Face(0).PlusDam.TimeLeft = FindBetweenCommas(NewData, 1, AComma)
                Face(1).PlusDam.TimeLeft = FindBetweenCommas(NewData, 2, AComma)
            Case "S"

                Dim TrailNums As Integer
                For n = 0 To UBound(Shooting)
                    With Shooting(n)
                        If .What < 0 Then
                            .What = FindBetweenCommas(NewData, 1, AComma)
                            .ShotX = FindBetweenCommas(NewData, 2, AComma)
                            .ShotY = FindBetweenCommas(NewData, 3, AComma)
                            TrailNums = ShotSets(.What).Setting(SetInd("TRAIL", ShotSets(.What).Setting)).SettingData
                            If TrailNums > 0 Then
                                Call frmPositions.NewTrail(n, TrailNums)
                            End If
                            Exit For
                        End If
                    End With
                Next n
            Case "BOOM"
                XVal = FindBetweenCommas(NewData, 1, AComma)
                YVal = FindBetweenCommas(NewData, 2, AComma)
                Call frmPositions.DoKaboom(XVal, YVal, FindBetweenCommas(NewData, 3, AComma))
            Case "POW"
                For n = LBound(PowerUpNow) To UBound(PowerUpNow)
                    With PowerUpNow(n)
                        If .What < 0 Then
                            .What = FindBetweenCommas(NewData, 1, AComma)
                            .X = FindBetweenCommas(NewData, 2, AComma)
                            .Y = FindBetweenCommas(NewData, 3, AComma)
                            Exit For
                        End If
                    End With
                Next n

            Case "CHS"
                If GameType = 1 Then
                    UsingShot = NewData
                    Call frmPositions.ShowShotStuff(1, NewData)
                Else
                    Call frmPositions.ShowShotStuff(0, NewData)
                End If
'
            'Don't do this! Send the keys themselves!
'            Case "CHD"
'                Dim DegChange As Byte 'Integer
'                If NewData = True Then
'                    DegChange = 3 '45
'                Else
'                    DegChange = 2 '-45
'                End If
'
'                'DegreeChange(1) = DegChange
'                PressKeys(1, DegChange) = True
'
'            Case "MOV"
'                Dim Direction As Byte 'String
'                If NewData = True Then
'                    Direction = 1 '"down"
'                Else
'                    Direction = 0 '"up"
'                End If
'
'                'Face(1).UpDown = Direction
'                PressKeys(1, Direction) = True
'
'            Case "SMOV"
'
'                'Face(1).UpDown = ""
'                PressKeys(1, 0) = False
'                PressKeys(1, 1) = False
'
'            Case "SHOOT"
'
'                'Face(1).HoldShot = True
'                PressKeys(1, 4) = True
'
'            Case "STOPS"
'
'                'Face(1).HoldShot = False
'                PressKeys(1, 4) = False
'
'            Case "STURN"
'                'DegreeChange(1) = 0
'                PressKeys(1, 2) = False
'                PressKeys(1, 3) = False

            'Do this instead!
            Case "KEY"
                PressKeys(1, NewData) = True
            Case "KEYUP"
                PressKeys(1, NewData) = False


            Case "NAME"

                OpReady = True

                If GameType = 1 Then
                    Face(1).Name = NewData
                    frmStats.lblTeamName(1).Caption = NewData
                ElseIf GameType = 2 Then
                    Face(0).Name = NewData
                    frmStats.lblTeamName(0).Caption = NewData
                End If

            Case "READY"
                OpChosen = True

            Case "PTYPE"
                PlayType = NewData

            Case "CTYPE"
                Call frmChoseShots.SrvAcceptClass(NewData)


        End Select
    End If

    i = i + 1

    If DataGroup = "" Then
        DataDone = True
    End If

Loop Until DataDone

If GameType = 2 Then
    With frmPositions

        For n = 0 To UBound(Face)
            Call .ShowHP(n)
            Call .ShowMods(n)
            Call .ShowModTime(n)
        Next n

        Call .DrawTrails

        Call .DrawPowers
        For n = 0 To UBound(Shooting)
            If Shooting(n).What >= 0 Then
                Call .DrawShots(n)
            End If
        Next n

        Call .DrawKaboom

        Call .DrawFaces

        Call .CheckForWins

        Call .DoKeyStates

    End With

    frmPositions.Refresh
End If

'End If

End Sub

Sub SrvSendHP()

If GameType = 1 Then
    Call MoreData("HP" & " " & Face(0).HP & AComma & Face(1).HP)
End If

End Sub

Sub SrvSendPOS(ByVal WFace As Byte)

If GameType = 1 Then
    Call MoreData("POS" & WFace & " " & Face(WFace).X & AComma & Face(WFace).Y)
End If

End Sub

Sub SendName(YourName As String)

frmGameType.LotsaSocks.SendStuff ("NAME" & " " & YourName & ASemi)

End Sub

Sub SendReady()

Call frmGameType.LotsaSocks.SendStuff("READY" & ASemi)

End Sub

Sub SendShot(ByVal CurShot As Integer)

Call frmGameType.LotsaSocks.SendStuff("CHS" & " " & CurShot & ASemi)

End Sub

Sub SendDegMove(Right As Boolean)

Call frmGameType.LotsaSocks.SendStuff("CHD" & " " & Right & ASemi)

End Sub

Sub SendUpDown(Down As Boolean)

Call frmGameType.LotsaSocks.SendStuff("MOV " & Down & ASemi)

End Sub

Sub SendStopMove()

Call frmGameType.LotsaSocks.SendStuff("SMOV" & ASemi)

End Sub

Sub SendStopTurn()

Call frmGameType.LotsaSocks.SendStuff("STURN" & ASemi)

End Sub

Sub SendShoot()

Call frmGameType.LotsaSocks.SendStuff("SHOOT" & ASemi)

End Sub

Sub SendStopShoot()

Call frmGameType.LotsaSocks.SendStuff("STOPS" & ASemi)

End Sub

Sub SendLastHurt(WFace As Byte)

Call MoreData("HURT" & WFace & " " & Face(WFace).LastHurt)
Call SrvSendHP

End Sub

Sub SendClassType(WType As Byte)

Call frmGameType.LotsaSocks.SendStuff("CTYPE" & " " & WType)

End Sub


Sub LostConnect()

MsgBox "The connection has been lost."
Call frmPositions.NewGame

End Sub
