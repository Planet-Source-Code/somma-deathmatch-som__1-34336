VERSION 5.00
Begin VB.Form frmChoseShots 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3135
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer WaitForOp 
      Interval        =   200
      Left            =   3840
      Top             =   1080
   End
   Begin VB.Timer AIWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   120
   End
   Begin VB.Label lblClass 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   270
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblClass 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblReady 
      BackStyle       =   0  'Transparent
      Caption         =   "Ready..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   270
      Index           =   1
      Left            =   2760
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblReady 
      BackStyle       =   0  'Transparent
      Caption         =   "Ready..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblInstr 
      BackStyle       =   0  'Transparent
      Caption         =   "Press [Fire] when done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label lblheading 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   270
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1350
   End
   Begin VB.Label lblheading 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1350
   End
   Begin VB.Label lblP2Weps 
      BackStyle       =   0  'Transparent
      Caption         =   "Lightning"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label lblP1Weps 
      BackStyle       =   0  'Transparent
      Caption         =   "Lightning"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2100
   End
End
Attribute VB_Name = "frmChoseShots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurInd(1) As Byte
Dim ChoseDone(1) As Boolean
Dim WhatClass(1) As Byte

Dim DragForm As DragIt

Private Sub AIWait_Timer()

'The AI waits a second before chosing shots.
'This prevents an error I kept on getting when AIChose
'ran in the Form_Load event
Dim i As Byte
For i = 0 To UBound(Face)
    If IsAI(i) Then
        Call PickRanClass(i)
        Call ShowDone(i)
    End If
Next i

AIWait.Enabled = False

End Sub

Private Sub Form_Activate()

Dim i As Byte

AIWait.Enabled = True

For i = 0 To UBound(Face)
    lblHeading(i).Caption = Face(i).Name
    lblReady(i).Visible = False
    lblClass(i).Visible = (PlayType = 0)
Next i

Call PlaySound(App.Path & "/sounds/country4.wav")

For i = 0 To UBound(Face)
    Call DoClassWeps(i)
Next i

Call LoadLabels

For i = 0 To UBound(Face)
    ChoseDone(i) = False
    Call AIChose(i)
    Call MakeBold(i)
    Call ShowClass(i)
    Call ShotStats(i, CurInd(i))
Next i


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'Selecting shots is done with the player's keys
'The AI choses the shots another way, so it doesn't
'take up time.

Const SelSound = "sounds/type.wav"

'I was thinking of using another sound for this
Const ChSound = "sounds/keystrok.wav"

Dim Player As Integer
Dim DontShow As Boolean

'Checks whose key is being pressed
Player = WhoseKey(KeyCode)

If (Player = 0 Or Player = 1) Then

    'Checks if its a network game, and the key pressed
    'belongs to the player on the computer
    If GameType = 0 Or (GameType = Player + 1) Then
        If Not IsAI(Player) Then
            If Not ChoseDone(Player) Then
                With Face(Player)
                    Select Case KeyCode
                    
                        'Up
                        Case .PKeys(0)
                            CurInd(Player) = ConvShot(CurInd(Player) - 1, UBound(.Weps))
                            Call MakeBold(Player)
                            Call PlaySound(SelSound)
                            
                        'Down
                        Case .PKeys(1)
                            CurInd(Player) = ConvShot(CurInd(Player) + 1, UBound(.Weps))
                            Call MakeBold(Player)
                            Call PlaySound(SelSound)
                            
                        'Left
                        Case .PKeys(2)
                            '.Weps(CurInd(Player)) = ConvShot(.Weps(CurInd(Player)) - 1)
                            Call NextShot(Player, CurInd(Player), True)
                            Call ShowShots(Player, CurInd(Player))
                            Call PlaySound(ChSound)
                            
                        'Right
                        Case .PKeys(3)
                            '.Weps(CurInd(Player)) = ConvShot(.Weps(CurInd(Player)) + 1)
                            Call NextShot(Player, CurInd(Player))
                            Call ShowShots(Player, CurInd(Player))
                            Call PlaySound(ChSound)
                        
                        Case .PKeys(5)
                            WhatClass(Player) = ConvShot(WhatClass(Player) - 1, UBound(Classes))
                            Call ShowClass(Player)
                            Call PlaySound(SelSound)
                            
                        Case .PKeys(6)
                            WhatClass(Player) = ConvShot(WhatClass(Player) + 1, UBound(Classes))
                            Call ShowClass(Player)
                            Call PlaySound(SelSound)
                            
                        'Fire
                        Case .PKeys(4)
                            Call ShowDone(Player)
                            DontShow = True
                            
                    End Select
                    
                    If Not DontShow Then
                        Call ShotStats(Player, CurInd(Player))
                    End If
                    
                End With
            End If
        End If
    End If
End If
            
End Sub

Sub NextShot(ByVal WFace As Byte, WInd As Byte, Optional Previous As Boolean)

Dim UpDown As Integer
Dim AllDone As Boolean
Dim StartWep As Byte
Dim i As Integer

If Previous Then
    UpDown = -1
Else
    UpDown = 1
End If

With Face(WFace)

    StartWep = .Weps(WInd)
    
    Do
        .Weps(WInd) = ConvShot(.Weps(WInd) + UpDown, UBound(ShotSets))
        
        If PlayType = 0 Then
            If CanUseThis(.Weps(WInd), WhatClass(WFace)) Then
                AllDone = True
            End If
        Else
            AllDone = True
        End If
    Loop Until AllDone Or .Weps(WInd) = StartWep

End With

End Sub

Sub ShowDone(ByVal Player As Byte)

'Shows that the player is done selecting shots

Const DoneCol = vbRed

Dim i As Byte

ChoseDone(Player) = True

'Plays a sound
Call PlaySound("sounds/hydrod1.wav", 1)

If Player = 0 Then
    For i = 0 To lblP1Weps.UBound
        lblP1Weps(i).FontBold = True
        lblP1Weps(i).ForeColor = DoneCol
    Next i
Else
    For i = 0 To lblP2Weps.UBound
        lblP2Weps(i).FontBold = True
        lblP2Weps(i).ForeColor = DoneCol
    Next i
End If

Call frmPositions.ShotChange(Player, True)

If GameType = 0 Then
    If ChoseDone(0) And ChoseDone(1) Then
        Call Finished
    End If
Else
    Call SendReady
End If


End Sub

Sub Finished()

'Waits for 2 seconds
Call Wait(1.5)
Unload Me

End Sub

Sub MakeBold(ByVal Player As Byte)

'Boldens the weapon slot currently selected
Dim i As Byte

If Player = 0 Then
    
    For i = 0 To lblP1Weps.UBound
        lblP1Weps(i).FontBold = False
    Next i
    lblP1Weps(CurInd(Player)).FontBold = True
    
Else
    For i = 0 To lblP2Weps.UBound
        lblP2Weps(i).FontBold = False
    Next i
    lblP2Weps(CurInd(Player)).FontBold = True
End If

End Sub

Sub LoadLabels()

'This sub loads all the labels when the form is first loaded
Const ASpace = 500
Dim i As Integer

For i = 0 To 20  'UBound(Face(0).Weps)
        
    
    With lblP1Weps(i)
    
        If i <= UBound(Face(0).Weps) Then
            If i > lblP1Weps.UBound Then
    
                If i > 0 Then
                    Load lblP1Weps(i)
                    .Top = lblP1Weps(i - 1).Top + lblP1Weps(i - 1).Height
                End If
                
            End If
                
            If GameType = 0 Or GameType = 1 Then
                .Visible = True
            End If
            
        ElseIf i <= lblP1Weps.UBound Then 'And i > UBound(Face(0).Weps) Then
            Unload lblP1Weps(i)
        End If

        
    End With
    

    With lblP2Weps(i)
    
        If i <= UBound(Face(1).Weps) Then
            If i > lblP2Weps.UBound Then
    
                If i > 0 Then
                    Load lblP2Weps(i)
                    .Top = lblP2Weps(i - 1).Top + lblP2Weps(i - 1).Height
                End If
                
            End If
    
            If GameType = 0 Or GameType = 2 Then
                .Visible = True
            End If
            
        ElseIf i <= lblP2Weps.UBound Then
            Unload lblP2Weps(i)
        End If

        
    End With
    
Next i

lblInstr.Top = lblP1Weps(lblP1Weps.UBound).Top + lblP1Weps(lblP1Weps.UBound).Height + ASpace
Me.Height = lblInstr.Top + lblInstr.Height + ASpace

End Sub

Sub RanChose(Player As Byte)

'Randomly choses shots
'Can produce duplicates
'Dim n As Byte
'Dim AllDone As Boolean
'Dim TestThis As Integer
'
'For n = 0 To UBound(Face(Player).Weps)
'
'    AllDone = False
'
'    Do
'        TestThis = SomRand(UBound(ShotSets), LBound(ShotSets))
'        If PlayType = 1 Or CanUseThis(TestThis, WhatClass(Player)) Then
'            Face(Player).Weps(n) = TestThis
'            AllDone = True
'        End If
'    Loop Until AllDone
'
'    Call ShowShots(Player, n)
'Next n

End Sub

Sub PickRanClass(ByVal Player As Byte)

WhatClass(Player) = SomRand(UBound(Classes), LBound(Classes))
Call ShowClass(Player)
Call Wait(0.1)

End Sub

Sub ShowShots(ByVal Player As Byte, LIndex As Byte)

'Shows the name of the selected shot in the label
If Player = 0 Then
    lblP1Weps(LIndex).Caption = ShotNameFind(Face(0).Weps(LIndex), ShotSets)
Else
    lblP2Weps(LIndex).Caption = ShotNameFind(Face(1).Weps(LIndex), ShotSets)
End If

End Sub

Sub ShotStats(ByVal Player As Byte, ByVal LIndex As Byte)

'Shows the currently selected shot in the stats window
If Player = 0 Then
    Call frmPositions.ShowShotStuff(0, Face(0).Weps(LIndex))
Else
    Call frmPositions.ShowShotStuff(1, Face(1).Weps(LIndex))
End If

End Sub

Sub SrvAcceptClass(ByVal NewClass As Byte)

Dim Who As Byte
Who = FindOp(GameType - 1)

WhatClass(Who) = NewClass
Call ShowClass(Who, True)

End Sub

Sub ShowClass(ByVal Player As Byte, Optional AlreadySent As Boolean)

'Sets the class characteristics when user is cycling through the classes
'This is so that the user can see some of the things like hitpoints, damage
'bonus, and armour of the class while going through them
If PlayType = 0 Then

    'Sets the selected weapon to 0, so there wont be any errors when going from
    'classes with 4 weapon slots to 3 weapon slots
    CurInd(Player) = 0
    
    'Sets the number of weapon slots
    Call DoClassWeps(Player)
    Call LoadLabels

    lblClass(Player).Caption = ShotNameFind(WhatClass(Player), Classes)
    
    'Randomises the weapon slots with allowed weapons
    'Call RanChose(Player)
    
    'Goes through and makes sure there are no repetitions (not just done
    'with the AI anymore
    Call AIChose(Player)
    
    'Makes the selected shot bold
    Call MakeBold(Player)
    
    If GameType = 0 Or GameType = 1 Then
    
        'Note that it is quite possible to just have a .Class property and
        'read off the settings on runtime. However, the few extra bytes
        'of memory this takes up is worth it, as it makes things a lot easier
        With Classes(WhatClass(Player))
            Face(Player).ClassArm = .Setting(SetInd("CARM", .Setting)).SettingData
            Face(Player).ClassDam = .Setting(SetInd("CDAM", .Setting)).SettingData
            Face(Player).ClassHP = .Setting(SetInd("CHP", .Setting)).SettingData
            Face(Player).ClassShotSize = .Setting(SetInd("SHOTSIZE", .Setting)).SettingData
            Face(Player).ClassShotSpeed = .Setting(SetInd("SHOTSPEED", .Setting)).SettingData
            Face(Player).ClassSpeed = .Setting(SetInd("SPEED", .Setting)).SettingData
            Face(Player).ClassShotAccel = .Setting(SetInd("SHOTACCEL", .Setting)).SettingData
        End With
        
    End If
    
    'Sends the class type to the other computer if on network play
    If GameType <> 0 And Not AlreadySent Then
        Call SendClassType(WhatClass(Player))
    End If
    
    'Sets the modified HP
    With Face(Player)
        .HP = StartHP * .ClassHP
        frmStats.SbarHP(Player).Max = .HP 'StartHP * .ClassHP
    End With
    
    'Show the HP and damage/armour mods
    With frmPositions
        Call .ShowHP(Player)
        Call .ShowMods(Player)
    End With
    
    'Display the class type on the stats window
    frmStats.lblClassType(Player).Caption = ShotNameFind(WhatClass(Player), Classes)
    
    'Makes it visible (not visible in free-play)
    frmStats.lblClassType(Player).Visible = True
End If

End Sub

Sub DoClassWeps(ByVal Player As Byte)

'Sets the number of weapon slots the player can use
Dim Slots As Byte
Dim i As Byte

If PlayType = 0 Then
    'Playing with classes
    With Classes(WhatClass(Player))
        Slots = .Setting(SetInd("WEPSLOTS", .Setting)).SettingData - 1
    End With
Else
    'With free play
    Slots = BSettings(SetInd("PICKSHOTS", BSettings)).SettingData - 1
End If

With Face(Player)
    ReDim .Weps(Slots)
    For i = 0 To UBound(.Weps)
        .Weps(i) = -1
    Next i
End With

End Sub

Sub AIChose(ByVal Player As Byte)

'Picks weapons completely at random, but it will not pick duplicate shots
'Does it by picking a random weapon, and then checking it to see if it is
'valid (that is, the class can use it and it hasn't already been picked).
'If not, then it will go on to the weapon after it, and so on. Faster than
'picking another one at random

Dim i As Byte
Dim n As Integer
Dim ChoseDone As Boolean
Dim TestThis As Byte
Dim TestThisToo As Byte

With Face(Player)
    
    'Loop through each 'weapon slot'
    For i = 0 To UBound(.Weps)
        
        'Picks a random weapon
        TestThis = SomRand(UBound(ShotSets), LBound(ShotSets))
        n = 0
        ChoseDone = False
        
        Do
            'Checks if that weapon has already been picked
            TestThisToo = ConvShot(TestThis + n)
            If PlayType = 1 Or CanUseThis(TestThisToo, WhatClass(Player)) Then
            
                'Checks if its been chosen, or if there are more weapon slots than
                'available weapons
                If Not Chosen(TestThisToo, Player) Or (PlayType = 0 And i >= TotalWeps(WhatClass(Player))) Then
                
                    'I actually need to check this again, even though its part
                    'of the For-next condition, because it is somehow possible for
                    'the user to change classes half-way through the loop
                    If i <= UBound(.Weps) Then
                    
                        'Sets the weapon slot to the chosen weapon
                        .Weps(i) = TestThisToo
                        ChoseDone = True
                    Else
                        Exit For
                    End If
                End If
                
            End If
            
            'If the weapon has already been picked, goes on to
            'the one after it
            n = n + 1
        Loop Until ChoseDone Or n = UBound(ShotSets) + 1
        
        'Display the weapon chosen
        Call ShowShots(Player, i)
    Next i
    
End With


End Sub

Function Chosen(ByVal WhatShot As Byte, ByVal WFace As Byte) As Boolean

'This function checks whether 'WhatShot' has been chosen
Dim i As Byte

With Face(WFace)
    For i = LBound(.Weps) To UBound(.Weps)
        If WhatShot = .Weps(i) Then
            Chosen = True
            Exit For
        End If
    Next i
End With

End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

With DragForm
    .XStart = X
    .YStart = Y
    .Dragging = True
End With

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

With DragForm
    If .Dragging Then
        Me.Left = Me.Left + X - .XStart
        Me.Top = Me.Top + Y - .YStart
    End If
End With

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

DragForm.Dragging = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

DragForm.Dragging = False

End Sub

Private Sub WaitForOp_Timer()

If GameType <> 0 Then

    If OpChosen Then
        lblReady(FindOp(GameType - 1)).Visible = True
    End If

    If ChoseDone(GameType - 1) And OpChosen Then
        Call PlaySound("sounds/ending.wav")
        Call Finished
        OpChosen = False
    End If
End If
    
End Sub

Function CanUseThis(ByVal WepInd As Byte, WClass As Byte) As Boolean

'Checks if the class can use a certain weapon

Dim AllDone As Boolean
Dim TheSetting As String
Dim TheWep As String
Dim i As Byte

With Classes(WClass)
    TheSetting = .Setting(SetInd("WEPS", .Setting)).SettingData
End With

'i = 1
'Do
For i = 1 To TotalWeps(WClass)
    TheWep = FindBetweenCommas(TheSetting, i)
    If TheWep = Trim(Str(WepInd)) Then
        CanUseThis = True
        Exit For
        'AllDone = True
    ElseIf TheWep = "" Then
        Exit For
        'AllDone = True
    End If
Next i

'    i = i + 1
'Loop Until AllDone
    

End Function

Function TotalWeps(ByVal WClass As Byte) As Byte

Dim i As Byte
Dim TheWeps As String
Dim AmDone As Boolean

With Classes(WClass)
    TheWeps = .Setting(SetInd("WEPS", .Setting)).SettingData
End With

Do
    If FindBetweenCommas(TheWeps, i + 1) = "" Then
        TotalWeps = i
        AmDone = True
    End If
    i = i + 1
Loop Until AmDone
    

End Function
