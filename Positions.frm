VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPositions 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   510
   ClientTop       =   1005
   ClientWidth     =   9510
   ForeColor       =   &H00000000&
   Icon            =   "Positions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Positions.frx":054A
   MousePointer    =   99  'Custom
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   4560
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picWinWep 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   1080
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   7
      Top             =   480
      Width           =   495
   End
   Begin VB.PictureBox picWinWep 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   1080
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblFrameRate 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label lblWinLose 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loser"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   270
      Index           =   1
      Left            =   2355
      TabIndex        =   9
      Top             =   570
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblWinKill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Winner"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   570
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblWinKill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Winner"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   210
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblWinLose 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loser"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   270
      Index           =   0
      Left            =   2355
      TabIndex        =   5
      Top             =   210
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblInstruct 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press F5 to continue..."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label lblWins 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "draw DRAW"
      BeginProperty Font 
         Name            =   "Matisse ITC"
         Size            =   65.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3135
      Left            =   3240
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblWinner 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Winner!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1095
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label lblPaused 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press F12"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   45
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuKeys 
         Caption         =   "Define &Keys"
      End
      Begin VB.Menu mnuBackGround 
         Caption         =   "Change &BackGround"
      End
      Begin VB.Menu mnuResize 
         Caption         =   "&Resize Game Window"
      End
      Begin VB.Menu mnuSkins 
         Caption         =   "Change &Skins"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "(&P)ause"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuClear 
         Caption         =   "Next &Round"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmPositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'>>Note to anyone who reads this<<
'----------------------------------
'After looking up some of the programs in Planet-source-code, I have realised that some of the
'methods I use here are not the best ways to do things. For example, I load all
'the pictures into separate picture boxes, when I could just load them into memory,
'and I dont use backbuffering. I have decided to let these go because 1) I can't
'be bothered learning a whole new set of api functions, and 2) I've just learnt
'DirectDraw anyway

'If you happen to find parts of this program that are badly coded, chances are i've
'realised this by the time you read it. That's how quickly im learning right now :)
'Still, you should leave a comment in PSC or e-mail me

'There are probably lots of other unconventional methods throughout the program,
'because I pretty much taught myself visual basic with this program. There are
'a few things which I had to learn from other sources, such as DirectSound and the
'API functions, but most of this is derived from my own thinking.

'Things I didn't learn myself (apart from the *really* basic stuff that my
'software design teacher tought me, like adding labels to forms) include:

'Automatic masking, thanks to Winter Grave's masking program
'The Bitblt API
'DirectSound, from Derek Hall's Directsound7 program
'SetPixelV and GetPixel, from Tanner Helland's tutorial
'The basics of the Winsock control
'All of these i found in PSC (planet-source-code)

'That's about it really. Believe it or not, everything else in here I learnt myself

'Note that although I have deleted a lot of old code, I have deliberately left
'some of it here, just in case some idiot thinks that I ripped the code off them
'If you're one of those idiots, I can tell you that I have archives of previous
'versions of this program going back a looooong way (even a v1.0.0 executable), so
'don't even think about it

'Also, although I have tried to put descriptive comments wherever I could,
'some of it might simply not make much sense to you. I never intended to release
'this as open-source until a while ago, so you should *feel lucky you're reading
'this at all*

Option Explicit

'This is now used for key-pressing. It can handle
'about six keys at once, rather than the 3 keys using form_keydown
'What's the difference between this and the GetAsyncKeyState function?
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

'For mask-making. A lot faster than pure vb
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

'Stores info on 'flowers'
Private Type Flower
    Shots As Integer
    ShotType As Integer
    ShotSpeed As Integer
    Pause As Integer
    PauseMax As Integer
    Player As Byte
    FollowYou As Byte
    X As Integer
    Y As Integer
End Type

Dim Flowers(5) As Flower
Dim KaboomNumbers As Byte

Dim GameOver As Boolean

Dim PowerTotal As Integer

Dim Loaded As Boolean

Dim NoRestart As Boolean

Dim FirstTime  As Boolean

Dim IsPaused As Boolean
Dim UnPausing As Boolean


Private Sub Form_Initialize()

FirstTime = True

'Sets the directory so program can read data files and pictures
ChDir App.Path

'Setting caption
Me.Caption = SBrack & "Deathmatch SOM" & EBrack & " v" & App.Major & "." & App.Minor & "." & App.Revision
Call ChColours(GetSetting("Positions", "Settings", "BackGround", vbBlack))

'Hides the main form while the Loading form shows
Me.Hide
Call NewGame 'Starts a new game

FirstTime = False

End Sub


Private Sub Form_Load()

'This is one of the oldest pieces of commented-out code i've left in here.

'YES, AT LAST! I HAVE A MUCH, MUCH, MUCH BETTER WAY OF DOING THIS NOW!!!
'NO HARD-CODING FOR ME!!!!!!
'-------------------------------
'We don't *have* to redim this. The array length can be set when its first
'declared. However, doing this makes it easier to manage when writing code
'ReDim BSettings(0)
'These things have to be hard-coded in, unfortunately
'BSettings(0).SetDesc = "PLAYERHP"
'BSettings(1).SetDesc = "EXPLOSIONWAIT"
''BSettings(2).SetDesc = "XCOORD"
''BSettings(3).SetDesc = "YCOORD"
'BSettings(2).SetDesc = "MOVEDIAG"
'BSettings(3).SetDesc = "MOVESTRAIGHT"
'BSettings(4).SetDesc = "DEFSHOTSPEED"
'BSettings(5).SetDesc = "DEFSHOTSOUND"
'BSettings(6).SetDesc = "GAMESPEED"

'Dim i As Byte
'For i = LBound(ShotSets) To UBound(ShotSets)
    'ReDim ShotSets(0)
'    With ShotSets(0)
'        ReDim .Setting(5)
'        .Setting(0).SetDesc = "SHOT"
'        .Setting(1).SetDesc = "SHOTDAM"
'        .Setting(2).SetDesc = "SHOTDAMLOW"
'        .Setting(3).SetDesc = "SHOTWAIT"
'        .Setting(4).SetDesc = "SHOTSPEED"
'        .Setting(5).SetDesc = "SHOTSOUND"
'    End With
'Next i
'---------------------------------------

DefHeight = Me.ScaleHeight + AddHeight
DefWidth = Me.ScaleWidth

End Sub

Public Sub Pause(Optional NoUpdate As Boolean)  'Optional UnPause As Boolean)

'Pauses/unpauses the game. More complicated than it sounds, because
'during the 'Ready...GO!' display, the user could press pause again

If Not GameOver And Not BadData Then
'    If Timer1.Enabled = True Then
'        Timer1.Enabled = False
'    Else
'        Timer1.Enabled = True
'    End If
    'If Timer1.Enabled = False Then
    
    If IsPaused Then
        Me.MousePointer = 99
        If Not UnPausing Then
            UnPausing = True
            lblPaused.Visible = True
            lblPaused.Caption = SBrack & "Ready..." & EBrack
            Call Wait(1)
            
            'Checks if user hasn't pressed pause again
            If UnPausing Then
                lblPaused.Caption = SBrack & "GO!" & EBrack
                Call Wait(1)
            End If
        Else
            UnPausing = False
        End If
    Else
        Me.MousePointer = 0
    End If
    
    'Timer1.Enabled = Not Timer1.Enabled
    
    If Not UnPausing Then
        IsPaused = False
    End If
    
    If Not NoUpdate Then
        Call PauseUpdate
    End If
    
    IsPaused = (Not IsPaused)
    UnPausing = False
    
End If

End Sub

Public Sub PauseUpdate()

'If Not BadData Then
    'Updates the display when pausing/unpausing
    'If Timer1.Enabled = False Then
    If Not IsPaused Then
        lblPaused.Caption = SBrack & "Paused" & EBrack
        lblPaused.Visible = True
        mnuPause.Caption = "Un&pause"
    Else
        'Call Wait(1)
        'Checks if the game is still paused after the 1 second wait
        'If Timer1.Enabled = True Then
        'If Not IsPaused Then
            lblPaused.Visible = False
            mnuPause.Caption = "&Pause"
        'End If
    End If
'End If
    
End Sub

Public Sub NewGame()

'If Timer1.Enabled Then
If IsPaused Then
    Call Pause
End If

Call Reset

Unload frmGameType

If Not BadData Then

    frmGameType.Show vbModal
    
    frmStats.Hide


    If GameType = 0 Or GameType = 1 Then
        Call ShowEnterName(0)
    End If
    
    If GameType = 0 Or GameType = 2 Then
        Call ShowEnterName(1)
    End If
    
    frmStats.Show
    Call GetReady

End If


'This is here in case BadData is true, and the Loadpics procedure
'doesn't run
Me.Show

End Sub

Sub GetReady()

frmChoseShots.Show vbModal

'If Not Timer1.Enabled Then
If IsPaused Then
    Call Pause
End If

Call MainLoop

End Sub

Public Sub ShowEnterName(Player As Byte)

ProcessPlayer = Player
'We want players 1 and 2, not 0 and 1!, hence the Player + 1
frmEnterName!txtEnterName.Text = "Player " & Player + 1
frmEnterName.Show vbModal

frmStats!lblTeamName(Player) = Face(ProcessPlayer).Name
    
End Sub

Public Sub ReadSettings()

Dim Temp As String
Dim TestLen As Integer
Dim KeyWords() As String
Dim Keyword As String
Dim i As Integer
Dim Count As Integer

Dim NewSet As String

Dim Free As Byte

Dim DoWhat As String
Free = FreeFile

'Open data file
Open SetFile For Input As #Free

Do Until EOF(Free)
    'Reads the current line of the file
    Input #Free, Temp
    
    'Checks to see where the space is.
    'A space is used to seperate the setting label from
    'the setting data e.g "PLAYERHP 100"
    Keyword = FindKeyWord(Temp)
    NewSet = FindNewSet(Temp)
    
    If Keyword <> "" Then
        'Checks if line is part of a list, or if its commented
        If Left(Keyword, 1) <> "'" Then
            If IsNumeric(Right(Keyword, 1)) = False Then
                If Keyword = "STARTSHOTS" Or Keyword = "STARTPOWERS" Or Keyword = "STARTCLASSES" Then
                    DoWhat = Keyword
                Else
                    Call LoadSets(BSettings, Keyword, NewSet)
                End If
            
            'These blocks set up things that go in lists,
            'such as the names of all the shots
            'e.g SHOT0,SHOT1,SHOT2...
            Else
                
                Dim SetListInd As Byte
                Dim ListName As String
                
                'ListName is the list label, e.g "SHOT"
                ListName = UCase(Left(Keyword, Len(Keyword) - FindStepRight(Keyword)))
                
                'SetListInd is the listindex
                SetListInd = Right(Keyword, FindStepRight(Keyword))
                
                Select Case DoWhat
                    Case "STARTSHOTS"
                        Call SetUpListSets(ShotSets, SetListInd)
                        Call LoadSets(ShotSets(SetListInd).Setting, ListName, NewSet)
                        
                    Case "STARTPOWERS"
                        Call SetUpListSets(PowerUps, SetListInd)
                        Call LoadSets(PowerUps(SetListInd).Setting, ListName, NewSet)
                    Case "STARTCLASSES"
                        Call SetUpListSets(Classes, SetListInd)
                        Call LoadSets(Classes(SetListInd).Setting, ListName, NewSet)
                        
                End Select
                        
            End If
        End If
    End If
Loop

On Error GoTo 0
'Close data file
Close #Free
    
'Sets the maximum number of powerups
ReDim Preserve PowerUpNow(1 To BSettings(SetInd("POWERUPMAX", BSettings)).SettingData)


For i = LBound(Face) To UBound(Face)
    Call ChDegrees(i, Face(i).Degrees)
Next i

End Sub

Sub LoadPics()

'Shows the loading form
frmLoading.Show

Call frmSkins.GetSprites

'Loads all the sprites into frmpics
'and creates the masks at runtime
Dim i As Integer

On Error GoTo NoBitmap
With frmPics

    Call LoadFacePics
    
    Call LoadShotPics
        
    Call LoadPowerPics
    
    'Loads explosion (kaboom) pictures
    'More pictures can be added easily without changing the code
    Dim PicFile As String
    'Dim Pic
    For i = 0 To 20
        PicFile = "pics\kaboom" & i & ".bmp"
        If FileExists(PicFile) Then
            KaboomNumbers = i
            If i > frmPics!picKaboom.UBound Then
                Load .picKaboom(i)
                Load .picKaboomMask(i)
                .picKaboom(i).Visible = True
                .picKaboomMask(i).Visible = True
                If i >= 1 Then
                    .picKaboom(i).Top = .picKaboom(i - 1).Top + .picKaboom(i - 1).Height
                End If
            End If
            .picKaboom(i).Picture = LoadPicture(PicFile)
            Call MakeMask(.picKaboom(i), .picKaboomMask(i))
            '.picKaboomMask(i).Picture = LoadPicture("pics\kaboommask" & i & ".bmp")
            KaboomNumbers = i
        End If
        frmLoading.SbarLoad.Value = i * 100 / 20
    Next i
    
    For i = 0 To 10
    
        PicFile = "pics/trail" & i & ".bmp"
        If FileExists(PicFile) Then
            If i > .picTrail.UBound Then
                Load .picTrail(i)
                Load .picTrailMask(i)
            End If
            
            .picTrail(i).Picture = LoadPicture(PicFile)
            Call MakeMask(.picTrail(i), .picTrailMask(i))
        Else
            Exit For
        End If
    Next i
    
    'frmPics!picShotMask.Picture = LoadPicture("pics\redshotmask.bmp")
    
    'A black bmp used to cover up pictures drawn using bltbit
    '.picBG.Picture = LoadPicture("pics\bground.bmp")
    Call MakeBG

End With

On Error GoTo 0

Unload frmLoading
Me.Show


Exit Sub

NoBitmap:

Call SetErr(Err.Description, 1)

End Sub

Sub LoadFacePics(Optional Skinning As Boolean)

Dim i As Integer

If Not Skinning Then
    On Error GoTo NoPic
Else
    On Error GoTo NoSkin
End If

With frmPics
    For i = 0 To 315 Step 45
        '.picFace(i).Picture = LoadPicture("pics\face" & i & ".bmp")
        .picFace(i).Picture = LoadPicture(FaceSprites(0) & i & ".bmp")
        'frmPics!picFaceMask(i).Picture = LoadPicture("pics\face" & i & "mask.bmp")
        
        '.picFace2(i).Picture = LoadPicture("pics\f2ce" & i & ".bmp")
        .picFace2(i).Picture = LoadPicture(FaceSprites(1) & i & ".bmp")
        'frmPics!picFace2Mask(i).Picture = LoadPicture("pics\f2ce" & i & "mask.bmp")
        
        Call MakeMask(.picFace(i), .picFaceMask(i))
        Call MakeMask(.picFace2(i), .picFace2Mask(i))
        
        frmLoading.SbarLoad.Value = i * 100 / 315
        
    Next i
        
End With


Exit Sub

NoPic:
Call SetErr(Err.Description, 1)
Exit Sub

NoSkin:
Call frmSkins.BadSkin

End Sub

Sub LoadShotPics(Optional Skinning As Boolean)

Dim i As Byte

If Not Skinning Then
    On Error GoTo NoPic
Else
    On Error GoTo NoSkin
End If

With frmPics
    'Loads the shot pictures
    For i = LBound(ShotSets) To UBound(ShotSets)
        If i > frmPics!picShot.UBound Then
            Load .picShot(i)
            Load .picShotMask(i)
            .picShot(i).Top = .picShot(i - 1).Top + .picShot(i - 1).Height
            .picShotMask(i).Top = .picShotMask(i - 1).Top + .picShotMask(i - 1).Height
            .picShot(i).Visible = True
            .picShotMask(i).Visible = True
        End If
        '.picShot(i).Picture = LoadPicture("pics\" & ShotSets(i).Setting(SetInd("SHOT", ShotSets(i).Setting)).SettingData & ".bmp")
        '.picShot(i).Picture = LoadPicture("pics\" & ShotSets(i).Setting(SetInd("NAME", ShotSets(i).Setting)).SettingData & ".bmp")
        .picShot(i).Picture = LoadPicture(ShotSprites & "\" & ShotSets(i).Setting(SetInd("NAME", ShotSets(i).Setting)).SettingData & ".bmp")
    
        Call MakeMask(.picShot(i), .picShotMask(i))
        
        frmLoading.SbarLoad.Value = i * 100 / UBound(ShotSets)
        
    Next i
End With


Exit Sub

NoPic:
Call SetErr(Err.Description, 1)
Exit Sub

NoSkin:
Call frmSkins.BadSkin

End Sub

Sub LoadPowerPics(Optional Skinning As Boolean)

Dim i As Byte

If Not Skinning Then
    On Error GoTo NoPic
Else
    On Error GoTo NoSkin
End If

With frmPics
    'Loads the Powerup pictures
    Dim PicName As String
    For i = LBound(PowerUps) To UBound(PowerUps)
    
        If i > frmPics!picPowerUp.UBound Then
            Load .picPowerUp(i)
            Load .picPowerUpMask(i)
            .picPowerUp(i).Top = .picPowerUp(i - 1).Top + .picPowerUp(i - 1).Height
            .picPowerUpMask(i).Top = .picPowerUpMask(i - 1).Top + .picPowerUpMask(i - 1).Height
            .picPowerUp(i).Visible = True
            .picPowerUpMask(i).Visible = True
        End If
        
        PicName = PowerUps(i).Setting(SetInd("PNAME", PowerUps(i).Setting)).SettingData
        '.picPowerUp(i).Picture = LoadPicture("pics\" & PicName & ".bmp")
        .picPowerUp(i).Picture = LoadPicture(PowerSprites & "\" & PicName & ".bmp")
        Call MakeMask(.picPowerUp(i), .picPowerUpMask(i))
        
        'frmPics!picPowerUpMask(i).Picture = LoadPicture("pics\" & PicName & "mask.bmp")
        frmLoading.SbarLoad.Value = i * 100 / UBound(PowerUps)
    Next i
End With


Exit Sub

NoPic:
Call SetErr(Err.Description, 1)
Exit Sub

NoSkin:
Call frmSkins.BadSkin

End Sub

Sub MakeBG()

'Makes the 'cover up' picture according to the colour of the background
Dim X As Integer
Dim Y As Integer

With frmPics.picBG
    For X = 0 To .Width
        For Y = 0 To .Height
             frmPics.picBG.PSet (X, Y), Me.BackColor
        Next Y
    Next X
End With
            
End Sub

Sub ChColours(WColour As String)

'Changes the background colours

Dim i As Byte
Me.BackColor = WColour

Call MakeBG

For i = picWinWep.LBound To picWinWep.UBound
    picWinWep(i).BackColor = WColour
Next i
    
End Sub

Sub MakeMask(PicBox As Object, MaskBox As Object)

Dim X As Integer
Dim Y As Integer
Dim ThePoint As Long

With PicBox

    MaskBox.Picture = .Picture

    For X = 0 To .Width
        For Y = 0 To .Height
        
            ThePoint = GetPixel(.hdc, X, Y) '.Point(x, y)
            If ThePoint > (vbWhite + 2000) Or ThePoint < (vbWhite - 2000) Then
                'MaskBox.PSet (x, y), vbBlack
                SetPixelV MaskBox.hdc, X, Y, vbBlack
                DoEvents
            End If
        Next Y
    Next X

End With


End Sub

Public Sub Reset()

DoEvents 'We need this when the program loads, otherwise the system timer
        'doesn't start until the load event finishes, and hence 'randomize'
        'will not work properly. *Stupid VB*

Dim i As Integer
Dim n As Integer
Dim PickNum As Byte

ReDim ShotSets(0)
ReDim PowerUps(0)
ReDim Classes(0)
ReDim BSettings(0)

ReDim PressKeys(UBound(Face), UBound(Face(0).PKeys))
ReDim AlreadySent(UBound(PressKeys), UBound(PressKeys, 2))

'If Timer1.Enabled Then
If Not IsPaused Then
    Call Pause
End If

'Loads the settings from the setting file
Call ReadSettings

'Gets the keys
Call GetKeys

'Loads the pictures and creates masks for them
If FirstTime Then
    Call LoadPics
End If

'Sets up buffers and Direct Sound
Call SetUpSounds(Me.hWnd)

'Hides all the stuff that appears at the end
lblWinner.Visible = False
lblWins.Visible = False
lblInstruct.Visible = False

For i = picWinWep.LBound To picWinWep.UBound
    lblWinKill(i).Visible = False
    lblWinLose(i).Visible = False
    picWinWep(i).Visible = False
    picWinWep(i).Cls
Next i


'On Error GoTo ReadError

'Starting HP. Used in other subs as well
StartHP = BSettings(SetIndMain("PLAYERHP")).SettingData
PickNum = BSettings(SetIndMain("PICKSHOTS")).SettingData

Me.Cls


'Reset explosions and Shooting properties
For i = LBound(Kaboom) To UBound(Kaboom)
    Kaboom(i).Wait = 0
Next i

For i = LBound(Trails) To UBound(Trails)
    Trails(i).Wait = 0
Next i

For i = LBound(Shooting) To UBound(Shooting)
    Call ShotReset(i)
Next i

For i = LBound(Flowers) To UBound(Flowers)
    Flowers(i).Shots = 0
    Flowers(i).ShotSpeed = -1
Next i

For i = LBound(PowerUpNow) To UBound(PowerUpNow)
    Call PowerUpReset(i)
Next i

'Sets players to default status
For i = LBound(Face) To UBound(Face)
    With Face(i)
        .Height = 33
        .Width = 33
'        .Degrees = SomRand(7, 0) * 45 'Faces a random direction
'        .UpDown = ""
        '.HoldShot = False
        .ShotCurrent = 0
        .ShotWait = 0
        .HP = StartHP
        .PlusArmour.Bonus = 0
        .PlusArmour.TimeLeft = 0
        .PlusDam.Bonus = 0
        .PlusDam.TimeLeft = 0
        .X = SomRand(Me.ScaleWidth - .Width, .Width)
        .Y = SomRand(Me.ScaleHeight - .Height, .Height)
        .ClassArm = 0
        .ClassDam = 0
        .ClassHP = 1
        .ClassShotSize = 0
        .ClassShotSpeed = 1
        .ClassSpeed = 1
        .ClassShotAccel = 0
        
        Call ChDegrees(i, SomRand(7, 0) * 45)
        
        If PlayType = 1 Then
            ReDim .Weps(PickNum - 1)
        End If
        
        frmStats.SbarHP(i).Max = StartHP * .ClassHP
        frmStats.lblClassType(i).Visible = False
            
    End With
    
    Call RandPlace(i)
    'DegreeChange(i) = 0
    
    'Update HP display
    Call ShowHP(i)

    'Update bonus display
    Call ShowMods(i)
    
    Call ShowModTime(i)
    

Next i


If GameType = 1 Then
    Call SrvSendHP
End If

Hor = Abs(BSettings(SetIndMain("MOVESTRAIGHT")).SettingData)
'Diag = BSettings(SetIndMain("MOVEDIAG")).SettingData

'Use pythagorus to find what the horizontal and vertical movement
'will be when moving one unit diagonally
Diag = Sqr(Hor ^ 2 / 2)

Call DrawFaces

On Error GoTo 0

'Updates power up stuff
Call SetUpPower

GameOver = False

'Timer1.Interval = Val(BSettings(SetInd("GAMESPEED", BSettings)).SettingData)

'Timer1.Enabled = False
'Call PauseUpdate
Me.Refresh

FirstTime = False

Exit Sub

'Displays the settings error message
ReadError:
Call SetErr(Err.Description)

End Sub

Public Sub RandPlace(ByVal WhichFace As Byte)

Call PlaySound("sounds/spawn.wav", 3)
With Face(WhichFace)
    .X = SomRand(Me.ScaleWidth - .Width, .Width)
    .Y = SomRand(Me.ScaleHeight - .Height, .Height)
End With

Call SrvSendPOS(WhichFace)

End Sub

Public Sub SetUpPower()

Dim i As Integer
PowerTotal = 0

For i = LBound(PowerUps) To UBound(PowerUps)
    With PowerUps(i)
        PowerTotal = PowerTotal + .Setting(SetInd("CHANCE", .Setting)).SettingData
    End With
Next i


End Sub
Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'Using GetKeyState now

'Dim WFace As Integer
'
'WFace = WhoseKey(KeyCode)
'
''Checks whether the key pressed belongs to a player
'If WFace >= 0 Then
'
'    'Checks whether the key pressed belongs to an AI player,
'    'so humans can't cheat and mess up the AI by pressing the
'    'other players keys :-)
'    If Not IsAI(WFace) Then
'        Call DoKeys(KeyCode)
'    End If
'
'ElseIf KeyCode = Asc("P") Then
'
'    Call Pause
'
'End If

If KeyCode = Asc("P") Then
    Call Pause
End If

End Sub


Sub DoKeys(Code As Integer, ByVal WFace As Byte, Optional RecievedData As Boolean)

'NOT USED ANYMORE. USING GETKEYSTATES NOW

'Dim SendIt As Boolean
'Dim MyKeys As Boolean
''Moving around using keyboard
'
''Note that AI uses this same event to move around,
''so in theory, there is nothing it can do that
''a human can't
'
''If Timer1.Enabled = True Then
'If Not IsPaused Then
'
'    'For i = 0 To UBound(Face)
'
'        MyKeys = (GameType = 0 Or GameType - 1 = WFace) Or RecievedData
'        SendIt = ((GameType = 2 And WFace = 1) And Not RecievedData)
'
'        If MyKeys Then
'            With Face(WFace)
'                Select Case Code
'                    Case .PKeys(2)
'                        'Turn left
'                        If Not SendIt Then
'                            DegreeChange(WFace) = -45
'                        Else
'                            Call SendDegMove(False)
'                        End If
'
'                    Case .PKeys(3)
'                        'Turn right
'                        If Not SendIt Then
'                            DegreeChange(WFace) = 45
'                        Else
'                            Call SendDegMove(True)
'                        End If
'
'                    Case .PKeys(0)
'                        'Forward
'                        If Not SendIt Then
'                            .UpDown = "up"
'                        Else
'                            Call SendUpDown(False)
'                        End If
'
'                    Case .PKeys(1)
'                        'Backward
'                        If Not SendIt Then
'                            .UpDown = "down"
'                        Else
'                            Call SendUpDown(True)
'                        End If
'
'                    Case .PKeys(4)
'                        'Shoot
'                        If Not SendIt Then
'                            .HoldShot = True
'                        Else
'                            Call SendShoot
'                        End If
'
'                    'Change weapon
'                    Case .PKeys(5)
'                        Call ShotChange(WFace, False)
'
'                    Case .PKeys(6)
'                        Call ShotChange(WFace, True)
'
'                End Select
'
'
'            End With
'        End If
'
'    'Next i
'
'End If

End Sub

Sub DoKeyStates()

Dim i As Byte
Dim n As Byte
Dim YesItIs As Boolean

For n = 0 To UBound(Face)
    If Not IsAI(n) Then
    
        With Face(n)
        
            For i = LBound(.PKeys) To UBound(.PKeys)
                YesItIs = KeyIsDown(.PKeys(i))
                PressKeys(n, i) = YesItIs
                
                'For network play. NOT DONE YET!
                If GameType = 2 Then
                    With frmGameType.LotsaSocks
                        If YesItIs And Not AlreadySent(n, i) Then
                            .SendStuff ("KEY " & i)
                            AlreadySent(n, i) = True
                        ElseIf Not YesItIs And AlreadySent(n, i) Then
                            .SendStuff ("KEYUP " & i)
                            AlreadySent(n, i) = False
                        End If
                    End With
                End If
                    
            Next i
            
        End With
        
    End If
Next n

End Sub

Function KeyIsDown(Ascii As Integer) As Boolean

If GetKeyState(Ascii) < -125 Then
    KeyIsDown = True
End If

End Function

Sub DoKeyUp(Code As Integer, ByVal WFace As Byte)

'Using getkeystates now
'-----------------------

'Because all movement is on the timer,
'the program needs to know when a player
'releases a key so it can stop moving the
'thingo's on screen

'Dim SendIt As Boolean

    
'    SendIt = (GameType = 2 And WFace = 1)
'    With Face(WFace)
'        If Code = .PKeys(2) Or Code = .PKeys(3) Then
'            If Not SendIt Then
'                WFace = 0
'            Else
'                Call SendStopTurn
'            End If
'        End If
'        If Code = .PKeys(0) Or Code = .PKeys(1) Then
'            If Not SendIt Then
'                .UpDown = ""
'            Else
'                Call SendStopMove
'            End If
'        End If
'        If Code = .PKeys(4) Then
'            If Not SendIt Then
'                .HoldShot = False
'            Else
'                Call SendStopShoot
'            End If
'        End If
'    End With

End Sub

Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Call DoKeyUp(KeyCode)

End Sub


Private Sub Form_Deactivate()

If Not GameOver Then
    'Updated in v3.4.0. Used to be timer1.enabled=false and call pauseupdate
    'If Timer1.Enabled Then
    If Not IsPaused Then
        Call Pause
    End If
End If

End Sub

Private Sub Form_LostFocus()

If Not GameOver Then
    'Updated in v3.4.0. Used to be timer1.enabled=false and call pauseupdate
    'If Timer1.Enabled Then
    If Not IsPaused Then
        Call Pause
    End If
End If

End Sub

Private Sub Form_Resize()


Call CenterIt(lblPaused)
Call CenterIt(lblWinner)
Call CenterIt(lblWins)
Call CenterIt(lblInstruct)

lblFrameRate.Top = frmPositions.ScaleHeight - lblFrameRate.Height

End Sub

Sub CenterIt(ByRef What As Object)

What.Left = Me.ScaleWidth / 2 - What.Width / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)

If Not NoRestart Then
    
    Dim Reply As Integer
    Reply = MsgBox("Quit now?", vbYesNo, "Quit")
    
    If Reply = vbNo Then
        Cancel = 2
    Else
        Call StopSounds
        Unload Me
        Unload frmPics
        Unload frmLoading
        Unload frmOptions
        Unload frmStats
        Unload frmEnterName
        Unload frmGameType
        End
    End If
Else
    Cancel = 2
End If

End Sub

Public Sub MoveFace(ByVal MoveDirection As String, ByVal Degrees As Integer, ByVal Index As Byte, ByRef XVal As Single, ByRef YVal As Single, ByVal Height As Integer, ByVal Width As Integer, Optional ByVal MoveFace As Boolean, Optional ByVal Units As Single = 1)

Dim MoveRight As Single
Dim MoveDown As Single
Dim Kaboom As Boolean
Dim i As Integer

Dim NewPlaceX As Integer
Dim NewPlaceY As Integer

Dim Radius As Integer
Dim IntUnits As Integer
Dim DecUnits As Single
Dim OneMore As Byte
Dim Fraction As Single

IntUnits = Int(Units)
If Units <> IntUnits Then '1 Then
    DecUnits = Units - IntUnits
    OneMore = 1
End If

Fraction = 1

'This is in a loop so that fast shots don't miss. It moves the shot
'one 'unit' at a time and checks if it has collided with anything
For i = 1 To IntUnits + OneMore

    If i > IntUnits Then
        Fraction = DecUnits
    End If
    
    Call MoveHowMuch(MoveDirection, Degrees, MoveRight, MoveDown)
    MoveDown = MoveDown * Fraction
    MoveRight = MoveRight * Fraction
'    MoveDown = MoveHowMuch(MoveDirection, Degrees).Down * Fraction
'    MoveRight = MoveHowMuch(MoveDirection, Degrees).Right * Fraction
    
    
    Radius = Width / 2
    
    'Stops things from moving off the screen
    If (YVal - Radius + MoveDown < 0) Then
        MoveDown = -YVal + Radius
        Kaboom = True
    
    End If
    
    If (YVal + Radius + MoveDown > Me.ScaleHeight) Then
        MoveDown = Me.ScaleHeight - YVal - Radius
        Kaboom = True
    
    End If
    
    If (XVal - Radius + MoveRight < 0) Then
        MoveRight = -XVal + Radius
        Kaboom = True
    
    End If
    
    If (XVal + Radius + MoveRight > Me.ScaleWidth) Then
        MoveRight = Me.ScaleWidth - XVal - Radius
        Kaboom = True
    
    End If
    
    '--------------------------------------
    'Moves whatever to its new position.
    'Xval and Yval are passed by reference, so sub can move them here
    XVal = XVal + MoveRight
    YVal = YVal + MoveDown
    
    If Not MoveFace Then
    
        'Stops the loop if shot has gone
        If Shooting(Index).What < 0 Then
            Exit For
        End If

        'Detects collsions
        Call Collision(Index)

        If Kaboom = True Then
            'Explodes if shot has hit wall
            Call Hit(Index)
        End If
        
    Else
        Call SrvSendPOS(Index)
    End If
    
Next i

End Sub

Sub CoverStuff()

Dim CoverWidth As Integer
Dim FaceWidth As Integer
Dim i As Integer

CoverWidth = frmPics.picBG.Width / 2

'Cover up faces
For i = 0 To 1
    Call CoverUp(Face(i).X - CoverWidth, Face(i).Y - CoverWidth)
Next i

'Cover up shots
For i = LBound(Shooting) To UBound(Shooting)
    With Shooting(i)
        If .What >= 0 And .ShotX >= 0 And .ShotY >= 0 Then
            Call CoverUp(.ShotX - frmPics!picShot(.What).Width / 2, .ShotY - frmPics!picShot(.What).Height / 2)
        End If
    End With
Next i

End Sub

Sub DrawShots(Ind As Integer)

Dim PicInd As Integer
Dim PicWidth As Integer
Dim PicHeight As Integer

PicInd = Shooting(Ind).What 'ShotIndex(Shooting(ind).What, ShotSets)
PicWidth = frmPics.picShot(PicInd).Width
PicHeight = frmPics.picShot(PicInd).Height

With Shooting(Ind)
    'Draws the shots
    Call BitBlt(Me.hdc, .ShotX - PicWidth / 2, .ShotY - PicHeight / 2, PicWidth, PicHeight, frmPics!picShotMask(PicInd).hdc, 0, 0, vbMergePaint)
    Call BitBlt(Me.hdc, .ShotX - PicWidth / 2, .ShotY - PicHeight / 2, PicWidth, PicHeight, frmPics!picShot(PicInd).hdc, 0, 0, vbSrcAnd)
End With

End Sub

Sub DrawKaboom()

Dim i As Byte
Dim PicHeight As Integer
Dim PicWidth As Integer
Dim RanNum As Integer

'Draws the explosions *after* everything else, so they are always *on top*
For i = LBound(Kaboom) To UBound(Kaboom)
    If Kaboom(i).Wait > 0 Then
        
        RanNum = SomRand(KaboomNumbers, 0)
        PicHeight = frmPics!picKaboom(RanNum).Height
        PicWidth = frmPics!picKaboom(RanNum).Width

        Call BitBlt(Me.hdc, Kaboom(i).X - PicWidth / 2, Kaboom(i).Y - PicWidth / 2, PicWidth, PicHeight, frmPics!picKaboomMask(RanNum).hdc, 0, 0, vbMergePaint)
        Call BitBlt(Me.hdc, Kaboom(i).X - PicHeight / 2, Kaboom(i).Y - PicHeight / 2, PicWidth, PicHeight, frmPics!picKaboom(RanNum).hdc, 0, 0, vbSrcAnd)
    End If
Next i

End Sub

Sub UpdateKaboom()

Dim i As Byte
Dim CoverWidth As Integer

CoverWidth = frmPics.picBG.Width / 2

'Cover up explosions and update time remaining before
'the explosion graphics disappear
For i = LBound(Kaboom) To UBound(Kaboom)
    If Kaboom(i).Wait > 0 Then
        'If DoProcess Then
            Kaboom(i).Wait = Kaboom(i).Wait - 1
        'End If
        Call CoverUp(Kaboom(i).X - CoverWidth, Kaboom(i).Y - CoverWidth)
    End If
Next i

End Sub

Sub UpdateTrails()

Dim i As Integer
Dim CoverWidth As Integer

CoverWidth = frmPics.picBG.Width / 2

'Cover up trails and update time remaining before they disappear
For i = LBound(Trails) To UBound(Trails)
    If Trails(i).Wait > 0 Then
        'If DoProcess Then
            Trails(i).Wait = Trails(i).Wait - 1
        'End If
        Call CoverUp(Trails(i).X - CoverWidth, Trails(i).Y - CoverWidth)
    End If
Next i

End Sub
Sub DrawFaces()

'Draws the 'Faces'
Dim WhichPic As Object
Dim WhichMask As Object
Dim FWidth As Integer
Dim FHeight As Integer
Dim i As Byte

For i = 0 To 1
    If i = 1 Then
        Set WhichPic = frmPics!picFace2
        Set WhichMask = frmPics!picFace2Mask
    Else
        Set WhichPic = frmPics!picFace
        Set WhichMask = frmPics!picFaceMask
    End If
    
    FWidth = WhichPic(0).Width
    FHeight = WhichPic(0).Height
 
    Call BitBlt(Me.hdc, Face(i).X - FWidth / 2, Face(i).Y - FHeight / 2, FWidth, FHeight, WhichMask(Face(i).Degrees).hdc, 0, 0, vbMergePaint)
    Call BitBlt(Me.hdc, Face(i).X - FWidth / 2, Face(i).Y - FHeight / 2, FWidth, FHeight, WhichPic(Face(i).Degrees).hdc, 0, 0, vbSrcAnd)
Next i


End Sub

Sub DrawTrails()

Dim i As Integer
Dim TestNum As Single
Dim Width As Integer

Width = frmPics.picTrail(0).Width

For i = LBound(Trails) To UBound(Trails)
    With Trails(i)
        If .Wait > 0 Then
            'If DoProcess Then
                TestNum = .Wait - 1 '(.Wait / 2) - 1
            'End If
'            If Int(TestNum) = TestNum Then
                Call BitBlt(Me.hdc, .X - Width / 2, .Y - Width / 2, Width, Width, frmPics.picTrailMask(TestNum).hdc, 0, 0, vbMergePaint)
                Call BitBlt(Me.hdc, .X - Width / 2, .Y - Width / 2, Width, Width, frmPics.picTrail(TestNum).hdc, 0, 0, vbSrcAnd)
'            End If
        End If
    End With
Next i
                

End Sub

Sub DrawPowers()

Dim i As Integer
Dim PicInd As Integer
Dim PicWidth As Integer
Dim PicHeight As Integer

'Draws powerups (always on the bottom, so drawn first)
For i = LBound(PowerUpNow) To UBound(PowerUpNow)
    With PowerUpNow(i)
        'If .What <> "" Then
        If .What >= 0 Then
            PicInd = .What 'ShotIndex(.What, PowerUps)
            PicWidth = frmPics!picPowerUp(PicInd).Width
            PicHeight = frmPics!picPowerUp(PicInd).Height
            '(PicHeight / 2)- (PicWidth / 2)
            Call BitBlt(Me.hdc, .X - (PicHeight / 2), .Y - (PicWidth / 2), PicWidth, PicHeight, frmPics!picPowerUpMask(PicInd).hdc, 0, 0, vbMergePaint)
            Call BitBlt(Me.hdc, .X - (PicHeight / 2), .Y - (PicWidth / 2), PicWidth, PicHeight, frmPics!picPowerUp(PicInd).hdc, 0, 0, vbSrcAnd)
        End If
    End With
Next i

End Sub

Public Sub Collision(ByVal SInd As Byte)

'Detects collisions between a shot and a player

Dim CenterX As Integer
Dim CenterY As Integer

'You cannot collide with your own shot,
'so the opponent's index number must be found
Dim OpNum As Byte
Dim YourNum As Byte

YourNum = Shooting(SInd).WhichFace

If YourNum = 0 Then
    OpNum = 1
Else
    OpNum = 0
End If

'Finds the center of the opponent
CenterX = Face(OpNum).X
CenterY = Face(OpNum).Y

'Because both faces are **circles**, we can simply check whether the
'distance between the shot and the face is less than or equal
'to the combined radius of the two circles, using the distance formula.
If Distance(Shooting(SInd).ShotX, CenterX, Shooting(SInd).ShotY, CenterY) < Face(OpNum).Width / 2 + ShotData(Shooting(SInd).What, "RADIUS") + Face(YourNum).ClassShotSize Then
    
    'Wont calculate damage if it is 0-0
    If ShotData(Shooting(SInd).What, "DAM") > 0 Then
        Call Ouch(OpNum, ShotDamage(Shooting(SInd).What), Shooting(SInd).What) 'ShotIndex(Shooting(SInd).What, ShotSets))
    End If
    Call Hit(SInd, True)
    
End If

End Sub

Sub Ouch(ByVal Defender As Byte, ByVal Damage As Integer, ByVal WepInd As Integer)

'This is the sub where all the shot damage is dealt. Damage from
'powerups is done in the 'GotPowerUps' sub

Dim ModDam As Integer
Dim Armour As Integer
Dim DamMod As Integer
Dim ArmMod As Integer

With Face(Defender)

    'Sets the last thing that hurt the player. Used for that '[Player1] used
    '[weapon] to kill [player2] display at the end.
    .LastHurt = WepInd
    
    'Calculates the damage after bonuses
    ModDam = Damage * (Face(FindOp(Defender)).PlusDam.Bonus + 100) / 100
    
    'Stops 'negative damage'. Remember that two negatives cancel out
    'so without this, your shot would *add* to your opponent's health
    'if you damage mod is < -100%
    If ModDam < 0 Then
        ModDam = 0
    End If
    
    Armour = .PlusArmour.Bonus
    
    'Ensures you can't do 'negative damage' if defender has
    'more armour than damage dealt
    If Armour > ModDam Then
        Armour = ModDam
    End If
    
    'Deals the damage
    .HP = .HP - GRound(ModDam) + Armour
    
    'Finds the damage and armour mod effect of the shot and gives a random
    'number between the min and max effects
    DamMod = SomRand(ShotData(WepInd, "DAMMODMAX"), ShotData(WepInd, "DAMMODMIN"))
    ArmMod = SomRand(ShotData(WepInd, "ARMMODMAX"), ShotData(WepInd, "ARMMODMIN"))
    
    'Subtracts the mod effects
    .PlusDam.TimeLeft = .PlusDam.TimeLeft - DamMod
    .PlusArmour.TimeLeft = .PlusArmour.TimeLeft - ArmMod
    
    'Update the HP display
    Call ShowHP(Defender)
    
    'I have to do this because of some stupid winsock 'feature' that
    'sends this *after* the restart. If a player is dead, the HP gets
    'sent with the 'Checkforwins' procedure. It's just the stupid way
    'the winsock control works in dos-based platforms like my win98.
    If .HP > 0 Then
        Call SrvSendHP
    End If
    
End With

End Sub

Sub ShowHP(ByVal WFace As Byte)

'Shows the HP in the stats display
With frmStats
    .lblHP(WFace).Caption = SBrack & Face(WFace).HP & EBrack
    .SbarHP(WFace).Value = Face(WFace).HP
        
    If Face(WFace).HP / StartHP * Face(WFace).ClassHP <= 0.2 Then
    
        Const LowHP = &HFF&
        
        'Turns red if player is almost dead
        .lblHP(WFace).ForeColor = LowHP
        .SbarHP(WFace).BarColour = LowHP
        
        'Turns the player's colour if over 100% health.
    ElseIf Face(WFace).HP > StartHP * Face(WFace).ClassHP Then
        .lblHP(WFace).ForeColor = .lblTeamName(WFace).ForeColor
        .SbarHP(WFace).BarColour = .lblTeamName(WFace).ForeColor
    Else
        .lblHP(WFace).ForeColor = &HD6E6E9
        .SbarHP(WFace).BarColour = &H808000
        
    End If
End With

End Sub

Public Sub GotPowerUp(ByVal PIndex As Byte, ByVal WhichFace As Byte)

'Checks to see if player has gotten a powerup, and if so
'performs the appropriate action and plays a sound.
'Remember that everything in the game is a circle, so checking for collisions
'involves the distance formula

'This sub is a good example of the downside of making a highly-customisable
'program. If, for instance, the user decides the put a string into the
''TIMEMIN' field, the whole thing will crash.
If Distance(Face(WhichFace).X, PowerUpNow(PIndex).X, Face(WhichFace).Y, PowerUpNow(PIndex).Y) <= Face(WhichFace).Width / 2 + BSettings(SetInd("POWERUPSIZE", BSettings)).SettingData Then

    Dim WhatUp As ShotSettings
    
    Dim Low As String
    Dim High As String
    Dim TMin As Long
    Dim TMax As Long
    
    'String operations are performed on this with flower powerups,
    'so it is a string
    Dim EffectYou As String
    
    Dim Effect As Integer
    Dim EffectTime As Long
    
    Dim Sound As String
    
    'Find the effect of the powerup
    WhatUp = PowerUps(PowerUpNow(PIndex).What)  'PowerUps(ShotIndex(PowerUpNow(PIndex).What, PowerUps))
    
    'Min and max effect
    Low = WhatUp.Setting(SetInd("LOW", WhatUp.Setting)).SettingData
    High = WhatUp.Setting(SetInd("HIGH", WhatUp.Setting)).SettingData
    
    'Minimum and maximum time that the powerup has effect for
    TMin = WhatUp.Setting(SetInd("TIMEMIN", WhatUp.Setting)).SettingData
    TMax = WhatUp.Setting(SetInd("TIMEMAX", WhatUp.Setting)).SettingData
    
    'Find the sound of the shot
    Sound = WhatUp.Setting(SetInd("SOUND", WhatUp.Setting)).SettingData
    
    'Whether the powerup effects the player who gets it or the opponent
    'The 'Weaken' powerup, for example, is like the 'DamPlus' powerup with
    'this set to 0.
    EffectYou = WhatUp.Setting(SetInd("YOU", WhatUp.Setting)).SettingData
    
    
    If IsNumeric(Low) And IsNumeric(High) Then
        Effect = SomRand(High, Low)
    End If
    
    EffectTime = SomRand(TMax, TMin)
    
    'Swaps faces if powerup is supposed to damage opponent
    If EffectYou = 0 Then
        WhichFace = FindOp(WhichFace)
    End If
    
    With Face(WhichFace)
    
        'Figures out what to do
        Select Case UCase(WhatUp.Setting(SetInd("EFFECT", WhatUp.Setting)).SettingData)
        
            Case "PLUSDAM"
                
                'Extra damage/weaken powerups.
                .PlusDam.TimeLeft = .PlusDam.TimeLeft + EffectTime

            Case "HEAL"
                .HP = .HP + Effect
                
                If EffectYou = 0 Then
                    
                    'The last thing that hurt the player was a powerup
                    Face(WhichFace).LastHurt = PowerUpNow(PIndex).What * -1 'ShotIndex(PowerUpNow(PIndex).What, PowerUps) * -1
                    
                    Call DoKaboom(.X, .Y, True, -1)
                    'Sound = "die.wav"
                End If
                
                'Update the HP display
                Call ShowHP(WhichFace)
                Call SrvSendHP
                
            Case "ARMOUR"
                
                'Armour/anti-armour powerups
                .PlusArmour.TimeLeft = .PlusArmour.TimeLeft + EffectTime
                
            Case "TIME"
                
                ''Extra time' powerups. Affect both armour and damage mods
                .PlusArmour.TimeLeft = .PlusArmour.TimeLeft + Effect
                .PlusDam.TimeLeft = .PlusDam.TimeLeft + Effect

                
            Case "SHOOT"
                
                ''Flowers' are positions on the game field where a ring of shots
                'will shoot out. After a set number of shots, they stop shooting
                
                Dim i As Integer
                'Finds an empty flower record
                For i = LBound(Flowers) To UBound(Flowers)
                    If Flowers(i).Shots = 0 Then
                        
                        'Sets the 'owner' of the flower
                        Flowers(i).Player = WhichFace
                        
                        'Number of times to shoot
                        Flowers(i).Shots = EffectTime
                        
                        'Uses the 'Low' field to store the type of shot the flower
                        'will use
                        Flowers(i).ShotType = Low
                        
                        'Uses the 'High' field to store the time between each shot
                        Flowers(i).PauseMax = High
                        
                        'Resets the pause time
                        Flowers(i).Pause = 0
                        
                        'Sets the original x and y positions of the flower.
                        'This will get changed if
                        Flowers(i).X = PowerUpNow(PIndex).X
                        
                        Flowers(i).Y = PowerUpNow(PIndex).Y
                        
                        'Uses the first character of the effectyou field
                        'This is where the flower will shoot from
                        'Either where the powerup was (0), wherever the
                        'player is (1), or random (2).
                        Flowers(i).FollowYou = Left(EffectYou, 1)
                        
                        'Uses the second character of the effectyou setting to set the new speed
                        'of the flower shots. If the field is empty,
                        'it will use the shot's original speed. See 'shootnow'
                        'for the details
                        If Len(EffectYou) > 1 Then
                            Flowers(i).ShotSpeed = Right(EffectYou, 1)
                        Else
                            Flowers(i).ShotSpeed = -1
                        End If
                        Exit For
                        
                    End If
                Next i
            
        End Select
    End With
    
    If Sound <> "" Then
        'Plays the sound
        Call PlaySound(App.Path & "\sounds\" & Sound, 4)
    End If
    
    Call PowerUpReset(PIndex)
'    'Update the bonus display on the stats form
'    Call ShowMods(WhichFace)
    
    'Covers up the image. This only needs to be done once because
    'powerups don't move
    Call CoverPowers(PIndex)
End If

End Sub

Sub CoverPowers(ByVal Ind As Byte)

'Covers the powerups when picked up
Call CoverUp(PowerUpNow(Ind).X - frmPics!picPowerUp(0).Width, PowerUpNow(Ind).Y - frmPics!picPowerUp(0).Height)

End Sub


Sub PowerUpReset(ByVal Ind As Integer)

'Reset the powerup
PowerUpNow(Ind).What = -1
PowerUpNow(Ind).GoForChance(0) = -1
PowerUpNow(Ind).GoForChance(1) = -1

End Sub

Sub ShowMods(ByVal WFace As Byte)

'Display bonus mods in the stats window
'This is now done on every timer loop because it is constantly fading
'Note that this sub updates the bonuses themselves as well as
'showing the update

'Font sizes. LrgSize is the size of the display when the mod is affected
'by a powerup. SmlSize is when the mod is normal
Const LrgSize = 16
Const SmlSize = 13

'The number of units of time per unit of the mod. e.g, plusarmour.timeleft =300
'would give an armour bonus of 2
Const ArmBonus = 150
Const DamBonus = 25

Dim Bonus As Integer
Dim TheCaption As String

With frmStats
    
    'Note: the main reason why there the armour and damage mods have a .bonus and
    '.timeleft property is because they used to work differently. i.e the mod
    'used to stay fixed until the time ran out. Now, the strength of the mods is
    'calculated directly from the time remaining. However, i decided to keep the
    '.bonus property because it is used elsewhere in the program, such as with
    'the AI functions.
    
    'Updates the armour bonus
    Face(WFace).PlusArmour.Bonus = Int(Face(WFace).PlusArmour.TimeLeft / ArmBonus) + Face(WFace).ClassArm
    
    TheCaption = SBrack & Face(WFace).PlusArmour.Bonus & EBrack
    
    If .lblPlusArmour(WFace).Caption <> TheCaption Then
        .lblPlusArmour(WFace).Caption = TheCaption
    End If
    
    If Face(WFace).PlusArmour.TimeLeft <> 0 Then
        .lblPlusArmour(WFace).FontBold = True
        .lblPlusArmour(WFace).FontSize = LrgSize
    Else
        .lblPlusArmour(WFace).FontBold = False
        .lblPlusArmour(WFace).FontSize = SmlSize
    End If
    
    Face(WFace).PlusDam.Bonus = Int(Face(WFace).PlusDam.TimeLeft / DamBonus) + Face(WFace).ClassDam
    
    TheCaption = SBrack & Face(WFace).PlusDam.Bonus & "%" & EBrack
    
    If .lblPlusDam(WFace).Caption <> TheCaption Then
        .lblPlusDam(WFace).Caption = TheCaption
    End If
    
    If Face(WFace).PlusDam.TimeLeft <> 0 Then
        .lblPlusDam(WFace).FontBold = True
        .lblPlusDam(WFace).FontSize = LrgSize
    Else
        .lblPlusDam(WFace).FontBold = False
        .lblPlusDam(WFace).FontSize = SmlSize
    End If
    
End With


End Sub

Sub ShowModTime(ByVal WFace As Byte)

''NO LONGER USED
'
'Dim Temp As Long
'
'With frmStats
'
'    Temp = Face(WFace).PlusArmour.TimeLeft
'    If .lblPlusArmourTime(WFace).Caption <> Temp Then
'        .lblPlusArmourTime(WFace).Caption = Temp
'    End If
'
'    Temp = Face(WFace).PlusDam.TimeLeft
'    If .lblPlusDamTime(WFace).Caption <> Temp Then
'        .lblPlusDamTime(WFace).Caption = Temp
'    End If
'
'End With

End Sub

Sub DecayHP(ByVal WFace As Byte)

Dim ExtraHP As Boolean

With Face(WFace)

    ExtraHP = (.HP > StartHP * .ClassHP)

    If ExtraHP Then
        If .HPTime <= 0 Then
            .HP = .HP - 1
            .HPTime = 15
            Call ShowHP(WFace)
        End If
        .HPTime = .HPTime - 1
    End If
End With

End Sub

Public Sub Hit(ByVal Index As Integer, Optional HitPlayer As Boolean)
'Displays the explosion when a shot hits something,
'and also resets the shot

Dim i As Integer
Dim WFace As Byte

With Shooting(Index)
    Call DoKaboom(.ShotX, .ShotY, HitPlayer, .What)
    WFace = .WhichFace
End With

If Shooting(Index).What >= 0 Then
    'Calculates damage for shots that explode
    Dim ExArea As Integer
    
    With ShotSets(Shooting(Index).What) 'ShotSets(ShotIndex(Shooting(Index).What, ShotSets))
    
        ExArea = .Setting(SetInd("EXAREA", .Setting)).SettingData
    
        If ExArea > 0 Then
            'For i = LBound(Face) To UBound(Face)
            Dim Opp As Byte
            Opp = FindOp(WFace)
            If Distance(Face(Opp).X, Shooting(Index).ShotX, Face(Opp).Y, Shooting(Index).ShotY) <= ExArea + Face(0).Width / 2 Then
                Dim ExLow As Integer
                Dim ExHigh As Integer
                ExLow = .Setting(SetInd("EXLOW", .Setting)).SettingData
                ExHigh = .Setting(SetInd("EXHIGH", .Setting)).SettingData
                Call Ouch(Opp, SomRand(ExHigh, ExLow), Shooting(Index).What) 'ShotIndex(Shooting(Index).What, ShotSets))
            End If
            'Next i
        End If
        
    End With
    
    Call ShotReset(Index)
End If

End Sub

Sub ShotReset(ByVal Index As Integer)

'Resets the shot
With Shooting(Index)
'    .Degrees = 0
    .Expire = -1
'    .WhichFace = 0
'    .What = ""
    .What = -1
'    .ShotX = -10
'    .ShotY = -10
    .NewSpeed = -1
    .AddSpeed = 0

End With

End Sub

Sub DoKaboom(X As Single, Y As Single, HitPlayer As Boolean, SInd As Integer)

'Creates an explosion. The graphics are drawn on the timer event, not here

Dim i As Integer
Dim BoomInd As Integer
Dim SoundFile As String
Dim OtherSound As String

'See's if there's a 'kaboom' that is free
For i = LBound(Kaboom) To UBound(Kaboom)
    If Kaboom(i).Wait <= 0 Then
        Kaboom(i).Wait = Val(BSettings(SetIndMain("EXPLOSIONWAIT")).SettingData)
        Kaboom(i).X = X
        Kaboom(i).Y = Y
        BoomInd = i

        Exit For
    End If
Next i

'Plays sound
Dim RanNum As Byte
Dim Prior As Byte


If HitPlayer = False Then

    If SInd >= 0 Then
        OtherSound = ShotData(SInd, "HITWALLSND")
        
        Select Case UCase(OtherSound)
        
            Case ""
'                RanNum = SomRand(2, 1)
'                SoundFile = "kaboom" & RanNum & ".wav"
'                Prior = 0
                Call PlayRanKaboom("kaboom", 0)
            Case "NONE"
                'SoundFile = ""
            Case Else
                
                Call PlayRanKaboom(OtherSound, 0)
'                Dim RanThis As Integer
'                RanThis = -1
'
'                For i = 0 To 5
'                    If FileExists("sounds/" & OtherSound & i & ".wav") Then
'                        RanThis = i
'                    Else
'                        Exit For
'                    End If
'                Next i
'
'                If RanThis >= 0 Then
'                    SoundFile = OtherSound & SomRand(RanThis, 0) & ".wav"
'                Else
'                    SoundFile = ""
'                End If
        
        End Select
        
    End If
    
Else
'    RanNum = SomRand(2, 1)
'    SoundFile = "hit" & RanNum & ".wav"
'    Prior = 1
    Call PlayRanKaboom("hit", 1)
End If
        
'If SoundFile <> "" Then
'    Call PlaySound("sounds/" & SoundFile, Prior)
'End If

If GameType = 1 Then
    Call MoreData("BOOM" & " " & X & AComma & Y & AComma & HitPlayer)
End If


End Sub

Sub PlayRanKaboom(WavFile As String, Optional Priority As Byte)

'This players a random sound for an explosions. The sounds must be in
'the format sound0.wav, sound1.wav, sound2.wav etc
Dim RanThis As Integer
Dim SoundFile As String
Dim i As Byte
Dim AllDone As Boolean
RanThis = -1

Do
    If FileExists("sounds/" & WavFile & i & ".wav") Then
        RanThis = i
    Else
        AllDone = True
    End If
    i = i + 1
Loop Until AllDone Or i >= 255

If RanThis >= 0 Then
    SoundFile = WavFile & SomRand(RanThis, 0) & ".wav"
    Call PlaySound("sounds/" & SoundFile, Priority)
End If

End Sub

Public Sub Turning(ByVal TurnWhat As Byte, ByVal DegreeChange As Integer)

If DegreeChange <> 0 Then
    Face(TurnWhat).Degrees = Face(TurnWhat).Degrees + DegreeChange
    Call ChDegrees(TurnWhat, Face(TurnWhat).Degrees)
End If

End Sub

Public Sub ChDegrees(ByVal WhichFace As Integer, ByRef FaceDegrees As Integer)
'This sub rotates the images by loading the corresponding bmp file

Select Case FaceDegrees
    'If Degrees is 360 (revolution), then set it back to 0
    Case 360
        FaceDegrees = 0
    'If it's -45, set it to 315
    Case -45
        FaceDegrees = 315
End Select

Face(WhichFace).Degrees = FaceDegrees

Call MoreData("DEG" & WhichFace & " " & FaceDegrees)

End Sub

Public Sub ShotChange(ByVal WhichFace As Byte, Up As Boolean)

'Changes the shot type of the player. Called either when
'the player presses one of the change keys, or when the
'AI calls the sub
With Face(WhichFace)
    
    If Up Then
       .ShotCurrent = ConvShot(.ShotCurrent + 1, UBound(.Weps))
    Else
        .ShotCurrent = ConvShot(.ShotCurrent - 1, UBound(.Weps))
    End If
    
    If GameType <> 0 Then
        Call SendShot(.Weps(.ShotCurrent))
    End If

End With

Call ShowShotStuff(WhichFace)
    
End Sub

Sub ShowShotStuff(WFace As Byte, Optional ByVal ShotInd As Integer = -1)

Dim SInd As Integer
With Face(WFace)

    If ShotInd >= 0 Then
        SInd = ShotInd
    Else
        SInd = .Weps(.ShotCurrent)
    End If

    'Displays an image of the shots in the Stats window, after clearing the
    'picture box
     frmStats.picShowShot(WFace).Cls
     Call BitBlt(frmStats.picShowShot(WFace).hdc, 0, 0, 50, 50, frmPics.picShotMask(SInd).hdc, 0, 0, vbMergePaint)
     Call BitBlt(frmStats.picShowShot(WFace).hdc, 0, 0, 50, 50, frmPics.picShot(SInd).hdc, 0, 0, vbSrcAnd)
     
     frmStats!picShowShot(WFace).Refresh
End With

Dim ExArea As Integer
'With ShotSets(SInd)

    ExArea = ShotData(SInd, "EXAREA") '.Setting(SetInd("EXAREA", .Setting)).SettingData
    
    frmStats.lblCurShot(WFace).Caption = ShotData(SInd, "NAME") '.Setting(SetInd("NAME", .Setting)).SettingData
    frmStats.lblShotDam(WFace) = ShotData(SInd, "DAMLOW") & " - " & ShotData(SInd, "DAM") '.Setting(SetInd("DAM", .Setting)).SettingData
    frmStats.lblShotEx(WFace) = ShotData(SInd, "EXLOW") & " - " & ShotData(SInd, "EXHIGH") & " (" & ExArea & ")"

'End With


End Sub

Public Sub ShootNew(ByVal WhichFace As Byte, ByVal Degrees As Integer, Optional ShotNum As Integer = -1, Optional X As Integer = -1, Optional Y As Integer = -1, Optional PlaySnd As Boolean, Optional ByVal Speed As Integer = -1)
'This creates a new shot by setting up an empty 'Shooting' box
'The actual graphics are drawn in the Timer1 event


Dim IsFlower As Boolean
Dim UseThis As Integer

If ShotNum >= 0 Then
    UseThis = ShotNum
Else
    With Face(WhichFace)
        UseThis = .Weps(.ShotCurrent)
    End With
End If

If ShotNum >= 0 And X >= 0 And Y >= 0 Then
    IsFlower = True
End If

'Only runs if player has 'reloaded' **This is now
'checked in the timer loop, not here**
'If Face(WhichFace).ShotWait = 0 Or IsFlower Then

Dim i As Integer
'Dim n As Integer
Dim SInd As Byte
Dim ExpireH As String
Dim ExpireL As String
Dim Spread As Single

For i = LBound(Shooting) To UBound(Shooting)

    'Searches for an empty array number
    If Shooting(i).What < 0 Then 'Shooting(i).What = "" Then
        
        Shooting(i).What = UseThis
        
        If Not IsFlower Then
        
            'This happens with normal shooting
             'Face(WhichFace).Weps(Face(WhichFace).ShotCurrent) '.Setting(SetInd("NAME", .Setting)).SettingData
            Shooting(i).ShotX = Face(WhichFace).X
            Shooting(i).ShotY = Face(WhichFace).Y
            
            'Face(WhichFace).Weps(Face(WhichFace).ShotCurrent)
        Else
        
            'This happens with shooting powerups
            'ShotSets(ShotNum).Setting(SetInd("NAME", ShotSets(ShotNum).Setting)).SettingData
            Shooting(i).ShotX = X
            Shooting(i).ShotY = Y
            If Speed >= 0 Then
                Shooting(i).NewSpeed = Speed
            End If

            'WhatNum = ShotNum
            
        End If
                
        'Finds the max and min expiry.
        With ShotSets(UseThis)
            ExpireL = .Setting(SetInd("EXPIRE", .Setting)).SettingData
            ExpireH = .Setting(SetInd("EXPIREHIGH", .Setting)).SettingData
        End With
        
        'If the shot expires, the expiry time is set
        If IsNumeric(ExpireL) And IsNumeric(ExpireH) Then
            Shooting(i).Expire = SomRand(ExpireH, ExpireL)
        End If
        
        'Calculates spread of shots
        Spread = ShotData(UseThis, "SPREAD")
        Spread = SomRand(Spread, -Spread, 1)
        
        Shooting(i).Degrees = Degrees + Spread
        Shooting(i).WhichFace = WhichFace
        
        
        If Not IsFlower Or PlaySnd = True Then
            Dim Sound As String
            Dim ShotSound As String
            SInd = Shooting(i).What 'ShotIndex(Shooting(i).What, ShotSets)
            
            With ShotSets(SInd)
                ShotSound = .Setting(SetInd("SOUND", .Setting)).SettingData
            End With
            
            'Checks to see if shot is assigned a sound.
            'If not, uses the default sound
            'If Trim(ShotSound) <> "" Then
            If FileExists("sounds/" & ShotSound) And ShotSound <> "" Then
                Sound = ShotSound
            Else
                Sound = BSettings(SetInd("DEFSHOTSOUND", BSettings)).SettingData
            End If
            
            Call PlaySound("sounds/" & Sound, 4)
        End If
        
        
        If Not IsFlower Then
            'Normal shooting
            With ShotSets(SInd)
                Face(WhichFace).ShotWait = .Setting(SetInd("WAIT", .Setting)).SettingData
            End With
        End If

                
        Exit For
        
    End If
    
Next i
    
'End If

End Sub

Sub FlowerShoot()

'Makes the 'flowers' shoot
Dim i As Integer
Dim n As Integer

Dim Snd As Boolean
For i = LBound(Flowers) To UBound(Flowers)

    With Flowers(i)
    
        'De-increments the pause time
        .Pause = .Pause - 1
        
        If .Pause <= 0 Then
            .Pause = .PauseMax
            
            If .Shots > 0 Then
                
                'If the flower follows the player, then
                'the Coordinates are moved to where the
                'player is.
                Select Case .FollowYou
                    Case 1
                        .X = Face(.Player).X
                        .Y = Face(.Player).Y
                    Case 2
                        Const Bound = 10
                        .X = SomRand(Me.ScaleWidth - Bound, Bound)
                        .Y = SomRand(Me.ScaleHeight - Bound, Bound)
                End Select
                                
                'De-increments the number of shots remaining
                .Shots = .Shots - 1
                
                For n = 0 To 315 Step 45
                    Call ShootNew(.Player, n, .ShotType, .X, .Y, , .ShotSpeed)
                Next n
                
            End If
            
        End If
        
    End With
    
Next i

End Sub
Public Function ShotWidth(ByVal ShotWhat As String) As Integer

'Select Case UCase(ShotWhat)
'    Case "REDSHOT"
        ShotWidth = frmPics!picShot(0).Width
        
'End Select

End Function

Public Function ShotHeight(ByVal ShotWhat As String) As Integer

'Select Case UCase(ShotWhat)
'    Case "REDSHOT"
        ShotHeight = frmPics!picShot(0).Height
    
'End Select
    
End Function


Public Function ShotDamage(ByVal ShotWhat As Integer) As Integer
Dim i As Integer

With ShotSets(ShotWhat)
    ShotDamage = SomRand(.Setting(SetInd("DAM", .Setting)).SettingData, .Setting(SetInd("DAMLOW", .Setting)).SettingData)
End With

End Function


Public Sub ShootNow(ByVal WhichFace As Byte, ByVal Index As Integer)

Dim i As Byte
Dim Times As Single

'Sets the number of times the shot moves in one timer event
Dim SInd As Byte
Dim ShotTime As String
Dim TrailNums As Integer
Dim Accel As Single
'Dim Seeking As Integer
'Dim SeekSpeed As Single
'Dim SeekNow As Boolean

SInd = Shooting(Index).What 'ShotIndex(Shooting(Index).What, ShotSets)

'Get data from settings
ShotTime = ShotData(SInd, "SPEED")
TrailNums = ShotData(SInd, "TRAIL")
Accel = ShotData(SInd, "ACCEL") + Face(WhichFace).ClassShotAccel
'Seeking = ShotData(SInd, "SEEKING")
'SeekSpeed = ShotData(SInd, "SEEKSPEED")

'Whether the shot should move towards opponent or not
'The higher 'seeking' is, the higher the chance
'SeekNow = (Seeking >= SomRand(100, 1))


'Checks if shot has been assigned a special movement time
With Shooting(Index)
    If .NewSpeed >= 0 Then
        Times = .NewSpeed
    Else
        Times = ShotTime
    End If
        
    If TrailNums > 0 Then
        Call NewTrail(Index, TrailNums)
    End If

'Was trying to make heat-seeking shots, but couldn't get it to work
'    If SeekNow Then
'        Dim OpFace As Byte
'        Dim AddAngle As Integer
'        Dim Target As Single
'        Dim ModTarget As Single
'        Dim TurnLeft As Single
'        Dim TurnRight As Single
'        Dim ChDirect As Integer
'
'        OpFace = FindOp(WhichFace)
'        Target = RadToDeg(Atn(Grad(Face(OpFace).Y, .ShotY, Face(OpFace).X, .ShotX))) + 90
'        AddAngle = FindLeftRight(RoundClosest(.Degrees, 45), .ShotX, .ShotY, Face(OpFace).X, Face(OpFace).Y)
'        Target = Target + AddAngle
'
'        If Target > 180 Then
'            ModTarget = Target - 360
'        Else
'            ModTarget = Target
'        End If
'
'        TurnLeft = .Degrees - Target
'        TurnRight = 360 - TurnLeft 'Target - .Degrees
'
'        If Abs(TurnLeft) < Abs(TurnRight) Then
'            ChDirect = -1
'        Else
'            ChDirect = 1
'        End If
'
'        .Degrees = .Degrees + SeekSpeed * ChDirect '* -Sgn(.Degrees - Target)
'
'    End If
        
    'Calculate acceleration
    .AddSpeed = .AddSpeed + Accel
    Times = Times + .AddSpeed
                
    'Because the players no longer have to move just one unit either, i have
    'put the loop in the moveface function
    Call MoveFace("up", .Degrees, Index, .ShotX, .ShotY, 0, 0, False, Times * Face(WhichFace).ClassShotSpeed)  ', Shooting(Index).What
End With

End Sub

Sub NewTrail(ByVal ShotNum As Integer, Amount As Integer)

Dim i As Integer

For i = LBound(Trails) To UBound(Trails)

    With Trails(i)
        If .Wait = 0 Then
            .X = Shooting(ShotNum).ShotX
            .Y = Shooting(ShotNum).ShotY
            .Wait = Amount
                
            Exit For
        End If
    End With
    
Next i

End Sub

Sub CheckForWins()

Dim i As Byte
'This is in a seperate loop so the details get updated
'before all the waiting
For i = LBound(Face) To UBound(Face)

    'Death!!!
    If Face(i).HP <= 0 Then
        
        If GameType = 1 Then
            Call SendLastHurt(i)
            Call ServerSend
            
        End If
                
        'Refreshes screen so that the explosion and everything is drawn
        Me.Refresh
                
        If GameType <> 2 Then
            'Timer1.Enabled = False
            'IsPaused = True
            If Not IsPaused Then
                Call Pause(True)
            End If
        End If
        
        GameOver = True

        'Checks if the 'winner' is actually still alive
        If Face(FindOp(i)).HP > 0 Then
            Call Ending(FindOp(i))
        Else
            'Draws the game
            Call Ending(-1)
            Exit For
        End If
        
    End If
Next i


End Sub
Sub Ending(Winner As Integer)

Const EPriority = 15
NoRestart = True

Dim EndSound As String
EndSound = "sounds/end.wav"

Call PlaySound("sounds/die.wav", EPriority)

If Winner >= 0 Then
    
    'Show who killed who with what
    Call ShowKills(0, Winner)
    
    'Displays the winner!
    Call Wait(1)
    Call PlaySound(EndSound, EPriority)
    lblWinner.Caption = Face(Winner).Name
    lblWinner.ForeColor = frmStats!lblTeamName(Winner).ForeColor
    lblWinner.Visible = True
    
    'Does a fancy pause before showing 'Wins!' :-)
    Call Wait(1.2)
    Call PlaySound(EndSound, EPriority)
    lblWins.Caption = "WINS"
    lblWins.Visible = True
    Call PlaySound("sounds/applause.wav", EPriority)
    
Else
    
    Call ShowKills(0, 0)
    Call ShowKills(1, 1)
    
    Call Wait(1)
    Call PlaySound(EndSound, EPriority)
    lblWins.Caption = "DRAW GAME"
    lblWins.Visible = True
End If

Call Wait(1.2)
lblInstruct.Visible = True
Call PlaySound(EndSound, EPriority)

NoRestart = False

End Sub

Sub ShowKills(Num As Byte, ByVal Killer As Byte)

Const Space = 7

'Display the 'Winner kills loser with (weapon)' thing
    Dim DrawWhat As Object
    Dim DrawWhatMask As Object
    Dim LastWep As Integer
    Dim PicWidth As Integer
    Dim PicHeight As Integer
    
    LastWep = Face(FindOp(Killer)).LastHurt
    lblWinKill(Num).Caption = Face(Killer).Name
    lblWinKill(Num).Visible = True
    lblWinLose(Num).Caption = Face(FindOp(Killer)).Name
    lblWinLose(Num).Visible = True
    
        
    If LastWep >= 0 Then
        Set DrawWhat = frmPics.picShot(LastWep)
        Set DrawWhatMask = frmPics.picShotMask(LastWep)
    Else
        Set DrawWhat = frmPics.picPowerUp(LastWep * -1)
        Set DrawWhatMask = frmPics.picPowerUpMask(LastWep * -1)
    End If
    
    PicWidth = DrawWhat.Width
    PicHeight = DrawWhat.Height
    
    'Move the labels and the picture box
    With picWinWep(Num)
        
        'Resizes the picture box to fit the picture (can't use autoresize with bitblt)
        .Width = PicWidth
        .Height = PicHeight
        .Left = lblWinKill(Num).Left + lblWinKill(Num).Width + Space
        .Top = lblWinKill(Num).Top + lblWinKill(Num).Height / 2 - .Height / 2
        lblWinLose(Num).Left = .Left + .Width + Space

        .Visible = True
        'Draws the killing shot in the picture box
        Call BitBlt(.hdc, 0, 0, PicWidth, PicHeight, DrawWhatMask.hdc, 0, 0, vbMergePaint)
        Call BitBlt(.hdc, 0, 0, PicWidth, PicHeight, DrawWhat.hdc, 0, 0, vbSrcAnd)
        .Refresh
    End With

End Sub

Public Sub AI(ByVal WFace As Byte)

Dim DoThis As String
Dim Key As Integer
Dim RanNum As Byte
Dim Action(4) As String

'Decides what actions to do depending on what
'checkboxes were checked at the start
If PlayerAI(WFace, 3) = True Then
    Action(3) = AIFindPower(WFace)
End If

If PlayerAI(WFace, 1) = True And Action(3) = "" Then
    Action(1) = AIShoot(WFace)
End If

'Will only choose a weapon if it has opponent in sight,
'so it doesn't interfere with mine planting (done in
'AIRAN)
If PlayerAI(WFace, 4) = True Then 'And Action(1) <> ""
    Action(4) = ChooseWep(WFace)
End If

If PlayerAI(WFace, 2) = True And (Action(1) <> "" Or PlayerAI(WFace, 1) = False) Then
    Action(2) = MUp
End If

If PlayerAI(WFace, 0) = True And Action(3) = "" And Action(1) = "" Then
    Action(0) = AIRan(WFace)
End If

    
'Select Case PlayerAI(WhichFace)
'    Case 1
'        Action(0) = AIRan(WhichFace)
'    Case 2
'
'        'Shoots at you if it has the chance
'        DoThis = AIShoot(WhichFace)
'        If DoThis <> "" Then
'            Action(0) = DoThis
'            Action(1) = ChooseWep(WhichFace)
'            Action(2) = AIMoveBack(WhichFace)
'        Else
'            Action(0) = AIRan(WhichFace)
'        End If
'    Case 3
'
'        'Chases you around everywhere *VERY ANNOYING!!!*
'        Action(0) = AIShoot(WhichFace)
'        Action(1) = ChooseWep(WhichFace)
''        Randomize
''        If Int(1 * Rnd) = 0 Then
'            Action(2) = MUp
''        End If
'        Action(3) = AIMoveBack(WhichFace)
'    Case 4
'
'        'Finds power-ups and shoots at you
'        'Would rather go for a powerup than shoot
'        Action(0) = AIFindPower(WhichFace)
'        If Action(0) = "" Then
'            DoThis = AIShoot(WhichFace)
'            If DoThis <> "" Then
'                Action(0) = DoThis
'                Action(1) = ChooseWep(WhichFace)
'                Action(2) = AIMoveBack(WhichFace)
'            Else
'                Action(0) = AIRan(WhichFace)
'            End If
'        End If
'
'End Select


'Uses the actual KeyDown events, so theoretically,
'there is nothing the computer can do that a
'human can't

Dim i As Integer

With Face(WFace)
'    Call Form_KeyUp(.PKeys(5), 0)
'    Call Form_KeyUp(.PKeys(6), 0)
'    Call Form_KeyUp(.PKeys(2), 0)
'    Call Form_KeyUp(.PKeys(3), 0)
    PressKeys(WFace, 5) = False
    PressKeys(WFace, 6) = False
    PressKeys(WFace, 2) = False
    PressKeys(WFace, 3) = False
    
    'Call Form_KeyUp(96, 0)
    'For i = LBound(Action) To UBound(Action)
    For i = UBound(Action) To LBound(Action) Step -1
        Select Case Action(i)
            Case TLeft: Key = 2 '.PKeys(2)
            Case TRight: Key = 3 '.PKeys(3)
            Case MDown: Key = 1 '.PKeys(1)
            Case MUp: Key = 0 '.PKeys(0)
            Case ShootAI: Key = 4 '.PKeys(4)
            Case ChDownAI: Key = 5 '.PKeys(5)
            Case ChUpAI: Key = 6 '.PKeys(6)
            Case Else
                Key = -1
        End Select
        If Key <> -1 Then
            'Call DoKeys(Key)
            PressKeys(WFace, Key) = True
        End If
    Next i
End With
    
    
'Else
'    Call Form_KeyUp(Asc("A"), 0)
'    Call Form_KeyUp(Asc("D"), 0)
'    Call Form_KeyUp(Asc("T"), 0)
'    Call Form_KeyUp(Asc("Y"), 0)
'
'    For i = LBound(Action) To UBound(Action)
'        Select Case Action(i)
'            Case TLeft
'                Key = Asc("A")
'            Case TRight
'                Key = Asc("D")
'            Case MDown
'                Key = Asc("S")
'            Case MUp
'                Key = Asc("W")
'            Case ShootAI
'                Key = Asc("G")
'            Case ChDownAI
'                Key = Asc("T")
'            Case ChUpAI
'                Key = Asc("Y")
'            Case Else
'                Key = -1
'        End Select
'        If Key <> -1 Then
'            Call DoKeys(Key)
'        End If
'
'    Next i
'End If


End Sub


Public Sub CoverUp(ByVal XVal As Integer, ByVal YVal As Integer)
'Covers up all sprites using a blank picture so that they don't lead a trail.

Call BitBlt(Me.hdc, XVal, YVal, 70, 70, frmPics!picBG.hdc, 0, 0, vbSrcCopy)

End Sub

Private Sub mnuBackGround_Click()

'If Timer1.Enabled Then
If Not IsPaused Then
    Call Pause
End If

With cdlg
    .Color = Me.BackColor
    .ShowColor
    Call ChColours(.Color)
End With

Call DrawFaces

Call SaveSetting(AppName, "Settings", "BackGround", Me.BackColor)

Me.Refresh

End Sub

Private Sub mnuClear_Click()

If Not NoRestart Then
    'This is needed to stop residue images from being drawn after reset
    'If Timer1.Enabled = True Then
    If Not IsPaused Then
        Call Pause
    End If
    Dim Reply As Byte
    Reply = MsgBox("Start next round? Current players will be kept", vbOKCancel, "Next Round")
    
    If Reply = vbOK Then
        Call Reset
        Call GetReady
    End If
End If

End Sub

Private Sub mnuKeys_Click()

frmKeys.Show vbModal

End Sub

Private Sub mnuNew_Click()

If Not NoRestart Then
    Dim Reply As Byte
    Reply = MsgBox("Start new game?", vbOKCancel, "New Game?")
    
    If Reply = vbOK Then
        Call NewGame
    End If
End If
    
End Sub

Sub mnuOptions_Click()

frmOptions.Show vbModal

End Sub

Private Sub mnuPause_Click()
Call Pause

End Sub

Private Sub mnuQuit_Click()
Unload Me
End Sub

Private Sub mnuResize_Click()

If MsgBox("Resizing the game window will restart the round.", vbOKCancel, "Resize") = vbOK Then

    frmResize.Show vbModal
    Call Reset
    frmChoseShots.Show vbModal
End If

End Sub

Private Sub mnuSkins_Click()
frmSkins.Show vbModal
End Sub

'Private Sub Timer1_Timer()
Sub MainLoop()
'This is where everything happens
'Runs pretty slowly at times (when there's lots of shots on the screen)
'and i'm not sure if its my code, or if using bitblt to draw stuff
'is just slow and i should learn DirectDraw.

'Doesn't run slowly anymore! Getting rid of the timer speeded the game
'up a lot.

'For the (sole) purpose of increasing program speed, I have
'not modularised many of the procedures in here. It saves dimming more variables
'and sending arguments back and forth.

If Not BadData Then
    Dim StartTime As Single
    
    If GameType <> 2 Then
    
        Dim RanNum As Single
        Dim i As Integer
        Dim n As Integer
        
        Dim PicWidth As Integer
        Dim PicHeight As Integer
        Dim PicInd As Integer
        
        Dim FrameCounter As Integer
        Dim StartFrameCount As Single

        Do
            
            If Not IsPaused And Timer - StartTime >= Val(BSettings(SetInd("GAMESPEED", BSettings)).SettingData) / 1000 Then
            
                StartTime = Timer
                
                'Calculate fps
                If Timer - StartFrameCount >= 1 Then
        
                    StartFrameCount = Timer
                    lblFrameRate.Caption = SBrack & FrameCounter & " fps" & EBrack
                    FrameCounter = 0
            
                End If
                
                FrameCounter = FrameCounter + 1
                
                Call DoKeyStates
                                    
                Call UpdateKaboom
                    
                Call UpdateTrails
                
                Call CoverStuff
                
                'Makes the flowers shoot
                Call FlowerShoot
                
                ''Teleports' players if they collide with each other
                If Distance(Face(0).X, Face(1).X, Face(0).Y, Face(1).Y) < Face(0).Width / 2 + Face(1).Width / 2 Then
                    Call RandPlace(0)
                    Call RandPlace(1)
                End If
                
                'Everything that happens to each face is in this loop
                For i = LBound(Face) To UBound(Face)
                    
                    With Face(i)
                        'Checks to see if player is turning
                        'If DegreeChange(i) <> 0 Then
                        If PressKeys(i, 2) Then
                            Call Turning(i, -45) 'DegreeChange(i))
                        End If
                        
                        If PressKeys(i, 3) Then
                            Call Turning(i, 45)
                        End If
                    
                        'Checks if player is moving up/down
                        'If .UpDown <> "" Then
                        
                        Dim UpDown As String
                        If PressKeys(i, 0) Then
                            UpDown = "up"
                        ElseIf PressKeys(i, 1) Then
                            UpDown = "down"
                        Else
                            UpDown = ""
                        End If
                        
                        If UpDown <> "" Then
                            Call MoveFace(UpDown, .Degrees, i, .X, .Y, .Height, .Width, True, .ClassSpeed)
                        End If
                        
                    
                        'Reduces the reload waiting time by 1
                        If .ShotWait <> 0 Then
                            .ShotWait = .ShotWait - 1
                        End If
                    
                        'Checks if player is holding shoot button and reload time has passed
                        'If so, creates a new shot
                        'If .HoldShot And .ShotWait = 0 Then
                        If PressKeys(i, 4) And .ShotWait = 0 Then
                            If GameType = 0 Or (GameType = 1 And i = 0) Then
                                Call ShootNew(i, .Degrees)
                            Else
                                Call ShootNew(i, .Degrees, UsingShot)
                            End If
                        End If
                        
                        If PressKeys(i, 5) Then
                            Call ShotChange(i, False)
                        ElseIf PressKeys(i, 6) Then
                            Call ShotChange(i, True)
                        End If
                        
                        
                    End With
                    
                    'Call ShowModTime(i)
                    Call ShowMods(i)
                    Call DecayHP(i)
                               
                Next i
                
                'Creates powerups
                RanNum = SomRand(100, 1, 2)
                If RanNum <= BSettings(SetInd("POWERUPPERCENT", BSettings)).SettingData Then
                
                    'This is probably a better way of determining what powerup to put in
                    'The last method I used was really messy
                    
                    Dim Ran2 As Integer
                    Dim PowerTemp As Integer
                    Dim PickedThis As Integer
                    
                    PickedThis = -1 'For debugging purposes, just in case
                    PowerTemp = 0
                    Ran2 = SomRand(PowerTotal, 0)
                    
                    For n = LBound(PowerUps) To UBound(PowerUps)
                        With PowerUps(n)
                            PowerTemp = PowerTemp + .Setting(SetInd("CHANCE", .Setting)).SettingData
                        End With
                        If PowerTemp >= Ran2 Then
                            PickedThis = n
                            Exit For
                        End If
                    Next n
                    
                    For n = LBound(PowerUpNow) To UBound(PowerUpNow)
                        'If PowerUpNow(n).What = "" Then
                        If PowerUpNow(n).What < 0 Then
                            With PowerUpNow(n)
                                .What = PickedThis 'ShotNameFind(PickedThis, PowerUps)
                                .X = SomRand(Me.ScaleWidth - 40, 20)
                                .Y = SomRand(Me.ScaleHeight - 40, 20)
                            End With
                            Call PlaySound("sounds/newpower.wav", 3)
                            Exit For
                        End If
                    Next n
                    
                End If
            
                Call DrawPowers
        
                'Moves the shots and draws them. Also updates the
                ''expiry' time
                
                For i = LBound(Shooting) To UBound(Shooting)
                    
                    With Shooting(i)
                        'Check for shotwait is now here
                        If .What >= 0 Then
                        
                            Call ShootNow(.WhichFace, i)
                            
                            'Updates the expiry time
                            'Wont do anything if expiry is <0
                            If .Expire > 0 Then
                                .Expire = .Expire - 1
                            ElseIf .Expire = 0 Then
                                Call Hit(i)
                            End If
                                    
                            'Checks if it hasn't run into something yet after ShootNow and expiry
                            'NOT REDUNDANT
                            If .What >= 0 Then
                                 Call DrawShots(i)
                            End If
                        End If
                    End With
                    
                Next i
                
                'Draws the trails
                Call DrawTrails
                
                'Draw explosions
                Call DrawKaboom
        
                'Check for contact with powerups
                'and updates powerup info
                
                'This must be done *before* the faces are drawn, otherwise
                'when a powerup is picked up, the coverup thing will cover
                'the face as well for one timer loop
                'If DoProcess Then
                For n = LBound(Face) To UBound(Face)
                    For i = LBound(PowerUpNow) To UBound(PowerUpNow)
                        'If PowerUpNow(i).What <> "" Then
                        If PowerUpNow(i).What >= 0 Then
                            Call GotPowerUp(i, n)
                        End If
                    Next i
                    
                    With Face(n)
                        If .PlusArmour.TimeLeft > 0 Then
                            .PlusArmour.TimeLeft = .PlusArmour.TimeLeft - 1
                        ElseIf .PlusArmour.TimeLeft < 0 Then
                            .PlusArmour.TimeLeft = .PlusArmour.TimeLeft + 1
'                        Else
'                            .PlusArmour.Bonus = 0
                        End If
                        
                        If .PlusDam.TimeLeft > 0 Then
                            .PlusDam.TimeLeft = .PlusDam.TimeLeft - 1
                        ElseIf .PlusDam.TimeLeft < 0 Then
                            .PlusDam.TimeLeft = .PlusDam.TimeLeft + 1
'                        Else
'                            .PlusDam.Bonus = 0
                        End If
                    End With
                
                Next n
                'End If
            
                'Draws the faces
                Call DrawFaces
        
                'Send information through network
                If GameType = 1 Then
                    Call ServerSend
                End If
                                    
                Call CheckForWins
                        
                'Checks for shot-player collisions
                For i = LBound(Shooting) To UBound(Shooting)
                    If Shooting(i).What >= 0 Then
                        Call Collision(i)
                    End If
                Next i
                
                'Runs the AI
                For i = LBound(Face) To UBound(Face)
                    If IsAI(i) = True Then
                        Call AI(i)
                    End If
                Next i
                
                'Refreshes the screen, or it wont draw everything
                'If Timer1.Enabled = True Then Me.Refresh
                Me.Refresh
                
            End If
            
            DoEvents
            
        Loop
        
    End If


End If

End Sub

'Function DoProcess() As Boolean
'
'If GameType <> 2 Then
'    DoProcess = True
'End If
'
'End Function
