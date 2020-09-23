VERSION 5.00
Begin VB.Form frmKeys 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customise Keys"
   ClientHeight    =   4485
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4200
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkChar 
      BackColor       =   &H00000000&
      Caption         =   "Characters"
      ForeColor       =   &H00BBEBFD&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox txtFocus 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   4480
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblInstr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Character representation not always correct (sorry)."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblP2 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   10
      Top             =   1080
      Width           =   800
   End
   Begin VB.Label lblPDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   9
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label lblPDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   8
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label lblP1 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   7
      Top             =   1080
      Width           =   800
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Next Wep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   1050
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Prev Wep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1050
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Shoot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1050
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Turn Right"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Turn Left"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1050
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Backward"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Forward"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1050
   End
End
Attribute VB_Name = "frmKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Entering As Boolean
Dim WhoFace As Byte
Dim WhichKey As Integer



Private Sub chkChar_Click()

Call ShowAsc
txtFocus.SetFocus

End Sub

Private Sub cmdOK_Click()

Call SaveKeys
Me.Hide

End Sub


Private Sub Form_Load()

Call ShowAsc
Call InstrClick

End Sub

Sub InstrClick()

lblInstr.Caption = "Click on a key to change"

End Sub

Sub InstrNewKey()

lblInstr.Caption = "Hit new key"

End Sub

Sub Reset()

Dim i As Integer
For i = 0 To lblP1.UBound
    lblP1(i).BorderStyle = 1
    lblP2(i).BorderStyle = 1
Next i
Entering = False

End Sub

Sub ShowAsc()

'Sub for loading the labels. Also for showing the current
'key configuration (I don't know why I put them in the
'same procedure...)
Dim i As Integer

For i = 0 To lblDesc.UBound

    If i > lblP1.UBound Then
        Load lblP1(i)
        Load lblP2(i)
    End If
    
    With lblP1(i)
        .Top = lblDesc(i).Top
        
        If chkChar.Value = 0 Then
            .Caption = Face(0).PKeys(i)
        Else
            .Caption = KeyChr(Face(0).PKeys(i))
        End If
        
        .Visible = True
    End With
    
    With lblP2(i)
        .Top = lblDesc(i).Top
        
        If chkChar.Value = 0 Then
            .Caption = Face(1).PKeys(i)
        Else
            .Caption = KeyChr(Face(1).PKeys(i))
        End If
        
        .Visible = True
    End With

Next i

End Sub

Sub SaveKeys()

'Saves the key settings to the registry

Dim n As Integer
Dim i As Integer
Dim Temp As String

For n = 0 To UBound(Face)
    For i = 0 To UBound(Face(n).PKeys)
        Temp = Temp & Face(n).PKeys(i) & AComma
    Next i
    
    Call SaveSetting(AppName, SetSec, "Player" & n + 1 & "Keys", Temp)
    Temp = ""
    
Next n

End Sub

Function ValidKey(Code As Integer) As Boolean

'Checks whether the key is already in use

'Checks the pause key first
If Code = Asc("P") Then
    Exit Function
End If

Dim i As Integer
Dim f As Integer

ValidKey = True

For f = 0 To UBound(Face)
    For i = 0 To UBound(Face(f).PKeys)
        
        'Check whether the key is assigned to another action. It will be valid
        'if the key pressed is the one already assigned to the action
        If Face(f).PKeys(i) = Code And (f <> WhoFace Or i <> WhichKey) Then
            ValidKey = False
            Exit For
        End If
    Next i
Next f

End Function

Private Sub lblP1_Click(Index As Integer)

Call Reset
lblP1(Index).BorderStyle = 0

Entering = True
WhoFace = 0
WhichKey = Index

Call InstrNewKey

End Sub

Private Sub lblP2_Click(Index As Integer)

Call Reset
lblP2(Index).BorderStyle = 0

Entering = True
WhoFace = 1
WhichKey = Index

Call InstrNewKey

End Sub

Private Sub txtFocus_KeyDown(KeyCode As Integer, Shift As Integer)

'Unfortunately, i have to maintain the focus on a textbox (which you can't
'see) so that the command button doesn't steal the focus. This is why
'the keydown event is here
If Entering Then
    If ValidKey(KeyCode) Then
        Face(WhoFace).PKeys(WhichKey) = KeyCode
        Call Reset
        Call ShowAsc
        Call InstrClick
    Else
        MsgBox "That key is already in use.", vbInformation
        Call Reset
        Call InstrClick
    End If
    
    'For some reason, the form will freeze up after this sub.
    'The only way to unfreeze it was to make it lose and gain focus,
    'so this does it automatically.
    'It seems to be a conflict with the getkeystate function, even though
    'its a private function and has nothing to do with this form
    Me.Hide
    Me.Show vbModal
End If

End Sub


Function KeyChr(KeyCode As Integer) As String

Select Case KeyCode
    
    Case 13
        KeyChr = "Enter"
    Case 33
        KeyChr = "PgUp"
    Case 34
        KeyChr = "PgDn"
    Case 35
        KeyChr = "End"
    Case 36
        KeyChr = "Home"
    Case 45
        KeyChr = "Ins"
    Case 46
        KeyChr = "Del"
    Case 38
        KeyChr = "Up"
    Case 40
        KeyChr = "Down"
    Case 37
        KeyChr = "Left"
    Case 39
        KeyChr = "Right"
    Case 96
        KeyChr = "Num 0"
    Case 97
        KeyChr = "Num 1"
    Case 98
        KeyChr = "Num 2"
    Case 99
        KeyChr = "Num 3"
    Case 100
        KeyChr = "Num 4"
    Case 101
        KeyChr = "Num 5"
    Case 102
        KeyChr = "Num 6"
    Case 103
        KeyChr = "Num 7"
    Case 104
        KeyChr = "Num 8"
    Case 105
        KeyChr = "Num 9"
        
    Case 111
        KeyChr = "Num /"


    Case Else
        KeyChr = Chr(KeyCode)
End Select

End Function
