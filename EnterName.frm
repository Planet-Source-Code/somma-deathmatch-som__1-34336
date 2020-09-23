VERSION 5.00
Begin VB.Form frmEnterName 
   BackColor       =   &H00000080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4635
   ClientLeft      =   2730
   ClientTop       =   3435
   ClientWidth     =   3525
   ControlBox      =   0   'False
   ForeColor       =   &H00808080&
   Icon            =   "EnterName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   3525
   Begin Positions.Somcmd somcmdOK 
      Height          =   1215
      Left            =   2160
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2880
      Top             =   960
   End
   Begin VB.OptionButton optIsAI 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "&AI"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.Frame fraAIOpts 
      BackColor       =   &H00000080&
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
      Begin VB.CheckBox chkAI 
         BackColor       =   &H00000080&
         Caption         =   "Plants Mines"
         Enabled         =   0   'False
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CheckBox chkAI 
         BackColor       =   &H00000080&
         Caption         =   "Chooses Weapon"
         Enabled         =   0   'False
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkAI 
         BackColor       =   &H00000080&
         Caption         =   "Finds Powerups"
         Enabled         =   0   'False
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkAI 
         BackColor       =   &H00000080&
         Caption         =   "Chases You"
         Enabled         =   0   'False
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkAI 
         BackColor       =   &H00000080&
         Caption         =   "Smart Shooting"
         Enabled         =   0   'False
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkAI 
         BackColor       =   &H00000080&
         Caption         =   "Random Moves"
         Enabled         =   0   'False
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
   End
   Begin VB.OptionButton optIsAI 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "&Human"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox txtEnterName 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   40
      MaxLength       =   30
      TabIndex        =   0
      Text            =   "Input Name..."
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label lblEnterYourName 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your name..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   255
      Left            =   45
      TabIndex        =   11
      Top             =   90
      Width           =   2055
   End
   Begin VB.Label lblWaiting 
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting for opponent..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   2055
   End
End
Attribute VB_Name = "frmEnterName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim OldName As String

Dim IamDone As Boolean
Dim DragForm As DragIt

Private Sub Form_Activate()

txtEnterName.SetFocus
lblWaiting.Visible = False

IamDone = False

'Sets the text to the same colour as the player
txtEnterName.ForeColor = frmStats.lblTeamName(ProcessPlayer).ForeColor
Call SelAll

End Sub

Private Sub Form_Load()

DoEvents
optIsAI(0).Value = 1

With somcmdOK
    .TextPos = 360
    .BorderShape = vbShapeCircle
    .BorderWidth = 3
    .FontSize = 20
    .Default = True
    .BorderColour = &HFF00&
    .BackColour = Me.BackColor
    .BorderOpaque = True
    .BorderFillColour = vbBlack
    .Caption = SBrack & "OK" & EBrack
End With

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Dragging the window
With DragForm
    .Dragging = True
    .XStart = X
    .YStart = Y
End With

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Dragging the window
With DragForm
    If .Dragging Then
        Me.Left = Me.Left + X - .XStart
        Me.Top = Me.Top + Y - .YStart
    End If
End With

Call somcmdOK.StopHover

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Stop dragging
DragForm.Dragging = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

DragForm.Dragging = False

End Sub

Private Sub optAI_Click(Index As Integer)

End Sub

Private Sub optIsAI_Click(Index As Integer)

Dim i As Integer
Dim AISel As Boolean

AISel = optIsAI(1).Value

'Enables/disables all the AI checkboxes
For i = chkAI.LBound To chkAI.UBound
    'I've disabled 'plant mines' for now
    If i <> 5 Then
        chkAI(i).Enabled = AISel
    End If
Next i

With txtEnterName
    If AISel Then
        OldName = .Text
        .Text = "Som-AI"
    Else
        .Text = OldName
    End If
End With

End Sub

Private Sub somcmdOK_Click()

With Face(ProcessPlayer)

    If txtEnterName.Text <> "" Then
        .Name = Trim(txtEnterName.Text)
    Else
        .Name = "Anonymous " & ProcessPlayer + 1
    End If
    
    If GameType <> 0 Then
        Call SendName(.Name)
    End If

End With

'Whether the player is AI
IsAI(ProcessPlayer) = optIsAI(1).Value

'Stores what AI procedures will run
If IsAI(ProcessPlayer) Then
    Dim i As Byte
    For i = chkAI.LBound To chkAI.UBound
        PlayerAI(ProcessPlayer, i) = chkAI(i).Value
    Next i
End If

If GameType = 0 Then
    Unload Me
Else
    IamDone = True
End If

End Sub

Private Sub Timer1_Timer()

If IamDone Then
    lblWaiting.Visible = True
    If GameType <> 0 And OpReady Then
        OpReady = False
        Unload Me
    End If
End If

End Sub

Private Sub txtEnterName_Change()

Call PlaySound("sounds/type.wav", 1)

End Sub

Private Sub txtEnterName_DblClick()

Call SelAll

End Sub

Sub SelAll()

'I can't seem to pass a textbox as an object,
'as VB defaults to the text property when no
'property is specified. Hence i cant simply
'make a sub with the textbox as the parameter
'to do this. *Stupid VB*
With txtEnterName
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub

