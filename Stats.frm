VERSION 5.00
Begin VB.Form frmStats 
   BackColor       =   &H00000040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2430
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   162
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   589
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin Positions.ctrSomBar SbarHP 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
   End
   Begin VB.PictureBox picShowShot 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   5160
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   13
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox picShowShot 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3360
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   12
      Top             =   1440
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   120
      Top             =   1200
   End
   Begin Positions.ctrSomBar SbarHP 
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   17
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Explosive (Range)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   25
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   24
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Armour Bonus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   7200
      TabIndex        =   23
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Armour Bonus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   22
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Damage Bonus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   21
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Damage Bonus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblClassType 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   300
      Index           =   1
      Left            =   4560
      TabIndex        =   19
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblClassType 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   300
      Index           =   0
      Left            =   3360
      TabIndex        =   18
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblShotEx 
      BackStyle       =   0  'Transparent
      Caption         =   "Shot Damage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   15
      Top             =   2160
      Width           =   1350
   End
   Begin VB.Label lblShotEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Shot Damage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   14
      Top             =   2160
      Width           =   1350
   End
   Begin VB.Label lblPlusDam 
      BackStyle       =   0  'Transparent
      Caption         =   "PlusDam"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label lblPlusArmour 
      BackStyle       =   0  'Transparent
      Caption         =   "Armour"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   450
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label lblPlusArmour 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Armour"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   450
      Index           =   1
      Left            =   7440
      TabIndex        =   9
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label lblPlusDam 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PlusDam"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Index           =   1
      Left            =   7440
      TabIndex        =   8
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label lblShotDam 
      BackStyle       =   0  'Transparent
      Caption         =   "Shot Damage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   7
      Top             =   1800
      Width           =   1350
   End
   Begin VB.Label lblShotDam 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Shot Damage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   6
      Top             =   1800
      Width           =   1350
   End
   Begin VB.Label lblTeamName 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
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
      Height          =   495
      Index           =   1
      Left            =   5760
      TabIndex        =   5
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D6E6E9&
      Height          =   450
      Index           =   1
      Left            =   5760
      TabIndex        =   4
      Top             =   480
      Width           =   1200
   End
   Begin VB.Label lblCurShot 
      BackStyle       =   0  'Transparent
      Caption         =   "CurrentShot"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D6E6E9&
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblCurShot 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CurrentShot"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D6E6E9&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblHP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D6E6E9&
      Height          =   450
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   1200
   End
   Begin VB.Label lblTeamName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
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
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
'The main window *always* has focus

On Error GoTo BadFocus
frmPositions.SetFocus

Exit Sub

BadFocus:
'MsgBox "That thing happened!"
Debug.Print

End Sub

Private Sub Form_Load()

Const BackCol = &HBBEBFD
SbarHP(0).Invert = True
SbarHP(0).BackColour = BackCol
SbarHP(1).BackColour = BackCol

End Sub

Private Sub Timer2_Timer()

'Moves the Stats window
If frmPositions.Height <> Me.Height Or frmPositions.Top <> Me.Top Or frmPositions.Left <> Me.Left + Me.Width Then
    'Moves the stats window so it is always next to the main window
'    frmStats.Height = Me.Height
'    frmStats.Top = Me.Top
'    frmStats.Left = Me.Left + Me.Width
    With frmPositions
        Me.Top = .Top + .Height
        Me.Left = .Left + .Width / 2 - Me.Width / 2
    End With
End If

End Sub
