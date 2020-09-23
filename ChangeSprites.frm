VERSION 5.00
Begin VB.Form frmSkins 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Skins"
   ClientHeight    =   3615
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3015
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPics 
      Caption         =   "Show &Pics"
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoadPowers 
      Caption         =   "Load"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   2370
      Width           =   735
   End
   Begin VB.CommandButton cmdLoadShots 
      Caption         =   "Load"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   1770
      Width           =   735
   End
   Begin VB.TextBox txtPowerSprites 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtShotSprites 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdLoadFace 
      Caption         =   "Load"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   750
      Width           =   735
   End
   Begin VB.TextBox txtFaceSprite 
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtFaceSprite 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Power-Ups Folder:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Shots Folder:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Face 2:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1800
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Face 1:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "frmSkins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim SkinBad As Boolean

Private Sub cmdClose_Click()

If Not SkinBad Then
    Unload Me
Else
    If MsgBox("This will exit the game. Are you sure?", vbYesNo, "Exit Game?") = vbYes Then
        End
    End If
End If

End Sub

Private Sub cmdLoadFace_Click()

Dim i As Byte

SkinBad = False
Call ShowBusy
For i = 0 To UBound(Face)
    FaceSprites(i) = txtFaceSprite(i).Text
Next i

Call SaveSprites
Call frmPositions.LoadFacePics(True)

If Not SkinBad Then
    Call ShowBusy(True)
Else
    Call ShowBadSkin
    cmdLoadFace.Enabled = True
End If

End Sub

Private Sub cmdLoadPowers_Click()

SkinBad = False

Call ShowBusy

PowerSprites = txtPowerSprites.Text
Call SaveSprites
Call frmPositions.LoadPowerPics(True)

If Not SkinBad Then
    Call ShowBusy(True)
Else
    Call ShowBadSkin
    cmdLoadPowers.Enabled = True
End If

End Sub

Private Sub cmdLoadShots_Click()

SkinBad = False

Call ShowBusy

ShotSprites = txtShotSprites.Text
Call SaveSprites
Call frmPositions.LoadShotPics(True)

If Not SkinBad Then
    Call ShowBusy(True)
Else
    Call ShowBadSkin
    cmdLoadShots.Enabled = True
End If

End Sub

Private Sub cmdPics_Click()

Unload Me
frmPics.Show vbModal

End Sub

Private Sub Form_Load()

Dim i As Byte
For i = 0 To UBound(Face)
    txtFaceSprite(i).Text = FaceSprites(i)
Next i

txtShotSprites.Text = ShotSprites
txtPowerSprites.Text = PowerSprites

End Sub

Sub ShowBusy(Optional StopIt As Boolean)

If StopIt Then
    Me.MousePointer = 0
Else
    Me.MousePointer = vbHourglass
End If

cmdLoadFace.Enabled = StopIt
cmdLoadShots.Enabled = StopIt
cmdLoadPowers.Enabled = StopIt
cmdClose.Enabled = StopIt
cmdPics.Enabled = StopIt

End Sub

Sub GetSprites()

FaceSprites(0) = GetSetting(AppName, SetSec, "FaceSprite" & 0, "pics/face")
FaceSprites(1) = GetSetting(AppName, SetSec, "FaceSprite" & 1, "pics/f2ce")

ShotSprites = GetSetting(AppName, SetSec, "ShotSprites", "pics")
PowerSprites = GetSetting(AppName, SetSec, "PowerSprites", "pics")

End Sub

Sub SaveSprites()

Dim i As Byte
For i = 0 To UBound(Face)
    Call SaveSetting(AppName, SetSec, "FaceSprite" & i, FaceSprites(i))
Next i

Call SaveSetting(AppName, SetSec, "ShotSprites", ShotSprites)
Call SaveSetting(AppName, SetSec, "PowerSprites", PowerSprites)

End Sub

Sub BadSkin()

MsgBox "An error occured while trying to load from the skins folder." & Chr(13) & "Make sure the path is correct, and all neccesary files are there.", vbExclamation, "Error"
SkinBad = True

End Sub

Sub ShowBadSkin()

cmdClose.Enabled = True
Me.MousePointer = 0

End Sub
