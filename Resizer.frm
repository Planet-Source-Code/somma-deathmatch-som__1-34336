VERSION 5.00
Begin VB.Form frmResize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1980
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdMax 
      Caption         =   "&Maximise"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdDef 
      Caption         =   "&Default"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblInstr 
      Caption         =   "Height"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblInstr 
      Caption         =   "Width"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const ASpace = 25


Private Sub cmdDef_Click()

'Sets settings back to the default values
txtWidth.Text = DefWidth
txtHeight.Text = DefHeight

End Sub

Private Sub cmdMax_Click()

'Maximises the window to fit onto the screen
Dim ScaleIt As Single
ScaleIt = frmPositions.ScaleWidth / frmPositions.Width

txtWidth.Text = Screen.Width * ScaleIt - ASpace
txtHeight.Text = (Screen.Height - frmStats.Height) * ScaleIt - AddHeight
End Sub

Private Sub cmdOK_Click()

If IsNumeric(txtWidth.Text) And IsNumeric(txtHeight.Text) Then
    If CInt(txtWidth.Text) >= 300 And CInt(txtHeight.Text) >= 200 Then
        
        Dim ScaleIt As Single
        
        'Resizes the game window
        With frmPositions
            ScaleIt = .Width / .ScaleWidth
            .Width = Int(txtWidth.Text) * ScaleIt
            .Height = Int(txtHeight.Text) * ScaleIt
            .Left = ASpace / 2 * ScaleIt
            .Top = 0
        End With
    Else
        MsgBox "Please make the dimensions bigger.", vbInformation
    End If
Else
    MsgBox "Please enter numbers only.", vbInformation
End If

Me.Hide

End Sub

Private Sub Form_Activate()

With frmPositions
    txtWidth.Text = .ScaleWidth
    txtHeight.Text = .ScaleHeight + AddHeight
End With

End Sub

