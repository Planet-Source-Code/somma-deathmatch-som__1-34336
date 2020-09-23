VERSION 5.00
Begin VB.Form frmLoading 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4035
   ControlBox      =   0   'False
   FillColor       =   &H00BBEBFD&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Positions.ctrSomBar SbarLoad 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   750
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   450
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label lblLoading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[Loading....]"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

'Centers the window
With frmPositions
    Me.Top = .Top + .Height / 2 - Me.Height / 2
    Me.Left = .Left + .Width / 2 - Me.Width / 2
End With

With SbarLoad
    .BackColour = vbBlack
    .BarColour = &HFF00&
    .Max = 100
End With

lblVersion.Caption = SBrack & "Version " & App.Major & "." & App.Minor & "." & App.Revision & EBrack

End Sub

