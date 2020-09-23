VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00000080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5640
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4035
   ControlBox      =   0   'False
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   269
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help!"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdPowerUpSets 
      Caption         =   "&Power-Up Settings"
      Height          =   615
      Left            =   3150
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdShotSets 
      Caption         =   "&Shot Settings"
      Height          =   615
      Left            =   3150
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtChSet 
      BackColor       =   &H00E2F4FC&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblInformation 
      BackStyle       =   0  'Transparent
      Caption         =   "Stuff goes in here"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblDesc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NAME OF SETTING"
      ForeColor       =   &H00E2F4FC&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'All the 'numsonly' stuff you see around here happened when I was
'trying to figure out how to stop textboxes from accepting non-number characters
'It doesn't work too well, but I left it all in anyway to show you one
'of my failures. I am not a perfect coder, and I believe my code should
'reflect that :-)

Dim NumsOnly() As NumsOnly

Dim DontRefreshSets As Boolean

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
'Dim Invalid As Boolean
'If IsNumeric(txtChHP.Text) Then
'    If CStr(Int(txtChHP.Text)) = txtChHP.Text Then
'        MsgBox Int(txtChHP.Text) & Chr(13) & txtChHP.Text
'        Unload Me
'    Else
'        Invalid = True
'    End If
'Else
'    Invalid = True
'End If
'
'If Invalid = True Then
'    MsgBox "Invalid Entry"
'End If
Call ChSettings(BSettings, "", txtChSet)
Unload Me
End Sub

Private Sub cmdPowerUpSets_Click()

With frmShotOptions
    ChangeWhat = PowerUps
    .lblInformation.Caption = "Select a Power-Up to change..."
    .Show vbModal
    .Caption = "Power-up options"
End With

End Sub

Private Sub cmdShotSets_Click()

With frmShotOptions
    ChangeWhat = ShotSets
    .lblInformation.Caption = "Select a Shot to change..."
    .Show vbModal
    .Caption = "Shot options"
End With

End Sub

Private Sub Form_Activate()
'txtChHP.Text = Settings.PlayerHP
'ChHPOld.OldText = txtChHP.Text
'ChHPOld.SelPosition = txtChHP.SelStart
ReDim NumsOnly(txtChSet.UBound) As NumsOnly
Dim i As Byte

'For i = txtChSet.LBound To txtChSet.UBound
'    Call NumsOnlyUpdate(i)
'Next i
Call NumsOnlyUpdate(txtChSet, NumsOnly)

If DontRefreshSets = False Then
    For i = txtChSet.LBound To txtChSet.UBound
        txtChSet(i).Text = BSettings(i).SettingData
    Next i
    DontRefreshSets = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
DontRefreshSets = False
End Sub

Private Sub txtChSet_Change(Index As Integer)
'Stops user from typing in letters in numbers-only boxes

'...but only if the setting was numbers only in the first place
'If IsNumeric(BSettings(Index).SettingData) Then
'    Dim NotNumber As Boolean
'
'    If IsNumeric(txtChSet(Index).Text) Then
'        If CStr(Int(txtChSet(Index).Text)) = txtChSet(Index).Text Then
''            ChHPOld.OldText = txtChHP.Text
''            ChHPOld.SelPosition = txtChHP.SelStart
'            Call NumsOnlyUpdate(txtChSet, NumsOnly)
'        Else
'            NotNumber = True
'        End If
'    Else
'        NotNumber = True
'    End If
'
'    If NotNumber = True Then
'        Beep
'        txtChSet(Index).Text = NumsOnly(Index).OldText
'        txtChSet(Index).SelStart = NumsOnly(Index).SelPosition
'    End If
'End If

Call NumsOnlyCheck(txtChSet, Index, NumsOnly, BSettings)

End Sub


Private Sub Form_Load()

Dim i As Byte

'The space between each desc label
Const Space = 4

lblInformation.Caption = "Some changes will not take effect until the beginning of the next round." & Chr(13) & Chr(13) & "PLEASE DON'T CHANGE ANYTHING IF YOU DON'T KNOW WHAT YOU'RE DOING!!!"

'Loads control arrays at runtime, since the number of
'settings will grow, and i can't be bothered copying and
'pasting each time i add a setting. Besides, this is more
'*elegant*.
For i = LBound(BSettings) To UBound(BSettings)
    'If BSettings(i).SetDesc <> "" Then
        If i <> 0 Then
            Load lblDesc(i)
            Load txtChSet(i)
            lblDesc(i).Visible = True
            txtChSet(i).Visible = True
            lblDesc(i).Left = lblDesc(0).Left
            lblDesc(i).Top = lblDesc(i - 1).Top + lblDesc(i - 1).Height + Space
        End If
        
        txtChSet(i).Left = txtChSet(0).Left
        txtChSet(i).Top = lblDesc(i).Top - 3
    'End If
Next i

'Sets the window height to accomadate the labels and text boxes
Me.Height = (((lblDesc.UBound + 1) * (lblDesc(0).Height + Space)) + lblInformation.Height + cmdOK.Height + 20) * (Me.Height / Me.ScaleHeight)

Dim ButtonTop As Integer
ButtonTop = Me.ScaleHeight - cmdOK.Height - Space
cmdOK.Top = ButtonTop
cmdCancel.Top = ButtonTop
cmdHelp.Top = ButtonTop

For i = lblDesc.LBound To lblDesc.UBound
    lblDesc(i).Caption = BSettings(i).SetDesc
Next i
    
    
End Sub
