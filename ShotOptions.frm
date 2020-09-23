VERSION 5.00
Begin VB.Form frmShotOptions 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4005
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   267
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   247
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtChShotSet 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   6
      Text            =   "Setting Data"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ComboBox cboShotIndex 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblInformation 
      BackStyle       =   0  'Transparent
      Caption         =   "Select something"
      ForeColor       =   &H00BBEBFD&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmShotOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DontRefresh As Boolean
'Dim NumsOnly() As NumsOnly

Dim DragForm As DragIt
Dim Changed As Boolean
Dim BoxInd As Integer


Private Sub cboShotIndex_Click()

Dim OK As Boolean

If cboShotIndex.ListIndex <> BoxInd Then
    If Changed Then
        OK = (MsgBox("You have not saved changes. Are you sure?", vbYesNo) = vbYes)
        
        If Not OK Then
            cboShotIndex.ListIndex = BoxInd
        End If
        
    Else
        OK = True
    End If
End If

If OK Then
    Dim i As Byte
    
    For i = LBound(ChangeWhat(0).Setting) To UBound(ChangeWhat(0).Setting)
        txtChShotSet(i).Text = ChangeWhat(cboShotIndex.ListIndex).Setting(i).SettingData
    Next i
    
    Changed = False
    
    BoxInd = cboShotIndex.ListIndex

End If
    

End Sub

Private Sub cmdApply_Click()
Call Save
Changed = False
'Call ChSettings(ChangeWhat(cboShotIndex.ListIndex).Setting, cboShotIndex.ListIndex, txtChShotSet)
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Call Save
Unload Me
End Sub

Public Sub Save()
Dim ChArray As SettingsBetter
Dim ChkThis As String

'Check which array (shot settings or power-up settings) it is using
'This is actually redundant, because the program gets told what array
'to use when the form gets activated (the changewhat array), but VB can't
'store arrays by reference.
ChkThis = ChangeWhat(0).Setting(0).SetDesc
Select Case ChkThis
    Case ShotSets(0).Setting(0).SetDesc
        Call ChSettings(ShotSets(cboShotIndex.ListIndex).Setting, cboShotIndex.ListIndex, txtChShotSet)
    Case PowerUps(0).Setting(0).SetDesc
        Call ChSettings(PowerUps(cboShotIndex.ListIndex).Setting, cboShotIndex.ListIndex, txtChShotSet)
End Select

End Sub

Private Sub Form_Activate()
Dim i As Byte
'ReDim NumsOnly(txtChShotSet.UBound)
'Call NumsOnlyUpdate(txtChShotSet, NumsOnly)

BoxInd = -2
Changed = False

With frmOptions
    Me.BackColor = .BackColor
    lblDesc(0).ForeColor = .lblDesc(0).ForeColor
    txtChShotSet(0).BackColor = .txtChSet(0).BackColor
End With

'Dim i As Byte
'The space between each desc label
Const Space = 4

'Loads control arrays at runtime, since the number of
'settings will grow. Slower, but it's much better than copying and pasting
'every time a new setting is added.

For i = LBound(ChangeWhat(0).Setting) To UBound(ChangeWhat(0).Setting)
    'If BSettings(i).SetDesc <> "" Then
        If i <> 0 Then
            Load lblDesc(i)
            Load txtChShotSet(i)
            lblDesc(i).Visible = True
            txtChShotSet(i).Visible = True
            lblDesc(i).Left = lblDesc(0).Left
            lblDesc(i).Top = lblDesc(i - 1).Top + lblDesc(i - 1).Height + Space
        End If
        
        lblDesc(i).Caption = ChangeWhat(0).Setting(i).SetDesc
        txtChShotSet(i).Left = txtChShotSet(0).Left
        txtChShotSet(i).Top = lblDesc(i).Top - 3
    'End If
Next i

'txtChShotSet(0).Locked = True

'Sets the window height to accomadate the labels and text boxes
Me.Height = (((lblDesc.UBound + 1) * (lblDesc(0).Height + Space)) + lblInformation.Height + cmdOK.Height + cboShotIndex.Height + 35) * (Me.Height / Me.ScaleHeight)

Dim ButtonTop As Integer
ButtonTop = Me.ScaleHeight - cmdOK.Height - Space
cmdOK.Top = ButtonTop
cmdCancel.Top = ButtonTop
cmdApply.Top = ButtonTop

If DontRefresh = False Then
    For i = LBound(ChangeWhat) To UBound(ChangeWhat)
        cboShotIndex.List(i) = ShotNameFind(i, ChangeWhat)
    Next i
    cboShotIndex.ListIndex = LBound(ChangeWhat)
End If


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'For dragging the form around.
With DragForm
    .XStart = X
    .YStart = Y
    .Dragging = True
End With

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Dragging!
With DragForm
    If .Dragging Then
        Me.Left = Me.Left + X - .XStart
        Me.Top = Me.Top + Y - .YStart
    End If
End With

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Stop dragging
DragForm.Dragging = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
DontRefresh = False

Dim i As Integer
For i = lblDesc.UBound To 1 Step -1
    Unload lblDesc(i)
    Unload txtChShotSet(i)
Next i
    

End Sub

Private Sub txtChShotSet_Change(Index As Integer)

Changed = True

'Call NumsOnlyCheck(txtChShotSet, Index, NumsOnly, ChangeWhat(cboShotIndex.ListIndex).Setting)
End Sub
