VERSION 5.00
Begin VB.Form frmGameType 
   BackColor       =   &H00000080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4590
   ClientLeft      =   2730
   ClientTop       =   3435
   ClientWidth     =   4860
   ControlBox      =   0   'False
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
   Begin Positions.Somcmd somcmdQuit 
      Height          =   375
      Left            =   3120
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin Positions.Somcmd somcmdOK 
      Height          =   585
      Left            =   2760
      Top             =   2760
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1032
   End
   Begin VB.Frame fraPlayType 
      BackColor       =   &H00000080&
      Caption         =   "Play Type"
      ForeColor       =   &H00FFFF80&
      Height          =   975
      Left            =   2880
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
      Begin VB.OptionButton optPlayType 
         BackColor       =   &H00000080&
         Caption         =   "Free-Play"
         ForeColor       =   &H00E2F4FC&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optPlayType 
         BackColor       =   &H00000080&
         Caption         =   "Use Classes"
         ForeColor       =   &H00E2F4FC&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin Positions.LotsaSocks LotsaSocks 
      Left            =   2160
      Top             =   3480
      _ExtentX        =   661
      _ExtentY        =   873
   End
   Begin VB.TextBox txtPort 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Text            =   "for now."
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox txtAddress 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "Network Play disabled"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.OptionButton optType 
      BackColor       =   &H00000080&
      Caption         =   "Connect to a Host"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.OptionButton optType 
      BackColor       =   &H00000080&
      Caption         =   "Host Network Game"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
   Begin VB.OptionButton optType 
      BackColor       =   &H00000080&
      Caption         =   "Single Player/ Keyboard Battle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E2F4FC&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Port Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "IP / Server Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBEBFD&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   2055
   End
End
Attribute VB_Name = "frmGameType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const SCancel = SBrack & "Stop" & EBrack
Const OK = SBrack & "OK" & EBrack

Dim Selected As Byte
Dim Busy As Boolean

Dim DragForm As DragIt

Sub Serve()

If IsNumeric(txtPort.Text) Then

    'On Error GoTo ServeErr
    
    Call DoBusy
    
'    With PosSock
'        .LocalPort = txtPort.Text
'        .Listen
'    End With
    With LotsaSocks
        Call .ServerListen(txtPort.Text)
        If .WhatError <> "" Then
            MsgBox .WhatError
            Call DoBusy(True)
            Exit Sub
        End If
    End With
    
    ErrorStop = False
    
    IsServer = True
Else
    MsgBox "Please enter a port number", vbInformation
End If


'Exit Sub
    
'ServeErr:
'MsgBox Err.Description & Chr(13) & "Another program is using the connection. Close all programs that are running and try again."
'Call DoBusy(True)

End Sub

Sub BeClient()

If IsNumeric(txtPort.Text) Then
    'On Error GoTo ClientErr
    
    Call DoBusy
    
    'Call PosSock.Connect(txtAddress.Text, txtPort.Text)
    With LotsaSocks
        Call .ClientConnect(txtAddress.Text, txtPort.Text)
        If .WhatError <> "" Then
            MsgBox .WhatError
            Call DoBusy(True)
            Exit Sub
        End If
    End With
    
    IsServer = False
    
End If


'Exit Sub

'ClientErr:
'MsgBox Err.Description '"Could not connect."
'Call DoBusy(True)

End Sub

Private Sub Form_Activate()

Call DoBusy(True)
Call DoEnabled
LotsaSocks.Amount = 4

End Sub


Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "/newgtypebg.bmp")

With somcmdOK
    .BackColour = &H80&
    .FontSize = 25
    .Default = True
    .SoundHover = App.Path & "/sounds/type.wav"
    .UsePlaySoundSub = True
End With

With somcmdQuit
    .BackColour = &H80&
    .Caption = SBrack & "Quit" & EBrack
    .FontSize = 14
    .Cancel = True
'    .BorderShape = vbShapeRoundedRectangle
'    .BorderWidth = 2
'    .BorderColour = &HEEEEEE
    .SoundHover = App.Path & "/sounds/type.wav"
    .UsePlaySoundSub = True
End With

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Fore dragging the window around
With DragForm
    .Dragging = True
    .XStart = X
    .YStart = Y
End With

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Dragging
With DragForm
    If .Dragging Then
        Me.Left = Me.Left - .XStart + X
        Me.Top = Me.Top - .YStart + Y
    End If
End With

'Tells the buttons that the mouse is elsewhere
somcmdOK.StopHover
somcmdQuit.StopHover

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Stop dragging
DragForm.Dragging = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

DragForm.Dragging = False

End Sub

Private Sub optType_Click(Index As Integer)

Selected = Index

Call DoEnabled

End Sub

Sub DoEnabled()

Select Case Selected
    Case 0
        txtPort.Enabled = False
        txtAddress.Enabled = False
        fraPlayType.Enabled = True
    Case 1
        txtPort.Enabled = True
        txtAddress.Enabled = False
        fraPlayType.Enabled = True
    Case 2
        txtPort.Enabled = True
        txtAddress.Enabled = True
        fraPlayType.Enabled = False
End Select
    

End Sub

Sub DoBusy(Optional CancelIt As Boolean)

Dim i As Byte

Busy = Not CancelIt

txtPort.Enabled = CancelIt
txtAddress.Enabled = CancelIt

For i = 0 To optType.UBound
    optType(i).Enabled = CancelIt
Next i

If CancelIt Then
    somcmdOK.Caption = OK
Else
    somcmdOK.Caption = SCancel
End If

End Sub

Private Sub LotsaSocks_DataArrive(Data As String)

'Call DataArrive(Data)

'With frmServeData.txtData
'    .Text = .Text & ">" & Data & Chr(13) & Chr(10) & Chr(13) & Chr(10)
'End With

End Sub

Private Sub LotsaSocks_Connected(Server As Boolean)

Me.Hide
If Server Then
    Call LotsaSocks.SendStuff("PTYPE" & " " & PlayType & ASemi)
End If
    
End Sub

'Private Sub PosSock_Connect()
'
'Me.Hide
'
'End Sub



'Private Sub PosSock_ConnectionRequest(ByVal RequestID As Long)

'PosSock.Close
'PosSock.Accept RequestID
'Me.Hide

'End Sub

'Private Sub PosSock_DataArrival(ByVal bytesTotal As Long)
'
'NewData = ""
'Call PosSock.GetData(NewData)
'
'With frmServeData.txtData
'    .Text = .Text & ">" & NewData & Chr(13) & Chr(10) & Chr(13) & Chr(10)
'End With
'
'Call DataArrive(NewData)
'
'End Sub

'Private Sub PosSock_SendComplete()
'
'UnsentData = False
'
'End Sub

Private Sub somcmdOK_Click()

Dim i As Byte

If Not Busy Then

    GameType = Selected
    
    Select Case Selected
    
        Case 0
            Me.Hide
        
        Case 1
            Call Serve
        Case 2
            Call BeClient
    
    End Select
Else
    'PosSock.Close
    Call DoBusy(True)
End If

For i = 0 To optPlayType.UBound
    If optPlayType(i).Value = True Then
        PlayType = i
        Exit For
    End If
Next i

End Sub

Private Sub somcmdQuit_Click()

End

End Sub
