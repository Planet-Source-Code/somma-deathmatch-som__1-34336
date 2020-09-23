VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl LotsaSocks 
   BackColor       =   &H00BBEBFD&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   465
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   465
   Begin VB.Timer TheWaiter 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   -120
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "LotsaSocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This user control is simply a way of getting around some
'stupid thing with how the winsock sends data in DOS based
'platforms, such as win98,95, and ME

'Dim NumSocks As Byte
Dim CurData As String
Dim CurSock As Byte
Dim SockConnected() As Boolean

'Dim WaitedLongEnough As Boolean
Dim LastInd As Byte
Dim Serving As Boolean

Dim TheError As String

'Event ClientConnected()
'Event ServeConnected()
Event Connected(Server As Boolean)
Event DataArrive(Data As String)

Property Let Amount(Num As Byte)

If Num > 0 Then
    'NumSocks = Num
    Call LoadSocks(Num)
Else
    MsgBox "Please enter a number more than 0"
End If


End Property

Property Get Amount() As Byte

Amount = Sock.UBound

End Property

Property Get WhatError() As String

'This cannot be modified by user, so
'no let.
WhatError = TheError

End Property

Sub SendStuff(ByVal Data As String)

On Error GoTo LostConnect

Sock(CurSock).SendData (Data)

CurSock = CurSock + 1
If CurSock > Sock.UBound Then
    CurSock = 0
End If

Exit Sub

LostConnect:
TheError = Err.Description

End Sub

Sub ClientConnect(Host As String, FirstPort As Integer)

Dim i As Byte

TheError = ""

On Error GoTo NoConnect

For i = 0 To Sock.UBound
    Call Sock(i).Close
    Call Sock(i).Connect(Host, FirstPort + i)
Next i

Exit Sub

NoConnect:
TheError = Err.Description

End Sub

Sub ServerListen(FirstPort As Integer)

Dim i As Byte

TheError = ""
On Error GoTo NoServe

For i = 0 To Sock.UBound
    With Sock(i)
        .Close
        .LocalPort = FirstPort + i
        .Listen
    End With
Next i

Exit Sub

NoServe:
TheError = Err.Description

End Sub

Private Sub BeginWait()

'WaitedLongEnough = False
TheWaiter.Enabled = True

End Sub

Private Sub ResetConnected()

ReDim SockConnected(Sock.UBound)

End Sub

Private Sub LoadSocks(NewNum)

'Loads/unloads the appropriate number of winsock controls
Dim i As Byte

If NewNum > Sock.UBound Then
    For i = Sock.UBound + 1 To NewNum
        Load Sock(i)
    Next i
Else
    For i = NewNum + 1 To Sock.UBound
        Unload Sock(i)
    Next i
End If

Call ResetConnected

End Sub

Private Sub Sock_Connect(Index As Integer)

If Index = 0 Then
    Call BeginWait
End If

If Index = Sock.UBound Then
    TheWaiter.Enabled = False

    RaiseEvent Connected(False)
End If

LastInd = Index

End Sub

Private Sub Sock_ConnectionRequest(Index As Integer, ByVal RequestID As Long)

Dim i As Byte
Dim AllConnected As Boolean

Call BeginWait
LastInd = Index

With Sock(Index)
    If .State <> sckClosed Then
        .Close
    End If
    
    .Accept RequestID
    
    SockConnected(Index) = True
    
End With


AllConnected = True
For i = 0 To UBound(SockConnected)
    If SockConnected(i) = False Then
        AllConnected = False
        Exit For
    End If
Next i

'AllConnected = True
'For i = 0 To Sock.UBound
'    If Sock(i).State = sckConnected Then
'        AllConnected = False
'        Exit For
'    End If
'Next i

'It will say its connected if all the socks have found a
'connection, or if at least one sock is connected and its waited
'for more than 1 second for the rest of them.


If AllConnected Then
    TheWaiter.Enabled = False
    RaiseEvent Connected(True)
End If

End Sub

Private Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)

Sock(Index).GetData CurData
RaiseEvent DataArrive(CurData)

End Sub

Private Sub TheWaiter_Timer()

Call LoadSocks(LastInd)
RaiseEvent Connected(Serving)
TheWaiter.Enabled = False

End Sub

Private Sub UserControl_Initialize()

Call ResetConnected

End Sub
