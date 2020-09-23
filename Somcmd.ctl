VERSION 5.00
Begin VB.UserControl Somcmd 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   CanGetFocus     =   0   'False
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   DefaultCancel   =   -1  'True
   ScaleHeight     =   1125
   ScaleWidth      =   1695
   Begin VB.Timer TimeFade 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1320
      Top             =   600
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   465
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   1035
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Border 
      BorderColor     =   &H00BBEBFD&
      Height          =   615
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Somcmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Event Click()

Dim NormalColour As Long
Dim HoverColour As Long
Dim Hovering As Boolean

Dim FadeIn As Boolean

Dim SelSound As String
Dim LostSound As String
Dim UsePlaySound As Boolean

'Private Type TriColour
'    R As Byte
'    G As Byte
'    B As Byte
'End Type
'
'Dim LColour As TriColour

Private Sub lblC_Click()

RaiseEvent Click

End Sub

Private Sub lblC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Not Hovering Then
    Hovering = True
    Call SetFontColour
End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Makes sure it only happens *once*
If Not Hovering Then
    Hovering = True
    Call SetFontColour
End If

End Sub


Private Sub TimeFade_Timer()

Dim Diff As Long
Dim Direct As Integer

Diff = NormalColour - HoverColour

If Diff > 0 Then
    Direct = 10
ElseIf Diff < 0 Then
    Direct = -10
Else
    TimeFade.Enabled = False
    Exit Sub
End If

If Hovering Then
    Direct = Direct * -1
End If

lblC.ForeColor = lblC.ForeColor + Direct

End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

RaiseEvent Click

End Sub

Private Sub UserControl_Click()

RaiseEvent Click

End Sub

Private Sub UserControl_Initialize()

lblC.Left = 0
With Border
    .Top = 0
    .Left = 0
    .Visible = False
End With

NormalColour = &HFF&
HoverColour = &HFF00&

Call SetFontColour

'With LColour
'    .B = 0
'    .R = 0
'    .G = 255
'End With

End Sub

Property Let Caption(What As String)

lblC.Caption = What

End Property

Property Get Caption() As String

Caption = lblC.Caption

End Property

Property Let BackColour(Colour As OLE_COLOR)

UserControl.BackColor = Colour

End Property

Property Get BackColour() As OLE_COLOR

BackColour = UserControl.BackColor

End Property

Property Let Alignment(Align As AlignmentConstants)

lblC.Alignment = Align

End Property

Property Get Alignment() As AlignmentConstants

Alignment = lblC.Alignment

End Property

Property Let FontType(WFont As String)

lblC.Font = WFont

End Property

Property Get FontType() As String

FontType = lblC.Font

End Property

Property Let FontBold(IsIt As Boolean)

lblC.FontBold = IsIt

End Property

Property Get FontBold() As Boolean

FontBold = lblC.FontBold

End Property

Property Let FontColour(Colour As OLE_COLOR)

NormalColour = Colour

End Property

Property Get FontColour() As OLE_COLOR

FontColour = NormalColour

End Property

Property Let FontSize(Size As Integer)

lblC.FontSize = Size

End Property

Property Get FontSize() As Integer

FontSize = lblC.FontSize

End Property

Property Let MouseOverColour(Colour As OLE_COLOR)

HoverColour = Colour

End Property

Property Get MouseOverColour() As OLE_COLOR

MouseOverColour = HoverColour

End Property

Property Let Fade(DoesIt As Boolean)

FadeIn = DoesIt

End Property

Property Get Fade() As Boolean

Fade = FadeIn

End Property

Property Let SoundHover(Sound As String)

SelSound = Sound

End Property

Property Get SoundHover() As String

SoundHover = SelSound

End Property

Property Let SoundSHover(Sound As String)

LostSound = Sound

End Property

Property Get SoundSHover() As String

SoundSHover = LostSound

End Property

'This is needed because in Positions, directx is used to play
'sounds. This tells it to use the 'playsound' sub rather than
'sndplaysound
Property Let UsePlaySoundSub(YesNo As Boolean)

UsePlaySound = YesNo

End Property

Property Get UsePlaySoundSub() As Boolean

UsePlaySoundSub = UsePlaySound

End Property


Property Let BorderColour(Colour As OLE_COLOR)

Border.BorderColor = Colour

End Property

Property Get BorderColour() As OLE_COLOR

BorderColour = Border.BorderColor

End Property

Property Let BorderWidth(Width As Byte)

If Width > 0 Then
    Border.Visible = True
    Border.BorderWidth = Width
Else
    Border.Visible = False
End If

End Property

Property Get BorderWidth() As Byte

With Border

    If .Visible Then
        BorderWidth = .BorderWidth
    Else
        BorderWidth = 0
    End If

End With

End Property

Property Let BorderShape(Style As ShapeConstants)

Border.Shape = Style

End Property

Property Get BorderShape() As ShapeConstants

BorderShape = Border.Shape

End Property

Property Let BorderOpaque(YesNo As Boolean)

Border.BackStyle = Abs(YesNo)

End Property

Property Get BorderOpaque() As Boolean

BorderOpaque = Border.BackStyle

End Property

Property Let BorderFillColour(Colour As OLE_COLOR)

Border.BackColor = Colour

End Property

Property Get BorderFillColour() As OLE_COLOR

BorderFillColour = Border.BackColor

End Property

Property Let TextPos(Pos As Integer)

lblC.Top = Pos

End Property

Property Get TextPos() As Integer

TextPos = lblC.Top

End Property



Private Sub SetFontColour()

Dim PlayThis As String
If Hovering Then

    lblC.ForeColor = HoverColour
    PlayThis = SelSound
    
Else

    lblC.ForeColor = NormalColour
    PlayThis = LostSound

End If

If FileExists(PlayThis) Then
    If UsePlaySound Then
        Call PlaySound(PlayThis)
    Else
        Call sndPlaySound(PlayThis, &H1)
    End If
End If

End Sub

Sub StopHover()

If Hovering = True Then
    Hovering = False
    Call SetFontColour
End If

End Sub

Private Function FileExists(FileName As String) As Boolean
'Checks if a file exists. There *has* to be
'a better way of doing this...
If FileName <> "" Then
    Dim CheckThis As String
    On Error Resume Next
    
    CheckThis = Dir(FileName)
    If CheckThis = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End If
End Function


'Property Let ColourRed(Num As Byte)
'
'LColour.R = Num
'Call ConvColour
'
'End Property
'
'Property Get ColourRed() As Byte
'
'ColourRed = LColour.R
'
'End Property
'
'Property Let ColourBlue(Num As Byte)
'
'LColour.B = Num
'Call ConvColour
'
'End Property
'
'Property Get ColourBlue() As Byte
'
'ColourBlue = LColour.B
'
'End Property
'
'Property Let ColourGreen(Num As Byte)
'
'LColour.G = Num
'Call ConvColour
'
'End Property
'
'Property Get ColourGreen() As Byte
'
'ColourGreen = LColour.G
'
'End Property
'
'Sub ConvColour()
'
'With LColour
'    lblC.ForeColor = RGB(.R, .G, .B)
'End With
'
'End Sub

Private Sub UserControl_Resize()

With UserControl
    lblC.Width = .Width
    lblC.Height = .Height
    Border.Width = .Width
    Border.Height = .Height
End With

End Sub
