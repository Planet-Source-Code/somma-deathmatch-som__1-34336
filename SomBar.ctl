VERSION 5.00
Begin VB.UserControl ctrSomBar 
   BackColor       =   &H80000007&
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4275
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   285
   Begin VB.Label lblBar 
      BackColor       =   &H000000C0&
      Height          =   315
      Left            =   390
      TabIndex        =   1
      Top             =   390
      Width           =   375
   End
   Begin VB.Label lblBack 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "ctrSomBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Its probably been done before many times, but this is my
'custom progress bar, which will have more functionality than
'the crappy one that comes with VB

Option Explicit

Const Bord = 4

Dim BarValue As Single

Dim MinVal As Single
Dim MaxVal As Single
Dim OldMin As Single
Dim OldMax As Single

Dim Inverted As Boolean
Dim IsVert As Boolean
Dim IsSunken As Boolean


'Colour of the moving bar
Property Get BarColour() As OLE_COLOR

BarColour = lblBar.BackColor

End Property

Property Let BarColour(Colour As OLE_COLOR)

lblBar.BackColor = Colour

End Property


'Colour of the thing at the back
Property Get BackColour() As OLE_COLOR

BackColour = lblBack.BackColor

End Property

Property Let BackColour(Colour As OLE_COLOR)

lblBack.BackColor = Colour

End Property


'Whether or not the thing at the back has that border
'that makes it look engraved
Property Get Sunken() As Boolean

Sunken = IsSunken

End Property

Property Let Sunken(YesNo As Boolean)

IsSunken = YesNo
lblBack.BorderStyle = Abs(CInt(IsSunken))

End Property


'The value of the bar
Property Get Value() As Single

Value = BarValue

End Property

Property Let Value(Num As Single)

BarValue = Num
Call Update

End Property


'The minumum value
Property Get Min() As Single

Min = MinVal

End Property

Property Let Min(Num As Single)

Call SaveOld
MinVal = Num
Call Update

End Property


'The maximum value
Property Get Max() As Single

Max = MaxVal

End Property

Property Let Max(Num As Single)

Call SaveOld
MaxVal = Num
Call Update

End Property


'Whether the bar will go from right to left (top to bottom)
Property Get Invert() As Boolean

Invert = Inverted

End Property

Property Let Invert(YesNo As Boolean)

Inverted = YesNo
Call MoveBar

End Property

'Whether bar is vertical or not
Property Get Vertical() As Boolean

Vertical = IsVert

End Property

Property Let Vertical(YesNo As Boolean)

IsVert = YesNo

'I've got a swap function, but the values wont change
'even though they are passed by reference *stupid VB*
Dim Temp As Long
Temp = Height
Height = Width
Width = Temp

Call MoveBar
    
End Property

Private Sub UserControl_Resize()

'Makes sure the control doesn't crash if it's too small
If ScaleHeight <= Bord Then
    ScaleHeight = Bord + 1
End If

If ScaleWidth <= Bord Then
    ScaleWidth = Bord + 1
End If

'Resizes both the labels
With lblBack
    .Top = 0
    .Left = 0
    .Height = ScaleHeight
    .Width = ScaleWidth
    
    If Not IsVert Then
        lblBar.Height = .Height - Bord
    Else
        lblBar.Width = .Width - Bord
    End If
    
    lblBar.Left = .Left + Bord / 2
    lblBar.Top = .Top + Bord / 2
End With

Call MoveBar

End Sub

Sub Update()

Call ChkRange
Call MoveBar

End Sub

Sub SaveOld()

'Stores the values of the min and max
'Used in ChkRange if an invalid min or max
'value is entered
OldMin = MinVal
OldMax = MaxVal

End Sub

Sub MoveBar()

If BarValue > MinVal Then

    'Calculate what percentage of the bar to fill up
    Dim Percentage As Single
    Percentage = (BarValue - MinVal) / (MaxVal - MinVal)
    
    lblBar.Visible = True
    
    If Not IsVert Then
    
        'Moves the inside label to the other side of the bar to invert it
        lblBar.Width = GRound((Percentage) * (lblBack.Width - Bord))
        If Inverted Then
            lblBar.Left = lblBack.Left + lblBack.Width - lblBar.Width - Bord / 2
        Else
            lblBar.Left = lblBack.Left + Bord / 2
        End If
        
    Else
    
        lblBar.Height = GRound((Percentage) * (lblBack.Height - Bord))
        If Inverted Then
            lblBar.Top = lblBack.Top + lblBack.Height - lblBar.Height - Bord / 2
        Else
            lblBar.Top = lblBack.Top + Bord / 2
        End If
        
    End If
    
Else

    'Inside label disappears if value=min value
    lblBar.Visible = False
    
End If

End Sub

Sub ChkRange()

'Checks if the min value is greater than max
If Min > Max Then

    MsgBox "Max value must be greater than or equal to the Min value", vbInformation, "Invalid property value"
    
    'Restores old values
    MinVal = OldMin
    MaxVal = OldMax
    
End If


'Checks if the value is outside the range
If BarValue >= Max Then
    BarValue = Max
ElseIf BarValue <= Min Then
    BarValue = Min
End If

End Sub
