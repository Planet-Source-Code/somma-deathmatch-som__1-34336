Attribute VB_Name = "modMathsFunctions"
Option Explicit

'There are lots of old functions here. They date back to early 2001 when
'i made a program that moves images around using the mouse, which gradually
'evolved into what i have today. See the 'History of Deathmatch SOM' for
'the full story.
Public Type NewCoords
    X As Integer
    Y As Integer
End Type

Public Function ConvY(ByVal ConvertWhat As Integer, ByVal WhichForm As Object) As Single
'This is needed because VB measures everything from the top of the form
'instead of the bottom

ConvY = WhichForm.ScaleHeight - ConvertWhat

End Function

Public Function UnConvY(ByVal UnConvertWhat As Integer, ByVal WhichForm As Object) As Single
'And this undoes what the function above does
UnConvY = WhichForm.ScaleHeight - UnConvertWhat
End Function

Public Function Grad(ByVal YOrigin As Integer, ByVal Y2 As Integer, ByVal XOrigin As Integer, ByVal X2 As Integer) As Double
'Finds the gradient of any two points

'Stop divide by zero
If X2 - XOrigin <> 0 Then
    Grad = (Y2 - YOrigin) / (X2 - XOrigin)
    'If Grad = 1234567890 Then MsgBox "The gradient thing happened!"
Else
    'Returns a dummy value if X2-Xorigin = 0
    Grad = 1234567890
End If
End Function

Public Function Distance(ByVal X2 As Integer, ByVal x1 As Integer, ByVal Y2 As Integer, ByVal y1 As Integer) As Single
'The Distance Formula. Just plug in some numbers, and Tada!

Distance = Sqr((X2 - x1) ^ 2 + (Y2 - y1) ^ 2)

End Function

Public Function FindXOrigin(ByVal WhichImage As Object) As Single
'This function finds the X ordinate of the center of any object
FindXOrigin = WhichImage.Left + (WhichImage.Width / 2)

End Function

Public Function FindYOrigin(ByVal WhichImage As Object) As Single
'...And this one finds the Y ordinate (abcissa?)
FindYOrigin = WhichImage.Top + (WhichImage.Height / 2)

End Function

Public Function FindX(ByVal XOrigin As Integer, ByVal YOrigin As Integer, ByVal Length As Single, ByVal XDirection As Integer, ByVal YDirection) As Single
'Finds the X-ord of one end of a line
'using division of interval in given ratio formula

'This can be used to draw a line that is always the same length,
'regardless of where the mouse pointer(or whatever) is
    Dim Extra As Single
    Extra = XTRALen(Length, XOrigin, YOrigin, XDirection, YDirection)
If Extra <> Length Then
    FindX = ((Length * XDirection) + (Extra * XOrigin)) / (Length + Extra)
Else
    FindX = XOrigin
End If
End Function

Public Function FindY(ByVal XOrigin As Integer, ByVal YOrigin As Integer, ByVal Length As Single, ByVal XDirection As Integer, ByVal YDirection) As Single
'Ditto
Dim Extra As Single
Extra = XTRALen(Length, XOrigin, YOrigin, XDirection, YDirection)

FindY = ((Length * YDirection) + (Extra * YOrigin)) / (Length + Extra)

End Function


Public Function XTRALen(ByVal Length As Single, ByVal XOrigin As Integer, ByVal YOrigin As Integer, ByVal XDirection As Integer, ByVal YDirection As Integer) As Single
'Finds the extra distance between the mouse pointer is from the line
'that will be drawn.
XTRALen = Distance(XDirection, XOrigin, YDirection, YOrigin) - Length

End Function

Public Sub DrawLine(ByVal XOrigin As Long, ByVal YOrigin As Long, ByVal X As Long, ByVal Y As Long)
'Draws the line
'Change 'Me' when neccessary

'Me.Cls
'Me.Line (XOrigin, YOrigin)-(FindX(XOrigin, UnConvY(YOrigin, Me), LineLen, x, ConvY(y, Me)), UnConvY(FindY(XOrigin, UnConvY(YOrigin, Me), LineLen, x, ConvY(y, Me)), Me)), LineColour

End Sub

'Public Function AngleMove(ByVal XOrigin As Long, ByVal YOrigin As Long, ByVal Length As Integer, ByVal Gradient As Double) As NewCoords
''Not done yet
'Dim XDir As Single
'Dim YDir As Single
'
'Dim i As Integer
'
'Gradient = GRound(Gradient, 2)
'For i = 1 To 100
'    If Gradient * i = Int(Gradient) Then
'        XDir = i
'        YDir = Gradient * i
'        Exit For
'    End If
'Next i
'
'AngleMove.X = FindX(XOrigin, YOrigin, Length, XDir, YDir)
'AngleMove.Y = FindY(XOrigin, YOrigin, Length, XDir, YDir)
'
'End Function


