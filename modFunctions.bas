Attribute VB_Name = "modFunctions"
Public Function ConvY(ByVal ConvertWhat As Integer, ByVal WhichForm As Object) As Single
'This is needed because VB measures everything from the top of the form
'instead of the bottom

ConvY = WhichForm.ScaleHeight - ConvertWhat

End Function

Public Function UnConvY(ByVal UnConvertWhat As Integer, ByVal WhichForm As Object) As Single
'And this undoes what the function above does
UnConvY = WhichForm.ScaleHeight - UnConvertWhat
End Function

Public Function Grad(ByVal YOrigin As Integer, ByVal Rise As Integer, ByVal XOrigin As Integer, ByVal Run As Integer) As Single
'Finds the gradient of any two points

Grad = (Rise - YOrigin) / (Run - XOrigin)

End Function

Public Function Distance(ByVal X2 As Integer, ByVal X1 As Integer, ByVal Y2 As Integer, ByVal Y1 As Integer) As Single
'The Distance Formula

Distance = Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2)

End Function

