Attribute VB_Name = "GoodRound"

Function GRound(ByVal RoundNum As Double, Optional DecPlaces As Integer = 0) As Double
'This is a 'patched' round function, which always rounds
'**UP** when the number is something .5

RoundNum = RoundNum * 10 ^ DecPlaces
If IsNumeric(RoundNum) = True Then
    If RoundNum > 0 Then
        If RoundNum >= Int(RoundNum) + 0.5 Then
            GRound = Int(RoundNum) + 1
        Else
            GRound = Int(RoundNum)
        End If
    Else
        If RoundNum <= Int(RoundNum) - 0.5 Then
            GRound = Int(RoundNum) - 1
        Else
            GRound = Int(RoundNum)
        End If
    End If
End If

GRound = GRound / 10 ^ DecPlaces

End Function

Public Function RoundUp(ByVal Number As Double, Optional DecPlaces As Integer = 0) As Long

If Int(Number) <> Number Then
    Number = Number * 10 ^ DecPlaces
    RoundUp = (Int(Number) + 1) / 10 ^ DecPlaces
Else
    RoundUp = Number
End If

End Function

Public Function RoundClosest(ByVal Number As Double, ByVal ClosestNum As Long) As Double

'Rounds number to closest whatever.
'e.g round to closest 360 for some angle thing

Number = Number / ClosestNum
Number = GRound(Number)
RoundClosest = Number * ClosestNum

End Function

'Public Function SomRand(ByVal High As Double, ByVal Low As Double, Optional DecPlaces As Integer = 0) As Double
''High = High * 10 ^ DecPlaces
''Low = Low * 10 ^ DecPlaces
'
'DoEvents
'
''If NoRandomize = False Then
'    Randomize
''End If
'
'This is really bad. Can *you* figure out why? :)
'SomRand = GRound(((High - Low) * Rnd) + Low, DecPlaces)
'
'End Function

Public Function SomRand(ByVal High As Double, ByVal Low As Double, Optional DecPlaces As Integer = 0, Optional NoRandomize As Boolean) As Double
'High = High * 10 ^ DecPlaces
'Low = Low * 10 ^ DecPlaces

DoEvents

If NoRandomize = False Then
    Randomize
End If
'SomRand = GRound(((High - Low) * Rnd) + Low, DecPlaces)

High = High * 10 ^ DecPlaces
Low = Low * 10 ^ DecPlaces

SomRand = (Int(((High - Low + 1) * Rnd) + Low)) / 10 ^ DecPlaces

End Function

