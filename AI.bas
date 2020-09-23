Attribute VB_Name = "modAI"
Option Explicit

'AI Constants, just to make it easier for myself

Public Const TLeft = "TurnLeft"
Public Const TRight = "TurnRight"
Public Const MUp = "MoveUp"
Public Const MDown = "MoveDown"
Public Const ShootAI = "Shoot"
Public Const ChUpAI = "ChUp"
Public Const ChDownAI = "ChDown"

Const Pi = 3.1415927

Dim Straight As Integer


Public Function DegToRad(ByVal Degree As Single) As Single

'Converts degrees entered, since VB calculates in Radian mode
'The constant is pi/180
DegToRad = Degree * (Pi / 180) '0.0174532

End Function

Public Function RadToDeg(ByVal Degree As Single) As Single

RadToDeg = (Degree / Pi) * 180

End Function

Public Function AIRan(ByVal WhichFace As Byte) As String

Dim RanNum As Byte
Dim Mines As Boolean

'Just a stupid Random AI
'It plants mines if it has nothing else to do,
'given that it has mines in its weapon selection
'and that the 'Plants Mines' option was selected

Mines = PlayerAI(WhichFace, 5)
RanNum = SomRand(7, 1)

If Not Mines Or RanNum < 6 Then
    Select Case RanNum
        Case 1: AIRan = TRight
        Case 2: AIRan = TLeft
        Case 3: AIRan = MDown
        Case 4: AIRan = MUp
        Case 5: AIRan = ShootAI
        Case 6: AIRan = ChDownAI
        Case 7: AIRan = ChUpAI
    End Select
Else
    AIRan = ChooseWep(WhichFace, True)
End If

End Function

Public Function AIMove(ByVal WhichFace As Byte) As String
'NOT DONE YET!!!!
'Nah, this doesn't work anyway

'Dim i As Integer
'Dim DontDoThis(315) As Boolean
'
'Dim CenterX As Integer
'Dim CenterY As Integer
'
'CenterX = Face(WhichFace).X + Face(WhichFace).Width / 2
'CenterY = Face(WhichFace).Y + Face(WhichFace).Height / 2
'
'For i = LBound(Shooting) To UBound(Shooting)
'    If Shooting(i).What <> "" And Shooting(i).WhichFace <> WhichFace Then
'        If Distance(CenterX, Shooting(i).ShotX, CenterY - Hor, Shooting(i).ShotY) <= Face(WhichFace).Width / 2 Then
'            DontDoThis(0) = True
'        End If
'        If Distance(CenterX - Diag, Shooting(i).ShotX, CenterY - Diag, Shooting(i).ShotY) <= Face(WhichFace).Width / 2 Then
'            DontDoThis(45) = True
'        End If
'        If Distance(CenterX + Hor, Shooting(i).ShotX, CenterY, Shooting(i).ShotY) <= Face(WhichFace).Width / 2 Then
'            DontDoThis(90) = True
'        End If
'        If Distance(CenterX + Diag, Shooting(i).ShotX, CenterY + Diag, Shooting(i).ShotY) <= Face(WhichFace).Width / 2 Then
'            DontDoThis(135) = True
'        End If
'        If Distance(CenterX, Shooting(i).ShotX, CenterY + Hor, Shooting(i).ShotY) <= Face(WhichFace).Width / 2 Then
'            DontDoThis(180) = True
'        End If
'        If Distance(CenterX - Diag, Shooting(i).ShotX, CenterY + Diag, Shooting(i).ShotY) <= Face(WhichFace).Width / 2 Then
'            DontDoThis(225) = True
'        End If
'        If Distance(CenterX - Hor, Shooting(i).ShotX, CenterY, Shooting(i).ShotY) <= Face(WhichFace).Width / 2 Then
'            DontDoThis(270) = True
'        End If
'        If Distance(CenterX - Diag, Shooting(i).ShotX, CenterY - Diag, Shooting(i).ShotY) <= Face(WhichFace).Width / 2 Then
'            DontDoThis(315) = True
'        End If
'    End If
'Next i
'
'For i = 0 To UBound(DontDoThis) Step 45
'    If DontDoThis(i) = False Then
'        If Face(WhichFace).Degrees = i Then
'            AIMove = MUp
'        Else
'            AIMove = ClosestTurn(WhichFace, i)
'        End If
'    End If
'Next i


End Function

Public Function AIDodge(ByVal WFace As Byte) As String

'How the hell am i going to do this...?

End Function

Public Function AIShoot(ByVal WhichFace As Byte) As String
'My vain attempt at writing an AI
'Damn, IT WORKS!!! Yay! Joy to the world!

'Anyway, this is probably the most complicated AI function here
'Basically, what it does is it checks if the opponent is at an angle
'where the AI can hit it, i.e at a 45º increment, plus or minus shot
'spreading
Dim TestGrad As Long
Dim i As Integer
Dim LeftRight As Integer

Dim OpX As Single
Dim OpY As Single
Dim AIX As Single
Dim AIY As Single
'Dim OpHeight As Integer
Dim OpWidth As Integer
'Dim GradStore As Double
'
'Dim GradLeftTop As Double
'Dim GradLeftBot As Double
'Dim GradRightTop As Double
'Dim GradRightBot As Double

Dim Angle As Single
Dim AnglePlus As Single
Dim AngleMinus As Single
Dim RadI As Single
Dim Dist As Single
Dim TheGrad As Single
Dim AngleSpread As Single
Dim ShotSize As Single
Dim CheckPlus As Integer

'Stores position of opponent
OpX = Face(FindOp(WhichFace)).X
OpY = ConvY((Face(FindOp(WhichFace)).Y), frmPositions)

'Finds size of opponent. (the .height property is obselete because the
'players are circles
'OpHeight = Face(FindOp(WhichFace)).Height / 2
OpWidth = Face(FindOp(WhichFace)).Width / 2

With Face(WhichFace)
    'Stores position of itself
    AIX = .X
    AIY = ConvY(.Y, frmPositions)
    
    'Finds spread and size of current weapon
    AngleSpread = DegToRad(ShotData(.Weps(.ShotCurrent), "SPREAD")) / 1.5
    ShotSize = ShotData(.Weps(.ShotCurrent), "RADIUS") / 2
End With

'Finds distance between players
Dist = Abs(Distance(AIX, OpX, AIY, OpY))

'Finds the gradient between the players
TheGrad = Grad(AIY, OpY, AIX, OpX)

'Finds angle between players. 1234567890 is a dummy value for when
'opx=aix, and gradient is undefined (division by 0). atn is arctangent,
'aka inverse tan, which is used to find an angle given the gradient in basic
'trigonometry
If TheGrad <> 1234567890 Then
    Angle = Atn(TheGrad)
Else
    Angle = 0
End If

'I have to do this because I made it hard for myself early in development
'Instead of letting facing right be 0º (like in trigonometry), I made facing
'upwards 0º, and made the angles increase clockwise. I have to convert the angle
'back 'trig friendly' form before performing trig functions on it
Angle = DegToRad(90) - Angle

'Finds the angle between AI and the edges of the opponent
AnglePlus = Angle + Tan(OpWidth / Dist) + AngleSpread
AngleMinus = Angle - Tan(OpWidth / Dist) - AngleSpread

'Test each 45º increment to see if the angle is between AnglePlus and
'AngleMinus.
For i = 0 To 135 Step 45

    'Finds the angle in radians
    RadI = DegToRad(i)
    
    'FindLeftRight just determines whether, for example, the angle is 45º or
    '225º (i.e 45º+180º)
    LeftRight = FindLeftRight(i, AIX, AIY, OpX, OpY)
    
    'Similar to FindLeftRight. Determines whether to test the angle as it is,
    'or test 180º + the angle.
    'It goes: If [we are testing 0º] AND ([the AI is above the Opponent] AND [the AI is to the left of the
    'opponent]) then we add the value of leftright to angle we are checking.
    'This would make it 180º
    'AND I CANT QUITE REMEMBER WHY I HAVE TO DO THIS...it seems to work without it
'    If i = 0 And ((AIY > OpY And AIX < OpX)) Then 'Or (AIY > OpY And AIX < OpX)) Then
'        CheckPlus = LeftRight
'    Else
'        CheckPlus = 0
'    End If
    
    'Checks if angle being tested is between AnglePlus and AngleMinus
    If (RadI + DegToRad(CheckPlus) > AngleMinus And RadI + DegToRad(CheckPlus) < AnglePlus) Then
        
        If Face(WhichFace).Degrees = i + LeftRight Then
            AIShoot = ShootAI
        Else
            AIShoot = ClosestTurn(WhichFace, i + LeftRight)
        End If

        Exit For
    End If
Next i

'>This is the old AI function
'All this tests if the opponent is at a 45º angle to the AI's face
'I can think of a more elegant way to do this, but I havn't figured out
'how to do it yet. It works, though, so im happy with it for now.
'(as you can tell, I wrote this sub a while ago when I didn't know much about VB :-))

'For i = 0 To 135 Step 45
'    'Sets which angle to test for
'    If i = 0 Or i = 180 Then
'        TestGrad = 1234567890
'    Else
'        TestGrad = Tan(DegToRad(90 - i))
'    End If
'
'    GradStore = Grad(OpY, AIY, OpX, AIX)
'    GradLeftTop = Grad(OpY + OpHeight, AIY, OpX, AIX)
'    GradLeftBot = Grad(OpY - OpHeight, AIY, OpX, AIX)
'    GradRightTop = Grad(OpY + OpHeight, AIY, OpX + OpWidth, AIX)
'    GradRightBot = Grad(OpY - OpHeight, AIY, OpX + OpWidth, AIX)
'
'    'If GRound(GradStore, 1) = 9 Then MsgBox GradStore
'    '(GradStore > 9 Or GradStore < 9)
'    If (OpX < AIX And i <> 0) Or (GradStore <> 9 And OpY < AIY And i = 0) Then '(GRound(OpX, -1) = GRound(AIX, -1) And OpY < AIY) Then
'        LeftRight = 180
'    Else
'        LeftRight = 0
'    End If
'
'    If (((TestGrad > GradRightBot And TestGrad < GradLeftTop) Or (TestGrad < GradRightTop And TestGrad > GradLeftBot) And LeftRight = 0) Or ((TestGrad < GradRightBot And TestGrad > GradLeftTop) Or (TestGrad > GradRightTop And TestGrad < GradLeftBot) And LeftRight = 180) And SameSign(GradLeftTop, GradRightTop)) Or ((Abs(GradStore) > 9 Or SameSign(GradLeftTop, GradRightTop) = False) And (i = 0 Or i = 180)) Then '(TestGrad >= (Grad(OpY - OpHeight, AIY, OpX - OpWidth, AIX)) And TestGrad <= Grad(OpY + OpHeight, AIY, OpX + OpWidth, AIX)) Then
'
'        If Face(WhichFace).Degrees = ConvDeg(i - LeftRight) Then
'            AIShoot = ShootAI
'        Else
'            AIShoot = ClosestTurn(WhichFace, ConvDeg(i - LeftRight))
'        End If
'        Exit For
'    End If
'Next i
    

End Function

Function FindLeftRight(ByVal Degrees As Single, X As Single, Y As Single, DestX As Single, DestY As Single) As Integer

'Explained in AIShoot function
If (X > DestX And Degrees <> 0) Or (Y > DestY And Degrees = 0) Then
    FindLeftRight = 180
Else
    FindLeftRight = 0
End If


End Function

Public Function AIFindPower(WhichFace As Byte) As String

'This is a basic chasing-type AI that moves towards a powerup
Dim i As Integer
Dim n As Integer

Dim Lowest As Integer
Dim LowestDeg As Integer
Dim ArePowerUps As Boolean

Dim XDist As Single
Dim YDist As Single

Dim OpNum As Byte
OpNum = FindOp(WhichFace)
Lowest = 30000

For i = LBound(PowerUpNow) To UBound(PowerUpNow)
    With PowerUpNow(i)
        If .What >= 0 Then
        
            'Dim OpDist As Single
            
            'OpDist = Distance(Face(OpNum).X, .X, Face(OpNum).Y, .Y)
            
            Dim TestThis As Single
            
            If GoForPower(WhichFace, OpNum, i) Then
                ArePowerUps = True
                'Figure out which direction to move in would be quickest
                For n = 0 To 315 Step 45
                    
                    Call MoveHowMuch("up", n, XDist, YDist)
                    TestThis = Distance(Face(WhichFace).X + XDist, .X, Face(WhichFace).Y + YDist, .Y)
                    
                    'Check to see who's closer to the power-up NOT DONE HERE
                    
                    'If (TestThis * 0.8) < OpDist Then
                        'ArePowerUps = True
                        
                    'Check if the current powerup is the closest, and if
                    'facing nº will get the face there in the shortest time.
                    If TestThis < Lowest Then
                        Lowest = TestThis
                        LowestDeg = n
                    End If
                    'End If
                    
                Next n
            End If
            
        End If
    End With
Next i
       
'Will move in the direction of LowestDeg if any GoForPowers were true
If ArePowerUps = True Then
    If Face(WhichFace).Degrees = LowestDeg Then
        AIFindPower = MUp
    Else
        AIFindPower = ClosestTurn(WhichFace, LowestDeg)
    End If
End If

End Function

Function GoForPower(WFace As Byte, OpFace As Byte, PInd As Integer) As Boolean

Dim Dist As Single
Dim OpDist As Single
Dim Chance As Single

With PowerUpNow(PInd)
    'Find the distance to the powerup, taking into account movement speeds
    Dist = Distance(Face(WFace).X, .X, Face(WFace).Y, .Y) / Face(WFace).ClassSpeed
    OpDist = Distance(Face(OpFace).X, .X, Face(OpFace).Y, .Y) / Face(FindOp(WFace)).ClassSpeed
    
    'Add a random factor to the way AI goes for powerups.
    'AI will go for a powerup even if the opponent is
    'a certain random percentage closer to it. Some they
    'might not go for even if they are closer.
    If .GoForChance(WFace) < 0 Then
        .GoForChance(WFace) = SomRand(1, 0.75, 2)
    End If
    
    If Dist * .GoForChance(WFace) <= OpDist Then
        GoForPower = True
    End If
    
End With

End Function

Public Function AIMoveBack(WhichFace As Byte) As String

If Distance(Face(WhichFace).X, Face(FindOp(WhichFace)).X, Face(WhichFace).Y, Face(FindOp(WhichFace)).Y) <= Face(WhichFace).Height / 2 + Face(FindOp(WhichFace)).Height / 2 Then
    AIMoveBack = MDown
End If

End Function

Public Function ChooseWep(WhichFace As Byte, Optional UseMines As Boolean) As String

'This is the AI's method of choosing the best shot to use.
'Note that the *AI does not cheat*: it must cycle through the
'shots just like a human player. It does NOT simply change
'its current shot to the best one straight away. Although this is
'not noticable in normal playing, it is if you use free-play and
'set the number of weapon slots to a high number

Dim Dist As Single
Dim BestSpeed As Integer
'Dim ShotSpeed As Integer
'Dim UseThisSpeed As Integer

Dist = Distance(Face(WhichFace).X, Face(FindOp(WhichFace)).X, Face(WhichFace).Y, Face(FindOp(WhichFace)).Y)
'ShotSpeed = RawSpeed * BSettings(SetIndMain("MOVESTRAIGHT")).SettingData
'Straight = BSettings(SetIndMain("MOVESTRAIGHT")).SettingData

BestSpeed = UseWhatSpeed(Dist, UseMines, WhichFace)

Dim i As Integer
Dim CurShot As Byte
'Dim Found(1) As Boolean
Dim FoundIt As Boolean
Dim ShotInd As Byte
Dim WepInd As Byte
'Dim Time As Integer

Dim NoGood As Boolean

Dim SDamLow As Integer
Dim SDamHigh As Integer
Dim SExLow As Integer
Dim SExHigh As Integer
Dim MinDam As Integer
Dim DamBonus As Single
Dim Acidic As Single
Dim Weaken As Single
Dim ROF As Integer
Dim DontDoMods As Boolean

Dim GoodExplode As Byte

Dim DamScore As Single
Dim ExDamScore As Single

Dim NegPos As Integer

Dim BestShot As Integer
Dim MostDam As Integer

Dim ShotScore() As Single
ReDim ShotScore(UBound(Face(WhichFace).Weps))

'Decide *at random* whether to search up or down for a matching weapon
Do
    NegPos = SomRand(1, -1)
Loop Until NegPos <> 0

MinDam = Face(FindOp(WhichFace)).PlusArmour.Bonus
DamBonus = (100 + Face(WhichFace).PlusDam.Bonus) / 100

CurShot = Face(WhichFace).ShotCurrent

i = 0

Do
    'Search through each shot in the direction of NegPos
    'Note that it starts at the current shot
    WepInd = ConvShot(CurShot + (i * NegPos), UBound(Face(WhichFace).Weps))
    ShotInd = Face(WhichFace).Weps(WepInd)
        
    'Go through each shot and give each a rating
    'After that, it choses the shot with the highest rating
    
    SDamHigh = Int(ShotData(ShotInd, "DAM")) 'Int(.Setting(SetInd("DAM", .Setting)).SettingData)
    SDamLow = Int(ShotData(ShotInd, "DAMLOW")) 'Int(.Setting(SetInd("DAMLOW", .Setting)).SettingData)
    SExHigh = Int(ShotData(ShotInd, "EXHIGH")) 'Int(.Setting(SetInd("EXHIGH", .Setting)).SettingData)
    SExLow = Int(ShotData(ShotInd, "EXLOW")) 'Int(.Setting(SetInd("EXLOW", .Setting)).SettingData)
    
    'Damage and armour mod effects
    Weaken = (Int(ShotData(ShotInd, "DAMMODMIN")) + Int(ShotData(ShotInd, "DAMMODMAX"))) / 2
    Acidic = (Int(ShotData(ShotInd, "ARMMODMIN")) + Int(ShotData(ShotInd, "ARMMODMAX"))) / 2
    
    'Rate of fire
    ROF = ShotData(ShotInd, "WAIT")
    
'        If ((SDamHigh + SDamLow) / 6) * DamBonus > MinDam Or ((SExHigh + SExLow) / 6) * DamBonus > MinDam Then
'            ShotScore(WepInd) = ShotScore(WepInd) + 6
'        End If
'
'        If ((SDamHigh + SDamLow) / 2) * DamBonus > MinDam Or ((SExHigh + SExLow) / 2) * DamBonus > MinDam Then
'            ShotScore(WepInd) = ShotScore(WepInd) + 5
'        End If

    'Get scores for normal and explosive damage
    If SDamHigh > 0 Then
        DamScore = DamageScore((SDamHigh + SDamLow) / 2, MinDam, DamBonus, ROF)
    Else
        DamScore = 0
    End If
    ExDamScore = DamageScore((SExHigh + SExLow) / 2, MinDam, DamBonus, ROF)
    
'        If SDamHigh * DamBonus > MinDam Or SExHigh * DamBonus > MinDam Then
'            ShotScore(WepInd) = ShotScore(WepInd) + 1
'        End If
            
    'If GoodShot(ShotInd, Dist, UseMines, BestSpeed, Face(WhichFace).ClassShotSpeed) Then
        
    'End If
    
    'If IsNumeric(ShotSets(ShotInd).Setting(SetInd("EXPIRE", ShotSets(ShotInd).Setting)).SettingData) Then
    
    DontDoMods = False
    
    With Face(WhichFace)
        GoodExplode = GoodExplodeShot(ShotInd, Dist, UseMines, .ClassShotSpeed)
        If GoodExplode = 1 Then
            ShotScore(WepInd) = ShotScore(WepInd) + (5 * (DamScore + ExDamScore))
        ElseIf SDamHigh = 0 And GoodExplode = 2 Then
            'Makes sure it doesn't calculate mod effects for this weapon
            DontDoMods = True
        End If
        
        If GoodExplode = 0 Then
            ShotScore(WepInd) = ShotScore(WepInd) + GoodShot(ShotInd, Dist, UseMines, BestSpeed, .ClassShotSpeed, .ClassShotAccel) * (DamScore + ExDamScore) '4
        End If
        
        If Not DontDoMods Then
            ShotScore(WepInd) = ShotScore(WepInd) + ModEffectScore(Weaken, ROF)
            ShotScore(WepInd) = ShotScore(WepInd) + ModEffectScore(Acidic, ROF)
        End If
    End With
                

    i = i + 1
            
Loop Until (i > UBound(Face(WhichFace).Weps))    '(Found(0) = True Or Found(1) = True) Or i > UBound(ShotSets)

'If AI hasnt got the best shot selected, then it will keep selecting the
'next/prev weapon until it does. IT DOES NOT SIMPLY SET THE CURRENT SHOT
'TO THE BEST SHOT
If CurShot <> ModFindBiggest(ShotScore, CurShot, NegPos, WhichFace) Then
    If NegPos = 1 Then
        ChooseWep = ChUpAI
    Else
        ChooseWep = ChDownAI
    End If
End If
    
End Function

Function GoodShot(ByVal SInd As Integer, ByVal WDist As Single, UseMines As Boolean, ByVal BestSpeed As Integer, CBonus As Single, CAccelBonus As Single) As Single

'Checks if the shot is appropriate for the distance away from the
'other player.
'Function returns a score for that shot

Dim Speed As Single
Dim Accel As Single
Dim SpeedDif As Single

With ShotSets(SInd)
    
    'Get data from settings
    Accel = ShotData(SInd, "ACCEL") + CAccelBonus
    Speed = AvgSpeed(WDist, ShotData(SInd, "SPEED") * CBonus, Accel)
        
    'Find difference between the speed and the bestspeed
    SpeedDif = Speed - BestSpeed
    
    'Gives more value to faster shots
    If SpeedDif < 0 Then
        SpeedDif = SpeedDif / 5
    End If
    
    'Prevents divide by 0
    If SpeedDif = -1 Then
        SpeedDif = SpeedDif + 0.1
    End If

    'As speed -> bestspeed, speedif -> 0, and hence the closer the
    'shot's speed is to the best speed, the higher the score.
    GoodShot = (5 / (Abs(SpeedDif) + 1))
            

End With

End Function

Function DamageScore(ByVal AvgDam As Single, OpArmour As Integer, DamBon As Single, FireRate As Integer) As Single

'Gives a score for damage
DamageScore = 3 * ((AvgDam * DamBon) - (OpArmour * 1.5)) / FireRate
If DamageScore < -1 Then
    DamageScore = DamageScore / 5 '-1
End If

End Function


Function GoodExplodeShot(ByVal SInd As Integer, ByVal WDist As Single, ByVal UseMines As Boolean, ByVal CBonus As Single) As Byte

'Checks if the shot explodes within the given distance
'It is counted as a good shot if it does
Dim Expire As String
Dim ExplodeArea As Integer
Dim Speed As Single
Dim Accel As Single
Dim ExpireMax As String

    
'Get data from settings
Expire = ShotData(SInd, "EXPIRE")
ExplodeArea = ShotData(SInd, "EXAREA")
Speed = ShotData(SInd, "SPEED")
Accel = ShotData(SInd, "ACCEL")
ExpireMax = ShotData(SInd, "EXPIREHIGH")

'Checks if opponent is standing between the min and max of where
'the shot will expire.
If IsNumeric(Expire) And IsNumeric(ExpireMax) Then 'And Not UseMines Then

    'If Speed > 0 Then
    
        
        If (CInt(Expire) * AvgSpeed(WDist + ExplodeArea, Speed, Accel) * Hor) * CBonus <= WDist + ExplodeArea Then
            If (CInt(ExpireMax) * AvgSpeed(WDist - ExplodeArea, Speed, Accel) * Hor) * CBonus >= WDist - ExplodeArea Then
                GoodExplodeShot = 1
            Else
                GoodExplodeShot = 2
            End If
        End If
        
    'End If
End If

End Function

Function AvgSpeed(ByVal Dist As Single, ByVal Speed As Single, ByVal Accel As Single) As Single

'Find the average speed over a given distance
'uses v^2 = u^2 +2*a*s, where v is the final velocity, u the initial vel,
'a is acceleration, and s is the distance to be covered.
'Note that if a shot that will never cover the distance because the
'accel is too great, then v^2 < 0, and v cannot be found.
Dim VSquared As Single
VSquared = Speed ^ 2 + (2 * Accel * (Dist / Hor))

If VSquared >= 0 Then
    If Accel <> 0 Then
        AvgSpeed = (Speed + Sqr(VSquared))
    Else
        AvgSpeed = Speed
    End If
Else
    AvgSpeed = 0
End If

End Function

Function ModEffectScore(ByVal Effect As Single, FireRate As Integer) As Single

'Gives score based on the armour/damage mod effect
If Effect > 0 Then
    ModEffectScore = (Effect / FireRate) * 0.9
End If

End Function

Function ModFindBiggest(Scores() As Single, ByVal StartWhere As Integer, UpDown As Integer, ByVal Player As Byte) As Integer

'Function finds the largest number in the array

Dim i As Integer

ModFindBiggest = StartWhere

i = StartWhere

Do

    If Scores(i) > Scores(ModFindBiggest) Then
        ModFindBiggest = i
    End If
    i = ConvShot(i + UpDown, UBound(Face(Player).Weps))
    
Loop Until i = StartWhere
    

End Function


Function UseWhatSpeed(Dist As Single, ByVal UseMines As Boolean, ByVal WhichFace As Byte) As Single

'>Old code...
'Selects the shot speed based on distance to opponent
'It wont use Slow Shot if opponent is backing away
'Now it will *only* use the fast shots if opponent is backing away
'*Usemines is only true when this function is called from the random AI, and
'user has selected it to be an active AI function.


'Const RawSpeed = 2 'BSettings(SetIndMain("DEFSHOTSPEED")).SettingData
'Dim ShotDist As Integer
'
''Distance a shot with a speed of 2 will travel in one timer event
'ShotDist = RawSpeed * Hor

'If Not UseMines Then
'    If (Face(WhichFace).UpDown <> "up" Or Face(FindOp(WhichFace)).UpDown <> "down") Then
'        If Dist <= ShotDist * 6 Then
'            UseWhatSpeed = RawSpeed - 1
'        ElseIf Dist <= ShotDist * 15 Then
'            UseWhatSpeed = RawSpeed
'        Else
'            UseWhatSpeed = RawSpeed + 1
'        End If
'    Else
'        UseWhatSpeed = RawSpeed + 1
'    End If
'Else
'    UseWhatSpeed = 0
'End If

'New code!
'The best weapon is a shot that can reach the opponent in 'optimalmoves'
Const OptimalMoves = 4.5
Dim AddRun As Single
Dim OpDeg As Integer
Dim FaceDeg As Integer

FaceDeg = Face(WhichFace).Degrees

With Face(FindOp(WhichFace))

    OpDeg = .Degrees
    
    If (FaceDeg = OpDeg And PressKeys(WhichFace, 0)) Or (FaceDeg = ConvDeg(OpDeg + 180) And PressKeys(WhichFace, 1)) Then
        AddRun = .ClassSpeed
    End If
    
    UseWhatSpeed = ((Dist / Hor) / OptimalMoves) + AddRun

End With
        

End Function

Public Function ConvShot(Number As Integer, Optional UpBound As Integer = -1) As Integer

'For changing down from shot 0 (changes to last shot)
'or up from the last shot (changes to shot 0)

If UpBound = -1 Then
    UpBound = UBound(ShotSets)
End If

Do
    If Number < 0 Then
        ConvShot = UpBound + Number + 1
    ElseIf Number > UpBound Then
        ConvShot = Number - UpBound - 1
    Else
        ConvShot = Number
    End If
Loop Until ConvShot >= 0 And ConvShot <= UpBound

End Function

Public Function SameSign(Num1 As Double, Num2 As Double) As Boolean

'I didn't know the sgn function existed before, OK?
'If (Num1 > 0 And Num2 > 0) Or (Num1 < 0 And Num2 < 0) Or (Num1 = 0 Or Num2 = 0) Then
'    SameSign = True
'End If

If Sgn(Num1) = Sgn(Num2) Then
    SameSign = True
End If

End Function

Public Function ConvDeg(ByVal Degree As Integer) As Integer
'Makes all degrees between 0 and 360
Do Until Degree >= 0 And Degree < 360
    If Degree >= 360 Then
        Degree = Degree - 360
    ElseIf Degree < 0 Then
        Degree = 360 + Degree
    End If
Loop

ConvDeg = Degree
End Function

Public Function FindOp(ByVal WhichFace As Byte) As Byte

'This just finds the opponent of 'WhichFace'
'It is done *a lot*, so I put it into a function
If WhichFace = 0 Then
    FindOp = 1
ElseIf WhichFace = 1 Then
    FindOp = 0
End If

End Function

Public Function ClosestTurn(ByVal WhichFace As Byte, ByVal Degree As Integer) As String
'This function finds which direction (left or right) to turn would
'reach a given degree quicker

Dim Deg As Integer
Dim Turns(1) As Byte
'It's a shame that variables can't hold operators (+,-,*,/,etc)
'I wouldn't have to do 2 seperate loops...

Deg = Face(WhichFace).Degrees
Do
    Deg = Deg + 45
    Turns(0) = Turns(0) + 1
'    If Deg = 360 Then
'        Deg = 0
'    End If
    If ConvDeg(Deg) = Degree Then
        Exit Do
    End If
    

Loop

Deg = Face(WhichFace).Degrees
Do
    Deg = Deg - 45
    Turns(1) = Turns(1) + 1
'    If Deg = -45 Then
'        Deg = 315
'    End If
    If ConvDeg(Deg) = Degree Then
        Exit Do
    End If
Loop

'Finds which took less turns
If Turns(0) < Turns(1) Then
    ClosestTurn = TRight
ElseIf Turns(1) < Turns(0) Then
    ClosestTurn = TLeft
Else
    'If both the same, then select one at random
    Dim RanNum As Byte
    Randomize
    RanNum = Int(1 * Rnd)
    If RanNum = 0 Then
        ClosestTurn = TRight
    Else
        ClosestTurn = TLeft
    End If
End If
    

End Function
