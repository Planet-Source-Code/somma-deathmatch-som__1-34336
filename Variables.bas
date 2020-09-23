Attribute VB_Name = "modStuff"
Option Explicit

'API functions
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'The maximum number of shots on screen at once
Public Const MaxShots = 50

'What data file to read from
Public Const SetFile = "data41.pos"

'Constants for saving to the registry
Public Const AppName = "Positions"
Public Const SetSec = "Settings"

'The [DeathMatch SOM] theme brackets
Public Const SBrack = "["
Public Const EBrack = "]"

'The program seems to need the height of the menu (or something) when
'resizing. Very strange, because it's not needed when I do collision
'detection with the bottom of the window
Public Const AddHeight = 40

'A universal error message for invalid settings
'Contains chr(13)'s so can't be a constant
Public SetError As String

'A (perhaps) better way of storing settings
Public Type SettingsBetter
    SetDesc As String
    SettingData As String
End Type

'Specific settings from the data file for the 'shots'
'or any other thing that goes in *lists*
Public Type ShotSettings
    Setting() As SettingsBetter
End Type

'See the frmoptions code
Public Type NumsOnly
    SelPosition As Byte
    OldText As String
End Type

'Used for damage/armour mods
Public Type PlayerPowerUp
    Bonus As Integer
    TimeLeft As Long
End Type

'Stores settings for each explosion. Now it's also used for shot trails,
'as they have exactly the same properties
Type Kaboom
    Wait As Integer
    X As Integer
    Y As Integer
End Type

'Settings for the two players
Public Type Face

    'Co-ordinates of player
    X As Single
    Y As Single
    
    'Hitpoints remaining
    HP As Integer
    
    'For HP decay if over 100% HP
    HPTime As Integer
    
    'Pretty Self-explanatory. The 'faces' were originally intended to
    'be rectangular, but they are circular now
    Height As Integer
    Width As Integer
    
    'Angle player is facing
    Degrees As Integer
    
    'These are not needed anymore, because I am using GetKeyStates
    '-------------------------------------------------------------
'    'Whether player has held down shoot button
'    HoldShot As Boolean
'
'    'Which way player is moving (if at all)
'    UpDown As String
'
    '-------------------------------------------------------------
    
    'Shot Delay time
    ShotWait As Byte

    'The current shot type selected
    ShotCurrent As Integer
    
    PlusDam As PlayerPowerUp   'Damage bonus given by powerup
    PlusArmour As PlayerPowerUp 'Armour bonus
    
    'The last thing that hurt the player
    LastHurt As Integer
    
    'Name of player (used to be PlayerName)
    Name As String
    
    'Keys that the player uses
    PKeys(6) As Integer
    
    'The weapons player uses
    Weps() As Integer
    
    'Class characteristics
    ClassArm As Integer
    ClassDam As Integer
    ClassHP As Single
    ClassShotSize As Integer
    ClassShotSpeed As Single
    ClassSpeed As Single
    ClassShotAccel As Single
    
End Type

'Data on each active shot
Public Type ShootingNow

    'The type of shot it is
    What As Integer
    
    'Who it belongs to
    WhichFace As Byte
    
    'For flowers: sets a new speed to use for the shot
    NewSpeed As Integer
    
    'Which direction its moving
    Degrees As Integer
    
    'Co-oridinates of shot
    ShotX As Single
    ShotY As Single
    
    'When it expires (if at all)
    Expire As Integer
    
    'Speed to add on to original speed (for acceleration)
    AddSpeed As Single
    
End Type

'For dragging stuff
Public Type DragIt
    XStart As Long
    YStart As Long
    Dragging As Boolean
End Type

'Stores stuff on active powerups. I used to use the kaboom type
'for this, but I added another property for the AI to use
Private Type PowerUpNow
    What As Integer
    X As Integer
    Y As Integer
        
    'Used by the power-up finding AI function
    '1 for each player (both players might be AI)
    GoForChance(1) As Single
End Type

'The players
Public Face(1) As Face

'Not used anymore because of GetKeyStates
'Public DegreeChange(1) As Integer

'Explosions
Public Kaboom(Int(MaxShots / 3)) As Kaboom

'Shot trails
Public Trails(MaxShots) As Kaboom

'The type of game being run
Public GameType As Byte

'Whether current machine is the server
Public IsServer As Boolean

'Shots currently active
Public Shooting(MaxShots) As ShootingNow

'Arrays of settings
'------------------

'Main settings
Public BSettings() As SettingsBetter

'Weapon data
Public ShotSets() As ShotSettings

'Power-up data
Public PowerUps() As ShotSettings

'Class data
Public Classes() As ShotSettings
'----------------

'Active powerups
Public PowerUpNow() As PowerUpNow

'The names of the two players
'Public PlayerName(1) As String

'Whether player is AI
Public IsAI(1) As Boolean

'Stores which AI states the user has enabled
Public PlayerAI(1, 5) As Boolean

'The player being 'processed' by the 'Enter your name' dialog.
Public ProcessPlayer As Byte

'A variable used in displaying a list of settings in frmshotoptions
'We can re-use the form for any set of options which uses the shotsettings type.
'this meanspowerups, weapons, and classes, although
'i've decided not to let the user edit classes from program
Public ChangeWhat() As ShotSettings

'When the SetErr procedure runs, the game will not run because
'there is 'baddata'.
Public BadData As Boolean

'Distance moved when moving at 90º angles
Public Hor As Single

'Found by pythagorus theorm. Horizontal and vertical distance moved when
'moving diagnally
Public Diag As Single

'The default height and width of the main window
Public DefHeight As Integer
Public DefWidth As Integer

'What files to use for the player sprites (pictures)
Public FaceSprites(1) As String

'Where to look for weapon sprites
Public ShotSprites As String

'Where to look for power-up sprites
Public PowerSprites As String

'Either use classes (0) or free-play (1). There might be more playtypes in the
'future, so this is not a boolean
Public PlayType As Byte

'The default HP before class characteristics are taken into account.
'Although this comes straight from the bsettings array, it is used so
'much throughout the program that it gets a variable on its own
Public StartHP As Integer

'What keys are being pressed. It is also used by the AI and
'for network play
Public PressKeys() As Boolean

'Whether there is a picture as the background
Public UsingBGPick As Boolean

