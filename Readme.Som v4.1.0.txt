~~ This is probably a bit out of date. See the help file for more up-to-date stuff~~

~~ This file has not been spell-checked or grammar checked ~~
~~ Parts may also be incomplete (I'm only human, ok) ~~

A bit about Positions© (DeathMatch SOM©)
a game by ºSomº


CONTENTS
--------

-A History Lesson

-Getting Started
	The 'Enter your name' dialogs

-The Stats window

-Shooting
	The weapons (descriptions)

-Powerups *NEW STUFF IN HERE*

-The Player Classes *BRAND NEW!*

-The Almighty SOM-AI

-Controls

-Network Play *BRAND NEW!*

-The Options dialog

-Acknowledgements *NEW STUFF IN HERE*

-Contact details
	



<<A History Lesson>>
	'Positions', as the name suggests, started off as a simple experiment with moving an
image box (the picture was the circle that looked a bit like a face) around a window using 
mouse clicks (yep, back near the beginning of year 11 when I knew jack-all about 
Visual Basic). The program would draw a line of a set length in the direction of the mouse 
(which is a lot harder than it sounds because this involves quite a few maths formulas), 
and when the mouse was clicked, the image box would move that set length in that direction 
to the new position. Inspiring stuff, I know.

	At that time, I was thinking about making some sort of turn-based game, so I added
another image box in the program, and they would take turns moving. If a player pointed
the line in the direction of the other player, the line would 'lock on' to the other image
box and turn red. As you can tell, I wasn't really getting anywhere with my 'game'.
	
	And then, by some twist of fate, I did something that would forever change the future
of this humble program: I ADDED KEYBOARD COMMANDS TO THE GAME!!! (**!EXCITING!**). The image
boxes could now be moved up, down, left, **AND** right, all with just the keyboard. After a
while, I got rid of the mouse altogether, and 'Positions' became a completely keyboard-
controlled game.
	
	Later on, i added shooting to the game. It took me ages figuring out how to make more
than one shot move at a time, until, eventually, I realised that I needed a timer. Then came
some basic collision detection, so the shots would disappear when they hit the opponent 
(or the wall). Another important development was when I stopped using the 'keydown' event 
alone to make the 'faces' move, because of the half second pause before the keyboard starts
'repeating', and because changing the repeat speed changed the speed of the faces (the 
program now uses the timer and keyup event to move the faces).

	The first time sound was programmed in, I used the SndPlaySound API to do the job.
It was pretty crappy because it could only play one sound at once (I use DirectSound now,
of course). I also moved away from image boxes (they are not transparent, so you could see
the little black box around them if they got near each other), and started using the Bitblt
(does anyone know what that means?) API to draw stuff. 

	I also added a sub-program in one of the earlier versions which simply picked 
random moves, solely for debugging purposes, and ended up writing a full-blown AI which 
actually shoots *at* you, goes after power-ups, and even picks the best shot to use depending
on how far away the opponent is. It was another thing which took me ages to do, but I guess
it sort of payed off in the end.

	Anyway, enough of me reminiscing about my past. GO AND PLAY THE GODDAM GAME!

	Oh, and please remember that this program was written with less than 1 year's worth of
programming experience. I am not exactly a '1337 |-|4><0r' or anything like that, so please 
keep this in mind before playing the game or reading the sourcecode.

Footnotes
---------
	If you are really that desparate to read more about all the crap I went through to make 
this program work, read versions.txt.
	If the program confuses you, or if its doing disturbing things to your computer, then
too bad: you'll just have to wait until I get HTML help


--------------------------------------------------------

<<Getting Started>>

	Ok then, ignore that last footnote. Here is a quick guide on 'Positions'.

	If you got this program in a zip file or something, the first thing you should do is
put everything in its correct place. If they aren't there already, all sounds (*.wav) should
be in a folder named 'Sounds', and all pictures (*.bmp) should be in 'Pics' (unless they are 
skins made by someone else). These  subdirectories should be in the same folder as the game. 
The settings file 'data34.pos' (or data.pos in earlier versions) should also be in where the 
game is. If you don't do this, you will almost certainly get some sort of 'File not found' 
message, before the program unceremoniously kills itself. After this has been done, simply run 
the .exe, wait for it to load, and that's it. If the program decides to commit suicide anyway, 
make sure you tell me, because it could either be a bug, or you've done something stupid like 
mess up the settings file.

!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
	NOTE: IF YOU ARE USING WIN NT/2000/XP, THE GAME MIGHT RUN A LOT FASTER THAN ON A
DOS-BASED OS LIKE WIN 98/95. THIS IS SIMPLY BECAUSE THE 'GAMESPEED' SETTING IS SET TOO
LOW (THIS SETTING IS THE TIME BETWEEN EACH 'UNIT' OF GAME TIME). SIMPLY SETTING THIS TO
55 WILL BRING THE GAME BACK TO ITS INTENDED SPEED. OTHER WINDOWS USERS SHOULD SET
GAMESPEED TO 30.
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


<<The 'Enter your name' dialogs>>

	If everything is working, these will be the first things you see after the loading 
thing disappears. You can enter the name of the players here, as well as select what is 
controlling them. If you click on 'AI', note that the player's name defaults to 'Som-AI',
and the little checkboxes will be enabled. Use these checkboxes to choose which aspects 
of the AI will be active. Using the default setting will give you the best AI player, so 
you can disable some of them if you are just getting used to the game. Make sure at least
one player is human, unless you want to see two AI's battle it out (*yawn*).


<<The Stats Window>>

	The aim of the game is to simply kill the other player. You can keep track of both
your own hitpoints and you opponent's by looking at the Stats window at the bottom of the
game window. The progress bar (which I made myself) gives a visual representation of your
HP, while the big numbers (the beige-coloured ones) under it is your exact HP. Note that 
both the numbers and the bar will turn red when your HP is low, or turn into your player 
colour if it is above your starting HP.


	The numbers on the side (the red and green numbers which show 0 when the game begins,
and the smaller white numbers under them) show the bonuses given by power-ups. (See Powerups
for more details)

- The red number is the damage modifier of the player as a percentage. 
- The green number is the amount of armour the player has. Any *shot damage* inflicted on the
player gets reduced by the amount of armour. For example, an armour rating of 2 means all shots
deal 2 less damage. Consequently, an armour of -2 would increase the damage of all shots by 2.


<<Shooting>>
	
	Shooting is the most common method of dealing damage. When the shoot key is pressed,
the 'face' on the screen will shoot whatever your selected shot is, provided your *reload
time* has passed. Every shot is unique (so long as nobody has changed the default shot
settings), and their speed, damage, explosion damage, reload time, etc varies from shot 
to shot. When a shot hits your opponent or the wall, it will explode and deal *explosion
damage* to your opponent, if your opponent is within the *explosion range*.
A new shot feature in version 3.4.0 is *expiring shots*. When a shot 'expires',
it explodes as if it has hit a wall, and will deal its exposive damage if your opponent is 
within range. 

<<The Weapons>>
	
	The first fully working version of Positions had a choice of 4 shots. In fact, up
until v3.4.0, the 4 shots covered the whole variety of things I could do with the shots.
Besides, with all the shots being circles of exactly the same size, it would have gotten
pretty repetitive.
	And then came exploding shots in v3.3.12 (which really should have been the first v3.4), 
and expiring shots in v3.4.0, and suddenly I was able to add more shots and still maintain 
*variety*. Now there are a total of 8 shots (the 4 new ones were added within 2 versions 
of each other!), and each of them are still surprisingly unique.
	From v3.5.0 onwards, a new system involving picking shots has been implemented. At
the beginning of each round, both the players select a limited number of shots from the
list. During the round, the players can only use the shots that they have selected. 
	From v4.1.0 onwards, some shots will now subtract from the opponents damage or
armour mod. This new feature allows even more variety.

Shot Descriptions (For exact values, look in the options or the data file.)
(All descriptions apply to v3.4.0+. Earlier versions would have varying properties
for the pre v3.4.0 shots.)

The version number beside each shot is the version in which they were first implemented.

-----------------

- REDSHOT (v1.0.0): Rapid-fire weapon that shoots a dense line of shots.
	:-) A continuous line of redshots are impossible to dodge. The damage/time
ratio is very good, which makes this the perfect medium range weapon. Deals some
explosive damage as well. 
	:-( Because its damage, which is split almost 50/50 between normal and explosive,
isn't very high, it is a terrible weapon if your opponent has any sort of armour at all.
	Damage/Time = 1.833


- YELLOWSHOT (v1.0.0): Similar to the redshot, but doesn't fire as rapidly. Deals no explosive damage.
 	:-) When the redshot is neutralised because your opponent has 2 or 3 armour, use the
yellowshot instead. From v4.1, this shot also affects your opponents damage mod.
	:-( If your opponent has gathered a few armour powerups, the yellowshot isn't very
effective. If you try to use it long range like a redshot, it wont work as well because
the shots are much easier to dodge, and aren't any faster.
	Damage/Time = 1.60

- WHITESHOT (v1.0.0): A small but fast rocket that deals substantial damage.
	:-) Good armour-piercing capability with its 13-26 normal damage. Also fast enough
to hit opponents backing away.
	:-( Slow rate of fire and small size can make it hard to hit anything with this shot.
	Damage/Time = 1.63

- PIPEBOMB (v3.4.0): A medium-short range bomb that explodes.
	:-) Good damage and massive explosion radius makes it perfect against opponents who like
getting armour or dodging a lot. 
	:-( Low Damage/Time ratio means you wont last long going head-to-head against someone
using other shots. It can only be used as a medium-short range weapon because of its expiry time.
	Damage/Time = 1.35

	
- SLOWSHOT (v1.0.0): Slow-moving artillery shot that explodes on impact.
 	:-) Massive Damage/Time ratio. It is also bigger than redshot and yellowshot, and explodes
when it hits the wall as well.
	:-( Simply way too slow for even short range shooting. A human player will easily
dodge the shots, and might even be smart enough to move away from walls before the shots hit.
	Damage/Time = 2.45
	

- MINE (v3.4.0): A bomb that doesn't move, but expires and explodes after a while.
	:-) Constantly placing mines all over the map can seriouly hinder your opponents 
movement. Its big size and explosion radius will keep human players away from your mines - 
and any powerups that are near them. Placing them in front of a chasing opponent is also
a good strategy.
	:-( More of a deterrent than a serious damage-dealing weapon. Its long reload time
also leaves you vulnerable for a second or so. 
	Damage/Time = 1.02
	

- SNIPER SHOT (v3.4.0): Very fast long-ranged weapon.
	:-) A shot that is on-target is virtually impossible to dodge. 
It deals enough damage to still be useful against an opponent with 1 or 2 armour powerups,
and its quick fire rate means you can get a few shots off before your opponent responds.
	:-( You will easily be out-gunned if you try using this weapon at close range.
It also has a very small size.
	Damage/Time = 1.5
	

- FLAME (v3.4.1): Close range weapon that fires very rapidly.
	:-) Has the best Damage/Time ratio, dealing almost ridiculous amounts of damage
in a short space of time. It is also pretty easy to aim, even though the explosion radius
is quite small, because it fires so quickly that one or two missed shots wont matter much.
	:-( The only catch is that you have to be almost right next to your opponent to get 
them. The flame expires very quickly, but at least they explode when they expire.
	Damage/Time = 2.5
	
-MACHINE GUN (v3.5.0): Extremely fast rapid-fire weapon.
	:-) Travels almost instantaneosly across the screen, (faster than Sniper Shot)
 so it is very good for long-range shots. Fires as rapidly as the Flamethrower as well.
	:-( Only deals 3 damage per shot, so any amount of armour would significantly reduce
its damage. It is also as small as the sniper shot.
	Damage/Time = 1.5

-LIGHTNING (v3.5.0): Extremely fast weapon that deals a wide range of damage.
	:-) Faster and bigger than the sniper shot. It deals pretty good average damage.
	:-( The damage is very unpredictable, which could be a bad thing sometimes.
	Damage/Time = 1.34

-ACID (v4.1.0): A shot that degrades your opponent's armour.
	:-) Sustained fire from this weapon can quickly take down armour. As the armour goes
down, each shot starts dealing more damage as well. It explodes as well so it is easy to aim.
Also, if a shot hits your opponent, *both* the normal and explosive damage will take effect, with
the armour-eating effect that goes with them.
	:-( Doesn't deal much damage to start off with. 
	Damage/Time: 0
	
-GAUSS RIFLE (v4.1.0): A powerful but slow-firing long range weapon.
	:-) The fastest weapon in the game. Like the machine gun, it is almost instantaneos.
Deals enough damage to go through almost any armour.
	:-( Extremely slow rate of fire. You will be out-gunned if you use this at close range.
	Damage/Time=1.12

-BOMBARD (v4.1.0): An artillery weapon that explodes at a distance. Like a long-range pipebomb.
	:-) Moves quite fast. The good explosion radius makes it even harder to dodge. Deals more
damage than the pipebomb.
	:-( Slow rate of fire. Doesn't work at medium-close range like the pipebomb does.
	Damage/Time=1.35

----------------------------------------------------
	

<<PowerUps>>

	(BRAND NEW!) From version 4.0.0 on, powerups work differently to earlier versions. Instead of
the modifications being fixed until the time wears off, the effects will now slowly fade
away to 0 as the time remaining decreases.
	Powerups are things that appear randomly in the game at random intervals. They can
be picked up by moving into them. Powerups can have various effects:

- HEALING: Despite the name, these powerups can affect you or your opponent. Powerups like
SmallHeal increase your HP, while ones like SmallHurt damage your opponent.

- MODIFY DAMAGE: These either increase your the damage of your shots, or weakens your 
opponents shots. If your damage has been modified, it will show in the stats window. 
For example, 10% means your shots deal 110% damage, and -25% means they only deal 
75% damage. -100% or less means you're stuffed, but at least will wont deal 'negative damage'.
The damage modifier affects both explosive and normal damage.

- ARMOUR: These modify the the amount of shot damage taken by the player by the amount of 
armour. For example, an armour of 2 means all shots that hit the player deal 2 less damage, 
while an armour of -2 increases the damage by 2. If the shot deals less damage than the
armour, the player does *not* recieve negative damage. Note that armour affects explosive
damage and normal damage individually, so armour is effectively 2 times as good against
shots with both types of damage (or 2 times as bad if the modification is negative).

- EXTRA TIME: The extra time powerup isn't really that special after v4.0.0. Now, all
it does is give a small boost to both damage and armour. I don't know why I left it in
the game at all, really.

- FLOWERS(v3.3.12+): These powerups create a ring of shots that spread outwards. The ring
can either originate from where the powerup was, where the player who picked it up is, or
in random positions (depending on the settings). After a certain number of shots, they stop shooting.


<<The Player Classes>> *BRAND NEW!*

From v4.1.0, a new player class feature is available. To use the classes, simply make sure the
'use classes' button is selected when a new game is started.
At the 'Chose your shots' window (v3.5 onwards), you will notice that the player's class now
appears at the top. To change your class, use the Next/Prev weapon buttons. (Note: there seems
to be a bug here that happens if you cycle through your class types to rapidly, i.e holding down
the button. I haven't figured it out yet.)


<<The Almighty Som-AI>>

	The AI found in this game was written entirely by myself, without help from any
outside source, and I'm pretty damn proud of that fact. It still needs improvement, of
course, but it at least offers quite a challenge before I get network play implemented
(Two players can still play on one computer, but windows only allows 3 key inputs at once,
so there's a lot of key jamming).
	The AI can perform a small variety of actions. It operates as a simple FSM 
(Finite State Machine). That is, it will go between different states (that each perform 
specific actions) according to the inputs it receives. AI states can be enabled or 
disabled at the 'Enter your Name' dialogs.

	Note: If you ever read my source code, you will see that the AI does not cheat in
any way. Although it would have actually been much simpler for me to write an AI that
takes shortcuts (like selecting a shot straight away, for example), I have chosen not
to do so. So if you find yourself constantly beaten by the computer, you can rest assured
that the AI won *fairly*.



AI procedures
-------------

- Random AI: computer will perform random moves, exactly like a person hitting random action keys.
This state is active when the AI can't find anything else better to do. (Version 1.0.0 +)

- Lay mines: computer will select a shot with a movement rate of 0 and place the shots
randomly. This procedure is only called from the Random AI, so it wont lay mines if it can do
something else. (Version 3.4.0 +)

- Smart Shooting: AI will go into this state if the opponent is at any 45º angle to it. It
will turn in the direction of the opponent and start shooting whatever shot is currently
selected. This was the first real AI that I wrote, and it hasn't changed at all since then.
(Version 2.0.0 +)

- Chasing AI: this isn't really an AI. All it does is make the AI move forward all the time.
Combined with Smart Shooting, and with the Random AI and Find Powerup AI disabled, it will
simply keep shooting and advancing towards its opponent mercilessly. It is easily killed
if it can't find you and gets stuck againts a wall.

- Choose Weapon: assuming all the shots are balanced, this procedure will make the AI choose
the best weapon to use, given the distance between it and its opponent. Also, it will only 
use the faster shots if its opponent is backing away and the AI is chasing after it/you.
From Version 3.4.0 onwards, it will also use exploding shots properly. 

- Find Powerups: this AI was written after Power-Ups were implemented in the game. 
The AI will firstly check if there are any powerups on the screen that are not too close to 
its opponent. If so, it will move towards the closest one. This state will activate and
overide all the other states.

	You've probably noticed that something seems to be missing. That's right: there is no 'defence'
state where the AI will dodge shots. This is the AI's main weakness, and exploiting it wisely
against the computer will usually win you the game. Mind you, the AI is by no means easy to 
hit when the Random AI is running.
	It is possible to pit two AI's against each other, but its pretty boring and not recommended
if your aim is to have fun.


<<Controls>>

	Customisable controls have now been implemented! However, some of the keys will not be
shown properly if you check the 'Characters' checkbox, so it might be a bit confusing.
Customised controls are saved into the registry.

These are the default controls:

Action            Player 1      Player 2
=========================================

Forward	      W             Up Key
Back              S             Down Key
Left              A             Left Key
Right             D             Right Key

Shoot	            G             NumPad 0
Select prev wep   T             NumPad 1
Select next wep   Y             NumPad 2


<<Network Play (BRAND NEW!)>>

Version 4.0.0 and above will come with a network play feature. It is pretty much fully
working, except that the client will not hear any sounds. It also runs a bit slowly at times,
especially when there are lots of shots and powerups on the screen. However, as I said,
it *does* work, so give it a try if you have a LAN network. (It works over the internet as
well, but unless you have Cable/DSL, it will lag pretty badly.)

How to play over a network
--------------------------
To play over a network, you will need two computers connected together somehow. One will
be the server, and the other will be a client. It doesn't matter which is which, but running
the server on the faster computer is a good idea.

Before you start:
-Make sure you know the IP address or name of the computer that will be serving. To do this,
either go to the DOS prompt and type in 'IPCONFIG', or go to Run and type in 'WINIPCFG'. Note
the IP address of the device you will be using.
-Close down any other progams that are using the same connection. This prevents conflicts, and 
might improve game performance. 
-Agree on a port number to use.

Starting a network game:
1. Start DeathMatch SOM on both computers.
2. On the SERVER, click on the Host game button.
3. Type in any port number between 5001 and 32000 (windows uses some of the ports below 5000)
4. Click on OK. The buttons should now grey out as the server waits for a connection.
5. On the CLIENT, click on the Connect to server button.
6. Type in either the name or IP address of the serving computer.
7. Click OK.
8. The two computers should now connect, and display the 'Enter you name' dialogs.



<<The Options dialog>>

	In case you missed it, there is a clear warning message on the top of the options
dialog window. This is there because changing a few settings can easily stuff up the game.
Most of the options were NOT intended to be modified by the end user, so consider yourself lucky I
included it at all.
	That said, here's a description on what everything does:


Main Options
------------

PLAYERHP - The amount of hitpoints players start with. This setting applies to both players,
and takes affect at the next round. Default is 250.

EXPLOSIONWAIT - Controls how many units of game time that the explosion animations will show.
If you set it to 0, you wont see any explosions at all (the shots just disappear when they hit
something). Setting it to a lower number will slightly improve game performance.

MOVESTRAIGHT - You really shouldn't change this, but if you do, it determines how far you can
move in one unit of game time. Default is 7. Takes effect at the beginning of the next round.

MOVEDIAG - Obselete after v3.4.1. In earlier versions, determines how far you move diagnally.
This is now calculated with MOVESTRAIGHT and pythagorus's threom.

DEFSHOTSPEED - Shots will move at this speed if they are not assigned a speed. Pretty obselete,
as all shots have been assigned a speed after v3.4.0. Default is 2.

DEFSHOTSOUND - Shots make this sound if not assigned a sound. Must be a wav file in the 'sounds'
sub-folder.

GAMESPEED - The number of milliseconds between each unit of game time. The lower the number, the
faster the game runs. It is pointless setting this to anything lower than 30 if you have a 
DOS-based operating system like Windows 98,95,etc. Default is 30, BUT IF YOU HAVE A WINDOWS NT BASED
OS (LIKE WIN 2000 AND XP), SET THIS TO 55.

POWERUPPERCENT - The chance of a powerup appearing during each unit of game time, as a percentage.
Setting it higher will make powerups appear more rapidly. If you set this too high, the 
find powerups AI will just go around picking up powerups all the time instead of shooting you.
Setting it at 0 means that powerups wont appear at all. Default is 2.5.

POWERUPSIZE - The size of the powerups. A larger size means they are picked up more easily.
You shouldn't try to change this, because the default size is roughly the same as the size
of the pictures. Default is 10.

POWERUPMAX - The maximum number of powerups in the game at once. If you set POWERUPPERCENT
really high and this really high, you can flood the game screen with powerups (not that I see
why you would want to.) Default is 10.

BUFFERS - The number of sounds that can play at once. If the game runs a bit slowly, set it to
something lower, like 4. It wont sound as good, but because all sounds are prioritised, it 
will simply cut off the less important sounds, such as shots hitting the wall. This takes
affect when the next round starts. Default is 8. This setting has a great effect on game
performance, especially if you have an old soundcard, or you don't have much memory (RAM).

PICKSHOTS (brand new!) - The number of shots you can pick to use at the beginning of each
round. 


Shot Options
------------

None of the Shot Options should be changed. The game is pretty well balanced as it is, so try
not to muck around with the settings.

NAME - The name of the shot. This is used to load the shot's bitmap, so it should not be changed.

DAMLOW - The minimum amount of damage dealt by the shot.

DAM - The max damage of the shot

WAIT - The amount of time you must wait after firing the shot ('reload time').

SPEED - The speed of the shot relative to the speed of the players. (e.g a speed
of 2 means the shot moves 2 times the speed of the players).

RADIUS - How big the shot is. Bigger shots have a better chance of hitting the target.

SOUND - The sound that plays when the shot is fired.

EXLOW - The minimum explosion damage of the shot.

EXHIGH - The maximum explosion damage.

EXAREA - The area covered by an explosion. 

EXPIRE - The minimum time taken for the shot to expire. This is empty if shot doesn't expire.

EXPIREHIGH - The max time taken for shot to expire.

TRAIL - How big a trail the shot leaves (if any). Maximum is 5.


Power-up Options
----------------

As with the shot options, these should not be changed. 

PNAME - Name of powerup. Used to load the powerup's picture.

EFFECT - The *type* of effect the powerup has (see PowerUps section for details).

LOW - The minimum effect of powerup. 

HIGH - The maximum effect of powerup.

CHANCE - The relative chance of this powerup being created. 

TIMEMIN - The minimum time this powerup will take effect (if applicable).

TIMEMAX - The maximum time the powerup will take effect.

YOU - Whether the powerup effects you or your opponent. 

SOUND - Sound made when powerup is picked up.


<<Acknowledgements>>

All sounds were made with my Panasonic© General MIDI keyboard, except for the following:

- Mine.wav (BLEEP6.WAV) 
- Fire.wav (FIRETRT1.WAV) 
- Type.wav (CASHDN1.WAV)
- Zap.wav  (APPEAR1.WAV)
- SILENCER.WAV
- KEYSTROK.WAV
- HYDROD1.WAV
- COUNTRY4.WAV
- STINGER.WAV (IRONCUR9.WAV)
- zapped 1,2,and 3.wav (Tesla2.wav)
From 'Red Alert: The Aftermath'

- ump45-1.wav
- acid.wav (headshot2.wav)
- awp1.wav
- elite_fire.wav
- ric0.wav, ric1.wav, ric2.wav (ric1.wav, ric3.wav, ric2.wav)
- hit0.wav (ric_metal-2.wav)
Shortened wavs from Counter-Strike


The automatic mask creating procedure in versions 3.4.0 and above was written with the help of the source
code from DSMaskCreator, by Winter Grave.
I learnt how to use Direct-Sound 
(Search for them in Planet-source-code.com).

All bitmaps were made by me with Paint Shot Pro 4, except for the following:

- MoreTime.bmp (Modified icon from Visual Basic)

- Mine.bmp (Modified icon from Red Alert)

- Ring of Destruction.bmp (Modified icon from Dark Reign)

- SmallHurt.bmp
- MinorArmour.bmp (Icons from Windows, found in pifmgr.dll)

In upcoming versions, I will implement a feature that allows customisable graphics,
sounds, shot settings, and powerup settings without having to replace existing files
and settings. This was inspired by Nathan ('Jin') who has personally designed many new
graphics for the game, which, unfortunately, I haven't been able to include...*yet*.

||||||||||||||||||||||||||||

Positions© [DeathMatch SOM]©.

You can contact me by email: forthelight@writeme.com
OR
ICQ: 55756521
