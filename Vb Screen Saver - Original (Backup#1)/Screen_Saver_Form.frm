VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MouseIcon       =   "Screen_Saver_Form.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   4680
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   3960
      Top             =   2160
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   4200
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1560
      Top             =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
'---Readme.txt---
'»VB Screen Saver« --> By Alex M
'
'_______________________________________________________________________________________________
'
'Before hand
'________________________________________________________________________________________________
'
'    This was originally a little muck around project with the mouse. I made the original program so that wherever the mouse moved it drew little circles or dots, etc. But then I remembered that upon opening a previous .scr file with a hex viewer, that it started with the characters "MZ". Now this is when you say, "So What?". Well This is important because normal Executables also start with "MZ" and thats when it hit me, What if I rename the .exe to .scr? Well to some of you more advanced programmers/computer users may be laughing at this point saying "Yeah, I already knew that you could rename the file from .exe to .scr" but It was a learning experience and I'm not even a real programmer as such, programming is just an interest not my profession.
'    Anyways, Back to something more relevant, I also used the command$ string to find out that windows sends data to the program to tell it whether it wants to change settings "/c" or to test settings "/p" or to run the actual screen saver "Null". Another problem was the fact that whenever the program started, it started minimised so I loaded a separate form that was maximised. Sorry about the poor structure of the code, I learned basic using GWBasic as many of its elements of structure still exist in my code (such as GOTO, lol). I Hope this code is some use to you, you CAN use any part of code you want, Copy what you want, delete what you don't (Just don't Take the whole program and pass it off as your own, Thats called plagiarism)
'
'________________________________________________________________________________________________
'
'Creating the Screen Saver
'________________________________________________________________________________________________
'
'Its pretty easy, just follow these steps,
'    1. Compile the Code to an .Exe file
'    2. Rename the .Exe to a .Scr
'    3. Right click the .Scr file and select install
'    4. Done.
'
'________________________________________________________________________________________________
'
'Trouble Shooting
'________________________________________________________________________________________________
'    Note, when compiling, if you encounter errors, change the compiling method to P-code instead of native code (Native code is faster, but can cause problems)
'
'    Any Further Questions, E-mail me at Alex_murray1@hotmail.com, I'll do my best to answer them (Also Make the subject relevant or I'll think its spam, and I check my e-mail atleast every week)
'---------------------------------------------------------
'
'Main Array
Dim X1%(255, 3)
'Position Variables
Dim Xtemp%
Dim Ytemp%
Dim Xtemp2%
Dim Ytemp2%
'Just a number
Dim I
'Defines Colour
Dim RandomRGB%
'position Variable
Dim Mult%
Dim Mult2%
'This is to get rid of some null variables at the start
Dim Starter%
'Work out whether the mouse has been moved (Due to resize
'Vb thinks the mouse has move from the time it loads and
'the time it draws)
Dim oldX%
Dim oldy%
'This is to get rid of some null variables at the start
Dim TEmpyY%

'On Key press Exit
Private Sub Form_KeyPress(KeyAscii As Integer)
End
End Sub

'Set the direction of the movement (ie +1 or -1)
Private Sub Form_Load()
Mult% = 1
Mult2% = 1
End Sub

'Close on mouse move
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'When loading, record mouse position (After resize)
If Starter% = 1 Then
    If oldX% <> Y And oldy% <> X Then
        End
    End If
Else
'If mouse has been moved, then exit
    oldX% = Y
    oldy% = X
End If
End Sub

Private Sub Timer1_Timer()
'If still loading, then don't draw (To get rid of null
'values)
If TEmpyY% = 0 Then GoTo 10
'Refresh screen
Me.Cls
'A bit of obselete Code
'-----------------------------------
'RandomRGB% = RandomRGB% + Mult%
'If RandomRGB% = 255 Or RandomRGB% = 0 Then Mult% = Mult% * -1
'-----------------------------------
'Find the next free position in the array
For I = 1 To 255
    If X1%(I, 1) = 0 Then
        X1%(I, 1) = 255
        X1%(I, 2) = Xtemp%
        X1%(I, 3) = Ytemp%
        Exit For
    End If
Next I

'This For loop is for drawing each Circle or line or etc
For I = 1 To 255
    'Make the colour fade to black (or there abouts)
    X1%(I, 1) = X1%(I, 1) - 1
    'if the colour is less than 0 then it is set to zero
    If X1%(I, 1) < 0 Then X1%(I, 1) = 0
    'This is the colour of the stuff thats Drawn
    Form1.ForeColor = RGB(0, X1%(I, 1), X1%(I, 1))
    'This is Obselete Code
    '--------------------------------------
    'Form1.ForeColor = RGB(X1%(I, 1),0, X1%(I, 1))
    'Form1.ForeColor = RGB(0, 0, X1%(I, 1))
    'Form1.ForeColor = RGB( X1%(I, 1), X1%(I, 1),X1%(I, 1))
    '--------------------------------------
    'Select which Drawing style/pattern to use
    Select Case Typer%
    Case 1
        DrawOneLine
    Case 2
        DrawTwoLines
    Case 3
        DrawCircle
    Case 4
        DrawThreeLines
    Case 5
        DrawFancyCircle
    Case 6
        DrawDots
    Case 7
        DrawManyDots
    Case 8
        DrawManyCircles
    Case 9
        DrawSOManyMoreDots
    Case 10
        DrawSuicidalDots
    Case 11
        DrawSlowerSuicidalDots
    Case 12
        DrawRippleEffect
    Case 13
        DrawTunnelEffect
    Case 14
        ParticleCannon
    Case 15
        RushingWater
    Case 16
        RushingWater2
    Case 17
        ParticleCannon2
    Case 18
        ParticleCannon3
    Case 19
        ParticleCannon4
    Case 20
        MovingLines
    End Select
Next I
10
End Sub

'Newly Added
'-----------------------------
Private Sub MovingLines()
    Form1.Line (0, (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))-(Me.Width, (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))
End Sub
Private Sub ParticleCannon2()
    Form1.PSet ((X1%(I, 2)) / 100 * (255 - X1%(I, 1)), (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))
    Form1.PSet (Me.Width - (X1%(I, 2)) / 100 * (255 - X1%(I, 1)), Me.Height - (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))
End Sub

Private Sub ParticleCannon3()
    Form1.PSet ((X1%(I, 2)) / 100 * (255 - X1%(I, 1)), (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))
    Form1.PSet (Me.Width - (X1%(I, 2)) / 100 * (255 - X1%(I, 1)), Me.Height - (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))
    Form1.PSet ((X1%(I, 2)) / 100 * (255 - X1%(I, 1)), Me.Height - (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))
End Sub

Private Sub ParticleCannon4()
    Form1.PSet ((X1%(I, 2)) / 100 * (255 - X1%(I, 1)), (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))
    Form1.PSet (Me.Width - (X1%(I, 2)) / 100 * (255 - X1%(I, 1)), Me.Height - (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))
    Form1.PSet ((X1%(I, 2)) / 100 * (255 - X1%(I, 1)), Me.Height - (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))
    Form1.PSet (Me.Width - (X1%(I, 2)) / 100 * (255 - X1%(I, 1)), (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))
End Sub

Private Sub ParticleCannon()
    Form1.PSet ((X1%(I, 2)) / 100 * (255 - X1%(I, 1)), (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))
End Sub

Private Sub RushingWater()
    Form1.Line ((X1%(I, 2)) / 100 * (255 - X1%(I, 1)), (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))-(0, 0)
End Sub

Private Sub RushingWater2()
    Form1.Line ((X1%(I, 2)) / 100 * (255 - X1%(I, 1)), (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))-(0, 0)
    Form1.Line ((X1%(I, 2)) / 100 * (255 - X1%(I, 1)), (X1%(I, 3)) / 100 * (255 - X1%(I, 1)))-(Me.Width, Me.Height)
End Sub
'--------------------------------------------

Private Sub DrawTunnelEffect()
'Draw circles that get smaller as their colour changes
'from 255 to 0
        Form1.Circle (X1%(I, 2), X1%(I, 3)), Size% * 10 * X1%(I, 1)
End Sub

Private Sub DrawRippleEffect()
'Draw circles that get bigger as their colour changes
'from 255 to 0
Form1.Circle (X1%(I, 2), X1%(I, 3)), Size% * 10 * (256 - X1%(I, 1))
End Sub

Private Sub DrawSuicidalDots()
'Randomize using the System Clock as the seed (more
'random than normal randomize)
Randomize Timer
'Draws 5 Dots at a random distance from a point which
'increases as their colour goes from 255 to 0
For Y = 1 To 5
        Form1.PSet (Int(X1%(I, 2) + (Rnd * 2 - 1) * Size% * 15 * (256 - X1%(I, 1))), Int(X1%(I, 3) + (Rnd * 2 - 1) * Size% * 15 * (256 - X1%(I, 1))))
Next Y
End Sub

Private Sub DrawSlowerSuicidalDots()
'Randomize using the System Clock as the seed (more
'random than normal randomize)
Randomize Timer
'Draws 4 Dots at a random small distance from a point
'which increases as their colour goes from 255 to 0
For Y = 1 To 4
        Form1.PSet (Int(X1%(I, 2) + (Rnd * 2 - 1) * Size% * 15 * (256 - X1%(I, 1)) / Size% / 5), Int(X1%(I, 3) + (Rnd * 2 - 1) * Size% * 15 * (256 - X1%(I, 1)) / Size% / 5))
Next Y
End Sub

Private Sub DrawSOManyMoreDots()
'Randomize using the System Clock as the seed (more
'random than normal randomize)
Randomize Timer
'Draws 20 Dots at a random distance from a point which is
'independent of its colour change
For Y = 1 To 20
        Form1.PSet (Int(X1%(I, 2) + (Rnd * 2 - 1) * Size% * 30), Int(X1%(I, 3) + (Rnd * 2 - 1) * Size% * 30))
Next Y
End Sub

Private Sub DrawDots()
'Just draw a dot where ever the movement has been
        Form1.PSet (X1%(I, 2), X1%(I, 3))
End Sub

Private Sub DrawManyCircles()
'Randomize using the System Clock as the seed (more
'random than normal randomize)
Randomize Timer
'Draw 4 Circles around set radii and random distances
'from a point
For Y = 1 To 4
        Form1.Circle (Int(X1%(I, 2) + (Rnd * 2 - 1) * Size% * 15), Int(X1%(I, 3) + (Rnd * 2 - 1) * Size% * 15)), Size% * 15
Next Y
End Sub

Private Sub DrawManyDots()
'Randomize using the System Clock as the seed (more
'random than normal randomize)
Randomize Timer
'Draw alot of dots at random distances about a set point
For Y = 1 To 4
        Form1.PSet (Int(X1%(I, 2) + (Rnd * 2 - 1) * Size% * 15), Int(X1%(I, 3) + (Rnd * 2 - 1) * Size% * 15))
Next Y
End Sub

Private Sub DrawFancyCircle()
'This draws a circle with two lines, one line through its
'centre the other somewhere around its circumference
        Form1.Circle (X1%(I, 2), X1%(I, 3)), Size% * 30
    If I = 1 Then
        Form1.Line (X1%(I, 2) + Size% * 30, X1%(I, 3) + Size% * 30)-(X1%(255, 2) + Size% * 30, X1%(255, 3) + Size% * 30)
        Form1.Line (X1%(I, 2) - Size% * 30, X1%(I, 3) - Size% * 30)-(X1%(255, 2) - Size% * 30, X1%(255, 3) - Size% * 30)
    Else
        Form1.Line (X1%(I, 2) + Size% * 30, X1%(I, 3) + Size% * 30)-(X1%(I - 1, 2) + Size% * 30, X1%(I - 1, 3) + Size% * 30)
        Form1.Line (X1%(I, 2) - Size% * 30, X1%(I, 3) - Size% * 30)-(X1%(I - 1, 2) - Size% * 30, X1%(I - 1, 3) - Size% * 30)
    End If
End Sub

Private Sub DrawCircle()
'just draw a circle about a point
        Form1.Circle (X1%(I, 2), X1%(I, 3)), Size% * 15
End Sub

Private Sub DrawOneLine()
'draws a single line from the current point to the
'previous point
    If I = 1 Then
        Form1.Line (X1%(I, 2), X1%(I, 3))-(X1%(255, 2), X1%(255, 3))
    Else
        Form1.Line (X1%(I, 2), X1%(I, 3))-(X1%(I - 1, 2), X1%(I - 1, 3))
    End If
End Sub

Private Sub DrawTwoLines()
'draws two lines from the current point to the previous
'point
    If I = 1 Then
        Form1.Line (X1%(I, 2), X1%(I, 3))-(X1%(255, 2), X1%(255, 3))
        Form1.Line (X1%(I, 2) + Size% * 30, X1%(I, 3) + Size% * 30)-(X1%(255, 2) + Size% * 30, X1%(255, 3) + Size% * 30)
    Else
        Form1.Line (X1%(I, 2), X1%(I, 3))-(X1%(I - 1, 2), X1%(I - 1, 3))
        Form1.Line (X1%(I, 2) + Size% * 30, X1%(I, 3) + Size% * 30)-(X1%(I - 1, 2) + Size% * 30, X1%(I - 1, 3) + Size% * 30)
    End If
End Sub

Private Sub DrawThreeLines()
'draws three lines from the current point to the previous
'point
    If I = 1 Then
        Form1.Line (X1%(I, 2), X1%(I, 3))-(X1%(255, 2), X1%(255, 3))
        Form1.Line (X1%(I, 2) + Size% * 30, X1%(I, 3) + Size% * 30)-(X1%(255, 2) + Size% * 30, X1%(255, 3) + Size% * 30)
        Form1.Line (X1%(I, 2) - Size% * 30, X1%(I, 3) - Size% * 30)-(X1%(255, 2) - Size% * 30, X1%(255, 3) - Size% * 30)
    Else
        Form1.Line (X1%(I, 2), X1%(I, 3))-(X1%(I - 1, 2), X1%(I - 1, 3))
        Form1.Line (X1%(I, 2) + Size% * 30, X1%(I, 3) + Size% * 30)-(X1%(I - 1, 2) + Size% * 30, X1%(I - 1, 3) + Size% * 30)
        Form1.Line (X1%(I, 2) - Size% * 30, X1%(I, 3) - Size% * 30)-(X1%(I - 1, 2) - Size% * 30, X1%(I - 1, 3) - Size% * 30)
    End If
End Sub

Private Sub Timer2_Timer()
'This timer defines the movement of the drawn objects

'Increment by 1 or decrement by 1
Xtemp2% = Xtemp2% + Mult% * Speed%
Ytemp2% = Ytemp2% + Mult2% * Speed%

Select Case MovementType%
Case 1
    'This causes the drawn objects to move like an inverse
    'Sine curve about the X axis and move Linear about the
    'Y axis
    Xtemp% = Sin(Xtemp2% / 500) * 300 * 15 + Me.Width / 2
    If Xtemp2% / 500 > 3.14159265358979 * 2 Then Xtemp2% = 0
    Ytemp% = Ytemp2%
    If Ytemp2% >= Form1.Height Or Ytemp2% < 0 Then Mult2% = Mult2% * -1
Case 2
    'Simple Linear movement about X,Y Axes
    If Ytemp2% >= Form1.Height Or Ytemp2% < 0 Then Mult2% = Mult2% * -1
    If Xtemp2% >= Form1.Width Or Xtemp2% < 0 Then Mult% = Mult% * -1
    Xtemp% = Xtemp2%
    Ytemp% = Ytemp2%
Case 3
    'This causes the drawn objects to move like an inverse
    'Sine curve about the X axis and move like an inverse
    'Sine curve about the Y axis at a slightly slower
    'speed
    Xtemp% = Sin(Xtemp2% / 500) * 20 * Speed% * 15 + Me.Width / 2
    If Xtemp2% / 500 > 3.14159265358979 * 2 Then Xtemp2% = 0
    Ytemp% = Sin(Ytemp2% / 1000) * 20 * Speed% * 15 + Me.Height / 2
    If Ytemp2% / 1000 > 3.14159265358979 * 2 Then Ytemp2% = 0
End Select
End Sub

Private Sub Timer3_Timer()
'Set a delay to remove null values at the start of the
'program
Starter% = 1
Timer3.Enabled = False
TEmpyY% = 1
End Sub
