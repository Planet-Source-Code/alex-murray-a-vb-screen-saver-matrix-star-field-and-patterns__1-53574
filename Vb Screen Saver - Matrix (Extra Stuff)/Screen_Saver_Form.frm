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
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
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
      Left            =   0
      Top             =   0
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
Dim X1#
Dim Array1#(23, 200)
Dim Spacing%
Dim Accel%
Dim oldX%
Dim oldy%
Dim TEmpyY%
Dim STarter%

Private Sub Timer3_Timer()
STarter% = 1
Timer3.Enabled = False
TEmpyY% = 1
Wid1% = Me.Width
Hig1% = Me.Height
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
End
End Sub

Private Sub Form_Load()
Randomize Timer
Spacing% = 300
For J = 0 To Amount%
    For I = 0 To 20
        Array1#(I, J) = Int(Rnd * 97 + 33)
    Next I
    Array1#(21, J) = Int(Rnd * Me.Height * 3 / Speed% - Spacing%)
    Array1#(22, J) = Int(Rnd * Me.Width / (10 * 30)) * (10 * 30)
    Array1#(23, J) = Rnd * 10 + 5
Next J
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If STarter% = 1 Then
    If oldX% <> Y And oldy% <> X Then
        End
    End If
Else
    oldX% = Y
    oldy% = X
End If
End Sub

Private Sub Timer1_Timer()
Randomize Timer
Me.Cls

For J = 0 To Amount%
    X3# = Array1#(23, J) / 10
    X2# = 20
    For I = 0 To 20
        Me.ForeColor = 0
        Me.PSet (Array1#(22, J), Array1#(21, J) * Speed% + (I * Spacing% + 1))
        Me.ForeColor = RGB(0, 255 - (X2# * 12), 0)
        If I = 20 Then Me.ForeColor = RGB(100, 255, 100)
        Me.Print Chr$(Array1#(I, J))
        X2# = X2# - 1
    Next I
    Array1#(21, J) = Array1#(21, J) + 5 * X3#
    If Array1#(21, J) * Speed% > Me.Height Then
        Array1#(21, J) = -(Spacing% + 5) * 20 / Speed% - Int(Rnd * 200)
        Array1#(22, J) = Int(Rnd * Me.Width / (10 * 30)) * (10 * 30)
        Array1#(23, J) = Rnd * 10 + 5
    End If
Next J
End Sub

