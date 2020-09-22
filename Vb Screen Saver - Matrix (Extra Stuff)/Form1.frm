VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1665
      Top             =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X1#
Dim Speed%
Dim Array1#(22, 100)
Dim Amount%
Dim Spacing%

Private Sub Form_Load()
Randomize Timer
Speed% = 35
Amount% = 30
Spacing% = 300
For J = 0 To Amount%
    For I = 0 To 20
        Array1#(I, J) = Int(Rnd * 97 + 33)
    Next I
    Array1#(22, J) = Int(Rnd * Me.Width / (10 * 30)) * (10 * 30)
    Array1#(21, J) = Int(Rnd * Me.Height * 3 / Speed% - Spacing%)
Next J
End Sub

Private Sub Timer1_Timer()
Randomize Timer
Me.Cls

For J = 0 To Amount%
    X2# = 20
    For I = 0 To 20
        Me.ForeColor = 0
        Me.PSet (Array1#(22, J), Array1#(21, J) * Speed% + (I * Spacing% + 1))
        Me.ForeColor = RGB(0, 255 - (X2# * 12), 0)
        If I = 20 Then Me.ForeColor = RGB(100, 255, 100)
        Me.Print Chr$(Array1#(I, J))
        X2# = X2# - 1
    Next I
    Array1#(21, J) = Array1#(21, J) + 5
    If Array1#(21, J) * Speed% > Me.Height Then
        Array1#(21, J) = -(Spacing% + 5) * 20 / Speed%
        Array1#(22, J) = Int(Rnd * Me.Width / (10 * 30)) * (10 * 30)
    End If
Next J
End Sub
