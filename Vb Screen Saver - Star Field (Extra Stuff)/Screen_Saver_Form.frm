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
Dim X1%(355, 3)
Dim I
Dim Starter%
Dim oldX%
Dim oldy%
Dim TEmpyY%

Private Sub Form_KeyPress(KeyAscii As Integer)
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Starter% = 1 Then
    If oldX% <> Y And oldy% <> X Then
        End
    End If
Else
    oldX% = Y
    oldy% = X
End If
End Sub

Private Sub Timer3_Timer()
Starter% = 1
Timer3.Enabled = False
TEmpyY% = 1
End Sub

Private Sub Timer1_Timer()
Randomize Timer
Me.Cls

For I = 1 To 355
    If X1%(I, 0) = 0 Then
        X1%(I, 0) = 255
        X1%(I, 1) = Rnd * Speed% - Speed% / 2
        X1%(I, 2) = Rnd * Speed% - Speed% / 2
        Exit For
    End If
Next I


For I = 1 To 355
    Me.ForeColor = RGB(255, 255, 255)
    X1%(I, 0) = X1%(I, 0) - 1
    If X1%(I, 0) < 0 Then X1%(I, 0) = 0
    Select Case Typer%
    Case 1
        DecayingParticle
    Case 2
        StarField
    Case 3
        StarField2
    Case 4
        DecayingParticle2
    End Select
    Me.ForeColor = RGB(0, 0, 0)
    Me.PSet (Me.Width / 2, Me.Height / 2)
Next I
End Sub

Private Sub DecayingParticle()
    Me.PSet (Me.Width / 2 + X1%(I, 1) * (256 - X1%(I, 0)), Me.Height / 2 + X1%(I, 2) * (256 - X1%(I, 0)))
End Sub

Private Sub StarField()
    Me.PSet (Me.Width / 2 + X1%(I, 1) * ((256 - X1%(I, 0)) ^ (Accel% / 10)), Me.Height / 2 + X1%(I, 2) * ((256 - X1%(I, 0)) ^ (Accel% / 10)))
End Sub

Private Sub StarField2()
Me.ForeColor = RGB(256 - X1%(I, 0), 256 - X1%(I, 0), 256 - X1%(I, 0))
    Me.PSet (Me.Width / 2 + X1%(I, 1) * ((256 - X1%(I, 0)) ^ (Accel% / 10)), Me.Height / 2 + X1%(I, 2) * ((256 - X1%(I, 0)) ^ (Accel% / 10)))
End Sub

Private Sub DecayingParticle2()
Me.ForeColor = RGB(256 - X1%(I, 0), 256 - X1%(I, 0), 256 - X1%(I, 0))
    Me.PSet (Me.Width / 2 + X1%(I, 1) * (256 - X1%(I, 0)), Me.Height / 2 + X1%(I, 2) * (256 - X1%(I, 0)))
End Sub

