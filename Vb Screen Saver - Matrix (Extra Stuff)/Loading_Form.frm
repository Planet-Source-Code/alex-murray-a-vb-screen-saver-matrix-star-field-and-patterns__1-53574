VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading...."
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   1605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Loading_Form.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   1605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   960
      Top             =   960
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempr%

Private Sub Form_Load()
On Error Resume Next
FIlepos$ = "C:\Settings645.cfg"
If Mid$(Command$, 1, 2) = "/p" Then End
If Mid$(Command$, 1, 2) = "/c" Then tempr% = 1
Open FIlepos$ For Input As #1
Input #1, B$
Input #1, C$
Close
If B$ = "" Then B$ = "30"
If C$ = "" Then C$ = "3"
Amount% = Val(B$)
Speed% = Val(C$) * 15
End Sub

Private Sub Timer1_Timer()
If tempr% = 1 Then
    Load Form3
    Me.Hide
    Form3.Show
Else
    Load Form1
    Me.Hide
    Form1.Show
End If
Timer1.Enabled = False
End Sub
