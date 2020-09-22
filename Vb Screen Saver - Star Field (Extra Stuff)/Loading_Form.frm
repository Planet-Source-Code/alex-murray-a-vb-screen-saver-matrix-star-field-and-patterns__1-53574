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
'If error occurs, then continue (mainly if config file
'does not exist)
On Error Resume Next
'Obselete data
'------------------------------
'Open "c:\temp.txt" For Append As #1
'Print #1, Time$ + " Line:-" + Command$
'Close
'------------------------------
'The File to save settings to
FilePos$ = "C:\Settings223.cfg"
'This is when windows tells the screen saver program to
'test the data entered
If Mid$(Command$, 1, 2) = "/p" Then End
'This is when you press settings in the select
'screen saver menu
If Mid$(Command$, 1, 2) = "/c" Then tempr% = 1
'Open the File to save settings to
Open FilePos$ For Input As #1
Input #1, A$
Input #1, B$
Input #1, C$
Input #1, D$
Close
'If the settings are incorrect then correct them
If A$ = "" Then A$ = "0"
If B$ = "" Then B$ = "3"
If C$ = "" Then C$ = "15"
If D$ = "" Then D$ = "1"
'Save the data to variables
Typer% = Val(A$) + 1
Accel% = Val(B$)
Speed% = Val(C$)
'MovementType% = Val(D$)
End Sub

Private Sub Timer1_Timer()
'If at the select screen saver menu the user selects
'settings then display the settings form instead of the
'actual screen saver
If tempr% = 1 Then
    Load Form3
    Me.Hide
    Form3.Show
Else
'open the screen saver form
    Load Form1
    Me.Hide
    Form1.Show
End If
'disable this timer
Timer1.Enabled = False
End Sub
