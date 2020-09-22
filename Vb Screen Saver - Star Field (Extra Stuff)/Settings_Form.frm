VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2985
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   2985
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Settings_Form.frx":0000
      Left            =   2415
      List            =   "Settings_Form.frx":000D
      TabIndex        =   10
      Text            =   "01 - Sine X Linear Y"
      Top             =   3750
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   240
      Max             =   100
      Min             =   10
      TabIndex        =   6
      Top             =   2025
      Value           =   10
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   240
      Max             =   20
      Min             =   10
      TabIndex        =   3
      Top             =   1200
      Value           =   10
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Settings_Form.frx":004E
      Left            =   240
      List            =   "Settings_Form.frx":005E
      TabIndex        =   0
      Text            =   "01 - Decaying Particle"
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Faster"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Slower"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Movement Type"
      Height          =   255
      Left            =   2295
      TabIndex        =   11
      Top             =   3510
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Speed"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1785
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Slower"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2265
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Faster"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2265
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Acceleration"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Option"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Save the data set in the form
On Error Resume Next
'Open the settings file, as set in form2, and save
'settings. I used Sequential file saveing (output)
'instead of random or binary because its easier
'(but slightly bigger)
Open FilePos$ For Output As #1
'This converts the values to strings (it also gets the
'first two characters of the name of the selected item
'in the combo box, look and you'll see why)
R$ = Str$(Val(Mid$(Combo1.Text, 1, 2)) - 1)
Print #1, R$
R$ = Str$(HScroll1.Value)
Print #1, R$
R$ = Str$(HScroll2.Value)
Print #1, R$
R$ = Str$(Val(Mid$(Combo2.Text, 1, 2)))
Print #1, R$
'Close all open files
Close
'close program
End
End Sub

Private Sub Command2_Click()
'close program
End
End Sub

Private Sub Form_Load()
On Error Resume Next
'Opens the current saved settings from the setting file
Open FilePos$ For Input As #1
Input #1, A$
Input #1, B$
Input #1, C$
Input #1, D$
Close
'If any values do not exist, then set them
If A$ = "" Then A$ = "0"
If B$ = "" Then B$ = "3"
If C$ = "" Then C$ = "15"
If D$ = "" Then D$ = "1"
'Set the data to Objects on the Form
Combo1.Text = Combo1.List(Val(A$))
HScroll1.Value = Val(B$)
HScroll2.Value = Val(C$)
Combo2.Text = Combo2.List(Val(D$) - 1)
End Sub

Private Sub HScroll1_Change()
'Update the label caption (let the user know what value
'they have selected
Label2.Caption = "Acceleration - " + Str$(HScroll1.Value)
End Sub

Private Sub HScroll2_Change()
'Update the label caption (let the user know what value
'they have selected
Label7.Caption = "Speed - " + Str$(HScroll2.Value)
End Sub


