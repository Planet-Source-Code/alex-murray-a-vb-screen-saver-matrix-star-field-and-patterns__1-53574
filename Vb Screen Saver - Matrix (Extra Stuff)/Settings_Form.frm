VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2985
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   2985
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   285
      Max             =   7
      Min             =   1
      TabIndex        =   6
      Top             =   1170
      Value           =   1
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1815
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1815
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   285
      Max             =   200
      Min             =   1
      TabIndex        =   1
      Top             =   345
      Value           =   1
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "Speed"
      Height          =   255
      Left            =   165
      TabIndex        =   9
      Top             =   930
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Slower"
      Height          =   255
      Left            =   45
      TabIndex        =   8
      Top             =   1410
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Faster"
      Height          =   375
      Left            =   2325
      TabIndex        =   7
      Top             =   1410
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Many"
      Height          =   375
      Left            =   2325
      TabIndex        =   3
      Top             =   585
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "One"
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   585
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Amount"
      Height          =   255
      Left            =   165
      TabIndex        =   0
      Top             =   105
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Open FIlepos$ For Output As #1
Print #1, Str$(HScroll1.Value)
Print #1, Str$(HScroll2.Value)
Close
End
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
On Error Resume Next
Open FIlepos$ For Input As #1
Input #1, A$
Input #1, B$
Close
If A$ = "" Then A$ = "30"
If B$ = "" Then B$ = "2"
HScroll1.Value = Val(A$)
HScroll2.Value = Val(B$)
End Sub

Private Sub HScroll1_Change()
Label2.Caption = "Amount - " + Str$(HScroll1.Value)
End Sub

Private Sub HScroll2_Change()
Label7.Caption = "Speed - " + Str$(HScroll2.Value)
End Sub


