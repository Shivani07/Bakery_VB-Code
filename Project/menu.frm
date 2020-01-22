VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "menu.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Height          =   1575
      Left            =   4680
      Picture         =   "menu.frx":163B12
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Height          =   1575
      Left            =   6360
      Picture         =   "menu.frx":164CEB
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Height          =   1575
      Left            =   2160
      Picture         =   "menu.frx":166154
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Height          =   1575
      Left            =   480
      Picture         =   "menu.frx":1676A7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Height          =   1575
      Left            =   3360
      Picture         =   "menu.frx":1692FE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   7560
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2760
      TabIndex        =   5
      Top             =   3240
      Width           =   3255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form8.Show
Form2.Hide
End Sub

Private Sub Command2_Click()
Form3.Show
Form2.Hide
End Sub

Private Sub Command3_Click()
Form9.Show
Form2.Hide
End Sub

Private Sub Command4_Click()
Form7.Show
Form2.Hide
End Sub

Private Sub Command5_Click()
Form6.Show
Form2.Hide
End Sub

