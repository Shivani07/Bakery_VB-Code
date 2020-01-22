VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   Picture         =   "CAKE16.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      LargeChange     =   350
      Left            =   8640
      Max             =   -13067
      TabIndex        =   24
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   22450
      Left            =   0
      Picture         =   "CAKE16.frx":1E009
      ScaleHeight     =   22395
      ScaleWidth      =   8940
      TabIndex        =   9
      Top             =   0
      Width           =   9000
      Begin VB.CommandButton Command15 
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "CAKE16.frx":52BE5
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add to Cart"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   4920
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add to Cart"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   720
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3840
         Width           =   2775
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add to Cart"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   5040
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   16200
         Width           =   2775
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add to Cart"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   720
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   20640
         Width           =   2775
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add to Cart"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   5040
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   20640
         Width           =   2775
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add to Cart"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   720
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   16200
         Width           =   2775
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add to Cart"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   5040
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   12000
         Width           =   2775
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add to Cart"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   960
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   12000
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add to Cart"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   5040
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   7920
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add to Cart"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   840
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   7920
         Width           =   2775
      End
      Begin VB.PictureBox Picture11 
         Height          =   1695
         Left            =   1320
         Picture         =   "CAKE16.frx":53881
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   34
         Top             =   1200
         Width           =   1815
      End
      Begin VB.PictureBox Picture10 
         Height          =   1695
         Left            =   5520
         Picture         =   "CAKE16.frx":5EDDB
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   33
         Top             =   1200
         Width           =   1815
      End
      Begin VB.PictureBox Picture7 
         Height          =   1695
         Left            =   1320
         Picture         =   "CAKE16.frx":686D9
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   29
         Top             =   5400
         Width           =   1815
      End
      Begin VB.PictureBox Picture2 
         Height          =   1695
         Left            =   5520
         Picture         =   "CAKE16.frx":7243F
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   28
         Top             =   5400
         Width           =   1815
      End
      Begin VB.PictureBox Picture6 
         Height          =   1695
         Left            =   1320
         Picture         =   "CAKE16.frx":7DB81
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   25
         Top             =   13680
         Width           =   1815
      End
      Begin VB.PictureBox Picture5 
         Height          =   1695
         Left            =   1320
         Picture         =   "CAKE16.frx":86AC3
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   23
         Top             =   18120
         Width           =   1815
      End
      Begin VB.PictureBox Picture8 
         Height          =   1695
         Left            =   5400
         Picture         =   "CAKE16.frx":91EA1
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   13
         Top             =   9480
         Width           =   1815
      End
      Begin VB.PictureBox Picture9 
         Height          =   1695
         Left            =   1560
         Picture         =   "CAKE16.frx":9D103
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   12
         Top             =   9480
         Width           =   1815
      End
      Begin VB.PictureBox Picture3 
         Height          =   1695
         Left            =   5520
         Picture         =   "CAKE16.frx":A996D
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   11
         Top             =   13680
         Width           =   1815
      End
      Begin VB.PictureBox Picture4 
         Height          =   1695
         Left            =   5520
         Picture         =   "CAKE16.frx":B423F
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   10
         Top             =   18120
         Width           =   1815
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. 42.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   270
         Left            =   1530
         TabIndex        =   42
         Top             =   20280
         Width           =   1275
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "CHOCO SQUARES"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   270
         Left            =   1050
         TabIndex        =   41
         Top             =   19920
         Width           =   2355
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. 25.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   270
         Left            =   5730
         TabIndex        =   40
         Top             =   11640
         Width           =   1275
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. 35.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   270
         Left            =   5850
         TabIndex        =   39
         Top             =   3360
         Width           =   1275
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FRUIT PASTRY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   270
         Left            =   1185
         TabIndex        =   38
         Top             =   3000
         Width           =   2025
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. 32.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   270
         Left            =   1530
         TabIndex        =   37
         Top             =   3360
         Width           =   1275
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FOREST CRUNCH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   270
         Left            =   5265
         TabIndex        =   36
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   0
         X2              =   8880
         Y1              =   8880
         Y2              =   8880
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DARK CHOCOLATE PASTRY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   375
         TabIndex        =   32
         Top             =   7200
         Width           =   3645
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. 55.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   270
         Left            =   1530
         TabIndex        =   31
         Top             =   7560
         Width           =   1275
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CHOCO CHIP BITE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   5160
         TabIndex        =   30
         Top             =   7200
         Width           =   2475
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   120
         X2              =   9000
         Y1              =   13200
         Y2              =   13200
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. 40.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   270
         Left            =   1530
         TabIndex        =   27
         Top             =   15840
         Width           =   1275
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CHOCO NUTS PASTRY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   690
         TabIndex        =   26
         Top             =   15480
         Width           =   2955
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NUT PINE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   270
         Left            =   5730
         TabIndex        =   22
         Top             =   11280
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. 40.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   270
         Left            =   5850
         TabIndex        =   21
         Top             =   7560
         Width           =   1275
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   -120
         X2              =   8760
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. 35.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   270
         Left            =   1890
         TabIndex        =   20
         Top             =   11640
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NUTTY ALMO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   270
         Left            =   1545
         TabIndex        =   19
         Top             =   11280
         Width           =   1785
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFC0FF&
         BorderWidth     =   3
         X1              =   2280
         X2              =   6720
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASTERIES"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   780
         Left            =   2580
         TabIndex        =   18
         Top             =   0
         Width           =   3945
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   240
         X2              =   9120
         Y1              =   17400
         Y2              =   17400
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   4320
         X2              =   4320
         Y1              =   840
         Y2              =   22320
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. 60.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   270
         Left            =   5850
         TabIndex        =   17
         Top             =   15840
         Width           =   1275
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GERMAN TOAST"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   270
         Left            =   5400
         TabIndex        =   16
         Top             =   15480
         Width           =   2175
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rs. 42.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   270
         Left            =   5850
         TabIndex        =   15
         Top             =   20280
         Width           =   1275
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLUSTER RED"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   270
         Left            =   5430
         TabIndex        =   14
         Top             =   19920
         Width           =   1875
      End
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0FF&
      Caption         =   "DETAILS"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   11280
      Width           =   2175
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add to Cart"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2280
      MaskColor       =   &H00404040&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10800
      Width           =   2175
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0FF&
      Caption         =   "DETAILS"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10680
      Width           =   2175
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add to Cart"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      MaskColor       =   &H00404040&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   10200
      Width           =   2175
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   0
      X2              =   8880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT CODE #HFK1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   2040
      TabIndex        =   7
      Top             =   10080
      Width           =   3135
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INR650.00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   2880
      TabIndex        =   6
      Top             =   10440
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INR650.00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   6840
      TabIndex        =   3
      Top             =   9840
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT CODE #HFK1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   9240
      Width           =   3135
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INR650.00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   1080
      TabIndex        =   0
      Top             =   9600
      Width           =   1455
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
Form10.List1.AddItem ("Fruit Pastry")
Form10.List2.AddItem ("32")
End Sub

Private Sub Command10_Click(Index As Integer)
Form10.List1.AddItem ("Cluster Red")
Form10.List2.AddItem ("42")
End Sub

Private Sub Command15_Click(Index As Integer)
Form2.Show
Form6.Hide
End Sub

Private Sub Command2_Click(Index As Integer)
Form10.List1.AddItem ("Forest Crunch")
Form10.List2.AddItem ("35")
End Sub

Private Sub Command3_Click(Index As Integer)
Form10.List1.AddItem ("Dark Chocolate Pastry")
Form10.List2.AddItem ("55")
End Sub

Private Sub Command4_Click(Index As Integer)
Form10.List1.AddItem ("Choco Chip Bite")
Form10.List2.AddItem ("40")
End Sub

Private Sub Command5_Click(Index As Integer)
Form10.List1.AddItem ("Nutty Almo")
Form10.List2.AddItem ("35")
End Sub

Private Sub Command6_Click(Index As Integer)
Form10.List1.AddItem ("Nut Pine")
Form10.List2.AddItem ("25")
End Sub

Private Sub Command7_Click(Index As Integer)
Form10.List1.AddItem ("Choco Nuts Pastry")
Form10.List2.AddItem ("40")
End Sub

Private Sub Command8_Click(Index As Integer)
Form10.List1.AddItem ("German Toast")
Form10.List2.AddItem ("60")
End Sub

Private Sub Command9_Click(Index As Integer)
Form10.List1.AddItem ("Choco Sqaures")
Form10.List2.AddItem ("42")
End Sub

Private Sub VScroll1_Scroll()
Picture1.Top = VScroll1.Value
End Sub
