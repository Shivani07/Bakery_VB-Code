VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   Picture         =   "CAKE15.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      LargeChange     =   450
      Left            =   8640
      Max             =   -7750
      SmallChange     =   50
      TabIndex        =   20
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   22000
      Left            =   0
      Picture         =   "CAKE15.frx":1E009
      ScaleHeight     =   21945
      ScaleWidth      =   8940
      TabIndex        =   9
      Top             =   0
      Width           =   9000
      Begin VB.CommandButton Command15 
         Height          =   495
         Index           =   1
         Left            =   120
         Picture         =   "CAKE15.frx":4223B
         Style           =   1  'Graphical
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   3840
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
         Left            =   840
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   15840
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
         Left            =   840
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   15840
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
         TabIndex        =   39
         Top             =   11880
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
         Left            =   840
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   11880
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
         TabIndex        =   37
         Top             =   7800
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
         TabIndex        =   31
         Top             =   7800
         Width           =   2775
      End
      Begin VB.PictureBox Picture11 
         Height          =   1695
         Left            =   1320
         Picture         =   "CAKE15.frx":42ED7
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   30
         Top             =   1200
         Width           =   1815
      End
      Begin VB.PictureBox Picture10 
         Height          =   1695
         Left            =   5520
         Picture         =   "CAKE15.frx":4D9CD
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   29
         Top             =   1200
         Width           =   1815
      End
      Begin VB.PictureBox Picture7 
         Height          =   1695
         Left            =   1320
         Picture         =   "CAKE15.frx":5806B
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   25
         Top             =   5160
         Width           =   1815
      End
      Begin VB.PictureBox Picture2 
         Height          =   1695
         Left            =   5520
         Picture         =   "CAKE15.frx":62271
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   24
         Top             =   5160
         Width           =   1815
      End
      Begin VB.PictureBox Picture6 
         Height          =   1695
         Left            =   1320
         Picture         =   "CAKE15.frx":6C033
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   21
         Top             =   13320
         Width           =   1815
      End
      Begin VB.PictureBox Picture8 
         Height          =   1695
         Left            =   5520
         Picture         =   "CAKE15.frx":75395
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   12
         Top             =   9240
         Width           =   1815
      End
      Begin VB.PictureBox Picture9 
         Height          =   1695
         Left            =   1320
         Picture         =   "CAKE15.frx":7F523
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   11
         Top             =   9240
         Width           =   1815
      End
      Begin VB.PictureBox Picture3 
         Height          =   1695
         Left            =   5520
         Picture         =   "CAKE15.frx":891CD
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   10
         Top             =   13320
         Width           =   1815
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
         Left            =   5970
         TabIndex        =   36
         Top             =   11400
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
         TabIndex        =   35
         Top             =   3360
         Width           =   1275
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CHOCO DELIGHT"
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
         Left            =   1065
         TabIndex        =   34
         Top             =   3000
         Width           =   2265
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
         TabIndex        =   33
         Top             =   3360
         Width           =   1275
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BLACK MUFFINS"
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
         Left            =   5445
         TabIndex        =   32
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   240
         X2              =   9120
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DARK DELIGHT"
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
         Left            =   1170
         TabIndex        =   28
         Top             =   6960
         Width           =   2055
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
         Left            =   1650
         TabIndex        =   27
         Top             =   7320
         Width           =   1275
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COCO NUTTIES"
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
         Left            =   5475
         TabIndex        =   26
         Top             =   6960
         Width           =   2085
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   0
         X2              =   9000
         Y1              =   12840
         Y2              =   12840
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
         TabIndex        =   23
         Top             =   15480
         Width           =   1275
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ROYAL BITE"
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
         Left            =   1335
         TabIndex        =   22
         Top             =   15120
         Width           =   1665
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
         Left            =   5850
         TabIndex        =   19
         Top             =   11040
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
         TabIndex        =   18
         Top             =   7320
         Width           =   1275
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   120
         X2              =   9000
         Y1              =   4680
         Y2              =   4680
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
         Left            =   1650
         TabIndex        =   17
         Top             =   11400
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
         Left            =   1305
         TabIndex        =   16
         Top             =   11040
         Width           =   1785
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         X1              =   2160
         X2              =   6600
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CHOCOLATES"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   780
         Left            =   2160
         TabIndex        =   15
         Top             =   0
         Width           =   4485
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   -360
         X2              =   9000
         Y1              =   8760
         Y2              =   8760
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   4320
         X2              =   4320
         Y1              =   720
         Y2              =   18000
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
         TabIndex        =   14
         Top             =   15480
         Width           =   1275
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CHOCO HEARTS"
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
         Left            =   5295
         TabIndex        =   13
         Top             =   15120
         Width           =   2145
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
      Index           =   0
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
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
Form10.List1.AddItem ("Choco Delight")
Form10.List2.AddItem ("32")
End Sub

Private Sub Command15_Click(Index As Integer)
Form2.Show
Form7.Hide
End Sub

Private Sub Command2_Click(Index As Integer)
Form10.List1.AddItem ("Black Muffins")
Form10.List2.AddItem ("35")
End Sub

Private Sub Command3_Click(Index As Integer)
Form10.List1.AddItem ("Dark Delight")
Form10.List2.AddItem ("55")
End Sub

Private Sub Command4_Click(Index As Integer)
Form10.List1.AddItem ("Coco Nutties")
Form10.List2.AddItem ("40")
End Sub

Private Sub Command5_Click(Index As Integer)
Form10.List1.AddItem ("Nutty Almo")
Form10.List2.AddItem ("35")
End Sub

Private Sub Command6_Click(Index As Integer)
Form10.List1.AddItem ("Nutty Pine")
Form10.List2.AddItem ("25")
End Sub

Private Sub Command7_Click(Index As Integer)
Form10.List1.AddItem ("Royal Bite")
Form10.List2.AddItem ("40")
End Sub

Private Sub Command8_Click(Index As Integer)
Form10.List1.AddItem ("Choco Hearts")
Form10.List2.AddItem ("60")
End Sub

Private Sub VScroll1_Scroll()
Picture1.Top = VScroll1.Value
End Sub
