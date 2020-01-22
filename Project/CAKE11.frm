VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   8535
      LargeChange     =   350
      Left            =   8520
      Max             =   -13067
      TabIndex        =   24
      Top             =   -120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   21450
      Left            =   0
      Picture         =   "CAKE11.frx":0000
      ScaleHeight     =   21390
      ScaleWidth      =   8535
      TabIndex        =   9
      Top             =   0
      Width           =   8595
      Begin VB.CommandButton C1 
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
         TabIndex        =   52
         Top             =   3840
         Width           =   2775
      End
      Begin VB.CommandButton C1 
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
         Left            =   480
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3840
         Width           =   2775
      End
      Begin VB.CommandButton C1 
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
         Left            =   5280
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   16080
         Width           =   2775
      End
      Begin VB.CommandButton C1 
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
         Left            =   480
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   20280
         Width           =   2775
      End
      Begin VB.CommandButton C1 
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
         Left            =   5160
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   20280
         Width           =   2775
      End
      Begin VB.CommandButton C1 
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
         Left            =   600
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   16080
         Width           =   2775
      End
      Begin VB.CommandButton C1 
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
         Left            =   4920
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   12000
         Width           =   2775
      End
      Begin VB.CommandButton C1 
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
         Left            =   600
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   12000
         Width           =   2775
      End
      Begin VB.CommandButton C1 
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
         TabIndex        =   44
         Top             =   7440
         Width           =   2775
      End
      Begin VB.PictureBox Picture12 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   8175
         TabIndex        =   43
         Top             =   21360
         Width           =   8175
      End
      Begin VB.CommandButton C1 
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
         Left            =   600
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   7440
         Width           =   2775
      End
      Begin VB.PictureBox Picture11 
         Height          =   1695
         Left            =   840
         Picture         =   "CAKE11.frx":218FB
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   34
         Top             =   1200
         Width           =   1815
      End
      Begin VB.PictureBox Picture10 
         Height          =   1695
         Left            =   5640
         Picture         =   "CAKE11.frx":231EA
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   33
         Top             =   1200
         Width           =   1815
      End
      Begin VB.PictureBox Picture7 
         Height          =   1695
         Left            =   840
         Picture         =   "CAKE11.frx":2487B
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   29
         Top             =   4920
         Width           =   1815
      End
      Begin VB.PictureBox Picture2 
         Height          =   1695
         Left            =   5640
         Picture         =   "CAKE11.frx":25730
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   28
         Top             =   4920
         Width           =   1815
      End
      Begin VB.PictureBox Picture6 
         Height          =   1695
         Left            =   960
         Picture         =   "CAKE11.frx":26832
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   25
         Top             =   13440
         Width           =   1815
      End
      Begin VB.PictureBox Picture5 
         Height          =   1695
         Left            =   960
         Picture         =   "CAKE11.frx":279D5
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   23
         Top             =   17640
         Width           =   1815
      End
      Begin VB.PictureBox Picture8 
         Height          =   1695
         Left            =   5760
         Picture         =   "CAKE11.frx":28D5E
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   13
         Top             =   9360
         Width           =   1815
      End
      Begin VB.PictureBox Picture9 
         Height          =   1695
         Left            =   960
         Picture         =   "CAKE11.frx":29D39
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   12
         Top             =   9360
         Width           =   1815
      End
      Begin VB.PictureBox Picture3 
         Height          =   1695
         Left            =   5760
         Picture         =   "CAKE11.frx":2AEFD
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   11
         Top             =   13440
         Width           =   1815
      End
      Begin VB.PictureBox Picture4 
         Height          =   1695
         Left            =   5760
         Picture         =   "CAKE11.frx":2BEE6
         ScaleHeight     =   1635
         ScaleWidth      =   1755
         TabIndex        =   10
         Top             =   17640
         Width           =   1815
      End
      Begin VB.Label Label26 
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
         Left            =   1200
         TabIndex        =   42
         Top             =   19920
         Width           =   1455
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT CODE #HFK4"
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
         Left            =   660
         TabIndex        =   41
         Top             =   19560
         Width           =   3135
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INR750.00"
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
         Left            =   5880
         TabIndex        =   40
         Top             =   11520
         Width           =   1455
      End
      Begin VB.Label Label23 
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
         Left            =   5880
         TabIndex        =   39
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FRIUT CAKE"
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
         Left            =   1020
         TabIndex        =   38
         Top             =   3000
         Width           =   1635
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INR600.00"
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
         Left            =   960
         TabIndex        =   37
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CHOCOLATE TREAT"
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
         Left            =   5070
         TabIndex        =   36
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   0
         X2              =   8880
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RED VELET"
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
         Left            =   1095
         TabIndex        =   32
         Top             =   6720
         Width           =   1485
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INR750.00"
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
         Left            =   960
         TabIndex        =   31
         Top             =   7080
         Width           =   1455
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BLACK FOREST"
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
         Left            =   5385
         TabIndex        =   30
         Top             =   6720
         Width           =   2025
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   0
         X2              =   8400
         Y1              =   13200
         Y2              =   13200
      End
      Begin VB.Label Label7 
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
         Left            =   1320
         TabIndex        =   27
         Top             =   15720
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT CODE #HFK2"
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
         Left            =   480
         TabIndex        =   26
         Top             =   15360
         Width           =   3135
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT CODE #HFK13"
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
         Left            =   4860
         TabIndex        =   22
         Top             =   11160
         Width           =   3315
      End
      Begin VB.Label Label4 
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
         Left            =   5760
         TabIndex        =   21
         Top             =   7080
         Width           =   1455
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
         TabIndex        =   20
         Top             =   11520
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VANILA TREAT"
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
         Left            =   960
         TabIndex        =   19
         Top             =   11160
         Width           =   1995
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   2520
         X2              =   5640
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CAKES"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   780
         Left            =   2880
         TabIndex        =   18
         Top             =   0
         Width           =   2385
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   -360
         X2              =   8520
         Y1              =   8280
         Y2              =   8280
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   -120
         X2              =   8760
         Y1              =   17400
         Y2              =   17400
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   4320
         X2              =   4320
         Y1              =   720
         Y2              =   21480
      End
      Begin VB.Label Label8 
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
         Left            =   5880
         TabIndex        =   17
         Top             =   15720
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT CODE #HFK3"
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
         Left            =   5040
         TabIndex        =   16
         Top             =   15360
         Width           =   3135
      End
      Begin VB.Label Label10 
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
         Left            =   5880
         TabIndex        =   15
         Top             =   19920
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT CODE #HFK5"
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
         Left            =   5040
         TabIndex        =   14
         Top             =   19560
         Width           =   3135
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub C1_Click(Index As Integer)
For i = 0 To 9
Form1.Show
Next
End Sub



Private Sub VScroll1_Scroll()
Picture1.Top = VScroll1.Value

End Sub
