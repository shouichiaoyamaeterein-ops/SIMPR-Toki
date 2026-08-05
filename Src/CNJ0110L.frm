VERSION 5.00
Begin VB.Form CNJ0110L 
   Caption         =   "ì‡é¶ï\"
   ClientHeight    =   4335
   ClientLeft      =   -270
   ClientTop       =   1005
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   15420
   Begin VB.Label B1_KMaiSa 
      AutoSize        =   -1  'True
      Caption         =   "ZZZZ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6000
      TabIndex        =   62
      Top             =   2520
      Width           =   420
   End
   Begin VB.Label B1_KMai 
      AutoSize        =   -1  'True
      Caption         =   "ZZZ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5250
      TabIndex        =   61
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label B1_TYJ 
      AutoSize        =   -1  'True
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   12510
      TabIndex        =   60
      Top             =   1950
      Width           =   105
   End
   Begin VB.Label B1_TY 
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   12120
      TabIndex        =   59
      Top             =   1950
      Width           =   210
   End
   Begin VB.Label B1_NJSU 
      AutoSize        =   -1  'True
      Caption         =   "99,999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   10380
      TabIndex        =   58
      Top             =   1920
      Width           =   1050
   End
   Begin VB.Line B1_K_LMTe2 
      Index           =   0
      X1              =   12750
      X2              =   12030
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Line B1_K_LSTe 
      BorderWidth     =   4
      Index           =   0
      X1              =   12720
      X2              =   12060
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Line B1_K_LMTe 
      Index           =   0
      X1              =   12750
      X2              =   12060
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line B1_K_LUTe 
      BorderWidth     =   5
      Index           =   0
      X1              =   12750
      X2              =   12060
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line B1_K_LTe2 
      BorderWidth     =   5
      Index           =   0
      X1              =   12720
      X2              =   12720
      Y1              =   1590
      Y2              =   2190
   End
   Begin VB.Line B1_K_LTe1 
      Index           =   0
      X1              =   12420
      X2              =   12420
      Y1              =   1710
      Y2              =   2190
   End
   Begin VB.Line B1_K_LINENJS 
      BorderWidth     =   4
      Index           =   0
      X1              =   12030
      X2              =   10080
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Line B1_K_LNS 
      BorderWidth     =   5
      X1              =   10080
      X2              =   9780
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Line B1_K_LINENJT 
      Index           =   0
      X1              =   12030
      X2              =   12030
      Y1              =   1590
      Y2              =   2190
   End
   Begin VB.Line B1_K_LINENJM 
      Index           =   0
      X1              =   12060
      X2              =   10080
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Label B1_Lb_NJm 
      AutoSize        =   -1  'True
      Caption         =   "ìñèâédä|ì‡é¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   0
      Left            =   11040
      TabIndex        =   55
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label B1_Lb_NJ 
      AutoSize        =   -1  'True
      Caption         =   "99îN 99åé"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   10140
      TabIndex        =   54
      Top             =   1650
      Width           =   915
   End
   Begin VB.Line B1_K_LINENJU 
      BorderWidth     =   4
      Index           =   0
      X1              =   12030
      X2              =   10080
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label B1_Lb_Nai4 
      AutoSize        =   -1  'True
      Caption         =   "ïÒ"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   9870
      TabIndex        =   53
      Top             =   2055
      Width           =   150
   End
   Begin VB.Label B1_Lb_Nai3 
      AutoSize        =   -1  'True
      Caption         =   "èÓ"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   9870
      TabIndex        =   52
      Top             =   1890
      Width           =   150
   End
   Begin VB.Label B1_Lb_Nai2 
      AutoSize        =   -1  'True
      Caption         =   "é¶"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   9870
      TabIndex        =   51
      Top             =   1755
      Width           =   150
   End
   Begin VB.Label B1_Lb_Nai1 
      AutoSize        =   -1  'True
      Caption         =   "ì‡"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   9870
      TabIndex        =   50
      Top             =   1620
      Width           =   150
   End
   Begin VB.Line B1_K_LNJT 
      X1              =   10080
      X2              =   10080
      Y1              =   1590
      Y2              =   2190
   End
   Begin VB.Line B1_K_LNJ 
      BorderWidth     =   5
      X1              =   10080
      X2              =   9780
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line B1_K_LINE 
      BorderWidth     =   8
      Index           =   26
      X1              =   9780
      X2              =   9780
      Y1              =   1590
      Y2              =   2190
   End
   Begin VB.Label B1_Kanri 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8910
      TabIndex        =   49
      Top             =   1920
      Width           =   630
   End
   Begin VB.Label B1_Lb_Kanri 
      AutoSize        =   -1  'True
      Caption         =   "ä«óùî‘çÜ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   8850
      TabIndex        =   48
      Top             =   1650
      Width           =   720
   End
   Begin VB.Line B1_K_LINE 
      Index           =   25
      X1              =   8670
      X2              =   8670
      Y1              =   1590
      Y2              =   2190
   End
   Begin VB.Label B1_Syuyo 
      AutoSize        =   -1  'True
      Caption         =   "99999"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7920
      TabIndex        =   47
      Top             =   1920
      Width           =   525
   End
   Begin VB.Label B1_Lb_Syuyo 
      AutoSize        =   -1  'True
      Caption         =   "é˚ óe êî"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   7770
      TabIndex        =   46
      Top             =   1650
      Width           =   720
   End
   Begin VB.Line B1_K_LINE 
      BorderWidth     =   8
      Index           =   0
      X1              =   9780
      X2              =   570
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line B1_K_LINE 
      Index           =   1
      X1              =   9780
      X2              =   570
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line B1_K_LINE 
      Index           =   24
      X1              =   7680
      X2              =   7680
      Y1              =   1590
      Y2              =   2190
   End
   Begin VB.Label B1_Cycle3 
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7365
      TabIndex        =   45
      Top             =   1920
      Width           =   210
   End
   Begin VB.Label B1_Cycle2 
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   44
      Top             =   1920
      Width           =   210
   End
   Begin VB.Label B1_Cycle1 
      AutoSize        =   -1  'True
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   43
      Top             =   1920
      Width           =   105
   End
   Begin VB.Line B1_K_LINE 
      Index           =   23
      X1              =   7680
      X2              =   6750
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Line B1_K_LINE 
      Index           =   22
      X1              =   7350
      X2              =   7350
      Y1              =   1710
      Y2              =   2190
   End
   Begin VB.Line B1_K_LINE 
      Index           =   21
      X1              =   7020
      X2              =   7020
      Y1              =   1710
      Y2              =   2190
   End
   Begin VB.Label B1_Lb_Cycle1 
      AutoSize        =   -1  'True
      Caption         =   "ì˙"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   6840
      TabIndex        =   40
      Top             =   1730
      Width           =   150
   End
   Begin VB.Line B1_K_LINE 
      Index           =   20
      X1              =   6750
      X2              =   6750
      Y1              =   1590
      Y2              =   2190
   End
   Begin VB.Line B1_K_LINE 
      Index           =   19
      X1              =   6060
      X2              =   6060
      Y1              =   1590
      Y2              =   2190
   End
   Begin VB.Line B1_K_LINE 
      Index           =   18
      X1              =   5580
      X2              =   5580
      Y1              =   1590
      Y2              =   2190
   End
   Begin VB.Label B1_Seban 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6270
      TabIndex        =   38
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label B1_Lb_Seban 
      AutoSize        =   -1  'True
      Caption         =   "îwî‘çÜ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6150
      TabIndex        =   37
      Top             =   1650
      Width           =   540
   End
   Begin VB.Label B1_Nouku 
      AutoSize        =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5760
      TabIndex        =   36
      Top             =   1920
      Width           =   105
   End
   Begin VB.Label B1_Lb_Nouku 
      AutoSize        =   -1  'True
      Caption         =   "î[ãÊ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5640
      TabIndex        =   35
      Top             =   1650
      Width           =   360
   End
   Begin VB.Line B1_K_LINE 
      Index           =   17
      X1              =   5130
      X2              =   5130
      Y1              =   1590
      Y2              =   2190
   End
   Begin VB.Label B1_Tanto 
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5250
      TabIndex        =   34
      Top             =   1920
      Width           =   210
   End
   Begin VB.Label B1_Lb_Tanto 
      AutoSize        =   -1  'True
      Caption         =   "íSìñ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   5160
      TabIndex        =   33
      Top             =   1650
      Width           =   360
   End
   Begin VB.Line B1_K_LINE 
      Index           =   16
      X1              =   4650
      X2              =   4650
      Y1              =   1590
      Y2              =   2190
   End
   Begin VB.Label B1_Lb_Seibi 
      AutoSize        =   -1  'True
      Caption         =   "êÆîı"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4710
      TabIndex        =   32
      Top             =   1650
      Width           =   360
   End
   Begin VB.Label B1_Seibi 
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4800
      TabIndex        =   31
      Top             =   1920
      Width           =   210
   End
   Begin VB.Line B1_K_LINE 
      Index           =   14
      X1              =   4170
      X2              =   4170
      Y1              =   1590
      Y2              =   2190
   End
   Begin VB.Line B1_K_LINE 
      Index           =   13
      X1              =   3300
      X2              =   3300
      Y1              =   1590
      Y2              =   2190
   End
   Begin VB.Line B1_K_LINEDayY 
      X1              =   7860
      X2              =   7380
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Label B1_NSU 
      AutoSize        =   -1  'True
      Caption         =   "ZZZZZ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   7380
      TabIndex        =   30
      Top             =   2550
      Width           =   450
   End
   Begin VB.Line B1_K_LINEDay 
      Index           =   0
      X1              =   7860
      X2              =   7860
      Y1              =   2190
      Y2              =   2790
   End
   Begin VB.Label B1_Lb_Day 
      Alignment       =   1  'âEëµÇ¶
      Caption         =   "XXì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   0
      Left            =   7470
      TabIndex        =   29
      Top             =   2250
      Width           =   345
   End
   Begin VB.Line B1_K_LINE 
      Index           =   15
      X1              =   7380
      X2              =   7380
      Y1              =   2190
      Y2              =   2790
   End
   Begin VB.Label B1_Lb_Day6 
      AutoSize        =   -1  'True
      Caption         =   "Å®"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   7200
      TabIndex        =   28
      Top             =   2460
      Width           =   165
   End
   Begin VB.Label B1_Lb_Day5 
      AutoSize        =   -1  'True
      Caption         =   "ó "
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   6990
      TabIndex        =   27
      Top             =   2550
      Width           =   180
   End
   Begin VB.Label B1_Lb_Day4 
      AutoSize        =   -1  'True
      Caption         =   "êî"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   6990
      TabIndex        =   26
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label B1_Lb_Day3 
      AutoSize        =   -1  'True
      Caption         =   "ÇË"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   6840
      TabIndex        =   25
      Top             =   2580
      Width           =   165
   End
   Begin VB.Label B1_Lb_Day2 
      AutoSize        =   -1  'True
      Caption         =   "ìñ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   6840
      TabIndex        =   24
      Top             =   2400
      Width           =   165
   End
   Begin VB.Label B1_Lb_Day1 
      AutoSize        =   -1  'True
      Caption         =   "ì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   6840
      TabIndex        =   23
      Top             =   2250
      Width           =   165
   End
   Begin VB.Line B1_K_LINEDayS 
      BorderWidth     =   4
      X1              =   7860
      X2              =   6780
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Line B1_K_LINE 
      BorderWidth     =   8
      Index           =   10
      X1              =   6780
      X2              =   5010
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Line B1_K_LINE 
      BorderWidth     =   8
      Index           =   9
      X1              =   6780
      X2              =   6780
      Y1              =   2190
      Y2              =   2790
   End
   Begin VB.Line B1_K_LINE 
      Index           =   8
      X1              =   5820
      X2              =   5820
      Y1              =   2190
      Y2              =   2790
   End
   Begin VB.Label B1_Lb_KMaiZ 
      AutoSize        =   -1  'True
      Caption         =   "+/-ëOâÒÇ∆ÇÃç∑"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   5880
      TabIndex        =   22
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label B1_Lb_KMai 
      AutoSize        =   -1  'True
      Caption         =   "Ç©ÇÒÇŒÇÒñáêî"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   5070
      TabIndex        =   21
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label B1_CMai 
      AutoSize        =   -1  'True
      Caption         =   "ZZZ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4470
      TabIndex        =   20
      Top             =   2520
      Width           =   315
   End
   Begin VB.Line B1_K_LINE 
      BorderWidth     =   8
      Index           =   7
      X1              =   5010
      X2              =   5010
      Y1              =   2190
      Y2              =   2790
   End
   Begin VB.Label B1_Lb_HMai 
      AutoSize        =   -1  'True
      Caption         =   "∂∞ƒﬁî≠çsêî"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   4320
      TabIndex        =   19
      Top             =   2280
      Width           =   750
   End
   Begin VB.Line B1_K_LINE 
      Index           =   6
      X1              =   4260
      X2              =   4260
      Y1              =   2190
      Y2              =   2790
   End
   Begin VB.Line B1_K_LINE 
      Index           =   3
      X1              =   6780
      X2              =   960
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Line B1_K_LINE 
      BorderWidth     =   4
      Index           =   5
      X1              =   960
      X2              =   960
      Y1              =   2190
      Y2              =   2790
   End
   Begin VB.Line B1_K_LINE 
      BorderWidth     =   4
      Index           =   4
      X1              =   5010
      X2              =   960
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Label B1_Lb_Comment 
      AutoSize        =   -1  'True
      Caption         =   "ÉR ÉÅ Éì Ég"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2130
      TabIndex        =   18
      Top             =   2250
      Width           =   990
   End
   Begin VB.Label B1_Comment 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   17
      Top             =   2520
      Width           =   3150
   End
   Begin VB.Label B1_Ukeir 
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4290
      TabIndex        =   16
      Top             =   1920
      Width           =   210
   End
   Begin VB.Line B1_K_LINE 
      BorderWidth     =   8
      Index           =   2
      X1              =   9780
      X2              =   570
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   5
      Index           =   3
      X1              =   8850
      X2              =   8850
      Y1              =   690
      Y2              =   1410
   End
   Begin VB.Label H1_KTORNM 
      AutoSize        =   -1  'True
      Caption         =   "ÇwÇwÇwÇwÇwÇwÇwÇwÇwÇwÇwÇwÇwÇwÇwÇwÇwÇwÇwÇw"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   15
      Top             =   1050
      Width           =   5700
   End
   Begin VB.Label H1_KTorcd 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "(XXXXXXX)"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   14
      Top             =   1110
      Width           =   945
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "å_ñÒêÊ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1500
      TabIndex        =   13
      Top             =   780
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "çHãÊ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   690
      TabIndex        =   12
      Top             =   780
      Width           =   390
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   5
      Index           =   0
      X1              =   570
      X2              =   8850
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Label B1_Lb_Ukeir 
      AutoSize        =   -1  'True
      Caption         =   "éÛì¸"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4230
      TabIndex        =   11
      Top             =   1650
      Width           =   360
   End
   Begin VB.Line B1_K_LINE 
      BorderWidth     =   8
      Index           =   12
      X1              =   570
      X2              =   570
      Y1              =   1590
      Y2              =   2190
   End
   Begin VB.Label B1_Hinbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   660
      TabIndex        =   10
      Top             =   1920
      Width           =   2625
   End
   Begin VB.Label H1_Kouku 
      AutoSize        =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   810
      TabIndex        =   9
      Top             =   1110
      Width           =   105
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   5
      Index           =   1
      X1              =   570
      X2              =   570
      Y1              =   690
      Y2              =   1410
   End
   Begin VB.Label B1_NTorcd 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   8
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label B1_Lb_Torcd 
      AutoSize        =   -1  'True
      Caption         =   "î[ì¸êÊ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3420
      TabIndex        =   7
      Top             =   1650
      Width           =   540
   End
   Begin VB.Label B1_Lb_Hinbn 
      AutoSize        =   -1  'True
      Caption         =   "ïiÅ@î‘"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1590
      TabIndex        =   6
      Top             =   1650
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ï≈"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   15135
      TabIndex        =   5
      Top             =   675
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "î≠çsì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   12690
      TabIndex        =   4
      Top             =   690
      Width           =   540
   End
   Begin VB.Label H1_PAGE 
      Alignment       =   1  'âEëµÇ¶
      Caption         =   "ZZ9"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   14475
      TabIndex        =   3
      Top             =   690
      Width           =   615
   End
   Begin VB.Label H1_HYMD 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD HH:MM"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   13335
      TabIndex        =   2
      Top             =   690
      Width           =   1440
   End
   Begin VB.Label H1_Naiji 
      AutoSize        =   -1  'True
      Caption         =   "ÅyíËèÌÅz"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9300
      TabIndex        =   1
      Top             =   1050
      Width           =   1260
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   5
      Index           =   2
      X1              =   570
      X2              =   8850
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Label B1_Lb_Cycle 
      AutoSize        =   -1  'True
      Caption         =   "î¿ ì¸ âÒ êî"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   4.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   90
      Left            =   7020
      TabIndex        =   39
      Top             =   1620
      Width           =   660
   End
   Begin VB.Label B1_Lb_Cycle3 
      AutoSize        =   -1  'True
      Caption         =   "å„"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   7470
      TabIndex        =   42
      Top             =   1730
      Width           =   150
   End
   Begin VB.Label B1_Lb_Cycle2 
      AutoSize        =   -1  'True
      Caption         =   "âÒ"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   7140
      TabIndex        =   41
      Top             =   1730
      Width           =   150
   End
   Begin VB.Label B1_Lb_Teki 
      AutoSize        =   -1  'True
      Caption         =   "ìKópäJén"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   4.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   90
      Index           =   0
      Left            =   12240
      TabIndex        =   56
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label B1_Lb_TekiY 
      AutoSize        =   -1  'True
      Caption         =   "åé"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   0
      Left            =   12180
      TabIndex        =   0
      Top             =   1730
      Width           =   150
   End
   Begin VB.Label B1_Lb_TekiJ 
      AutoSize        =   -1  'True
      Caption         =   "è{"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   0
      Left            =   12540
      TabIndex        =   57
      Top             =   1730
      Width           =   150
   End
End
Attribute VB_Name = "CNJ0110L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  ì‡é¶ï\
'**       ÉtÉHÅ[ÉÄID    :  CNJ0110L
'**       èàóùäTóv      :  ópéÜ: A3â°
'**
'**       çÏ  ê¨  ì˙    :  2003/12/17
'**       ïœ  çX  ì˙    :  2006/01/17  ÉTÉCÉNÉãåÖêîëùâ¡Ç…î∫Ç§èCê≥
'**                        2006/01/25  ì˙ïtññì˙à»ç~ÇÃï\é¶Ç™è¡Ç¶ÇƒÇ¢Ç»Ç¢ÇÃÇèCê≥
'**
'**********************************************************'
    Dim wCnt            As Long
    Dim OutStr          As String   ' CSVèoóÕì‡óe
    Dim mPrtCnt         As Long

    Dim PrtTime         As String   ' î≠çsì˙
    Dim GokeiSu(36)     As Long     'çáåv
    
    Dim Line_End        As Long
    Dim Col_Width       As Long
    Dim Day_Width       As Long
    Dim End_Day         As Long
    Dim Nj_Width        As Long
    Dim Teki_Width      As Long
    
    Dim Brk_Kouku       As String
    Dim Brk_KTorcd      As String
    Dim Brk_Naiji       As Integer

    Dim w_Int           As Integer  '2006/01/17
'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub Form_Load()

' àÛç¸ëŒè€åèêîÅ@ëﬁî
    mPrtCnt = CisDB.RecordCount
' ëŒè€îNåé(ì‡é¶îNåé)ÇÃññì˙ÇãÅÇﬂÇÈ
    With CisFun
        .Date_RsFormat = fDD
        .Date_End NTT.ì‡é¶îNåé
        End_Day = .Val2(.Date_Result)
    End With
' ïiî‘ï“èWèâä˙ílÇãÅÇﬂÇÈ   ïiî‘ñ¢ìoò^Ç≈Ç‡ï“èWÇµÇΩÇ¢éûÇÕÉRÉÅÉìÉgÉAÉEÉgÇ∑ÇÈ
'    Call EditHinbnInit

    B1_Lb_Hinbn.Tag = B1_Lb_Hinbn
    B1_Lb_Torcd.Tag = B1_Lb_Torcd
    B1_Lb_Ukeir.Tag = B1_Lb_Ukeir
    B1_Lb_Seibi.Tag = B1_Lb_Seibi
    B1_Lb_Tanto.Tag = B1_Lb_Tanto
    B1_Lb_Nouku.Tag = B1_Lb_Nouku
    B1_Lb_Seban.Tag = B1_Lb_Seban
    B1_Lb_Cycle.Tag = B1_Lb_Cycle
    B1_Lb_Cycle1.Tag = B1_Lb_Cycle1
    B1_Lb_Cycle2.Tag = B1_Lb_Cycle2
    B1_Lb_Cycle3.Tag = B1_Lb_Cycle3
    B1_Lb_Syuyo.Tag = B1_Lb_Syuyo
    B1_Lb_Kanri.Tag = B1_Lb_Kanri
    B1_Lb_Nai1.Tag = B1_Lb_Nai1
    B1_Lb_Nai2.Tag = B1_Lb_Nai2
    B1_Lb_Nai3.Tag = B1_Lb_Nai3
    B1_Lb_Nai4.Tag = B1_Lb_Nai4
    B1_Lb_Comment.Tag = B1_Lb_Comment
    B1_Lb_HMai.Tag = B1_Lb_HMai
    B1_Lb_KMai.Tag = B1_Lb_KMai
    B1_Lb_KMaiZ.Tag = B1_Lb_KMaiZ
    B1_Lb_Day1.Tag = B1_Lb_Day1
    B1_Lb_Day2.Tag = B1_Lb_Day2
    B1_Lb_Day3.Tag = B1_Lb_Day3
    B1_Lb_Day4.Tag = B1_Lb_Day4
    B1_Lb_Day5.Tag = B1_Lb_Day5
    B1_Lb_Day6.Tag = B1_Lb_Day6

'A3ópèâä˙ê›íËÅiçÄñ⁄à⁄ìÆÅj
    Line_End = 23240        'ÉâÉCÉìèIà íu(X)
    Col_Width = 510         'Colïù
    Day_Width = B1_K_LINEDayY.X1 - B1_K_LINEDayY.X2 + 30       'ì˙ïtïù
    
    H1_HYMD.Left = H1_HYMD.Left + 7200          'î≠çsì˙
    Label2.Left = Label2.Left + 7200            '"î≠çsì˙"
    H1_PAGE.Left = H1_PAGE.Left + 7700          'ï≈
    Label3.Left = Label3.Left + 7700            '"ï≈"
    
    ' ì˙èÓïÒ
    B1_K_LINEDayS.X1 = B1_K_LINEDayS.X1 + Day_Width * 30
    B1_K_LINEDayY.X1 = B1_K_LINEDayY.X1 + Day_Width * 30
    B1_Lb_Day(0).Tag = " 1ì˙"
    For gLong = 1 To 30
        Load B1_Lb_Day(gLong)
        B1_Lb_Day(gLong).Visible = True
        If gLong <= End_Day Then
            B1_Lb_Day(gLong).Tag = CisFun.RSetFld(gLong + 1, 2, "##") & "ì˙"
    '--------------------------------------------------------------- 2006/01/25 Insert
        Else
            B1_Lb_Day(gLong).Tag = ""
    '--------------------------------------------------------------- 2006/01/25 Insert End
        End If
        B1_Lb_Day(gLong).Move B1_Lb_Day(gLong - 1).Left + Day_Width
        
        Load B1_NSU(gLong)
        B1_NSU(gLong).Visible = True
        B1_NSU(gLong).Move B1_NSU(gLong - 1).Left + Day_Width
        
        Load B1_K_LINEDay(gLong)
        B1_K_LINEDay(gLong).Visible = True
        B1_K_LINEDay(gLong).X1 = B1_K_LINEDay(gLong - 1).X1 + Day_Width
        B1_K_LINEDay(gLong).X2 = B1_K_LINEDay(gLong - 1).X2 + Day_Width
    Next gLong
    B1_K_LINEDay(30).BorderWidth = 4
    
    ' ì‡é¶èÓïÒ
    Nj_Width = B1_K_LINENJU(0).X1 - B1_K_LINENJU(0).X2
    For gInt = 1 To 5
        Load B1_K_LINENJU(gInt)
        Load B1_K_LINENJM(gInt)
        Load B1_K_LINENJS(gInt)
        Load B1_K_LINENJT(gInt)
        Load B1_Lb_NJ(gInt)
        Load B1_Lb_NJm(gInt)
        Load B1_NJSU(gInt)
        
        B1_K_LINENJU(gInt).Visible = True
        B1_K_LINENJM(gInt).Visible = True
        B1_K_LINENJS(gInt).Visible = True
        B1_K_LINENJT(gInt).Visible = True
        B1_Lb_NJ(gInt).Visible = True
        B1_Lb_NJm(gInt).Visible = True
        B1_NJSU(gInt).Visible = True
                
        B1_K_LINENJU(gInt).X1 = B1_K_LINENJU(gInt - 1).X1 + Nj_Width
        B1_K_LINENJU(gInt).X2 = B1_K_LINENJU(gInt - 1).X2 + Nj_Width
        
        B1_K_LINENJM(gInt).X1 = B1_K_LINENJM(gInt - 1).X1 + Nj_Width
        B1_K_LINENJM(gInt).X2 = B1_K_LINENJM(gInt - 1).X2 + Nj_Width
                
        B1_K_LINENJS(gInt).X1 = B1_K_LINENJS(gInt - 1).X1 + Nj_Width
        B1_K_LINENJS(gInt).X2 = B1_K_LINENJS(gInt - 1).X2 + Nj_Width
        
        B1_K_LINENJT(gInt).X1 = B1_K_LINENJT(gInt - 1).X1 + Nj_Width
        B1_K_LINENJT(gInt).X2 = B1_K_LINENJT(gInt - 1).X2 + Nj_Width
    
        B1_Lb_NJ(gInt).Left = B1_Lb_NJ(gInt - 1).Left + Nj_Width
        B1_Lb_NJm(gInt).Left = B1_Lb_NJm(gInt - 1).Left + Nj_Width
        
        B1_NJSU(gInt).Left = B1_NJSU(gInt - 1).Left + Nj_Width
        
        If gInt = 2 Or gInt = 5 Then
           B1_K_LINENJT(gInt).BorderWidth = 8
        Else
           B1_K_LINENJT(gInt).BorderWidth = 4
        End If
        If gInt > 1 Then
           B1_Lb_NJm(gInt).Font.Size = 9
           B1_Lb_NJm(gInt).Top = B1_Lb_NJm(gInt).Top - 30
           If gInt = 2 Then
              B1_Lb_NJm(gInt).Left = B1_Lb_NJm(gInt).Left + 60
           End If
        End If
        If gInt > 2 Then
           B1_K_LINENJU(gInt).BorderWidth = 8
           B1_K_LINENJS(gInt).BorderWidth = 8
           B1_Lb_NJm(gInt).Tag = "édä|ì‡é¶"
        End If

    Next gInt
    B1_Lb_NJm(0).Tag = B1_Lb_NJm(0)
    B1_Lb_NJm(1).Tag = "åvâÊïœçXêî"
    B1_Lb_NJm(2).Tag = "à¯éÊì‡é¶"
    
    gStr = NTT.ì‡é¶îNåé
    B1_Lb_NJ(0).Tag = Mid(gStr, 3, 2) & "îN " & CisFun.RSetFld(Mid(gStr, 5, 2), 2, "##") & "åé"
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add gStr & "01", M, 1, LastDate
        gStr = .Date_Result
        B1_Lb_NJ(1).Tag = Mid(gStr, 3, 2) & "îN " & CisFun.RSetFld(Mid(gStr, 5, 2), 2, "##") & "åé"
        .Date_RsFormat = fYM
        .Date_Add gStr & "01", M, 1, LastDate
        gStr = .Date_Result
        B1_Lb_NJ(2).Tag = Mid(gStr, 3, 2) & "îN " & CisFun.RSetFld(Mid(gStr, 5, 2), 2, "##") & "åé"
    End With
    B1_Lb_NJ(3).Tag = B1_Lb_NJ(0).Tag
    B1_Lb_NJ(4).Tag = B1_Lb_NJ(1).Tag
    B1_Lb_NJ(5).Tag = B1_Lb_NJ(2).Tag
    ' ìKópä˙ä‘
    Teki_Width = B1_K_LUTe(0).X1 - B1_K_LUTe(0).X2
    B1_K_LUTe(0).X1 = B1_K_LINENJT(5).X1 + Teki_Width
    B1_K_LUTe(0).X2 = B1_K_LINENJT(5).X2
    B1_K_LMTe2(0).X1 = B1_K_LINENJT(5).X1 + Teki_Width
    B1_K_LMTe2(0).X2 = B1_K_LINENJT(5).X2
    B1_K_LMTe(0).X1 = B1_K_LINENJT(5).X1 + Teki_Width
    B1_K_LMTe(0).X2 = B1_K_LINENJT(5).X2
    B1_K_LSTe(0).X1 = B1_K_LINENJT(5).X1 + Teki_Width
    B1_K_LSTe(0).X2 = B1_K_LINENJT(5).X2
    B1_Lb_Teki(0).Left = B1_Lb_Teki(0).Left + Nj_Width * 5
    B1_Lb_TekiY(0).Left = B1_Lb_TekiY(0).Left + Nj_Width * 5
    B1_Lb_TekiJ(0).Left = B1_Lb_TekiJ(0).Left + Nj_Width * 5
    
    B1_K_LTe1(0).X1 = B1_K_LINENJT(5).X1 + Teki_Width - 260
    B1_K_LTe1(0).X2 = B1_K_LINENJT(5).X1 + Teki_Width - 260
    B1_K_LTe2(0).X1 = B1_K_LINENJT(5).X1 + Teki_Width
    B1_K_LTe2(0).X2 = B1_K_LINENJT(5).X1 + Teki_Width
    
    Load B1_K_LUTe(1)
    Load B1_K_LMTe2(1)
    Load B1_K_LMTe(1)
    Load B1_K_LSTe(1)
    Load B1_Lb_Teki(1)
    Load B1_Lb_TekiY(1)
    Load B1_Lb_TekiJ(1)
    Load B1_K_LTe1(1)
    Load B1_K_LTe2(1)
    Load B1_TY(1)
    Load B1_TYJ(1)
    
    B1_K_LUTe(1).Visible = True
    B1_K_LMTe2(1).Visible = True
    B1_K_LMTe(1).Visible = True
    B1_K_LSTe(1).Visible = True
    B1_Lb_Teki(1).Visible = True
    B1_Lb_TekiY(1).Visible = True
    B1_Lb_TekiJ(1).Visible = True
    B1_K_LTe1(1).Visible = True
    B1_K_LTe2(1).Visible = True
    B1_TY(1).Visible = True
    B1_TYJ(1).Visible = True
    
    B1_Lb_Teki(0).Tag = B1_Lb_Teki(0)
    B1_Lb_Teki(1).Tag = "ìKópèIóπ"
    B1_Lb_TekiY(0).Tag = B1_Lb_TekiY(0)
    B1_Lb_TekiY(1).Tag = B1_Lb_TekiY(0)
    B1_Lb_TekiJ(0).Tag = B1_Lb_TekiJ(0)
    B1_Lb_TekiJ(1).Tag = B1_Lb_TekiJ(1)

    B1_K_LUTe(1).X1 = B1_K_LUTe(0).X1 + Teki_Width
    B1_K_LUTe(1).X2 = B1_K_LUTe(0).X2 + Teki_Width
    B1_K_LMTe2(1).X1 = B1_K_LMTe2(0).X1 + Teki_Width
    B1_K_LMTe2(1).X2 = B1_K_LMTe2(0).X2 + Teki_Width
    B1_K_LMTe(1).X1 = B1_K_LMTe(0).X1 + Teki_Width
    B1_K_LMTe(1).X2 = B1_K_LMTe(0).X2 + Teki_Width
    B1_K_LSTe(1).X1 = B1_K_LSTe(0).X1 + Teki_Width
    B1_K_LSTe(1).X2 = B1_K_LSTe(0).X2 + Teki_Width
    B1_Lb_Teki(1).Left = B1_Lb_Teki(0).Left + Teki_Width
    B1_Lb_TekiY(1).Left = B1_Lb_TekiY(0).Left + Teki_Width
    B1_Lb_TekiJ(1).Left = B1_Lb_TekiJ(0).Left + Teki_Width
    
    B1_TY(0).Left = B1_TY(0).Left + Nj_Width * 5
    B1_TYJ(0).Left = B1_TYJ(0).Left + Nj_Width * 5
    B1_TY(1).Left = B1_TY(0).Left + Teki_Width
    B1_TYJ(1).Left = B1_TYJ(0).Left + Teki_Width
    
    B1_K_LTe1(1).X1 = B1_K_LTe1(0).X1 + Teki_Width
    B1_K_LTe1(1).X2 = B1_K_LTe1(0).X2 + Teki_Width
    B1_K_LTe2(1).X1 = B1_K_LTe2(0).X1 + Teki_Width
    B1_K_LTe2(1).X2 = B1_K_LTe2(0).X2 + Teki_Width
    
    PrtTime = Format(Now(), "yyyy/mm/dd hh:nn")             'î≠çsì˙ÅEéûä‘
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    With CisPrt
        .iPrtTitle = "ì‡é¶ï\"
        .iLineMax = 60
'        .iLinePitch = 8
        .iPaperSize = p_A3
        .iOrientation = po_Landscape
        If Not .PrintIni Then
           With CisFun
                .MB_Lines = 5
                .MB_MSG(1) = "     Å@ÉvÉäÉìÉ^èâä˙ê›íËÉGÉâÅ[        "
                .MB_Title = "àÛç¸¥◊∞"
                .MB_Button = Error
                .MBOX
           End With
           Exit Sub
        .cLine = 0
        .cPage = 0
        End If
    End With
    
    Call PrintProc
    
End Sub
'*---------------------*
'*Å@  àÛ  ç¸  èà  óù
'*---------------------*
Private Function PrintProc()
    
    Brk_Kouku = ""
    Brk_KTorcd = ""
    Brk_Naiji = 0

    Do Until Not NTT_RDSTS
        With NTT
            
            Call HeadWrite(0)
            
            Call CisPrt.ClrLstFld("B")
            
            B1_Lb_Hinbn = B1_Lb_Hinbn.Tag
            B1_Lb_Torcd = B1_Lb_Torcd.Tag
            B1_Lb_Ukeir = B1_Lb_Ukeir.Tag
            B1_Lb_Seibi = B1_Lb_Seibi.Tag
            B1_Lb_Tanto = B1_Lb_Tanto.Tag
            B1_Lb_Nouku = B1_Lb_Nouku.Tag
            B1_Lb_Seban = B1_Lb_Seban.Tag
            B1_Lb_Cycle = B1_Lb_Cycle.Tag
            B1_Lb_Cycle1 = B1_Lb_Cycle1.Tag
            B1_Lb_Cycle2 = B1_Lb_Cycle2.Tag
            B1_Lb_Cycle3 = B1_Lb_Cycle3.Tag
            B1_Lb_Syuyo = B1_Lb_Syuyo.Tag
            B1_Lb_Kanri = B1_Lb_Kanri.Tag
            B1_Lb_Nai1 = B1_Lb_Nai1.Tag
            B1_Lb_Nai2 = B1_Lb_Nai2.Tag
            B1_Lb_Nai3 = B1_Lb_Nai3.Tag
            B1_Lb_Nai4 = B1_Lb_Nai4.Tag
            B1_Lb_Comment = B1_Lb_Comment.Tag
            B1_Lb_HMai = B1_Lb_HMai.Tag
            B1_Lb_KMai = B1_Lb_KMai.Tag
            B1_Lb_KMaiZ = B1_Lb_KMaiZ.Tag
            B1_Lb_Day1 = B1_Lb_Day1.Tag
            B1_Lb_Day2 = B1_Lb_Day2.Tag
            B1_Lb_Day3 = B1_Lb_Day3.Tag
            B1_Lb_Day4 = B1_Lb_Day4.Tag
            B1_Lb_Day5 = B1_Lb_Day5.Tag
            B1_Lb_Day6 = B1_Lb_Day6.Tag
            For gInt = 0 To 5
                B1_Lb_NJ(gInt) = B1_Lb_NJ(gInt).Tag
                B1_Lb_NJm(gInt) = B1_Lb_NJm(gInt).Tag
                
            Next gInt
            B1_Lb_Teki(0) = B1_Lb_Teki(0).Tag
            B1_Lb_TekiY(0) = B1_Lb_TekiY(0).Tag
            B1_Lb_TekiJ(0) = B1_Lb_TekiJ(0).Tag
            B1_Lb_Teki(1) = B1_Lb_Teki(1).Tag
            B1_Lb_TekiY(1) = B1_Lb_TekiY(1).Tag
            B1_Lb_TekiJ(1) = B1_Lb_TekiJ(1).Tag
            
            
            Brk_Kouku = NTT.íuä∑çHãÊ
            Brk_KTorcd = NTT.å_ñÒêÊ
            Brk_Naiji = NTT.ì‡é¶å`ë‘
            
            B1_Hinbn = EditHinbn(.ïiî‘, .ïiî‘ï“èW)
            B1_NTorcd = RTrim(.î[ì¸êÊ)
            B1_Ukeir = RTrim(.éÛì¸)
            B1_Seibi = RTrim(.êÆîı)
            B1_Tanto = RTrim(.íSìñ)
            B1_Nouku = RTrim(.î[ì¸ãÊï™)
            B1_Seban = RTrim(.îwî‘çÜ)
'*********** 2006/01/17 Update Start *******************************************
'            If Trim(.ÉTÉCÉNÉã) <> "" Then
'                B1_Cycle1 = CisFun.RSetFld(Mid(.ÉTÉCÉNÉã, 1, 1), 1, "#")
'                B1_Cycle2 = CisFun.RSetFld(Mid(.ÉTÉCÉNÉã, 2, 2), 2, "#")
'                B1_Cycle3 = CisFun.RSetFld(Mid(.ÉTÉCÉNÉã, 4, 2), 2, "#")
'            Else
'                B1_Cycle1 = ""
'                B1_Cycle2 = ""
'                B1_Cycle3 = ""
'            End If
            If Trim(.ÉTÉCÉNÉã) <> "" Then
               If Len(Trim(.ÉTÉCÉNÉã)) < 6 Then
                  B1_Cycle1 = CisFun.RSetFld(Mid(.ÉTÉCÉNÉã, 1, 1), 1, "#")
                  B1_Cycle2 = CisFun.RSetFld(Mid(.ÉTÉCÉNÉã, 2, 2), 2, "#")
                  B1_Cycle3 = CisFun.RSetFld(Mid(.ÉTÉCÉNÉã, 4, 2), 2, "#")
               Else
                  B1_Cycle1 = CisFun.RSetFld(CisFun.Val2(Mid(.ÉTÉCÉNÉã, 1, 2)), 1, "#")
                  B1_Cycle2 = CisFun.RSetFld(Mid(.ÉTÉCÉNÉã, 3, 2), 2, "#")
                  If CisFun.Val2(Mid(.ÉTÉCÉNÉã, 8, 3)) <> 0 Then        'ç≈å„ÇRåÖÇÕè¨êîì_à»â∫
                     w_Int = CisFun.Val2(Mid(.ÉTÉCÉNÉã, 5, 3)) + 1      'è¨êîì_êîílÇ™ë∂ç›Ç∑ÇÈèÍçáÇÕêÆêîïîÇ…Å{ÇP
                  Else
                     w_Int = CisFun.Val2(Mid(.ÉTÉCÉNÉã, 5, 3))          'ë∂ç›ÇµÇ»Ç¢èÍçáÇÕÇªÇÃÇ‹Ç‹ÇÃílÇégóp
                  End If
                  B1_Cycle3 = CisFun.RSetFld(Format(w_Int, "###"), 3, "#")
               End If
            Else
                B1_Cycle1 = ""
                B1_Cycle2 = ""
                B1_Cycle3 = ""
            End If
'*********** 2006/01/17 Update End   *******************************************
            
            B1_Syuyo = CisFun.RSetFld(.é˚óeêî, 5, "##")
            B1_Kanri = .ä«óùNO
                        
            B1_NJSU(0) = CisFun.RSetFld(.ìñèâì‡é¶êî, 10, "#,###")
            If .í≤êÆêîïÑçÜ = "-" Then .í≤êÆêî = .í≤êÆêî * -1
            B1_NJSU(1) = CisFun.RSetFld(.í≤êÆêî, 10, "#,###")
            If .ïœçXêîïÑçÜ = "-" Then .ïœçXêî = .ïœçXêî * -1
            B1_NJSU(2) = CisFun.RSetFld(.ïœçXêî, 10, "#,###")
            B1_NJSU(3) = CisFun.RSetFld(.ì‡é¶êî, 10, "#,###")
            B1_NJSU(4) = CisFun.RSetFld(.óÇåéì‡é¶êî, 10, "#,###")
            B1_NJSU(5) = CisFun.RSetFld(.óÇÅXåéì‡é¶êî, 10, "#,###")
            
            If Trim(.ìKópäJénîNåéì˙) <> "" Then
                B1_TY(0) = CisFun.RSetFld(Mid(.ìKópäJénîNåéì˙, 5, 2), 2, "00")
                gLong = CisFun.Val2(Mid(.ìKópäJénîNåéì˙, 7, 2))
                If gLong > 0 And gLong < 11 Then B1_TYJ(0) = "1"
                If gLong > 10 And gLong < 21 Then B1_TYJ(0) = "2"
                If gLong > 20 Then B1_TYJ(0) = "3"
            Else
                B1_TY(0) = ""
                B1_TYJ(0) = ""
            End If
            If Trim(.ìKópèIóπîNåéì˙) <> "" Then
                B1_TY(1) = CisFun.RSetFld(Mid(.ìKópèIóπîNåéì˙, 5, 2), 2, "00")
                gLong = CisFun.Val2(Mid(.ìKópèIóπîNåéì˙, 7, 2))
                If gLong > 0 And gLong < 11 Then B1_TYJ(1) = "1"
                If gLong > 10 And gLong < 21 Then B1_TYJ(1) = "2"
                If gLong > 20 Then B1_TYJ(1) = "3"
            Else
                B1_TY(1) = ""
                B1_TYJ(1) = ""
            End If
            
            B1_Comment = RTrim(.ÉRÉÅÉìÉg)
            B1_CMai = CisFun.RSetFld(.î[ïièëî≠çsñáêî, 3, "###")
            B1_KMai = CisFun.RSetFld(.Ç©ÇÒÇŒÇÒâÒì]ñáêî, 3, "###")
            If .Ç©ÇÒÇŒÇÒëùå∏ïÑçÜ = "-" Then
                .Ç©ÇÒÇŒÇÒëùå∏ñáêî = .Ç©ÇÒÇŒÇÒëùå∏ñáêî * -1
            End If
            B1_KMaiSa = CisFun.RSetFld(.Ç©ÇÒÇŒÇÒëùå∏ñáêî, 4, "####")
            B1_Lb_Day(0) = B1_Lb_Day(0).Tag
            For gLong = 1 To 31
                B1_NSU(gLong - 1) = CisFun.RSetFld(.ì‡é¶êîó (gLong), 5, "#####")
                B1_Lb_Day(gLong - 1) = B1_Lb_Day(gLong - 1).Tag
            Next gLong
            
        End With

    '*---------------------------------*
    '*  BodyçsÅ@àÛç¸
    '*---------------------------------*
        Call CisPrt.PrintFld("B1")
                
        With CisPrt
            .cLine = .cLine + 6
            .cDataCnt = .cDataCnt + 1
            If .cDataCnt < 200 Or .cDataCnt Mod 50 = 0 Then
               .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(mPrtCnt, "#,###")
            End If
        End With
        
        Call NTTReadNext
    Loop
    Call NTTClose
    
    ' àÛç¸èIóπ
    Call CisPrt.PrintEndDoc
    
End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite(KB As Byte)
    Dim NewPageFlg As Boolean
'
    HeadWrite = False
    With CisPrt
        ' ÇPï≈ì‡àÛç¸çsêî¡™Ø∏
        NewPageFlg = False
        If Brk_Kouku <> NTT.íuä∑çHãÊ Or _
           Brk_KTorcd <> NTT.å_ñÒêÊ Or _
           Brk_Naiji <> NTT.ì‡é¶å`ë‘ Then
           NewPageFlg = True
        End If
        If .iLineMax >= .cLine + 1 And Not NewPageFlg Then Exit Function
        
        
        .cLine = 0
        .cPage = .cPage + 1
        
        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé
    
        H1_Kouku = NTT.íuä∑çHãÊ
        H1_KTorcd = "(" & NTT.å_ñÒêÊ & ")"
        H1_KTORNM = NTT.ê≥éÆñºèÃ
    
        If NTT.ì‡é¶å`ë‘ = 0 Then
           H1_Naiji = "Åy" & CNJ0110_Naiji & "Åz"
        Else
           H1_Naiji = "ÅyåvâÊïœçXÅz"
        End If
        H1_PAGE = CisFun.RSetFld(.cPage, 5, "###")
        H1_HYMD = PrtTime
        
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With

    HeadWrite = True
End Function
