VERSION 5.00
Begin VB.Form COY0311L 
   Caption         =   "îÑè„é¿ê—àÍóóï\"
   ClientHeight    =   3330
   ClientLeft      =   -1080
   ClientTop       =   1620
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   15420
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "îwî‘çÜ"
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
      Index           =   8
      Left            =   2700
      TabIndex        =   54
      Top             =   1545
      Width           =   720
   End
   Begin VB.Label B1_Seban 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXX"
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
      Left            =   2685
      TabIndex        =   53
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label B1_Kari 
      AutoSize        =   -1  'True
      Caption         =   "X"
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
      Index           =   2
      Left            =   13515
      TabIndex        =   52
      Top             =   1920
      Width           =   90
   End
   Begin VB.Label B1_Kari 
      AutoSize        =   -1  'True
      Caption         =   "X"
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
      Index           =   1
      Left            =   9105
      TabIndex        =   51
      Top             =   1920
      Width           =   90
   End
   Begin VB.Label B1_Kari 
      AutoSize        =   -1  'True
      Caption         =   "X"
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
      Left            =   4665
      TabIndex        =   50
      Top             =   1935
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "íPÅ@âø"
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
      Index           =   29
      Left            =   12855
      TabIndex        =   49
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "êî ó "
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
      Index           =   23
      Left            =   13980
      TabIndex        =   46
      Top             =   1560
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ã‡ äz"
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
      Index           =   21
      Left            =   15015
      TabIndex        =   45
      Top             =   1560
      Width           =   570
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   12
      X1              =   12405
      X2              =   13405
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Label H1_NJYm 
      AutoSize        =   -1  'True
      Caption         =   "MM"
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
      Index           =   2
      Left            =   13800
      TabIndex        =   44
      Top             =   1215
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "åé"
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
      Index           =   19
      Left            =   14190
      TabIndex        =   43
      Top             =   1260
      Width           =   195
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   11
      X1              =   12405
      X2              =   12405
      Y1              =   1350
      Y2              =   1550
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   10
      X1              =   14670
      X2              =   15670
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   9
      X1              =   15690
      X2              =   15690
      Y1              =   1350
      Y2              =   1550
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "íPÅ@âø"
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
      Index           =   17
      Left            =   8430
      TabIndex        =   42
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "êî ó "
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
      Index           =   10
      Left            =   9555
      TabIndex        =   39
      Top             =   1560
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ëOåéî‰"
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
      Index           =   6
      Left            =   11580
      TabIndex        =   38
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ã‡ äz"
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
      Index           =   5
      Left            =   10590
      TabIndex        =   37
      Top             =   1560
      Width           =   570
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   8
      X1              =   7980
      X2              =   9580
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Label H1_NJYm 
      AutoSize        =   -1  'True
      Caption         =   "MM"
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
      Index           =   1
      Left            =   9900
      TabIndex        =   36
      Top             =   1215
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "åé"
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
      Index           =   4
      Left            =   10290
      TabIndex        =   35
      Top             =   1260
      Width           =   195
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   7
      X1              =   7980
      X2              =   7980
      Y1              =   1350
      Y2              =   1550
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   6
      X1              =   10725
      X2              =   12325
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   5
      X1              =   12330
      X2              =   12330
      Y1              =   1350
      Y2              =   1550
   End
   Begin VB.Label B1_Tanka 
      AutoSize        =   -1  'True
      Caption         =   "999,999.99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   12435
      TabIndex        =   34
      Top             =   1905
      Width           =   1050
   End
   Begin VB.Label B1_Tanka 
      AutoSize        =   -1  'True
      Caption         =   "999,999.99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   8040
      TabIndex        =   33
      Top             =   1905
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "íPÅ@âø"
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
      Index           =   26
      Left            =   4020
      TabIndex        =   32
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label B1_Tanka 
      AutoSize        =   -1  'True
      Caption         =   "999,999.99"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   3600
      TabIndex        =   31
      Top             =   1920
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ïiÅ@Å@î‘"
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
      Index           =   25
      Left            =   675
      TabIndex        =   30
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label B1_Hinbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX-XXXXX-XX-XXXXXXXXXX"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   330
      TabIndex        =   29
      Top             =   1920
      Width           =   2250
   End
   Begin VB.Label H1_Kouku 
      AutoSize        =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3525
      TabIndex        =   28
      Top             =   615
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "çHãÊ:"
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
      Index           =   24
      Left            =   2850
      TabIndex        =   27
      Top             =   630
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÅyíPà ÅFêÁâ~Åz"
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
      Index           =   20
      Left            =   10950
      TabIndex        =   26
      Top             =   630
      Width           =   1365
   End
   Begin VB.Label B1_Suryo 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   13635
      TabIndex        =   23
      Top             =   1890
      Width           =   945
   End
   Begin VB.Label B1_Kingk 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   14640
      TabIndex        =   22
      Top             =   1890
      Width           =   945
   End
   Begin VB.Label B1_Zenhi 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   11355
      TabIndex        =   21
      Top             =   1905
      Width           =   945
   End
   Begin VB.Label B1_Suryo 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   9255
      TabIndex        =   20
      Top             =   1905
      Width           =   945
   End
   Begin VB.Label B1_Kingk 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   10260
      TabIndex        =   19
      Top             =   1905
      Width           =   945
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   4
      X1              =   7920
      X2              =   7920
      Y1              =   1350
      Y2              =   1550
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   2
      X1              =   6315
      X2              =   7915
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   3
      X1              =   3570
      X2              =   3570
      Y1              =   1350
      Y2              =   1550
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "åé"
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
      Index           =   2
      Left            =   5880
      TabIndex        =   18
      Top             =   1260
      Width           =   195
   End
   Begin VB.Label H1_NJYm 
      AutoSize        =   -1  'True
      Caption         =   "MM"
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
      Index           =   0
      Left            =   5490
      TabIndex        =   17
      Top             =   1215
      Width           =   300
   End
   Begin VB.Label B1_Kingk 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   5865
      TabIndex        =   16
      Top             =   1920
      Width           =   945
   End
   Begin VB.Line B1_LINE 
      Visible         =   0   'False
      X1              =   900
      X2              =   15800
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   4
      Index           =   1
      X1              =   375
      X2              =   15800
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   0
      X1              =   3570
      X2              =   5170
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Label H1_MM 
      AutoSize        =   -1  'True
      Caption         =   "MM"
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
      Left            =   1515
      TabIndex        =   15
      Top             =   600
      Width           =   300
   End
   Begin VB.Label H1_YY 
      AutoSize        =   -1  'True
      Caption         =   "YYYY"
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
      Left            =   435
      TabIndex        =   14
      Top             =   600
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "îN"
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
      Index           =   13
      Left            =   1110
      TabIndex        =   13
      Top             =   600
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "åéìx"
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
      Index           =   12
      Left            =   1845
      TabIndex        =   12
      Top             =   585
      Width           =   570
   End
   Begin VB.Label B1_Suryo 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4815
      TabIndex        =   11
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label B1_Zenhi 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   6960
      TabIndex        =   10
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label H1_TORCD 
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
      Left            =   1155
      TabIndex        =   9
      Top             =   990
      Width           =   735
   End
   Begin VB.Label H1_TORNM 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XX"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1935
      TabIndex        =   8
      Top             =   960
      Width           =   2640
   End
   Begin VB.Label H1_PAGE 
      AutoSize        =   -1  'True
      Caption         =   "9,999"
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
      Left            =   14715
      TabIndex        =   7
      Top             =   600
      Width           =   600
   End
   Begin VB.Label H1_YMD 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD HH:MM"
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
      Left            =   12795
      TabIndex        =   6
      Top             =   600
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "å_ñÒêÊ"
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
      Index           =   11
      Left            =   435
      TabIndex        =   5
      Top             =   975
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ã‡ äz"
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
      Index           =   9
      Left            =   6180
      TabIndex        =   4
      Top             =   1560
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ëOåéî‰"
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
      Index           =   7
      Left            =   7170
      TabIndex        =   3
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "êî ó "
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
      Index           =   3
      Left            =   5145
      TabIndex        =   2
      Top             =   1560
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ï≈"
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
      Index           =   1
      Left            =   15345
      TabIndex        =   1
      Top             =   600
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Å¶ îÑè„ó\ëzï\ÅyïîïiÅz Å¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   4890
      TabIndex        =   0
      Top             =   600
      Width           =   4320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ås"
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
      Index           =   15
      Left            =   5055
      TabIndex        =   24
      Top             =   1245
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Åt"
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
      Index           =   18
      Left            =   6225
      TabIndex        =   25
      Top             =   1245
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Åt"
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
      Index           =   28
      Left            =   14535
      TabIndex        =   48
      Top             =   1245
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ås"
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
      Index           =   27
      Left            =   13365
      TabIndex        =   47
      Top             =   1245
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ås"
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
      Index           =   14
      Left            =   9480
      TabIndex        =   40
      Top             =   1245
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Åt"
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
      Index           =   16
      Left            =   10635
      TabIndex        =   41
      Top             =   1245
      Width           =   225
   End
End
Attribute VB_Name = "COY0311L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  îÑè„ó\ëzï\ÅyïîïiÅEñæç◊Åz
'**       ÉtÉHÅ[ÉÄID    :  COY0311L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2004/10/05
'**       ïœ  çX  ì˙    :  2008/04/11 é‘å^5åÖ
'**
'**********************************************************'
    Dim OutStr          As String   ' CSVèoóÕì‡óe
    Dim KeyKouku        As String
    Dim KeyTorcd        As String
    Dim KeySyakei       As String
    Dim KeyHinbn        As String
    
    Dim wASuryo(3)      As Currency     'é‘å^åv
    Dim wAKinGk(3)      As Currency
    Dim wAZenHi(3)      As Currency
    Dim wTSuryo(3)      As Currency     'å_ñÒêÊåv
    Dim wTKinGk(3)      As Currency
    Dim wTZenHi(3)      As Currency
    Dim wKSuryo(3)      As Currency     'çHãÊåv
    Dim wKKinGk(3)      As Currency
    Dim wKZenHi(3)      As Currency
    Dim wSSuryo(3)      As Currency     'ëççáåv
    Dim wSKinGk(3)      As Currency
    Dim wSZenHi(3)      As Currency
    
    Dim wInCnt          As Long
    Dim wSyaCnt         As Long
    Dim mPrtCnt         As Long
    
'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub Form_Load()
    App.Title = "îÑè„ó\ëzï\ÅyïîïiÅz"
    
    If Not DataRead Then Exit Sub
    
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    mPrtCnt = CisDB.RecordCount

    With CisPrt
        .iPrtTitle = "îÑè„ó\ëzï\ÅyïîïiÅz"
        .iLineMax = 38
        .iLinePitch = 6
        .iPaperSize = p_A4
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
'*-------------------------*
'*      àÛç¸ÉfÅ[É^ì«çû
'*-------------------------*
Private Function DataRead()
    DataRead = False
    
    Call SQLEdit0311(UY_HinKeta)
    If Not UYTRead(gSL_Select, 0) Then
        Call UYTClose
        Exit Function
    End If
    
    DataRead = True
End Function
'*---------------------*
'*Å@  àÛ  ç¸  èà  óù
'*---------------------*
Private Function PrintProc()
    
    Call CisPrt.ClrLstFld("B*")
    CisPrt.cLine = 99
    wInCnt = 0: wSyaCnt = 0
    KeyKouku = "": KeyTorcd = "": KeySyakei = "": KeyHinbn = ""
    Erase wASuryo: Erase wAKinGk: Erase wAZenHi
    Erase wTSuryo: Erase wTKinGk: Erase wTZenHi
    Erase wKSuryo: Erase wKKinGk: Erase wKZenHi
    Erase wSSuryo: Erase wSKinGk: Erase wSZenHi
    
    Do Until Not UYT_RDSTS
        With UYT
            
            If RTrim(KeyKouku) <> RTrim(.çHãÊ) Or RTrim(KeyTorcd) <> RTrim(.å_ñÒêÊ) Then
                If wInCnt <> 0 Then
                    If RTrim(KeyKouku) <> RTrim(.çHãÊ) Then
                        Call TotalWrite(2)
                    Else
                        Call TotalWrite(1)
                    End If
                    CisPrt.cLine = 99
                    Call HeadWrite
                End If
            Else
'                If RTrim(KeySyakei) <> RTrim(.é‘å^) Then
'                    Call TotalWrite(0)
'                End If
            End If
            Call HeadWrite
'            If RTrim(KeySyakei) <> RTrim(.é‘å^) Then
'                B1_Syakei = RTrim(.é‘å^)
'                KeyHinbn = ""
'            End If
            If RTrim(KeyHinbn) <> RTrim(.ïiî‘) Then
                B1_Hinbn = EditHinbn(.ïiî‘, .ïiî‘ï“èW)
            End If
            
            B1_Seban = .é‘å^
            
            B1_Tanka(0) = Format(Format(.íPâø1, "#,##0.00;-#,##0.00;#"), String(10, "@"))
            If RTrim(.âºãÊï™1) = "1" Then B1_Kari(0) = "*"
            B1_Suryo(0) = Format(Format(.êîó 1, "#,##0;-#,##0;#"), String(9, "@"))
            gLong = .ã‡äz1 / 1000
            B1_Kingk(0) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
            gLong = .ëOåéî‰1 / 1000
            B1_Zenhi(0) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
            
            B1_Tanka(1) = Format(Format(.íPâø2, "#,##0.00;-#,##0.00;#"), String(10, "@"))
            If RTrim(.âºãÊï™2) = "1" Then B1_Kari(1) = "*"
            B1_Suryo(1) = Format(Format(.êîó 2, "#,##0;-#,##0;#"), String(9, "@"))
            gLong = .ã‡äz2 / 1000
            B1_Kingk(1) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
            gLong = .ëOåéî‰2 / 1000
            B1_Zenhi(1) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
            
            If RTrim(.âºãÊï™3) = "1" Then B1_Kari(2) = "*"
            B1_Tanka(2) = Format(Format(.íPâø3, "#,##0.00;-#,##0.00;#"), String(10, "@"))
            B1_Suryo(2) = Format(Format(.êîó 3, "#,##0;-#,##0;#"), String(9, "@"))
            gLong = .ã‡äz3 / 1000
            B1_Kingk(2) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
            
            Call CisPrt.PrintFld("B1")

            Call CisPrt.ClrLstFld("B*")
            
            With CisPrt
                .cLine = .cLine + 1
                .cDataCnt = .cDataCnt + 1
                .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(mPrtCnt, "#,###")
            End With
            B1_LINE.Visible = False
            
            wASuryo(1) = wASuryo(1) + .êîó 1
            wAKinGk(1) = wAKinGk(1) + .ã‡äz1
            wAZenHi(1) = wAZenHi(1) + .ëOåéî‰1
            wASuryo(2) = wASuryo(2) + .êîó 2
            wAKinGk(2) = wAKinGk(2) + .ã‡äz2
            wAZenHi(2) = wAZenHi(2) + .ëOåéî‰2
            wASuryo(3) = wASuryo(3) + .êîó 3
            wAKinGk(3) = wAKinGk(3) + .ã‡äz3
            
            wInCnt = wInCnt + 1
            wSyaCnt = wSyaCnt + 1
            KeyKouku = .çHãÊ
            KeyTorcd = .å_ñÒêÊ
'            KeySyakei = .é‘å^
            KeyHinbn = .ïiî‘
        End With

        Call UYTReadNext
    Loop
    Call UYTClose
    
    Call HeadWrite
    Call TotalWrite(3)
    
    ' àÛç¸èIóπ
    Call CisPrt.PrintEndDoc

End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite()
'
    HeadWrite = False
    With CisPrt
        ' ÇPï≈ì‡àÛç¸çsêî¡™Ø∏
        If .iLineMax >= .cLine + 1 Then Exit Function
        .cLine = 0
        .cPage = .cPage + 1
        
        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé
    
        With UYT
            H1_PAGE = CisFun.RSetFld(CisPrt.cPage, 5, "#,###")
            H1_YMD = Format(Now(), "YYYY/MM/DD HH:NN")
            H1_Yy = Mid(.ì‡é¶îNåé, 1, 4)
            H1_Mm = Mid(.ì‡é¶îNåé, 5, 2)
            H1_Kouku = RTrim(.çHãÊ)
            H1_TORCD = RTrim(.å_ñÒêÊ)
            H1_TORNM = RTrim(.å_ñÒêÊñº)
            H1_NJYm(0) = Mid(.ì‡é¶îNåé, 5, 2)
            H1_NJYm(1) = Format(DateAdd("m", 1, CDate(Left(.ì‡é¶îNåé, 4) & "/" & Mid(.ì‡é¶îNåé, 5, 2) & "/01")), "mm")
            H1_NJYm(2) = Format(DateAdd("m", 2, CDate(Left(.ì‡é¶îNåé, 4) & "/" & Mid(.ì‡é¶îNåé, 5, 2) & "/01")), "mm")
        End With
    
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With
    
    KeySyakei = "": KeyHinbn = ""
    B1_LINE.Visible = False
    HeadWrite = True
End Function
'*----------------*
'*   çáåvàÛç¸     *
'*----------------*
Private Sub TotalWrite(Optional Kb As Integer = 0)
    With CisPrt
        If wSyaCnt > 1 Then
'            B1_Hinbn = "( é‘ å^ åv )"
'            B1_Suryo(0) = Format(Format(wASuryo(1), "#,##0;-#,##0;#"), String(9, "@"))
'            gLong = wAKinGk(1) / 1000
'            B1_Kingk(0) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
'            gLong = wAZenHi(1) / 1000
'            B1_Zenhi(0) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
'            B1_Suryo(1) = Format(Format(wASuryo(2), "#,##0;-#,##0;#"), String(9, "@"))
'            gLong = wAKinGk(2) / 1000
'            B1_Kingk(1) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
'            gLong = wAZenHi(2) / 1000
'            B1_Zenhi(1) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
'            B1_Suryo(2) = Format(Format(wASuryo(3), "#,##0;-#,##0;#"), String(9, "@"))
'            gLong = wAKinGk(3) / 1000
'            B1_Kingk(2) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
'            B1_LINE.Visible = True
'            B1_LINE.X1 = 900: B1_LINE.BorderWidth = 1
'            Call .PrintFld("B1")
'            Call .ClrLstFld("B*")
'            CisPrt.cLine = CisPrt.cLine + 1
'            B1_LINE.X1 = 375: B1_LINE.BorderWidth = 2
        Else
            B1_LINE.Visible = True
            B1_LINE.X1 = 375: B1_LINE.BorderWidth = 2
        End If
        
        wTSuryo(1) = wTSuryo(1) + wASuryo(1)
        wTKinGk(1) = wTKinGk(1) + wAKinGk(1)
        wTZenHi(1) = wTZenHi(1) + wAZenHi(1)
        wTSuryo(2) = wTSuryo(2) + wASuryo(2)
        wTKinGk(2) = wTKinGk(2) + wAKinGk(2)
        wTZenHi(2) = wTZenHi(2) + wAZenHi(2)
        wTSuryo(3) = wTSuryo(3) + wASuryo(3)
        wTKinGk(3) = wTKinGk(3) + wAKinGk(3)
        Erase wASuryo: Erase wAKinGk: Erase wAZenHi: wSyaCnt = 0
    End With
    
    If Kb = 0 Then Exit Sub
    
    With CisPrt
        
        B1_Hinbn = "[ å_ ñÒ êÊ åv ]"
        B1_Suryo(0) = Format(Format(wTSuryo(1), "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wTKinGk(1) / 1000
        B1_Kingk(0) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wTZenHi(1) / 1000
        B1_Zenhi(0) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        B1_Suryo(1) = Format(Format(wTSuryo(2), "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wTKinGk(2) / 1000
        B1_Kingk(1) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wTZenHi(2) / 1000
        B1_Zenhi(1) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        B1_Suryo(2) = Format(Format(wTSuryo(3), "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wTKinGk(3) / 1000
        B1_Kingk(2) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        
        B1_LINE.Visible = True
        B1_LINE.X1 = 375: B1_LINE.BorderWidth = 2
        Call .PrintFld("B1")
        Call .ClrLstFld("B*")
        CisPrt.cLine = CisPrt.cLine + 1
        B1_LINE.Visible = True
        
        wKSuryo(1) = wKSuryo(1) + wTSuryo(1)
        wKKinGk(1) = wKKinGk(1) + wTKinGk(1)
        wKZenHi(1) = wKZenHi(1) + wTZenHi(1)
        wKSuryo(2) = wKSuryo(2) + wTSuryo(2)
        wKKinGk(2) = wKKinGk(2) + wTKinGk(2)
        wKZenHi(2) = wKZenHi(2) + wTZenHi(2)
        wKSuryo(3) = wKSuryo(3) + wTSuryo(3)
        wKKinGk(3) = wKKinGk(3) + wTKinGk(3)
        Erase wTSuryo: Erase wTKinGk: Erase wTZenHi
    End With
    
    If Kb = 1 Then Exit Sub
    
    With CisPrt
        
        Call HeadWrite
        B1_Hinbn = "Ås çH ãÊ åv Åt"
        B1_Suryo(0) = Format(Format(wKSuryo(1), "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wKKinGk(1) / 1000
        B1_Kingk(0) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wKZenHi(1) / 1000
        B1_Zenhi(0) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        B1_Suryo(1) = Format(Format(wKSuryo(2), "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wKKinGk(2) / 1000
        B1_Kingk(1) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wKZenHi(2) / 1000
        B1_Zenhi(1) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        B1_Suryo(2) = Format(Format(wKSuryo(3), "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wKKinGk(3) / 1000
        B1_Kingk(2) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        
        B1_LINE.Visible = True
        B1_LINE.X1 = 375: B1_LINE.BorderWidth = 3
        Call .PrintFld("B1")
        Call .ClrLstFld("B*")
        CisPrt.cLine = CisPrt.cLine + 1
        
        wSSuryo(1) = wSSuryo(1) + wKSuryo(1)
        wSKinGk(1) = wSKinGk(1) + wKKinGk(1)
        wSZenHi(1) = wSZenHi(1) + wKZenHi(1)
        wSSuryo(2) = wSSuryo(2) + wKSuryo(2)
        wSKinGk(2) = wSKinGk(2) + wKKinGk(2)
        wSZenHi(2) = wSZenHi(2) + wKZenHi(2)
        wSSuryo(3) = wSSuryo(3) + wKSuryo(3)
        wSKinGk(3) = wSKinGk(3) + wKKinGk(3)
        Erase wKSuryo: Erase wKKinGk: Erase wKZenHi
    End With
    
    If Kb = 2 Then Exit Sub
    
    With CisPrt
        
        Call HeadWrite
        B1_Hinbn = "Åy ëç çá åv Åz"
        B1_Suryo(0) = Format(Format(wSSuryo(1), "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wSKinGk(1) / 1000
        B1_Kingk(0) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wSZenHi(1) / 1000
        B1_Zenhi(0) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        B1_Suryo(1) = Format(Format(wSSuryo(2), "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wSKinGk(2) / 1000
        B1_Kingk(1) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wSZenHi(2) / 1000
        B1_Zenhi(1) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        B1_Suryo(2) = Format(Format(wSSuryo(3), "#,##0;-#,##0;#"), String(9, "@"))
        gLong = wSKinGk(3) / 1000
        B1_Kingk(2) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
        
        B1_LINE.Visible = True
        B1_LINE.X1 = 375: B1_LINE.BorderWidth = 5
        Call .PrintFld("B1")
        Call .ClrLstFld("B*")
        CisPrt.cLine = CisPrt.cLine + 1
        
        Erase wSSuryo: Erase wSKinGk: Erase wSZenHi
    End With
    
End Sub

