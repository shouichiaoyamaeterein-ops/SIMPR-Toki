VERSION 5.00
Begin VB.Form COY0310L 
   Caption         =   "îÑè„é¿ê—àÍóóï\"
   ClientHeight    =   3330
   ClientLeft      =   390
   ClientTop       =   2400
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   14610
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
      Left            =   1410
      TabIndex        =   43
      Top             =   1095
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
      Left            =   735
      TabIndex        =   42
      Top             =   1110
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
      Left            =   10170
      TabIndex        =   41
      Top             =   630
      Width           =   1365
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
      Index           =   21
      Left            =   11850
      TabIndex        =   38
      Top             =   1500
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
      Index           =   19
      Left            =   12960
      TabIndex        =   37
      Top             =   1500
      Width           =   570
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   12
      X1              =   11475
      X2              =   11975
      Y1              =   1290
      Y2              =   1290
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
      Left            =   12375
      TabIndex        =   36
      Top             =   1155
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
      Index           =   17
      Left            =   12765
      TabIndex        =   35
      Top             =   1200
      Width           =   195
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   11
      X1              =   11475
      X2              =   11475
      Y1              =   1290
      Y2              =   1490
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   10
      X1              =   13170
      X2              =   13670
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   9
      X1              =   13665
      X2              =   13665
      Y1              =   1290
      Y2              =   1490
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
      Left            =   8355
      TabIndex        =   32
      Top             =   1500
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
      Left            =   10485
      TabIndex        =   31
      Top             =   1500
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
      Left            =   9465
      TabIndex        =   30
      Top             =   1500
      Width           =   570
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   8
      X1              =   8025
      X2              =   9025
      Y1              =   1290
      Y2              =   1290
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
      Left            =   9375
      TabIndex        =   29
      Top             =   1155
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
      Left            =   9765
      TabIndex        =   28
      Top             =   1200
      Width           =   195
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   7
      X1              =   8025
      X2              =   8025
      Y1              =   1290
      Y2              =   1490
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   6
      X1              =   10230
      X2              =   11230
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   5
      X1              =   11235
      X2              =   11235
      Y1              =   1290
      Y2              =   1490
   End
   Begin VB.Label B1_Suryo 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Left            =   11520
      TabIndex        =   25
      Top             =   1860
      Width           =   945
   End
   Begin VB.Label B1_Kingk 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   12600
      TabIndex        =   24
      Top             =   1860
      Width           =   945
   End
   Begin VB.Label B1_Zenhi 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Index           =   1
      Left            =   10245
      TabIndex        =   23
      Top             =   1875
      Width           =   945
   End
   Begin VB.Label B1_Suryo 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Index           =   1
      Left            =   8040
      TabIndex        =   22
      Top             =   1875
      Width           =   945
   End
   Begin VB.Label B1_Kingk 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Index           =   1
      Left            =   9120
      TabIndex        =   21
      Top             =   1875
      Width           =   945
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   4
      X1              =   7785
      X2              =   7785
      Y1              =   1290
      Y2              =   1490
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   2
      X1              =   6780
      X2              =   7780
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   3
      X1              =   4575
      X2              =   4575
      Y1              =   1290
      Y2              =   1490
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
      Left            =   6315
      TabIndex        =   20
      Top             =   1200
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
      Left            =   5925
      TabIndex        =   19
      Top             =   1155
      Width           =   300
   End
   Begin VB.Label B1_Kingk 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Left            =   5685
      TabIndex        =   18
      Top             =   1890
      Width           =   945
   End
   Begin VB.Line B1_LINE 
      Visible         =   0   'False
      X1              =   3585
      X2              =   14000
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   4
      Index           =   1
      X1              =   555
      X2              =   14000
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   0
      X1              =   4575
      X2              =   5575
      Y1              =   1290
      Y2              =   1290
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
      Left            =   1800
      TabIndex        =   17
      Top             =   720
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
      Left            =   720
      TabIndex        =   16
      Top             =   720
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
      Left            =   1395
      TabIndex        =   15
      Top             =   720
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
      Left            =   2130
      TabIndex        =   14
      Top             =   705
      Width           =   570
   End
   Begin VB.Label B1_Syakei 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
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
      Left            =   3660
      TabIndex        =   13
      Top             =   1890
      Width           =   525
   End
   Begin VB.Label B1_Suryo 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Left            =   4560
      TabIndex        =   12
      Top             =   1890
      Width           =   945
   End
   Begin VB.Label B1_Zenhi 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Left            =   6810
      TabIndex        =   11
      Top             =   1890
      Width           =   945
   End
   Begin VB.Label B1_TORCD 
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
      Left            =   600
      TabIndex        =   10
      Top             =   1890
      Width           =   735
   End
   Begin VB.Label B1_TORNM 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
      Left            =   1380
      TabIndex        =   9
      Top             =   1890
      Width           =   2100
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
      Left            =   13935
      TabIndex        =   8
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
      Left            =   12015
      TabIndex        =   7
      Top             =   600
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "å_Å@ñÒÅ@êÊ"
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
      Left            =   1215
      TabIndex        =   6
      Top             =   1515
      Width           =   1125
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
      Left            =   6015
      TabIndex        =   5
      Top             =   1500
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Left            =   3690
      TabIndex        =   4
      Top             =   1500
      Width           =   120
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
      Left            =   7035
      TabIndex        =   3
      Top             =   1500
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
      Left            =   4905
      TabIndex        =   2
      Top             =   1500
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
      Left            =   14565
      TabIndex        =   1
      Top             =   600
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Å¶ îÑè„ó\ëzèWåvï\ÅyïîïiÅz Å¶"
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
      Left            =   4290
      TabIndex        =   0
      Top             =   600
      Width           =   5040
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
      Left            =   5490
      TabIndex        =   26
      Top             =   1185
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
      Left            =   6660
      TabIndex        =   27
      Top             =   1185
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
      Left            =   8940
      TabIndex        =   33
      Top             =   1185
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
      Left            =   10110
      TabIndex        =   34
      Top             =   1185
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
      Index           =   23
      Left            =   13080
      TabIndex        =   40
      Top             =   1185
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
      Index           =   22
      Left            =   11910
      TabIndex        =   39
      Top             =   1185
      Width           =   225
   End
End
Attribute VB_Name = "COY0310L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  îÑè„ó\ëzï\ÅyïîïiÅEèWåvÅz
'**       ÉtÉHÅ[ÉÄID    :  COY0310L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2004/10/05
'**       ïœ  çX  ì˙    :  2008/04/11 é‘å^5åÖ
'**
'**********************************************************'
    Dim OutStr          As String   ' CSVèoóÕì‡óe
    Dim KeyKouku        As String
    Dim KeyTorcd        As String
    
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
    Dim mPrtCnt          As Long
'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub Form_Load()
    App.Title = "îÑè„ó\ëzèWåvï\ÅyïîïiÅz"
    
    If Not DataRead Then Exit Sub
    
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    mPrtCnt = CisDB.RecordCount

    With CisPrt
        .iPrtTitle = "îÑè„ó\ëzèWåvï\ÅyïîïiÅz"
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
    
    Call SQLEdit0310
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
    wInCnt = 0: KeyKouku = "": KeyTorcd = ""
    Erase wKSuryo: Erase wKKinGk: Erase wKZenHi
    Erase wSSuryo: Erase wSKinGk: Erase wSZenHi
    
    Do Until Not UYT_RDSTS
        With UYT
            
            If RTrim(KeyKouku) <> RTrim(.çHãÊ) Or RTrim(KeyTorcd) <> RTrim(.å_ñÒêÊ) Then
                If wInCnt <> 0 Then
                    If RTrim(KeyKouku) <> RTrim(.çHãÊ) Then
                        Call TotalWrite(1)
                        CisPrt.cLine = 99
                        Call HeadWrite
                    Else
                        Call TotalWrite(0)
                    End If
                End If
                Erase wTSuryo: Erase wTKinGk: Erase wTZenHi
                
            End If
            Call HeadWrite
            If RTrim(KeyTorcd) <> RTrim(.å_ñÒêÊ) Then
                If Trim(.å_ñÒêÊ) <> "" Then
                    B1_TORCD = RTrim(.å_ñÒêÊ)
                End If
                If Trim(.å_ñÒêÊñº) <> "" Then
                    B1_TORNM = RTrim(.å_ñÒêÊñº)
                End If
            End If
            
            B1_Syakei = RTrim(.é‘å^)
            
            B1_Suryo(0) = Format(Format(.êîó 1, "#,##0;-#,##0;#"), String(9, "@"))
            gLong = .ã‡äz1 / 1000
            B1_Kingk(0) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
            gLong = .ëOåéî‰1 / 1000
            B1_Zenhi(0) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
            B1_Suryo(1) = Format(Format(.êîó 2, "#,##0;-#,##0;#"), String(9, "@"))
            gLong = .ã‡äz2 / 1000
            B1_Kingk(1) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
            gLong = .ëOåéî‰2 / 1000
            B1_Zenhi(1) = Format(Format(gLong, "#,##0;-#,##0;#"), String(9, "@"))
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
            
            wTSuryo(1) = wTSuryo(1) + .êîó 1
            wTKinGk(1) = wTKinGk(1) + .ã‡äz1
            wTZenHi(1) = wTZenHi(1) + .ëOåéî‰1
            wTSuryo(2) = wTSuryo(2) + .êîó 2
            wTKinGk(2) = wTKinGk(2) + .ã‡äz2
            wTZenHi(2) = wTZenHi(2) + .ëOåéî‰2
            wTSuryo(3) = wTSuryo(3) + .êîó 3
            wTKinGk(3) = wTKinGk(3) + .ã‡äz3
            
            wInCnt = wInCnt + 1
            KeyKouku = .çHãÊ
            KeyTorcd = .å_ñÒêÊ
        End With

        Call UYTReadNext
    Loop
    Call UYTClose
    
    Call HeadWrite
    Call TotalWrite(2)
    
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
            H1_NJYm(0) = Mid(.ì‡é¶îNåé, 5, 2)
            H1_NJYm(1) = Format(DateAdd("m", 1, CDate(Left(.ì‡é¶îNåé, 4) & "/" & Mid(.ì‡é¶îNåé, 5, 2) & "/01")), "mm")
            H1_NJYm(2) = Format(DateAdd("m", 2, CDate(Left(.ì‡é¶îNåé, 4) & "/" & Mid(.ì‡é¶îNåé, 5, 2) & "/01")), "mm")
        End With
    
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With
    
    KeyTorcd = ""
    B1_LINE.Visible = False
    HeadWrite = True
End Function
'*----------------*
'*   çáåvàÛç¸     *
'*----------------*
Private Sub TotalWrite(Optional Kb As Integer = 0)
    With CisPrt
        
        B1_TORNM = "Ås å_ ñÒ êÊ åvÅt"
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
        B1_LINE.X1 = 1350: B1_LINE.BorderWidth = 1
        Call .PrintFld("B1")
        Call .ClrLstFld("B*")
        CisPrt.cLine = CisPrt.cLine + 1
        B1_LINE.X1 = 555: B1_LINE.BorderWidth = 2
        
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
    
    If Kb = 0 Then Exit Sub
    
    With CisPrt
        
        Call HeadWrite
        B1_TORNM = "Åy çH ãÊ åv Åz"
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
        B1_LINE.X1 = 555: B1_LINE.BorderWidth = 3
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
    
    If Kb = 1 Then Exit Sub
    
    With CisPrt
        
        Call HeadWrite
        B1_TORNM = "[ ëç çá åv ]"
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
        B1_LINE.X1 = 555: B1_LINE.BorderWidth = 5
        Call .PrintFld("B1")
        Call .ClrLstFld("B*")
        CisPrt.cLine = CisPrt.cLine + 1
        B1_LINE.Visible = False
        
        Erase wSSuryo: Erase wSKinGk: Erase wSZenHi
    End With
    
End Sub

