VERSION 5.00
Begin VB.Form COS0510L 
   Caption         =   "ç∑àŸñæç◊ï\"
   ClientHeight    =   2850
   ClientLeft      =   -1950
   ClientTop       =   2265
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2850
   ScaleWidth      =   15420
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ç∑àŸñæç◊ï\"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   12
      Left            =   7440
      TabIndex        =   36
      Top             =   480
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "î≠çsì˙éûÅF"
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
      Index           =   22
      Left            =   12750
      TabIndex        =   35
      Top             =   465
      Width           =   900
   End
   Begin VB.Label H1_YMD 
      AutoSize        =   -1  'True
      Caption         =   "YYYY.MM.DD"
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
      Left            =   13695
      TabIndex        =   34
      Top             =   465
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÉyÅ[ÉW"
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
      Index           =   4
      Left            =   14835
      TabIndex        =   33
      Top             =   690
      Width           =   495
   End
   Begin VB.Label H1_PAGE 
      AutoSize        =   -1  'True
      Caption         =   "999"
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
      Left            =   14475
      TabIndex        =   32
      Top             =   675
      Width           =   270
   End
   Begin VB.Line H1_K_LINE 
      Index           =   8
      X1              =   12540
      X2              =   12540
      Y1              =   1815
      Y2              =   2055
   End
   Begin VB.Line B1_K_LINE 
      Index           =   6
      X1              =   12540
      X2              =   12540
      Y1              =   1875
      Y2              =   2445
   End
   Begin VB.Line H1_K_LINE 
      Index           =   3
      X1              =   8355
      X2              =   8355
      Y1              =   1800
      Y2              =   2040
   End
   Begin VB.Line B1_K_LINE 
      Index           =   1
      X1              =   8355
      X2              =   8355
      Y1              =   1875
      Y2              =   2445
   End
   Begin VB.Label B1_STanka 
      AutoSize        =   -1  'True
      Caption         =   "999,999,999.99"
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
      Left            =   11190
      TabIndex        =   31
      Top             =   2145
      Width           =   1260
   End
   Begin VB.Label B1_TTanka 
      AutoSize        =   -1  'True
      Caption         =   "999,999,999.99"
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
      Left            =   7005
      TabIndex        =   30
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "íPÅ@âø"
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
      Index           =   3
      Left            =   7725
      TabIndex        =   29
      Top             =   1845
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "íPÅ@âø"
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
      Index           =   0
      Left            =   11895
      TabIndex        =   28
      Top             =   1845
      Width           =   540
   End
   Begin VB.Line B1_K_LINE 
      Index           =   11
      X1              =   15420
      X2              =   15420
      Y1              =   1890
      Y2              =   2445
   End
   Begin VB.Line H1_K_LINE 
      Index           =   14
      X1              =   15420
      X2              =   15420
      Y1              =   1455
      Y2              =   2280
   End
   Begin VB.Line B1_K_LINE 
      Index           =   10
      X1              =   11085
      X2              =   11085
      Y1              =   1875
      Y2              =   2445
   End
   Begin VB.Line B1_K_LINE 
      Index           =   9
      X1              =   9720
      X2              =   9720
      Y1              =   1875
      Y2              =   2445
   End
   Begin VB.Line B1_K_LINE 
      Index           =   8
      X1              =   6900
      X2              =   6900
      Y1              =   1875
      Y2              =   2445
   End
   Begin VB.Line H1_K_LINE 
      Index           =   13
      X1              =   11085
      X2              =   11085
      Y1              =   1800
      Y2              =   2025
   End
   Begin VB.Line H1_K_LINE 
      Index           =   12
      X1              =   6900
      X2              =   6900
      Y1              =   1800
      Y2              =   2040
   End
   Begin VB.Line H1_K_LINE 
      Index           =   11
      X1              =   9720
      X2              =   9720
      Y1              =   1455
      Y2              =   2280
   End
   Begin VB.Label B1_Sagaku 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999,999"
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
      Left            =   14145
      TabIndex        =   27
      Top             =   2160
      Width           =   1170
   End
   Begin VB.Line B1_K_LINE 
      Index           =   7
      X1              =   13905
      X2              =   13905
      Y1              =   1890
      Y2              =   2445
   End
   Begin VB.Line B1_K_LINE 
      Index           =   5
      X1              =   4500
      X2              =   4500
      Y1              =   1875
      Y2              =   2445
   End
   Begin VB.Line B1_K_LINE 
      Index           =   4
      X1              =   5550
      X2              =   5550
      Y1              =   1875
      Y2              =   2445
   End
   Begin VB.Line B1_K_LINE 
      Index           =   3
      X1              =   3345
      X2              =   3345
      Y1              =   1890
      Y2              =   2445
   End
   Begin VB.Line B1_K_LINE 
      Index           =   2
      X1              =   2100
      X2              =   2100
      Y1              =   1890
      Y2              =   2445
   End
   Begin VB.Line H1_K_LINE 
      Index           =   10
      X1              =   5550
      X2              =   13905
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line H1_K_LINE 
      Index           =   9
      X1              =   13905
      X2              =   13905
      Y1              =   1455
      Y2              =   2280
   End
   Begin VB.Line H1_K_LINE 
      Index           =   7
      X1              =   4500
      X2              =   4500
      Y1              =   1455
      Y2              =   2280
   End
   Begin VB.Line H1_K_LINE 
      Index           =   6
      X1              =   5550
      X2              =   5550
      Y1              =   1455
      Y2              =   2280
   End
   Begin VB.Line H1_K_LINE 
      Index           =   5
      X1              =   3345
      X2              =   3345
      Y1              =   1455
      Y2              =   2280
   End
   Begin VB.Line H1_K_LINE 
      Index           =   4
      X1              =   2100
      X2              =   2100
      Y1              =   1455
      Y2              =   2280
   End
   Begin VB.Line T1_K_LINE 
      X1              =   645
      X2              =   15420
      Y1              =   2445
      Y2              =   2445
   End
   Begin VB.Line H1_K_LINE 
      Index           =   2
      X1              =   645
      X2              =   15420
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line B1_K_LINE 
      Index           =   0
      X1              =   645
      X2              =   645
      Y1              =   1890
      Y2              =   2445
   End
   Begin VB.Label B1_SSu 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999,999"
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
      Left            =   9825
      TabIndex        =   26
      Top             =   2160
      Width           =   1170
   End
   Begin VB.Label B1_SKingk 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999,999"
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
      Left            =   12645
      TabIndex        =   25
      Top             =   2160
      Width           =   1170
   End
   Begin VB.Line H1_K_LINE 
      Index           =   1
      X1              =   645
      X2              =   15420
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Line H1_K_LINE 
      Index           =   0
      X1              =   645
      X2              =   645
      Y1              =   1455
      Y2              =   2280
   End
   Begin VB.Label B1_TKingk 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999,999"
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
      Left            =   8460
      TabIndex        =   24
      Top             =   2160
      Width           =   1170
   End
   Begin VB.Label B1_TSu 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999,999"
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
      Left            =   5655
      TabIndex        =   23
      Top             =   2160
      Width           =   1170
   End
   Begin VB.Label B1_Sikyu 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
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
      Left            =   4875
      TabIndex        =   22
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label B1_NouNo 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1"
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
      Left            =   2250
      TabIndex        =   21
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label B1_HINBN 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXX"
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
      Left            =   750
      TabIndex        =   20
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Label B1_NYMD 
      AutoSize        =   -1  'True
      Caption         =   "YY-MM"
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
      Left            =   3720
      TabIndex        =   19
      Top             =   2160
      Width           =   450
   End
   Begin VB.Label H1_KCODE 
      AutoSize        =   -1  'True
      Caption         =   "XXX:XXXXXXXXX1XXXXXXXXX2"
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
      Left            =   1485
      TabIndex        =   18
      Top             =   555
      Width           =   2880
   End
   Begin VB.Label H1_TORCD 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX:XXXXXXXXXXXXXXXXXXXX"
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
      Left            =   1485
      TabIndex        =   17
      Top             =   255
      Width           =   3120
   End
   Begin VB.Label H1_YY 
      AutoSize        =   -1  'True
      Caption         =   "YY"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6150
      TabIndex        =   16
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ç∑äzã‡äz"
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
      Index           =   21
      Left            =   14595
      TabIndex        =   15
      Top             =   1845
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "êîÅ@ó "
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
      Index           =   20
      Left            =   6270
      TabIndex        =   14
      Top             =   1845
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ã‡Å@äz"
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
      Index           =   19
      Left            =   9075
      TabIndex        =   13
      Top             =   1845
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "êîÅ@ó "
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
      Index           =   18
      Left            =   10455
      TabIndex        =   12
      Top             =   1845
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ã‡Å@äz"
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
      Index           =   17
      Left            =   13275
      TabIndex        =   11
      Top             =   1845
      Width           =   540
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
      Index           =   16
      Left            =   660
      TabIndex        =   10
      Top             =   255
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "åéìx"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   13
      Left            =   6585
      TabIndex        =   9
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ãÊï™"
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
      Left            =   660
      TabIndex        =   8
      Top             =   555
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "î[ì¸(éÛì¸)"
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
      Index           =   10
      Left            =   3495
      TabIndex        =   7
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ïiÅ@Å@î‘"
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
      Index           =   9
      Left            =   1020
      TabIndex        =   6
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "åéÅ@ì˙"
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
      Index           =   8
      Left            =   3675
      TabIndex        =   5
      Top             =   1845
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "î[ì¸áÇ"
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
      Index           =   6
      Left            =   2430
      TabIndex        =   4
      Top             =   1575
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "éÛì¸èÍèä"
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
      Index           =   5
      Left            =   4695
      TabIndex        =   3
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label H1_TORNM 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "ìæà”êÊÅiÇ`Åj"
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
      Left            =   7140
      TabIndex        =   2
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "édì¸êÊÅiÇaÅj"
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
      Index           =   2
      Left            =   11295
      TabIndex        =   1
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÅiÇ`ÅjÅ|ÅiÇaÅj"
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
      Index           =   1
      Left            =   14010
      TabIndex        =   0
      Top             =   1560
      Width           =   1260
   End
End
Attribute VB_Name = "COS0510L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  ç∑àŸñæç◊ÉäÉXÉg
'**       ÉtÉHÅ[ÉÄID    :  COS0510L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2004/11/17   By CIS
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
    Dim wCnt            As Long
    Dim OutStr          As String   ' CSVèoóÕì‡óe
    Dim Sl_Select       As String

    Dim BK_Torcd        As String   ' ìæà”êÊ
    Dim BK_KCODE        As String   ' ãÊï™∫∞ƒﬁ
    Dim wSagaku         As Currency ' ç∑äzã‡äz
    Dim mTornm          As String   ' î≠íçâÔé–ñº

    Dim mPrtDateTime    As String
'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub Form_Load()
    App.Title = "îÉä|ã‡ñæç◊ÉäÉXÉg"
    
    If Not DataRead Then Exit Sub
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    With CisPrt
        .iPrtTitle = "îÑè„ç∑àŸñæç◊ÉäÉXÉg"
        .iLineMax = 35
        .iLinePitch = 6
'        .iLineInterval = 1
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
        End If
        .cLine = 0
        .cPage = 0
        .cDataMaxCnt = CisDB.RecordCount
    End With
    
    Call PrintProc
    
End Sub
'*-------------------------*
'*      àÛç¸ÉfÅ[É^ì«çû
'*-------------------------*
Private Function DataRead()
    DataRead = False

    DataRead = True
End Function
'*---------------------*
'*Å@  àÛ  ç¸  èà  óù
'*---------------------*
Private Function PrintProc()
    
    mPrtDateTime = Format(Now(), "YYYY.MM.DD hh:nn:ss")

    Call CisPrt.ClrLstFld("B*")
    CisPrt.cLine = 99: gCnt1 = 0
    
    Do Until Not UST_RDSTS
        With UST
            Call HeadWrite(0)
            
            If gCnt1 <> 0 And (BK_KCODE <> Mid(.è∆çáåãâ , 1, 1) Or BK_Torcd <> .å_ñÒêÊ) Then
                With CisPrt
                    If .cLine <> .iLineMax Then
                        Do Until .cLine = .iLineMax
                            Call .PrintFld("B1")
                            .cLine = .cLine + 1
                            Call .ClrLstFld("B*")
                        Loop
                    End If
    
                    If .cLine = .iLineMax Then
                       .cLine = .cLine - 1
                       Call .PrintFld("T1")
                       .cLine = .cLine + 1
                    End If
    
                    .cLine = 99
                End With
            End If

            Call HeadWrite(0)
            
            B1_NYMD = Mid(.ç≈ëÂì˙ït, 5, 2) & "-" & Mid(.ç≈ëÂì˙ït, 7, 2)
            If Trim(.ï\é¶ïiî‘) = "" Then
                If Len(Trim(.ïiî‘)) > 10 Then
                    B1_HINBN = Mid(.ïiî‘, 1, 5) & "-" & Mid(.ïiî‘, 6, 5) & "-" & Mid(.ïiî‘, 11, 2)
                Else
                    B1_HINBN = Mid(.ïiî‘, 1, 5) & "-" & Mid(.ïiî‘, 6, 5)
                End If
            Else
                B1_HINBN = .ï\é¶ïiî‘
            End If
            
            B1_NouNo = Trim(.î[ì¸î‘çÜ)
            B1_Sikyu = Trim(.éÛì¸)
            'å_ñÒêÊ
            B1_TSu = Format(Format(.å_ñÒêÊêîó , "#,##0;-#,##0;#"), String(13, "@"))
            B1_TTanka = Format(Format(.å_ñÒêÊíPâø, "#,##0.00;-#,##0.00;#"), String(14, "@"))
            B1_TKingk = Format(Format(.å_ñÒêÊã‡äz, "#,##0;-#,##0;#"), String(13, "@"))
            'édì¸êÊ
            B1_SSu = Format(Format(.é©é–êîó , "#,##0;-#,##0;#"), String(13, "@"))
            B1_STanka = Format(Format(.é©é–íPâø, "#,##0.00;-#,##0.00;#"), String(14, "@"))
            B1_SKingk = Format(Format(.é©é–ã‡äz, "#,##0;-#,##0;#"), String(13, "@"))
            'ç∑äzã‡äz
            wSagaku = .å_ñÒêÊã‡äz - .é©é–ã‡äz
            B1_Sagaku = Format(Format(wSagaku, "#,##0;-#,##0;#"), String(13, "@"))

            wSagaku = 0
            BK_Torcd = RTrim(.å_ñÒêÊ)
            BK_KCODE = Mid(.è∆çáåãâ , 1, 1)
        End With

        Call CisPrt.PrintFld("B1")
        CisPrt.cLine = CisPrt.cLine + 1
        Call CisPrt.ClrLstFld("B*")

        With CisPrt
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(CisPrt.cDataMaxCnt, "#,###")
        End With

        Call USTReadNext
        gCnt1 = gCnt1 + 1
    Loop
    Call USTClose

    With CisPrt
        If .cLine <> .iLineMax Then
            Do Until .cLine = .iLineMax
                Call .PrintFld("B1")
                .cLine = .cLine + 1
                Call .ClrLstFld("B*")
            Loop
        End If

        If .cLine = .iLineMax Then
           .cLine = .cLine - 1
           Call .PrintFld("T1")
           .cLine = .cLine + 1
        End If

        .cLine = 99
    End With
    
    ' àÛç¸èIóπ
    Call CisPrt.PrintEndDoc
    
End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite(KB As Byte)
    HeadWrite = False
    With CisPrt
        ' ÇPï≈ì‡àÛç¸çsêî¡™Ø∏
        If .iLineMax >= .cLine + 1 Then Exit Function
        
        If .iLineMax = .cLine Then
            .cLine = .cLine - 1
            Call .PrintFld("T1")
            .cLine = .cLine + 1
        End If
        .cLine = 0
        .cPage = .cPage + 1
        
        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé
    
        With UST
            H1_YMD = mPrtDateTime
            H1_PAGE = CisFun.RSetFld(CisPrt.cPage, 3, "###")
            
            H1_YY = Mid(.è∆çáîNåé, 5, 2)
            H1_TORCD = Trim(.å_ñÒêÊ) & ":" & RTrim(.å_ñÒêÊñº)
            H1_KCODE = Trim(.è∆çáåãâ ñº)
            mTornm = "å_ñÒêÊÅiÇ`Åj"
            H1_TORNM = mTornm
        End With
    
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With
    
    HeadWrite = True
End Function
