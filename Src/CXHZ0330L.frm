VERSION 5.00
Begin VB.Form CXHZ0330L 
   Caption         =   "明細書"
   ClientHeight    =   5940
   ClientLeft      =   -15
   ClientTop       =   3285
   ClientWidth     =   18435
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   18435
   Begin VB.Label H1_Nonyu 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX-XX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   10695
      TabIndex        =   55
      Top             =   195
      Width           =   720
   End
   Begin VB.Label H1_Nonyunm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9945
      TabIndex        =   54
      Top             =   435
      Width           =   2400
   End
   Begin VB.Label H1_Nonyukj 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9945
      TabIndex        =   53
      Top             =   705
      Width           =   2400
   End
   Begin VB.Label H1_NonyuM 
      Caption         =   "手配商社"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9915
      TabIndex        =   52
      Top             =   180
      Width           =   735
   End
   Begin VB.Label H1_NOncyu 
      Caption         =   "御中"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   12420
      TabIndex        =   51
      Top             =   690
      Width           =   465
   End
   Begin VB.Line H1_K_NLINE 
      Index           =   1
      X1              =   9870
      X2              =   12975
      Y1              =   390
      Y2              =   390
   End
   Begin VB.Line H1_K_NLINE 
      Index           =   3
      X1              =   9870
      X2              =   12975
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line H1_K_NLINE 
      Index           =   0
      X1              =   9870
      X2              =   9870
      Y1              =   390
      Y2              =   960
   End
   Begin VB.Line H1_K_NLINE 
      Index           =   2
      X1              =   12975
      X2              =   12975
      Y1              =   390
      Y2              =   960
   End
   Begin VB.Label B1_Tani 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   13950
      TabIndex        =   50
      Top             =   2100
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "材　種"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   7125
      TabIndex        =   49
      Top             =   1755
      Width           =   990
   End
   Begin VB.Label B1_Zaisy 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6540
      TabIndex        =   48
      Top             =   2100
      Width           =   1800
   End
   Begin VB.Label Label4 
      Caption         =   "寸　　法"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   4590
      TabIndex        =   47
      Top             =   1755
      Width           =   990
   End
   Begin VB.Label B1_Sunpo 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3765
      TabIndex        =   46
      Top             =   2100
      Width           =   2760
   End
   Begin VB.Label B1_Kensu 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "9,999"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7740
      TabIndex        =   45
      Top             =   2085
      Width           =   600
   End
   Begin VB.Label B1_KensuM 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "《頁件数》"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6570
      TabIndex        =   44
      Top             =   2115
      Width           =   930
   End
   Begin VB.Label H1_Page 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14580
      TabIndex        =   43
      Top             =   1185
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "頁"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   15555
      TabIndex        =   42
      Top             =   1200
      Width           =   225
   End
   Begin VB.Line B1_K_LINE2 
      X1              =   1080
      X2              =   16500
      Y1              =   2085
      Y2              =   2085
   End
   Begin VB.Label B1_HakoKei 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "《箱数頁合計》"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9300
      TabIndex        =   41
      Top             =   2130
      Width           =   1365
   End
   Begin VB.Line H1_K_LINE 
      Index           =   5
      X1              =   3285
      X2              =   3285
      Y1              =   390
      Y2              =   960
   End
   Begin VB.Line H1_K_LINE 
      Index           =   4
      X1              =   180
      X2              =   180
      Y1              =   390
      Y2              =   960
   End
   Begin VB.Line H1_K_LINE 
      Index           =   3
      X1              =   180
      X2              =   3285
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line H1_K_LINE 
      Index           =   2
      X1              =   180
      X2              =   3285
      Y1              =   390
      Y2              =   390
   End
   Begin VB.Label B1_Hasu 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "9,999.999"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   11415
      TabIndex        =   40
      Top             =   2100
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "端数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   11940
      TabIndex        =   39
      Top             =   1755
      Width           =   435
   End
   Begin VB.Line B1_K_LINE1 
      X1              =   195
      X2              =   16500
      Y1              =   2085
      Y2              =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "受入"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   7350
      TabIndex        =   38
      Top             =   1200
      Width           =   435
   End
   Begin VB.Label H1_Ukeir 
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7845
      TabIndex        =   37
      Top             =   1065
      Width           =   390
   End
   Begin VB.Label B1_Bikou 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   14355
      TabIndex        =   36
      Top             =   2100
      Width           =   1800
   End
   Begin VB.Label B1_HSuryo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999,999.999"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12555
      TabIndex        =   35
      Top             =   2100
      Width           =   1320
   End
   Begin VB.Label B1_Maisu 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "99,999"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10680
      TabIndex        =   34
      Top             =   2100
      Width           =   720
   End
   Begin VB.Label B1_Syuyo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999,999.999"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9375
      TabIndex        =   33
      Top             =   2100
      Width           =   1320
   End
   Begin VB.Label B1_Seban 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8370
      TabIndex        =   32
      Top             =   2100
      Width           =   960
   End
   Begin VB.Label B1_Zaist 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1245
      TabIndex        =   31
      Top             =   2100
      Width           =   3000
   End
   Begin VB.Label B1_NPage 
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   945
      TabIndex        =   30
      Top             =   2100
      Width           =   240
   End
   Begin VB.Label B1_Nouno 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   29
      Top             =   2100
      Width           =   600
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   1
      X1              =   195
      X2              =   16500
      Y1              =   2055
      Y2              =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "備　考"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   14640
      TabIndex        =   28
      Top             =   1755
      Width           =   810
   End
   Begin VB.Label Label4 
      Caption         =   "納入数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   13095
      TabIndex        =   27
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "箱数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   10965
      TabIndex        =   26
      Top             =   1755
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "収容数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   9945
      TabIndex        =   25
      Top             =   1755
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "背番号"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   8520
      TabIndex        =   24
      Top             =   1755
      Width           =   630
   End
   Begin VB.Label Label4 
      Caption         =   "材　　質"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1680
      TabIndex        =   23
      Top             =   1755
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "頁"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   945
      TabIndex        =   22
      Top             =   1755
      Width           =   225
   End
   Begin VB.Label Label4 
      Caption         =   "御中"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   2730
      TabIndex        =   21
      Top             =   690
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "納品書№"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   1755
      Width           =   645
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   0
      X1              =   195
      X2              =   16500
      Y1              =   1605
      Y2              =   1605
   End
   Begin VB.Label H1_NBinM 
      Caption         =   "便"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6555
      TabIndex        =   19
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "日"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   5745
      TabIndex        =   18
      Top             =   1200
      Width           =   225
   End
   Begin VB.Label Label4 
      Caption         =   "月"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5040
      TabIndex        =   17
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "年"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   4380
      TabIndex        =   16
      Top             =   1200
      Width           =   225
   End
   Begin VB.Label H1_NBin 
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6150
      TabIndex        =   15
      Top             =   1065
      Width           =   390
   End
   Begin VB.Label H1_NDd 
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5340
      TabIndex        =   14
      Top             =   1065
      Width           =   390
   End
   Begin VB.Label Label4 
      Caption         =   "納入日"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   3000
      TabIndex        =   13
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label H1_NMm 
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4635
      TabIndex        =   12
      Top             =   1065
      Width           =   390
   End
   Begin VB.Label H1_NYy 
      AutoSize        =   -1  'True
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      TabIndex        =   11
      Top             =   1065
      Width           =   780
   End
   Begin VB.Label H1_TEKbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1365
      TabIndex        =   10
      Top             =   1065
      Width           =   720
   End
   Begin VB.Label H1_HCKbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   255
      TabIndex        =   9
      Top             =   1065
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "仕入先"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   8
      Top             =   180
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "発行日"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   11
      Left            =   13500
      TabIndex        =   7
      Top             =   120
      Width           =   540
   End
   Begin VB.Label H1_HYmd 
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99 hh:mm"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14130
      TabIndex        =   6
      Top             =   90
      Width           =   1680
   End
   Begin VB.Label H1_JisyaKj 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   13500
      TabIndex        =   5
      Top             =   660
      Width           =   2400
   End
   Begin VB.Label H1_SiireKj 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   255
      TabIndex        =   4
      Top             =   705
      Width           =   2400
   End
   Begin VB.Label H1_SiireNm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   255
      TabIndex        =   3
      Top             =   435
      Width           =   2400
   End
   Begin VB.Label H1_Siire 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX-XX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   870
      TabIndex        =   2
      Top             =   195
      Width           =   720
   End
   Begin VB.Label H1_JisyaNm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   13500
      TabIndex        =   1
      Top             =   405
      Width           =   2400
   End
   Begin VB.Label H1_SKKbn 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "発注明細書"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   26.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6150
      TabIndex        =   0
      Top             =   315
      Width           =   2700
   End
End
Attribute VB_Name = "CXHZ0330L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  [材料]発注明細書　発行
'**       フォームID    :  CXHZ0330L
'**       処理概要      :
'**
'**       作  成  日    :  2004/02/11
'**       変  更  日    :
'**
'**********************************************************'
'#比較退避用#
    Dim Key_MeiNo       As Long
    Dim Key_Nouno       As String
    Dim Key_NPage       As Byte
    Dim Key_Siire       As String
    Dim Key_SKKbn       As Byte
    Dim GPMaisu         As Long         '頁内箱数合計
    Dim GTMaisu         As Long         '明細№箱数合計
    Dim GPKensu         As Long         '頁内件数
    Dim GTKensu         As Long         '明細№件数
    Dim TotalPage       As Long         '仕入先総頁数
    Dim SiirePage       As Long         '仕入先現頁数
    Dim MeiPage         As Long         '明細書頁数
    
    Dim mPrtDate        As String
    Dim mPrtTime        As String
    Dim mJisyaNm        As String
    Dim mJisyaKj        As String

'*------------------------------------*
'*      Ｆｏｒｍ＿Ｌｏａｄ イベント     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")             '発行日
    mPrtTime = Format(Now(), "hh:nn")                  '発行時間
    '*---------------------*
    '*  プリント系初期設定
    '*---------------------*
    With CisPrtCXHZ0330
        .iPrtTitle = "[材料]発注明細書"
        .iLineMax = 35
        .iLinePitch = 6
        .iPaperSize = p_A4
        .iOrientation = po_Landscape
        
        If Not .PrintIni Then
           With CisFun
                .MB_Lines = 5
                .MB_MSG(1) = "     　プリンタ初期設定エラー        "
                .MB_Title = "印刷ｴﾗｰ"
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
'*---------------------------------------------------------------*
'*　 << 印  刷  処  理 >>
'*---------------------------------------------------------------*
Private Function PrintProc()
'*---------------------------------*
'*  自社名称・工場名を求める
'*---------------------------------*
' システム情報(自社工場)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    Call IRNRead("", 1)
' 取引先マスタより名称を求める
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(gCompany) & Left(IRN.oyian, 1) & "'"
    If TRMRead(gSL_Select, 1) Then
        mJisyaNm = TRM.略称
        mJisyaKj = TRM.工場名
    End If
    
'*---------------------------------*
'*  初期設定
'*---------------------------------*
    GPMaisu = 0: GPMaisu = 0: GPKensu = 0: GPKensu = 0
    TotalPage = 0: SiirePage = 0: MeiPage = 0
    Key_MeiNo = 0: Key_Nouno = "": Key_NPage = 0: Key_SKKbn = 0
    Call CisPrtCXHZ0330.ClrLstFld("B1")
    
'*---------------------------------*
'*  材料発注テーブル読込み
'*---------------------------------*
    Do Until Not ZCT_RDSTS
        If Key_MeiNo <> ZCT.明細書NO Or CisPrtCXHZ0330.cLine >= CisPrtCXHZ0330.iLineMax Then
            If Key_MeiNo <> 0 Then
                If Key_MeiNo <> ZCT.明細書NO Then
                    B1_K_LINE1.Visible = True
                    If MeiPage > 1 Then
                        B1_KensuM = "《頁件数》"
                        B1_Kensu = Format(GPKensu, "#,##0")
                        If Key_SKKbn = 0 Then
                            B1_HakoKei = "《箱数頁合計》"
                            B1_Maisu = Format(GPMaisu, "##,##0")
                        End If
                        Call CisPrtCXHZ0330.PrintFld("B1*")
                        CisPrtCXHZ0330.cLine = CisPrtCXHZ0330.cLine + 1
                        Call CisPrtCXHZ0330.ClrLstFld("B1")
                        B1_K_LINE1.Visible = False
                    End If
                '明細№合計印刷
                    B1_K_LINE2.Visible = False
                    B1_KensuM = "【総件数】"
                    B1_Kensu = Format(GTKensu, "#,##0")
                    If Key_SKKbn = 0 Then
                        B1_HakoKei = "【箱数総合計】"
                        B1_Maisu = Format(GTMaisu, "##,##0")
                    Else
                        B1_HakoKei = "": B1_Maisu = ""
                    End If
                    Call CisPrtCXHZ0330.PrintFld("B1*")
                    GTMaisu = 0: GTKensu = 0: MeiPage = 0
                Else
                    '頁合計印刷
                    B1_K_LINE1.Visible = True
                    B1_KensuM = "《頁件数》"
                    B1_Kensu = Format(GPKensu, "#,##0")
                    If Key_SKKbn = 0 Then
                        B1_HakoKei = "《箱数頁合計》"
                        B1_Maisu = Format(GPMaisu, "##,##0")
                    Else
                        B1_HakoKei = "": B1_Maisu = ""
                    End If
                    Call CisPrtCXHZ0330.PrintFld("B1*")
                    CisPrtCXHZ0330.cLine = CisPrtCXHZ0330.cLine + 1
                End If
                GPMaisu = 0: GPKensu = 0
            End If
            Call HeadWrite(0)
        End If
        B1_K_LINE1.Visible = False: B1_K_LINE2.Visible = False
        B1_HakoKei = "": B1_KensuM = "": B1_Kensu = ""
        With ZCT
        '*---------------------------------*
        '*  Body行　印刷
        '*---------------------------------*
            If RTrim(Key_Nouno) <> RTrim(.納入番号) Then
                B1_K_LINE1.Visible = True
                B1_Nouno = RTrim(.納入番号)
            End If
            If RTrim(Key_Nouno) <> RTrim(.納入番号) Or Key_NPage <> .納品書頁 Then
                B1_NPage = RTrim(.納品書頁)
                If Key_NPage <> .納品書頁 Then
                    B1_K_LINE2.Visible = True
                End If
            End If
            B1_Zaist = RTrim(.材質)
            B1_Sunpo = RTrim(.表示寸法)
            If RTrim(.材種コード) = "" Then
                B1_Zaisy = RTrim(.材種名)
            Else
                B1_Zaisy = RTrim(.材種コード)
            End If
            B1_Seban = RTrim(.背番号)
            If .枚数 <> 0 Then
                B1_Syuyo = Format(.収容数, "##,##0.000")
            End If
            If .枚数 <> 0 Then
                B1_Maisu = Format(.枚数, "##,##0")
                GPMaisu = GPMaisu + .枚数
                GTMaisu = GTMaisu + .枚数
            End If
            If .端数 <> 0 Then
                B1_Hasu = Format(.端数, "##,##0.000")
            End If
            If .発注数 <> 0 Then
                B1_HSuryo = Format(.発注数, "##,##0.000")
            End If
            B1_Tani = RTrim(.手配単位名)
           '備考（打切情報）
            gStr = ""
            Select Case .打切区分
                Case 1:
                        If .打切種別 = 0 Then
                            gStr = gStr & "打切中 残( " & .打切残数 & " )"
                        Else
                            With CisFun
                                .DateE_BefVal = ZCT.打切日
                                .DateEdit Slash
                                gStr = gStr & "打切日( " & .DateE_AfVal & " )"
                            End With
                        End If
                Case 9: gStr = "打切完了 "
                Case Else: gStr = ""
            End Select
            B1_Bikou = gStr
        End With
        GPKensu = GPKensu + 1
        GTKensu = GTKensu + 1
    'BODY部印刷
        Call CisPrtCXHZ0330.PrintFld("B1*")
        CisPrtCXHZ0330.cLine = CisPrtCXHZ0330.cLine + 1
        Call CisPrtCXHZ0330.ClrLstFld("B1")
            
        '*---------------------------------*
        '*  カウンタ　ＵＰ
        '*---------------------------------*
        With CisPrtCXHZ0330
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(CXHZ0330_PrtCnt, "#,###")
        End With
        
        ' 材料発注行テーブル　発行日更新(新規発行のみ)
        If CXHZ0330_PrintKb = 28 Then
            gSL_Select = "UPDATE 材料発注テーブル SET "
            gSL_Select = gSL_Select & " 明細書発行日 = CONVERT(VARCHAR(8),GETDATE(),112)"
            gSL_Select = gSL_Select & " WHERE 発注管理NO = " & ZCT.発注管理NO & ""
            With CisDB
                .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
                On Error GoTo CoPrintHZ0330_Err
                .SQL = gSL_Select
                .DBExec
                .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
            End With
        End If
    'KEY項目退避
        Key_MeiNo = ZCT.明細書NO
        Key_Nouno = ZCT.納入番号
        Key_NPage = ZCT.納品書頁
        Key_Siire = ZCT.仕入先
        Key_SKKbn = ZCT.SK区分
        
        Call ZCTReadNext(1)
    Loop
'最終頁　枚数計印刷
    If Key_MeiNo <> 0 Then
       '頁合計印刷
        B1_K_LINE1.Visible = True
        If MeiPage > 1 Then
            B1_KensuM = "《頁件数》"
            B1_Kensu = Format(GPKensu, "#,##0")
            If Key_SKKbn = 0 Then
                B1_HakoKei = "《箱数頁合計》"
                B1_Maisu = Format(GPMaisu, "##,##0")
            End If
            Call CisPrtCXHZ0330.PrintFld("B1*")
            CisPrtCXHZ0330.cLine = CisPrtCXHZ0330.cLine + 1
            B1_K_LINE1.Visible = False
        End If
       '明細№合計印刷
        B1_K_LINE2.Visible = False
        Call CisPrtCXHZ0330.ClrLstFld("B1")
        B1_KensuM = "【総件数】"
        B1_Kensu = Format(GTKensu, "#,##0")
        If Key_SKKbn = 0 Then
            B1_HakoKei = "【箱数総合計】"
            B1_Maisu = Format(GTMaisu, "##,##0")
        End If
        Call CisPrtCXHZ0330.PrintFld("B1*")
    End If
    
    Call ZCTClose(1)
    
    ' 印刷終了
    Call CisPrtCXHZ0330.PrintEndDoc
    
    Exit Function
CoPrintHZ0330_Err:
    CisFun.ErrorBox
'    Call PrintSet2("", True)
    End
End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim NewPageFlg As Boolean
    HeadWrite = False
    
    With CisPrtCXHZ0330
        NewPageFlg = False
        .cLine = 0
        .cPage = .cPage + 1
        .ClrLstFld ("H1")   'HEAD部内容消去
    '頁
        MeiPage = MeiPage + 1
        If Key_Siire <> ZCT.仕入先 Then
            Call Torcd_PageGet
            SiirePage = 0
        End If
        SiirePage = SiirePage + 1
        H1_Page = "( " & Format(SiirePage, "##0") & " / " & Format(TotalPage, "##0") & " )"
    '仕入先・仕入先名
        H1_Siire = RTrim(ZCT.仕入先CD)
        If RTrim(ZCT.仕入先工場CD) <> "" Then
            H1_Siire = H1_Siire & " - " & RTrim(ZCT.仕入先工場CD)
        End If
        H1_SiireNm = RTrim(ZCT.仕入先名)
        H1_SiireKj = RTrim(ZCT.仕入先工場名)
    '直送先・手配商社
        For gInt = 0 To 3
            H1_K_NLINE(gInt).Visible = False
        Next
        If RTrim(ZCT.納入先) <> "" Then
            If RTrim(ZCT.手配区分) = "1" Then
                H1_NonyuM = "直送先"
            Else
                H1_NonyuM = "手配商社"
            End If
            H1_Nonyu = RTrim(ZCT.納入先CD)
            If RTrim(ZCT.納入先工場CD) <> "" Then
                H1_Nonyu = H1_Nonyu & " - " & RTrim(ZCT.納入先工場CD)
            End If
            H1_Nonyunm = RTrim(ZCT.納入先名)
            H1_Nonyukj = RTrim(ZCT.納入先工場名)
            H1_NOncyu = "御中"
            For gInt = 0 To 3
                H1_K_NLINE(gInt).Visible = True
            Next
        End If
    '帳票種別（かんばん・指示）
        If ZCT.SK区分 = 0 Then
            H1_SKKbn = "かんばん明細書"
        Else
            H1_SKKbn = "指示明細書"
        End If
    '発行日・自社名
        H1_HYmd = mPrtDate & " " & mPrtTime
        H1_JisyaNm = mJisyaNm
        H1_JisyaKj = mJisyaKj
    '発注区分・手配区分・納入日・便・受入
        H1_HCKbn = RTrim(ZCT.発注区分名)
        H1_TEKbn = RTrim(ZCT.手配区分名)
        H1_NYy = Left(ZCT.納入日, 4)
        H1_NMm = Mid(ZCT.納入日, 5, 2)
        H1_NDd = Mid(ZCT.納入日, 7, 2)
        If RTrim(ZCT.納入便) <> "" Then
            H1_NBin = RTrim(ZCT.納入便)
            H1_NBinM = "便"
        End If
        H1_Ukeir = RTrim(ZCT.受入)
'
        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With
'
    Key_Nouno = ""
    HeadWrite = True
End Function
Private Function Torcd_PageGet()
    Torcd_PageGet = False
'
    TotalPage = 0
    gSL_Select = "SELECT HC.明細書NO,COUNT(HC.材料管理番号)"
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料発注テーブル HC"
    gSL_Select = gSL_Select & "   ON HC.明細書NO = CH.管理NO"
    gSL_Select = gSL_Select & " WHERE CH.端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND CH.帳票区分 = " & CXHZ0330_PrintKb
    gSL_Select = gSL_Select & "   AND CH.管理文字 = '1'"
    gSL_Select = gSL_Select & "   AND HC.仕入先 = '" & RTrim(ZCT.仕入先) & "'"
    gSL_Select = gSL_Select & " GROUP BY HC.明細書NO"
    
    With CisDB
        .SQL = gSL_Select
        .ReadStatus = True
        If .DBRead(0, 0) Then
            Do Until .Recordset.EOF
                gLong = CisDB.Recordset(1)
                TotalPage = TotalPage + (gLong \ CisPrtCXHZ0330.iLineMax)
                If gLong Mod CisPrtCXHZ0330.iLineMax <> 0 Then
                    TotalPage = TotalPage + 1
                End If
                .Recordset.MoveNext
            Loop
        End If
        .Recordset.Close
    End With
'
    Torcd_PageGet = True
End Function
