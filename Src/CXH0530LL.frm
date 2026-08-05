VERSION 5.00
Begin VB.Form CXH0530LL 
   Caption         =   "明細書"
   ClientHeight    =   11715
   ClientLeft      =   2550
   ClientTop       =   3510
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11715
   ScaleWidth      =   11700
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "受  領  書"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4590
      TabIndex        =   49
      Top             =   6270
      Width           =   2220
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  '点線
      BorderWidth     =   2
      X1              =   240
      X2              =   11340
      Y1              =   5790
      Y2              =   5790
   End
   Begin VB.Label H1_Bun 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXX"
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
      Left            =   8385
      TabIndex        =   48
      Top             =   1140
      Width           =   1080
   End
   Begin VB.Label B2_Kara 
      AutoSize        =   -1  'True
      Caption         =   "※※※　今回発注はありません　※※※"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1500
      TabIndex        =   47
      Top             =   3720
      Width           =   8640
   End
   Begin VB.Label H1_Hikae 
      Caption         =   "(控)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   26.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7440
      TabIndex        =   46
      Top             =   330
      Width           =   1155
   End
   Begin VB.Label B1_Kensu 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "9,999"
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
      Left            =   3930
      TabIndex        =   45
      Top             =   2100
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
      Left            =   2985
      TabIndex        =   44
      Top             =   2130
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
      Left            =   9885
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
      Left            =   10860
      TabIndex        =   42
      Top             =   1200
      Width           =   225
   End
   Begin VB.Line B1_K_LINE2 
      X1              =   1080
      X2              =   11175
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
      Left            =   5475
      TabIndex        =   41
      Top             =   2130
      Width           =   1365
   End
   Begin VB.Line H1_K_LINE 
      Index           =   5
      X1              =   3285
      X2              =   3285
      Y1              =   495
      Y2              =   1065
   End
   Begin VB.Line H1_K_LINE 
      Index           =   4
      X1              =   180
      X2              =   180
      Y1              =   495
      Y2              =   1065
   End
   Begin VB.Line H1_K_LINE 
      Index           =   3
      X1              =   180
      X2              =   3285
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Line H1_K_LINE 
      Index           =   2
      X1              =   180
      X2              =   3285
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Label B1_Hasu 
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
      Left            =   7605
      TabIndex        =   40
      Top             =   2100
      Width           =   720
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
      Left            =   7905
      TabIndex        =   39
      Top             =   1755
      Width           =   435
   End
   Begin VB.Line B1_K_LINE1 
      X1              =   195
      X2              =   11160
      Y1              =   2085
      Y2              =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "納入番号"
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
      Width           =   870
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
      Height          =   375
      Left            =   7845
      TabIndex        =   37
      Top             =   840
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
      Left            =   9270
      TabIndex        =   36
      Top             =   2100
      Width           =   1800
   End
   Begin VB.Label B1_HSuryo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999,999"
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
      TabIndex        =   35
      Top             =   2100
      Width           =   840
   End
   Begin VB.Label B1_Maisu 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "99,999"
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
      Left            =   6870
      TabIndex        =   34
      Top             =   2100
      Width           =   720
   End
   Begin VB.Label B1_Syuyo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999,999"
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
      Left            =   5925
      TabIndex        =   33
      Top             =   2100
      Width           =   840
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
      Left            =   4935
      TabIndex        =   32
      Top             =   2100
      Width           =   960
   End
   Begin VB.Label B1_Hinbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
      Left            =   1275
      TabIndex        =   31
      Top             =   2100
      Width           =   3600
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
      Left            =   975
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
      Left            =   330
      TabIndex        =   29
      Top             =   2100
      Width           =   600
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   1
      X1              =   195
      X2              =   11160
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
      Left            =   9555
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
      Left            =   8565
      TabIndex        =   27
      Top             =   1755
      Width           =   630
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
      Left            =   7185
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
      Left            =   6165
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
      Left            =   5040
      TabIndex        =   24
      Top             =   1755
      Width           =   630
   End
   Begin VB.Label Label4 
      Caption         =   "品    番"
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
      Left            =   1710
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
      Left            =   975
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
      Top             =   585
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
      Width           =   675
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   0
      X1              =   195
      X2              =   11160
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
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6150
      TabIndex        =   15
      Top             =   1155
      Width           =   330
   End
   Begin VB.Label H1_NDd 
      AutoSize        =   -1  'True
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5340
      TabIndex        =   14
      Top             =   1155
      Width           =   330
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
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4635
      TabIndex        =   12
      Top             =   1155
      Width           =   330
   End
   Begin VB.Label H1_NYy 
      AutoSize        =   -1  'True
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Top             =   1155
      Width           =   660
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
      Top             =   1155
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
      Top             =   1155
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
      Top             =   285
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
      Left            =   8760
      TabIndex        =   7
      Top             =   285
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
      Left            =   9390
      TabIndex        =   6
      Top             =   285
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
      Left            =   8670
      TabIndex        =   5
      Top             =   825
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
      Top             =   810
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
      Top             =   540
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
      Top             =   300
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
      Left            =   8670
      TabIndex        =   1
      Top             =   570
      Width           =   2400
   End
   Begin VB.Label H1_SKKbn 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "納  品  書"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4440
      TabIndex        =   0
      Top             =   315
      Width           =   2220
   End
End
Attribute VB_Name = "CXH0530LL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  発注明細書　発行
'**       フォームID    :  CXH0530L
'**       処理概要      :
'**
'**       作  成  日    :  2008/08/20
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

Private Sub B2_Kara_Click()

End Sub

'*------------------------------------*
'*      Ｆｏｒｍ＿Ｌｏａｄ イベント     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")             '発行日
    mPrtTime = Format(Now(), "hh:nn")                  '発行時間
    
    '*---------------------*
    '*  プリント系初期設定
    '*---------------------*
    With CisPrtCXH0530
        .iPrtTitle = "発注明細書"
        .iLineMax = 55
        .iLinePitch = 6
        .iPaperSize = p_A4
'        .iOrientation = po_PORTRAIT
        
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
    Key_Siire = "": GTKensu = 0: GTMaisu = 0 ' 2004/05/10 Add
    Call CisPrtCXH0530.ClrLstFld("B1")
    
'*---------------------------------*
'*  発注テーブル読込み
'*---------------------------------*
    Do Until Not HCT_RDSTS
        If Key_MeiNo <> HCT.明細書NO Or CisPrtCXH0530.cLine >= CisPrtCXH0530.iLineMax Then
            If Key_MeiNo <> 0 Then
                If Key_MeiNo <> HCT.明細書NO Then
                    B1_K_LINE1.Visible = True
                    If MeiPage > 1 Then
                        B1_KensuM = "《頁件数》"
                        B1_Kensu = Format(GPKensu, "#,##0")
                        If Key_SKKbn = 0 Then
                            B1_HakoKei = "《箱数頁合計》"
                            B1_Maisu = Format(GPMaisu, "##,##0")
                        End If
                        If Not CXH0530_Kara Then
                           Call CisPrtCXH0530.PrintFld("B1*")
                        End If
                        CisPrtCXH0530.cLine = CisPrtCXH0530.cLine + 1
                        Call CisPrtCXH0530.ClrLstFld("B1")
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
                    If Not CXH0530_Kara Then
                       Call CisPrtCXH0530.PrintFld("B1*")
                    End If
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
                    If Not CXH0530_Kara Then
                       Call CisPrtCXH0530.PrintFld("B1*")
                    End If
                    CisPrtCXH0530.cLine = CisPrtCXH0530.cLine + 1
                End If
                GPMaisu = 0: GPKensu = 0
            End If
            Call HeadWrite(0)
        End If
        B1_K_LINE1.Visible = False: B1_K_LINE2.Visible = False
        B1_HakoKei = "": B1_KensuM = "": B1_Kensu = ""
        
        If Not CXH0530_Kara Then
            With HCT
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
                B1_Hinbn = RTrim(.表示品番)
                B1_Seban = RTrim(.背番号)
                If .枚数 <> 0 Then
                    B1_Syuyo = Format(.収容数, "##,##0")
                End If
                If .枚数 <> 0 Then
                    B1_Maisu = Format(.枚数, "##,##0")
                    GPMaisu = GPMaisu + .枚数
                    GTMaisu = GTMaisu + .枚数
                End If
                If .端数 <> 0 Then
                    B1_Hasu = Format(.端数, "##,##0")
                End If
                If .発注数 <> 0 Then
                    B1_HSuryo = Format(.発注数, "##,##0")
                End If
               '備考（打切情報）
                gStr = ""
                Select Case .打切区分
                    Case 1:
                            If .打切種別 = 0 Then
                                gStr = gStr & "打切中 残( " & .打切残数 & " )"
                            Else
                                With CisFun
                                    .DateE_BefVal = HCT.打切日
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
            Call CisPrtCXH0530.PrintFld("B1*")
            CisPrtCXH0530.cLine = CisPrtCXH0530.cLine + 1
            Call CisPrtCXH0530.ClrLstFld("B1")
        Else
            B2_Kara = "※※※　今回発注はありません　※※※"
            Call CisPrtCXH0530.PrintFld("B2*")
            CisPrtCXH0530.cLine = CisPrtCXH0530.cLine + 1
            Call CisPrtCXH0530.ClrLstFld("B2")
        End If
        '*---------------------------------*
        '*  カウンタ　ＵＰ
        '*---------------------------------*
        With CisPrtCXH0530
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(CXH0530_PrtCnt, "#,###")
        End With
        
        If Not CXH0530_Kara Then
            ' 発注行テーブル　発行日更新(新規発行のみ)
            If CXH0530_PrintKb = 8 Then
                gSL_Select = "UPDATE 発注テーブル SET "
                gSL_Select = gSL_Select & " 明細書発行日 = CONVERT(VARCHAR(8),GETDATE(),112)"
                gSL_Select = gSL_Select & " WHERE 発注管理NO = " & HCT.発注管理NO & ""
                With CisDB
                    .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
                    On Error GoTo CoPrintH0530_Err
                    .SQL = gSL_Select
                    .DBExec
                    .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
                End With
            End If
        Else
            ' 空明細書テーブル　発行日更新(新規発行のみ)
            If CXH0530_PrintKb2 = 15 Then
                gSL_Select = "UPDATE 空明細書テーブル SET "
                gSL_Select = gSL_Select & " 明細書発行日 = CONVERT(VARCHAR(8),GETDATE(),112)"
                gSL_Select = gSL_Select & " WHERE 明細書NO = " & HCT.明細書NO & ""
                With CisDB
                    .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
                    On Error GoTo CoPrintH0530_Err
                    .SQL = gSL_Select
                    .DBExec
                    .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
                End With
            End If
        End If
    'KEY項目退避
        Key_MeiNo = HCT.明細書NO
        Key_Nouno = HCT.納入番号
        Key_NPage = HCT.納品書頁
        Key_Siire = HCT.仕入先
        Key_SKKbn = HCT.SK区分
        
        Call HCTReadNext(1)
    Loop
'最終頁　枚数計印刷
    If Not CXH0530_Kara Then
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
                Call CisPrtCXH0530.PrintFld("B1*")
                CisPrtCXH0530.cLine = CisPrtCXH0530.cLine + 1
                B1_K_LINE1.Visible = False
            End If
           '明細№合計印刷
            B1_K_LINE2.Visible = False
            Call CisPrtCXH0530.ClrLstFld("B1")
            B1_KensuM = "【総件数】"
            B1_Kensu = Format(GTKensu, "#,##0")
            If Key_SKKbn = 0 Then
                B1_HakoKei = "【箱数総合計】"
                B1_Maisu = Format(GTMaisu, "##,##0")
            End If
            Call CisPrtCXH0530.PrintFld("B1*")
        End If
    End If
    
    Call HCTClose(1)
    
    ' 印刷終了
    Call CisPrtCXH0530.PrintEndDoc
    
    Exit Function
CoPrintH0530_Err:
    CisFun.ErrorBox
    End
End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim NewPageFlg As Boolean
    HeadWrite = False
    
    With CisPrtCXH0530
        NewPageFlg = False
        .cLine = 0
        .cPage = .cPage + 1
        .ClrLstFld ("H1")   'HEAD部内容消去
    '頁
        MeiPage = MeiPage + 1
        If Key_Siire <> HCT.仕入先 Then
            Call Torcd_PageGet
            SiirePage = 0
        End If
        SiirePage = SiirePage + 1
        H1_Page = "( " & Format(SiirePage, "##0") & " / " & Format(TotalPage, "##0") & " )"
    '仕入先・仕入先名
        H1_Siire = RTrim(HCT.仕入先CD)
        If RTrim(HCT.仕入先工場CD) <> "" Then
            H1_Siire = H1_Siire & " - " & RTrim(HCT.仕入先工場CD)
        End If
        H1_SiireNm = RTrim(HCT.仕入先名)
        H1_SiireKj = RTrim(HCT.仕入先工場名)
    '帳票種別（かんばん・指示）
        If HCT.SK区分 = 0 Then
            H1_SKKbn = "かんばん明細書"
        Else
            H1_SKKbn = "指示明細書"
        End If
    '発行日・自社名
        H1_HYmd = mPrtDate & " " & mPrtTime
        H1_JisyaNm = mJisyaNm
        H1_JisyaKj = mJisyaKj
        
        If CXH0530_Hikae Then
           H1_Hikae = "(控)"
        End If
    '発注区分・手配区分・納入日・便・受入
        H1_HCKbn = RTrim(HCT.発注区分名)
        H1_TEKbn = RTrim(HCT.手配区分名)
        H1_NYy = Left(HCT.納入日, 4)
        H1_NMm = Mid(HCT.納入日, 5, 2)
        H1_NDd = Mid(HCT.納入日, 7, 2)
        If RTrim(HCT.納入便) <> "" Then
            H1_NBin = RTrim(HCT.納入便)
            H1_NBinM = "便"
        End If
        H1_Ukeir = RTrim(HCT.受入)
    ' 分割
        If Trim(HCT.納入番号) <> "" Then
            If Right(HCT.納入番号, 1) <> "0" Then
               H1_Bun = "【分割】"
            End If
        End If
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
    If Not CXH0530_Kara Then
        gSL_Select = "SELECT HC.明細書NO,COUNT(HC.品番)"
        gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 発注テーブル HC"
        gSL_Select = gSL_Select & "   ON HC.明細書NO = CH.管理NO"
        gSL_Select = gSL_Select & " WHERE CH.端末番号 = HOST_NAME()"
        gSL_Select = gSL_Select & "   AND CH.帳票区分 = " & CXH0530_PrintKb
    '    gSL_Select = gSL_Select & "   AND CH.管理文字 = '" & CXH0530_PrtOut & "'"
        If CXH0530_Hikae Then
           gSL_Select = gSL_Select & "   AND CH.管理文字 IN ('1','3','4')"
        Else
           If Not CXH0530_FaxProc Then
              gSL_Select = gSL_Select & "   AND CH.管理文字 IN ('1','3')"
           Else
              gSL_Select = gSL_Select & "   AND CH.管理文字 = '" & CXH0530_PrtOut & "'"
           End If
'           gSL_Select = gSL_Select & "   AND CH.管理文字 = '" & CXH0530_PrtOut & "'"
        End If
        
        gSL_Select = gSL_Select & "   AND HC.仕入先 = '" & RTrim(HCT.仕入先) & "'"
        gSL_Select = gSL_Select & " GROUP BY HC.明細書NO"
        
        With CisDB
            .SQL = gSL_Select
            .ReadStatus = True
            If .DBRead(0, 0) Then
                Do Until .Recordset.EOF
                    gLong = CisDB.Recordset(1)
                    TotalPage = TotalPage + (gLong \ CisPrtCXH0530.iLineMax)
                    If gLong Mod CisPrtCXH0530.iLineMax <> 0 Then
                        TotalPage = TotalPage + 1
                    End If
                    .Recordset.MoveNext
                Loop
            End If
            .Recordset.Close
        End With
    '
        Torcd_PageGet = True
        Exit Function
    Else
        TotalPage = 1
        gSL_Select = "SELECT HC.仕入先,Count(*)"
        gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 空明細書テーブル HC"
        gSL_Select = gSL_Select & "   ON HC.明細書NO = CH.管理NO"
        gSL_Select = gSL_Select & " WHERE CH.端末番号 = HOST_NAME()"
        gSL_Select = gSL_Select & "   AND CH.帳票区分 = " & CXH0530_PrintKb2
        If CXH0530_Hikae Then
           gSL_Select = gSL_Select & "   AND CH.管理文字 IN ('1','3','4')"
        Else
'           gSL_Select = gSL_Select & "   AND CH.管理文字 = '" & CXH0530_PrtOut & "'"
           If Not CXH0530_FaxProc Then
              gSL_Select = gSL_Select & "   AND CH.管理文字 IN ('1','3')"
           Else
              gSL_Select = gSL_Select & "   AND CH.管理文字 = '" & CXH0530_PrtOut & "'"
           End If
        End If

        gSL_Select = gSL_Select & "   AND HC.仕入先 = '" & RTrim(HCT.仕入先) & "'"
        gSL_Select = gSL_Select & " Group By HC.仕入先"
'
        With CisDB
            .SQL = gSL_Select
            .ReadStatus = True
            If .DBRead(0, 0) Then
                gLong = CisDB.Recordset(1)
                TotalPage = gLong
            End If
            .Recordset.Close
        End With
    '
        Torcd_PageGet = True
        Exit Function
    End If
End Function

