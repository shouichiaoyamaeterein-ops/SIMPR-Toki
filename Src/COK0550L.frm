VERSION 5.00
Begin VB.Form COK0550L 
   Caption         =   "有償売上明細表"
   ClientHeight    =   3330
   ClientLeft      =   1410
   ClientTop       =   2325
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3330
   ScaleWidth      =   15420
   Begin VB.Line B1_K_LINE 
      BorderStyle     =   3  '点線
      X1              =   120
      X2              =   16150
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label H1_Kojyo 
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
      Left            =   855
      TabIndex        =   43
      Top             =   930
      Width           =   3045
   End
   Begin VB.Label Label1 
      Caption         =   "工　区"
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
      Index           =   6
      Left            =   225
      TabIndex        =   42
      Top             =   945
      Width           =   600
   End
   Begin VB.Label B9_KDays 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD"
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
      Left            =   14400
      TabIndex        =   41
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label B9_Suryo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Left            =   15090
      TabIndex        =   40
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label B9_KDays 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD"
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
      Left            =   12510
      TabIndex        =   39
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label B9_Suryo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Left            =   13200
      TabIndex        =   38
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label B9_KDays 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD"
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
      Left            =   10560
      TabIndex        =   37
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label B9_Suryo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Left            =   11250
      TabIndex        =   36
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label B9_KDays 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD"
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
      Left            =   8670
      TabIndex        =   35
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label B9_Suryo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Left            =   9360
      TabIndex        =   34
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label B9_KDays 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD"
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
      Left            =   6720
      TabIndex        =   33
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label B9_Suryo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Left            =   7410
      TabIndex        =   32
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label B9_KDays 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD"
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
      Index           =   2
      Left            =   4845
      TabIndex        =   31
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label B9_Suryo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Index           =   2
      Left            =   5535
      TabIndex        =   30
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label B9_KDays 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD"
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
      Index           =   1
      Left            =   2895
      TabIndex        =   29
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label B9_Suryo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Index           =   1
      Left            =   3585
      TabIndex        =   28
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label B1_Tanjyu 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "###,##0.000"
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
      Left            =   11460
      TabIndex        =   27
      Top             =   1590
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "単　　重"
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
      Left            =   12015
      TabIndex        =   26
      Top             =   1230
      Width           =   840
   End
   Begin VB.Label B1_Tanka 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
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
      Left            =   9975
      TabIndex        =   25
      Top             =   1590
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "単　　価"
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
      Left            =   10410
      TabIndex        =   24
      Top             =   1230
      Width           =   840
   End
   Begin VB.Label B1_Seban 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXX"
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
      Left            =   2895
      TabIndex        =   23
      Top             =   1590
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Index           =   3
      Left            =   3015
      TabIndex        =   22
      Top             =   1230
      Width           =   630
   End
   Begin VB.Label B1_Hinbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXX*"
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
      Left            =   150
      TabIndex        =   21
      Top             =   1590
      Width           =   2625
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "品　　番"
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
      Index           =   0
      Left            =   825
      TabIndex        =   20
      Top             =   1230
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "＞＞"
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
      Index           =   5
      Left            =   11070
      TabIndex        =   19
      Top             =   105
      Width           =   750
   End
   Begin VB.Label H1_Title 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "YYYY年MM月度　売 上 明 細 書"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5145
      TabIndex        =   18
      Top             =   105
      Width           =   5340
   End
   Begin VB.Label B1_Tannm 
      AutoSize        =   -1  'True
      Caption         =   "XX"
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
      Left            =   9420
      TabIndex        =   17
      Top             =   1575
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "単位"
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
      Left            =   9330
      TabIndex        =   16
      Top             =   1230
      Width           =   390
   End
   Begin VB.Label H1_Page 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999"
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
      Left            =   15555
      TabIndex        =   15
      Top             =   270
      Width           =   315
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
      Index           =   3
      Left            =   15915
      TabIndex        =   14
      Top             =   285
      Width           =   225
   End
   Begin VB.Label B1_Suryo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0"
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
      Left            =   13485
      TabIndex        =   13
      Top             =   1590
      Width           =   945
   End
   Begin VB.Label B1_Kingak 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   14895
      TabIndex        =   12
      Top             =   1575
      Width           =   1155
   End
   Begin VB.Label B1_Sunpo 
      AutoSize        =   -1  'True
      Caption         =   "9999.99x9999.99x9999.99"
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
      Left            =   6660
      TabIndex        =   11
      Top             =   1590
      Width           =   2415
   End
   Begin VB.Label B1_Zaisitu 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXX*"
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
      Left            =   3840
      TabIndex        =   10
      Top             =   1590
      Width           =   2625
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   16150
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "数　　量"
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
      Left            =   13605
      TabIndex        =   9
      Top             =   1230
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "金　　額"
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
      Index           =   9
      Left            =   15195
      TabIndex        =   8
      Top             =   1230
      Width           =   840
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "寸　　　法"
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
      Index           =   2
      Left            =   7560
      TabIndex        =   7
      Top             =   1230
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "材    質"
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
      Index           =   1
      Left            =   4470
      TabIndex        =   6
      Top             =   1230
      Width           =   810
   End
   Begin VB.Label Label1 
      Caption         =   "得意先"
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
      Index           =   4
      Left            =   225
      TabIndex        =   5
      Top             =   615
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "作成日"
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
      Index           =   2
      Left            =   12675
      TabIndex        =   4
      Top             =   285
      Width           =   540
   End
   Begin VB.Label H1_HYmd 
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99"
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
      Left            =   13305
      TabIndex        =   3
      Top             =   270
      Width           =   1920
   End
   Begin VB.Label H1_Tornm 
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
      Left            =   1890
      TabIndex        =   2
      Top             =   600
      Width           =   3045
   End
   Begin VB.Label H1_Torcd 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX-XX"
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
      Left            =   870
      TabIndex        =   1
      Top             =   600
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "＜＜　"
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
      Index           =   1
      Left            =   3705
      TabIndex        =   0
      Top             =   90
      Width           =   960
   End
End
Attribute VB_Name = "COK0550L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  有償売上明細書　発行
'**       フォームID    :  COK0550L
'**       処理概要      :
'**
'**       作  成  日    :  2008/07/02
'**       変  更  日    :  2009/02/10
'**
'**********************************************************'
'   #比較退避用#
    Dim Key_Torcd       As String       '>得意先
    Dim Key_Jisya       As String       '>自社工区
    
    Dim Key_Hinbn       As String       '>品番
    Dim Key_Seban       As String       '>背番号
    Dim Key_Zaisitu     As String       '>材質
    Dim Key_Sunpou      As String       '>寸法
    Dim Key_Tani        As String       '>単位
    Dim Key_Tanka       As Currency     '>単価
    Dim Key_Tanjyu      As Currency     '>単重(実績換算値)
    
    Dim KJ_GKingak       As Long         '>金額(工場計)
    Dim TK_GKingak       As Long         '>金額(得意先計)
    
    Dim mPrtDate        As String
    Dim mPrtTime        As String

'*------------------------------------*
'*      Ｆｏｒｍ＿Ｌｏａｄ イベント     *
'*------------------------------------*
Private Sub form_load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")             '発行日
    mPrtTime = Format(Now(), "hh:nn")                  '発行時間
    '*---------------------*
    '*  プリント系初期設定
    '*---------------------*
    With CisPrtCOK0550
        .iPrtTitle = "有償売上明細書"
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
    Dim wI      As Integer

'*---------------------------------*
'*  初期設定
'*---------------------------------*
    Key_Torcd = RTrim(YUW.納入先CD)
    Key_Jisya = RTrim(YUW.自社工区名)
    
    KJ_GKingak = 0
    TK_GKingak = 0
    
    Call CisPrtCOK0550.ClrLstFld("B1")
    Call CisPrtCOK0550.ClrLstFld("B9")

    Key_Hinbn = RTrim(YUW.品番)
    Key_Seban = RTrim(YUW.背番号)
    Key_Zaisitu = RTrim(YUW.材質)
    Key_Sunpou = RTrim(YUW.表示寸法)
    Key_Tani = RTrim(YUW.手配単位)
    Key_Tanka = YUW.単価
    Key_Tanjyu = YUW.単重
    
    Call Set_Line
    
'*---------------------------------*
'*  データ読込み
'*---------------------------------*
    Do Until Not YUW_RDSTS
        If RTrim(Key_Torcd) <> YUW.納入先CD Or _
           RTrim(Key_Jisya) <> YUW.自社工区名 Then
        '   合計印刷
            If RTrim(Key_Torcd) = YUW.納入先CD Then
               Call Total_Print(0)
            Else
               Call Total_Print(1)
            End If
        '   改頁判定
            If RTrim(Key_Torcd) <> YUW.納入先CD Or _
               RTrim(Key_Jisya) <> YUW.自社工区名 Then
                CisPrtCOK0550.cLine = CisPrtCOK0550.iLineMax
            End If
        End If

        Call HeadWrite(0)

        CisPrtCOK0550.ClrLstFld ("B1")
        '*---------------------------------*
        '*  Body行　編集
        '*---------------------------------*
        With YUW
            
            B1_Hinbn = RTrim(.品番)
            B1_Seban = RTrim(.背番号)
            B1_Zaisitu = RTrim(.材質)
            B1_Sunpo = RTrim(.表示寸法)
            B1_Tannm = RTrim(.手配単位)
            B1_Tanka = Format(.単価, "#,###,##0.00")
            If .単重 <> 0 Then
               B1_Tanjyu = Format(.単重, "###,##0.000")
            Else
               B1_Tanjyu = ""
            End If
            B1_Suryo = Format(.納入数, "#,###,##0")
            B1_Kingak = Format(.金額, "###,###,##0")
            
            ' 合計計算
            KJ_GKingak = KJ_GKingak + .金額    '--工場単位
            TK_GKingak = TK_GKingak + .金額    '--得意先単位
            ' KEY項目退避
            Key_Torcd = .納入先CD
            Key_Jisya = .自社工区名
            
            '*---------------------------------*
            '*  Body部(B1)　印刷
            '*---------------------------------*
            With CisPrtCOK0550
                .PrintFld ("B1*")
                .cLine = .cLine + 1
                .ClrLstFld ("B1")
            End With

            If .明細件数 <> 0 Then
               gSL_Select = "SELECT ISNULL(納入日,'') 入荷日,ISNULL(納入数,0) 入荷数,ISNULL(金額,0) FROM 有償売上書ワーク"
               gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(.品番) & "'"
               gSL_Select = gSL_Select & "   AND 背番号 = '" & RTrim(.背番号) & "'"
               gSL_Select = gSL_Select & "   AND 材質 = '" & RTrim(.材質) & "'"
               gSL_Select = gSL_Select & "   AND 表示寸法 = '" & RTrim(.表示寸法) & "'"
               gSL_Select = gSL_Select & "   AND 手配単位 = '" & RTrim(.手配単位) & "'"
               gSL_Select = gSL_Select & "   AND 単価 = " & .単価
               gSL_Select = gSL_Select & "   AND 単重 = " & .単重
               gSL_Select = gSL_Select & "   AND 処理端末 = SUBSTRING(HOST_NAME(), 1, 20)"          '2009/02/10
               gSL_Select = gSL_Select & " ORDER BY 納入日"
               If ZCTRead(gSL_Select) Then
                  wI = 0
                  Do Until Not ZCT_RDSTS
                     wI = wI + 1
                     If wI <= 7 Then
                        B9_KDays(wI) = Mid(ZCT.入荷日, 5, 2) & "/" & Mid(ZCT.入荷日, 7, 2)
                        B9_Suryo(wI) = Format(ZCT.入荷数, "#,###,##0")
                     Else
                        '*---------------------------------*
                        '*  Body部(B9)　印刷
                        '*---------------------------------*
                        With CisPrtCOK0550
                            .PrintFld ("B9*")
                            .cLine = .cLine + 1
                            .ClrLstFld ("B9")
                        End With
                     
                        wI = 1
                        B9_KDays(wI) = Mid(ZCT.入荷日, 5, 2) & "/" & Mid(ZCT.入荷日, 7, 2)
                        B9_Suryo(wI) = Format(ZCT.入荷数, "#,###,##0")
                     End If
                     
                     Call ZCTReadNext
                  Loop
                  '*---------------------------------*
                  '*  Body部(B9)　印刷
                  '*---------------------------------*
                  With CisPrtCOK0550
                      .PrintFld ("B9*")
                      .cLine = .cLine + 1
                      .ClrLstFld ("B9")
                  End With
                  
                  Call ZCTClose
               End If
            End If
        
        End With
        
        ' カウンタ　ＵＰ
        With CisPrtCOK0550
             .cDataCnt = .cDataCnt + 1
             .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With
                     
        Call YUWReadNext(1)
    Loop
    Call YUWClose(1)

'   合計印刷
    Call Total_Print(1)
'   印刷終了
    Call CisPrtCOK0550.PrintEndDoc

End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim TtlName     As String
    Dim YmdName     As String

    HeadWrite = False
    
    With CisPrtCOK0550
    '   １頁内印刷行数ﾁｪｯｸ
        If .iLineMax >= .cLine + 1 Then Exit Function

        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")   'HEAD部内容消去

    '   発行日・頁
        H1_HYmd = mPrtDate & " " & mPrtTime
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")
    '   タイトル
        H1_Title = Mid(RV_SYmd, 1, 4) & "年" & Mid(RV_SYmd, 5, 2) & "月度" & " 有 償 売 上 明 細 書"
    '   得意先
        H1_Torcd = RTrim(YUW.納入先CD)
        H1_Tornm = RTrim(YUW.納入先名)
        H1_Kojyo = RTrim(YUW.自社工区名)

        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With

    HeadWrite = True
End Function
Private Function Total_Print(KB As Byte)
    Total_Print = False
'
    B1_Tanjyu = "< 工区計 >"
    B1_Kingak = Format(KJ_GKingak, "###,###,##0")
    
    Call CisPrtCOK0550.PrintFld("B1*")
    CisPrtCOK0550.cLine = CisPrtCOK0550.cLine + 2
    Call CisPrtCOK0550.ClrLstFld("B1")
    
    KJ_GKingak = 0
    
    If KB = 1 Then
       B1_Tanjyu = "< 得意先計 >"
       B1_Kingak = Format(TK_GKingak, "###,###,##0")
       
       Call CisPrtCOK0550.PrintFld("B1*")
       Call CisPrtCOK0550.ClrLstFld("B1")
    
       TK_GKingak = 0
    End If
       
'
    Total_Print = True
End Function
Private Sub Set_Line()
    Dim wI As Integer

    For wI = 1 To 7
        B9_KDays(wI).Top = B1_Seban.Top - 30
        B9_Suryo(wI).Top = B1_Seban.Top - 30
    Next wI
    
    B1_K_LINE.X1 = 120
    B1_K_LINE.X2 = 16150
    B1_K_LINE.Y1 = 1530
    B1_K_LINE.Y2 = 1530
    

End Sub
