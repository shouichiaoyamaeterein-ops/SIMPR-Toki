VERSION 5.00
Begin VB.Form COK0530L 
   Caption         =   "仕入明細書"
   ClientHeight    =   3330
   ClientLeft      =   2145
   ClientTop       =   900
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3330
   ScaleWidth      =   15420
   Begin VB.Label H1_KTornm 
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
      Left            =   4980
      TabIndex        =   36
      Top             =   915
      Width           =   3045
   End
   Begin VB.Label Label1 
      Caption         =   "契約先"
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
      Left            =   4320
      TabIndex        =   35
      Top             =   945
      Width           =   600
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
      Index           =   11
      Left            =   14910
      TabIndex        =   34
      Top             =   1335
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "外注購入"
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
      Left            =   13530
      TabIndex        =   33
      Top             =   1335
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "外注材料"
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
      Left            =   12135
      TabIndex        =   32
      Top             =   1335
      Width           =   840
   End
   Begin VB.Label H1_TorTitle2 
      AutoSize        =   -1  'True
      Caption         =   "仕  入  先"
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
      Left            =   8370
      TabIndex        =   31
      Top             =   1335
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "有償支給"
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
      Left            =   10785
      TabIndex        =   30
      Top             =   1335
      Width           =   840
   End
   Begin VB.Label B1_Kingak2 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "ZZZZZZZZ9.00"
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
      Left            =   14775
      TabIndex        =   29
      Top             =   1710
      Width           =   1260
   End
   Begin VB.Label B1_GKonyu2 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "ZZZZZZZZ9.00"
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
      Left            =   13395
      TabIndex        =   28
      Top             =   1710
      Width           =   1260
   End
   Begin VB.Label B1_GZairyo2 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "ZZZZZZZZ9.00"
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
      Left            =   12000
      TabIndex        =   27
      Top             =   1710
      Width           =   1260
   End
   Begin VB.Label B1_STOR2 
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
      Left            =   8370
      TabIndex        =   26
      Top             =   1710
      Width           =   1800
   End
   Begin VB.Label B1_Yusyo2 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "ZZZZZZZZ9.00"
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
      Left            =   10620
      TabIndex        =   25
      Top             =   1710
      Width           =   1260
   End
   Begin VB.Label B1_Line 
      AutoSize        =   -1  'True
      Caption         =   "｜"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7980
      TabIndex        =   24
      Top             =   1695
      Width           =   225
   End
   Begin VB.Label B1_Kingak1 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "ZZZZZZZZ9.00"
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
      Left            =   6600
      TabIndex        =   23
      Top             =   1710
      Width           =   1260
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
      Index           =   2
      Left            =   6750
      TabIndex        =   22
      Top             =   1335
      Width           =   840
   End
   Begin VB.Label B1_GKonyu1 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "ZZZZZZZZ9.00"
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
      Left            =   5220
      TabIndex        =   21
      Top             =   1710
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "外注購入"
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
      Left            =   5370
      TabIndex        =   20
      Top             =   1335
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "～"
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
      Index           =   6
      Left            =   2205
      TabIndex        =   19
      Top             =   570
      Width           =   240
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
      TabIndex        =   18
      Top             =   915
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
      TabIndex        =   17
      Top             =   945
      Width           =   600
   End
   Begin VB.Label B1_Yusyo1 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "ZZZZZZZZ9.00"
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
      Left            =   2445
      TabIndex        =   16
      Top             =   1710
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "有償支給"
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
      Left            =   2625
      TabIndex        =   15
      Top             =   1335
      Width           =   840
   End
   Begin VB.Label B1_STOR1 
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
      Left            =   195
      TabIndex        =   14
      Top             =   1710
      Width           =   1800
   End
   Begin VB.Label H1_TorTitle1 
      AutoSize        =   -1  'True
      Caption         =   "仕  入  先"
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
      Left            =   210
      TabIndex        =   13
      Top             =   1335
      Width           =   1050
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
      TabIndex        =   12
      Top             =   105
      Width           =   750
   End
   Begin VB.Label H1_Title 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "YYYY年MM月度　仕 入 明 細 書"
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
      TabIndex        =   11
      Top             =   105
      Width           =   5340
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "｜"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   7980
      TabIndex        =   10
      Top             =   1335
      Width           =   225
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   285
      Width           =   225
   End
   Begin VB.Label B1_GZairyo1 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "ZZZZZZZZ9.00"
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
      Left            =   3825
      TabIndex        =   7
      Top             =   1710
      Width           =   1260
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   1
      X1              =   165
      X2              =   16195
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "外注材料"
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
      Left            =   3975
      TabIndex        =   6
      Top             =   1335
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "仕入日"
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
      Top             =   585
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
   Begin VB.Label H1_EYMD 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD"
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
      Left            =   2580
      TabIndex        =   2
      Top             =   570
      Width           =   1200
   End
   Begin VB.Label H1_SYMD 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD"
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
      Top             =   570
      Width           =   1200
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
Attribute VB_Name = "COK0530L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  仕入明細書　発行
'**       フォームID    :  COK0530L
'**       処理概要      :
'**
'**       作  成  日    :  2009/02/19
'**       変  更  日    :
'**
'**********************************************************'
'   #比較退避用#
    Dim Key_KBN         As String       '>区分(名称)
    
    Dim M_YusyoG        As Currency         '>有償支給計(区分計)
    Dim M_ZairyoG       As Currency         '>外注材料計(区分計)
    Dim M_KounyuG       As Currency         '>外注購入計(区分計)
    Dim M_KingakG       As Currency         '>金額計(区分計)
    
    Dim M_YusyoG1       As Currency         '>有償支給計_1(中計)
    Dim M_ZairyoG1      As Currency         '>外注材料計_1(中計)
    Dim M_KounyuG1      As Currency         '>外注購入計_1(中計)
    Dim M_KingakG1      As Currency         '>金額計_1(中計)
    
    Dim M_YusyoG2       As Currency         '>有償支給計_2(中計)
    Dim M_ZairyoG2      As Currency         '>外注材料計_2(中計)
    Dim M_KounyuG2      As Currency         '>外注購入計_2(中計)
    Dim M_KingakG2      As Currency         '>金額計_2(中計)
    
    Dim A_YusyoG        As Currency         '>有償支給計(総合計)
    Dim A_ZairyoG       As Currency         '>外注材料計(総合計)
    Dim A_KounyuG       As Currency         '>外注購入計(総合計)
    Dim A_KingakG       As Currency         '>金額計(総合計)
    
    Dim MeisaiLR        As Byte         '0:左 1:右
    
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
    With CisPrtCOK0530
        .iPrtTitle = "仕入明細書"
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
    Key_KBN = RTrim(SMW.区分)
    
    M_YusyoG = 0
    M_ZairyoG = 0
    M_KounyuG = 0
    M_KingakG = 0
    
    M_YusyoG1 = 0
    M_ZairyoG1 = 0
    M_KounyuG1 = 0
    M_KingakG1 = 0
    
    M_YusyoG2 = 0
    M_ZairyoG2 = 0
    M_KounyuG2 = 0
    M_KingakG2 = 0
    
    A_YusyoG = 0
    A_ZairyoG = 0
    A_KounyuG = 0
    A_KingakG = 0
    
    Call CisPrtCOK0530.ClrLstFld("B1")

    Call Set_Line
    
    MeisaiLR = 0
    
'*---------------------------------*
'*  データ読込み
'*---------------------------------*
    Do Until Not SMW_RDSTS
        If RTrim(Key_KBN) <> SMW.区分 Then
        '   前回明細が左側で終了している場合は、右側空白にて明細印刷
           If MeisaiLR = 1 Then
              B1_STOR2 = ""
              B1_Yusyo2 = ""
              B1_GZairyo2 = ""
              B1_GKonyu2 = ""
              B1_Kingak2 = ""
              With CisPrtCOK0530
                   .PrintFld ("B1*")
                   .cLine = .cLine + 1
                   .ClrLstFld ("B1")
              End With
           End If
        
        '   合計印刷(中計)/区分計
            Call Total_Print(0)
            
            CisPrtCOK0530.cLine = 99
        
'        '   改頁判定
'            If RTrim(Key_Torcd) <> YUW.納入先CD Or _
'               RTrim(Key_Jisya) <> YUW.自社工区名 Then
'                CisPrtCOK0530.cLine = CisPrtCOK0530.iLineMax
'            End If
        End If

        Call HeadWrite(0)

        'CisPrtCOK0530.ClrLstFld ("B1")
        '*---------------------------------*
        '*  Body行　編集
        '*---------------------------------*
        With SMW
            
            B1_Line = "｜"
            
            If MeisaiLR = 0 Then
               B1_STOR1 = CisFun.Mid2(RTrim(.仕入先名), 1, 26)
               B1_Yusyo1 = Format(.有償支給計, "########0.00")
               B1_GZairyo1 = Format(.外注材料計, "########0.00")
               B1_GKonyu1 = Format(.外注購入計, "########0.00")
               B1_Kingak1 = Format(.金額, "########0.00")
            
               ' 合計計算(中計)
               M_YusyoG1 = M_YusyoG1 + .有償支給計
               M_ZairyoG1 = M_ZairyoG1 + .外注材料計
               M_KounyuG1 = M_KounyuG1 + .外注購入計
               M_KingakG1 = M_KingakG1 + .金額
               
               MeisaiLR = 1
            Else
               B1_STOR2 = CisFun.Mid2(RTrim(.仕入先名), 1, 26)
               B1_Yusyo2 = Format(.有償支給計, "########0.00")
               B1_GZairyo2 = Format(.外注材料計, "########0.00")
               B1_GKonyu2 = Format(.外注購入計, "########0.00")
               B1_Kingak2 = Format(.金額, "########0.00")
            
               ' 合計計算(中計)
               M_YusyoG2 = M_YusyoG2 + .有償支給計
               M_ZairyoG2 = M_ZairyoG2 + .外注材料計
               M_KounyuG2 = M_KounyuG2 + .外注購入計
               M_KingakG2 = M_KingakG2 + .金額
               
               MeisaiLR = 0
            End If
            
            ' 合計計算
            M_YusyoG = M_YusyoG + .有償支給計
            M_ZairyoG = M_ZairyoG + .外注材料計
            M_KounyuG = M_KounyuG + .外注購入計
            M_KingakG = M_KingakG + .金額
            
            A_YusyoG = A_YusyoG + .有償支給計
            A_ZairyoG = A_ZairyoG + .外注材料計
            A_KounyuG = A_KounyuG + .外注購入計
            A_KingakG = A_KingakG + .金額
            
            ' KEY項目退避
            Key_KBN = .区分
            
            '*---------------------------------*
            '*  Body部(B1)　印刷 (明細右側時）
            '*---------------------------------*
            If MeisaiLR = 0 Then
                With CisPrtCOK0530
                    .PrintFld ("B1*")
                    .cLine = .cLine + 1
                    .ClrLstFld ("B1")
                End With
            End If
        
        End With
        
        ' カウンタ　ＵＰ
        With CisPrtCOK0530
             .cDataCnt = .cDataCnt + 1
             .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With
                     
        Call SMWReadNext(1)
    Loop
    
    Call SMWClose(1)

'   合計印刷
    Call Total_Print(1)

'   印刷終了
    Call CisPrtCOK0530.PrintEndDoc

End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim TtlName     As String
    Dim YmdName     As String

    HeadWrite = False
    
    With CisPrtCOK0530
    '   １頁内印刷行数ﾁｪｯｸ
        If .iLineMax >= .cLine + 1 Then Exit Function

        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")   'HEAD部内容消去

    '   発行日・頁
        H1_HYmd = mPrtDate & " " & mPrtTime
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")
    '   タイトル
        H1_Title = Mid(RV_SYmd, 1, 4) & "年" & Mid(RV_SYmd, 5, 2) & "月度" & " 仕 入 明 細 書"
    '   仕入日
        H1_SYMD = Mid(RV_SYmd, 1, 4) & "/" & Mid(RV_SYmd, 5, 2) & "/" & Mid(RV_SYmd, 7, 2)
        H1_EYMD = Mid(RV_EYmd, 1, 4) & "/" & Mid(RV_EYmd, 5, 2) & "/" & Mid(RV_EYmd, 7, 2)
    '   自社工区
        H1_Kojyo = RTrim(COK0530_Kojyo)
    '   契約先
        If RTrim(COK0530_KTornm) <> "" Then
           H1_KTornm = COK0530_KTornm
        Else
           H1_KTornm = ""
        End If

    '   取引先見出し
        If COK0530_OutPut = 0 Then
           H1_TorTitle1 = "仕  入  先"
           H1_TorTitle2 = "仕  入  先"
        Else
           H1_TorTitle1 = "契　約　先"
           H1_TorTitle2 = "契　約　先"
        End If
        
        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With

    HeadWrite = True
End Function
'*---------------------------------------------------------------*
'*      合計印刷                                                *
'*---------------------------------------------------------------*
Private Function Total_Print(KB As Byte)
    Total_Print = False
'
    B1_Line = "｜"
            
    '左側
    B1_STOR1 = "【 中　計 】"
    B1_Yusyo1 = Format(M_YusyoG1, "########0.00")
    B1_GZairyo1 = Format(M_ZairyoG1, "########0.00")
    B1_GKonyu1 = Format(M_KounyuG1, "########0.00")
    B1_Kingak1 = Format(M_KingakG1, "########0.00")
    
    '右側
    B1_STOR2 = "【 中　    計 】"
    B1_Yusyo2 = Format(M_YusyoG2, "########0.00")
    B1_GZairyo2 = Format(M_ZairyoG2, "########0.00")
    B1_GKonyu2 = Format(M_KounyuG2, "########0.00")
    B1_Kingak2 = Format(M_KingakG2, "########0.00")
    
    Call CisPrtCOK0530.PrintFld("B1*")
    CisPrtCOK0530.cLine = CisPrtCOK0530.cLine + 1
    Call CisPrtCOK0530.ClrLstFld("B1")
    
    B1_Line = "｜"
            
    '区分計
    Select Case Key_KBN
           Case 1: B1_STOR2 = "【 外注加工費 】"
           Case 2: B1_STOR2 = "【 補助材料費 】"
           Case 3: B1_STOR2 = "【 有償仕入材 】"
           Case 4: B1_STOR2 = "【  一般材料  】"
    End Select
    B1_STOR1 = ""
    B1_Yusyo1 = ""
    B1_GZairyo1 = ""
    B1_GKonyu1 = ""
    B1_Kingak1 = ""
    B1_Yusyo2 = Format(M_YusyoG, "########0.00")
    B1_GZairyo2 = Format(M_ZairyoG, "########0.00")
    B1_GKonyu2 = Format(M_KounyuG, "########0.00")
    B1_Kingak2 = Format(M_KingakG, "########0.00")
    
    Call CisPrtCOK0530.PrintFld("B1*")
    CisPrtCOK0530.cLine = CisPrtCOK0530.cLine + 1
    Call CisPrtCOK0530.ClrLstFld("B1")
    
    
    M_YusyoG = 0
    M_ZairyoG = 0
    M_KounyuG = 0
    M_KingakG = 0
    
    M_YusyoG1 = 0
    M_ZairyoG1 = 0
    M_KounyuG1 = 0
    M_KingakG1 = 0
    
    M_YusyoG2 = 0
    M_ZairyoG2 = 0
    M_KounyuG2 = 0
    M_KingakG2 = 0
    
    MeisaiLR = 0
    
'    If KB = 1 Then
'       B1_Tanjyu = "< 得意先計 >"
'       B1_Kingak = Format(TK_GKingak, "###,###,##0")
'
'       Call CisPrtCOK0530.PrintFld("B1*")
'       Call CisPrtCOK0530.ClrLstFld("B1")
'
'       TK_GKingak = 0
'    End If
       
'
    '総合計印刷
    If KB = 1 Then
       B1_Line = "｜"
           
       B1_STOR2 = "【 総  合  計 】"
       B1_Yusyo2 = Format(A_YusyoG, "########0.00")
       B1_GZairyo2 = Format(A_ZairyoG, "########0.00")
       B1_GKonyu2 = Format(A_KounyuG, "########0.00")
       B1_Kingak2 = Format(A_KingakG, "########0.00")
       
       Call CisPrtCOK0530.PrintFld("B1*")
       CisPrtCOK0530.cLine = CisPrtCOK0530.cLine + 1
       Call CisPrtCOK0530.ClrLstFld("B1")
    
       A_YusyoG = 0
       A_ZairyoG = 0
       A_KounyuG = 0
       A_KingakG = 0
    
    End If
    
    Total_Print = True
End Function
Private Sub Set_Line()
    Dim wI As Integer

'    For wI = 1 To 7
'        B9_KDays(wI).Top = B1_Seban.Top - 30
'        B9_Suryo(wI).Top = B1_Seban.Top - 30
'    Next wI
    
'    B1_K_LINE.X1 = 120
'    B1_K_LINE.X2 = 16150
'    B1_K_LINE.Y1 = 1530
'    B1_K_LINE.Y2 = 1530
'

End Sub
