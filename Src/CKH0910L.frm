VERSION 5.00
Begin VB.Form CKH0910L 
   Caption         =   "かんばん読取エラーリスト"
   ClientHeight    =   4770
   ClientLeft      =   900
   ClientTop       =   2580
   ClientWidth     =   18795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4770
   ScaleWidth      =   18795
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "ｻｲｸﾙ"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11100
      TabIndex        =   34
      Top             =   960
      Width           =   420
   End
   Begin VB.Label B1_Cycle 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "9-99-99"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10965
      TabIndex        =   33
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label B1_NO 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   11775
      TabIndex        =   32
      Top             =   1470
      Width           =   315
   End
   Begin VB.Label B1_NO 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   12180
      TabIndex        =   31
      Top             =   1470
      Width           =   315
   End
   Begin VB.Label B1_NO 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   12570
      TabIndex        =   30
      Top             =   1470
      Width           =   315
   End
   Begin VB.Label B1_NO 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   12960
      TabIndex        =   29
      Top             =   1470
      Width           =   315
   End
   Begin VB.Label B1_NO 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   13350
      TabIndex        =   28
      Top             =   1470
      Width           =   315
   End
   Begin VB.Label B1_NO 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   15300
      TabIndex        =   27
      Top             =   1470
      Width           =   315
   End
   Begin VB.Label B1_NO 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   13740
      TabIndex        =   26
      Top             =   1470
      Width           =   315
   End
   Begin VB.Label B1_NO 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   14130
      TabIndex        =   25
      Top             =   1470
      Width           =   315
   End
   Begin VB.Label B1_NO 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   14520
      TabIndex        =   24
      Top             =   1470
      Width           =   315
   End
   Begin VB.Label B1_NO 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   14910
      TabIndex        =   23
      Top             =   1470
      Width           =   315
   End
   Begin VB.Line B1_Line 
      X1              =   11715
      X2              =   16350
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Line B1_K_Line 
      Index           =   1
      X1              =   12120
      X2              =   12120
      Y1              =   1410
      Y2              =   1665
   End
   Begin VB.Line B1_K_Line 
      Index           =   0
      X1              =   11715
      X2              =   11715
      Y1              =   1395
      Y2              =   1665
   End
   Begin VB.Line B1_K_Line 
      Index           =   2
      X1              =   12510
      X2              =   12510
      Y1              =   1410
      Y2              =   1665
   End
   Begin VB.Line B1_K_Line 
      Index           =   3
      X1              =   12900
      X2              =   12900
      Y1              =   1410
      Y2              =   1665
   End
   Begin VB.Line B1_K_Line 
      Index           =   4
      X1              =   13305
      X2              =   13305
      Y1              =   1410
      Y2              =   1665
   End
   Begin VB.Line B1_K_Line 
      Index           =   5
      X1              =   13695
      X2              =   13695
      Y1              =   1410
      Y2              =   1665
   End
   Begin VB.Line B1_K_Line 
      Index           =   6
      X1              =   14085
      X2              =   14085
      Y1              =   1410
      Y2              =   1665
   End
   Begin VB.Line B1_K_Line 
      Index           =   7
      X1              =   14475
      X2              =   14475
      Y1              =   1410
      Y2              =   1665
   End
   Begin VB.Line B1_K_Line 
      Index           =   8
      X1              =   14865
      X2              =   14865
      Y1              =   1410
      Y2              =   1665
   End
   Begin VB.Line B1_K_Line 
      Index           =   9
      X1              =   15270
      X2              =   15270
      Y1              =   1410
      Y2              =   1665
   End
   Begin VB.Line B1_K_Line 
      Index           =   10
      X1              =   15660
      X2              =   15660
      Y1              =   1410
      Y2              =   1665
   End
   Begin VB.Label B1_MAISU 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "ZZ9"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   15885
      TabIndex        =   22
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "枚数"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   15780
      TabIndex        =   21
      Top             =   960
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "├──── 《 発行ＮＯ 》────┤"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11610
      TabIndex        =   20
      Top             =   960
      Width           =   4200
   End
   Begin VB.Label B1_ERRCMNT 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   165
      TabIndex        =   19
      Top             =   1440
      Width           =   2700
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "エラー内容"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1365
      TabIndex        =   18
      Top             =   945
      Width           =   1200
   End
   Begin VB.Label B1_UKEIR 
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5760
      TabIndex        =   17
      Top             =   1440
      Width           =   210
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "発行日時："
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   11670
      TabIndex        =   16
      Top             =   405
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "仕  入  先"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   11
      Left            =   4230
      TabIndex        =   15
      Top             =   945
      Width           =   1200
   End
   Begin VB.Label B1_TORNM 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3615
      TabIndex        =   14
      Top             =   1440
      Width           =   2100
   End
   Begin VB.Label B1_TORCD 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   13
      Top             =   1440
      Width           =   735
   End
   Begin VB.Line H1_K_LINE 
      X1              =   255
      X2              =   16305
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label B1_Seban 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6015
      TabIndex        =   12
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label H1_HYMD 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD hh:ss"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12840
      TabIndex        =   11
      Top             =   405
      Width           =   1920
   End
   Begin VB.Label H1_PAGE 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "Z,ZZ9"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   15315
      TabIndex        =   10
      Top             =   405
      Width           =   615
   End
   Begin VB.Label B1_HINBN 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7035
      TabIndex        =   9
      Top             =   1440
      Width           =   2625
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "受"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5745
      TabIndex        =   8
      Top             =   960
      Width           =   240
   End
   Begin VB.Label B1_TEHAIKBN 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9780
      TabIndex        =   7
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label B1_SYUYO 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "ZZ,ZZ9"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10260
      TabIndex        =   6
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "※※ かんばん読取エラーリスト ※※"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4545
      TabIndex        =   5
      Top             =   285
      Width           =   5610
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "背番号"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6045
      TabIndex        =   4
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "頁"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   15960
      TabIndex        =   3
      Top             =   405
      Width           =   225
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "品　　番"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   7605
      TabIndex        =   2
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "手配"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9735
      TabIndex        =   1
      Top             =   960
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "収容数"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10245
      TabIndex        =   0
      Top             =   960
      Width           =   630
   End
End
Attribute VB_Name = "CKH0910L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  かんばん読取エラーリスト
'**       フォームID    :  CKH0910L
'**       処理概要      :
'**
'**       作  成  日    :  2004/01/06
'**       変  更  日    :  2004/11/26 By CIS ｴﾗｰﾒｯｾｰｼﾞ追加(手配区分相違)
'**       変  更  日    :  2004/05/24  便算出時、カレンダ無しエラーが
'**                                     『EXECUTE後のトランザクション数は、COMMIT TRANまたは･･･』となるのを
'**                                         カレンダ未登録エラーに変更( ERROR No = 6 )
'**       変  更  日    :  2006/12/20  ErrNo = 9 版数異常　エラー追加
'**
'**********************************************************'
'#比較退避用#
    Dim BK_ErrCD        As Byte         'エラーコード
    Dim BK_Torcd        As String       '仕入先CD(取引先+工場CD)
    Dim BK_Ukeir        As String       '受入
    Dim BK_Seban        As String       '背番号
    Dim BK_Hinbn        As String       '品番
    Dim BK_TehaiKB      As String       '手配区分(0/購入,1/直送,2/社内,3/受給)
    Dim BK_Syuyo        As Long         '収容数
        
    Dim lIx             As Byte         '添字
    Dim TMai            As Long         '枚数
    
    Dim mNo(10)         As Long
    
    Dim mPrtDate        As String

'*------------------------------------*
'*      Ｆｏｒｍ＿Ｌｏａｄ イベント     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd hh:nn")             '発行日・時間
    '*---------------------*
    '*  プリント系初期設定
    '*---------------------*
    With CisPrtCKH0910
        .iPrtTitle = "かんばん読取エラーリスト"
        .iLineMax = 40
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
'*---------------------*
'*　  印  刷  処  理
'*---------------------*
Private Function PrintProc()
    Dim FirstFLG    As Boolean      '初回回避用(TRUE:Body行 印刷)
    Dim BreakFLG    As Boolean      'KEY-BREAK(FALSE:keyBreak,TRUE:not_KeyBreak)
    
    lIx = 0: TMai = 0
    Erase mNo
    BK_ErrCD = 0
    
    Call CisPrtCKH0910.ClrLstFld("B1")
    
    Do Until Not KRE_RDSTS
        With KRE
            
        '*---------------------------------*
        '*  Head部　印刷
        '*---------------------------------*
            Call HeadWrite(0)
        '*---------------------------------*
        '*  KEY-BRAEK　判定
        '*---------------------------------*
        If .エラー要因 <> BK_ErrCD Or _
           .取引先 <> BK_Torcd Or _
           .受入 <> BK_Ukeir Or _
           .背番号 <> BK_Seban Or _
           .品番 <> BK_Hinbn Or _
           .手配区分 <> BK_TehaiKB Or _
           .収容数 <> BK_Syuyo Then
            BreakFLG = False
        Else
            BreakFLG = True
        End If
        
        '*---------------------------------*
        '*  Body行　印刷
        '*---------------------------------*
            If FirstFLG Then
                If lIx + 1 > 10 Or Not BreakFLG Then
                    B1_Line.Visible = Not BreakFLG
                    If Not BreakFLG Then
                       B1_MAISU = CisFun.RSetFld(TMai, 3, "##0")
                       TMai = 0
                    End If
                    Call CisPrtCKH0910.PrintFld("B1*")
                    CisPrtCKH0910.cLine = CisPrtCKH0910.cLine + 1
                    Call CisPrtCKH0910.ClrLstFld("B1")
                    lIx = 0: Erase mNo
                End If
            Else
                FirstFLG = True
            End If
            
        '*---------------------------------*
        '*  印字項目　転送
        '*---------------------------------*
            '#エラー要因
            If CisPrtCKH0910.cLine >= CisPrtCKH0910.iLineMax Or _
               .エラー要因 <> BK_ErrCD Then
                Select Case .エラー要因
                    Case 1: B1_ERRCMNT = "対象外かんばん(かんばん区分)"
                    Case 2: B1_ERRCMNT = "品番取引先未登録"
                    Case 3: B1_ERRCMNT = "発注サイクルマスタ未登録"
                    Case 4: B1_ERRCMNT = "かんばん流動管理ﾃｰﾌﾞﾙ未登録"
                    Case 5: B1_ERRCMNT = "かんばん重複読取"
                    
'----------------------------------------------------------------< 2005.02.24 ADD START >
                    Case 6: B1_ERRCMNT = "カレンダマスタ未登録"
'----------------------------------------------------------------< 2005.02.24 ADD START >
'----------------------------------------------------------------< 2006.12.20 ADD START >
                    Case 9: B1_ERRCMNT = "版数異常(紛失かんばん)"
'----------------------------------------------------------------< 2006.12.20 ADD START >
                    
                    Case 10: B1_ERRCMNT = "収容数相違"
                    Case 11: B1_ERRCMNT = "背番号相違"
                    Case 15: B1_ERRCMNT = "無効かんばん"
                    Case 16: B1_ERRCMNT = "未流動かんばん"
                    Case 17: B1_ERRCMNT = "滞留かんばん"
                    Case 18: B1_ERRCMNT = "打切り済"
                    Case 19: B1_ERRCMNT = "有効期限エラー"
                    Case 20: B1_ERRCMNT = "サイクルエラー"
                    Case 21: B1_ERRCMNT = "減かんばん(無効)"
                    Case 30: B1_ERRCMNT = "最終発注日算出エラー"
                    Case 31: B1_ERRCMNT = "開始発注日前"
                    Case 32: B1_ERRCMNT = "サイクル相違"
                    Case 33: B1_ERRCMNT = "手配区分相違"
'----------------------------------------------------------------< 2005.03.10 ADD START >
                    Case 35: B1_ERRCMNT = "最終発注日･便より前の発注日"
'----------------------------------------------------------------< 2005.03.10 ADD START >
                    Case 45: B1_ERRCMNT = "グループ外"
                End Select
                BK_Torcd = ""
            End If
            
            '#取引先,取引先略称
            If .取引先 <> BK_Torcd Then
                B1_TORCD = .取引先                     '.取引先
                B1_TORNM = .取引先名     '.取引先略称
                BK_Ukeir = ""
            End If
            
            '#受入
            If .受入 <> BK_Ukeir Then
                B1_UKEIR = .受入
                BK_Seban = ""
            End If
                
            '#背番号
            If .背番号 <> BK_Seban Then
                B1_Seban = .背番号
                BK_Hinbn = ""
            End If

            '#品番,(背番号)
            If .品番 <> BK_Hinbn Then
                B1_HINBN = .表示品番
                B1_Seban = .背番号
            End If

            '#手配区分,収容数
            If Not BreakFLG Then
               B1_TEHAIKBN = .手配区分名
               B1_SYUYO = CisFun.RSetFld(.収容数, 6, "#,##0")
               If Trim(.サイクル) <> "" Then
                  B1_Cycle = Mid(.サイクル, 1, 1) & "-" & Mid(.サイクル, 2, 2) & "-" & _
                             Mid(.サイクル, 4, 2)
               End If
            End If
            
            '#発行NO
            lIx = lIx + 1
            B1_NO(lIx) = CisFun.RSetFld(.発行NO, 3, "000")
            ' 枚数カウント
            TMai = TMai + 1
        '*---------------------------------*
        '*  比較項目　退避
        '*---------------------------------*
            BK_ErrCD = .エラー要因             '.エラー要因
            BK_Torcd = .取引先                  '.取引先
            BK_Ukeir = .受入
            BK_Seban = .背番号
            BK_Hinbn = .品番
            BK_TehaiKB = .手配区分              '.手配区分
            BK_Syuyo = .収容数
        
        End With
        
        '*---------------------------------*
        '*  カウンタ　ＵＰ
        '*---------------------------------*
        With CisPrtCKH0910
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(CKH0910_PrtCnt, "#,###")
        End With
        
        Call KREReadNext(1)
    Loop
    
    '*---------------------------------*
    '*  最終Body行　印刷
    '*---------------------------------*
    Call HeadWrite(0)
    B1_Line.Visible = True
    B1_MAISU = CisFun.RSetFld(TMai, 3, "##0")
    Call CisPrtCKH0910.PrintFld("B1")
    CisPrtCKH0910.cLine = CisPrtCKH0910.cLine + 1
    
    Call KREClose(1)
    
    ' 印刷終了
    Call CisPrtCKH0910.PrintEndDoc
    
End Function
'*----------------*
'*  ヘッダ部印刷  *
'*----------------*
Private Function HeadWrite(KB As Byte)
    
    Dim NewPageFlg As Boolean
'
    HeadWrite = False
    With CisPrtCKH0910
        ' １頁内印刷行数ﾁｪｯｸ
        NewPageFlg = False
        
        If .iLineMax >= .cLine + 1 And Not NewPageFlg Then Exit Function
        
        .cLine = 0
        .cPage = .cPage + 1
        
        .ClrLstFld ("H1")   'HEAD部内容消去
    
        H1_PAGE = CisFun.RSetFld(.cPage, 5, "###")
        H1_HYMD = mPrtDate
        
        
        
        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With
    HeadWrite = True
End Function

