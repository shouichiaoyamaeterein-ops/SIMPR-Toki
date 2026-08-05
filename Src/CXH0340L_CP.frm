VERSION 5.00
Object = "{0D300FC0-B2EA-11D1-8D3B-444553540000}#1.30#0"; "Qrmaker.ocx"
Begin VB.Form CXH0340L 
   Caption         =   "臨時かんばん"
   ClientHeight    =   5940
   ClientLeft      =   3510
   ClientTop       =   3510
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   11640
   Begin QRMAKERLib.QRmaker B1_QRmaker1 
      Height          =   1080
      Left            =   9270
      TabIndex        =   32
      Top             =   1350
      Width           =   1185
      _Version        =   65566
      _ExtentX        =   2090
      _ExtentY        =   1905
      _StockProps     =   1
      Picture         =   "CXH0340L.frx":0000
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   585
      Index           =   8
      Left            =   225
      Top             =   4155
      Width           =   2445
   End
   Begin VB.Label B1_OkibaM 
      AutoSize        =   -1  'True
      Caption         =   "置場"
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
      Left            =   300
      TabIndex        =   34
      Top             =   4380
      Width           =   390
   End
   Begin VB.Label B1_Okiba 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   810
      TabIndex        =   33
      Top             =   4305
      Width           =   1515
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   585
      Index           =   6
      Left            =   2670
      Top             =   4155
      Width           =   6045
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   585
      Index           =   5
      Left            =   7185
      Top             =   2955
      Width           =   1530
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   585
      Index           =   4
      Left            =   2670
      Top             =   3555
      Width           =   6045
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   585
      Index           =   3
      Left            =   5085
      Top             =   2955
      Width           =   2070
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   585
      Index           =   2
      Left            =   2670
      Top             =   2955
      Width           =   2385
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   1350
      Index           =   1
      Left            =   9315
      Top             =   3060
      Width           =   1380
   End
   Begin VB.Label B1_HinnmM 
      AutoSize        =   -1  'True
      Caption         =   "品名"
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
      Left            =   2745
      TabIndex        =   30
      Top             =   4365
      Width           =   390
   End
   Begin VB.Label B1_Hinnm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
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
      Left            =   3270
      TabIndex        =   29
      Top             =   4350
      Width           =   4800
   End
   Begin VB.Label B1_HYmdM 
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
      Left            =   9120
      TabIndex        =   28
      Top             =   4575
      Width           =   540
   End
   Begin VB.Label B1_HYmd 
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
      Left            =   9750
      TabIndex        =   27
      Top             =   4545
      Width           =   1050
   End
   Begin VB.Label B1_UkeirM 
      AutoSize        =   -1  'True
      Caption         =   "受入"
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
      Left            =   9750
      TabIndex        =   26
      Top             =   3180
      Width           =   390
   End
   Begin VB.Label B1_Ukeir 
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   36
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   9585
      TabIndex        =   25
      Top             =   3450
      Width           =   900
   End
   Begin VB.Label B1_SyakeiM 
      AutoSize        =   -1  'True
      Caption         =   "車型"
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
      Left            =   7290
      TabIndex        =   24
      Top             =   3195
      Width           =   390
   End
   Begin VB.Label B1_Syakei 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7755
      TabIndex        =   23
      Top             =   3120
      Width           =   825
   End
   Begin VB.Label B1_HinbnM 
      AutoSize        =   -1  'True
      Caption         =   "品番"
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
      Left            =   2745
      TabIndex        =   22
      Top             =   3780
      Width           =   390
   End
   Begin VB.Label B1_Hinbn 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3255
      TabIndex        =   21
      Top             =   3645
      Width           =   5610
   End
   Begin VB.Label B1_SyuyoM 
      AutoSize        =   -1  'True
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
      Left            =   5145
      TabIndex        =   20
      Top             =   3180
      Width           =   585
   End
   Begin VB.Label B1_Syuyo 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "99,999"
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
      Left            =   5730
      TabIndex        =   19
      Top             =   3075
      Width           =   1350
   End
   Begin VB.Label B1_SyukiM 
      AutoSize        =   -1  'True
      Caption         =   "収容器"
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
      Left            =   2730
      TabIndex        =   18
      Top             =   3195
      Width           =   585
   End
   Begin VB.Label B1_Syuki 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3420
      TabIndex        =   17
      Top             =   3120
      Width           =   1515
   End
   Begin VB.Label B1_Maisu 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "( 999 / 999 )"
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
      Left            =   705
      TabIndex        =   16
      Top             =   3765
      Width           =   1170
   End
   Begin VB.Label B1_NBin 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   540
      TabIndex        =   15
      Top             =   3045
      Width           =   600
   End
   Begin VB.Label B1_NYmdM 
      AutoSize        =   -1  'True
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
      Height          =   180
      Left            =   255
      TabIndex        =   14
      Top             =   2445
      Width           =   540
   End
   Begin VB.Label B1_NYmd 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   13
      Top             =   2685
      Width           =   2100
   End
   Begin VB.Label B1_JisyaKj 
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
      Left            =   8700
      TabIndex        =   12
      Top             =   585
      Width           =   2400
   End
   Begin VB.Label B1_NonyuKj 
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
      Left            =   240
      TabIndex        =   11
      Top             =   1830
      Width           =   2400
   End
   Begin VB.Label B1_Nonyu 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
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
      Left            =   855
      TabIndex        =   10
      Top             =   1335
      Width           =   630
   End
   Begin VB.Label B1_NonyuNm 
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
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   2400
   End
   Begin VB.Label B1_NonyuM 
      AutoSize        =   -1  'True
      Caption         =   "納入先"
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
      Left            =   225
      TabIndex        =   8
      Top             =   1335
      Width           =   540
   End
   Begin VB.Label B1_SiireKj 
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
      Left            =   240
      TabIndex        =   7
      Top             =   705
      Width           =   2400
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   4
      Height          =   1335
      Index           =   0
      Left            =   2685
      Top             =   1380
      Width           =   6015
   End
   Begin VB.Label B1_SiireM 
      AutoSize        =   -1  'True
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
      Height          =   180
      Left            =   225
      TabIndex        =   6
      Top             =   195
      Width           =   540
   End
   Begin VB.Label B1_Tehai 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5385
      TabIndex        =   5
      Top             =   765
      Width           =   720
   End
   Begin VB.Label B1_SiireNm 
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
      Left            =   240
      TabIndex        =   4
      Top             =   435
      Width           =   2400
   End
   Begin VB.Label B1_Siire 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
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
      Left            =   855
      TabIndex        =   3
      Top             =   195
      Width           =   630
   End
   Begin VB.Label B1_Seban 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   60
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   2790
      TabIndex        =   2
      Top             =   1425
      Width           =   5895
   End
   Begin VB.Label B1_JisyaNm 
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
      Left            =   8685
      TabIndex        =   1
      Top             =   300
      Width           =   2400
   End
   Begin VB.Label B1_PrintKb 
      AutoSize        =   -1  'True
      Caption         =   "臨時かんばん"
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
      Left            =   4170
      TabIndex        =   0
      Top             =   165
      Width           =   3240
   End
   Begin VB.Label B1_Hinbn30 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3285
      TabIndex        =   31
      Top             =   3720
      Width           =   5400
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      Height          =   4890
      Index           =   7
      Left            =   30
      Top             =   0
      Width           =   11220
   End
End
Attribute VB_Name = "CXH0340L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  臨時かんばん　発行
'**       フォームID    :  CXH0340L
'**       処理概要      :
'**
'**       作  成  日    :  2004/01/30
'**       変  更  日    :  2008/04/11 By CIS - 背番号８桁対応、車型５桁対応
'**                        2008/11/14 By CIS - 工区出力変更
'**********************************************************'
'#比較退避用#
    Dim MaxGyo          As Byte         '頁行数
    Dim GMaisu          As Long         '現枚数
    Dim TMaisu          As Long         '総枚数
    
    Dim mPrtDate        As String
    Dim mPrtTime        As String
    Dim mQRData         As String * 150
    Dim mHMoto          As String
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
    With CisPrtCXH0340
        .iPrtTitle = "臨時かんばん"
        .iLineMax = 80
'        .iLinePitch = 8
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
    Dim GGyo    As Integer
'*---------------------------------*
'*   システム情報(発行元)
'*---------------------------------*
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 4
    Call IRNRead("", 1)
    mHMoto = Mid(IRN.oyian, 1, 7)
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
    MaxGyo = 3: GGyo = 0
    Call CisPrtCXH0340.ClrLstFld("B1")
    Call HeadWrite(0)
    
'*---------------------------------*
'*  発注テーブル読込み
'*---------------------------------*
    Do Until Not HCT_RDSTS
        With HCT
            '打切端数の時の端数分は端数ラベルで発行
            If .端数 <> 0 And .端数区分 <> 0 Then
                TMaisu = .枚数 - 1
            Else
                TMaisu = .枚数
            End If
        '*---------------------------------*
        '*  Body行　印刷（枚数分繰返し）
        '*---------------------------------*
            For GMaisu = 1 To TMaisu
                If GGyo + 1 > MaxGyo Then
                    Call HeadWrite(0)
                    GGyo = 0
                End If
                GGyo = GGyo + 1
                
                B1_PrintKb = "臨時かんばん"
                B1_Tehai = RTrim(.手配区分名)
                B1_JisyaNm = RTrim(mJisyaNm)
                '=================== *** 2008/11/14 Start
                'B1_JisyaKj = RTrim(mJisyaKj)
                gSL_Select = "SELECT 値名称 FROM 名称マスタ "
                gSL_Select = gSL_Select & " WHERE 区分名称 = '自社工区'"
                gSL_Select = gSL_Select & "   AND 英数字1 = '" & RTrim(.受入) & "'"
                gSL_Select = gSL_Select & "   AND 区分タイプ = ''"
                If SYMRead(gSL_Select, 1, 0) Then
                    B1_JisyaKj = RTrim(SYM.値名称)
                Else
                    B1_JisyaKj = RTrim(mJisyaKj)
                End If
                '=================== *** 2008/11/14 End
                
                B1_SiireM = "仕入先"
                B1_Siire = RTrim(.仕入先CD)
                If RTrim(.仕入先工場CD) <> "" Then
                    B1_Siire = B1_Siire & "-" & RTrim(.仕入先工場CD)
                End If
                B1_SiireNm = RTrim(.仕入先名)
                B1_SiireKj = RTrim(.仕入先工場名)
                If RTrim(.納入先) <> "" Then
                    B1_NonyuM = "納入先"
                    B1_Nonyu = RTrim(.納入先CD)
                    If RTrim(.納入先工場CD) <> "" Then
                        B1_Nonyu = B1_Nonyu & "-" & RTrim(.納入先工場CD)
                    End If
                    B1_NonyuNm = RTrim(.納入先名)
                    B1_NonyuKj = RTrim(.納入先工場名)
                End If
                
                B1_NYmdM = "納入日"
                CisFun.DateE_BefVal = .納入日
                CisFun.DateEdit (J_YMD)
                B1_NYmd = CisFun.DateE_AfVal
                If RTrim(.納入便) <> "" Then
                    B1_NBin = RTrim(.納入便) & "便"
                End If
                B1_Maisu = "( " & Format(GMaisu, "#,##0") & " / " & Format(TMaisu, "#,##0") & " )"
                
                B1_Seban = RTrim(.背番号)
                B1_SyukiM = "収容器"
                B1_Syuki = RTrim(.収容器)
                B1_SyuyoM = "収容数"
                B1_Syuyo = Format(.収容数, "#,##0")
                B1_SyakeiM = "車型"
                B1_Syakei = RTrim(.車型)
                B1_HinbnM = "品番"
                If Len(RTrim(.表示品番)) <= 22 Then
                    B1_Hinbn = RTrim(.表示品番)
                Else
                    B1_Hinbn30 = RTrim(.表示品番)
                End If
                B1_HinnmM = "品名"
                B1_Hinnm = RTrim(.品名)
                
                B1_UkeirM = "受入"
                B1_Ukeir = RTrim(.受入)
                B1_HYmdM = "発行日"
                B1_HYmd = RTrim(mPrtDate)
                B1_OkibaM = "置場"
                B1_Okiba = RTrim(.置場)
            'QR
                Call QR_Edit
                B1_QRmaker1.InputData = mQRData
                B1_QRmaker1.Refresh
            'BODY部印刷
                Call CisPrtCXH0340.PrintFld("B1*")
                CisPrtCXH0340.cLine = CisPrtCXH0340.cLine + 23
                Call CisPrtCXH0340.ClrLstFld("B1")
                
                '*---------------------------------*
                '*  カウンタ　ＵＰ
                '*---------------------------------*
                With CisPrtCXH0340
                    .cDataCnt = .cDataCnt + 1
                    .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(CXH0340_PrtCnt, "#,###")
                End With
            Next
            
            
        
        End With
        Call HCTReadNext(1)
    Loop
    
    Call HCTClose(1)
    
    ' 印刷終了
    Call CisPrtCXH0340.PrintEndDoc
    
End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim NewPageFlg As Boolean
    HeadWrite = False
    
    With CisPrtCXH0340
        NewPageFlg = False
        .cLine = 0
        .cPage = .cPage + 1
'
        .ClrLstFld ("H1")   'HEAD部内容消去
'
        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With
    HeadWrite = True
End Function
'*---------------------------------------------------------------*
'*   << QR内容編集 >>                                             *
'*---------------------------------------------------------------*
Private Sub QR_Edit()
    mQRData = Space(150)
    
    Mid(mQRData, 1, 7) = Left(mHMoto, 7)                '発行元
    Mid(mQRData, 8, 5) = gCompany                       '自社コード
    Mid(mQRData, 13, 1) = "R"                           'QR識別(R:臨時かんばん、S:現品票、H:端数ラベル)
    Mid(mQRData, 14, 25) = HCT.品番                     '品番
    Mid(mQRData, 39, 1) = "1"                           '品目(0:受注、1:発注)
    Mid(mQRData, 40, 7) = HCT.仕入先                    '取引先コード
    Mid(mQRData, 47, 2) = HCT.受入                      '取引先受入
    Mid(mQRData, 49, 1) = Format(HCT.手配区分, "0")     '手配区分
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
'    Mid(mQRData, 50, 5) = HCT.背番号                    '背番号
'    Mid(mQRData, 55, 5) = HCT.納入番号                  '納入番号
'    Mid(mQRData, 60, 8) = HCT.納入日                    '納入日
'    Mid(mQRData, 68, 2) = HCT.納入便                    '納入便
'    Mid(mQRData, 70, 7) = HCT.納入先                    '納入先コード
'    Mid(mQRData, 77, 2) = HCT.納入先受入                '納入先受入
'    Mid(mQRData, 79, 6) = Format(HCT.収容数, "000000")  '収容数
'    Mid(mQRData, 85, 10) = HCT.収容器                   '収容器
'    Mid(mQRData, 95, 6) = Format(HCT.収容数, "000000")  '数量
'    Mid(mQRData, 101, 3) = Format(GMaisu, "000")          '現枚数
'    Mid(mQRData, 104, 3) = Format(TMaisu, "000")        '総枚数
'    CisFun.DateE_AfVal = mPrtDate
'    CisFun.DateEditUn
'    Mid(mQRData, 107, 8) = CisFun.DateE_BefVal          '発行日
'    Mid(mQRData, 115, 5) = mPrtTime                     '発行時間
'    Mid(mQRData, 120, 5) = HCT.サイクル                 'サイクル
'    Mid(mQRData, 125, 10) = HCT.置場                    '置場
'    Mid(mQRData, 135, 4) = HCT.車型                     '車型
'    Mid(mQRData, 139, 1) = Format(HCT.発注区分, "0")    '発注区分
    Mid(mQRData, 50, 8) = HCT.背番号                    '背番号
    Mid(mQRData, 58, 5) = HCT.納入番号                  '納入番号
    Mid(mQRData, 63, 8) = HCT.納入日                    '納入日
    Mid(mQRData, 71, 2) = HCT.納入便                    '納入便
    Mid(mQRData, 73, 7) = HCT.納入先                    '納入先コード
    Mid(mQRData, 80, 2) = HCT.納入先受入                '納入先受入
    Mid(mQRData, 82, 6) = Format(HCT.収容数, "000000")  '収容数
    Mid(mQRData, 88, 10) = HCT.収容器                   '収容器
    Mid(mQRData, 98, 6) = Format(HCT.収容数, "000000")  '数量
    Mid(mQRData, 104, 3) = Format(GMaisu, "000")          '現枚数
    Mid(mQRData, 107, 3) = Format(TMaisu, "000")        '総枚数
    CisFun.DateE_AfVal = mPrtDate
    CisFun.DateEditUn
    Mid(mQRData, 110, 8) = CisFun.DateE_BefVal          '発行日
    Mid(mQRData, 118, 5) = mPrtTime                     '発行時間
    Mid(mQRData, 123, 5) = HCT.サイクル                 'サイクル
    Mid(mQRData, 128, 10) = HCT.置場                    '置場
    Mid(mQRData, 138, 5) = HCT.車型                     '車型
    Mid(mQRData, 143, 1) = Format(HCT.発注区分, "0")    '発注区分
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
End Sub

