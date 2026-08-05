VERSION 5.00
Begin VB.Form CXHZ0510L 
   Caption         =   "明細書"
   ClientHeight    =   3750
   ClientLeft      =   3375
   ClientTop       =   1575
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   11505
   Begin VB.Label B1_Sunpo 
      AutoSize        =   -1  'True
      Caption         =   "9999.99-9999.99-9999.99"
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
      Left            =   2820
      TabIndex        =   27
      Top             =   1860
      Width           =   2415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "寸    法"
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
      Left            =   3600
      TabIndex        =   26
      Top             =   1515
      Width           =   810
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
      Left            =   10740
      TabIndex        =   25
      Top             =   90
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
      Index           =   2
      Left            =   11100
      TabIndex        =   24
      Top             =   120
      Width           =   225
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
   Begin VB.Label B1_Ymd 
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99-99"
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
      Left            =   5310
      TabIndex        =   23
      Top             =   1860
      Width           =   1365
   End
   Begin VB.Label B1_Suryo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999,990.000"
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
      Left            =   9510
      TabIndex        =   22
      Top             =   1860
      Width           =   1275
   End
   Begin VB.Label B1_Maisu 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "99,999"
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
      Left            =   8790
      TabIndex        =   21
      Top             =   1860
      Width           =   630
   End
   Begin VB.Label B1_Syuyo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999,990.000"
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
      Left            =   7500
      TabIndex        =   20
      Top             =   1860
      Width           =   1155
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
      Left            =   10905
      TabIndex        =   19
      Top             =   1860
      Width           =   210
   End
   Begin VB.Label B1_Zaist 
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
      Left            =   120
      TabIndex        =   18
      Top             =   1860
      Width           =   2625
   End
   Begin VB.Label B1_Nouno 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
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
      Left            =   6810
      TabIndex        =   17
      Top             =   1860
      Width           =   525
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   11190
      Y1              =   1815
      Y2              =   1815
   End
   Begin VB.Label H1_Ymd 
      AutoSize        =   -1  'True
      Caption         =   "実 績 日"
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
      Left            =   5550
      TabIndex        =   16
      Top             =   1515
      Width           =   795
   End
   Begin VB.Label H1_Suryo 
      AutoSize        =   -1  'True
      Caption         =   "実績数"
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
      Left            =   9945
      TabIndex        =   15
      Top             =   1515
      Width           =   585
   End
   Begin VB.Label Label3 
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
      Left            =   9000
      TabIndex        =   14
      Top             =   1515
      Width           =   435
   End
   Begin VB.Label Label3 
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
      Index           =   6
      Left            =   7830
      TabIndex        =   13
      Top             =   1515
      Width           =   585
   End
   Begin VB.Label Label3 
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
      Index           =   3
      Left            =   10800
      TabIndex        =   12
      Top             =   1515
      Width           =   390
   End
   Begin VB.Label Label3 
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
      Left            =   600
      TabIndex        =   11
      Top             =   1515
      Width           =   810
   End
   Begin VB.Label Label1 
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
      Index           =   3
      Left            =   2730
      TabIndex        =   10
      Top             =   480
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "納入番号"
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
      Left            =   6690
      TabIndex        =   9
      Top             =   1515
      Width           =   780
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   0
      X1              =   120
      X2              =   11190
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Label H1_TorTtl 
      Caption         =   "実績先"
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
      Index           =   1
      Left            =   8160
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
      Left            =   8790
      TabIndex        =   6
      Top             =   90
      Width           =   1680
   End
   Begin VB.Label H1_JisyaKj 
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
      Left            =   8670
      TabIndex        =   5
      Top             =   660
      Width           =   2400
   End
   Begin VB.Label H1_Torkj 
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
      Left            =   255
      TabIndex        =   3
      Top             =   435
      Width           =   2400
   End
   Begin VB.Label H1_Torcd 
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
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8670
      TabIndex        =   1
      Top             =   405
      Width           =   2400
   End
   Begin VB.Label H1_JKbn 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "※　実績明細書　※"
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
      Left            =   3855
      TabIndex        =   0
      Top             =   315
      Width           =   3390
   End
End
Attribute VB_Name = "CXHZ0510L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  材料　入出庫・支給明細書　発行
'**       フォームID    :  CXHZ0510L
'**       処理概要      :
'**
'**       作  成  日    :  2004/02/26
'**       変  更  日    :
'**
'**********************************************************'
'   #比較退避用#
    Dim Key_Torcd       As String       '>実績先
    Dim Key_Zkno        As Long         '>材料管理番号
    Dim GMaisu          As Currency     ' 箱数　合計
    Dim GSuryo          As Currency     ' 実績数合計
    Dim GKensu          As Long         ' 件数

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
    With CisPrtCXHZ0510
        Select Case ZJK.実績区分
            Case 1: .iPrtTitle = "材料・入庫明細書"
            Case 2: .iPrtTitle = "材料・出庫明細書"
            Case 3: .iPrtTitle = "材料・支給明細書"
        End Select
        .iLineMax = 58
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
'   システム情報(自社工場)
    Call ItemsClearIRN
    IRN.bk = "Option"
    IRN.rno = 7
    Call IRNRead("", 1)
'   取引先マスタより名称を求める
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
    Key_Torcd = ZJK.実績先
    Key_Zkno = ZJK.材料管理番号
    GMaisu = 0
    GSuryo = 0
    GKensu = 0

    Call CisPrtCXHZ0510.ClrLstFld("B1")

'*---------------------------------*
'*  実績管理テーブル読込み
'*---------------------------------*
    Do Until Not ZJK_RDSTS
        If Key_Torcd <> ZJK.実績先 Or _
           Key_Zkno <> ZJK.材料管理番号 Then
        '   合計印刷
            Call Total_Print
        '   改頁判定
            If Key_Torcd <> ZJK.実績先 Then
                CisPrtCXHZ0510.cLine = CisPrtCXHZ0510.iLineMax
            End If
        End If

        Call HeadWrite(0)

        '*---------------------------------*
        '*  Body行　編集
        '*---------------------------------*
        With ZJK
            If 0 = GKensu Or _
               0 = CisPrtCXHZ0510.cLine = 0 Then
                B1_Zaist = RTrim(.材質)
                B1_Sunpo = RTrim(.表示寸法)
                Call MeisyoGet("単位区分", .単位)
            End If
            If RTrim(.実績便) <> "" Then
                B1_Ymd = Format(.実績日, "0000/00/00") & "-" & .実績便
            Else
                B1_Ymd = Format(.実績日, "0000/00/00") & "   "
            End If
            B1_Nouno = RTrim(.納入番号)
            B1_Syuyo = Format(.収容数, "##,##0.000")
            B1_Maisu = Format(.実績枚数, "##,###")
            B1_Suryo = Format(.実績数, "##,##0.000")
            B1_Tannm = SYM_Meisyo
    '       合計計算
            GMaisu = GMaisu + .実績枚数
            GSuryo = GSuryo + .実績数
            GKensu = GKensu + 1
    '       KEY項目退避
            Key_Torcd = .実績先
            Key_Zkno = .材料管理番号
        End With
        '*---------------------------------*
        '*  Body部　印刷
        '*---------------------------------*
        With CisPrtCXHZ0510
            .PrintFld ("B1*")
            .cLine = .cLine + 1
            .ClrLstFld ("B1")
    '       カウンタ　ＵＰ
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With

        Call ZJKReadNext(1)
    Loop
    Call ZJKClose(1)

'   合計印刷
    Call Total_Print
'   印刷終了
    Call CisPrtCXHZ0510.PrintEndDoc

End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim TtlName     As String
    Dim YmdName     As String

    HeadWrite = False
    
    With CisPrtCXHZ0510
    '   １頁内印刷行数ﾁｪｯｸ
        If .iLineMax >= .cLine + 1 Then Exit Function

        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")   'HEAD部内容消去

    '   明細書ﾀｲﾄﾙ
        Select Case ZJK.実績区分
            Case 1: TtlName = "入庫": YmdName = "入庫"
            Case 2: TtlName = "出庫": YmdName = "出庫"
            Case 3: TtlName = "支給": YmdName = "納入"
        End Select
        H1_JKbn = "※　材料・" & TtlName & "明細書　※"
        H1_TorTtl = TtlName & "先"
        H1_Suryo = TtlName & "数"
        H1_Ymd = YmdName & "日"
    '   発行日・頁
        H1_HYmd = mPrtDate & " " & mPrtTime
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")
    '   自社名
        H1_JisyaNm = mJisyaNm
        H1_JisyaKj = mJisyaKj
    '   実績先
        If RTrim(ZJK.工場CD) <> "" Then
            H1_Torcd = ZJK.取引先 & " - " & RTrim(ZJK.工場CD)
        Else
            H1_Torcd = ZJK.取引先
        End If
        H1_Tornm = RTrim(ZJK.実績先名)
        H1_Torkj = RTrim(ZJK.工場名)

        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With

    HeadWrite = True
End Function
Private Function Total_Print()
    Total_Print = False
'
    If GKensu > 1 Then
        B1_Zaist = "-----<< 材 料 合 計 >>-----"
        B1_Maisu = Format(GMaisu, "##,##0")
        B1_Suryo = Format(GSuryo, "##,##0.000")
        Call CisPrtCXHZ0510.PrintFld("B1*")
        CisPrtCXHZ0510.cLine = CisPrtCXHZ0510.cLine + 2
        Call CisPrtCXHZ0510.ClrLstFld("B1")
    Else
        CisPrtCXHZ0510.cLine = CisPrtCXHZ0510.cLine + 1
    End If
    GMaisu = 0
    GSuryo = 0
    GKensu = 0
'
    Total_Print = True
End Function

