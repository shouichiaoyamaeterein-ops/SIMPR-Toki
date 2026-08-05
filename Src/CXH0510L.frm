VERSION 5.00
Begin VB.Form CXH0510L 
   Caption         =   "明細書"
   ClientHeight    =   5940
   ClientLeft      =   3390
   ClientTop       =   1545
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   11700
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
      Index           =   12
      Left            =   11100
      TabIndex        =   24
      Top             =   120
      Width           =   225
   End
   Begin VB.Line B1_K_LINE2 
      X1              =   1080
      X2              =   11175
      Y1              =   1845
      Y2              =   1845
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
   Begin VB.Line B1_K_LINE1 
      X1              =   195
      X2              =   11160
      Y1              =   1845
      Y2              =   1845
   End
   Begin VB.Label B1_Ymd 
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99-99"
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
      Left            =   5190
      TabIndex        =   23
      Top             =   1860
      Width           =   1560
   End
   Begin VB.Label B1_Suryo 
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
      Left            =   10140
      TabIndex        =   22
      Top             =   1860
      Width           =   840
   End
   Begin VB.Label B1_Maisu 
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
      Left            =   9120
      TabIndex        =   21
      Top             =   1860
      Width           =   840
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
      Left            =   8085
      TabIndex        =   20
      Top             =   1860
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
      Left            =   4095
      TabIndex        =   19
      Top             =   1860
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
      Left            =   405
      TabIndex        =   18
      Top             =   1860
      Width           =   3600
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
      Left            =   7170
      TabIndex        =   17
      Top             =   1860
      Width           =   600
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   1
      X1              =   195
      X2              =   11160
      Y1              =   1815
      Y2              =   1815
   End
   Begin VB.Label H1_Ymd 
      Caption         =   "実績日"
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
      TabIndex        =   16
      Top             =   1515
      Width           =   810
   End
   Begin VB.Label H1_Suryo 
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
      Left            =   10365
      TabIndex        =   15
      Top             =   1515
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
      Left            =   9555
      TabIndex        =   14
      Top             =   1515
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
      Left            =   8325
      TabIndex        =   13
      Top             =   1515
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
      Left            =   4080
      TabIndex        =   12
      Top             =   1515
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
      Left            =   840
      TabIndex        =   11
      Top             =   1515
      Width           =   990
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
      TabIndex        =   10
      Top             =   480
      Width           =   465
   End
   Begin VB.Label Label1 
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
      Index           =   3
      Left            =   7080
      TabIndex        =   9
      Top             =   1515
      Width           =   795
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   0
      X1              =   195
      X2              =   11160
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
      Index           =   11
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
Attribute VB_Name = "CXH0510L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  入出庫・支給明細書　発行
'**       フォームID    :  CXH0510L
'**       処理概要      :
'**
'**       作  成  日    :  2004/02/19
'**       変  更  日    :
'**
'**********************************************************'
'   #比較退避用#
    Dim Key_Torcd       As String       '>実績先
    Dim Key_Hinbn       As String       '>品番
    Dim GMaisu          As Long         ' 箱数　合計
    Dim GSuryo          As Long         ' 実績数合計
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
    With CisPrtCXH0510
        Select Case JKT.実績区分
            Case 1: .iPrtTitle = "入庫明細書"
            Case 2: .iPrtTitle = "出庫明細書"
            Case 3: .iPrtTitle = "支給明細書"
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
    Key_Torcd = JKT.実績先
    Key_Hinbn = JKT.品番
    GMaisu = 0
    GSuryo = 0
    GKensu = 0

    B1_K_LINE1.Visible = False
    B1_K_LINE2.Visible = False
    Call CisPrtCXH0510.ClrLstFld("B1")

'*---------------------------------*
'*  実績管理テーブル読込み
'*---------------------------------*
    Do Until Not JKT_RDSTS
        If Key_Torcd <> JKT.実績先 Or _
           Key_Hinbn <> JKT.品番 Then
        '   合計印刷
            Call Total_Print
        '   改頁判定
            If Key_Torcd <> JKT.実績先 Then
                CisPrtCXH0510.cLine = CisPrtCXH0510.iLineMax
            End If
        End If

        Call HeadWrite(0)

        '*---------------------------------*
        '*  Body行　編集
        '*---------------------------------*
        With JKT
            If 0 = GKensu Or _
               0 = CisPrtCXH0510.cLine = 0 Then
                B1_Hinbn = RTrim(.表示品番)
                B1_Seban = RTrim(.背番号)
            End If
            If RTrim(.実績便) <> "" Then
                B1_Ymd = Format(.実績日, "0000/00/00") & "-" & .実績便
            Else
                B1_Ymd = Format(.実績日, "0000/00/00") & "   "
            End If
            B1_Nouno = RTrim(.納入番号)
            B1_Syuyo = Format(.収容数, "##,###")
            B1_Maisu = Format(.実績枚数, "##,###")
            B1_Suryo = Format(.実績数, "##,###")
    '       合計計算
            GMaisu = GMaisu + .実績枚数
            GSuryo = GSuryo + .実績数
            GKensu = GKensu + 1
    '       KEY項目退避
            Key_Torcd = .実績先
            Key_Hinbn = .品番
        End With
        '*---------------------------------*
        '*  Body部　印刷
        '*---------------------------------*
        With CisPrtCXH0510
            .PrintFld ("B1*")
            .cLine = .cLine + 1
            .ClrLstFld ("B1")
    '       カウンタ　ＵＰ
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With

        Call JKTReadNext(1)
    Loop
    Call JKTClose(1)

'   合計印刷
    Call Total_Print
'   印刷終了
    Call CisPrtCXH0510.PrintEndDoc

End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim TtlName     As String
    Dim YmdName     As String

    HeadWrite = False
    
    With CisPrtCXH0510
    '   １頁内印刷行数ﾁｪｯｸ
        If .iLineMax >= .cLine + 1 Then Exit Function

        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")   'HEAD部内容消去

    '   明細書ﾀｲﾄﾙ
        Select Case JKT.実績区分
            Case 1: TtlName = "入庫": YmdName = "入庫"
            Case 2: TtlName = "出庫": YmdName = "出庫"
            Case 3: TtlName = "支給": YmdName = "納入"
        End Select
        H1_JKbn = "※　" & TtlName & "明細書　※"
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
        If RTrim(JKT.工場CD) <> "" Then
            H1_Torcd = JKT.取引先 & " - " & RTrim(JKT.工場CD)
        Else
            H1_Torcd = JKT.取引先
        End If
        H1_Tornm = RTrim(JKT.実績先名)
        H1_Torkj = RTrim(JKT.工場名)

        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With

    HeadWrite = True
End Function
Private Function Total_Print()
    Total_Print = False
'
    If GKensu > 1 Then
        B1_Hinbn = "-----<< 品 番 合 計 >>-----"
        B1_Maisu = Format(GMaisu, "##,##0")
        B1_Suryo = Format(GSuryo, "##,##0")
        Call CisPrtCXH0510.PrintFld("B1*")
        CisPrtCXH0510.cLine = CisPrtCXH0510.cLine + 2
        Call CisPrtCXH0510.ClrLstFld("B1")
    Else
        CisPrtCXH0510.cLine = CisPrtCXH0510.cLine + 1
    End If
    GMaisu = 0
    GSuryo = 0
    GKensu = 0
'
    Total_Print = True
End Function
