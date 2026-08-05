VERSION 5.00
Begin VB.Form CXHZ0250L 
   Caption         =   "材料入荷一覧表"
   ClientHeight    =   2625
   ClientLeft      =   3375
   ClientTop       =   1545
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2625
   ScaleWidth      =   11565
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "番号"
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
      Left            =   6840
      TabIndex        =   25
      Top             =   1515
      Width           =   390
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
      Left            =   9690
      TabIndex        =   24
      Top             =   1860
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
      Left            =   9660
      TabIndex        =   23
      Top             =   1515
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
      Left            =   10740
      TabIndex        =   22
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
      Index           =   3
      Left            =   11100
      TabIndex        =   21
      Top             =   120
      Width           =   225
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
      Left            =   5340
      TabIndex        =   20
      Top             =   1860
      Width           =   1365
   End
   Begin VB.Label B1_Suryo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "###,###.000"
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
      Left            =   8385
      TabIndex        =   19
      Top             =   1860
      Width           =   1155
   End
   Begin VB.Label B1_Hasu 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "###,###.000"
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
      Left            =   10065
      TabIndex        =   18
      Top             =   1860
      Width           =   1155
   End
   Begin VB.Label B1_Keicd 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
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
      Left            =   7485
      TabIndex        =   17
      Top             =   1860
      Width           =   735
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
      Left            =   2850
      TabIndex        =   16
      Top             =   1860
      Width           =   2415
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
      TabIndex        =   15
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
      Left            =   6840
      TabIndex        =   14
      Top             =   1860
      Width           =   525
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   11310
      Y1              =   1815
      Y2              =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "入荷日"
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
      Left            =   5640
      TabIndex        =   13
      Top             =   1515
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "入荷数"
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
      Left            =   8910
      TabIndex        =   12
      Top             =   1515
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "入荷端数"
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
      Left            =   10440
      TabIndex        =   11
      Top             =   1515
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "契約先"
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
      Left            =   7485
      TabIndex        =   10
      Top             =   1515
      Width           =   585
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
      Left            =   3600
      TabIndex        =   9
      Top             =   1515
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
      Left            =   600
      TabIndex        =   8
      Top             =   1515
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "納入"
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
      Left            =   6840
      TabIndex        =   7
      Top             =   1290
      Width           =   390
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
      Index           =   4
      Left            =   225
      TabIndex        =   6
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
      Index           =   2
      Left            =   8160
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   90
      Width           =   1680
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   195
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "※　材料入荷一覧表　※"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   315
      Width           =   4140
   End
End
Attribute VB_Name = "CXHZ0250L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  材料入荷一覧表　発行
'**       フォームID    :  CXHZ0250L
'**       処理概要      :
'**
'**       作  成  日    :  2004/02/21
'**       変  更  日    :
'**
'**********************************************************'
'   #比較退避用#
    Dim Key_Torcd       As String       '>仕入先
    Dim Key_Zkno        As Long         '>材料管理番号
    Dim GSuryo          As Currency     ' 入荷数　合計
    Dim GHasu           As Currency     ' 入荷端数合計
    Dim GKensu          As Long         ' 件数

    Dim mPrtDate        As String
    Dim mPrtTime        As String
'*------------------------------------*
'*      Ｆｏｒｍ＿Ｌｏａｄ イベント     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")             '発行日
    mPrtTime = Format(Now(), "hh:nn")                  '発行時間
    '*---------------------*
    '*  プリント系初期設定
    '*---------------------*
    With CisPrtCXHZ0250
        .iPrtTitle = "材料入荷一覧表"
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
'*  初期設定
'*---------------------------------*
    Key_Torcd = ZYT.仕入先
    Key_Zkno = ZYT.材料管理番号
    GSuryo = 0
    GHasu = 0
    GKensu = 0

    Call CisPrtCXHZ0250.ClrLstFld("B1")

'*---------------------------------*
'*  入荷テーブル読込み
'*---------------------------------*
    Do Until Not ZYT_RDSTS
        If Key_Torcd <> ZYT.仕入先 Or _
           Key_Zkno <> ZYT.材料管理番号 Then
        '   合計印刷
            Call Total_Print
        '   改頁判定
            If Key_Torcd <> ZYT.仕入先 Then
                CisPrtCXHZ0250.cLine = CisPrtCXHZ0250.iLineMax
            End If
        End If

        Call HeadWrite(0)

        '*---------------------------------*
        '*  Body行　編集
        '*---------------------------------*
        With ZYT
            If 0 = GKensu Or _
               0 = CisPrtCXHZ0250.cLine Then
                B1_Zaist = RTrim(.材質)
                B1_Sunpo = RTrim(.表示寸法)
                Call MeisyoGet("単位区分", .単位)
            End If
            If .入荷便 = "" Then
                B1_Ymd = Format(.入荷日, "0000/00/00")
            Else
                B1_Ymd = Format(.入荷日, "0000/00/00") & "-" & Format(.入荷便, "00")
            End If
            B1_Nouno = RTrim(.納入番号)
            B1_Keicd = RTrim(.契約先)
            B1_Suryo = Format(.入荷数, "##,##0.000")
            B1_Tannm = SYM_Meisyo
            If .入荷端数 <> 0 Then
                B1_Hasu = Format(.入荷端数, "##,##0.000")
            End If
    '       合計計算
            GSuryo = GSuryo + .入荷数
            GHasu = GHasu + .入荷端数
            GKensu = GKensu + 1
    '       KEY項目退避
            Key_Torcd = .仕入先
            Key_Zkno = .材料管理番号
        End With
        '*---------------------------------*
        '*  Body部　印刷
        '*---------------------------------*
        With CisPrtCXHZ0250
            .PrintFld ("B1*")
            .cLine = .cLine + 1
            .ClrLstFld ("B1")
    '       カウンタ　ＵＰ
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With

        Call ZYTReadNext(1)
    Loop
    Call ZYTClose(1)

'   合計印刷
    Call Total_Print
'   印刷終了
    Call CisPrtCXHZ0250.PrintEndDoc

End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim TtlName     As String
    Dim YmdName     As String

    HeadWrite = False
    
    With CisPrtCXHZ0250
    '   １頁内印刷行数ﾁｪｯｸ
        If .iLineMax >= .cLine + 1 Then Exit Function

        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")   'HEAD部内容消去

    '   発行日・頁
        H1_HYmd = mPrtDate & " " & mPrtTime
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")
    '   仕入先
        If RTrim(ZYT.工場CD) <> "" Then
            H1_Torcd = ZYT.取引先 & " - " & RTrim(ZYT.工場CD)
        Else
            H1_Torcd = ZYT.取引先
        End If
        H1_Tornm = RTrim(ZYT.仕入先名)
        H1_Torkj = RTrim(ZYT.工場名)

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
        B1_Suryo = Format(GSuryo, "##,##0.000")
        B1_Hasu = Format(GHasu, "##,##0.000")
        Call CisPrtCXHZ0250.PrintFld("B1*")
        CisPrtCXHZ0250.cLine = CisPrtCXHZ0250.cLine + 2
        Call CisPrtCXHZ0250.ClrLstFld("B1")
    Else
        CisPrtCXHZ0250.cLine = CisPrtCXHZ0250.cLine + 1
    End If
    GSuryo = 0
    GHasu = 0
    GKensu = 0
'
    Total_Print = True
End Function

