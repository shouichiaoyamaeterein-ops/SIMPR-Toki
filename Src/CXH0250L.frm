VERSION 5.00
Begin VB.Form CXH0250L 
   Caption         =   "入荷一覧表"
   ClientHeight    =   5940
   ClientLeft      =   1380
   ClientTop       =   1710
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
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5340
      TabIndex        =   20
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
      Left            =   8940
      TabIndex        =   19
      Top             =   1860
      Width           =   840
   End
   Begin VB.Label B1_Hasu 
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
      Left            =   10080
      TabIndex        =   18
      Top             =   1860
      Width           =   840
   End
   Begin VB.Label B1_Keicd 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
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
      Left            =   7845
      TabIndex        =   17
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
      Left            =   4200
      TabIndex        =   16
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
      TabIndex        =   15
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
      Left            =   7110
      TabIndex        =   14
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
   Begin VB.Label Label2 
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
      Left            =   5745
      TabIndex        =   13
      Top             =   1515
      Width           =   810
   End
   Begin VB.Label Label2 
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
      Index           =   6
      Left            =   9165
      TabIndex        =   12
      Top             =   1515
      Width           =   720
   End
   Begin VB.Label Label2 
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
      Index           =   7
      Left            =   10155
      TabIndex        =   11
      Top             =   1515
      Width           =   795
   End
   Begin VB.Label Label2 
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
      Index           =   5
      Left            =   7965
      TabIndex        =   10
      Top             =   1515
      Width           =   630
   End
   Begin VB.Label Label2 
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
      Index           =   2
      Left            =   4200
      TabIndex        =   9
      Top             =   1515
      Width           =   630
   End
   Begin VB.Label Label2 
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
      Index           =   1
      Left            =   840
      TabIndex        =   8
      Top             =   1515
      Width           =   990
   End
   Begin VB.Label Label2 
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
      Index           =   4
      Left            =   7080
      TabIndex        =   7
      Top             =   1515
      Width           =   645
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
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
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
      Caption         =   "※　入荷一覧表　※"
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
      Left            =   3855
      TabIndex        =   0
      Top             =   315
      Width           =   3390
   End
End
Attribute VB_Name = "CXH0250L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  入荷一覧表　発行
'**       フォームID    :  CXH0250L
'**       処理概要      :
'**
'**       作  成  日    :  2004/02/20
'**       変  更  日    :
'**
'**********************************************************'
'   #比較退避用#
    Dim Key_Torcd       As String       '>仕入先
    Dim Key_Hinbn       As String       '>品番
    Dim GSuryo          As Long         ' 入荷数　合計
    Dim GHasu           As Long         ' 入荷端数合計
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
    With CisPrtCXH0250
        .iPrtTitle = "入荷一覧表"
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
    Key_Torcd = NYT.仕入先
    Key_Hinbn = NYT.品番
    GSuryo = 0
    GHasu = 0
    GKensu = 0

    Call CisPrtCXH0250.ClrLstFld("B1")

'*---------------------------------*
'*  入荷テーブル読込み
'*---------------------------------*
    Do Until Not NYT_RDSTS
        If Key_Torcd <> NYT.仕入先 Or _
           Key_Hinbn <> NYT.品番 Then
        '   合計印刷
            Call Total_Print
        '   改頁判定
            If Key_Torcd <> NYT.仕入先 Then
                CisPrtCXH0250.cLine = CisPrtCXH0250.iLineMax
            End If
        End If

        Call HeadWrite(0)

        '*---------------------------------*
        '*  Body行　編集
        '*---------------------------------*
        With NYT
            If 0 = GKensu Or _
               0 = CisPrtCXH0250.cLine Then
                B1_Hinbn = RTrim(.表示品番)
                B1_Seban = RTrim(.背番号)
            End If
            If .入荷便 = "" Then
                B1_Ymd = Format(.入荷日, "0000/00/00")
            Else
                B1_Ymd = Format(.入荷日, "0000/00/00") & "-" & Format(.入荷便, "00")
            End If
            B1_Nouno = RTrim(.納入番号)
            B1_Keicd = RTrim(.契約先)
            B1_Suryo = Format(.入荷数, "##,###")
            B1_Hasu = Format(.入荷端数, "##,###")
    '       合計計算
            GSuryo = GSuryo + .入荷数
            GHasu = GHasu + .入荷端数
            GKensu = GKensu + 1
    '       KEY項目退避
            Key_Torcd = .仕入先
            Key_Hinbn = .品番
        End With
        '*---------------------------------*
        '*  Body部　印刷
        '*---------------------------------*
        With CisPrtCXH0250
            .PrintFld ("B1*")
            .cLine = .cLine + 1
            .ClrLstFld ("B1")
    '       カウンタ　ＵＰ
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With

        Call NYTReadNext(1)
    Loop
    Call NYTClose(1)

'   合計印刷
    Call Total_Print
'   印刷終了
    Call CisPrtCXH0250.PrintEndDoc

End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim TtlName     As String
    Dim YmdName     As String

    HeadWrite = False
    
    With CisPrtCXH0250
    '   １頁内印刷行数ﾁｪｯｸ
        If .iLineMax >= .cLine + 1 Then Exit Function

        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")   'HEAD部内容消去

    '   発行日・頁
        H1_HYmd = mPrtDate & " " & mPrtTime
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")
    '   仕入先
        If RTrim(NYT.工場CD) <> "" Then
            H1_Torcd = NYT.取引先 & " - " & RTrim(NYT.工場CD)
        Else
            H1_Torcd = NYT.取引先
        End If
        H1_Tornm = RTrim(NYT.仕入先名)
        H1_Torkj = RTrim(NYT.工場名)

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
        B1_Suryo = Format(GSuryo, "##,##0")
        B1_Hasu = Format(GHasu, "##,##0")
        Call CisPrtCXH0250.PrintFld("B1*")
        CisPrtCXH0250.cLine = CisPrtCXH0250.cLine + 2
        Call CisPrtCXH0250.ClrLstFld("B1")
    Else
        CisPrtCXH0250.cLine = CisPrtCXH0250.cLine + 1
    End If
    GSuryo = 0
    GHasu = 0
    GKensu = 0
'
    Total_Print = True
End Function
