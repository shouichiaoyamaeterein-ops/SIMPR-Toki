VERSION 5.00
Begin VB.Form CNJ1130L 
   Caption         =   "明細書"
   ClientHeight    =   9405
   ClientLeft      =   6255
   ClientTop       =   1350
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   11700
   Begin VB.Label Label1 
      Caption         =   "当月所要量"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   9840
      TabIndex        =   20
      Top             =   1305
      Width           =   1425
   End
   Begin VB.Label B1_Yotei 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "###,##0"
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
      Left            =   10185
      TabIndex        =   19
      Top             =   1710
      Width           =   840
   End
   Begin VB.Label Label4 
      Caption         =   "品　　名"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4935
      TabIndex        =   18
      Top             =   1305
      Width           =   990
   End
   Begin VB.Label B1_Hinnm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
      Left            =   3945
      TabIndex        =   17
      Top             =   1710
      Width           =   3600
   End
   Begin VB.Label H1_YY 
      AutoSize        =   -1  'True
      Caption         =   "YYYY"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5310
      TabIndex        =   16
      Top             =   615
      Width           =   600
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "年"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   6000
      TabIndex        =   15
      Top             =   615
      Width           =   285
   End
   Begin VB.Label H1_MM 
      AutoSize        =   -1  'True
      Caption         =   "MM"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6390
      TabIndex        =   14
      Top             =   615
      Width           =   300
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "月度"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   6720
      TabIndex        =   13
      Top             =   615
      Width           =   570
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
      Left            =   9915
      TabIndex        =   12
      Top             =   495
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
      Left            =   10890
      TabIndex        =   11
      Top             =   510
      Width           =   225
   End
   Begin VB.Line H1_K_LINE 
      Index           =   5
      X1              =   3885
      X2              =   3885
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
      X2              =   3900
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line H1_K_LINE 
      Index           =   2
      X1              =   180
      X2              =   3900
      Y1              =   390
      Y2              =   390
   End
   Begin VB.Label B1_Mikomi 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "###,##0"
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
      Left            =   8580
      TabIndex        =   10
      Top             =   1710
      Width           =   840
   End
   Begin VB.Label B1_Hinbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
      Left            =   270
      TabIndex        =   9
      Top             =   1710
      Width           =   3600
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   1
      X1              =   195
      X2              =   11160
      Y1              =   1635
      Y2              =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "受注実績見込"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   8025
      TabIndex        =   8
      Top             =   1305
      Width           =   1545
   End
   Begin VB.Label Label4 
      Caption         =   "品    番"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1260
      TabIndex        =   7
      Top             =   1305
      Width           =   990
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   0
      X1              =   195
      X2              =   11160
      Y1              =   1185
      Y2              =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "手配先"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   105
      Width           =   750
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
      Left            =   8835
      TabIndex        =   5
      Top             =   255
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
      Left            =   9465
      TabIndex        =   4
      Top             =   225
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
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
      Width           =   3600
   End
   Begin VB.Label H1_Torcd 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
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
      Left            =   1005
      TabIndex        =   1
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label4 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "※　工事手配書　※"
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
      Left            =   4605
      TabIndex        =   0
      Top             =   225
      Width           =   3390
   End
End
Attribute VB_Name = "CNJ1130L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  工事手配書　発行
'**       フォームID    :  CNJ1130L
'**       処理概要      :
'**
'**       作  成  日    :  2010/01/17
'**       変  更  日    :  2010/11/11  品番取引先マスタから出力,計画年月入力
'**
'**********************************************************'
'   #比較退避用#
    Dim Key_Torcd       As String       '>仕入先
    Dim Key_Hinbn       As String       '>品番
    Dim GSuryo          As Long         ' 所要量　合計
    Dim GJSuryo         As Long         ' 受注実績見込 合計
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
    With CisPrtCNJ1130
        .iPrtTitle = "工事手配書"
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
    Key_Torcd = KJT.手配先
    Key_Hinbn = KJT.品番
    GSuryo = 0
    GJSuryo = 0
    GKensu = 0

    Call CisPrtCNJ1130.ClrLstFld("B1")

'*---------------------------------*
'*  工事手配書ワーク読込み
'*---------------------------------*
    Do Until Not KJT_RDSTS
        If Key_Torcd <> KJT.手配先 Then
        '   合計印刷
            Call Total_Print
        '   改頁判定
            If Key_Torcd <> KJT.手配先 Then
                CisPrtCNJ1130.cLine = CisPrtCNJ1130.iLineMax
            End If
        End If

        Call HeadWrite(0)

        '*---------------------------------*
        '*  Body行　編集
        '*---------------------------------*
        With KJT
'            If 0 = GKensu Or _
'               0 = CisPrtCNJ1130.cLine Then
'                B1_Hinbn = RTrim(.表示品番)
'                B1_Hinnm = CisFun.Mid2(RTrim(.品名), 1, 30)
'            End If
            B1_Hinbn = RTrim(.表示品番)
            B1_Hinnm = CisFun.Mid2(RTrim(.品名), 1, 30)
            B1_Mikomi = Format(.受注実績見込, "###,###")
            B1_Yotei = Format(.当月所要量, "###,###")
    '       合計計算
            GSuryo = GSuryo + .当月所要量
            GJSuryo = GJSuryo + .受注実績見込
            GKensu = GKensu + 1
    '       KEY項目退避
            Key_Torcd = .手配先
        End With
        '*---------------------------------*
        '*  Body部　印刷
        '*---------------------------------*
        With CisPrtCNJ1130
            .PrintFld ("B1*")
            .cLine = .cLine + 1
            .ClrLstFld ("B1")
    '       カウンタ　ＵＰ
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With

        Call KJTReadNext
    Loop
    Call KJTClose

'   合計印刷
    Call Total_Print
'   印刷終了
    Call CisPrtCNJ1130.PrintEndDoc

End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim TtlName     As String
    Dim YmdName     As String

    HeadWrite = False
    
    With CisPrtCNJ1130
    '   １頁内印刷行数ﾁｪｯｸ
        If .iLineMax >= .cLine + 1 Then Exit Function

        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")   'HEAD部内容消去
    '   対象年月
'        H1_YY = Mid(CNJ1130.iH1_YM, 1, 4)                  '2010/11/11
'        H1_MM = Mid(CNJ1130.iH1_YM, 5, 2)                  '2010/11/11
        H1_YY = Mid(CNJ1130.iH1_NYM, 1, 4)                  '2010/11/11
        H1_MM = Mid(CNJ1130.iH1_NYM, 5, 2)                  '2010/11/11

    '   発行日・頁
        H1_HYmd = mPrtDate & " " & mPrtTime
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")
    '   仕入先
        H1_Torcd = KJT.手配先
        H1_Tornm = RTrim(KJT.手配先名)
        H1_Torkj = RTrim(KJT.工場名)

        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With

    HeadWrite = True
End Function
Private Function Total_Print()
    Total_Print = False
'
    If GKensu > 1 Then
        B1_Hinnm = "-----<< 手配先合計 >>-----"
        B1_Yotei = Format(GSuryo, "###,##0")
        B1_Mikomi = Format(GJSuryo, "###,##0")
        Call CisPrtCNJ1130.PrintFld("B1*")
        CisPrtCNJ1130.cLine = CisPrtCNJ1130.cLine + 2
        Call CisPrtCNJ1130.ClrLstFld("B1")
    Else
        CisPrtCNJ1130.cLine = CisPrtCNJ1130.cLine + 1
    End If
    GSuryo = 0
    GJSuryo = 0
    GKensu = 0
'
    Total_Print = True
End Function


