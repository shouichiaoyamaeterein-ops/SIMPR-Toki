VERSION 5.00
Begin VB.Form CNJ0510L 
   Caption         =   "かんばん増減リスト"
   ClientHeight    =   2745
   ClientLeft      =   105
   ClientTop       =   1530
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   11685
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
      Left            =   8235
      TabIndex        =   23
      Top             =   405
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "仕入先："
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
      Index           =   11
      Left            =   225
      TabIndex        =   22
      Top             =   870
      Width           =   960
   End
   Begin VB.Label H1_Tornm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
      Left            =   2085
      TabIndex        =   21
      Top             =   870
      Width           =   2400
   End
   Begin VB.Label H1_Torcd 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
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
      Left            =   1215
      TabIndex        =   20
      Top             =   870
      Width           =   840
   End
   Begin VB.Line H1_K_LINE 
      X1              =   180
      X2              =   10800
      Y1              =   1800
      Y2              =   1800
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
      Left            =   3945
      TabIndex        =   19
      Top             =   300
      Width           =   570
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
      Left            =   3615
      TabIndex        =   18
      Top             =   300
      Width           =   300
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
      Left            =   3225
      TabIndex        =   17
      Top             =   300
      Width           =   285
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
      Left            =   2535
      TabIndex        =   16
      Top             =   300
      Width           =   600
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
      Left            =   210
      TabIndex        =   15
      Top             =   1920
      Width           =   960
   End
   Begin VB.Label H1_HYmd 
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
      Left            =   9405
      TabIndex        =   14
      Top             =   405
      Width           =   1920
   End
   Begin VB.Label H1_Page 
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
      Left            =   10170
      TabIndex        =   13
      Top             =   795
      Width           =   600
   End
   Begin VB.Label B1_Hinbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
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
      Left            =   1335
      TabIndex        =   12
      Top             =   1920
      Width           =   3000
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "前月枚数"
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
      Left            =   4995
      TabIndex        =   11
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label B1_ZenSu 
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
      Left            =   5430
      TabIndex        =   10
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label B1_TouSu 
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
      Left            =   7110
      TabIndex        =   9
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label B1_AddSu 
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
      Left            =   8700
      TabIndex        =   8
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label B1_KaiSu 
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
      Left            =   10200
      TabIndex        =   7
      Top             =   1905
      Width           =   360
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "かんばん増減リスト"
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
      TabIndex        =   6
      Top             =   285
      Width           =   2835
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "背番号"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1440
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
      Left            =   10815
      TabIndex        =   4
      Top             =   795
      Width           =   225
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "品　　番"
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
      Index           =   0
      Left            =   2310
      TabIndex        =   3
      Top             =   1455
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "当月枚数"
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
      Left            =   6600
      TabIndex        =   2
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "追加枚数"
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
      Left            =   8175
      TabIndex        =   1
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "回収枚数"
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
      Left            =   9615
      TabIndex        =   0
      Top             =   1440
      Width           =   960
   End
End
Attribute VB_Name = "CNJ0510L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  かんばん増減リスト
'**       フォームID    :  CNJ0510L
'**       処理概要      :
'**
'**       作  成  日    :  2004/01/06
'**       変  更  日    :  2008/04/11 背番号8桁
'**
'**********************************************************'
'   # 比較退避用 #
    Dim Key_Torcd       As String       '>仕入先

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
    With CisPrtCNJ0510
        .iPrtTitle = "かんばん増減リスト"
        .iLineMax = 60
       '.iLinePitch = 8
        .iPaperSize = p_A4
        .iOrientation = po_Portrait

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
'*---------------------------------*
'*      初期設定
'*---------------------------------*
    Key_Torcd = KKT.仕入先

'*---------------------------------*
'*  かんばん回転枚数テーブル読込み
'*---------------------------------*
    Do Until Not KKT_RDSTS
        If Key_Torcd <> KKT.仕入先 Then
            Key_Torcd = KKT.仕入先
            CisPrtCNJ0510.cLine = CisPrtCNJ0510.iLineMax
        End If

        Call HeadWrite(0)

        With CisPrtCNJ0510
            .ClrLstFld ("B1")
        End With
        '*---------------------------------*
        '*  Body行　編集
        '*---------------------------------*
        With KKT
            B1_Seban = .背番号
            B1_Hinbn = .表示品番
            B1_TouSu = Format(.回転枚数, "#,##0")
            B1_ZenSu = Format(.前月回転枚数, "#,##0")
            Select Case .増減区分
                Case 1: B1_AddSu = Format(.増発枚数, "#,##0")
                Case 0: B1_KaiSu = Format(.抜取枚数, "#,##0")
            End Select
        End With
        '*---------------------------------*
        '*  Body部　印刷
        '*---------------------------------*
        With CisPrtCNJ0510
            .PrintFld ("B1*")
            .cLine = .cLine + 1
    '       カウンタ　ＵＰ
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With

        Call KKTReadNext(1)
    Loop
    Call KKTClose(1)

    ' 印刷終了
    Call CisPrtCNJ0510.PrintEndDoc

End Function
'*----------------*
'*  ヘッダ部印刷  *
'*----------------*
Private Function HeadWrite(KB As Byte)
    HeadWrite = False

    With CisPrtCNJ0510
    '   １頁内印刷行数ﾁｪｯｸ
        If .iLineMax >= .cLine + 1 Then Exit Function

        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")   'HEAD部内容消去

    '   発行日・頁
        H1_HYmd = mPrtDate & " " & mPrtTime
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")

        H1_YY = CisFun.Mid2(KKT.年月, 1, 4)
        H1_MM = CisFun.Mid2(KKT.年月, 5, 2)
        H1_Torcd = KKT.仕入先
        H1_Tornm = KKT.仕入先名

        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With

    HeadWrite = True
End Function
