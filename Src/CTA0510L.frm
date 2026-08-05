VERSION 5.00
Begin VB.Form CTA0510L 
   Caption         =   "棚卸HHT読取結果取込エラーリスト"
   ClientHeight    =   5940
   ClientLeft      =   1380
   ClientTop       =   1710
   ClientWidth     =   16920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   16920
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
      Left            =   16065
      TabIndex        =   14
      Top             =   180
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
      Left            =   16425
      TabIndex        =   13
      Top             =   195
      Width           =   225
   End
   Begin VB.Label B1_Gappi 
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99"
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
      TabIndex        =   12
      Top             =   1860
      Width           =   1200
   End
   Begin VB.Label B1_Naiyo 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4XXXXXXXXX5XXXXXXXXX6XXXXXXXXX7XXXXXX"
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
      Left            =   7485
      TabIndex        =   11
      Top             =   1860
      Width           =   9120
   End
   Begin VB.Label B1_Renbn 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999999"
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
      TabIndex        =   10
      Top             =   1860
      Width           =   720
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
      Left            =   300
      TabIndex        =   9
      Top             =   1860
      Width           =   3600
   End
   Begin VB.Label B1_Jikan 
      AutoSize        =   -1  'True
      Caption         =   "hh:mm:ss"
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
      Left            =   6255
      TabIndex        =   8
      Top             =   1860
      Width           =   960
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   1
      X1              =   195
      X2              =   16650
      Y1              =   1810
      Y2              =   1810
   End
   Begin VB.Label Label2 
      Caption         =   "処理日"
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
      Left            =   5295
      TabIndex        =   7
      Top             =   1515
      Width           =   810
   End
   Begin VB.Label Label2 
      Caption         =   "読取内容"
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
      Left            =   11430
      TabIndex        =   6
      Top             =   1515
      Width           =   1710
   End
   Begin VB.Label Label2 
      Caption         =   "連番"
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
      Left            =   4230
      TabIndex        =   5
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
      Left            =   1635
      TabIndex        =   4
      Top             =   1515
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "時間"
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
      Left            =   6495
      TabIndex        =   3
      Top             =   1515
      Width           =   840
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
      Left            =   13485
      TabIndex        =   2
      Top             =   180
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
      Left            =   14115
      TabIndex        =   1
      Top             =   180
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "【　棚卸HHT読取結果取込エラーリスト　】　"
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
      Left            =   4965
      TabIndex        =   0
      Top             =   420
      Width           =   7710
   End
End
Attribute VB_Name = "CTA0510L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  棚卸ＨＨＴ読取結果取込エラーリスト　発行
'**       フォームID    :  CTA0510L  (Based CXH0250L:入荷一覧表　発行)
'**       処理概要      :
'**
'**       作  成  日    :  2019/07/05
'**       変  更  日    :
'**
'**********************************************************'
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 DELETE START
''   #比較退避用#
'    Dim Key_Torcd       As String       '>仕入先
'    Dim Key_Hinbn       As String       '>品番
'    Dim GSuryo          As Long         ' 入荷数　合計
'    Dim GHasu           As Long         ' 入荷端数合計
'    Dim GKensu          As Long         ' 件数
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 DELETE END

    Dim mPrtDate        As String
    Dim mPrtTime        As String

    Dim CellsData       As Variant      ' 2019/07/05 ADD

'*------------------------------------*
'*      Ｆｏｒｍ＿Ｌｏａｄ イベント     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")              ' 発行日
    mPrtTime = Format(Now(), "hh:nn")                   ' 発行時間
    '*---------------------*
    '*  プリント系初期設定
    '*---------------------*
    With CisPrtCTA0510
        .iPrtTitle = "棚卸ＨＴＴ読取結果取込エラーリスト"
        .iLineMax = 40                                  ' 2019/07/05 UPD：55
        .iLinePitch = 6
        .iPaperSize = p_A4                              ' 用紙サイズ
        .iOrientation = po_Landscape                    ' 用紙方向（横）
        
        If Not .PrintIni Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(1) = "     　プリンタ初期設定エラー        "
                .MB_Title = "印刷ｴﾗｰ"
                .MB_Button = Error
                .MBOX
            End With
            Exit Sub
        End If
        .cLine = 0
        .cPage = 0
    End With

    Call PrintProc

End Sub
'*---------------------------------------------------------------*
'*　 << 印  刷  処  理 >>
'*---------------------------------------------------------------*
Private Function PrintProc()
    Dim iLoop   As Long
'*---------------------------------*
'*  初期設定
'*---------------------------------*
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 DELETE START
'    Key_Torcd = NYT.仕入先
'    Key_Hinbn = NYT.品番
'    GSuryo = 0
'    GHasu = 0
'    GKensu = 0
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 DELETE START

    Call CisPrtCTA0510.ClrLstFld("B1")

'*---------------------------------*
'*  エラーデータ読込み
'*---------------------------------*
    Call HeadWrite(1)                   ' 2019/07/08 ADD：初回ヘッダ作成用

    For iLoop = 1 To iErrorCnt          ' 2019/07/05 ADD
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 DELETE START
'    Do Until Not NYT_RDSTS
'        If Key_Torcd <> NYT.仕入先 Or _
'           Key_Hinbn <> NYT.品番 Then
'        '   合計印刷
'            Call Total_Print
'        '   改頁判定
'            If Key_Torcd <> NYT.仕入先 Then
'                CisPrtCTA0510.cLine = CisPrtCTA0510.iLineMax
'            End If
'        End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 DELETE END
                                                                    
        Call HeadWrite(0)


        '*---------------------------------*
        '*  Body行　編集
        '*---------------------------------*
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 DELETE START
'        With NYT
'            If 0 = GKensu Or _
'               0 = CisPrtCTA0510.cLine Then
'                B1_Hinbn = RTrim(.表示品番)
'                B1_Seban = RTrim(.背番号)
'            End If
'            If .入荷便 = "" Then
'                B1_Ymd = Format(.入荷日, "0000/00/00")
'            Else
'                B1_Ymd = Format(.入荷日, "0000/00/00") & "-" & Format(.入荷便, "00")
'            End If
'            B1_Nouno = RTrim(.納入番号)
'            B1_Keicd = RTrim(.契約先)
'            B1_Suryo = Format(.入荷数, "##,###")
'            B1_Hasu = Format(.入荷端数, "##,###")
'    '       合計計算
'            GSuryo = GSuryo + .入荷数
'            GHasu = GHasu + .入荷端数
'            GKensu = GKensu + 1
'    '       KEY項目退避
'            Key_Torcd = .仕入先
'            Key_Hinbn = .品番
'        End With
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 DELETE END

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 INSERT START
        CellsData = Split(wErrorTbl(iLoop), ",")                    ' カンマ区切りで列データを分割
        B1_Hinbn = CisFun.Mid2(RTrim(CStr(CellsData(7))), 1, 30)    ' CellsData(7)
        B1_Renbn = CisFun.Mid2(RTrim(CStr(CellsData(5))), 1, 10)    ' CellsData(5)
        B1_Gappi = CisFun.Mid2(RTrim(CStr(CellsData(0))), 1, 10)    ' CellsData(0)
        B1_Jikan = CisFun.Mid2(RTrim(CStr(CellsData(1))), 1, 8)     ' CellsData(1)
        B1_Naiyo = CisFun.Mid2(RTrim(CStr(CellsData(6))), 1, 76)    ' CellsData(6)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 INSERT END
        
        '*---------------------------------*
        '*  Body部　印刷
        '*---------------------------------*
        With CisPrtCTA0510
            .PrintFld ("B1*")
            .cLine = .cLine + 1
            .ClrLstFld ("B1")
    '       カウンタ　ＵＰ
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 DELETE START
'        Call NYTReadNext(1)
'    Loop
'    Call NYTClose(1)
''   合計印刷
'    Call Total_Print
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 DELETE END
    Next iLoop                          ' 2019/07/05 ADD

'   印刷終了
    Call CisPrtCTA0510.PrintEndDoc

End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim TtlName     As String
    Dim YmdName     As String

    HeadWrite = False
    
    With CisPrtCTA0510
    '   １頁内印刷行数ﾁｪｯｸ
      ''If .iLineMax >= .cLine + 1 Then Exit Function                       ' 2019/07/08 DEL
        If KB = 0 And .iLineMax >= .cLine + 1 Then Exit Function            ' 2019/07/08 ADD

        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")   'HEAD部内容消去

    '   発行日・頁
        H1_HYmd = mPrtDate & " " & mPrtTime
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")
'    '   仕入先
'        If RTrim(NYT.工場CD) <> "" Then
'            H1_Torcd = NYT.取引先 & " - " & RTrim(NYT.工場CD)
'        Else
'            H1_Torcd = NYT.取引先
'        End If
'        H1_Tornm = RTrim(NYT.仕入先名)
'        H1_Torkj = RTrim(NYT.工場名)
'
        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With

    HeadWrite = True
End Function
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 DELETE START
'Private Function Total_Print()
'    Total_Print = False
''
'    If GKensu > 1 Then
'        B1_Hinbn = "-----<< 品 番 合 計 >>-----"
'        B1_Suryo = Format(GSuryo, "##,##0")
'        B1_Hasu = Format(GHasu, "##,##0")
'        Call CisPrtCTA0510.PrintFld("B1*")
'        CisPrtCTA0510.cLine = CisPrtCTA0510.cLine + 2
'        Call CisPrtCTA0510.ClrLstFld("B1")
'    Else
'        CisPrtCTA0510.cLine = CisPrtCTA0510.cLine + 1
'    End If
'    GSuryo = 0
'    GHasu = 0
'    GKensu = 0
''
'    Total_Print = True
'End Function
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2019/07/05 DELETE END


