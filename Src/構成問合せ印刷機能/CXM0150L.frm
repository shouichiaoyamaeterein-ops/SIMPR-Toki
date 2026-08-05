VERSION 5.00
Begin VB.Form CXM0150L 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "入荷一覧表"
   ClientHeight    =   5940
   ClientLeft      =   1365
   ClientTop       =   1695
   ClientWidth     =   16665
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   16665
   Begin VB.Label H1_JyuBiko 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
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
      Left            =   13575
      TabIndex        =   29
      Top             =   495
      Width           =   3210
   End
   Begin VB.Label B1_Syuyo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXXX"
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
      Left            =   12375
      TabIndex        =   28
      Top             =   1140
      Width           =   825
   End
   Begin VB.Label B1_Seban 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXX"
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
      Left            =   11385
      TabIndex        =   27
      Top             =   1140
      Width           =   840
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
      Index           =   10
      Left            =   11475
      TabIndex        =   26
      Top             =   780
      Width           =   705
   End
   Begin VB.Label Label2 
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
      Index           =   9
      Left            =   12480
      TabIndex        =   25
      Top             =   780
      Width           =   705
   End
   Begin VB.Label H1_Title2 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
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
      Left            =   315
      TabIndex        =   24
      Top             =   495
      Width           =   4170
   End
   Begin VB.Label H1_Title 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
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
      Left            =   315
      TabIndex        =   23
      Top             =   255
      Width           =   4170
   End
   Begin VB.Label B1_Tehai 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
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
      Left            =   990
      TabIndex        =   22
      Top             =   1140
      Width           =   2325
   End
   Begin VB.Label Label2 
      Caption         =   "手配先"
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
      Left            =   1545
      TabIndex        =   21
      Top             =   780
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "展開区分"
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
      Left            =   15180
      TabIndex        =   20
      Top             =   780
      Width           =   945
   End
   Begin VB.Label B1_Tenkai 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXXX"
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
      Left            =   15120
      TabIndex        =   19
      Top             =   1140
      Width           =   1080
   End
   Begin VB.Label B1_Gouho 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
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
      Left            =   14475
      TabIndex        =   18
      Top             =   1140
      Width           =   420
   End
   Begin VB.Label Label2 
      Caption         =   "号補区分"
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
      Left            =   14265
      TabIndex        =   17
      Top             =   780
      Width           =   870
   End
   Begin VB.Label Label2 
      Caption         =   "使用数"
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
      Left            =   13425
      TabIndex        =   16
      Top             =   780
      Width           =   690
   End
   Begin VB.Label B1_Siyosu 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXXX"
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
      Left            =   13515
      TabIndex        =   15
      Top             =   1140
      Width           =   750
   End
   Begin VB.Label Label2 
      Caption         =   "品名 / 材料名"
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
      Left            =   9315
      TabIndex        =   14
      Top             =   810
      Width           =   1410
   End
   Begin VB.Label B1_Hinnm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
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
      Left            =   8730
      TabIndex        =   13
      Top             =   1140
      Width           =   2625
   End
   Begin VB.Label B1_KB 
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
      Left            =   3540
      TabIndex        =   12
      Top             =   1140
      Width           =   225
   End
   Begin VB.Label B1_Level 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXX"
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
      Left            =   270
      TabIndex        =   11
      Top             =   1140
      Width           =   630
   End
   Begin VB.Label Label2 
      Caption         =   "寸　　法"
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
      Left            =   6555
      TabIndex        =   10
      Top             =   780
      Width           =   870
   End
   Begin VB.Label B1_Sunpou 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
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
      Left            =   6000
      TabIndex        =   9
      Top             =   1140
      Width           =   2625
   End
   Begin VB.Label Label2 
      Caption         =   "区分"
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
      Left            =   3465
      TabIndex        =   8
      Top             =   780
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "レベル"
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
      Left            =   270
      TabIndex        =   7
      Top             =   780
      Width           =   630
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
      Left            =   15945
      TabIndex        =   6
      Top             =   0
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
      Left            =   16305
      TabIndex        =   5
      Top             =   30
      Width           =   225
   End
   Begin VB.Label B1_Hinbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
      Left            =   3840
      TabIndex        =   4
      Top             =   1140
      Width           =   2100
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   1
      X1              =   210
      X2              =   16520
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "品番 / 材質"
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
      Left            =   4275
      TabIndex        =   3
      Top             =   780
      Width           =   1200
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
      Left            =   13365
      TabIndex        =   2
      Top             =   30
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
      Left            =   13995
      TabIndex        =   1
      Top             =   0
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "※　構成マスタ照会一覧　※"
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
      Left            =   5325
      TabIndex        =   0
      Top             =   225
      Width           =   4890
   End
End
Attribute VB_Name = "CXM0150L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  構成マスタ照会 一覧　発行
'**       フォームID    :  CXM0150L
'**       処理概要      :
'**
'**       作  成  日    :  2009/03/13
'**       変  更  日    :  2011/07/21 CIS 逆展開時：背番号・収容数印字追加
'**                                       逆展開時：品名の印字を25バイト分に制限
'**       変  更  日    :  2012/07/03 CIS 号補区分表示無
'**       変  更  日    :  2013/03/15  By CIS  CIS8606参照[手配先]取得
'**       変  更  日    :  2013/09/16  By CIS
'**       変  更  日    :   2014/02/13  By CIS  受注備考欄追加
'**       変  更  日    :   2015/01/21  By CIS  背番号桁数枠拡大(8桁)
'**       変  更  日    :   2016/11/23  By CIS  背番号表示変更
'**********************************************************'
'   #比較退避用#
    Dim Key_Torcd       As String       '>仕入先
    Dim Key_Hinbn       As String       '>品番
    Dim GSuryo          As Long         ' 入荷数　合計
    Dim GHasu           As Long         ' 入荷端数合計
    Dim GKensu          As Long         ' 件数

    Dim mPrtDate        As String
    Dim mPrtTime        As String

    Dim wFlg            As Boolean        '2013/03/15
'*------------------------------------*
'*      Ｆｏｒｍ＿Ｌｏａｄ イベント     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")             '発行日
    mPrtTime = Format(Now(), "hh:nn")                  '発行時間
    '*---------------------*
    '*  プリント系初期設定
    '*---------------------*
    With CisPrtCXM0150
        .iPrtTitle = "構成マスタ照会一覧"
        .iLineMax = 30
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

    Call CisPrtCXM0150.ClrLstFld("B1")

'*---------------------------------*
'*  入荷テーブル読込み
'*---------------------------------*
    Do Until Not KOL_RDSTS

        Call HeadWrite(0)

        '*---------------------------------*
        '*  Body行　編集
        '*---------------------------------*
        With KOL
            ' レベル
            If L_wsHKbn = 0 Then                '2009/11/10
                Select Case KOL.レベル
                    Case 0: B1_Level = "0"
                    Case 1: B1_Level = ".1"
                    Case 2: B1_Level = "..2"
                    Case 3: B1_Level = "...3"
                    Case 4: B1_Level = "....4"
                    Case Else
                        B1_Level = "....." & Format(KOL.レベル, "##")
                End Select
                If .手配先 <> "" Then
                   'B1_Tehai = RTrim(.手配先)
            '=========================== 2013/03/15 Start
'                    B1_Tehai = CisFun.Mid2(RTrim(.手配先), 1, 22) 'Update 2011/03/01
                    ' 手配先
                    If KOL.区分 = "" Then
                        gSL_Select = "SELECT 連番 FROM CIS8606"
                        '==================== 2013/09/16 Start
                        'gSL_Select = gSL_Select & " WHERE REPLACE(品番,'-','') = '" & Replace(RTrim(KOL.親品番), "-", "") & "'"
                        If RTrim(KOL.親品番) <> "" Then
                            gSL_Select = gSL_Select & " WHERE REPLACE(品番,'-','') = '" & Replace(RTrim(KOL.親品番), "-", "") & "'"
                        Else
                            gSL_Select = gSL_Select & " WHERE REPLACE(品番,'-','') = '" & Replace(RTrim(KOL.子品番), "-", "") & "'"
                        End If
                        '==================== 2013/09/16 End
                        gSL_Select = gSL_Select & " ORDER BY CONVERT(INT,連番) DESC "
                        If CISRead(gSL_Select, 1, 2) Then
                            gSL_Select = "SELECT TOP 100 * FROM CIS8606"
                            gSL_Select = gSL_Select & " WHERE CONVERT(INT,連番) >= " & CisFun.Val2(CIS.連番)
                            gSL_Select = gSL_Select & " ORDER BY CONVERT(INT,連番) "
                            If CISRead(gSL_Select, , 2) Then
                                wFlg = False
                                Do Until Not CIS_RDSTS
                                    If Replace(CIS.品番, "-", "") = Replace(KOL.子品番, "-", "") Then
                                        wFlg = True
                                        Call TorNmGet(RTrim(CIS.手配先コード) & RTrim(CIS.手配先工場コード), 9, 3)
                                        If RTrim(TRM.略称) <> "" Then
                                            'gSL_Select = "SELECT 取引先 FROM 品番取引先マスタ "                '2016/11/23 del
                                            gSL_Select = "SELECT 取引先,背番号 FROM 品番取引先マスタ "          '2016/11/23
                                            gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(KOL.子品番) & "'"
                                            gSL_Select = gSL_Select & " AND 取引先 = '" & RTrim(CIS.手配先コード) & RTrim(CIS.手配先工場コード) & "'"
                                            '================= 2016/11/23 start
                                            If RTrim(CIS.受入発注) <> "" Then
                                                gSL_Select = gSL_Select & " AND 受入 = '" & RTrim(CIS.受入発注) & "'"
                                            Else
                                                gSL_Select = gSL_Select & " AND 受入 = '01'"
                                            End If
                                            '================= 2016/11/23 end
                                            gSL_Select = gSL_Select & " AND 品目 = 1 "
                                            If HTMRead(gSL_Select, 1, 3) Then
                                                gStr = RTrim(CIS.手配先コード) & RTrim(CIS.手配先工場コード) & ":" & RTrim(TRM.略称)
                                                B1_Tehai = CisFun.Mid2(gStr, 1, 22)
                                    
                                                '================= 2016/11/23 start
                                                If RTrim(CIS.SD発注) <> "" Then
                                                    B1_Seban = RTrim(HTM.背番号)
                                                End If
                                                '================= 2016/11/23 end
                                
                                            Else
                                                B1_Tehai = RTrim(KOL.手配先)
                                            End If
                                        Else
                                            B1_Tehai = CisFun.Mid2(RTrim(KOL.手配先), 1, 22)
                                        End If
                                        Exit Do
                                    End If
                                    Call CISReadNext(2)
                                Loop
                                Call CISClose(2)
                                If wFlg = False Then
                                    B1_Tehai = RTrim(KOL.手配先)
                                End If
                            Else
                                B1_Tehai = RTrim(KOL.手配先)
                            End If
                        Else
                            B1_Tehai = RTrim(KOL.手配先)
                        End If
                    Else
                        B1_Tehai = RTrim(KOL.手配先)
                    End If
            '=========================== 2013/03/15 End
                
                End If
                If .区分 <> "" Then
                   B1_KB = "材"
                End If
                B1_Hinbn = CisFun.Mid2(RTrim(.表示品番), 1, 20) 'Update 2011/03/01
                'B1_Hinbn = RTrim(.表示品番)
                B1_Sunpou = CisFun.Mid2(RTrim(.表示寸法), 1, 25) 'Update 2011/03/01
                'B1_Sunpou = RTrim(.表示寸法)
                B1_Hinnm = CisFun.Mid2(RTrim(.品名), 1, 25) 'Update 2011/03/01
                'B1_Hinnm = RTrim(.品名)
                
                '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert Start
                If KOL.区分 = "" Then
                    gSL_Select = ""
                    gSL_Select = gSL_Select & "SELECT HT1.背番号"
                    gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(HT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= HT1.変更日 THEN HT1.収容数2 ELSE HT1.収容数1 END AS 収容数1"
                    gSL_Select = gSL_Select & "  FROM 品番取引先マスタ HT1"
                    gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                    gSL_Select = gSL_Select & "               FROM 品番取引先マスタ"
                    gSL_Select = gSL_Select & "              WHERE 取引先 = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "                AND 品番   = '" & KOL.子品番 & "') HT2"
                    gSL_Select = gSL_Select & "       ON  HT1.取引先 = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "       AND HT1.品番   = '" & KOL.子品番 & "'"
                    gSL_Select = gSL_Select & "       AND HT1.受入   = HT2.受入"
                Else
                    gSL_Select = ""
                    gSL_Select = gSL_Select & "SELECT ZT1.背番号"
                    gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(ZT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= ZT1.変更日 THEN ZT1.収容数2 ELSE ZT1.収容数1 END AS 収容数1"
                    gSL_Select = gSL_Select & "  FROM 材料取引先マスタ ZT1"
                    gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                    gSL_Select = gSL_Select & "               FROM 材料取引先マスタ"
                    gSL_Select = gSL_Select & "              WHERE 取引先       = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "                AND 材料管理番号 =  " & KOL.材料管理番号 & ") ZT2"
                    gSL_Select = gSL_Select & "       ON  ZT1.取引先       = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "       AND ZT1.材料管理番号 = '" & KOL.材料管理番号 & "'"
                    gSL_Select = gSL_Select & "       AND ZT1.受入         = ZT2.受入"
                End If
        
                Call HTMRead(gSL_Select, , 2)
        
                ' 背番号
                ''B1_Seban = HTM.背番号             '2016/11/23 del
                ' 収容数
                B1_Syuyo = Format(HTM.収容数1, "#,###")
                '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert End
            
                B1_Siyosu = Format(.構成数, "#,##0.000;;#")
'                B1_Gouho = .号補区分                               '2012/07/03
                B1_Gouho = ""                                       '2012/07/03
                ' 展開区分
                Select Case KOL.展開区分
                    Case 0: B1_Tenkai = ""
                    Case 1: B1_Tenkai = "まで対象"
                    Case 2: B1_Tenkai = "以降対象外"
                    Case 3: B1_Tenkai = "のみ対象外"
                End Select
             End If
             
             '============================= 2009/11/10 Start
             If L_wsHKbn = 1 Or L_wsHKbn = 2 Then
                B1_Level = ""
                If .手配先 <> "" Then
                   B1_Tehai = RTrim(.手配先)
                End If
                If .区分 <> "" Then
                   B1_KB = "材"
                End If
                B1_Hinbn = RTrim(.表示品番)
                B1_Sunpou = RTrim(.表示寸法)
                '-----2011/07/21 EDIT START-----
                B1_Hinnm = CisFun.Mid2(RTrim(.品名), 1, 25)
                'B1_Hinnm = RTrim(.品名)
                '-----2011/07/21 EDIT END-------
                '-----2011/07/21 ADD START-----
                gSL_Select = ""
                gSL_Select = gSL_Select & "SELECT HT1.背番号"
                gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(HT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= HT1.変更日 THEN HT1.収容数2 ELSE HT1.収容数1 END AS 収容数1"
                gSL_Select = gSL_Select & "  FROM 品番取引先マスタ HT1"
                gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                gSL_Select = gSL_Select & "               FROM 品番取引先マスタ"
                gSL_Select = gSL_Select & "              WHERE 取引先 = '" & KOL.取引先 & "'"
                gSL_Select = gSL_Select & "                AND 品番   = '" & KOL.親品番 & "') HT2"
                gSL_Select = gSL_Select & "       ON  HT1.取引先 = '" & KOL.取引先 & "'"
                gSL_Select = gSL_Select & "       AND HT1.品番   = '" & KOL.親品番 & "'"
                gSL_Select = gSL_Select & "       AND HT1.受入   = HT2.受入"
        
                Call HTMRead(gSL_Select, , 2)
                
                ' 背番号
                B1_Seban = HTM.背番号
                ' 収容数
                B1_Syuyo = Format(HTM.収容数1, "#,###")
                '-----2011/07/21 ADD END-------

                B1_Siyosu = Format(.構成数, "#,##0.000;;#")
                B1_Gouho = .号補区分
                ' 展開区分
                Select Case KOL.展開区分
                    Case 0: B1_Tenkai = ""
                    Case 1: B1_Tenkai = "まで対象"
                    Case 2: B1_Tenkai = "以降対象外"
                    Case 3: B1_Tenkai = "のみ対象外"
                End Select
             End If
             '============================= 2009/11/10 End
        End With
        '*---------------------------------*
        '*  Body部　印刷
        '*---------------------------------*
        With CisPrtCXM0150
            .PrintFld ("B1*")
            .cLine = .cLine + 1
            .ClrLstFld ("B1")
    '       カウンタ　ＵＰ
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With

        Call KOLReadNext(1)
    Loop
    Call KOLClose(1)
    
'   印刷終了
    Call CisPrtCXM0150.PrintEndDoc

End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim TtlName     As String
    Dim YmdName     As String

    HeadWrite = False
    
    With CisPrtCXM0150
    '   １頁内印刷行数ﾁｪｯｸ
        If .iLineMax >= .cLine + 1 Then Exit Function

        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")   'HEAD部内容消去

    '============================== 2009/11/10 Start
    ' Title
        Select Case L_wsHKbn
               Case 0: H1_Title = "【部品 - 正展開】"
               Case 1: H1_Title = "【部品 - 逆展開】"
               Case 2: H1_Title = "【材　料】"
        End Select
        
    ' Title2
        Select Case L_wsHKbn
               Case 0: H1_Title2 = ""
               Case 1: H1_Title2 = "子品番：" & RTrim(KOL.子品番)
               Case 2: H1_Title2 = "材質：" & RTrim(KOL.子品番)
        End Select
    '============================== 2009/11/10 End
    
    '============================== 2014/02/13 Start
        If L_wsHKbn = 0 Then
            H1_JyuBiko = RTrim(CXM0150.lb_JyutyuBiko)
        Else
            H1_JyuBiko = ""
        End If
    '============================== 2014/02/13 End
    
    '   発行日・頁
        H1_HYmd = mPrtDate & " " & mPrtTime
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")

        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With

    HeadWrite = True
End Function


