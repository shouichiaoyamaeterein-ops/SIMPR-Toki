VERSION 5.00
Begin VB.Form CXH0065L 
   Caption         =   "発注予定一覧表"
   ClientHeight    =   3090
   ClientLeft      =   3345
   ClientTop       =   3570
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   11640
   Begin VB.Label B1_ZHKbn 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "X"
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
      Left            =   9030
      TabIndex        =   25
      Top             =   1275
      Width           =   120
   End
   Begin VB.Label B1_NYmd 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD-99"
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
      Left            =   4845
      TabIndex        =   24
      Top             =   1275
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "納入日-便"
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
      Left            =   4845
      TabIndex        =   23
      Top             =   840
      Width           =   885
   End
   Begin VB.Label B1_HYmd 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD-99"
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
      Left            =   9240
      TabIndex        =   22
      Top             =   1275
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "発注日-便"
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
      Left            =   9240
      TabIndex        =   21
      Top             =   825
      Width           =   885
   End
   Begin VB.Label B1_HTime 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   20
      Top             =   1275
      Width           =   675
   End
   Begin VB.Label H1_HSYmd 
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99"
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
      Left            =   1245
      TabIndex        =   19
      Top             =   300
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "発注処理日"
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
      Left            =   165
      TabIndex        =   18
      Top             =   375
      Width           =   990
   End
   Begin VB.Line B1_K_LINE 
      BorderStyle     =   3  '点線
      Index           =   2
      X1              =   1005
      X2              =   11240
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "確認"
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
      Index           =   15
      Left            =   10605
      TabIndex        =   17
      Top             =   825
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "便"
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
      Index           =   4
      Left            =   1095
      TabIndex        =   16
      Top             =   855
      Width           =   180
   End
   Begin VB.Label B1_HBin 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "#9"
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
      Left            =   1080
      TabIndex        =   15
      Top             =   1275
      Width           =   240
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   2
      X1              =   150
      X2              =   11235
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "最終発注日-便"
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
      Left            =   7560
      TabIndex        =   14
      Top             =   825
      Width           =   1275
   End
   Begin VB.Line T1_K_LINE 
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   11235
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line B1_K_LINE 
      Index           =   1
      X1              =   150
      X2              =   11235
      Y1              =   1200
      Y2              =   1200
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
      TabIndex        =   12
      Top             =   225
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
      TabIndex        =   11
      Top             =   225
      Width           =   225
   End
   Begin VB.Label B1_Cycle 
      AutoSize        =   -1  'True
      Caption         =   "9-99-99"
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
      Left            =   6540
      TabIndex        =   10
      Top             =   1275
      Width           =   840
   End
   Begin VB.Label B1_Tornm 
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
      Left            =   2310
      TabIndex        =   9
      Top             =   1275
      Width           =   2400
   End
   Begin VB.Label B1_Torcd 
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
      Left            =   1470
      TabIndex        =   8
      Top             =   1275
      Width           =   840
   End
   Begin VB.Label B1_Ukeir 
      AutoSize        =   -1  'True
      Caption         =   "XX"
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
      Left            =   6075
      TabIndex        =   7
      Top             =   1275
      Width           =   240
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   1
      X1              =   150
      X2              =   11235
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label Label2 
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
      Index           =   3
      Left            =   6015
      TabIndex        =   6
      Top             =   825
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "発注時間"
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
      Index           =   6
      Left            =   195
      TabIndex        =   5
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "サイクル"
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
      Left            =   6570
      TabIndex        =   4
      Top             =   825
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "仕　入　先"
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
      Left            =   2460
      TabIndex        =   3
      Top             =   825
      Width           =   975
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
      TabIndex        =   2
      Top             =   225
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
      TabIndex        =   1
      Top             =   225
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "※　発注予定一覧表　※"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3900
      TabIndex        =   0
      Top             =   285
      Width           =   3300
   End
   Begin VB.Label B1_ENDHDate 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "YY/MM/DD-99"
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
      Left            =   7545
      TabIndex        =   13
      Top             =   1275
      Width           =   1320
   End
End
Attribute VB_Name = "CXH0065L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  発注予定一覧表　発行
'**       フォームID    :  CXH0065L
'**       処理概要      :
'**
'**       作  成  日    :  2005/04/27
'**       変  更  日    :
'**
'**********************************************************'
    Dim mPrtDate        As String
    Dim mPrtTime        As String
'   印字サプレス用
    Dim Key_HTime       As String

'*------------------------------------*
'*      Ｆｏｒｍ＿Ｌｏａｄ イベント     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")          '発行日
    mPrtTime = Format(Now(), "hh:nn")               '発行時間
    '*---------------------*
    '*  プリント系初期設定
    '*---------------------*
    With CisPrtCXH0065
        .iPrtTitle = "発注予定一覧表"
        .iLineMax = 82
        .iLinePitch = 8
        .iPaperSize = p_A4
       '.iOrientation = po_PORTRAIT

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
    Call CisPrtCXH0065.ClrLstFld("B1")
    CisPrtCXH0065.cLine = 99
'*---------------------------------*
'*  発注サイクルマスタ読込み
'*---------------------------------*
    Do Until Not HYH_RDSTS

        Call HeadWrite(0)

        '*---------------------------------*
        '*  Body行　編集
        '*---------------------------------*
        With HYH
            If Key_HTime <> RTrim(.発注時間) Then
                B1_HTime = Left(.発注時間, 2) & ":" & Mid(.発注時間, 3, 2)
                B1_K_LINE(1).Visible = True
                B1_K_LINE(2).Visible = False
            Else
                B1_K_LINE(1).Visible = False
                B1_K_LINE(2).Visible = True
            End If
            B1_HBin = CisFun.Val2(.発注便) & "便"
            B1_Torcd = RTrim(.仕入先)
            B1_Tornm = RTrim(.仕入先名)
            B1_Ukeir = RTrim(.受入)
            If .曜日発注 = 0 Then
                B1_Cycle = Left(.サイクル, 1) & "-" & Format(CisFun.Val2(Mid(.サイクル, 2, 2)), "@@") & "-" & Format(CisFun.Val2(Mid(.サイクル, 4, 2)), "@@")
            Else
                B1_Cycle = Format(CDate(Left(.発注日, 4) & "/" & Mid(.発注日, 5, 2) & "/" & Mid(.発注日, 7, 2)), "aaa") & "曜日"
            End If
            
            If Trim(.最終発注日) <> "" Then
               With CisFun
                    .DateE_BefVal = HYH.最終発注日
                    .DateEdit
                    B1_ENDHDate = Mid(.DateE_AfVal, 3)
               End With
            End If
            If RTrim(.最終発注便) <> "" Then
                B1_ENDHDate = B1_ENDHDate & "-" & Format(CisFun.Val2(.最終発注便), "@@")
            End If
            If .発行加算日 <> 0 Then
                B1_ZHKbn = "*"
            Else
                B1_ZHKbn = ""
            End If
            If Trim(.発注日) <> "" Then
               With CisFun
                    .DateE_BefVal = HYH.発注日
                    .DateEdit
                    B1_HYmd = Mid(.DateE_AfVal, 6)
               End With
            End If
            If RTrim(.発注便) <> "" Then
                B1_HYmd = B1_HYmd & "-" & Format(CisFun.Val2(.発注便), "@@")
            End If
            If Trim(.納入日) <> "" Then
               With CisFun
                    .DateE_BefVal = HYH.納入日
                    .DateEdit
                    B1_NYmd = Mid(.DateE_AfVal, 6)
               End With
            End If
            If RTrim(.納入便) <> "" Then
                B1_NYmd = B1_NYmd & "-" & Format(CisFun.Val2(.納入便), "@@")
            End If
            With CisPrtCXH0065
                .PrintFld ("B1*")
                .cLine = .cLine + 2
                .ClrLstFld ("B1")
            End With
            
        
            With CisPrtCXH0065
            '   カウンタ　ＵＰ
                .cDataCnt = .cDataCnt + 1
                .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
            End With
        '   印字サプレス用
            Key_HTime = RTrim(.発注時間)
        End With

        Call HYHReadNext(1)

    Loop
    Call HYHClose(1)
    If RTrim(Key_HTime) <> "" Then
        '*---------------------------------*
        '*  Body部　印刷
        '*---------------------------------*
        With CisPrtCXH0065
            .PrintFld ("T1*")
        End With
    End If

'   印刷終了
    Call CisPrtCXH0065.PrintEndDoc
End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    HeadWrite = False

    With CisPrtCXH0065
    '   １頁内印刷行数ﾁｪｯｸ
        If .iLineMax >= .cLine + 1 Then Exit Function
        If .cPage > 0 Then
            .PrintFld ("T1*")
        End If

        .cLine = 0
        .cPage = .cPage + 1
        .ClrLstFld ("H1")   'HEAD部内容消去
    '   印字サプレス用
        Key_HTime = ""
    '   発注処理日
        If Trim(HYH.発行日) <> "" Then
            H1_HSYmd = Left(HYH.発行日, 4) & "年" & Mid(HYH.発行日, 5, 2) & "月" & Mid(HYH.発行日, 7, 2) & "日"
        End If
    '   発行日・頁
        H1_HYmd = mPrtDate & " " & mPrtTime
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")

        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With

    HeadWrite = True
End Function
