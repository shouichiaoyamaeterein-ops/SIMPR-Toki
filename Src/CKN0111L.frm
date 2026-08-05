VERSION 5.00
Begin VB.Form CKN0111L 
   Caption         =   "発注サイクル一覧表"
   ClientHeight    =   4020
   ClientLeft      =   810
   ClientTop       =   4995
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4020
   ScaleWidth      =   12165
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "最終発注日"
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
      Left            =   4950
      TabIndex        =   27
      Top             =   1260
      Width           =   1050
   End
   Begin VB.Line T1_K_LINE 
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11310
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Line B1_K_LINE 
      BorderStyle     =   3  '点線
      Index           =   2
      X1              =   120
      X2              =   11310
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Line B1_K_LINE 
      BorderStyle     =   3  '点線
      Index           =   3
      X1              =   3510
      X2              =   11310
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Line B1_K_LINE 
      Index           =   1
      X1              =   120
      X2              =   11310
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Label B1_NJikan 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
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
      Left            =   10560
      TabIndex        =   25
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label B1_KJikan 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
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
      Left            =   9675
      TabIndex        =   24
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label B1_HJikan 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
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
      Left            =   8850
      TabIndex        =   23
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "入荷時間"
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
      Left            =   10470
      TabIndex        =   22
      Top             =   1260
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "回収時間"
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
      Index           =   9
      Left            =   9615
      TabIndex        =   21
      Top             =   1260
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "発行時間"
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
      Index           =   8
      Left            =   8790
      TabIndex        =   20
      Top             =   1260
      Width           =   720
   End
   Begin VB.Label B1_NBin 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
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
      Left            =   6120
      TabIndex        =   19
      Top             =   1680
      Width           =   480
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
      TabIndex        =   18
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
      TabIndex        =   17
      Top             =   120
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
      Left            =   3600
      TabIndex        =   16
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label B1_KDay 
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
      Left            =   8325
      TabIndex        =   15
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label B1_FJikan 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
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
      Left            =   7440
      TabIndex        =   14
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label B1_HYKbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
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
      Left            =   6765
      TabIndex        =   13
      Top             =   1680
      Width           =   480
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
      Left            =   1080
      TabIndex        =   12
      Top             =   1680
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
      Left            =   150
      TabIndex        =   11
      Top             =   1680
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
      Left            =   4560
      TabIndex        =   10
      Top             =   1680
      Width           =   240
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   11310
      Y1              =   1575
      Y2              =   1575
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
      Left            =   4500
      TabIndex        =   9
      Top             =   1260
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "加算日"
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
      Index           =   7
      Left            =   8145
      TabIndex        =   8
      Top             =   1260
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "振出時間"
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
      Left            =   7350
      TabIndex        =   7
      Top             =   1260
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "昼夜勤"
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
      Index           =   5
      Left            =   6750
      TabIndex        =   6
      Top             =   1260
      Width           =   540
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
      Left            =   3600
      TabIndex        =   5
      Top             =   1260
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
      Left            =   720
      TabIndex        =   4
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "入荷便"
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
      Left            =   6060
      TabIndex        =   3
      Top             =   1260
      Width           =   540
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11310
      Y1              =   1125
      Y2              =   1125
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
      TabIndex        =   1
      Top             =   90
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "※　発注サイクル一覧表　※"
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
      Left            =   3540
      TabIndex        =   0
      Top             =   540
      Width           =   3900
   End
   Begin VB.Label B1_ENDHDate 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "YY/MM/DD"
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
      Left            =   5010
      TabIndex        =   26
      Top             =   1680
      Width           =   960
   End
End
Attribute VB_Name = "CKN0111L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  発注サイクル一覧表　発行
'**       フォームID    :  CKN0111L
'**       処理概要      :
'**
'**       作  成  日    :  2004/05/11
'**       変  更  日    :
'**
'**********************************************************'
    Dim mPrtDate        As String
    Dim mPrtTime        As String
'   印字サプレス用
    Dim wSircd          As String
    Dim wTorcd          As String

'*------------------------------------*
'*      Ｆｏｒｍ＿Ｌｏａｄ イベント     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")          '発行日
    mPrtTime = Format(Now(), "hh:nn")               '発行時間
    '*---------------------*
    '*  プリント系初期設定
    '*---------------------*
    With CisPrtCKN0110
        .iPrtTitle = "発注サイクル一覧表"
        .iLineMax = 59
        .iLinePitch = 6
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

    Call CisPrtCKN0110.ClrLstFld("B1")

'*---------------------------------*
'*  発注サイクルマスタ読込み
'*---------------------------------*
    Do Until Not HCM_RDSTS

        Call HeadWrite(0)
        B1_K_LINE(1).Visible = False
        B1_K_LINE(2).Visible = False
        B1_K_LINE(3).Visible = False

        '*---------------------------------*
        '*  Body行　編集
        '*---------------------------------*
        With HCM
            If wSircd <> RTrim(.仕入先) Then
                If wTorcd <> RTrim(.取引先) Then
                    B1_Torcd = RTrim(.取引先) & "-" & RTrim(.工場CD)
                    If CisPrtCKN0110.cLine <> 0 Then
                        B1_K_LINE(1).Visible = True
                    End If
                Else
                    B1_Torcd = Space(CisFun.Len2(RTrim(.取引先))) & "-" & RTrim(.工場CD)
                    B1_K_LINE(2).Visible = True
                End If
                B1_Tornm = RTrim(.仕入先名)
            Else
                B1_K_LINE(3).Visible = True
            End If
            B1_Cycle = CisFun.Mid2(.サイクル, 1, 1) & "-" & CisFun.Mid2(.サイクル, 2, 2) & "-" & CisFun.Mid2(.サイクル, 4, 2)
            B1_Ukeir = RTrim(.受入)
            
            If Trim(.最終発注日) <> "" Then
               With CisFun
                    .DateE_BefVal = HCM.最終発注日
                    .DateEdit
                    B1_ENDHDate = Mid(.DateE_AfVal, 3)
               End With
            End If
            B1_NBin = CisFun.Val2(.振出便) & "便"
            Select Case .昼夜区分
                Case 0:     B1_HYKbn = "昼勤"
                Case Else:  B1_HYKbn = "夜勤"
            End Select
            B1_FJikan = CisFun.Mid2(.振出時間, 1, 2) & ":" & CisFun.Mid2(.振出時間, 3, 2)
            B1_KDay = Format(.発注加算日, "##")
            B1_HJikan = CisFun.Mid2(.発行時間, 1, 2) & ":" & CisFun.Mid2(.発行時間, 3, 2)
            If RTrim(.回収時間) <> "" Then
                B1_KJikan = CisFun.Mid2(.回収時間, 1, 2) & ":" & CisFun.Mid2(.回収時間, 3, 2)
            End If
            If RTrim(.入荷時間) <> "" Then
                B1_NJikan = CisFun.Mid2(.入荷時間, 1, 2) & ":" & CisFun.Mid2(.入荷時間, 3, 2)
            End If
        '   印字サプレス用
            wSircd = RTrim(.仕入先)
            wTorcd = RTrim(.取引先)
        End With
        '*---------------------------------*
        '*  Body部　印刷
        '*---------------------------------*
        With CisPrtCKN0110
            .PrintFld ("B1*")
            .cLine = .cLine + 1
            .ClrLstFld ("B1")
        '   カウンタ　ＵＰ
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With

        Call HCMReadNext(1)

    '   Tail部　印刷
       'With CisPrtCKN0110
       '    If Not HCM_RDSTS Or .iLineMax = .cLine Then
       '        .PrintFld ("T1*")
       '    End If
       'End With
    Loop
    Call HCMClose(1)

'   印刷終了
    Call CisPrtCKN0110.PrintEndDoc
End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    HeadWrite = False

    With CisPrtCKN0110
    '   １頁内印刷行数ﾁｪｯｸ
        If .iLineMax >= .cLine + 1 Then Exit Function

        .cLine = 0
        .cPage = .cPage + 1
        .ClrLstFld ("H1")   'HEAD部内容消去
    '   印字サプレス用
        wSircd = ""
        wTorcd = ""

    '   発行日・頁
        H1_HYmd = mPrtDate & " " & mPrtTime
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")

        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With

    HeadWrite = True
End Function

Private Sub Label3_Click()

End Sub

