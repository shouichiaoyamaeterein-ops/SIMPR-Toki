VERSION 5.00
Object = "{0D300FC0-B2EA-11D1-8D3B-444553540000}#1.30#0"; "Qrmaker.ocx"
Begin VB.Form CXJ0130L 
   Caption         =   "現品票"
   ClientHeight    =   4215
   ClientLeft      =   7695
   ClientTop       =   3075
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   10140
   Begin QRMAKERLib.QRmaker B1_QRmaker1 
      Height          =   1080
      Left            =   8205
      TabIndex        =   30
      Top             =   2280
      Width           =   1185
      _Version        =   65566
      _ExtentX        =   2090
      _ExtentY        =   1905
      _StockProps     =   1
      Picture         =   "CXJ0130L.frx":0000
   End
   Begin VB.Line B1_Line4 
      BorderWidth     =   2
      X1              =   5760
      X2              =   5760
      Y1              =   3105
      Y2              =   3930
   End
   Begin VB.Label B1_Syuka 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9090
      TabIndex        =   29
      Top             =   3705
      Width           =   585
   End
   Begin VB.Label B1_Jisya 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7155
      TabIndex        =   28
      Top             =   3705
      Width           =   1890
   End
   Begin VB.Line B1_Line3 
      BorderWidth     =   2
      X1              =   2145
      X2              =   7021
      Y1              =   3105
      Y2              =   3105
   End
   Begin VB.Line B1_Line2 
      BorderWidth     =   2
      X1              =   2145
      X2              =   7035
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line B1_Line1 
      BorderWidth     =   2
      X1              =   2160
      X2              =   7035
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   3615
      Index           =   0
      Left            =   2145
      Top             =   330
      Width           =   4905
   End
   Begin VB.Label B1_Ha 
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5025
      TabIndex        =   27
      Top             =   3420
      Width           =   450
   End
   Begin VB.Label B1_Mai 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "999/999"
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
      Left            =   5820
      TabIndex        =   26
      Top             =   3435
      Width           =   1155
   End
   Begin VB.Label B1_SSebanM 
      AutoSize        =   -1  'True
      Caption         =   "社内背番"
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
      TabIndex        =   25
      Top             =   2940
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label B1_SebanM 
      AutoSize        =   -1  'True
      Caption         =   "背番"
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
      Left            =   2220
      TabIndex        =   24
      Top             =   2430
      Width           =   630
   End
   Begin VB.Label B1_Time 
      AutoSize        =   -1  'True
      Caption         =   "99:99"
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
      Left            =   9015
      TabIndex        =   23
      Top             =   90
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label B1_SBinM 
      AutoSize        =   -1  'True
      Caption         =   "便"
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
      Index           =   1
      Left            =   1470
      TabIndex        =   22
      Top             =   2130
      Width           =   225
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      BorderWidth     =   2
      Height          =   555
      Index           =   2
      Left            =   195
      Top             =   330
      Width           =   1560
   End
   Begin VB.Label B1_HinnmM 
      AutoSize        =   -1  'True
      Caption         =   "品名"
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
      Left            =   2220
      TabIndex        =   20
      Top             =   1380
      Width           =   450
   End
   Begin VB.Label B1_Hinnm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2730
      TabIndex        =   19
      Top             =   1650
      Width           =   4260
   End
   Begin VB.Label B1_DateM 
      AutoSize        =   -1  'True
      Caption         =   "日付"
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
      Left            =   7065
      TabIndex        =   18
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label B1_Date 
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99"
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
      Left            =   7665
      TabIndex        =   17
      Top             =   90
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label B1_UkeirM 
      AutoSize        =   -1  'True
      Caption         =   "受入"
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
      Left            =   7125
      TabIndex        =   16
      Top             =   2310
      Width           =   450
   End
   Begin VB.Label B1_Ukeir 
      AutoSize        =   -1  'True
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7695
      TabIndex        =   15
      Top             =   2310
      Width           =   270
   End
   Begin VB.Label B1_HinbnM 
      AutoSize        =   -1  'True
      Caption         =   "品番"
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
      Left            =   2220
      TabIndex        =   14
      Top             =   465
      Width           =   450
   End
   Begin VB.Label B1_SyuyoM 
      AutoSize        =   -1  'True
      Caption         =   "収容数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2220
      TabIndex        =   13
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label B1_Syuyo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "99,999"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3300
      TabIndex        =   12
      Top             =   3420
      Width           =   1350
   End
   Begin VB.Label B1_SBinM 
      AutoSize        =   -1  'True
      Caption         =   "出荷便"
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
      Index           =   0
      Left            =   105
      TabIndex        =   11
      Top             =   2130
      Width           =   675
   End
   Begin VB.Label B1_SBin 
      AutoSize        =   -1  'True
      Caption         =   "99"
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
      Left            =   1035
      TabIndex        =   10
      Top             =   2100
      Width           =   330
   End
   Begin VB.Label B1_SYmdM 
      AutoSize        =   -1  'True
      Caption         =   "納入日"
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
      Left            =   120
      TabIndex        =   9
      Top             =   1170
      Width           =   675
   End
   Begin VB.Label B1_SYmd 
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99"
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
      Left            =   105
      TabIndex        =   8
      Top             =   1440
      Width           =   1950
   End
   Begin VB.Label B1_Tornm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
      Left            =   7125
      TabIndex        =   7
      Top             =   780
      Width           =   2940
   End
   Begin VB.Label B1_TornmM 
      AutoSize        =   -1  'True
      Caption         =   "納入先名"
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
      Left            =   7095
      TabIndex        =   6
      Top             =   510
      Width           =   930
   End
   Begin VB.Label B1_TorkjM 
      AutoSize        =   -1  'True
      Caption         =   "工場"
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
      Left            =   7140
      TabIndex        =   5
      Top             =   1395
      Width           =   450
   End
   Begin VB.Label B1_SSeban 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "XXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   36
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   720
      Left            =   150
      TabIndex        =   4
      Top             =   3180
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label B1_Torkj 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
      Left            =   7140
      TabIndex        =   3
      Top             =   1665
      Width           =   2505
   End
   Begin VB.Label B1_Jykb 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "指　示"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   21.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   285
      TabIndex        =   2
      Top             =   390
      Width           =   1350
   End
   Begin VB.Label B1_Seban 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   36
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3180
      TabIndex        =   1
      Top             =   2235
      Width           =   3015
   End
   Begin VB.Label B1_PrintKb 
      AutoSize        =   -1  'True
      Caption         =   "ノーかんばんエフ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3060
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Shape B1_Shape1 
      BackColor       =   &H00C0FFC0&
      Height          =   4125
      Index           =   1
      Left            =   45
      Top             =   45
      Width           =   9660
   End
   Begin VB.Label B1_Hinbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   21.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2700
      TabIndex        =   21
      Top             =   735
      Width           =   4335
   End
End
Attribute VB_Name = "CXJ0130L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  ノーかんばんエフ　発行
'**       フォームID    :  CXJ0130L
'**       処理概要      :
'**
'**       作  成  日    :  2004/03/24
'**       変  更  日    :  2008/04/09 By CIS 背番号8桁
'**       変  更  日    :  2010/03/03 By CIS 土岐津Ver
'**       変  更  日    :  2010/03/10 By CIS 土岐津Ver
'**       変  更  日    :  2019/08/02 By CIS QR対応
'**********************************************************'
'#比較退避用#
    Dim mPrtDate        As String
    Dim mPrtTime        As String
    Dim MaxGyo          As Byte         '頁行数
    Dim NowGyo          As Byte
    Dim GMaisu          As Long         '現枚数
    Dim TMaisu          As Long         '総枚数
    Dim mJisyaNm        As String       '自社名

    Dim mQRData         As String * 150             ' 2019/08/02 ADD
    Dim mHMoto          As String                   ' 2019/08/02 ADD

'*------------------------------------*
'*      Ｆｏｒｍ＿Ｌｏａｄ イベント     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")             '発行日
    mPrtTime = Format(Now(), "hh:nn")                  '発行時間
    '*---------------------*
    '*  プリント系初期設定
    '*---------------------*
    With CisPrtCXJ0130
        .iPrtTitle = "ノーかんばんエフ"
'''        .iLineMax = 80
        .iLineMax = 60
'       .iLinePitch = 8
'''        .iPaperSize = p_A4
        .iPaperSize = p_B5
'       .iOrientation = po_PORTRAIT

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

'--( 自社名称 )--  取引先マスタより自社名称を求める
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(gCompany) & Left(IRN.oyian, 1) & "'"
    If TRMRead(gSL_Select, 1) Then
        mJisyaNm = TRM.略称
    Else
        mJisyaNm = ""
    End If
    
    
    Call PrintProc

End Sub
'*---------------------------------------------------------------*
'*　 << 印  刷  処  理 >>
'*---------------------------------------------------------------*
Private Function PrintProc()
    
'===============================================' 2019/08/02 ADD START
'*---------------------------------*
'*   システム情報(発行元)
'*---------------------------------*
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 4
    Call IRNRead("", 1)
    mHMoto = Mid(IRN.oyian, 1, 7)
'===============================================' 2019/08/02 ADD END

'*---------------------------------*
'*  初期設定
'*---------------------------------*
    MaxGyo = 3
    NowGyo = MaxGyo
    Call CisPrtCXJ0130.ClrLstFld("B1")

'*---------------------------------*
'*      受注テーブル読込み
'*---------------------------------*
    Do Until Not JYT_RDSTS
        With JYT
        '   >> 収容数のある時は枚数、無い時は枚数
            If .収容数 <> 0 Then
                TMaisu = .枚数
            Else
                TMaisu = 1
            End If
        End With
'       *---------------------------------*
'       *   Body行　印刷（枚数分繰返し）
'       *---------------------------------*
        For GMaisu = 1 To TMaisu
        '   >> 改頁の判定
            If NowGyo >= MaxGyo Then
                Call HeadWrite(0)
                NowGyo = 0
            End If
            NowGyo = NowGyo + 1
'
            Call CisPrtCXJ0130.ClrLstFld("B1")
        '   >> BODY部編集
            With JYT
                B1_PrintKb = "ノーかんばんエフ"
                B1_DateM = "日付"
                B1_Date = mPrtDate
                B1_Time = mPrtTime
                B1_Jykb = RTrim(.受注区分名)
                B1_HinbnM = "品番"
                B1_Hinbn = RTrim(.表示品番)
                B1_HinnmM = "品名"
                B1_Hinnm = RTrim(.品名)
                B1_SYmdM = "納入日"
                B1_SYmd = Format(.出荷予定日, "0000/00/00")
                B1_SBinM(0) = "出荷便"
                B1_SBinM(1) = "便"
                B1_SBin = RTrim(.出荷予定便)
                B1_TornmM = "納入先名"
                B1_Tornm = RTrim(.納入先名)
                B1_TorkjM = "工場"
                B1_Torkj = RTrim(.工場名)
                B1_UkeirM = "受入"
                B1_Ukeir = RTrim(.受入)
                B1_SebanM = "背番"
                B1_Seban = RTrim(.背番号)
                B1_SSebanM = "社内背番"
                B1_SSeban = RTrim(.社内背番)
                If .収容数 <> 0 Then
                    If GMaisu = TMaisu And .端数 <> 0 Then
                        B1_SyuyoM = "納入数"
                        B1_Syuyo = Format(.端数, "#,##0")
                        B1_Ha = "端"
                    Else
                        B1_SyuyoM = "収容数"
                        B1_Syuyo = Format(.収容数, "#,##0")
                        B1_Ha = ""
                    End If
                Else
                    B1_SyuyoM = "納入数"
                    B1_Syuyo = Format(.納入数, "#,##0")
                End If
                
                B1_Mai = Format(GMaisu, "##0") & "/" & Format(TMaisu, "##0")
                
                B1_Jisya = RTrim(mJisyaNm)
                B1_Syuka = RTrim(.出荷管理部署名)
                
                
            End With
        
        '===============================================' 2019/08/02 ADD START
        'QR
            Call QR_Edit
            B1_QRmaker1.InputData = mQRData
            B1_QRmaker1.Refresh
        '===============================================' 2019/08/02 ADD END

        '   BODY部印刷
            With CisPrtCXJ0130
                .PrintFld ("B1*")
'''                .cLine = .cLine + 22
                .cLine = .cLine + 20
'
'           *---------------------------------*
'           *  カウンタ　ＵＰ
'           *---------------------------------*
                .cDataCnt = .cDataCnt + 1
                .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(CXJ0130_PrtCnt, "#,###")
            End With
        Next
'
        Call JYTReadNext(1)
    Loop
    Call JYTClose(1)

'   >> 印刷終了
    Call CisPrtCXJ0130.PrintEndDoc

End Function
'*---------------------------------------------------------------*
'*      ヘッダ部印刷                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    HeadWrite = False
    With CisPrtCXJ0130
        .cLine = 0
        .cPage = .cPage + 1
'
        .ClrLstFld ("H1")   'HEAD部内容消去
'
        .PrintNewPage       '改頁処理
        .PrintFld ("H1")    'HEAD部印刷
    End With
    HeadWrite = True
End Function

'=======================================================================================' 2019/08/02 ADD START
'*---------------------------------------------------------------*
'*   << QR内容編集 >>                                             *
'*---------------------------------------------------------------*
Private Sub QR_Edit()
    mQRData = Space(150)
    
    With JYT
        Mid(mQRData, 1, 7) = Left(mHMoto, 7)                '発行元
        Mid(mQRData, 8, 5) = gCompany                       '自社コード
        Mid(mQRData, 13, 1) = "J"                           'QR識別(R:臨時かんばん、S:現品票、H:端数ラベル) ' 2019/08/02 UPD："R"→"J"
        Mid(mQRData, 14, 25) = .品番                     '品番
        Mid(mQRData, 39, 1) = "0"                           '品目(0:受注、1:発注)   ' 2019/08/05 UPD：1:発注 → 0:受注　に変更
        Mid(mQRData, 40, 7) = .納入先                    '取引先コード              ' 2019/08/02 UPD：仕入先→納入先
        Mid(mQRData, 47, 2) = .受入                      '取引先受入
        Mid(mQRData, 49, 1) = " "                        '手配区分                  ' 2019/08/02 UPD：Format(.手配区分, "0")→" "
    '    Mid(mQRData, 50, 5) = HCT.背番号                    '背番号
    '    Mid(mQRData, 55, 5) = HCT.納入番号                  '納入番号
    '    Mid(mQRData, 60, 8) = HCT.納入日                    '納入日
    '    Mid(mQRData, 68, 2) = HCT.納入便                    '納入便
    '    Mid(mQRData, 70, 7) = HCT.納入先                    '納入先コード
    '    Mid(mQRData, 77, 2) = HCT.納入先受入                '納入先受入
    '    Mid(mQRData, 79, 6) = Format(HCT.収容数, "000000")  '収容数
    '    Mid(mQRData, 85, 10) = HCT.収容器                   '収容器
    '    Mid(mQRData, 95, 6) = Format(HCT.収容数, "000000")  '数量
    '    Mid(mQRData, 101, 3) = Format(GMaisu, "000")          '現枚数
    '    Mid(mQRData, 104, 3) = Format(TMaisu, "000")        '総枚数
    '    CisFun.DateE_AfVal = mPrtDate
    '    CisFun.DateEditUn
    '    Mid(mQRData, 107, 8) = CisFun.DateE_BefVal          '発行日
    '    Mid(mQRData, 115, 5) = mPrtTime                     '発行時間
    '    Mid(mQRData, 120, 5) = HCT.サイクル                 'サイクル
    '    Mid(mQRData, 125, 10) = HCT.置場                    '置場
    '    Mid(mQRData, 135, 4) = HCT.車型                     '車型
    '    Mid(mQRData, 139, 1) = Format(HCT.発注区分, "0")    '発注区分
        Mid(mQRData, 50, 8) = .背番号                    '背番号
        Mid(mQRData, 58, 5) = .納入番号                  '納入番号
        Mid(mQRData, 63, 8) = .納入日                    '納入日
        Mid(mQRData, 71, 2) = .納入便                    '納入便
        Mid(mQRData, 73, 7) = .納入先                    '納入先コード
''        Mid(mQRData, 80, 2) = .納入先受入                '納入先受入
        Mid(mQRData, 82, 6) = Format(.収容数, "000000")  '収容数
''        Mid(mQRData, 88, 10) = .収容器                   '収容器
        Mid(mQRData, 98, 6) = Format(.収容数, "000000")  '数量
        Mid(mQRData, 104, 3) = Format(GMaisu, "000")          '現枚数
        Mid(mQRData, 107, 3) = Format(TMaisu, "000")        '総枚数
        CisFun.DateE_AfVal = mPrtDate
        CisFun.DateEditUn
        Mid(mQRData, 110, 8) = CisFun.DateE_BefVal          '発行日
        Mid(mQRData, 118, 5) = mPrtTime                     '発行時間
''        Mid(mQRData, 123, 5) = .サイクル                 'サイクル
        Mid(mQRData, 128, 10) = "          "               '置場                    ' 2019/08/02 UPD：. 置場 →Space(10)
''        Mid(mQRData, 138, 5) = .車型                     '車型
        Mid(mQRData, 143, 1) = " "                         '発注区分                ' 2019/08/02 UPD：Format(.発注区分, "0")→" "
    End With
End Sub
'=======================================================================================' 2019/08/02 ADD END


