VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form CXM2240L 
   Caption         =   "スクラップ単価変動額一覧表"
   ClientHeight    =   4455
   ClientLeft      =   30
   ClientTop       =   2775
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   19080
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   360
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label T1_UTKei 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14640
      TabIndex        =   36
      Top             =   3255
      Width           =   1335
   End
   Begin VB.Label B1_UTanH 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,###,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1041
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14640
      TabIndex        =   35
      Top             =   2835
      Width           =   1335
   End
   Begin VB.Label B1_UTanK 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   12900
      TabIndex        =   34
      Top             =   2835
      Width           =   1320
   End
   Begin VB.Label B1_UTanZ 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11250
      TabIndex        =   33
      Top             =   2835
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "売上単価変動額"
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
      Index           =   8
      Left            =   14535
      TabIndex        =   32
      Top             =   2250
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "今回売上単価"
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
      Index           =   7
      Left            =   12900
      TabIndex        =   31
      Top             =   2250
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "前回売上単価"
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
      Index           =   6
      Left            =   11265
      TabIndex        =   30
      Top             =   2250
      Width           =   1440
   End
   Begin VB.Label H1_KonTK 
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
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
      Left            =   4965
      TabIndex        =   29
      Top             =   1680
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "今回単価："
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
      Index           =   6
      Left            =   3570
      TabIndex        =   28
      Top             =   1695
      Width           =   1200
   End
   Begin VB.Label H1_ZenTK 
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
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
      Left            =   1725
      TabIndex        =   27
      Top             =   1680
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "前回単価："
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
      Index           =   5
      Left            =   315
      TabIndex        =   26
      Top             =   1695
      Width           =   1200
   End
   Begin VB.Label H1_ZSiTu 
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
      Left            =   6210
      TabIndex        =   25
      Top             =   1290
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "材　質："
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
      Index           =   4
      Left            =   5175
      TabIndex        =   24
      Top             =   1275
      Width           =   960
   End
   Begin VB.Label T1_STKei 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9450
      TabIndex        =   23
      Top             =   3270
      Width           =   1350
   End
   Begin VB.Label T1_Title 
      AutoSize        =   -1  'True
      Caption         =   "得意先計"
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
      Left            =   8010
      TabIndex        =   22
      Top             =   3255
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "今回ｽｸﾗｯﾌﾟ単価"
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
      Index           =   4
      Left            =   7260
      TabIndex        =   21
      Top             =   2250
      Width           =   1680
   End
   Begin VB.Label B1_STanK 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7365
      TabIndex        =   20
      Top             =   2835
      Width           =   1335
   End
   Begin VB.Label H1_Kikan 
      AutoSize        =   -1  'True
      Caption         =   "(YYYY/MM/DD～YYYY/MM/DD)"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2760
      TabIndex        =   19
      Top             =   960
      Width           =   2520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "対象期："
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
      Index           =   2
      Left            =   345
      TabIndex        =   18
      Top             =   915
      Width           =   960
   End
   Begin VB.Label H1_NenKi 
      AutoSize        =   -1  'True
      Caption         =   "YYYY年上期"
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
      Left            =   1365
      TabIndex        =   17
      Top             =   930
      Width           =   1200
   End
   Begin VB.Line T1_Line1 
      X1              =   150
      X2              =   16300
      Y1              =   3165
      Y2              =   3165
   End
   Begin VB.Label B1_STanH 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9450
      TabIndex        =   16
      Top             =   2835
      Width           =   1350
   End
   Begin VB.Line B1_LINE 
      Visible         =   0   'False
      X1              =   150
      X2              =   11220
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   5
      Index           =   1
      X1              =   150
      X2              =   16285
      Y1              =   2610
      Y2              =   2610
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   5
      Index           =   0
      X1              =   150
      X2              =   16300
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Label B1_HinBn 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   15
      Top             =   2820
      Width           =   2625
   End
   Begin VB.Label B1_Juryo 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "#,##0.00"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      TabIndex        =   14
      Top             =   2820
      Width           =   825
   End
   Begin VB.Label B1_STanZ 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5595
      TabIndex        =   13
      Top             =   2820
      Width           =   1275
   End
   Begin VB.Label B1_SeBan 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2970
      TabIndex        =   12
      Top             =   2820
      Width           =   840
   End
   Begin VB.Label H1_TORCD 
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
      Left            =   1365
      TabIndex        =   11
      Top             =   1290
      Width           =   2400
   End
   Begin VB.Label H1_TORNM 
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
      Left            =   2265
      TabIndex        =   10
      Top             =   1290
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Label H1_PAGE 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      Caption         =   "9,999"
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
      Left            =   15150
      TabIndex        =   9
      Top             =   450
      Width           =   600
   End
   Begin VB.Label H1_YMD 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD HH:MM"
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
      Left            =   12645
      TabIndex        =   8
      Top             =   465
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "得意先："
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
      Index           =   3
      Left            =   345
      TabIndex        =   7
      Top             =   1290
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ｽｸﾗｯﾌﾟ重量"
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
      Index           =   2
      Left            =   3960
      TabIndex        =   6
      Top             =   2250
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "品　　番"
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
      Left            =   420
      TabIndex        =   5
      Top             =   2250
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "前回ｽｸﾗｯﾌﾟ単価"
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
      Index           =   3
      Left            =   5430
      TabIndex        =   4
      Top             =   2250
      Width           =   1650
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ｽｸﾗｯﾌﾟ単価変動額"
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
      Index           =   5
      Left            =   9135
      TabIndex        =   3
      Top             =   2250
      Width           =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "背番号"
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
      Left            =   2940
      TabIndex        =   2
      Top             =   2250
      Width           =   720
   End
   Begin VB.Label Label1 
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
      Left            =   15780
      TabIndex        =   1
      Top             =   450
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "※ スクラップ単価変動額一覧表 ※"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   5235
      TabIndex        =   0
      Top             =   345
      Width           =   5760
   End
End
Attribute VB_Name = "CXM2240L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  スクラップ単価変動額一覧表
'**       フォームID    :  CXM2240L
'**       処理概要      :
'**
'**       作  成  日    :  2008/08/27
'**       変  更  日    :
'**
'**********************************************************'
    Dim wCnt            As Long
    Dim wNo             As Long
    Dim OutStr          As String
    Dim mPrtCnt         As Long
    
    Dim wTorcd          As String
    Dim wHinbn          As String
    Dim wHinCnt         As Long
    Dim wHSuryo         As Currency
    Dim wHUriGak        As Currency
    Dim wGSuryo         As Currency
    Dim wGKatahi        As Currency
    Dim wGUriGak        As Currency
    Dim wSSuryo         As Currency
    Dim wSKatahi        As Currency
    Dim wSUriGak        As Currency
'

'*------------------------------------*
'*      Ｆｏｒｍ＿Ｌｏａｄ イベント     *
'*------------------------------------*
Private Sub form_load()
    App.Title = "スクラップ単価変動額一覧表"
'    If Not DataRead Then Exit Sub
    
    wTorcd = ""

    '*---------------------*
    '*  プリント系初期設定
    '*---------------------*
'    Set CisPrtCXM2240 = New CISFormPrint

    With CisPrtCXM2240
        .iPrtTitle = "スクラップ単価変動額一覧表"
        .iLineMax = 28
        .iLinePitch = 5
        .iPaperSize = p_A4
        .iOrientation = po_Landscape
                
        Set .iPrtForm = CXM2240L

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
    
    PrtTim = Format(Now(), "yyyy/MM/dd HH:mm")      ' 発行日
    CXM2240_PrtOut = PrtCnt                         ' 件数
    Call PrintProc
End Sub

'*-------------------------*
'*      印刷データ読込
'*-------------------------*
Private Function DataRead()
    Dim c1 As Integer
    DataRead = False
    For c1 = 0 To Adodc2.Recordset.Fields.Count - 1
    If IsNull(Adodc2.Recordset.Fields(c1)) Then
        PrtDat(c1) = " "
    Else
        PrtDat(c1) = Adodc2.Recordset.Fields(c1)
    End If
    Next
    DataRead = True
End Function

'*---------------------*
'*　  印  刷  処  理
'*---------------------*
Private Function PrintProc()
    '-- データを読み込む
    Adodc2.ConnectionString = CisDB.ConnectString
    Adodc2.RecordSource = gSL_Select
    Adodc2.Refresh
    
    If Adodc2.Recordset.RecordCount = 0 Then Exit Function
    PrtCnt = Adodc2.Recordset.RecordCount
    
    Kei_S = 0: Kei_U = 0
    PrtRow = 0
    With CisPrtCXM2240
        .ClrLstFld ("H1")
        .ClrLstFld ("B1")
        .cLine = 0
        .cPage = 0
        .cDataCnt = 0
    End With
    
    Do While Not Adodc2.Recordset.EOF
        Call DataRead                   ' １行読込
        
        If PrtDat(0) <> H1_TORCD Or PrtDat(1) <> H1_ZSiTu _
        Or CisPrtCXM2240.iLineMax < CisPrtCXM2240.cLine Then
           Call HeadWrite               ' ヘッダ印刷 および 合計印刷
        End If
        
        B1_HinBn = PrtDat(4)                        ' 品番
        B1_SeBan = PrtDat(5)                        ' 背番号
        B1_Juryo = Format(PrtDat(6), "#,##0.00")    ' スクラップ重量
        B1_STanZ = Format(PrtDat(7), "#,##0.00")    ' 前回スクラップ単価
        B1_STanK = Format(PrtDat(8), "#,##0.00")    ' 今回スクラップ単価
        B1_STanH = Format(PrtDat(9), "#,##0.00")    ' スクラップ単価変動額
        B1_UTanZ = Format(PrtDat(10), "#,##0.00")   ' 前回売上単価
        B1_UTanK = Format(PrtDat(11), "#,##0.00")   ' 今回売上単価
        B1_UTanH = Format(PrtDat(12), "#,##0.00")   ' 売上単価変動額
        Kei_S = Kei_S + Round(PrtDat(9), 4)         ' スクラップ単価変動額合計
        Kei_U = Kei_U + Round(PrtDat(12), 4)        ' 売上単価変動合計
        
        With CisPrtCXM2240
            Call .PrintFld("B1")        ' １行印刷
            Call .ClrLstFld("B1")
            .cLine = .cLine + 1
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(PrtCnt, "#,###")
            DoEvents
        End With
        
        PrtRow = PrtRow + 1
        Adodc2.Recordset.MoveNext
    Loop
    Call TotalWrite                     ' 最終ページの合計を印刷する
    
    ' 印刷終了
    Call CisPrtCXM2240.PrintEndDoc
End Function

'*----------------*
'*  ヘッダ部印刷  *
'*----------------*
Private Function HeadWrite()
    HeadWrite = False
    
    With CisPrtCXM2240
        '-- いちばん初めのヘッダ印刷でなければ合計を印刷 -------
        If .cDataCnt <> 0 Then
           If PrtDat(0) <> H1_TORCD _
           Or PrtDat(1) <> H1_ZSiTu Then    ' ページまたぎチェック
              Call TotalWrite               ' 同得意先ならば合計は印刷しない
           End If
        End If
        .cLine = 0
        .cPage = .cPage + 1
        If .cPage <> 1 Then       ' 改ページ
            .PrintNewPage
        End If

        .ClrLstFld ("H1")       ' HEAD部内容消去
        H1_YMD = PrtTim         ' 印刷年月日時分
        H1_PAGE = .cPage        ' ページ
        H1_NenKi = PrtNKi       ' 対象期
        H1_Kikan = PrtKKn       ' 期間
        H1_TORCD = PrtDat(0)    ' 得意先
        H1_ZSiTu = PrtDat(1)    ' 材質
        H1_ZenTK = PrtDat(2)    ' 前回単価
        H1_KonTK = PrtDat(3)    ' 今回単価
        If .iPreview = pv_Preview Then .PrintNewPage
        .PrintFld ("H1")        ' HEAD部印刷
'        .cDataCnt = .cDataCnt + 1
'        .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(CXH0535_PrtCnt, "#,###")
    End With

    HeadWrite = True
End Function

'*----------------*
'*   合計印刷     *
'*----------------*
Private Sub TotalWrite(Optional Kb As Integer = 0)
'    Dim iHosei  As Integer
'    iHosei = 300                ' 表示位置補正値
'    T1_Line1.Y1 = T1_Line1.Y1 - iHosei
'    T1_Line1.Y2 = T1_Line1.Y2 - iHosei
'    T1_Title.Top = T1_Title.Top - iHosei
'    T1_STKei.Top = T1_STKei.Top - iHosei
'    T1_UTKei.Top = T1_UTKei.Top - iHosei
    
    Dim iHosei  As Integer
    iHosei = 90 ' 表示位置補正値
    
    T1_Line1.Y1 = B1_HinBn.Top
    T1_Line1.Y2 = T1_Line1.Y1
    T1_Title.Top = T1_Line1.Y1 + iHosei
    T1_STKei.Top = T1_Title.Top
    T1_UTKei.Top = T1_Title.Top
    
    With CisPrtCXM2240
        T1_STKei = Format(Kei_S, "#,##0.00")
        T1_UTKei = Format(Kei_U, "#,##0.00")
        Call .PrintFld("T1")
    End With

    Kei_S = 0   ' 集計値クリア
    Kei_U = 0
End Sub
        
        
'---------------------------------------------------------------------------------------------------
'        T1_Line.Visible = True
'        If wHinCnt > 1 Then
'            T1_MIDASI = "((品 番 計))"
'            T1_SU = Format(Format(wHSuryo, "#,##0;-#,##0;#"), String(9, "@"))
'            T1_UriGak = Format(Format(wHUriGak, "#,##0;-#,##0;#"), String(13, "@"))
'            Call .PrintFld("T1")
'            Call .ClrLstFld("T*")
'            cisprt.cLine = cisprt.cLine + 1
'        End If
'
'        wHSuryo = 0: wHUriGak = 0: wHinCnt = 0
'    End With
'
'    If Kb = 0 Then Exit Sub
'
'    B1_LINE.BorderWidth = 2
'    B1_LINE.Visible = True
'    Call cisprt.PrintFld("B1")
'    Call cisprt.ClrLstFld("B*")
'    With cisprt
''        .cDataCnt = .cDataCnt + 1
'        .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(mPrtCnt, "#,###")
'    End With
'    B1_LINE.BorderWidth = 1
'
'    T1_Line.Visible = False
'    T1_MIDASI = "《契約先計》"
'    T1_SU = Format(Format(wGSuryo, "#,##0;-#,##0;#"), String(9, "@"))
'    T1_UriGak = Format(Format(wGUriGak, "#,##0;-#,##0;#"), String(13, "@"))
'    Call cisprt.PrintFld("T1")
'    Call cisprt.ClrLstFld("T*")
'    cisprt.cLine = cisprt.cLine + 1
'
'    wGSuryo = 0: wGUriGak = 0: wNo = 0
'    If Kb = 1 Then
'       cisprt.cLine = 99
'    End If
'
'    If Kb = 1 Then Exit Sub
'
'    T1_MIDASI = "【総 合 計】"
'    T1_SU = Format(Format(wSSuryo, "#,##0;-#,##0;#"), String(9, "@"))
'    T1_UriGak = Format(Format(wSUriGak, "#,##0;-#,##0;#"), String(13, "@"))
'    Call cisprt.PrintFld("T1")
'
'End Sub
'---------------------------------------------------------------------------------------------------
        
        
'        With UJT
'            If wCnt <> 0 Then
'                If wTorcd <> .契約先 Then
'                    Call TotalWrite(1)
'                End If
'            End If
'
'            Call HeadWrite
'            wNo = wNo + 1
'
'            B1_No = CisFun.RSetFld(wNo, 5, "#,###")
'            If RTrim(wHinbn) <> RTrim(.品番) Then
'                If RTrim(wHinbn) <> "" Then
'                '品番合計
'                    Call TotalWrite(0)
'                End If
'                If Trim(.品番) <> "" Then
'                    B1_HinBn = .品番
'                End If
'                B1_LINE.Visible = True
'            Else
''
''< 2004.12.07 ADD START >-----------------------------------------------------------------------
''
'                If cisprt.cLine = 0 Then B1_HinBn = .品番
''
''< 2004.12.07 ADD  END  >-----------------------------------------------------------------------
''
'                B1_LINE.Visible = False
'            End If
'            If Trim(.検収年月日) <> "" Then
''< 2005/06/09 Delete START >-----------------------------------------------------------------------
''                With CisFun
''                    .DateE_BefVal = UJT.検収年月日
''                    .DateEdit
''                     B1_NYMD = .DateE_AfVal
''                End With
''< 2005/06/09 Delete End >-----------------------------------------------------------------------
''< 2005/06/09 ADD START >-----------------------------------------------------------------------
'                With CisFun
'                    .DateE_BefVal = UJT.検収年月日
'                    If Len(UJT.検収年月日) > 6 Then
'                       .DateEdit
'                       B1_NYMD = .DateE_AfVal
'                    Else
'                       B1_NYMD = Mid(UJT.検収年月日, 1, 4) & "/" & Mid(UJT.検収年月日, 5, 2)
'                    End If
'                End With
''< 2005/06/09 ADD End >-----------------------------------------------------------------------
'            End If
'            B1_NouNo = RTrim(.納入番号)
'            B1_SU = Format(Format(.納入数, "#,##0;-#,##0;#"), String(9, "@"))
'            B1_TANKA = Format(Format(.単価, "#,##0.00;-#,##0.00;#"), String(14, "@"))
'            B1_UriGak = Format(Format(.金額, "#,##0;-#,##0;#"), String(13, "@"))
'            If Trim(.仮単価区分) <> "" Then
'                B1_Kari = "*"
'            End If
'
'            Call cisprt.PrintFld("B1")
'
'            Call cisprt.ClrLstFld("B*")
'
'            With cisprt
'                .cLine = .cLine + 1
'                .cDataCnt = .cDataCnt + 1
'                .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(mPrtCnt, "#,###")
'            End With
'
'            wHinCnt = wHinCnt + 1
'            wHSuryo = wHSuryo + .納入数
'            wHUriGak = wHUriGak + .金額
'            wGSuryo = wGSuryo + .納入数
'            wGUriGak = wGUriGak + .金額
'            wSSuryo = wSSuryo + .納入数
'            wSUriGak = wSUriGak + .金額
'            wTorcd = .契約先
'            wHinbn = .品番
'        End With
'

