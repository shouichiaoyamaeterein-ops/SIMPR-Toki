VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form CXM2540L 
   Caption         =   "ÉXÉNÉâÉbÉvíPâøïœìÆäzàÍóóï\"
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
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Alignment       =   1  'âEëµÇ¶
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
         Name            =   "ÇlÇr ñæí©"
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
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "îÑè„íPâøïœìÆäz"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "ç°âÒîÑè„íPâø"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "ëOâÒîÑè„íPâø"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "ç°âÒíPâøÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "ëOâÒíPâøÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "çﬁÅ@éøÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "éxï•êÊåv"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "ç°âÒΩ∏◊ØÃﬂíPâø"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "(YYYY/MM/DDÅ`YYYY/MM/DD)"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "ëŒè€ä˙ÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "YYYYîNè„ä˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
         Name            =   "ÇlÇr ñæí©"
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
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#,##0.00"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#,###,##0.00"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
         Name            =   "ÇlÇr ñæí©"
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
         Name            =   "ÇlÇr ñæí©"
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
         Name            =   "ÇlÇr ñæí©"
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
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "9,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "éxï•êÊÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "Ω∏◊ØÃﬂèdó "
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "ïiÅ@Å@î‘"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "ëOâÒΩ∏◊ØÃﬂíPâø"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "Ω∏◊ØÃﬂíPâøïœìÆäz"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "îwî‘çÜ"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "ï≈"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
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
      Caption         =   "Å¶ ÉXÉNÉâÉbÉvíPâøïœìÆäzàÍóóï\(éxï•) Å¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   4500
      TabIndex        =   0
      Top             =   345
      Width           =   6840
   End
End
Attribute VB_Name = "CXM2540L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  ÉXÉNÉâÉbÉvíPâøïœìÆäzàÍóóï\(éxï•)
'**       ÉtÉHÅ[ÉÄID    :  CXM2540L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2008/09/08
'**       ïœ  çX  ì˙    :
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
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub form_load()
    App.Title = "ÉXÉNÉâÉbÉvíPâøïœìÆäzàÍóóï\(éxï•)"
    
    wTorcd = ""
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    With CisPrtCXM2540
        .iPrtTitle = "ÉXÉNÉâÉbÉvíPâøïœìÆäzàÍóóï\"
        .iLineMax = 28
        .iLinePitch = 5
        .iPaperSize = p_A4
        .iOrientation = po_Landscape
                
        Set .iPrtForm = CXM2540L

        If Not .PrintIni Then
           With CisFun
                .MB_Lines = 5
                .MB_MSG(1) = "     Å@ÉvÉäÉìÉ^èâä˙ê›íËÉGÉâÅ[        "
                .MB_Title = "àÛç¸¥◊∞"
                .MB_Button = Error
                .MBOX
           End With
           Exit Sub
        .cLine = 0
        .cPage = 0
        End If
    End With
    
    PrtTim = Format(Now(), "yyyy/MM/dd HH:mm")      ' î≠çsì˙
    CXM2540_PrtOut = PrtCnt                         ' åèêî
    Call PrintProc
End Sub

'*-------------------------*
'*      àÛç¸ÉfÅ[É^ì«çû
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
'*Å@  àÛ  ç¸  èà  óù
'*---------------------*
Private Function PrintProc()
    '-- ÉfÅ[É^Çì«Ç›çûÇﬁ
    Adodc2.ConnectionString = CisDB.ConnectString
    Adodc2.RecordSource = gSL_Select
    Adodc2.Refresh
    
    If Adodc2.Recordset.RecordCount = 0 Then Exit Function
    PrtCnt = Adodc2.Recordset.RecordCount
    
    Kei_S = 0: Kei_U = 0
    PrtRow = 0
    With CisPrtCXM2540
        .ClrLstFld ("H1")
        .ClrLstFld ("B1")
        .cLine = 0
        .cPage = 0
        .cDataCnt = 0
    End With
    
    Do While Not Adodc2.Recordset.EOF
        Call DataRead                   ' ÇPçsì«çû
        
        If PrtDat(0) <> H1_TORCD Or PrtDat(1) <> H1_ZSiTu _
        Or CisPrtCXM2540.iLineMax < CisPrtCXM2540.cLine Then
           Call HeadWrite               ' ÉwÉbÉ_àÛç¸ Ç®ÇÊÇ— çáåvàÛç¸
        End If
        
        B1_HinBn = PrtDat(4)                        ' ïiî‘
        B1_SeBan = PrtDat(5)                        ' îwî‘çÜ
        B1_Juryo = Format(PrtDat(6), "#,##0.00")    ' ÉXÉNÉâÉbÉvèdó 
        B1_STanZ = Format(PrtDat(7), "#,##0.00")    ' ëOâÒÉXÉNÉâÉbÉvíPâø
        B1_STanK = Format(PrtDat(8), "#,##0.00")    ' ç°âÒÉXÉNÉâÉbÉvíPâø
        B1_STanH = Format(PrtDat(9), "#,##0.00")    ' ÉXÉNÉâÉbÉvíPâøïœìÆäz
        B1_UTanZ = Format(PrtDat(10), "#,##0.00")   ' ëOâÒîÑè„íPâø
        B1_UTanK = Format(PrtDat(11), "#,##0.00")   ' ç°âÒîÑè„íPâø
        B1_UTanH = Format(PrtDat(12), "#,##0.00")   ' îÑè„íPâøïœìÆäz
        Kei_S = Kei_S + Round(PrtDat(9), 4)         ' ÉXÉNÉâÉbÉvíPâøïœìÆäzçáåv
        Kei_U = Kei_U + Round(PrtDat(12), 4)        ' îÑè„íPâøïœìÆçáåv
        
        With CisPrtCXM2540
            Call .PrintFld("B1")        ' ÇPçsàÛç¸
            Call .ClrLstFld("B1")
            .cLine = .cLine + 1
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(PrtCnt, "#,###")
            DoEvents
        End With

        PrtRow = PrtRow + 1
        Adodc2.Recordset.MoveNext
    Loop
    Call TotalWrite                     ' ç≈èIÉyÅ[ÉWÇÃçáåvÇàÛç¸Ç∑ÇÈ

    ' àÛç¸èIóπ
    Call CisPrtCXM2540.PrintEndDoc
End Function

'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite()
    HeadWrite = False

    With CisPrtCXM2540
        '-- Ç¢ÇøÇŒÇÒèâÇﬂÇÃÉwÉbÉ_àÛç¸Ç≈Ç»ÇØÇÍÇŒçáåvÇàÛç¸ -------
        If .cDataCnt <> 0 Then
           If PrtDat(0) <> H1_TORCD _
           Or PrtDat(1) <> H1_ZSiTu Then    ' ÉyÅ[ÉWÇ‹ÇΩÇ¨É`ÉFÉbÉN
              Call TotalWrite               ' ìØéxï•êÊÇ»ÇÁÇŒçáåvÇÕàÛç¸ÇµÇ»Ç¢
           End If
        End If
        .cLine = 0
        .cPage = .cPage + 1
        If .cPage <> 1 Then         ' â¸ÉyÅ[ÉW
            .PrintNewPage
        End If

        .ClrLstFld ("H1")           ' HEADïîì‡óeè¡ãé
        H1_YMD = PrtTim             ' àÛç¸îNåéì˙éûï™
        H1_PAGE = .cPage            ' ÉyÅ[ÉW
        H1_NenKi = PrtNKi           ' ëŒè€ä˙
        H1_Kikan = PrtKKn           ' ä˙ä‘
        H1_TORCD = PrtDat(0)        ' éxï•êÊ
        H1_ZSiTu = PrtDat(1)        ' çﬁéø
        H1_ZenTK = PrtDat(2)        ' ëOâÒíPâø
        H1_KonTK = PrtDat(3)        ' ç°âÒíPâø
        If .iPreview = pv_Preview Then .PrintNewPage
        .PrintFld ("H1")            ' HEADïîàÛç¸
    End With

    HeadWrite = True
End Function

'*----------------*
'*   çáåvàÛç¸     *
'*----------------*
Private Sub TotalWrite(Optional Kb As Integer = 0)
   Dim iHosei  As Integer
    iHosei = 90     ' ï\é¶à íuï‚ê≥íl
    
    T1_Line1.Y1 = B1_HinBn.Top
    T1_Line1.Y2 = T1_Line1.Y1
    T1_Title.Top = T1_Line1.Y1 + iHosei
    T1_STKei.Top = T1_Title.Top
    T1_UTKei.Top = T1_Title.Top
    
    With CisPrtCXM2540
        T1_STKei = Format(Kei_S, "#,##0.00")
        T1_UTKei = Format(Kei_U, "#,##0.00")
        Call .PrintFld("T1")
    End With

    Kei_S = 0       ' èWåvílÉNÉäÉA
    Kei_U = 0
End Sub
