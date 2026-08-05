VERSION 5.00
Begin VB.Form CZK0110L 
   Caption         =   "çﬁóøç›å…ã‡äzàÍóóï\"
   ClientHeight    =   3735
   ClientLeft      =   -450
   ClientTop       =   480
   ClientWidth     =   16395
   BeginProperty Font 
      Name            =   "ÇlÇr ÉSÉVÉbÉN"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   16395
   Begin VB.Line L2_Line1 
      BorderStyle     =   3  'ì_ê¸
      X1              =   0
      X2              =   16320
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label H1_YM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "YYYYîNMMåé"
      BeginProperty Font 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2280
      TabIndex        =   35
      Top             =   300
      Width           =   1590
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ç≈èIì¸å…ì˙"
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
      Index           =   21
      Left            =   14745
      TabIndex        =   34
      Top             =   990
      Width           =   1050
   End
   Begin VB.Label B1_NDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "YYYY/MM/DD"
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
      Left            =   14760
      TabIndex        =   33
      Top             =   1245
      Width           =   1050
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "< åªç›ç›å… >"
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
      Index           =   18
      Left            =   13230
      TabIndex        =   32
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ã‡äz"
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
      Index           =   20
      Left            =   13920
      TabIndex        =   31
      Top             =   990
      Width           =   420
   End
   Begin VB.Label B1_GZKin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "999999"
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
      Left            =   13830
      TabIndex        =   30
      Top             =   1245
      Width           =   630
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "êîó "
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
      Index           =   19
      Left            =   13290
      TabIndex        =   29
      Top             =   990
      Width           =   420
   End
   Begin VB.Label B1_GZKaz 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "999999"
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
      Left            =   13170
      TabIndex        =   28
      Top             =   1245
      Width           =   630
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "< ìñåéèoå… >"
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
      Index           =   12
      Left            =   11730
      TabIndex        =   27
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "< ìñåéì¸å… >"
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
      Index           =   9
      Left            =   10260
      TabIndex        =   26
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "<ëOåéññç›å…>"
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
      Index           =   6
      Left            =   8790
      TabIndex        =   25
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ã‡äz"
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
      Index           =   14
      Left            =   12480
      TabIndex        =   24
      Top             =   990
      Width           =   420
   End
   Begin VB.Label B1_TOKin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "999999"
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
      Left            =   12360
      TabIndex        =   23
      Top             =   1245
      Width           =   630
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "êîó "
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
      Index           =   13
      Left            =   11820
      TabIndex        =   22
      Top             =   990
      Width           =   420
   End
   Begin VB.Label B1_TOKaz 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "999999"
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
      Left            =   11685
      TabIndex        =   21
      Top             =   1245
      Width           =   630
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ã‡äz"
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
      Index           =   11
      Left            =   10995
      TabIndex        =   20
      Top             =   990
      Width           =   420
   End
   Begin VB.Label B1_TIKin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "999999"
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
      Left            =   10890
      TabIndex        =   19
      Top             =   1245
      Width           =   630
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "êîó "
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
      Index           =   10
      Left            =   10350
      TabIndex        =   18
      Top             =   990
      Width           =   420
   End
   Begin VB.Label B1_TIKaz 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "999999"
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
      Left            =   10230
      TabIndex        =   17
      Top             =   1245
      Width           =   630
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ã‡äz"
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
      Index           =   8
      Left            =   9570
      TabIndex        =   16
      Top             =   990
      Width           =   420
   End
   Begin VB.Label B1_ZZKin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "999999"
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
      Left            =   9420
      TabIndex        =   15
      Top             =   1245
      Width           =   630
   End
   Begin VB.Label B1_Tornm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
      Left            =   5370
      TabIndex        =   14
      Top             =   1245
      Width           =   2100
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ê°ñ@"
      BeginProperty Font 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   3435
      TabIndex        =   13
      Top             =   990
      Width           =   390
   End
   Begin VB.Label B1_Sunpo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
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
      Left            =   2700
      TabIndex        =   12
      Top             =   1245
      Width           =   2625
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "édì¸íPâø"
      BeginProperty Font 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   7650
      TabIndex        =   11
      Top             =   990
      Width           =   780
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "éËîzêÊ"
      BeginProperty Font 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   5970
      TabIndex        =   10
      Top             =   990
      Width           =   585
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "Å¶Å@Å@çﬁ óø ç› å… ã‡ äz àÍ óó ï\Å@Å@Å¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   4605
      TabIndex        =   9
      Top             =   105
      Width           =   6840
   End
   Begin VB.Label H1_Page 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "9,999"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   15690
      TabIndex        =   8
      Top             =   300
      Width           =   450
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "PageÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   22
      Left            =   15180
      TabIndex        =   7
      Top             =   300
      Width           =   540
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "î≠çsì˙:"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   12255
      TabIndex        =   6
      Top             =   300
      Width           =   735
   End
   Begin VB.Label H1_Date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "9999/99/99 "
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   13020
      TabIndex        =   5
      Top             =   300
      Width           =   1155
   End
   Begin VB.Label B1_Hinbn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
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
      Left            =   45
      TabIndex        =   4
      Top             =   1245
      Width           =   2625
   End
   Begin VB.Label B1_Tanka 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "######.00"
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
      Left            =   7620
      TabIndex        =   3
      Top             =   1245
      Width           =   945
   End
   Begin VB.Label B1_ZZKaz 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "999999"
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
      Left            =   8730
      TabIndex        =   2
      Top             =   1245
      Width           =   630
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "çﬁéø"
      BeginProperty Font 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1095
      TabIndex        =   1
      Top             =   990
      Width           =   390
   End
   Begin VB.Label LBL_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "êîó "
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
      Index           =   7
      Left            =   8850
      TabIndex        =   0
      Top             =   990
      Width           =   420
   End
   Begin VB.Line L1_Line1 
      X1              =   45
      X2              =   16320
      Y1              =   1230
      Y2              =   1230
   End
End
Attribute VB_Name = "CZK0110L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  çﬁóøç›å…ã‡äzàÍóóï\àÛç¸
'**       ÉtÉHÅ[ÉÄID    :  CZK0110L
'**       èàóùäTóv      :  ópéÜ: A4â°
'**
'**       çÏ  ê¨  ì˙    :  2010/01/06
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
    Dim mPrtCnt         As Long
    Dim PrtTime         As String       ' î≠çsì˙
'    Dim wHin_Flg        As Byte         ' ïiî‘èÓïÒÉZÉbÉgÉtÉâÉO ' 2010.01.07 del
'    Dim wMSu(2)         As Long         ' åéó êîëﬁîóp         ' 2010.01.07 del
'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtCnt = CisDB.RecordCount                     'àÛç¸ëŒè€åèêîÅ@ëﬁî
  ''PrtTime = Format(Now(), "yyyy/mm/dd hh:nn")     'î≠çsì˙ÅEéûä‘
    PrtTime = Format(Now(), "yyyy/mm/dd")           'î≠çsì˙

'---------------------------------------------------' 2010.01.06 del start
''''    'ëŒè€îNåéëﬁî
''''    wNengetu(0) = TGet_YM
''''    With CisFun
''''        .Date_RsFormat = fYM
''''        .Date_Add TGet_YM, M, 1
''''
''''        wNengetu(1) = .Date_Result
''''
''''        .Date_RsFormat = fYM
''''        .Date_Add TGet_YM, M, 2
''''
''''        wNengetu(2) = .Date_Result
''''    End With
'---------------------------------------------------' 2010.01.06 del end

    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    With CisPrt
        .iPrtTitle = "çﬁóøç›å…ã‡äzàÍóóï\"           ' 2010.01.06 upd
        .iLineMax = 42
        .iPaperSize = p_A4
        .iOrientation = po_Landscape
        .iLineInterval = 0
''''    .iLinePitch = 6
        
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
    Call PrintProc
End Sub
'*---------------------*
'*Å@  àÛ  ç¸  èà  óù
'*---------------------*
Private Function PrintProc()
    Dim wCnt        As Integer
    Dim wPage       As Integer  ' 2010.01.07 add
    Dim wLine       As Integer  ' 2010.01.07 add
    
''''PgBk_Torcd = ""             ' 2010.01.07 del
''''PgBk_Syake = ""             ' 2010.01.07 del
''''Erase wMSu:    wHin_Flg = 0 ' 2010.01.07 del
''''Erase wKadou                ' 2010.01.06 del

    CisPrt.cLine = 0            ' 2010.01.07 add
    Do Until Not ZKW_RDSTS
        With ZKW
            With CisPrt
                .cDataCnt = .cDataCnt + 1
                .dpCountDisp Format(.cPage, "#,###") & " ï≈"
            End With
'-----------------------------------------------------------------------------------------------------------' 2010.01.06 add start
            With CisPrt
                If .cLine = 0 Then
                   Call HeadWrite
                End If
            End With
            
            B1_Hinbn = Mid(RTrim$(.çﬁéø), 1, 25)
            B1_Sunpo = Mid(RTrim$(.ê°ñ@), 1, 25)
            B1_Tornm = Mid(RTrim$(.éËîzêÊñº), 1, 20)
            B1_Tanka = CisFun.RSetFld(.édì¸íPâø, 9, "#####0.00")
            B1_ZZKaz = CisFun.RSetFld(.ëOåéññç›å…êî, 6, "#####0")
            B1_ZZKin = CisFun.RSetFld(.ëOåéññç›å…ã‡äz, 6, "#####0")
            B1_TIKaz = CisFun.RSetFld(.ìñåéì¸å…êî, 6, "#####0")
            B1_TIKin = CisFun.RSetFld(.ìñåéì¸å…ã‡äz, 6, "#####0")
            B1_TOKaz = CisFun.RSetFld(.ìñåéèoå…êî, 6, "#####0")
            B1_TOKin = CisFun.RSetFld(.ìñåéèoå…ã‡äz, 6, "#####0")
'            B1_CSKaz = CisFun.RSetFld(.í≤êÆêî, 6, "#####0")
'            B1_CSKin = CisFun.RSetFld(.í≤êÆã‡äz, 6, "#####0")
            B1_GZKaz = CisFun.RSetFld(.åªç›ç›å…êî, 6, "#####0")
            B1_GZKin = CisFun.RSetFld(.åªç›ç›å…ã‡äz, 6, "#####0")
            If Trim(.ç≈èIì¸å…ì˙) <> "" Then
               B1_NDate = Mid(.ç≈èIì¸å…ì˙, 1, 4) & "/" & Mid(.ç≈èIì¸å…ì˙, 5, 2) & "/" & Mid(.ç≈èIì¸å…ì˙, 7, 2)
            End If
'-----------------------------------------------------------------------------------------------------------' 2010.01.06 add end
                With CisPrt
                    'BodyçsàÛç¸
                    Call .PrintFld("B1")
                    .cLine = .cLine + 1
                    Call .ClrLstFld("B1")
                    'årê¸àÛç¸
                    If (.cLine Mod 2) = 0 Then                  ' 2010.01.07 upd : 5-->2
                        L2_Line1.BorderWidth = 1
                        .PrintFld "L2"                          ' 2010.07.07 upd : "L1"-->"L2"
                        L2_Line1.BorderWidth = 1
                    End If
                    If .cLine > .iLineMax - 1 Then .cLine = 0   ' 2010.01.07 add
                End With
                
                Call ZKWReadNext(1)                             ' 2010.01.07 add
'-----------------------------------------------------------------------------------------------------------' 2010.01.07 del start
'                Erase wMSu: wHin_Flg = 0
'                PrtK_Torcd = RTrim$(.å_ñÒêÊ)
'                PrtK_Hinbn = RTrim$(.ïiî‘)
'                PrtK_Syake = RTrim$(.é‘å^)
'                PrtK_Seban = RTrim$(.îwî‘çÜ)
'                PrtK_Syuyo = RTrim$(.é˚óeêî)
'            End If
'-----------------------------------------------------------------------------------------------------------' 2010.01.07 del end
        End With
    Loop
    
''''Call TailWrite      ' 2010.01.07 del
    Call ZKWClose

    ' àÛç¸èIóπ
    Call CisPrt.PrintEndDoc
End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite()
    HeadWrite = False
    
    With CisPrt
''''        If .iLineMax >= .cLine + 1 Then Exit Function           âºDEL

''''    If .cPage > 0 Then Call TailWrite           ' 2010.01.07 del
        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")       ' HEADïîì‡óeè¡ãé
        .ClrLstFld ("B1")
''''        .ClrLstFld ("Label")    ' ïsóvÉâÉxÉãì‡óeè¡ãé

        H1_YM = CZK0110_YM                              ' 2010.01.07 add : åƒÇ—èoÇµå≥Ç≈ÉZÉbÉgÇµÇƒÇ®Ç≠Ç±Ç∆
        H1_Page = CisFun.RSetFld(.cPage, 5, "#,###")
        H1_Date = PrtTime
        
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("LBL")          ' 2010.01.07 add
        .PrintFld ("H1")    'HEADïîàÛç¸
        
        'årê¸àÛç¸
        L1_Line1.BorderWidth = 1
        .PrintFld "L1"
        L1_Line1.BorderWidth = 1
    End With

    HeadWrite = True
End Function
'-----------------------------------------------------------------------------------------------------------' 2010.01.07 del start
''*----------------*
''*  ÉeÅ[ÉãïîàÛç¸  * Åi ï∂åæ Åj
''*----------------*
'Private Function TailWrite()
'
'    If CisPrt.cLine <> CisPrt.iLineMax Then CisPrt.cLine = CisPrt.iLineMax
'''''    Call CisPrt.PrintFld("T1")
'    Call CisPrt.PrintFld("LB")
'
'End Function
'-----------------------------------------------------------------------------------------------------------' 2010.01.07 del end
