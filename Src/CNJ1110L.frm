VERSION 5.00
Begin VB.Form CNJ1110L 
   Caption         =   "îÑè„é¿ê—àÍóóï\"
   ClientHeight    =   3000
   ClientLeft      =   6750
   ClientTop       =   6045
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   11640
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "åÎç∑ó¶"
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
      Index           =   19
      Left            =   10470
      TabIndex        =   36
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(%)"
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
      Index           =   18
      Left            =   10740
      TabIndex        =   35
      Top             =   1800
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "êîó "
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
      Index           =   17
      Left            =   9000
      TabIndex        =   34
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "êîó "
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
      Left            =   8160
      TabIndex        =   33
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label B1_RITSU 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "ZZZZ9"
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
      Left            =   10590
      TabIndex        =   32
      Top             =   2220
      Width           =   540
   End
   Begin VB.Label B1_NOUNM 
      AutoSize        =   -1  'True
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
      Left            =   1230
      TabIndex        =   31
      Top             =   2220
      Width           =   2100
   End
   Begin VB.Label H1_YMD 
      AutoSize        =   -1  'True
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
      Left            =   9060
      TabIndex        =   11
      Top             =   510
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "%"
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
      Index           =   16
      Left            =   10980
      TabIndex        =   30
      Top             =   1200
      Width           =   105
   End
   Begin VB.Label H1_KYOYO 
      AutoSize        =   -1  'True
      Caption         =   "99"
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
      Left            =   10680
      TabIndex        =   29
      Top             =   1200
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ãñóeó¶ÅF"
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
      Index           =   15
      Left            =   9900
      TabIndex        =   28
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "î≠çsì˙ÅF"
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
      Left            =   8250
      TabIndex        =   27
      Top             =   510
      Width           =   840
   End
   Begin VB.Label H1_NAIYO 
      AutoSize        =   -1  'True
      Caption         =   "ëOâÒóÇåéì‡é¶ Å©Å® ìñåéì‡é¶"
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
      Left            =   6570
      TabIndex        =   26
      Top             =   1200
      Width           =   2850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "É`ÉFÉbÉNì‡óeÅF"
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
      Index           =   5
      Left            =   5145
      TabIndex        =   25
      Top             =   1200
      Width           =   1470
   End
   Begin VB.Label B1_KAZSA 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "ZZZ,ZZ9"
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
      Left            =   9750
      TabIndex        =   24
      Top             =   2220
      Width           =   735
   End
   Begin VB.Line H1_LINE 
      X1              =   300
      X2              =   11190
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   5
      Index           =   0
      Visible         =   0   'False
      X1              =   300
      X2              =   11175
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Label H1_MM 
      AutoSize        =   -1  'True
      Caption         =   "MM"
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
      Left            =   2640
      TabIndex        =   23
      Top             =   480
      Width           =   240
   End
   Begin VB.Label H1_YY 
      AutoSize        =   -1  'True
      Caption         =   "YYYY"
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
      Left            =   1560
      TabIndex        =   22
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "îN"
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
      Index           =   13
      Left            =   2235
      TabIndex        =   21
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "åéìx"
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
      Index           =   12
      Left            =   2970
      TabIndex        =   20
      Top             =   480
      Width           =   480
   End
   Begin VB.Label B1_SEBAN 
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
      Left            =   3900
      TabIndex        =   19
      Top             =   2220
      Width           =   840
   End
   Begin VB.Label B1_HINBN 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
      Left            =   4860
      TabIndex        =   18
      Top             =   2220
      Width           =   3150
   End
   Begin VB.Label B1_ZNKAZ 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "ZZZ,ZZ9"
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
      Left            =   8100
      TabIndex        =   17
      Top             =   2220
      Width           =   735
   End
   Begin VB.Label B1_KNKAZ 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "ZZZ,ZZ9"
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
      Left            =   8940
      TabIndex        =   16
      Top             =   2220
      Width           =   735
   End
   Begin VB.Label B1_UKEIR 
      AutoSize        =   -1  'True
      Caption         =   "XX"
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
      Left            =   3450
      TabIndex        =   15
      Top             =   2220
      Width           =   210
   End
   Begin VB.Label H1_TORCD 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
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
      Left            =   1380
      TabIndex        =   14
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label H1_TORNM 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
      Left            =   2310
      TabIndex        =   13
      Top             =   1200
      Width           =   2400
   End
   Begin VB.Label H1_PAGE 
      AutoSize        =   -1  'True
      Caption         =   "9,999"
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
      Left            =   10320
      TabIndex        =   12
      Top             =   510
      Width           =   525
   End
   Begin VB.Label B1_NOUCD 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
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
      Left            =   450
      TabIndex        =   10
      Top             =   2220
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "å_ñÒêÊÅF"
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
      Index           =   11
      Left            =   450
      TabIndex        =   9
      Top             =   1200
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "î[ì¸êÊ"
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
      Index           =   10
      Left            =   450
      TabIndex        =   8
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label1 
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
      Index           =   9
      Left            =   4020
      TabIndex        =   7
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label1 
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
      Index           =   8
      Left            =   5610
      TabIndex        =   6
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ëOâÒ"
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
      Left            =   8160
      TabIndex        =   5
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ç°âÒ"
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
      Left            =   9000
      TabIndex        =   4
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "êîó ç∑"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   4
      Left            =   9840
      TabIndex        =   3
      Top             =   1800
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "éÛì¸"
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
      Left            =   3360
      TabIndex        =   2
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ï≈"
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
      Index           =   1
      Left            =   10890
      TabIndex        =   1
      Top             =   510
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ì‡é¶êîó É`ÉFÉbÉNÉäÉXÉg"
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
      Left            =   3855
      TabIndex        =   0
      Top             =   360
      Width           =   3960
   End
   Begin VB.Line H1_K_LINE_ 
      BorderWidth     =   5
      Index           =   1
      Visible         =   0   'False
      X1              =   300
      X2              =   11190
      Y1              =   2730
      Y2              =   2730
   End
End
Attribute VB_Name = "CNJ1110L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  ì‡é¶êîó É`ÉFÉbÉNÉäÉXÉg
'**       ÉtÉHÅ[ÉÄID    :  CNJ1110L
'**       èàóùäTóv      :
'**
'**                         à¯êîÅiëºÉÇÉWÉÖÅ[ÉãÇÃÇ‡ÇÃÇó¨ópÅj
'**                             RV_YM       îNåé
'**                             Rv_LongStr  É`ÉFÉbÉNì‡óe
'**                             RV_Code     ãñóeó¶
'**                             RV_TorKb    ì‡óeãÊï™
'**
'**       çÏ  ê¨  ì˙    :  2009/10/14
'**       ïœ  çX  ì˙    :
'**
'**
'**********************************************************'
'** H1_MM     H1_YY     H1_YMD    H1_PAGE
'** H1_TORCD  H1_TORNM  H1_NAIYO  H1_KYOYO
'** B1_NOUCD  B1_NOUNM  B1_UKEIR  B1_SEBAN  B1_HINBN
'** B1_ZNKAZ  B1_KNKAZ  T1_KAZSA
'**********************************************************'
    Dim wCnt            As Long
    Dim wNo             As Long
    Dim OutStr          As String   ' CSVèoóÕì‡óe
    Dim mPrtCnt         As Long
    
    Dim wTorcd          As String
'    Dim wKouku          As String
'    Dim wHinbn          As String
'    Dim wHinCnt         As Long
'    Dim wHSuryo         As Currency
'    Dim wHUriGak        As Currency
'    Dim wKSuryo         As Currency '-------- 2008/06/06 ADD
'    Dim wKUriGak        As Currency '-------- 2008/06/06 ADD
'    Dim wGSuryo         As Currency
'    Dim wGUriGak        As Currency
'    Dim wSSuryo         As Currency
'    Dim wSUriGak        As Currency
'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub Form_Load()
    App.Title = "ì‡é¶êîó É`ÉFÉbÉNÉäÉXÉg"
'   If Not DataRead Then Exit Sub
    
    wTorcd = ""

    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    mPrtCnt = CisDB.RecordCount

    With CisPrt
        .iPrtTitle = "ì‡é¶êîó É`ÉFÉbÉNÉäÉXÉg"
        .iLineMax = 45
        .iLinePitch = 5
        .iPaperSize = p_A4
        .iOrientation = po_Portrait
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
'*-------------------------*
'*      àÛç¸ÉfÅ[É^ì«çû
'*-------------------------*
Private Function DataRead()
    DataRead = False
    
    DataRead = True
End Function
'*---------------------*
'*Å@  àÛ  ç¸  èà  óù
'*---------------------*
Private Function PrintProc()
    
    Call CisPrt.ClrLstFld("B*")
    CisPrt.cLine = 99
    wNo = 0: wTorcd = "" ': wHinbn = "": wHinCnt = 0
    Do Until Not NJT_RDSTS
        With NJT
            If wCnt <> 0 Then
                If wTorcd <> .å_ñÒêÊ Then
                   CisPrt.cLine = 99
                End If
            End If
            
            Call HeadWrite
            wNo = wNo + 1

            B1_NOUCD = .î[ì¸êÊ
            B1_NOUNM = .î[ì¸êÊñº
            B1_UKEIR = .éÛì¸
            B1_SEBAN = .îwî‘çÜ
            B1_HINBN = .ï\é¶ïiî‘
            B1_KNKAZ = Format(.ì‡é¶êî, "###,##0")
            If RV_TorKb = 0 Then
               B1_ZNKAZ = Format(.óÇåéì‡é¶êî, "###,##0")        ' ëOÇÃåéÇÃóÇåéì‡é¶êî
            B1_KAZSA.Caption = Format(.ì‡é¶êî - .óÇåéì‡é¶êî, "###,##0")
            Else
               B1_ZNKAZ = Format(.óÇÅXåéì‡é¶êî, "###,##0")      ' ëOÇÃåéÇÃì‡é¶êî
               B1_KAZSA.Caption = Format(.ì‡é¶êî - .óÇÅXåéì‡é¶êî, "###,##0")
            End If
            B1_RITSU = Format(.ì‡é¶î‰ó¶ * 100, "  0")

            Call CisPrt.PrintFld("B1")
            Call CisPrt.ClrLstFld("B*")
            
            With CisPrt
                .cLine = .cLine + 1
                .cDataCnt = .cDataCnt + 1
                .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(mPrtCnt, "#,###")
            End With
            
            wTorcd = .å_ñÒêÊ
        End With

        Call NJTReadNext
        wCnt = wCnt + 1
    Loop
    Call NJTClose
    Call HeadWrite
    
    ' àÛç¸èIóπ
    Call CisPrt.PrintEndDoc

End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite()
    HeadWrite = False
    
    With CisPrt
        ' ÇPï≈ì‡àÛç¸çsêî¡™Ø∏
        If .iLineMax >= .cLine + 1 Then Exit Function
        .cLine = 0
        .cPage = .cPage + 1
        
        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé
    
        With NJT
            H1_YY = Mid(RV_YM, 1, 4)                            ' îN
            H1_MM = Mid(RV_YM, 5, 2)                            ' åé
            H1_YMD = Format(Now(), "YYYY/MM/DD")                ' î≠çsì˙
            H1_PAGE = CisFun.RSetFld(CisPrt.cPage, 5, "#,###")  ' ÉyÅ[ÉW
            H1_TORCD = Trim(.å_ñÒêÊ)                            ' å_ñÒêÊ
            H1_TORNM = Trim(.å_ñÒêÊñº)                          ' å_ñÒêÊñº
            H1_NAIYO = Rv_LongStr                               ' É`ÉFÉbÉNì‡óe
            H1_KYOYO = Format(CisFun.Val2(RV_Code), " 0")       ' ãñóeó¶
        End With
    
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With
    
    HeadWrite = True
End Function
