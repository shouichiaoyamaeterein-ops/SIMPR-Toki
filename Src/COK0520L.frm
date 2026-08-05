VERSION 5.00
Begin VB.Form COK0520L 
   Caption         =   "édì¸êÊï îÉä|é¿ê—"
   ClientHeight    =   5490
   ClientLeft      =   1860
   ClientTop       =   4020
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   15420
   Begin VB.Line B3_LINE 
      BorderWidth     =   3
      X1              =   300
      X2              =   16150
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Label B1_KSU3 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   9435
      TabIndex        =   36
      Top             =   2250
      Width           =   1155
   End
   Begin VB.Line B2_LINE 
      BorderWidth     =   4
      X1              =   300
      X2              =   16150
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Label B1_KDaySU 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   14400
      TabIndex        =   35
      Top             =   2250
      Width           =   1155
   End
   Begin VB.Label B1_KSU6 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   13095
      TabIndex        =   34
      Top             =   2250
      Width           =   1155
   End
   Begin VB.Label B1_KSU5 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   11865
      TabIndex        =   33
      Top             =   2250
      Width           =   1155
   End
   Begin VB.Label B1_KSU4 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   10650
      TabIndex        =   32
      Top             =   2250
      Width           =   1155
   End
   Begin VB.Label B1_KSU2 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   8205
      TabIndex        =   31
      Top             =   2250
      Width           =   1155
   End
   Begin VB.Label B1_KSU1 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   6960
      TabIndex        =   30
      Top             =   2250
      Width           =   1155
   End
   Begin VB.Label B1_KSC 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "##0.0%"
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
      Left            =   6225
      TabIndex        =   29
      Top             =   2250
      Width           =   630
   End
   Begin VB.Label B1_KJSU 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#,###,###,##0"
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
      Left            =   4755
      TabIndex        =   28
      Top             =   2250
      Width           =   1365
   End
   Begin VB.Label B1_KYSU 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#,###,###,##0"
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
      Left            =   3330
      TabIndex        =   27
      Top             =   2250
      Width           =   1365
   End
   Begin VB.Label B1_KojoNM 
      AutoSize        =   -1  'True
      Caption         =   "á@ çá åv"
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
      Left            =   360
      TabIndex        =   26
      Top             =   2235
      Width           =   960
   End
   Begin VB.Line B1_LINE 
      BorderStyle     =   3  'ì_ê¸
      X1              =   300
      X2              =   16150
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Label B1_DaySU 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   14400
      TabIndex        =   25
      Top             =   1935
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ì˙ìñÇË"
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
      Left            =   14910
      TabIndex        =   24
      Top             =   1485
      Width           =   720
   End
   Begin VB.Label B1_SU6 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   13095
      TabIndex        =   23
      Top             =   1935
      Width           =   1155
   End
   Begin VB.Label H1_MD6 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD"
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
      Left            =   13665
      TabIndex        =   22
      Top             =   1485
      Width           =   600
   End
   Begin VB.Label B1_SU5 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   11865
      TabIndex        =   21
      Top             =   1935
      Width           =   1155
   End
   Begin VB.Label H1_MD5 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD"
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
      Left            =   12435
      TabIndex        =   20
      Top             =   1485
      Width           =   600
   End
   Begin VB.Label B1_SU4 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   10650
      TabIndex        =   19
      Top             =   1935
      Width           =   1155
   End
   Begin VB.Label H1_MD4 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD"
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
      Left            =   11220
      TabIndex        =   18
      Top             =   1485
      Width           =   600
   End
   Begin VB.Label B1_SU3 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   9435
      TabIndex        =   17
      Top             =   1935
      Width           =   1155
   End
   Begin VB.Label H1_MD3 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD"
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
      Left            =   10005
      TabIndex        =   16
      Top             =   1485
      Width           =   600
   End
   Begin VB.Label B1_SU2 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   8205
      TabIndex        =   15
      Top             =   1935
      Width           =   1155
   End
   Begin VB.Label H1_MD2 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD"
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
      Left            =   8775
      TabIndex        =   14
      Top             =   1485
      Width           =   600
   End
   Begin VB.Label B1_SU1 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###,##0"
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
      Left            =   6960
      TabIndex        =   13
      Top             =   1935
      Width           =   1155
   End
   Begin VB.Label H1_MD1 
      AutoSize        =   -1  'True
      Caption         =   "MM/DD"
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
      Left            =   7530
      TabIndex        =   12
      Top             =   1485
      Width           =   600
   End
   Begin VB.Label B1_SC 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "##0.0%"
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
      Left            =   6225
      TabIndex        =   11
      Top             =   1935
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "êiíª"
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
      Left            =   6375
      TabIndex        =   10
      Top             =   1485
      Width           =   480
   End
   Begin VB.Label H1_JS 
      AutoSize        =   -1  'True
      Caption         =   "MMåéìxé¿ê—"
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
      Left            =   4950
      TabIndex        =   9
      Top             =   1485
      Width           =   1200
   End
   Begin VB.Label B1_JSU 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#,###,###,##0"
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
      Left            =   4755
      TabIndex        =   8
      Top             =   1935
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "å_ñÒêÊñº"
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
      Left            =   1155
      TabIndex        =   7
      Top             =   1485
      Width           =   960
   End
   Begin VB.Label B1_KTorNM 
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
      Left            =   1140
      TabIndex        =   6
      Top             =   1935
      Width           =   2100
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   1
      X1              =   300
      X2              =   16150
      Y1              =   1845
      Y2              =   1845
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   0
      X1              =   300
      X2              =   16150
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Label B1_KTorcd 
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
      Left            =   330
      TabIndex        =   5
      Top             =   1935
      Width           =   735
   End
   Begin VB.Label B1_YSU 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "#,###,###,##0"
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
      Left            =   3330
      TabIndex        =   4
      Top             =   1935
      Width           =   1365
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
      Left            =   14115
      TabIndex        =   3
      Top             =   555
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "å_ñÒêÊ"
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
      Left            =   330
      TabIndex        =   2
      Top             =   1485
      Width           =   720
   End
   Begin VB.Label H1_YS 
      AutoSize        =   -1  'True
      Caption         =   "MMåéìxó\ë™"
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
      Left            =   3525
      TabIndex        =   1
      Top             =   1485
      Width           =   1200
   End
   Begin VB.Label H1_Title 
      AutoSize        =   -1  'True
      Caption         =   "Å¶ YYYYîNMMåéìx édì¸êÊï îÉä|é¿ê— Å¶"
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
      Left            =   4710
      TabIndex        =   0
      Top             =   405
      Width           =   6300
   End
End
Attribute VB_Name = "COK0520L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  édì¸êÊï îÉä|é¿ê—ï\
'**       ÉtÉHÅ[ÉÄID    :  COK0520L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2008/05/22
'**       ïœÅ@çX  ì˙    :  2009/02/10
'**       ïœÅ@çX  ì˙    :  2009/02/12
'**
'**********************************************************'
    Dim wCnt            As Long
    Dim wNo             As Long
    Dim OutStr          As String   ' CSVèoóÕì‡óe
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
'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub Form_Load()
    App.Title = "îÉä|é¿ê—àÍóóï\"
'    If Not DataRead Then Exit Sub
    
    wTorcd = ""

    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    mPrtCnt = CisDB.RecordCount

    With CisPrt
        .iPrtTitle = "édì¸êÊï îÉä|é¿ê—àÍóóï\"
        .iLineMax = 30
        .iLinePitch = 5
        .iPaperSize = p_A4
        .iOrientation = po_Landscape
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
    
    Dim GKeiYs      As Currency     ' ó\ë™ã‡äz
    Dim GKeiJs      As Currency     ' é¿ê—ã‡äz
    Dim GKeiMD1     As Currency     ' XX/XX1
    Dim GKeiMD2     As Currency     ' XX/XX2
    Dim GKeiMD3     As Currency     ' XX/XX3
    Dim GKeiMD4     As Currency     ' XX/XX4
    Dim GKeiMD5     As Currency     ' XX/XX5
    Dim GKeiMD6     As Currency     ' XX/XX6
    Dim GKeiDA      As Currency     ' ì˙ìñÇË
    
    GKeiYs = 0: GKeiJs = 0
    GKeiMD1 = 0: GKeiMD2 = 0: GKeiMD3 = 0
    GKeiMD4 = 0: GKeiMD5 = 0: GKeiMD6 = 0
    GKeiDA = 0
    
    Call CisPrt.ClrLstFld("B*")
    CisPrt.cLine = 99
    wNo = 0: wTorcd = "": wHinbn = "": wHinCnt = 0
'    Call NYTRead(COK0520_SQL1)
    Call NYTRead(COK0520_SQL3)              '2009/02/10
    
    mPrtCnt = CisDB.RecordCount
    
    '+----------- å_ñÒêÊï îÉä|é¿ê— ------------+
    Do Until Not NYT_RDSTS
        With NYT
            Call HeadWrite
            
            B1_KTorcd = RTrim(.å_ñÒêÊ)
            B1_KTorNM = RTrim(.å_ñÒêÊñº)
            B1_YSU = Format(.ó\ë™ã‡äz, "#,###")
            B1_JSU = Format(.é¿ê—ã‡äz, "#,###")
            If .êiíª <> 0 Then
               B1_SC = Format(.êiíª, "#,##0.0;;0") & "%"
            End If
            B1_SU1 = Format(.é¿ê—ã‡äz0, "#,###")
            B1_SU2 = Format(.é¿ê—ã‡äz1, "#,###")
            B1_SU3 = Format(.é¿ê—ã‡äz2, "#,###")
            B1_SU4 = Format(.é¿ê—ã‡äz3, "#,###")
            B1_SU5 = Format(.é¿ê—ã‡äz4, "#,###")
            B1_SU6 = Format(.é¿ê—ã‡äz5, "#,###")
            B1_DaySU = Format(.ì˙ìñÇËêî, "#,###")
            
            GKeiYs = GKeiYs + NYT.ó\ë™ã‡äz
            GKeiJs = GKeiJs + NYT.é¿ê—ã‡äz
            GKeiMD1 = GKeiMD1 + NYT.é¿ê—ã‡äz0
            GKeiMD2 = GKeiMD2 + NYT.é¿ê—ã‡äz1
            GKeiMD3 = GKeiMD3 + NYT.é¿ê—ã‡äz2
            GKeiMD4 = GKeiMD4 + NYT.é¿ê—ã‡äz3
            GKeiMD5 = GKeiMD5 + NYT.é¿ê—ã‡äz4
            GKeiMD6 = GKeiMD6 + NYT.é¿ê—ã‡äz5
            GKeiDA = GKeiDA + NYT.ì˙ìñÇËêî
            
            B1_LINE.Visible = True
            
            Call CisPrt.PrintFld("B1")

            Call CisPrt.ClrLstFld("B*")
            
            With CisPrt
                .cLine = .cLine + 1
                .cDataCnt = .cDataCnt + 1
                .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(mPrtCnt, "#,###")
            End With
        End With

        Call NYTReadNext
        wCnt = wCnt + 1
    Loop
    
    '+------- å_ñÒêÊï Ç∆çHèÍï ÇÃãÊêÿÇË --------+
    CisPrt.cLine = CisPrt.cLine - 1
    
    B2_LINE.Visible = True
    Call CisPrt.PrintFld("B2")
    Call CisPrt.ClrLstFld("B*")
    
    '+------------- çáåvçsàÛç¸ ----------------+
    Call HeadWrite
    B1_KojoNM = " á@Å@  çá åv"
    B1_KYSU = Format(GKeiYs, "#,###")
    B1_KJSU = Format(GKeiJs, "#,###")
    If GKeiYs <> 0 And GKeiJs <> 0 Then
       B1_KSC = Format(CisFun.HasuProc(GKeiJs / GKeiYs * 100, éléÃå‹ì¸, 1), "#,##0.0;;0") & "%"
    End If
    B1_KSU1 = Format(GKeiMD1, "#,###")
    B1_KSU2 = Format(GKeiMD2, "#,###")
    B1_KSU3 = Format(GKeiMD3, "#,###")
    B1_KSU4 = Format(GKeiMD4, "#,###")
    B1_KSU5 = Format(GKeiMD5, "#,###")
    B1_KSU6 = Format(GKeiMD6, "#,###")
    B1_KDaySU = Format(GKeiDA, "#,###")
    
    Call CisPrt.PrintFld("B1")
    Call CisPrt.PrintFld("B3")

    Call CisPrt.ClrLstFld("B*")
    
    With CisPrt
        .cLine = .cLine + 1
    End With
    
    '+------------ çHèÍï îÉä|é¿ê— -------------+
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/02/12 UPDATE SART
    'Call NYTRead(COK0520_SQL2)
    Call NYTRead(COK0520_SQL4)
'Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢Å§Å¢ 2009/02/12 UPDATE END
    Do Until Not NYT_RDSTS
       With NYT
            Call HeadWrite
            
            B1_KojoNM = "       " & NYT.é©é–çHãÊñº
            B1_KYSU = Format(.ó\ë™ã‡äz, "#,###")
            B1_KJSU = Format(.é¿ê—ã‡äz, "#,###")
            If .é¿ê—ã‡äz <> 0 And .ó\ë™ã‡äz <> 0 Then
               B1_KSC = Format(CisFun.HasuProc(.é¿ê—ã‡äz / .ó\ë™ã‡äz * 100, éléÃå‹ì¸, 1), "#,##0.0;;0") & "%"
            End If
            B1_KSU1 = Format(.é¿ê—ã‡äz0, "#,###")
            B1_KSU2 = Format(.é¿ê—ã‡äz1, "#,###")
            B1_KSU3 = Format(.é¿ê—ã‡äz2, "#,###")
            B1_KSU4 = Format(.é¿ê—ã‡äz3, "#,###")
            B1_KSU5 = Format(.é¿ê—ã‡äz4, "#,###")
            B1_KSU6 = Format(.é¿ê—ã‡äz5, "#,###")
            B1_KDaySU = Format(.ì˙ìñÇËêî, "#,###")
            
            B3_LINE.Visible = True
            B1_LINE.Visible = False
            
            Call CisPrt.PrintFld("B1")
            Call CisPrt.PrintFld("B3")

            Call CisPrt.ClrLstFld("B*")
            
            With CisPrt
                .cLine = .cLine + 1
            End With
       End With
       
       Call NYTReadNext
       
    Loop
    
    Call NYTClose
    
    '+------- ç≈èIçsÇÃê¸ --------+
    B2_LINE.Visible = True
    Call CisPrt.PrintFld("B2")
    Call CisPrt.ClrLstFld("B*")
    
    ' àÛç¸èIóπ
    Call CisPrt.PrintEndDoc

End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite()
'
    HeadWrite = False
    With CisPrt
        ' ÇPï≈ì‡àÛç¸çsêî¡™Ø∏
        If .iLineMax >= .cLine + 1 Then Exit Function
        .cLine = 0
        .cPage = .cPage + 1
        
        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé
    
        With NYT
            H1_YMD = Format(Now(), "yyyy/MM/dd HH:mm")
        End With
        
        With CisFun
            H1_Title = "Å¶ " & .Mid2(COK0520_YM, 1, 4) & "îN" & .Mid2(COK0520_YM, 5, 2) & "åéìx édì¸êÊï îÉä|é¿ê— Å¶"
            H1_YS = .Mid2(COK0520_YM, 5, 2) & "åéìxó\ë™"
            H1_JS = .Mid2(COK0520_YM, 5, 2) & "åéìxé¿ê—"
            H1_MD1 = COK0520_MD1
            H1_MD2 = COK0520_MD2
            H1_MD3 = COK0520_MD3
            H1_MD4 = COK0520_MD4
            H1_MD5 = COK0520_MD5
            H1_MD6 = COK0520_MD6
        End With
    
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With
    
    HeadWrite = True
End Function
