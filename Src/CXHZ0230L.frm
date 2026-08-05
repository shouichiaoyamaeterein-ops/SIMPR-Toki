VERSION 5.00
Begin VB.Form CXHZ0230L 
   Caption         =   "ñæç◊èë"
   ClientHeight    =   3360
   ClientLeft      =   -825
   ClientTop       =   4590
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   15420
   Begin VB.Label B1_Tani 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
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
      Left            =   16065
      TabIndex        =   35
      Top             =   1710
      Width           =   420
   End
   Begin VB.Label B1_Zaisy 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
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
      Left            =   8370
      TabIndex        =   34
      Top             =   1710
      Width           =   2625
   End
   Begin VB.Label Label4 
      Caption         =   "çﬁÅ@éÌ"
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
      Index           =   5
      Left            =   8745
      TabIndex        =   33
      Top             =   1335
      Width           =   990
   End
   Begin VB.Label Label4 
      Caption         =   "ê°Å@ñ@"
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
      Index           =   2
      Left            =   6420
      TabIndex        =   32
      Top             =   1335
      Width           =   990
   End
   Begin VB.Label B1_Sunpo 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXX"
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
      Left            =   5925
      TabIndex        =   31
      Top             =   1710
      Width           =   2415
   End
   Begin VB.Label H1_KTorcd 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX-XX"
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
      Left            =   2595
      TabIndex        =   30
      Top             =   195
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "å_ñÒêÊ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1950
      TabIndex        =   29
      Top             =   180
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "î≠ãÊ"
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
      Left            =   2820
      TabIndex        =   28
      Top             =   1335
      Width           =   450
   End
   Begin VB.Label H1_Page 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "99"
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
      Left            =   14670
      TabIndex        =   27
      Top             =   495
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "ï≈"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   15645
      TabIndex        =   26
      Top             =   510
      Width           =   225
   End
   Begin VB.Line B1_K_LINE2 
      X1              =   1650
      X2              =   16400
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Line H1_K_LINE 
      Index           =   5
      X1              =   3285
      X2              =   3285
      Y1              =   390
      Y2              =   960
   End
   Begin VB.Line H1_K_LINE 
      Index           =   4
      X1              =   180
      X2              =   180
      Y1              =   390
      Y2              =   960
   End
   Begin VB.Line H1_K_LINE 
      Index           =   3
      X1              =   180
      X2              =   3285
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line H1_K_LINE 
      Index           =   2
      X1              =   180
      X2              =   3285
      Y1              =   390
      Y2              =   390
   End
   Begin VB.Label B1_Hasu 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "99,999.999"
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
      Left            =   13755
      TabIndex        =   25
      Top             =   1710
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "í[êî"
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
      Index           =   10
      Left            =   14265
      TabIndex        =   24
      Top             =   1335
      Width           =   435
   End
   Begin VB.Line B1_K_LINE1 
      X1              =   195
      X2              =   16400
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Label B1_HSuryo 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999.999"
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
      Left            =   14850
      TabIndex        =   23
      Top             =   1710
      Width           =   1155
   End
   Begin VB.Label B1_Maisu 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "99,999"
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
      Left            =   13095
      TabIndex        =   22
      Top             =   1710
      Width           =   630
   End
   Begin VB.Label B1_Syuyo 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999.999"
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
      Left            =   11925
      TabIndex        =   21
      Top             =   1710
      Width           =   1155
   End
   Begin VB.Label B1_Seban 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXX"
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
      Left            =   11055
      TabIndex        =   20
      Top             =   1710
      Width           =   840
   End
   Begin VB.Label B1_Zaist 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
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
      Left            =   3270
      TabIndex        =   19
      Top             =   1710
      Width           =   2625
   End
   Begin VB.Label B1_Ukeir 
      AutoSize        =   -1  'True
      Caption         =   "XX"
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
      Left            =   2325
      TabIndex        =   18
      Top             =   1710
      Width           =   210
   End
   Begin VB.Label B1_Nouno 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
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
      Left            =   1650
      TabIndex        =   17
      Top             =   1710
      Width           =   525
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   1
      X1              =   195
      X2              =   16400
      Y1              =   1635
      Y2              =   1635
   End
   Begin VB.Label Label4 
      Caption         =   "î[ì¸êî"
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
      Index           =   4
      Left            =   15090
      TabIndex        =   16
      Top             =   1335
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "î†êî"
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
      Index           =   7
      Left            =   13260
      TabIndex        =   15
      Top             =   1335
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "é˚óeêî"
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
      Index           =   6
      Left            =   12330
      TabIndex        =   14
      Top             =   1335
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "îwî‘çÜ"
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
      Index           =   5
      Left            =   11100
      TabIndex        =   13
      Top             =   1335
      Width           =   630
   End
   Begin VB.Label Label4 
      Caption         =   "çﬁÅ@éø"
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
      Index           =   3
      Left            =   3720
      TabIndex        =   12
      Top             =   1335
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "éÛì¸"
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
      Index           =   4
      Left            =   2340
      TabIndex        =   11
      Top             =   1335
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "î[ïièëáÇ"
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
      Index           =   3
      Left            =   1650
      TabIndex        =   10
      Top             =   1335
      Width           =   645
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   0
      X1              =   195
      X2              =   16400
      Y1              =   1185
      Y2              =   1185
   End
   Begin VB.Label Label4 
      Caption         =   "î[ ì¸ ì˙ - ï÷"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   1335
      Width           =   1245
   End
   Begin VB.Label B1_NYmd 
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99-99"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1710
      Width           =   1365
   End
   Begin VB.Label B1_HCKbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
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
      Left            =   2805
      TabIndex        =   7
      Top             =   1710
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "édì¸êÊ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   6
      Top             =   180
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "î≠çsì˙"
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
      Index           =   11
      Left            =   13590
      TabIndex        =   5
      Top             =   255
      Width           =   540
   End
   Begin VB.Label H1_HYmd 
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99 hh:mm"
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
      Left            =   14220
      TabIndex        =   4
      Top             =   225
      Width           =   1680
   End
   Begin VB.Label H1_SiireKj 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   255
      TabIndex        =   3
      Top             =   705
      Width           =   2400
   End
   Begin VB.Label H1_SiireNm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   255
      TabIndex        =   2
      Top             =   435
      Width           =   2400
   End
   Begin VB.Label H1_Siire 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX-XX"
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
      Left            =   870
      TabIndex        =   1
      Top             =   195
      Width           =   720
   End
   Begin VB.Label Label4 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "çﬁóøì¸â◊ÉvÉãÅ[ÉtÉäÉXÉg"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   5805
      TabIndex        =   0
      Top             =   315
      Width           =   4140
   End
End
Attribute VB_Name = "CXHZ0230L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  [çﬁóø]ì¸â◊ÉvÉãÅ[ÉtÉäÉXÉgÅ@î≠çs
'**       ÉtÉHÅ[ÉÄID    :  CXHZ0230L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2004/02/11
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
'#î‰ärëﬁîóp#
    Dim Key_Torcd       As String
    Dim Key_KTorcd      As String
    Dim Key_NYmd        As String
    Dim Key_Nouno       As String
    Dim SiirePage       As Long         'édì¸êÊåªï≈êî
    
    Dim mPrtDate        As String
    Dim mPrtTime        As String

'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")             'î≠çsì˙
    mPrtTime = Format(Now(), "hh:nn")                  'î≠çséûä‘
    
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    With CisPrtCXHZ0230
        .iPrtTitle = "[çﬁóø]ì¸â◊ÉvÉãÅ[ÉtÉäÉXÉg"
        .iLineMax = 38
        .iLinePitch = 6
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
'*---------------------------------------------------------------*
'*Å@ << àÛ  ç¸  èà  óù >>
'*---------------------------------------------------------------*
Private Function PrintProc()
'*---------------------------------*
'*  èâä˙ê›íË
'*---------------------------------*
    Key_Torcd = "": Key_KTorcd = "": Key_Nouno = "": Key_NYmd = ""
    Call CisPrtCXHZ0230.ClrLstFld("B1")
    
'*---------------------------------*
'*  î≠íçÉeÅ[ÉuÉãì«çûÇ›
'*---------------------------------*
    Do Until Not ZNG_RDSTS
        If Key_Torcd <> ZNG.édì¸êÊ Or Key_KTorcd <> ZNG.å_ñÒêÊ Or _
            CisPrtCXHZ0230.cLine >= CisPrtCXHZ0230.iLineMax Then
            If RTrim(Key_Torcd) <> "" Then
                B1_K_LINE1.Visible = True
                Call CisPrtCXHZ0230.PrintFld("B1*")
                CisPrtCXHZ0230.cLine = CisPrtCXHZ0230.cLine + 1
                Call CisPrtCXHZ0230.ClrLstFld("B1")
            End If
            Call HeadWrite(0)
        End If
        B1_K_LINE1.Visible = False: B1_K_LINE2.Visible = False
        With ZNG
        '*---------------------------------*
        '*  BodyçsÅ@àÛç¸
        '*---------------------------------*
            If RTrim(Key_NYmd) <> RTrim(.î[ì¸ì˙) & RTrim(.î[ì¸ï÷) Then
                B1_K_LINE1.Visible = True
                With CisFun
                    .DateE_BefVal = RTrim(ZNG.î[ì¸ì˙)
                    .DateEdit Slash
                    B1_NYmd = RTrim(.DateE_AfVal)
                End With
                If RTrim(.î[ì¸ï÷) <> "" Then
                    B1_NYmd = RTrim(B1_NYmd) & "-" & RTrim(.î[ì¸ï÷)
                End If
            End If
            If RTrim(Key_NYmd) <> RTrim(.î[ì¸ì˙) & RTrim(.î[ì¸ï÷) Or RTrim(Key_Nouno) <> RTrim(.î[ì¸î‘çÜ) Then
                B1_K_LINE2.Visible = True
                B1_Nouno = RTrim(.î[ì¸î‘çÜ)
                B1_Ukeir = RTrim(.éÛì¸)
                B1_HCKbn = RTrim(.î≠íçãÊï™ñº)
            End If
            B1_Zaist = RTrim(.çﬁéø)
            B1_Sunpo = RTrim(.ï\é¶ê°ñ@)
            gStr = RTrim(.çﬁéÌÉRÅ[Éh)
            If RTrim(.çﬁéÌÉRÅ[Éh) <> "" And RTrim(.çﬁéÌñº) <> "" Then
                gStr = gStr & ":" & RTrim(.çﬁéÌñº)
            Else
                gStr = gStr & RTrim(.çﬁéÌñº)
            End If
            B1_Zaisy = CisFun.Left2(gStr, 25)
            B1_Seban = RTrim(.îwî‘çÜ)
            If .ñáêî <> 0 Then
                B1_Syuyo = Format(.é˚óeêî, "##,##0.000")
            End If
            If .ñáêî <> 0 Then
                B1_Maisu = Format(.ñáêî, "##,##0")
'                GPMaisu = GPMaisu + .ñáêî
'                GTMaisu = GTMaisu + .ñáêî
            End If
            If .í[êî <> 0 Then
                B1_Hasu = Format(.í[êî, "##,##0.000")
            End If
            If .î≠íçêî <> 0 Then
                B1_HSuryo = Format(.î≠íçêî, "##,##0.000")
            End If
            B1_Tani = RTrim(.éËîzíPà ñº)
        End With
    'BODYïîàÛç¸
        Call CisPrtCXHZ0230.PrintFld("B1*")
        CisPrtCXHZ0230.cLine = CisPrtCXHZ0230.cLine + 1
        Call CisPrtCXHZ0230.ClrLstFld("B1")
            
        '*---------------------------------*
        '*  ÉJÉEÉìÉ^Å@ÇtÇo
        '*---------------------------------*
        With CisPrtCXHZ0230
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(CXHZ0230_PrtCnt, "#,###")
        End With
        
    'KEYçÄñ⁄ëﬁî
        Key_KTorcd = ZNG.å_ñÒêÊ
        Key_Torcd = ZNG.édì¸êÊ
        Key_NYmd = RTrim(ZNG.î[ì¸ì˙) & RTrim(ZNG.î[ì¸ï÷)
        Key_Nouno = ZNG.î[ì¸î‘çÜ
        
        Call ZNGReadNext(1)
    Loop
'ç≈èIï≈Å@ñáêîåvàÛç¸
    If RTrim(Key_Torcd) <> "" Then
        B1_K_LINE1.Visible = True
        Call CisPrtCXHZ0230.PrintFld("B1*")
    End If
    
    Call ZNGClose(1)
    
    ' àÛç¸èIóπ
    Call CisPrtCXHZ0230.PrintEndDoc
    
End Function
'*---------------------------------------------------------------*
'*      ÉwÉbÉ_ïîàÛç¸                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim NewPageFlg As Boolean
    HeadWrite = False
    
    With CisPrtCXHZ0230
        NewPageFlg = False
        .cLine = 0
        .cPage = .cPage + 1
        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé
    'ï≈
        If RTrim(Key_Torcd) <> RTrim(ZNG.édì¸êÊ) Then
            SiirePage = 0
        End If
        SiirePage = SiirePage + 1
        H1_Page = Format(SiirePage, "##0")
    'édì¸êÊÅEédì¸êÊñº
        H1_Siire = RTrim(ZNG.édì¸êÊCD)
        If RTrim(ZNG.édì¸êÊçHèÍCD) <> "" Then
            H1_Siire = H1_Siire & " - " & RTrim(ZNG.édì¸êÊçHèÍCD)
        End If
        H1_SiireNm = RTrim(ZNG.édì¸êÊñº)
        H1_SiireKj = RTrim(ZNG.édì¸êÊçHèÍñº)
    'édì¸êÊÅEédì¸êÊñº
        H1_KTorcd = RTrim(ZNG.å_ñÒêÊ)
    'î≠çsì˙ÅEé©é–ñº
        H1_HYmd = mPrtDate & " " & mPrtTime
'
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With
'
    Key_NYmd = ""
    HeadWrite = True
End Function
