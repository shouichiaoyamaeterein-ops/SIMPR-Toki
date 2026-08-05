VERSION 5.00
Begin VB.Form CXJ0230L 
   Caption         =   "ñæç◊èë"
   ClientHeight    =   3360
   ClientLeft      =   1125
   ClientTop       =   5370
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   11700
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
      Caption         =   "éÛíçãÊ"
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
      Left            =   3225
      TabIndex        =   28
      Top             =   1335
      Width           =   615
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
      Left            =   9915
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
      Left            =   10890
      TabIndex        =   26
      Top             =   510
      Width           =   225
   End
   Begin VB.Line B1_K_LINE2 
      X1              =   1740
      X2              =   11145
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
      Left            =   9465
      TabIndex        =   25
      Top             =   1710
      Width           =   630
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
      Left            =   9600
      TabIndex        =   24
      Top             =   1335
      Width           =   435
   End
   Begin VB.Line B1_K_LINE1 
      X1              =   195
      X2              =   11160
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Label B1_HSuryo 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
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
      Left            =   10320
      TabIndex        =   23
      Top             =   1710
      Width           =   735
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
      Left            =   8715
      TabIndex        =   22
      Top             =   1710
      Width           =   630
   End
   Begin VB.Label B1_Syuyo 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "999,999"
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
      Left            =   7875
      TabIndex        =   21
      Top             =   1710
      Width           =   735
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
      Left            =   7005
      TabIndex        =   20
      Top             =   1710
      Width           =   840
   End
   Begin VB.Label B1_Hinbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
      Left            =   3810
      TabIndex        =   19
      Top             =   1710
      Width           =   3150
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
      Left            =   2760
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
      Left            =   1830
      TabIndex        =   17
      Top             =   1710
      Width           =   525
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   1
      X1              =   195
      X2              =   11160
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
      Left            =   10410
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
      Left            =   8880
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
      Left            =   7995
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
      Left            =   7080
      TabIndex        =   13
      Top             =   1335
      Width           =   630
   End
   Begin VB.Label Label4 
      Caption         =   "ïi    î‘"
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
      Left            =   4170
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
      Left            =   2715
      TabIndex        =   11
      Top             =   1335
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "î[ì¸î‘çÜ"
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
      Left            =   1740
      TabIndex        =   10
      Top             =   1335
      Width           =   795
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   3
      Index           =   0
      X1              =   195
      X2              =   11160
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
      Left            =   405
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
      Left            =   3270
      TabIndex        =   7
      Top             =   1710
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "î[ì¸êÊ"
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
      Left            =   8835
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
      Left            =   9465
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
      Caption         =   "èoâ◊ÉvÉãÅ[ÉtÉäÉXÉg"
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
      Left            =   3855
      TabIndex        =   0
      Top             =   315
      Width           =   3390
   End
End
Attribute VB_Name = "CXJ0230L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  èoâ◊ÉvÉãÅ[ÉtÉäÉXÉgÅ@î≠çs
'**       ÉtÉHÅ[ÉÄID    :  CXJ0230L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2005/06/06
'**       ïœ  çX  ì˙    :  2008/04/10 By CIS îwî‘çÜ8åÖ
'**
'**********************************************************'
'#î‰ärëﬁîóp#
    Dim Key_Torcd       As String
    Dim Key_KTorcd      As String
    Dim Key_NYmd        As String
    Dim Key_Nouno       As String
    Dim SiirePage       As Long         'î[ì¸êÊåªï≈êî
    
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
    With CisPrtCXJ0230
        .iPrtTitle = "èoâ◊ÉvÉãÅ[ÉtÉäÉXÉg"
        .iLineMax = 60
        .iLinePitch = 6
        .iPaperSize = p_A4
'        .iOrientation = po_PORTRAIT
        
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
    Call CisPrtCXJ0230.ClrLstFld("B1")
    
'*---------------------------------*
'*  èoâ◊ÉeÅ[ÉuÉãì«çûÇ›
'*---------------------------------*
    Do Until Not SGW_RDSTS
        If Key_Torcd <> SGW.î[ì¸êÊ Or Key_KTorcd <> SGW.å_ñÒêÊ Or _
            CisPrtCXJ0230.cLine >= CisPrtCXJ0230.iLineMax Then
            If RTrim(Key_Torcd) <> "" Then
                B1_K_LINE1.Visible = True
                Call CisPrtCXJ0230.PrintFld("B1*")
                CisPrtCXJ0230.cLine = CisPrtCXJ0230.cLine + 1
                Call CisPrtCXJ0230.ClrLstFld("B1")
            End If
            Call HeadWrite(0)
        End If
        B1_K_LINE1.Visible = False: B1_K_LINE2.Visible = False
        With SGW
        '*---------------------------------*
        '*  BodyçsÅ@àÛç¸
        '*---------------------------------*
            If RTrim(Key_NYmd) <> RTrim(.î[ì¸ì˙) & RTrim(.î[ì¸ï÷) Then
                B1_K_LINE1.Visible = True
                With CisFun
                    .DateE_BefVal = RTrim(SGW.î[ì¸ì˙)
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
                B1_HCKbn = RTrim(.éÛíçãÊï™ñº)
            End If
            B1_Hinbn = RTrim(.ï\é¶ïiî‘)
            B1_Seban = RTrim(.îwî‘çÜ)
            If .ñáêî <> 0 Then
                B1_Syuyo = Format(.é˚óeêî, "##,##0")
            End If
            If .ñáêî <> 0 Then
                B1_Maisu = Format(.ñáêî, "##,##0")
'                GPMaisu = GPMaisu + .ñáêî
'                GTMaisu = GTMaisu + .ñáêî
            End If
            If .í[êî <> 0 Then
                B1_Hasu = Format(.í[êî, "##,##0")
            End If
            If .î[ì¸êî <> 0 Then
                B1_HSuryo = Format(.î[ì¸êî, "##,##0")
            End If
        End With
    'BODYïîàÛç¸
        Call CisPrtCXJ0230.PrintFld("B1*")
        CisPrtCXJ0230.cLine = CisPrtCXJ0230.cLine + 1
        Call CisPrtCXJ0230.ClrLstFld("B1")
            
        '*---------------------------------*
        '*  ÉJÉEÉìÉ^Å@ÇtÇo
        '*---------------------------------*
        With CisPrtCXJ0230
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(CXJ0230_PrtCnt, "#,###")
        End With
        
    'KEYçÄñ⁄ëﬁî
        Key_KTorcd = SGW.å_ñÒêÊ
        Key_Torcd = SGW.î[ì¸êÊ
        Key_NYmd = RTrim(SGW.î[ì¸ì˙) & RTrim(SGW.î[ì¸ï÷)
        Key_Nouno = SGW.î[ì¸î‘çÜ
        
        Call SGWReadNext(1)
    Loop
'ç≈èIï≈Å@ñáêîåvàÛç¸
    If RTrim(Key_Torcd) <> "" Then
        B1_K_LINE1.Visible = True
        Call CisPrtCXJ0230.PrintFld("B1*")
    End If
    
    Call SGWClose(1)
    
    ' àÛç¸èIóπ
    Call CisPrtCXJ0230.PrintEndDoc
    
End Function
'*---------------------------------------------------------------*
'*      ÉwÉbÉ_ïîàÛç¸                                                *
'*---------------------------------------------------------------*
Private Function HeadWrite(KB As Byte)
    Dim NewPageFlg As Boolean
    HeadWrite = False
    
    With CisPrtCXJ0230
        NewPageFlg = False
        .cLine = 0
        .cPage = .cPage + 1
        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé
    'ï≈
        If RTrim(Key_Torcd) <> RTrim(SGW.î[ì¸êÊ) Then
            SiirePage = 0
        End If
        SiirePage = SiirePage + 1
        H1_Page = Format(SiirePage, "##0")
    'î[ì¸êÊÅEî[ì¸êÊñº
        H1_Siire = RTrim(SGW.î[ì¸êÊCD)
        If RTrim(SGW.î[ì¸êÊçHèÍCD) <> "" Then
            H1_Siire = H1_Siire & " - " & RTrim(SGW.î[ì¸êÊçHèÍCD)
        End If
        H1_SiireNm = RTrim(SGW.î[ì¸êÊñº)
        H1_SiireKj = RTrim(SGW.î[ì¸êÊçHèÍñº)
    'î[ì¸êÊÅEî[ì¸êÊñº
        H1_KTorcd = RTrim(SGW.å_ñÒêÊ)
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
