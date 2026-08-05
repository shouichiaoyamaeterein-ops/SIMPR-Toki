VERSION 5.00
Begin VB.Form CZK0030L 
   Caption         =   "íIâµÉvÉåÉVÅ[Ég(çﬁóø)"
   ClientHeight    =   3615
   ClientLeft      =   2850
   ClientTop       =   2685
   ClientWidth     =   15810
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   15810
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "íIâµÉvÉåÉVÅ[Ég(çﬁóø)"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   24
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3825
      TabIndex        =   23
      Top             =   165
      Width           =   4965
   End
   Begin VB.Line B1_LINE 
      BorderStyle     =   3  'ì_ê¸
      BorderWidth     =   2
      X1              =   0
      X2              =   11650
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label H1_DenNo 
      AutoSize        =   -1  'True
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1350
      TabIndex        =   22
      Top             =   765
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ì`ï[áÇ"
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
      Left            =   150
      TabIndex        =   21
      Top             =   795
      Width           =   720
   End
   Begin VB.Label B1_TanaBan 
      AutoSize        =   -1  'True
      Caption         =   "XX-XX"
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
      Left            =   570
      TabIndex        =   20
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "íIî‘"
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
      Left            =   630
      TabIndex        =   19
      Top             =   1215
      Width           =   480
   End
   Begin VB.Label B1_Tani 
      AutoSize        =   -1  'True
      Caption         =   "XX"
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
      Left            =   5940
      TabIndex        =   18
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label B1_Zaiko 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "###,###"
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
      Left            =   5055
      TabIndex        =   17
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "óùò_ç›å…êî"
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
      Index           =   1
      Left            =   4950
      TabIndex        =   16
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label H1_Basyo 
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
      Index           =   1
      Left            =   1380
      TabIndex        =   15
      Top             =   450
      Width           =   2400
   End
   Begin VB.Label B1_ZaiNm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
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
      Left            =   8490
      TabIndex        =   14
      Top             =   1680
      Width           =   3000
   End
   Begin VB.Label B1_Tani 
      AutoSize        =   -1  'True
      Caption         =   "XX"
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
      Left            =   8070
      TabIndex        =   13
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label B1_Tanasu 
      AutoSize        =   -1  'True
      Caption         =   "(  Å@Å@Å@Å@Å@Åj"
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
      Left            =   6330
      TabIndex        =   12
      Top             =   1680
      Width           =   1800
   End
   Begin VB.Label B1_No 
      AutoSize        =   -1  'True
      Caption         =   "00"
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
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "áÇ"
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
      Index           =   0
      Left            =   135
      TabIndex        =   10
      Top             =   1215
      Width           =   240
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "î≠çsì˙ÅF"
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
      Left            =   8355
      TabIndex        =   9
      Top             =   750
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ç›å…èÍèäÅF"
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
      Index           =   11
      Left            =   165
      TabIndex        =   8
      Top             =   195
      Width           =   1200
   End
   Begin VB.Label H1_Basyo 
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
      Index           =   0
      Left            =   1380
      TabIndex        =   7
      Top             =   195
      Width           =   2400
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   2
      X1              =   -15
      X2              =   12985
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label H1_HYmd 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD"
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
      Left            =   9315
      TabIndex        =   6
      Top             =   750
      Width           =   1200
   End
   Begin VB.Label H1_Page 
      AutoSize        =   -1  'True
      Caption         =   "Z,ZZ9"
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
      Left            =   10770
      TabIndex        =   5
      Top             =   750
      Width           =   600
   End
   Begin VB.Label B1_Zaishitsu 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   3600
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "( íIÅ@âµÅ@êî Åj"
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
      Index           =   0
      Left            =   6330
      TabIndex        =   3
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label Label28 
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
      Left            =   11415
      TabIndex        =   2
      Top             =   750
      Width           =   225
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "çﬁÅ@Å@óø"
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
      Index           =   0
      Left            =   1350
      TabIndex        =   1
      Top             =   1215
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "çﬁÅ@óøÅ@ñº"
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
      Index           =   0
      Left            =   8490
      TabIndex        =   0
      Top             =   1200
      Width           =   1200
   End
End
Attribute VB_Name = "CZK0030L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  íIâµÉvÉåÉVÅ[Ég
'**       ÉtÉHÅ[ÉÄID    :  CZK0030L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2010/01/07
'**
'**********************************************************'
    Dim mPrtDate        As String
    
    Dim wInt            As Long

'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub form_load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")             'î≠çsì˙
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    With CisPrtCZK0030
        .iPrtTitle = "íIâµÉvÉåÉVÅ[ÉgÅiçﬁóøÅj"
        .iLineMax = 50
        .iLinePitch = 10
        .iLineInterval = 1
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
'*---------------------*
'*Å@  àÛ  ç¸  èà  óù
'*---------------------*
Private Function PrintProc()
'*---------------------------------*
'*      èâä˙ê›íË
'*---------------------------------*
    wInt = 0

'*--------------------------------------*
'*  çﬁóøíIâµÉvÉåÉVÅ[Égî≠çsÉeÅ[ÉuÉãì«çûÇ›
'*--------------------------------------*
    Do Until Not ZTP_RDSTS
        If wInt <> ZTP.ì`ï[NO Then
            wInt = ZTP.ì`ï[NO
            CisPrtCZK0030.cLine = CisPrtCZK0030.iLineMax
        End If

        Call HeadWrite(0)

        With CisPrtCZK0030
            .ClrLstFld ("B1")
        End With
        '*---------------------------------*
        '*  BodyçsÅ@ï“èW
        '*---------------------------------*
        With ZTP
            B1_No = .ì`ï[çsNO
            B1_TanaBan = .íIî‘
            B1_Zaishitsu = RTrim(.çﬁéø) & " " & Trim(.ï\é¶ê°ñ@)
            B1_Zaiko = Format(.ç›å…êî, "#,###")
            B1_Tani(0) = RTrim(.éËîzíPà )
            B1_Tani(1) = RTrim(.éËîzíPà )
            B1_Tanasu = "(  Å@Å@Å@Å@Å@Åj"
            B1_ZaiNm = RTrim(.çﬁóøñº)
            
            If Fix(.ì`ï[çsNO) Mod 2 = 0 Then
                If Not .ì`ï[çsNO = 50 Then
                    B1_LINE.Visible = True
                End If
            Else
                B1_LINE.Visible = False
            End If
            
            Call ZTPReadNext(1)
            
            If wInt <> ZTP.ì`ï[NO Or Not ZTP_RDSTS Then
                B1_LINE.Visible = False
            End If
        End With
        '*---------------------------------*
        '*  BodyïîÅ@àÛç¸
        '*---------------------------------*
        With CisPrtCZK0030
            .PrintFld ("B1*")
            .cLine = .cLine + 1
    '       ÉJÉEÉìÉ^Å@ÇtÇo
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With
    Loop
    Call ZTPClose(1)

    ' àÛç¸èIóπ
    Call CisPrtCZK0030.PrintEndDoc

End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite(KB As Byte)
    HeadWrite = False

    With CisPrtCZK0030
    '   ÇPï≈ì‡àÛç¸çsêî¡™Ø∏
        If .iLineMax >= .cLine + 1 Then Exit Function

        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé

    '   î≠çsì˙ÅEï≈
        H1_HYmd = mPrtDate
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")

        H1_DenNo = ZTP.ì`ï[NO
        H1_Basyo(0) = ZTP.ç›å…èÍèäñº1
        H1_Basyo(1) = ZTP.ç›å…èÍèäñº2

        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With

    HeadWrite = True
End Function
