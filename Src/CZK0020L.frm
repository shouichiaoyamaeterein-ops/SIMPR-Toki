VERSION 5.00
Begin VB.Form CZK0020L 
   Caption         =   "íIâµÉvÉåÉVÅ[Ég"
   ClientHeight    =   3870
   ClientLeft      =   3255
   ClientTop       =   3165
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   11370
   Begin VB.Label H1_kubun1 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXX"
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
      TabIndex        =   29
      Top             =   750
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ïîïiãÊï™ÅF"
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
      Left            =   165
      TabIndex        =   28
      Top             =   795
      Width           =   1200
   End
   Begin VB.Label H1_Kubun2 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXX"
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
      Left            =   2550
      TabIndex        =   27
      Top             =   750
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "îıÅ@Å@çl"
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
      Left            =   10140
      TabIndex        =   26
      Top             =   1455
      Width           =   960
   End
   Begin VB.Line B1_LINE 
      BorderStyle     =   3  'ì_ê¸
      X1              =   0
      X2              =   13000
      Y1              =   2145
      Y2              =   2145
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
      Left            =   450
      TabIndex        =   25
      Top             =   1920
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
      Left            =   495
      TabIndex        =   24
      Top             =   1470
      Width           =   480
   End
   Begin VB.Label B1_Seban 
      AutoSize        =   -1  'True
      Caption         =   "XXXXX"
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
      Left            =   1140
      TabIndex        =   23
      Top             =   1920
      Width           =   600
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "îwî‘çÜ"
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
      Left            =   1065
      TabIndex        =   22
      Top             =   1470
      Width           =   720
   End
   Begin VB.Label B1_Ko1 
      AutoSize        =   -1  'True
      Caption         =   "å¬"
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
      Left            =   5595
      TabIndex        =   21
      Top             =   1920
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
      Left            =   4710
      TabIndex        =   20
      Top             =   1920
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
      Left            =   4605
      TabIndex        =   19
      Top             =   1455
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
      TabIndex        =   18
      Top             =   450
      Width           =   2400
   End
   Begin VB.Label B1_Hinmoku 
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
      Left            =   7905
      TabIndex        =   17
      Top             =   1920
      Width           =   2445
   End
   Begin VB.Label B1_Ko2 
      AutoSize        =   -1  'True
      Caption         =   "å¬"
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
      Left            =   7575
      TabIndex        =   16
      Top             =   1920
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
      Left            =   5895
      TabIndex        =   15
      Top             =   1920
      Width           =   1800
   End
   Begin VB.Label B1_No 
      Alignment       =   1  'âEëµÇ¶
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
      Left            =   105
      TabIndex        =   14
      Top             =   1920
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
      Left            =   120
      TabIndex        =   13
      Top             =   1470
      Width           =   240
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
      TabIndex        =   12
      Top             =   1110
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
      Index           =   0
      Left            =   165
      TabIndex        =   11
      Top             =   1125
      Width           =   720
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
      Left            =   7905
      TabIndex        =   10
      Top             =   405
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   195
      Width           =   2400
   End
   Begin VB.Line H1_K_LINE 
      X1              =   0
      X2              =   13000
      Y1              =   1830
      Y2              =   1830
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
      Left            =   8865
      TabIndex        =   7
      Top             =   405
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
      Left            =   10320
      TabIndex        =   6
      Top             =   405
      Width           =   600
   End
   Begin VB.Label B1_Hinban 
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
      Left            =   2280
      TabIndex        =   5
      Top             =   1920
      Width           =   2430
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
      Left            =   5895
      TabIndex        =   4
      Top             =   1455
      Width           =   1800
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "íIâµÉvÉåÉVÅ[Ég"
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
      Left            =   3930
      TabIndex        =   3
      Top             =   165
      Width           =   3735
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
      Left            =   10965
      TabIndex        =   2
      Top             =   405
      Width           =   225
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "ïiÅ@Å@î‘"
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
      Left            =   2310
      TabIndex        =   1
      Top             =   1470
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ïiÅ@ñ⁄Å@ñº"
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
      Left            =   7920
      TabIndex        =   0
      Top             =   1455
      Width           =   1200
   End
End
Attribute VB_Name = "CZK0020L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  íIâµÉvÉåÉVÅ[Ég
'**       ÉtÉHÅ[ÉÄID    :  CZK0020L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2010/01/07
'**       ïœ  çX  ì˙    :  2011/01/20  à íuèCê≥
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
    With CisPrtCZK0020
        .iPrtTitle = "íIâµÉvÉåÉVÅ[Ég"
        .iLineMax = 60
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
        End If
        .cLine = 0
        .cPage = 0
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
    'wInt = TPH.ì`ï[NO
    wInt = 0
'*---------------------------------*
'*  íIâµÉvÉåÉVÅ[Égî≠çsÉeÅ[ÉuÉãì«çûÇ›
'*---------------------------------*
    Do Until Not TPH_RDSTS
        If wInt <> TPH.ì`ï[NO Then
            wInt = TPH.ì`ï[NO
            CisPrtCZK0020.cLine = CisPrtCZK0020.iLineMax
        End If

        Call HeadWrite(0)

        With CisPrtCZK0020
            .ClrLstFld ("B1")
        End With
        '*---------------------------------*
        '*  BodyçsÅ@ï“èW
        '*---------------------------------*
        With TPH
            B1_No = .ì`ï[çsNO
            B1_TanaBan = .íIî‘
            B1_Seban = .îwî‘çÜ
            B1_Hinban = .ï\é¶ïiî‘
            B1_Zaiko = Format(.ç›å…êî, "#,###")
            B1_Hinmoku = .ïiñº
            B1_Ko1 = "å¬"
            B1_Ko2 = "å¬"
            B1_Tanasu = "(  Å@Å@Å@Å@Å@Åj"
            
            If Fix(.ì`ï[çsNO) Mod 2 = 0 Then
                If Not .ì`ï[çsNO = 50 Then
                    B1_LINE.Visible = True
                End If
            Else
                B1_LINE.Visible = False
            End If
            
            Call TPHReadNext(1)
            
            If wInt <> TPH.ì`ï[NO Or Not TPH_RDSTS Then
                B1_LINE.Visible = False
            End If
        End With
        '*---------------------------------*
        '*  BodyïîÅ@àÛç¸
        '*---------------------------------*
        With CisPrtCZK0020
            .PrintFld ("B1*")
            .cLine = .cLine + 1
    '       ÉJÉEÉìÉ^Å@ÇtÇo
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With
    Loop
    Call TPHClose(1)

    ' àÛç¸èIóπ
    Call CisPrtCZK0020.PrintEndDoc

End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite(KB As Byte)
    HeadWrite = False

    With CisPrtCZK0020
    '   ÇPï≈ì‡àÛç¸çsêî¡™Ø∏
        If .iLineMax >= .cLine + 1 Then Exit Function

        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé

    '   î≠çsì˙ÅEï≈
        H1_HYmd = mPrtDate
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")

        H1_DenNo = TPH.ì`ï[NO
        H1_Basyo(0) = TPH.ç›å…èÍèäñº1
        H1_Basyo(1) = TPH.ç›å…èÍèäñº2
        
        If TPH.ïîïiãÊï™ = 0 Then
            H1_kubun1 = "äÆê¨ïi"
        Else
            If TPH.ïîïiãÊï™ = 1 Then
                H1_kubun1 = "çwì¸ïi"
            Else
                If TPH.ïîïiãÊï™ = 2 Then
                    H1_kubun1 = "édä|ïi"
                Else
                    H1_kubun1 = ""
                End If
            End If
        End If
        
        If TPH.é©ìÆé‘ÉKÉX = 0 Then
            H1_Kubun2 = "é©ìÆé‘"
        Else
            If TPH.é©ìÆé‘ÉKÉX = 1 Then
                H1_Kubun2 = "ÉKÉX"
            Else
                H1_Kubun2 = ""
            End If
        End If
        
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With

    HeadWrite = True
End Function

