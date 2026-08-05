VERSION 5.00
Begin VB.Form CZK0150L 
   Caption         =   "çﬁóøíIâµÉäÉXÉg"
   ClientHeight    =   3555
   ClientLeft      =   1860
   ClientTop       =   3165
   ClientWidth     =   15585
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
   ScaleHeight     =   3555
   ScaleWidth      =   15585
   Begin VB.Label B2_SumS 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "[ ëç  çá  åv ]"
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
      Left            =   6720
      TabIndex        =   27
      Top             =   1935
      Width           =   1695
   End
   Begin VB.Label B2_Kingaku 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "###,###,##0"
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
      Left            =   12705
      TabIndex        =   26
      Top             =   1935
      Width           =   1320
   End
   Begin VB.Line B1_LINE 
      BorderStyle     =   3  'ì_ê¸
      X1              =   0
      X2              =   15985
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Label B1_Saisyu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
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
      Left            =   14265
      TabIndex        =   25
      Top             =   1665
      Width           =   1200
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ç≈èIèoå…ì˙"
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
      Left            =   14235
      TabIndex        =   24
      Top             =   1395
      Width           =   1200
   End
   Begin VB.Label H1_TanaDay 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
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
      Left            =   6172
      TabIndex        =   23
      Top             =   795
      Width           =   3000
   End
   Begin VB.Label B1_Shiyo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
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
      Left            =   6735
      TabIndex        =   22
      Top             =   1665
      Width           =   2430
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "édÅ@Å@ól"
      BeginProperty Font 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   7470
      TabIndex        =   21
      Top             =   1395
      Width           =   810
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "çﬁ  éø  /  ê°Å@ñ@"
      BeginProperty Font 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   1020
      TabIndex        =   20
      Top             =   1395
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "éËîzêÊÅ@ÅF"
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
      TabIndex        =   19
      Top             =   705
      Width           =   1200
   End
   Begin VB.Label H1_ToriNm 
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
      Left            =   1560
      TabIndex        =   18
      Top             =   720
      Width           =   2400
   End
   Begin VB.Label H1_ZPlace 
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
      Left            =   1560
      TabIndex        =   17
      Top             =   435
      Width           =   2400
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
      Left            =   345
      TabIndex        =   16
      Top             =   420
      Width           =   1200
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ã‡äz"
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
      Left            =   13500
      TabIndex        =   15
      Top             =   1395
      Width           =   480
   End
   Begin VB.Label B1_Kingaku 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "###,###,##0"
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
      Left            =   12705
      TabIndex        =   14
      Top             =   1665
      Width           =   1320
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "íPâø"
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
      Left            =   12135
      TabIndex        =   13
      Top             =   1395
      Width           =   480
   End
   Begin VB.Label B1_Tanka 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "##,###,##0.00"
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
      Left            =   11115
      TabIndex        =   12
      Top             =   1665
      Width           =   1560
   End
   Begin VB.Label B1_Tani 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
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
      Left            =   10785
      TabIndex        =   11
      Top             =   1665
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "íIâµêî"
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
      Left            =   9990
      TabIndex        =   10
      Top             =   1395
      Width           =   720
   End
   Begin VB.Label B1_Tanasu 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "#,##0.00"
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
      Left            =   9780
      TabIndex        =   9
      Top             =   1665
      Width           =   960
   End
   Begin VB.Label B1_Zainm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
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
      Left            =   3690
      TabIndex        =   8
      Top             =   1665
      Width           =   3000
   End
   Begin VB.Label B1_Zaisun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
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
      Left            =   60
      TabIndex        =   7
      Top             =   1665
      Width           =   3600
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "çﬁ  óø  ñº"
      BeginProperty Font 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4455
      TabIndex        =   6
      Top             =   1395
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "(CZK0150L)"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   0
      TabIndex        =   5
      Top             =   60
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "Å¶Å@Å@çﬁ  óø  íIÅ@âµÅ@ÉäÅ@ÉXÅ@ÉgÅ@Å@Å¶"
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
      Left            =   4515
      TabIndex        =   4
      Top             =   105
      Width           =   6315
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
      Left            =   14805
      TabIndex        =   3
      Top             =   300
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "Page:"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14355
      TabIndex        =   2
      Top             =   300
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "î≠çsì˙:"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   12210
      TabIndex        =   1
      Top             =   300
      Width           =   570
   End
   Begin VB.Label H1_Date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "9999/99/99 "
      Height          =   225
      Left            =   12810
      TabIndex        =   0
      Top             =   270
      Width           =   1320
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   15985
      Y1              =   1635
      Y2              =   1635
   End
End
Attribute VB_Name = "CZK0150L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  çﬁóøíIâµÉäÉXÉgÅ@î≠çs
'**       ÉtÉHÅ[ÉÄID    :  CZK0150L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2010/01/15
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
    Dim mPrtDate        As String
    
    Dim wPlace          As String
    Dim wTori           As String
    Dim wTanaDay        As String
    Dim wGyoCnt         As Long
    Dim wKinS           As Long

'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub form_load()
    mPrtDate = Format(Now(), "yyyy/mm/dd")             'î≠çsì˙
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    With CisPrtCZK0150
        .iPrtTitle = "çﬁóøíIâµÉäÉXÉg"
        .iLineMax = 40
        .iLinePitch = 6
        .iPaperSize = p_A4
        .iOrientation = po_Landscape
        
        If Not .PrintIni Then
           With CisFun
                .MB_Lines = 4
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

    Dim WTanka As Currency

'*---------------------------------*
'*  èâä˙ê›íË
'*---------------------------------*
    wPlace = ""
    wTori = ""
    wTanaDay = ""
    
    wGyoCnt = 0
    wKinS = 0
'*--------------------------------------*
'*       ç›å…éÊà¯êÊÉ}ÉXÉ^ì«çûÇ›         *
'*--------------------------------------*
    Do Until Not ZTM_RDSTS
        If wPlace <> ZTM.ç›å…èÍèä Or wTori <> ZTM.éÊà¯êÊñº Or wTanaDay <> ZTM.é¿íIì˙ Then
            wPlace = ZTM.ç›å…èÍèä
            wTori = ZTM.éÊà¯êÊñº
            wTanaDay = ZTM.é¿íIì˙
            wGyoCnt = 0
            CisPrtCZK0150.cLine = CisPrtCZK0150.iLineMax
        End If

        Call HeadWrite(0)

        With CisPrtCZK0150
            .ClrLstFld ("B1")
        End With
        '*---------------------------------*
        '*  BodyçsÅ@ï“èW
        '*---------------------------------*
        With ZTM
            B1_Zaisun = .çﬁê°
            B1_Zainm = .çﬁóøñº
            B1_Shiyo = .édól
            B1_Tanasu = Format(.ç›å…êî, "#,##0")
            B1_Tani = .éËîzíPà ñº
            
            
            ' íPâøéZèoãyÇ—ã‡äzåvéZ
            With CisDB
                .SQL = "çﬁóøçwì¸íPâøåàíË"
                .StoadoCount = 9
                .ParaNo = 0:    .ParaIO = Returnóp
                .ParaNo = 1:    .ParaIO = Inputóp
                .ParaNo = 2:    .ParaIO = Inputóp
                .ParaNo = 3:    .ParaIO = Inputóp
                .ParaNo = 4:    .ParaIO = Inputóp
                .ParaNo = 5:    .ParaIO = OutPutóp
                .ParaNo = 6:    .ParaIO = OutPutóp
                .ParaNo = 7:    .ParaIO = OutPutóp
                .ParaNo = 8:    .ParaIO = OutPutóp
                .ParaNo = 9:    .ParaIO = OutPutóp
    
                .ParaNo = 1:    .ParaValue = ZTM.çﬁóøä«óùî‘çÜ               ' çﬁóøä«óùî‘çÜ
                .ParaNo = 2:    .ParaValue = RTrim(ZTM.éÊà¯êÊ)              ' éÊà¯êÊ
                .ParaNo = 3:    .ParaValue = ""                             ' å_ñÒêÊ
                If Len(RTrim(ZTM.é¿íIì˙)) = 6 Then
                    .ParaNo = 4:    .ParaValue = "01"
                Else
                    .ParaNo = 4:    .ParaValue = RTrim(ZTM.é¿íIì˙)          ' íPâøåàíËì˙ït
                End If
                
                .DBStored
    
               .ParaNo = 0
               If .ParaValue <> 0 Then
                   CisFun.MB_Lines = 5
                   CisFun.MB_MSG(2) = "    çﬁóøçwì¸íPâøåàíËÅ@àŸèÌÅI      "
                   CisFun.MB_MSG(4) = "    Error No = " & .ParaValue
                   CisFun.MB_Button = Error
                   CisFun.MBOX
                   End
               End If
    
                .ParaNo = 5
                WTanka = 0
                WTanka = RTrim(.ParaValue)
    
            End With
            
            B1_Tanka = Format(WTanka, "#,##0.00")
            B1_Kingaku = Format((WTanka * .ç›å…êî), "#,##0")
            
            If .ç≈èIèoå…ì˙ = "" Then
                B1_Saisyu = ""
            Else
                B1_Saisyu = Mid(.ç≈èIèoå…ì˙, 1, 4) & "/" & Mid(.ç≈èIèoå…ì˙, 5, 2) & "/" & Mid(.ç≈èIèoå…ì˙, 7, 2)
            End If
            wGyoCnt = wGyoCnt + 1
            wKinS = wKinS + (WTanka * .ç›å…êî)
            
            If Fix(wGyoCnt) Mod 2 = 0 Then
                If Not wGyoCnt = 40 Then
                    B1_LINE.Visible = True
                Else
                    B1_LINE.Visible = False
                    wGyoCnt = 0
                End If
            Else
                B1_LINE.Visible = False
            End If
            
            Call ZTMReadNext
            
            If (wPlace <> ZTM.ç›å…èÍèä Or wTori <> ZTM.éÊà¯êÊ Or wTanaDay <> ZTM.é¿íIì˙) Or Not ZTM_RDSTS Then
                B1_LINE.Visible = False
            End If
                        
        End With
        
        '*---------------------------------*
        '*  BodyïîÅ@àÛç¸
        '*---------------------------------*
        With CisPrtCZK0150
            .PrintFld ("B1*")
            .cLine = .cLine + 1
    '       ÉJÉEÉìÉ^Å@ÇtÇo
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(.cDataMaxCnt, "#,###")
        End With
    Loop
    Call HeadWrite(0)
    With CisPrtCZK0150
        .ClrLstFld ("B2")
        B2_SumS = "[ ëç  çá  åv ]"
        B2_Kingaku = Format(wKinS, "#,##0")
        wGyoCnt = wGyoCnt + 1
        '*---------------------------------*
        '*  Body3ïîÅ@àÛç¸
        '*---------------------------------*
        If Not .cLine = 0 Then .cLine = .cLine - 1
        .PrintFld ("B2*")
        .cLine = .cLine + 1
    End With
    Call ZTMClose(1)

    ' àÛç¸èIóπ
    Call CisPrtCZK0150.PrintEndDoc

End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite(KB As Byte)
    HeadWrite = False

    With CisPrtCZK0150
    '   ÇPï≈ì‡àÛç¸çsêî¡™Ø∏
        If .iLineMax >= .cLine + 1 Then Exit Function

        .cLine = 0
        .cPage = .cPage + 1

        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé

    '   î≠çsì˙ÅEï≈
        H1_Date = mPrtDate
        H1_Page = CisFun.RSetFld(.cPage, 5, "###")
        
        H1_ZPlace = wPlace
        H1_ToriNm = wTori
        If Len(wTanaDay) = 8 Then
            H1_TanaDay = "ÅyÅ@íIâµì˙  Å@ÅFÅ@" & Mid(wTanaDay, 1, 4) & "/" & Mid(wTanaDay, 5, 2) & "/" & Mid(wTanaDay, 7, 2) & "Å@Åz"
        Else
            H1_TanaDay = "ÅyÅ@íIâµîNåéÅ@ÅFÅ@" & Mid(wTanaDay, 1, 4) & "/" & Mid(wTanaDay, 5, 2) & "Å@Åz"
        End If

        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With

    HeadWrite = True
End Function

