VERSION 5.00
Begin VB.Form CXH0040L 
   Caption         =   "î≠íçë≈êÿÉäÉXÉg"
   ClientHeight    =   4770
   ClientLeft      =   675
   ClientTop       =   4740
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   12570
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "î[Å@ä˙"
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
      Left            =   6015
      TabIndex        =   23
      Top             =   1230
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "äÆóπì˙ÅEë≈êÿì˙"
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
      Left            =   9525
      TabIndex        =   22
      Top             =   1230
      Width           =   1680
   End
   Begin VB.Label B1_HSuryo 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "ZZ,ZZ9"
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
      Left            =   7200
      TabIndex        =   21
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label B1_NYmd 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99"
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
      Left            =   5790
      TabIndex        =   20
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Label B1_KANRYOBI 
      AutoSize        =   -1  'True
      Caption         =   "9999/99/99-99"
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
      Left            =   9675
      TabIndex        =   19
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "î≠íçêî"
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
      Left            =   7170
      TabIndex        =   18
      Top             =   1230
      Width           =   720
   End
   Begin VB.Label B1_UKEIR 
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
      Left            =   3180
      TabIndex        =   17
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "î≠çsì˙éûÅF"
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
      Left            =   7395
      TabIndex        =   16
      Top             =   405
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "éd  ì¸  êÊ"
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
      Left            =   795
      TabIndex        =   15
      Top             =   915
      Width           =   1200
   End
   Begin VB.Label B1_TORNM 
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
      Left            =   1020
      TabIndex        =   14
      Top             =   1680
      Width           =   2100
   End
   Begin VB.Label B1_TORCD 
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
      Left            =   255
      TabIndex        =   13
      Top             =   1680
      Width           =   735
   End
   Begin VB.Line H1_K_LINE 
      X1              =   255
      X2              =   11400
      Y1              =   1560
      Y2              =   1560
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
      Left            =   4050
      TabIndex        =   12
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label H1_HYMD 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD hh:ss"
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
      Left            =   8520
      TabIndex        =   11
      Top             =   405
      Width           =   1920
   End
   Begin VB.Label H1_PAGE 
      Alignment       =   1  'âEëµÇ¶
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
      Left            =   10425
      TabIndex        =   10
      Top             =   405
      Width           =   615
   End
   Begin VB.Label B1_HINBN 
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
      Left            =   1005
      TabIndex        =   9
      Top             =   1680
      Width           =   3000
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "éÛì¸"
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
      Left            =   3180
      TabIndex        =   8
      Top             =   915
      Width           =   480
   End
   Begin VB.Label B1_SETEI 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XX"
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
      Left            =   8280
      TabIndex        =   7
      Top             =   1680
      Width           =   1260
   End
   Begin VB.Label B1_SYUYO 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "ZZ,ZZ9"
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
      Left            =   4890
      TabIndex        =   6
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "Å¶ î≠íçë≈êÿäÆóπçœÉäÉXÉg Å¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2985
      TabIndex        =   5
      Top             =   285
      Width           =   4110
   End
   Begin VB.Label Label29 
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
      Left            =   3990
      TabIndex        =   4
      Top             =   1230
      Width           =   720
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
      Left            =   11100
      TabIndex        =   3
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
      Left            =   1545
      TabIndex        =   2
      Top             =   1230
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ê›íË"
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
      Left            =   8685
      TabIndex        =   1
      Top             =   1230
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "é˚óeêî"
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
      Left            =   4890
      TabIndex        =   0
      Top             =   1230
      Width           =   720
   End
End
Attribute VB_Name = "CXH0040L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  î≠íçë≈êÿäÆóπçœÉäÉXÉg
'**       ÉtÉHÅ[ÉÄID    :  CXH0040L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2004/03/26
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
'#î‰ärëﬁîóp#
    Dim BK_Torcd        As String       'édì¸êÊCD
    Dim BK_Ukeir        As String       'éÛì¸
    Dim BK_Hinbn        As String       'ïiî‘
    Dim BK_Seban        As String       'îwî‘çÜ
    Dim BK_Syuyo        As Long         'é˚óeêî
    Dim BK_UCSyu        As Byte         'ë≈êÿéÌï 
    
    Dim mPrtDate        As String
'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd hh:nn")             'î≠çsì˙ÅEéûä‘
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    With CisPrtCXH0040
        .iPrtTitle = "î≠íçë≈êÿäÆóπçœÉäÉXÉg"
        .iLineMax = 50
        .iLineInterval = 1
        .iLinePitch = 10
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
    
    Do Until Not KRE_RDSTS
        With KRE
            
        '*---------------------------------*
        '*  HeadïîÅ@àÛç¸
        '*---------------------------------*
            Call HeadWrite(0)
            
        '*---------------------------------*
        '*  BodyïîÅ@ÉNÉäÉA
        '*---------------------------------*
            Call CisPrtCXH0040.ClrLstFld("B1")
    
        '*---------------------------------*
        '*  àÛéöçÄñ⁄Å@ì]ëó
        '*---------------------------------*
            '#édì¸êÊCD,édì¸êÊó™èÃ
            If .éÊà¯êÊ <> BK_Torcd Then
                If CisPrtCXH0040.cLine <> 0 Then
                    Call CisPrtCXH0040.PrintFld("B1")
                    CisPrtCXH0040.cLine = CisPrtCXH0040.cLine + 1
                End If
                B1_TORCD = .éÊà¯êÊ
                B1_TORNM = .éÊà¯êÊñº
                BK_Ukeir = ""
            End If
            
            '#éÛì¸
            If .éÛì¸ <> BK_Ukeir Then
                B1_UKEIR = .éÛì¸
                BK_Hinbn = ""
                Call CisPrtCXH0040.PrintFld("B1")
                CisPrtCXH0040.cLine = CisPrtCXH0040.cLine + 1
            End If
                B1_TORCD = ""
                B1_TORNM = ""
                B1_UKEIR = ""
            '#ïiî‘,(îwî‘çÜ),é˚óeêî
            If .ïiî‘ <> BK_Hinbn Then
                B1_Seban = .îwî‘çÜ
                B1_HINBN = .ïiî‘
                BK_Seban = ""
            End If
            
            '#îwî‘çÜ
            If .îwî‘çÜ <> BK_Seban Then
                B1_Seban = .îwî‘çÜ
                BK_Syuyo = 0
                BK_UCSyu = 99
            End If
            
            '#é˚óeêî
            If .é˚óeêî <> BK_Syuyo Then
                B1_SYUYO = CisFun.RSetFld(.é˚óeêî, 6, "#,##0")
            End If
            
            '#î[ä˙
            CisFun.DateE_BefVal = Left(RV_SYmd, 6) & Format(.î≠çsNO, "00")
            Call CisFun.DateEdit
            B1_NYmd = CisFun.DateE_AfVal
            
            '#î≠íçêî
            B1_HSuryo = CisFun.RSetFld(.ç°âÒî≠íçêî, 6, "#,##0")
            
            If .ë≈êÿéÌï  <> BK_UCSyu Then
                '#ê›íË
                Select Case .ë≈êÿéÌï   '.ë≈êÿéÌï (0/écêîê›íËÅA1/ïiî‘íPà ÅA2/ïiî‘ÅEédì¸êÊíPà ÅA3/édì¸êÊíPà )
                    Case "0":
                        B1_SETEI = "écêî"
                    Case "1":
                        B1_SETEI = "ïiî‘"
                    Case "2":
                        B1_SETEI = "ïiî‘ÅEédì¸êÊ"
                    Case "3":
                        B1_SETEI = "édì¸êÊ"
                End Select
                
                If .ë≈êÿéÌï  = 0 Then
                '#äÆóπì˙-ï÷
                    If RTrim(.äÆóπì˙) <> "" Then
                        CisFun.DateE_BefVal = .äÆóπì˙
                        Call CisFun.DateEdit
                        B1_KANRYOBI = CisFun.DateE_AfVal
                        If RTrim(.äÆóπï÷) <> "" Then
                            B1_KANRYOBI = B1_KANRYOBI & "-" & .äÆóπï÷
                        End If
                    End If
                Else
                '#ë≈êÿì˙
                    If RTrim(.ë≈êÿì˙) <> "" Then
                        CisFun.DateE_BefVal = .ë≈êÿì˙
                        Call CisFun.DateEdit
                        B1_KANRYOBI = "(" & CisFun.DateE_AfVal & ")"
                    End If
                End If
            End If
            
        '*---------------------------------*
        '*  î‰ärçÄñ⁄Å@ëﬁî
        '*---------------------------------*
            BK_Torcd = .éÊà¯êÊ
            BK_Ukeir = .éÛì¸
            BK_Hinbn = .ïiî‘
            BK_Seban = .îwî‘çÜ
            BK_Syuyo = .é˚óeêî
            BK_UCSyu = .ë≈êÿéÌï 
            
        End With
    '*---------------------------------*
    '*  BodyçsÅ@àÛç¸
    '*---------------------------------*
        Call CisPrtCXH0040.PrintFld("B1")
                
        With CisPrtCXH0040
            .cLine = .cLine + 1
            .cDataCnt = .cDataCnt + 1
            .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(CXH0040_PrtCnt, "#,###")
        End With
        
        Call KREReadNext(1)
    Loop
    Call KREClose(1)
    
    ' àÛç¸èIóπ
    Call CisPrtCXH0040.PrintEndDoc
    
End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite(KB As Byte)
    
    Dim NewPageFlg As Boolean
'
    HeadWrite = False
    With CisPrtCXH0040
        ' ÇPï≈ì‡àÛç¸çsêî¡™Ø∏
        NewPageFlg = False
        
        If .iLineMax >= .cLine + 1 And Not NewPageFlg Then Exit Function
        
        .cLine = 0
        .cPage = .cPage + 1
        
        .ClrLstFld ("H1")   'HEADïîì‡óeè¡ãé
    
        H1_PAGE = CisFun.RSetFld(.cPage, 5, "###")
        H1_HYMD = mPrtDate
        
        
        
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With
    
    BK_Torcd = ""
    HeadWrite = True
End Function
