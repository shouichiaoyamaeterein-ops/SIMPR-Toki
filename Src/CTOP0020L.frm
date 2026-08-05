VERSION 5.00
Begin VB.Form CTOP0020L 
   Caption         =   "TOPPSèoâ◊ïœä∑ÉGÉâÅ[ÉäÉXÉg"
   ClientHeight    =   4770
   ClientLeft      =   1320
   ClientTop       =   1455
   ClientWidth     =   16755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4770
   ScaleWidth      =   16755
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ERRì‡óe"
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
      Left            =   14550
      TabIndex        =   27
      Top             =   960
      Width           =   810
   End
   Begin VB.Label B1_ERR 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX1"
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
      Left            =   14340
      TabIndex        =   26
      Top             =   1455
      Width           =   2100
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "î“âÒì˙ - ï÷"
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
      Left            =   12855
      TabIndex        =   25
      Top             =   960
      Width           =   1260
   End
   Begin VB.Label B1_BDate 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD-XX"
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
      Left            =   12690
      TabIndex        =   24
      Top             =   1440
      Width           =   1560
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "î[ì¸ì˙ - ï÷"
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
      Left            =   9060
      TabIndex        =   23
      Top             =   960
      Width           =   1260
   End
   Begin VB.Label B1_NDate 
      AutoSize        =   -1  'True
      Caption         =   "YYYY/MM/DD-XX"
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
      Left            =   8835
      TabIndex        =   22
      Top             =   1440
      Width           =   1560
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "TEXT-ID"
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
      TabIndex        =   21
      Top             =   960
      Width           =   840
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "î[ì¸î‘çÜ"
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
      TabIndex        =   20
      Top             =   960
      Width           =   900
   End
   Begin VB.Label B1_Nouno 
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
      Left            =   2415
      TabIndex        =   19
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label B1_FName 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   150
      TabIndex        =   18
      Top             =   1440
      Width           =   1800
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "êî ó "
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
      Left            =   12000
      TabIndex        =   17
      Top             =   975
      Width           =   495
   End
   Begin VB.Label B1_Suryo 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "Z,ZZZ,ZZ9"
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
      Left            =   11520
      TabIndex        =   16
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label B1_Ukeir 
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
      Left            =   8370
      TabIndex        =   15
      Top             =   1440
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
      Left            =   11670
      TabIndex        =   14
      Top             =   405
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "î[  ì¸  êÊ"
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
      Left            =   6255
      TabIndex        =   13
      Top             =   945
      Width           =   1155
   End
   Begin VB.Label B1_Tornm 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXX"
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
      Left            =   6690
      TabIndex        =   12
      Top             =   1455
      Width           =   1575
   End
   Begin VB.Label B1_Torcd 
      AutoSize        =   -1  'True
      Caption         =   "XXXX"
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
      Left            =   6165
      TabIndex        =   11
      Top             =   1440
      Width           =   480
   End
   Begin VB.Line H1_K_LINE 
      X1              =   15
      X2              =   16200
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label B1_Seban 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXX"
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
      Left            =   5175
      TabIndex        =   10
      Top             =   1440
      Width           =   960
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
      Left            =   12840
      TabIndex        =   9
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
      Left            =   15315
      TabIndex        =   8
      Top             =   405
      Width           =   615
   End
   Begin VB.Label B1_Hinbn 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXX"
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
      Left            =   3330
      TabIndex        =   7
      Top             =   1440
      Width           =   1800
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "éÛ"
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
      Left            =   8370
      TabIndex        =   6
      Top             =   960
      Width           =   240
   End
   Begin VB.Label B1_Syuyo 
      Alignment       =   1  'âEëµÇ¶
      AutoSize        =   -1  'True
      Caption         =   "ZZZ,ZZ9"
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
      Left            =   10575
      TabIndex        =   5
      Top             =   1440
      Width           =   840
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "Å¶Å¶ TOPPSèoâ◊ïœä∑ÉGÉâÅ[ÉäÉXÉg Å¶Å¶"
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
      Left            =   3360
      TabIndex        =   4
      Top             =   285
      Width           =   5565
   End
   Begin VB.Label Label29 
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
      Left            =   5310
      TabIndex        =   3
      Top             =   960
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
      Left            =   15960
      TabIndex        =   2
      Top             =   405
      Width           =   225
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "ïi    î‘"
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
      Left            =   3690
      TabIndex        =   1
      Top             =   945
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "é˚óeêî"
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
      Left            =   10860
      TabIndex        =   0
      Top             =   975
      Width           =   630
   End
End
Attribute VB_Name = "CTOP0020L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  TOPPSéÛíçïœä∑ÉGÉâÅ[ÉäÉXÉg
'**       ÉtÉHÅ[ÉÄID    :  CTOP0010L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2004/04/20
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
'#î‰ärëﬁîóp#
    Dim BK_FName        As String
    Dim BK_Nouno        As String
        
    Dim lIx             As Byte         'ìYéö
    Dim TMai            As Long         'ñáêî
    
    Dim mNo(10)         As Long
    
    Dim mPrtDate        As String

'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg   *
'*------------------------------------*
Private Sub Form_Load()
    mPrtDate = Format(Now(), "yyyy/mm/dd hh:nn")             'î≠çsì˙ÅEéûä‘
    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    With CisPrt
        .iPrtTitle = "TOPPSèoâ◊ïœä∑ÉGÉâÅ[ÉäÉXÉg"
        .iLineMax = 40
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
    Dim FirstFLG    As Boolean      'èââÒâÒîóp(TRUE:Bodyçs àÛç¸)
    Dim BreakFLG    As Boolean      'KEY-BREAK(FALSE:keyBreak,TRUE:not_KeyBreak)
    
    lIx = 0: TMai = 0
    Erase mNo
    
    gSL_Select = "SELECT TJ.*,"
    gSL_Select = gSL_Select & " ISNULL(TR.ó™èÃ,'') î[ì¸êÊñº,"
    gSL_Select = gSL_Select & " ISNULL(HN.ï\é¶ïiî‘,'') ï\é¶ïiî‘"
    gSL_Select = gSL_Select & " FROM TOPPSèoâ◊ïœä∑ÉeÉLÉXÉg TJ "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN éÊà¯êÊÉ}ÉXÉ^ TR "
    gSL_Select = gSL_Select & "   ON TR.éÊà¯êÊãÊï™ = 0 "
    gSL_Select = gSL_Select & "  AND TR.éÊà¯êÊCD = TJ.î[ì¸êÊ "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN ïiî‘É}ÉXÉ^ HN "
    gSL_Select = gSL_Select & "   ON HN.ïiî‘ = TJ.ïiî‘ "
    gSL_Select = gSL_Select & " WHERE TJ.í[ññî‘çÜ = HOST_NAME() "
    gSL_Select = gSL_Select & "   AND TJ.ERR = '1' "
    gSL_Select = gSL_Select & " ORDER BY ÉtÉ@ÉCÉãID,î[ì¸î‘çÜ"
    Call PSWRead(gSL_Select)
    
    
    Do Until Not PSW_RDSTS
        With PSW
        
            Call CisPrt.ClrLstFld("B1")
            '*---------------------------------*
            '*  HeadïîÅ@àÛç¸
            '*---------------------------------*
                Call HeadWrite(0)
            
            If BK_FName <> .ÉtÉ@ÉCÉãID Then
                B1_FName = RTrim(.ÉtÉ@ÉCÉãID)
                BK_FName = RTrim(.ÉtÉ@ÉCÉãID)
                
                BK_Nouno = ""
            End If
            
            If BK_Nouno <> .î[ì¸î‘çÜ Then
                B1_Nouno = RTrim(.î[ì¸î‘çÜ)
                BK_Nouno = RTrim(.î[ì¸î‘çÜ)
            End If
            
            If RTrim(.ï\é¶ïiî‘) <> "" Then
                B1_Hinbn = RTrim(.ï\é¶ïiî‘)
            Else
                B1_Hinbn = RTrim(.ïiî‘)
            End If
            B1_Seban = RTrim(.îwî‘çÜ)
            B1_Torcd = RTrim(.î[ì¸êÊ)
            B1_Tornm = CisFun.Left2(RTrim(.î[ì¸êÊñº), 14)
            B1_Ukeir = RTrim(.éÛì¸)
            B1_Syuyo = CisFun.RSetFld(.é˚óeêî, 7, "#,###")
            B1_Suryo = CisFun.RSetFld(.èoâ◊å¬êî, 9, "###")
            
            gStr = ""
            If RTrim(PSW.î[ì¸ì˙) <> "" Then
                With CisFun
                    .DateE_BefVal = RTrim(PSW.î[ì¸ì˙)
                    .DateEdit
                    gStr = .DateE_AfVal
                End With
                If RTrim(PSW.î[ì¸ï÷) <> "" Then
                    gStr = gStr & "-" & Format(PSW.î[ì¸ï÷, "00")
                End If
            End If
            B1_NDate = gStr
            
            gStr = ""
            If RTrim(PSW.î“âÒì˙) <> "" Then
                With CisFun
                    .DateE_BefVal = RTrim(PSW.î“âÒì˙)
                    .DateEdit
                    gStr = .DateE_AfVal
                End With
                If RTrim(PSW.î“âÒï÷) <> "" Then
                    gStr = gStr & "-" & Format(PSW.î“âÒï÷, "00")
                End If
            End If
            B1_BDate = gStr
            
            Select Case RTrim(PSW.ERR)
                Case "1": B1_ERR = "éÛíçèÓïÒ ñ¢ìoò^"
                Case "2": B1_ERR = "èoâ◊ïœä∑çœÇ›"
                Case Else: B1_ERR = ""
            End Select
            '*---------------------------------*
            '*  BodyçsÅ@àÛç¸
            '*---------------------------------*
            With CisPrt
                Call .PrintFld("B1")
                .cLine = .cLine + 1
                Call .ClrLstFld("B1")
            End With
        
        End With
        
        Call PSWReadNext
    Loop
    
    Call PSWClose
    
    ' àÛç¸èIóπ
    Call CisPrt.PrintEndDoc
    
End Function
'*----------------*
'*  ÉwÉbÉ_ïîàÛç¸  *
'*----------------*
Private Function HeadWrite(KB As Byte)
    
    Dim NewPageFlg As Boolean
'
    HeadWrite = False
    With CisPrt
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
    
    BK_FName = "": BK_Nouno = ""
    HeadWrite = True
End Function
