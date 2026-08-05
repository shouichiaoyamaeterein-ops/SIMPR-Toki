VERSION 5.00
Begin VB.Form CXU0125L 
   Caption         =   "îÑè„é¿ê—ï\"
   ClientHeight    =   3330
   ClientLeft      =   5025
   ClientTop       =   6105
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3330
   ScaleWidth      =   11640
   Begin VB.Label B1_Torcd 
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
      Left            =   990
      TabIndex        =   21
      Top             =   2220
      Width           =   735
   End
   Begin VB.Label H1_Jisya 
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
      Left            =   945
      TabIndex        =   20
      Top             =   1185
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "çHãÊÅF"
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
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   1185
      Width           =   720
   End
   Begin VB.Line T1_Line 
      X1              =   4800
      X2              =   11190
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Label T1_UriGak 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999,999"
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
      Left            =   9135
      TabIndex        =   18
      Top             =   2205
      Width           =   1365
   End
   Begin VB.Label T1_SU 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Left            =   7290
      TabIndex        =   17
      Top             =   2220
      Width           =   945
   End
   Begin VB.Label T1_MIDASI 
      AutoSize        =   -1  'True
      Caption         =   "çáÅ@Å@Å@åv"
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
      Left            =   4980
      TabIndex        =   16
      Top             =   2220
      Width           =   1050
   End
   Begin VB.Line B1_LINE 
      X1              =   300
      X2              =   11190
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   5
      Index           =   1
      X1              =   300
      X2              =   11190
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Line H1_K_LINE 
      BorderWidth     =   5
      Index           =   0
      X1              =   300
      X2              =   11175
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Label H1_MM 
      AutoSize        =   -1  'True
      Caption         =   "MM"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   720
      Width           =   300
   End
   Begin VB.Label H1_YY 
      AutoSize        =   -1  'True
      Caption         =   "YYYY"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   14
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "îN"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   1035
      TabIndex        =   13
      Top             =   720
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "åéìx"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   1770
      TabIndex        =   12
      Top             =   720
      Width           =   570
   End
   Begin VB.Label B1_TORNM 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
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
      Left            =   1815
      TabIndex        =   11
      Top             =   2220
      Width           =   4200
   End
   Begin VB.Label B1_SU 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999"
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
      Left            =   7290
      TabIndex        =   10
      Top             =   2220
      Width           =   945
   End
   Begin VB.Label B1_UriGak 
      AutoSize        =   -1  'True
      Caption         =   "9,999,999,999"
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
      Left            =   9135
      TabIndex        =   9
      Top             =   2205
      Width           =   1365
   End
   Begin VB.Label H1_PAGE 
      AutoSize        =   -1  'True
      Caption         =   "9,999"
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
      TabIndex        =   8
      Top             =   600
      Width           =   600
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
      Left            =   8220
      TabIndex        =   7
      Top             =   600
      Width           =   1920
   End
   Begin VB.Label B1_No 
      AutoSize        =   -1  'True
      Caption         =   "999"
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
      Left            =   405
      TabIndex        =   6
      Top             =   2220
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "áÇ"
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
      Left            =   465
      TabIndex        =   5
      Top             =   1770
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "å_Å@ñÒÅ@êÊ"
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
      Left            =   2460
      TabIndex        =   4
      Top             =   1770
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "î[ì¸êî"
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
      Left            =   7515
      TabIndex        =   3
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ã‡Å@äz"
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
      Index           =   4
      Left            =   9765
      TabIndex        =   2
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label1 
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
      Left            =   10950
      TabIndex        =   1
      Top             =   600
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Å¶ îÑè„é¿ê—ï\ Å¶"
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
      Left            =   4185
      TabIndex        =   0
      Top             =   600
      Width           =   2880
   End
End
Attribute VB_Name = "CXU0125L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  îÑè„é¿ê—ï\(ÉJÉKÉ~)
'**       ÉtÉHÅ[ÉÄID    :  CXU0125L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2008/06/06 By CIS
'**
'**********************************************************'
    Dim wCnt            As Long
    Dim wNo             As Long
    Dim OutStr          As String   ' CSVèoóÕì‡óe
    Dim mPrtCnt         As Long
    
    Dim wKouku          As String
    Dim wHinbn          As String
    Dim wHinCnt         As Long
    Dim wKSuryo         As Currency
    Dim wKUriGak        As Currency
    Dim wGSuryo         As Currency
    Dim wGUriGak        As Currency
    Dim wSSuryo         As Currency
    Dim wSUriGak        As Currency
'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub Form_Load()
    App.Title = "îÑè„é¿ê—ï\"
    If Not DataRead Then Exit Sub
    
    wKouku = ""

    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    mPrtCnt = CisDB.RecordCount

    With CisPrt
        .iPrtTitle = "îÑè„é¿ê—ï\"
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
    
    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT UJ.å_ñÒêÊ, ISNULL(TR.ó™èÃ, '') å_ñÒêÊñº," & vbCr
    gSL_Select = gSL_Select & "        HT.é©é–çHãÊ, ISNULL(SY1.ílñºèÃ, '') é©é–çHãÊñº," & vbCr
    gSL_Select = gSL_Select & "        SUM(UJ.î[ì¸êî) î[ì¸êî, SUM(UJ.ã‡äz) ã‡äz, UJ.é¿ê—îNåé" & vbCr
    gSL_Select = gSL_Select & "   FROM îÑè„é¿ê—ÉeÅ[ÉuÉã UJ" & vbCr
    gSL_Select = gSL_Select & "        Left Outer Join éÊà¯êÊÉ}ÉXÉ^ TR" & vbCr
    gSL_Select = gSL_Select & "             ON TR.éÊà¯êÊãÊï™ = 0" & vbCr
    gSL_Select = gSL_Select & "            And TR.éÊà¯êÊCD = UJ.å_ñÒêÊ" & vbCr
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN ïiî‘éÊà¯êÊÉ}ÉXÉ^ HT" & vbCr
    gSL_Select = gSL_Select & "             ON HT.ïiî‘ = UJ.ïiî‘" & vbCr
    gSL_Select = gSL_Select & "            AND HT.éÊà¯êÊ = UJ.î[ì¸êÊ" & vbCr
    gSL_Select = gSL_Select & "            AND HT.éÛì¸ = UJ.éÛì¸" & vbCr
    gSL_Select = gSL_Select & "            AND HT.ïiñ⁄ = 0" & vbCr
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN ñºèÃÉ}ÉXÉ^ SY1" & vbCr
    gSL_Select = gSL_Select & "             ON SY1.ãÊï™ñºèÃ = 'é©é–çHãÊ'" & vbCr
    gSL_Select = gSL_Select & "            AND SY1.íl = HT.é©é–çHãÊ" & vbCr
    gSL_Select = gSL_Select & "            AND SY1.ãÊï™É^ÉCÉv = ''" & vbCr
    gSL_Select = gSL_Select & "  WHERE UJ.é¿ê—îNåé = '" & RV_YM & "'" & vbCr
    If Trim(RV_Torcd) <> "" Then
       gSL_Select = gSL_Select & "    AND UJ.å_ñÒêÊ = '" & RV_Torcd & "'" & vbCr
    End If
    If Trim(RV_KOKBN) <> "" Then
       gSL_Select = gSL_Select & "    AND HT.é©é–çHãÊ = '" & RV_KOKBN & "'" & vbCr
    End If
    gSL_Select = gSL_Select & "  GROUP BY UJ.å_ñÒêÊ, TR.ó™èÃ, HT.é©é–çHãÊ, SY1.ílñºèÃ, UJ.é¿ê—îNåé"
    gSL_Select = gSL_Select & "  ORDER BY HT.é©é–çHãÊ, UJ.å_ñÒêÊ"
    If Not UJTRead(gSL_Select) Then Exit Function
    
    DataRead = True
End Function
'*---------------------*
'*Å@  àÛ  ç¸  èà  óù
'*---------------------*
Private Function PrintProc()
    
    Call CisPrt.ClrLstFld("B*")
    CisPrt.cLine = 99
    wNo = 0: wKouku = "": wHinbn = "": wHinCnt = 0
    Do Until Not UJT_RDSTS
        With UJT
            If wCnt <> 0 Then
                If wKouku <> .é©é–çHãÊ Then
                    Call TotalWrite(1)
                End If
            End If
            
            Call HeadWrite
            wNo = wNo + 1
            
            B1_No = CisFun.RSetFld(wNo, 5, "#,###")
            B1_Torcd = .å_ñÒêÊ
            B1_TORNM = .å_ñÒêÊñº
            B1_SU = Format(Format(.î[ì¸êî, "#,##0;-#,##0;#"), String(9, "@"))
            B1_UriGak = Format(Format(.ã‡äz, "#,##0;-#,##0;#"), String(13, "@"))
            
            Call CisPrt.PrintFld("B1")

            Call CisPrt.ClrLstFld("B*")
            
            With CisPrt
                .cLine = .cLine + 1
                .cDataCnt = .cDataCnt + 1
                .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(mPrtCnt, "#,###")
            End With
            
            wHinCnt = wHinCnt + 1
            wKSuryo = wKSuryo + .î[ì¸êî
            wKUriGak = wKUriGak + .ã‡äz
            wSSuryo = wSSuryo + .î[ì¸êî
            wSUriGak = wSUriGak + .ã‡äz
            wKouku = .é©é–çHãÊ
            wHinbn = .ïiî‘
        End With

        Call UJTReadNext
        wCnt = wCnt + 1
    Loop
    Call UJTClose
    
    Call HeadWrite
    Call TotalWrite(2)
    
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
    
        With UJT
            H1_PAGE = CisFun.RSetFld(CisPrt.cPage, 5, "#,###")
            H1_YMD = Format(Now(), "YYYY/MM/DD HH:NN")
            H1_YY = Mid(.é¿ê—îNåé, 1, 4)
            H1_MM = Mid(.é¿ê—îNåé, 5, 2)
            H1_Jisya = .é©é–çHãÊñº
        End With
    
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With
    
    HeadWrite = True
End Function
'*----------------*
'*   çáåvàÛç¸     *
'*----------------*
Private Sub TotalWrite(Optional Kb As Integer = 0)
    
    If Kb = 0 Then Exit Sub
    
    B1_LINE.BorderWidth = 2
    B1_LINE.Visible = True
    Call CisPrt.PrintFld("B1")
    Call CisPrt.ClrLstFld("B*")
    With CisPrt
        .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(mPrtCnt, "#,###")
    End With
    B1_LINE.BorderWidth = 1
    
    T1_Line.Visible = True
    T1_MIDASI = "Åsçá      åvÅt"
    T1_SU = Format(Format(wKSuryo, "#,##0;-#,##0;#"), String(9, "@"))
    T1_UriGak = Format(Format(wKUriGak, "#,##0;-#,##0;#"), String(13, "@"))
    Call CisPrt.PrintFld("T1")
    Call CisPrt.ClrLstFld("T*")
    CisPrt.cLine = CisPrt.cLine + 1
    
    wKSuryo = 0: wKUriGak = 0: wNo = 0
    If Kb = 1 Then
       CisPrt.cLine = 99
    End If
    
    If Kb = 1 Then Exit Sub
    
    T1_MIDASI = "Åyëç  çá  åvÅz"
    T1_SU = Format(Format(wSSuryo, "#,##0;-#,##0;#"), String(9, "@"))
    T1_UriGak = Format(Format(wSUriGak, "#,##0;-#,##0;#"), String(13, "@"))
    Call CisPrt.PrintFld("T1")
    
End Sub
