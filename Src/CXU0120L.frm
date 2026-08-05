VERSION 5.00
Begin VB.Form CXU0120L 
   Caption         =   "îÑè„é¿ê—àÍóóï\"
   ClientHeight    =   3000
   ClientLeft      =   6750
   ClientTop       =   6045
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   11640
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
      Left            =   5490
      TabIndex        =   31
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
      Left            =   4785
      TabIndex        =   30
      Top             =   1185
      Width           =   720
   End
   Begin VB.Line T1_Line 
      X1              =   3630
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
      Left            =   8850
      TabIndex        =   29
      Top             =   2235
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
      Left            =   5850
      TabIndex        =   28
      Top             =   2235
      Width           =   945
   End
   Begin VB.Label T1_MIDASI 
      AutoSize        =   -1  'True
      Caption         =   "å_ñÒêÊåv"
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
      TabIndex        =   27
      Top             =   2235
      Width           =   840
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   720
      Width           =   570
   End
   Begin VB.Label B1_NouNo 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX"
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
      Left            =   4815
      TabIndex        =   22
      Top             =   2220
      Width           =   945
   End
   Begin VB.Label B1_HINBN 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
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
      Left            =   870
      TabIndex        =   21
      Top             =   2220
      Width           =   2625
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
      Left            =   5850
      TabIndex        =   20
      Top             =   2220
      Width           =   945
   End
   Begin VB.Label B1_TANKA 
      AutoSize        =   -1  'True
      Caption         =   "999,999,999.00"
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
      Left            =   7065
      TabIndex        =   19
      Top             =   2220
      Width           =   1470
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
      Left            =   8850
      TabIndex        =   18
      Top             =   2220
      Width           =   1365
   End
   Begin VB.Label B1_NYMD 
      AutoSize        =   -1  'True
      Caption         =   "YYY/MM/DD"
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
      Left            =   3645
      TabIndex        =   17
      Top             =   2220
      Width           =   945
   End
   Begin VB.Label B1_Kari 
      AutoSize        =   -1  'True
      Caption         =   "x"
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
      Left            =   10725
      TabIndex        =   16
      Top             =   2220
      Width           =   105
   End
   Begin VB.Label H1_TORCD 
      AutoSize        =   -1  'True
      Caption         =   "XXXXXXX"
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
      Left            =   1260
      TabIndex        =   15
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label H1_TORNM 
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
      Left            =   2160
      TabIndex        =   14
      Top             =   1200
      Width           =   2400
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
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   600
      Width           =   1920
   End
   Begin VB.Label B1_No 
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
      Left            =   180
      TabIndex        =   11
      Top             =   2220
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "å_ñÒêÊÅF"
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
      Left            =   300
      TabIndex        =   10
      Top             =   1200
      Width           =   960
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
      Left            =   420
      TabIndex        =   9
      Top             =   1770
      Width           =   225
   End
   Begin VB.Label Label1 
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
      Index           =   9
      Left            =   4815
      TabIndex        =   8
      Top             =   1770
      Width           =   960
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
      Left            =   1770
      TabIndex        =   7
      Top             =   1770
      Width           =   900
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
      Left            =   6075
      TabIndex        =   6
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "íPÅ@âø"
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
      Left            =   7905
      TabIndex        =   5
      Top             =   1770
      Width           =   675
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
      Left            =   9480
      TabIndex        =   4
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "åüé˚ì˙"
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
      Left            =   3780
      TabIndex        =   3
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "âºíPâø"
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
      Left            =   10440
      TabIndex        =   2
      Top             =   1770
      Width           =   675
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
      Caption         =   "Å¶ îÑè„é¿ê—àÍóóï\ Å¶"
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
      Left            =   4095
      TabIndex        =   0
      Top             =   600
      Width           =   3600
   End
End
Attribute VB_Name = "CXU0120L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  îÑè„é¿ê—àÍóóï\
'**       ÉtÉHÅ[ÉÄID    :  CXU0120L
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2004/03/23
'**       ïœ  çX  ì˙    :  2005/06/09 By CIS å^îÔÇÃåüé˚ì˙Ç™îNåéì˙Ç…ÇƒàÛç¸Ç≥ÇÍÇƒÇ¢ÇÈïsãÔçáÇèCê≥
'**       ïœ  çX  ì˙    :  2008/06/06 By CIS çHãÊï Ç…àÛç¸Ç≈Ç´ÇÈÇÊÇ§ïœçX
'**
'**********************************************************'
    Dim wCnt            As Long
    Dim wNo             As Long
    Dim OutStr          As String   ' CSVèoóÕì‡óe
    Dim mPrtCnt         As Long
    
    Dim wTorcd          As String
    Dim wKouku          As String
    Dim wHinbn          As String
    Dim wHinCnt         As Long
    Dim wHSuryo         As Currency
    Dim wHUriGak        As Currency
    Dim wKSuryo         As Currency '-------- 2008/06/06 ADD
    Dim wKUriGak        As Currency '-------- 2008/06/06 ADD
    Dim wGSuryo         As Currency
    Dim wGUriGak        As Currency
    Dim wSSuryo         As Currency
    Dim wSUriGak        As Currency
'*------------------------------------*
'*      ÇeÇèÇíÇçÅQÇkÇèÇÅÇÑ ÉCÉxÉìÉg     *
'*------------------------------------*
Private Sub Form_Load()
    App.Title = "îÑè„é¿ê—àÍóóï\"
'    If Not DataRead Then Exit Sub
    
    wTorcd = ""

    '*---------------------*
    '*  ÉvÉäÉìÉgånèâä˙ê›íË
    '*---------------------*
    mPrtCnt = CisDB.RecordCount

    With CisPrt
        .iPrtTitle = "îÑè„é¿ê—àÍóóï\"
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
    wNo = 0: wTorcd = "": wHinbn = "": wHinCnt = 0
    Do Until Not UJT_RDSTS
        With UJT
            If wCnt <> 0 Then
                If wTorcd <> .å_ñÒêÊ Then
                    Call TotalWrite(2)
                Else
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 INSERT START
                    If wKouku <> .é©é–çHãÊ Then
                        Call TotalWrite(1)
                    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 INSERT END
                End If
            End If
            
            Call HeadWrite
            wNo = wNo + 1
            
            B1_No = CisFun.RSetFld(wNo, 5, "#,###")
            If RTrim(wHinbn) <> RTrim(.ïiî‘) Then
                If RTrim(wHinbn) <> "" Then
                'ïiî‘çáåv
                    Call TotalWrite(0)
                End If
                If Trim(.ïiî‘) <> "" Then
                    B1_HINBN = .ïiî‘
                End If
                B1_LINE.Visible = True
            Else
'
'< 2004.12.07 ADD START >-----------------------------------------------------------------------
'
                If CisPrt.cLine = 0 Then B1_HINBN = .ïiî‘
'
'< 2004.12.07 ADD  END  >-----------------------------------------------------------------------
'
                B1_LINE.Visible = False
            End If
            If Trim(.åüé˚îNåéì˙) <> "" Then
'< 2005/06/09 Delete START >-----------------------------------------------------------------------
'                With CisFun
'                    .DateE_BefVal = UJT.åüé˚îNåéì˙
'                    .DateEdit
'                     B1_NYMD = .DateE_AfVal
'                End With
'< 2005/06/09 Delete End >-----------------------------------------------------------------------
'< 2005/06/09 ADD START >-----------------------------------------------------------------------
                With CisFun
                    .DateE_BefVal = UJT.åüé˚îNåéì˙
                    If Len(UJT.åüé˚îNåéì˙) > 6 Then
                       .DateEdit
                       B1_NYMD = .DateE_AfVal
                    Else
                       B1_NYMD = Mid(UJT.åüé˚îNåéì˙, 1, 4) & "/" & Mid(UJT.åüé˚îNåéì˙, 5, 2)
                    End If
                End With
'< 2005/06/09 ADD End >-----------------------------------------------------------------------
            End If
            B1_NouNo = RTrim(.î[ì¸î‘çÜ)
            B1_SU = Format(Format(.î[ì¸êî, "#,##0;-#,##0;#"), String(9, "@"))
            B1_TANKA = Format(Format(.íPâø, "#,##0.00;-#,##0.00;#"), String(14, "@"))
            B1_UriGak = Format(Format(.ã‡äz, "#,##0;-#,##0;#"), String(13, "@"))
            If Trim(.âºíPâøãÊï™) <> "" Then
                B1_Kari = "*"
            End If
            
            Call CisPrt.PrintFld("B1")

            Call CisPrt.ClrLstFld("B*")
            
            With CisPrt
                .cLine = .cLine + 1
                .cDataCnt = .cDataCnt + 1
                .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(mPrtCnt, "#,###")
            End With
            
            wHinCnt = wHinCnt + 1
            wHSuryo = wHSuryo + .î[ì¸êî
            wHUriGak = wHUriGak + .ã‡äz
            wKSuryo = wKSuryo + .î[ì¸êî '-------- 2008/06/06 ADD
            wKUriGak = wKUriGak + .ã‡äz '-------- 2008/06/06 ADD
            wGSuryo = wGSuryo + .î[ì¸êî
            wGUriGak = wGUriGak + .ã‡äz
            wSSuryo = wSSuryo + .î[ì¸êî
            wSUriGak = wSUriGak + .ã‡äz
            wKouku = .é©é–çHãÊ          '-------- 2008/06/06 ADD
            wTorcd = .å_ñÒêÊ
            wHinbn = .ïiî‘
        End With

        Call UJTReadNext
        wCnt = wCnt + 1
    Loop
    Call UJTClose
    
    Call HeadWrite
    Call TotalWrite(3)
    
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
            H1_TORCD = Trim(.å_ñÒêÊ)
            H1_TORNM = Trim(.å_ñÒêÊñº)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 INSERT START
            H1_Jisya = Trim(.é©é–çHãÊñº)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 INSERT END
        End With
    
        .PrintNewPage       'â¸ï≈èàóù
        .PrintFld ("H1")    'HEADïîàÛç¸
    End With

'---------------------------------------------< 2004.12.07 MNT START >
'    wHinbn = ""
'---------------------------------------------< 2004.12.07 MNT  END  >
    
    HeadWrite = True
End Function
'*----------------*
'*   çáåvàÛç¸     *
'*----------------*
Private Sub TotalWrite(Optional Kb As Integer = 0)
    With CisPrt
        
        T1_Line.Visible = True
        If wHinCnt > 1 Then
            T1_MIDASI = "((ïi î‘ åv))"
            T1_SU = Format(Format(wHSuryo, "#,##0;-#,##0;#"), String(9, "@"))
            T1_UriGak = Format(Format(wHUriGak, "#,##0;-#,##0;#"), String(13, "@"))
            Call .PrintFld("T1")
            Call .ClrLstFld("T*")
            CisPrt.cLine = CisPrt.cLine + 1
        End If
        
        wHSuryo = 0: wHUriGak = 0: wHinCnt = 0
    End With

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 INSERT START
'-Å´-Å´-Å´-Å´-Å´-Å´-- çHãÊåvï\é¶ --Å´-Å´-Å´-Å´-Å´-Å´-
    If Kb = 0 Then Exit Sub
    
    B1_LINE.BorderWidth = 2
    B1_LINE.Visible = True
    Call CisPrt.PrintFld("B1")
    Call CisPrt.ClrLstFld("B*")
    With CisPrt
        .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(mPrtCnt, "#,###")
    End With
    B1_LINE.BorderWidth = 1
    
    T1_Line.Visible = False
    T1_MIDASI = "ÅsçH ãÊ åvÅt"
    T1_SU = Format(Format(wKSuryo, "#,##0;-#,##0;#"), String(9, "@"))
    T1_UriGak = Format(Format(wKUriGak, "#,##0;-#,##0;#"), String(13, "@"))
    Call CisPrt.PrintFld("T1")
    Call CisPrt.ClrLstFld("T*")
    CisPrt.cLine = CisPrt.cLine + 1
    
    wKSuryo = 0: wKUriGak = 0: wNo = 0
    If Kb = 1 Then
       CisPrt.cLine = 99
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/06 INSERT END

'-Å´-Å´-Å´-Å´-Å´-Å´-- å_ñÒêÊåvï\é¶ --Å´-Å´-Å´-Å´-Å´-Å´-
    If Kb = 1 Then Exit Sub 'If Kb = 0 Then Exit Sub-------- 2008/06/06 MNT
    
    B1_LINE.BorderWidth = 2
    B1_LINE.Visible = True
    Call CisPrt.PrintFld("B1")
    Call CisPrt.ClrLstFld("B*")
    With CisPrt
        .dpCountDisp Format(.cDataCnt, "#,###") & " / " & Format(mPrtCnt, "#,###")
    End With
    B1_LINE.BorderWidth = 1
    
    T1_Line.Visible = False
    T1_MIDASI = "Åså_ñÒêÊåvÅt"
    T1_SU = Format(Format(wGSuryo, "#,##0;-#,##0;#"), String(9, "@"))
    T1_UriGak = Format(Format(wGUriGak, "#,##0;-#,##0;#"), String(13, "@"))
    Call CisPrt.PrintFld("T1")
    Call CisPrt.ClrLstFld("T*")
    CisPrt.cLine = CisPrt.cLine + 1
    
    wGSuryo = 0: wGUriGak = 0: wNo = 0
    If Kb = 2 Then
       CisPrt.cLine = 99
    End If
    
'-Å´-Å´-Å´-Å´-Å´-Å´-- ëççáåvï\é¶ --Å´-Å´-Å´-Å´-Å´-Å´-
    If Kb = 2 Then Exit Sub 'If Kb = 1 Then Exit Sub--------- 2008/06/06 MNT
    
    T1_MIDASI = "Åyëç çá åvÅz"
    T1_SU = Format(Format(wSSuryo, "#,##0;-#,##0;#"), String(9, "@"))
    T1_UriGak = Format(Format(wSUriGak, "#,##0;-#,##0;#"), String(13, "@"))
    Call CisPrt.PrintFld("T1")
    
End Sub
