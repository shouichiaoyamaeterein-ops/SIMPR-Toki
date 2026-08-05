VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Begin VB.Form CXM0310 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "JISãKäiïœä∑ÉÅÉìÉeÉiÉìÉX"
   ClientHeight    =   2475
   ClientLeft      =   6195
   ClientTop       =   3870
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "ÇlÇr ÉSÉVÉbÉN"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z µ∞¿ﬁ∞
   ScaleHeight     =   2475
   ScaleWidth      =   7530
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   30
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "Åy JISãKäiïœä∑ÉÅÉìÉeÉiÉìÉX  Åz"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      cAlingnment     =   1
      cPositionX      =   1200
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   16711935
         Caption         =   "XXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1005
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   7320
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1410
      Width           =   7380
      Begin Cis3D_v60.CIS3D UC_3D 
         Height          =   885
         Index           =   0
         Left            =   45
         Top             =   30
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   1561
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   795
            Index           =   5
            Left            =   45
            Top             =   45
            Width           =   7170
            _ExtentX        =   12647
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            Begin Cis3D_v60.CIS3D LB1_Zai 
               Height          =   345
               Left            =   75
               Top             =   345
               Width           =   3060
               _ExtentX        =   5398
               _ExtentY        =   609
               BackColor       =   16777152
               Caption         =   "SPC440"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   8421504
               cBoderColor2    =   -2147483643
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   90
            End
            Begin Cis3D_v60.CIS3D CIS3D2 
               Height          =   315
               Index           =   0
               Left            =   1095
               Top             =   30
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   556
               ForeColor       =   16711680
               Caption         =   "çﬁÅ@Å@éø"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cBoderStyle     =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D1 
               Height          =   360
               Left            =   3270
               Top             =   345
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   635
               Caption         =   "Å®"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   20.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   2
            End
            Begin CisText_V60.CisText iH1_Zaist 
               Height          =   360
               Left            =   75
               TabIndex        =   3
               Top             =   345
               Visible         =   0   'False
               Width           =   3120
               _ExtentX        =   5503
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "SPC440"
               MaxLength       =   25
            End
            Begin CisText_V60.CisText iH1_JIS 
               Height          =   360
               Left            =   3930
               TabIndex        =   4
               Top             =   345
               Width           =   3210
               _ExtentX        =   5662
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   25
               IMEMode         =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D2 
               Height          =   315
               Index           =   1
               Left            =   4905
               Top             =   30
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   556
               ForeColor       =   16711680
               Caption         =   "ÇiÇhÇrãKäi"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cBoderStyle     =   2
            End
         End
      End
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      IMEMode         =   3  'µÃå≈íË
      Left            =   180
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   150
      Width           =   855
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4770
      ScaleHeight     =   615
      ScaleWidth      =   2610
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   465
      Width           =   2670
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   885
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   2
         BtnCaption      =   "ì¸óÕ"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   128
         BtnShadowColor  =   16777215
         PFCaption       =   "(Enter)"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   16777215
         PFShadowColor   =   0
         PFBackColor     =   128
      End
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   1740
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "ñﬂÇÈ"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16711680
         BtnShadowColor  =   16777215
         PFCaption       =   "( F12 )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   12648447
         PFShadowColor   =   0
         PFBackColor     =   16711680
      End
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   30
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   3
         BtnCaption      =   "éÊè¡"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   32768
         BtnShadowColor  =   16777215
         PFCaption       =   "( Esc )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   65535
         PFShadowColor   =   0
         PFBackColor     =   49152
      End
   End
End
Attribute VB_Name = "CXM0310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   JISãKäiïœä∑ÉÅÉìÉeÉiÉìÉXâÊñ 
'**       ÉtÉHÅ[ÉÄID    :   CXM0310
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2009/09/23  By CIS
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB          As String * 2   ' ì¸óÕàÊêÿë÷
    Dim FormCap         As String
    Dim CisFrm          As CISFormContorl
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås    ÇkÇèÇÅÇÑÉCÉxÉìÉg    Åt              +
'+----------------------------------------------------------+
Private Sub form_load()
'   #-------------------#
'   #  çÄ ñ⁄ èâ ä˙ âª    #
'   #-------------------#
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
'
    Dummy.Left = -1000
    Dummy.Enabled = False
'   #------------------#
'   #    èâä˙ì¸óÕê›íË   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    
    iH1_Zaist.Visible = False
    LB1_Zai.Caption = RV_JIS_Zai
    iH1_JIS.Text = RV_JIS_Num
'   #------------------#
'   # âÊñ íÜâõà íuï\é¶  #
'   #------------------#
    If RV_Left <> 0 Or RV_Top <> 0 Then
        Me.Move RV_Left, RV_Top
    Else
        Call CisFrm.Frm_Center(False, x1024y768)
    End If
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    ÅsÇ`ÇÉÇîÇâÇñÇÅÇîÇÖ ÉCÉxÉìÉgÅt                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormCap = Me.Caption
    FormAct = Not FormAct
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås ÇjÇÖÇôÇcÇèÇóÇéÉCÉxÉìÉg Åt                 +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄÅs ÉtÉ@ÉìÉNÉVÉáÉìÉLÅ[ópÉAÉNÉZÉâÉåÅ[É^ Åt         +
'+-------------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    Select Case Key_Code
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    'ÅyéÊè¡Åz
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    'ÅyèIóπÅz
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      'Åyì¸óÕÅz
                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                  Dummy.Enabled = False
                               End If
           Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås QueryUnloadÉCÉxÉìÉg Åt                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Call PB_END_Click
End Sub
'+---------------------+
'+    èI  óπ  ÉL Å[     +
'+---------------------+
Private Sub PB_END_Click()
    Unload CXM0310
End Sub
'+----------------------------+
'+    ì¸  óÕ  ÉL Å[(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Call ReturnPress
    Call PB_END_Click
End Sub
'+----------------------------+
'+    éÊ  è¡  ÉL Å[(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    Call CisFrm.MousePT(11)
    iH1_JIS.Text = RV_JIS_Num       ' åƒÇŒÇÍÇΩéûÇÃèÛë‘Ç…ñﬂÇ∑
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*                          *
'*      HEAD/BODY èàóù      *
'*                          *
'****************************
'****************************
'*      ì¸óÕêÿë÷ èàóù       *
'****************************
Private Function DispChange(Pro As String)
' ¿ﬁ–∞çÄñ⁄(âÊñ ÇÃÉ`ÉâÉcÉLñhé~)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If

    H1_Area1.Enabled = True
    H1_Area1.BackColor = gPLostSel
'    PB_END.Visible = H1Mode
' ¿ﬁ–∞çÄñ⁄(ì¸óÕã÷é~)
    Dummy.Enabled = False
End Function
'****************************
'*      ì¸óÕèàóù(Enter)      *
'****************************
Private Sub ReturnPress()
    Dim flg As Boolean      ' ë∂ç›óLñ≥ÉtÉâÉO
    
    Call CisFrm.MousePT(11)
    flg = False
   
    gSL_Select = "SELECT * FROM çﬁóøJISãKäiïœä∑É}ÉXÉ^ WHERE çﬁéø = '" & Trim(LB1_Zai.Caption) & "'"
    If ZJSRead(gSL_Select, 1) Then
        flg = True
    End If
    
    iH1_JIS.Text = Trim(iH1_JIS.Text)
    If iH1_JIS.Text <> "" Then
        If flg = True Then
            If iH1_JIS.Text <> RV_JIS_Num Then
                With CisDB
                    On Error GoTo ReturnPress_Err
                    .DBTran (TransBegin)    ' ƒ◊›ªﬁ∏ºÆ›äJén
                    .SQL = "UPDATE çﬁóøJISãKäiïœä∑É}ÉXÉ^ SET JISãKäi = '" & iH1_JIS.Text & "',"
                    .SQL = .SQL & "çÏê¨ì˙ = GETDATE(),"
                    .SQL = .SQL & "çÏê¨é“ = '" & gTanto & "',"
                    .SQL = .SQL & "çÏê¨í[ññ = SubString(Host_Name(),1,20)"
                    .SQL = .SQL & " WHERE çﬁéø = '" & Trim(LB1_Zai.Caption) & "'"
                    .DBExec
                    .DBTran (TransCommit)   ' ƒ◊›ªﬁ∏ºÆ› èIóπ
                End With
            End If
'            With ZJS
'                .JISãKäi = iH1_JIS.Text
'            End With
'            Call ZJSUpdate
        Else
            Call ItemsClearZJS
            With ZJS
                .çﬁéø = Trim(LB1_Zai.Caption)
                .JISãKäi = iH1_JIS.Text
            End With
            Call ZJSInsert
        End If
    Else
        If flg = True Then
            With CisDB
                On Error GoTo ReturnPress_Err
                .DBTran (TransBegin)    ' ƒ◊›ªﬁ∏ºÆ›äJén
                .SQL = "DELETE FROM çﬁóøJISãKäiïœä∑É}ÉXÉ^ WHERE çﬁéø = '" & Trim(LB1_Zai.Caption) & "'"
                .DBExec
                .DBTran (TransCommit)   ' ƒ◊›ªﬁ∏ºÆ› èIóπ
             End With
        End If
    End If
    GoTo ReturnPress_Ed

ReturnPress_Err:
    CisFun.ErrorBox
ReturnPress_Ed:
    On Error GoTo 0
    Call CisFrm.MousePT(1)
    RV_JIS_Num = iH1_JIS
End Sub
