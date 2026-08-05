VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmMenuSel 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "メニュー選択"
   ClientHeight    =   4530
   ClientLeft      =   705
   ClientTop       =   7500
   ClientWidth     =   7065
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7065
   Begin VB.PictureBox Picture1 
      Height          =   1740
      Left            =   5430
      ScaleHeight     =   1680
      ScaleWidth      =   1320
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   150
      Width           =   1380
      Begin VB.CommandButton PB_CAN 
         Caption         =   "戻る"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   75
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   945
         Width           =   1185
      End
      Begin VB.CommandButton PB_ENT 
         Caption         =   "選択"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   75
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.PictureBox H1_Area3 
      Height          =   705
      Left            =   75
      ScaleHeight     =   645
      ScaleWidth      =   1080
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1140
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   570
         Left            =   45
         Top             =   30
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1005
         Caption         =   "修正"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
   End
   Begin VB.PictureBox H1_Area2 
      BackColor       =   &H00808080&
      Height          =   720
      Left            =   4455
      ScaleHeight     =   660
      ScaleWidth      =   2385
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2445
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   555
         Left            =   1575
         Top             =   45
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnType         =   4
         BtnCaption      =   "削除"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   255
         BtnShadowColor  =   16777215
         PFCaption       =   "( F3 )"
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
         PFBackColor     =   255
      End
      Begin CisBtn_60.CisBtn PB_MNT 
         Height          =   555
         Left            =   825
         Top             =   45
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnType         =   8
         BtnCaption      =   "修正"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16776960
         BtnShadowColor  =   0
         PFCaption       =   "( F2 )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   0
         PFShadowColor   =   16777215
         PFBackColor     =   16776960
      End
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   555
         Left            =   75
         Top             =   45
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnType         =   7
         BtnCaption      =   "新規"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   65535
         BtnShadowColor  =   0
         PFCaption       =   "( F1 )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   0
         PFShadowColor   =   16777215
         PFBackColor     =   65535
      End
   End
   Begin VB.PictureBox H1_Area1 
      Appearance      =   0  'ﾌﾗｯﾄ
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   75
      ScaleHeight     =   1920
      ScaleWidth      =   6795
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   60
      Width           =   6825
      Begin VB.ListBox MN_List 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   90
         TabIndex        =   8
         Top             =   75
         Width           =   5175
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3585
      Top             =   5790
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Dummy 
      BackColor       =   &H00FFC0FF&
      Height          =   270
      Left            =   135
      TabIndex        =   0
      Text            =   "Dummy"
      Top             =   5910
      Width           =   180
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   1305
      Left            =   75
      ScaleHeight     =   1245
      ScaleWidth      =   6765
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3060
      Width           =   6825
      Begin VB.CommandButton PB_Text 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6090
         TabIndex        =   3
         Top             =   765
         Width           =   345
      End
      Begin CisText_V60.CisText iB1_Title 
         Height          =   375
         Left            =   1425
         TabIndex        =   1
         Top             =   105
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         MaxLength       =   40
         IMEMode         =   4
      End
      Begin CisText_V60.CisText iB1_Text 
         Height          =   345
         Left            =   105
         TabIndex        =   2
         Top             =   765
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         MaxLength       =   80
         IMEMode         =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "メニューテキスト"
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   555
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "メニュータイトル"
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   195
         Width           =   1215
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   0
      X1              =   -30
      X2              =   7100
      Y1              =   2160
      Y2              =   2145
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   -15
      X2              =   7100
      Y1              =   2145
      Y2              =   2145
   End
End
Attribute VB_Name = "FrmMenuSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   メニュー選択
'**       フォームID    :   FrmMenuSel
'**       処理概要      :
'**
'**       作  成  日    :   2006/07/06  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKB                 As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM                 As String       ' 処理名称(新規･修正･削除 等)
    
    Dim mGotColor               As Long
    Dim mLostColor              As Long

    Dim w_LIdx                  As Byte         ' 業務メニューListIndex
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
'
    Dummy.Left = -1000
    Dummy.Enabled = False
    
    Me.Width = 7155
    If gK_MNSelKB = 0 Then
'   [ ﾒﾆｭｰ選択 ]
        Me.Height = 2490
    Else
'   [ ﾒﾆｭｰ登録 ]
        Me.Height = 5010
    End If
    
'   #-----------------------#
'   #    メニュー種類獲得   #
'   #-----------------------#
    Call MenuKinds_Get

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)

'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = "ADD"
    Call SyoriName(SyoriKB)

'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(False, x1024y768)
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    gK_MNSelRtn = 0
    
    Call List_Set
    If gK_MNSelKB = 0 Then
        PB_ENT.Caption = "選択"
    Else
        PB_ENT.Caption = "入力"
    End If
    
    MN_List.SetFocus
    FormAct = Not FormAct
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》                 +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+-------------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》         +
'+-------------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False
    
    Select Case Key_Code
           Case vbKeyF1:        If PB_ADD.Visible Then Call PB_ADD_Click                       '【新規】
           Case vbKeyF2:        If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:        If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_Can_Click    '【取消】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      '【入力】
                                   Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                   Dummy.Enabled = False
                                End If
           Case Else:           Exit Function
    End Select
    
    Key_Acc = True
End Function
'+---------------------------+
'+    業務メニューリスト
'+----------------------------+
Private Sub MN_List_Click()
    w_LIdx = MN_List.ListIndex

    Call EditDisp(w_LIdx)

End Sub
'+----------------------------+
'+    メニューテキスト選択
'+----------------------------+
Private Sub PB_Text_Click()
    On Error Resume Next

    With CommonDialog1
        .Flags = cdlOFNExplorer + cdlOFNHideReadOnly
        .FileName = gIniExe & "MENU.Txt"
        .Filter = "MENU TEXT (*.TXT)|*.txt"
        .CancelError = True
        .ShowOpen

        If Trim(.FileName) <> "" Then
            iB1_Text = .FileName
        End If
    End With

    iB1_Text.SetFocus
    On Error GoTo 0
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Call ReturnPress
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_Can_Click()
    If ProcHB = "H1" Then
        Unload FrmMenuSel
        Exit Sub
    End If

    Call HeadBodyClear(ProcHB)
    ProcHB = "H1"
    Call DispChange(ProcHB)
    Call SyoriName(SyoriKB)
    
End Sub
'+----------------------------+
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    If SyoriKB <> "ADD" Then
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
    End If
    MN_List.SetFocus
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
    End If
    MN_List.SetFocus
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        SyoriKB = "DEL"
    Else
        SyoriKB = ""
    End If
    Call SyoriName(SyoriKB)
    MN_List.SetFocus
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim W_Color     As Long
   W_Color = &H80000008 '初期色ｾｯﾄ(黒)
   Select Case SyoriKB
     Case "ADD"
          SyoriNM = "追加"
          K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
          PB_ADD.Visible = False
     Case "MNT"
          SyoriNM = "修正"
          K_Sykbnm.ForeColor = &HFF0000  '(青色)
     Case "DEL"
          SyoriNM = "削除"
          K_Sykbnm.ForeColor = &HFF&     '(赤色)
          W_Color = &HFF&
     Case Else
          SyoriNM = ""
   End Select
   K_Sykbnm = SyoriNM
   
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
'      * PF-Key 表示設定 *
       If SyoriKB <> "LOC" Then
          If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
             gObj.Name = "PB_DEL" Then
             gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
          End If
       End If
   Next gObj
End Sub
'****************************
'*                          *
'*      HEAD/BODY 処理      *
'*                          *
'****************************
'****************************
'*      入力切替 処理        *
'****************************
Private Function DispChange(Pro As String)
    Dim H1Mode      As Boolean
    Dim H1Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    Dim wKB         As Boolean
    
    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
        Case "H1":  H1Mode = True:  H1Color = &HFF&
        Case "B1":  B1Mode = True:  B1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color

    If gK_MNSelKB = 0 Then
'   [ ﾒﾆｭｰ選択 ]
        wKB = False
    Else
'   [ ﾒﾆｭｰ登録 ]
        wKB = True
    End If
    H1_Area2.Enabled = wKB
    H1_Area3.Enabled = wKB
    B1_Area1.Enabled = wKB
    
    B1_Area1.Enabled = B1Mode
    
    PB_ADD.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_DEL.Visible = H1Mode

' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = "": .Enabled = True
            If .Name Like wbc_Enm & "Op_*" Then .Value = False: .BackColor = mLostColor
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)

    If gK_MNSelKB = 0 Then
'   [ ﾒﾆｭｰ選択 ]
        With CisFun
            .INI_FileID = gIniExe & gIniName
            .INI_Section = .WSID
            .INI_Key = "MenuLast"
            If MN_List.ListIndex = 0 Then
                .INI_String = vbNullString
            Else
                .INI_String = MN_List.ListIndex
            End If
            .PutIni
        End With
        gK_MNSelRtn = 1
        Call PB_Can_Click
    Else
'   [ ﾒﾆｭｰ登録 ]
        If ProcHB = "H1" Then
            If Not Head1Chk Then GoTo ReturnPress_Ed
    ' << 削除処理 >>
            If SyoriKB = "DEL" Then
                If Not DBDelete Then GoTo ReturnPress_Ed
                gK_EditKB = 1
                Call HeadBodyClear("B")
                Call MenuKinds_Get
                Call List_Set
                ProcHB = "H1"
                Call DispChange(ProcHB)
                Call SyoriName(SyoriKB)
            Else
                ProcHB = "B1"
                Call DispChange(ProcHB)
            End If
            GoTo ReturnPress_Ed
        End If
        
        If ProcHB = "B1" Then
            If Not Body1Chk Then GoTo ReturnPress_Ed
            If Not DBPut Then GoTo ReturnPress_Ed
            gK_EditKB = 1
            Call MenuKinds_Get
            Call List_Set
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
            GoTo ReturnPress_Ed
        End If
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
    If SyoriKB <> "ADD" Then
        If MN_List.ListIndex < 0 Then MN_List.ListIndex = 0
        w_LIdx = MN_List.ListIndex
        Call EditDisp(w_LIdx)
    End If
    
    If SyoriKB = "DEL" Then
        If MN_List.ListIndex = 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    標準メニューは削除出来ません         "
                .MB_Button = Error
                .MBOX
            End With
            MN_List.SetFocus
            Exit Function
        End If
    End If
    
    Head1Chk = True
End Function
'****************************
'*      BODY1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
   If RTrim$(iB1_Title) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    タイトルを入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iB1_Title.SetFocus
        Exit Function
    Else
        If SyoriKB = "ADD" Then
            For gCnt1 = 0 To gKind_Cnt
                If RTrim$(iB1_Title) = RTrim$(gK_Title(gCnt1)) Then
                     With CisFun
                         .MB_Lines = 4
                         .MB_MSG(2) = "    タイトルが重複します         "
                         .MB_Button = Error
                         .MBOX
                     End With
                     iB1_Title.SetFocus
                     Exit Function
                 End If
            Next gCnt1
        End If
    End If
    
   If RTrim$(iB1_Text) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    テキストを入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iB1_Text.SetFocus
        Exit Function
    Else
        If SyoriKB = "ADD" Then
            For gCnt1 = 0 To gKind_Cnt
                If RTrim$(iB1_Text) = RTrim$(gK_Text(gCnt1)) Then
                     With CisFun
                         .MB_Lines = 4
                         .MB_MSG(2) = "    テキストが重複します         "
                         .MB_Button = Error
                         .MBOX
                     End With
                     iB1_Text.SetFocus
                     Exit Function
                 End If
            Next gCnt1
        End If
    End If
    
    If RTrim$(Dir(iB1_Text)) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "    テキストがありません         "
            .MB_MSG(3) = "    　よろしいですか？         "
            .MB_Button = OK_CAN
            If Not .MBOX Then
                iB1_Text.SetFocus
                Exit Function
            End If
        End With
    End If

    Body1Chk = True
End Function
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+      削 除 処 理      +
'+----------------------+
Private Function DBDelete()
    DBDelete = False

'  ( 確認メッセージ )
    If Not CheckMessage Then Exit Function
    
    With CisFun
        .INI_FileID = gIniExe & gIniName
        .INI_Section = .WSID
        .INI_String = vbNullString

'       < Menu Text >
        .INI_Key = "MenuText-" & Format(MN_List.ListIndex, "00")
        .PutIni

'       < Menu Title >
        .INI_Key = "MenuTextTiTle-" & Format(MN_List.ListIndex, "00")
        .PutIni
    
'       < メニュー数 >
        .INI_Key = "MenuKinds"
        If (MN_List.ListCount - 1) = 1 Then
            .INI_String = vbNullString
        Else
            .INI_String = (MN_List.ListCount - 1)
        End If
        .PutIni
    End With
    
    DBDelete = True
    Exit Function

DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False

'  ( 確認メッセージ )
    If Not CheckMessage Then Exit Function
    
    If SyoriKB = "ADD" Then
        gCnt1 = MN_List.ListCount
    Else
        gCnt1 = MN_List.ListIndex
    End If
    
    With CisFun
        .INI_FileID = gIniExe & gIniName
        .INI_Section = .WSID

'       < Menu Text >
        .INI_Key = "MenuText" & Format(gCnt1, "-00;-00;#")
        .INI_String = RTrim$(iB1_Text)
        .PutIni
    
'       < Menu Title >
        If gCnt1 = 0 And _
            RTrim$(iB1_Title) = "標準メニュー" Then
                .INI_String = vbNullString
        Else
            .INI_String = RTrim$(iB1_Title)
        End If
        .INI_Key = "MenuTextTiTle" & Format(gCnt1, "-00;-00;#")
        .PutIni

        If SyoriKB = "ADD" Then
'           < メニュー数 >
            .INI_Key = "MenuKinds"
            .INI_String = gCnt1 + 1
            .PutIni
        End If
    End With

    DBPut = True
End Function
'*-------------------------------------------------*
'*      デ ー タ 更 新  確 認 メ ッ セ ー ジ
'*-------------------------------------------------*
Private Function CheckMessage() As Boolean
    CheckMessage = False
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     " & SyoriNM & "処理を行います      "
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Title = SyoriNM & "処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    CheckMessage = True
End Function
'*-------------------------*
'*      編集内容表示        *
'*-------------------------*
Private Function EditDisp(Idx As Byte)

    iB1_Title = RTrim$(gK_Title(Idx))
    iB1_Text = RTrim$(gK_Text(Idx))
        
End Function
'*--------------------------------*
'*      メニュー種類情報　獲得
'*--------------------------------*
Private Function MenuKinds_Get()
    Static swI          As Byte
    Static swS1         As String
    Static swS2         As String

    Erase gK_Title: Erase gK_Text
    gKind_Cnt = 0
    ReDim Preserve gK_Title(gKind_Cnt)
    ReDim Preserve gK_Text(gKind_Cnt)

    With CisFun
        .INI_Key = .WSID
        .INI_Default = ""

'       < 標準-ﾃｷｽﾄ >
        .INI_Section = "MenuText"
        .INI_FileID = gIniExe & gIniName
        .GetIni
        swS1 = RTrim$(.INI_String)
        If RTrim$(swS1) = "" Then swS1 = gIniExe & "Menu.Txt"
        
'       < 標準-ﾀｲﾄﾙ >
        .INI_Section = .WSID
        .INI_Key = "MenuTextTiTle"
        .GetIni
        swS2 = RTrim$(.INI_String)
        If RTrim$(swS2) = "" Then swS2 = "標準メニュー"
    End With

    gK_Title(0) = swS2
    gK_Text(0) = gFT_Text

'[ メニュー種類獲得 ]
    For swI = 1 To 20
        With CisFun
            .INI_FileID = gIniExe & gIniName
            .INI_Section = .WSID
            .INI_Default = ""

'       < Menu Text >
            .INI_Key = "MenuText" & Format(swI, "-00;-00;#")
            .GetIni

            If RTrim$(.INI_String) = "" Then Exit For
            swS1 = .INI_String

'       < Menu Title >
            .INI_Key = "MenuTextTiTle" & Format(swI, "-00;-00;#")
            .GetIni

            swS2 = RTrim$(.INI_String)
            If RTrim$(swS2) = "" Then swS2 = "業務メニュー" & Format(swI, "00;00;#")

            gKind_Cnt = gKind_Cnt + 1
            ReDim Preserve gK_Title(gKind_Cnt)
            ReDim Preserve gK_Text(gKind_Cnt)

            gK_Text(gKind_Cnt) = RTrim$(swS1)
            gK_Title(gKind_Cnt) = RTrim$(swS2)
        End With
    Next swI

End Function
'*--------------------------------*
'*      リストボックス　セット
'*--------------------------------*
Private Function List_Set()
    With MN_List
        .Clear
        For gCnt1 = 0 To gKind_Cnt
            .AddItem gK_Title(gCnt1), gCnt1
        Next gCnt1
    End With
End Function
