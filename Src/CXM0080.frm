VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CXM0080 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "担当者管理マスタメンテナンス"
   ClientHeight    =   10260
   ClientLeft      =   -15
   ClientTop       =   450
   ClientWidth     =   15360
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10260
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin VB.ComboBox DummyComb 
      BackColor       =   &H00FFC0C0&
      Height          =   345
      Left            =   12240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   4
      Top             =   1410
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   8220
      Left            =   4320
      ScaleHeight     =   8160
      ScaleWidth      =   6960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1620
      Width           =   7020
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   7995
         Left            =   75
         Top             =   60
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   14102
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         Begin vsFlexLib.vsFlexArray vsGrid1 
            DragIcon        =   "CXM0080.frx":0000
            Height          =   7965
            Left            =   0
            TabIndex        =   0
            Top             =   0
            Width           =   6810
            _Version        =   196608
            _ExtentX        =   12012
            _ExtentY        =   14049
            _StockProps     =   228
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "№|担当者ｺｰﾄﾞ|担当者名             |ﾊﾟｽﾜｰﾄﾞﾚﾍﾞﾙ|PLEBEL|INFO|"
            Rows            =   26
            Cols            =   7
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            RowHeightMin    =   300
            SelectionMode   =   1
            BackColorAlternate=   12648447
         End
      End
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   10755
      ScaleHeight     =   660
      ScaleWidth      =   4365
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   585
      Width           =   4425
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   2625
         Top             =   45
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   2
         BtnCaption      =   "入力"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Left            =   3480
         Top             =   45
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   1
         BtnCaption      =   "終了"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1755
         Top             =   45
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
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
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   45
         Top             =   45
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
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
      Begin CisBtn_60.CisBtn PB_MNT 
         Height          =   585
         Left            =   900
         Top             =   45
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
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
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 担当者管理メンテナンス  】"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   11940
         Top             =   60
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   503
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cBoderStyle     =   1
         cPositionX      =   50
      End
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14520
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   450
         BackColor       =   14737632
         ForeColor       =   16711680
         Caption         =   "SIMPR-A"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   14737632
         cBoderColor2    =   14737632
         cFont3DColor1   =   8454143
         cFont3DColor2   =   16777215
         cFont3D         =   1
         cAlingnment     =   7
         cBoderStyle     =   2
      End
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'ｵﾌ固定
      Left            =   330
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   630
      Top             =   -60
   End
   Begin VB.Timer Timer2 
      Left            =   1050
      Top             =   -60
   End
End
Attribute VB_Name = "CXM0080"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   担当者管理マスタメンテナンス
'**       フォームID    :   CXM0080
'**       処理概要      :
'**
'**       作  成  日    :   2003/11/17  By CIS
'**       変  更  日    :
'**
'**********************************************************'

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim CISFrm          As CISFormContorl
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   # 二 重 起 動 防 止  #
'   #-------------------#
    If Not Dupli_Start Then End
'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
    If Not IniGet Then End
    
    In_Pass = CisDB.PassWord(CisDB.InputPass)
    In_Tanto = CisDB.Tanto
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CISFrm = New CISFormContorl
    
    Set CISFrm.ThisForm = Me
    Call CISFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If

    If gTanto = "" And Mid(Command, 1, 1) <> "N" Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(1) = "     担当者入力なしに設定されています。            "
           .MB_MSG(3) = "           処理を終了します。"
           .MB_Button = OK
           .MB_Title = "担当者メンテ"
           .MBOX
           Call PB_END_Click
       End With
    End If
'
    Dummy.Left = -1000
    Dummy.Enabled = False
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
'   #------------------#
'   #  ｸﾞﾘｯﾄﾞ初期処理   #
'   #------------------#
    With c080_VsGrid
        Set .GridObj = vsGrid1              ' ｸﾞﾘｯﾄﾞｵﾌﾞｼﾞｪｸﾄ名 ｾｯﾄ
        .FixedHeight = 380                  ' 固定行(見出し)高さ設定(省略時は300)
        .PatanMax = 1                       ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ最大数
        .PatanInit = 0                      ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ初期値
        .SelectCol = 5                     ' セル選択処理ﾜｰｸ列
        .NoSpaceCol = 1
        .NumberSet = NoSpace
'        Set .SelectDispObj = B1lb_SKensu    ' セル選択件数表示ｵﾌﾞｼﾞｪｸﾄｾｯﾄ
' ﾌﾟﾛｸﾞﾗﾑ作成時、以下のﾌﾗｸﾞをｾｯﾄし、ｸﾘｯﾌﾟﾎﾞｰﾄﾞよりｺｰﾄﾞの貼り付けを行う。(GridInit)
        .InitGet = False                    ' ｸﾞﾘｯﾄﾞ内容初期設定 ｸﾘｯﾌﾟﾎﾞｰﾄﾞ作成ﾌﾗｸﾞ

        .Init                               ' ｸﾞﾘｯﾄﾞ初期処理 実行
    End With
    Call GridInit                           ' ｸﾞﾘｯﾄﾞ初期表示
'+---------------------+
'+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
'+---------------------+
    Set c080_Conbo = New CisCombo
    With c080_Conbo
        Set .CombObj = DummyComb            'ｺﾝﾎﾞﾎﾞｯｸｽｵﾌﾞｼﾞｪｸﾄ名 ｾｯﾄ
        .Control = "名称マスタ"
        .ComboName = "パスワードレベル"
        .SetSpace = 50
        .Widht = 0
        .Ess = True
        If Not .ComboInit Then
           Call .CombowkSet(3)
           .Code = "0": .CombText = "担当者"
           Call .CombowkSet
           .Code = "1": .CombText = "管理者"
           Call .CombowkSet
           .Code = "2": .CombText = "管理責任者"
           Call .CombowkSet
        End If
'+--------------------------------+
'+   ｺﾝﾎﾞﾎﾞｯｸｽへ内容をセットする
'+--------------------------------+
        Call .ComboSet
    End With
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CISFrm.Frm_Center(True, x1024y768)

    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    Call GridSet
    
    Call PFDsip
    
    FormAct = Not FormAct
'
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
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    Select Case Key_Code
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    '【新規】
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                  Dummy.Enabled = False
                               End If
           Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Cancel = 1
    Call PB_END_Click
End Sub

'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'+---------------------+
'+   ﾃﾞｰﾀﾍﾞｰｽ切断
'+---------------------+
    Call CisDB.DBDISConnect
'+---------------------+
'+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
'+---------------------+
    Call CISFrm.UnLoadDisp
    End
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
   Call ReturnPress
End Sub
'+----------------------------+
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    SyoriKB = "ADD"
    Call DelChg
    c080_ProcKB = "ADD"
    vsGrid1.HighLight = flexHighlightWithFocus
        
    CXM0085.Show vbModal
    Unload CXM0085
    Set CXM0085 = Nothing

    vsGrid1.HighLight = flexHighlightAlways

End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    SyoriKB = "MNT"
    Call DelChg
    c080_ProcKB = "MNT"
        
    CXM0085.Show vbModal
    Unload CXM0085
    Set CXM0085 = Nothing
    
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB = "DEL" Then
       SyoriKB = ""
    Else
       SyoriKB = "DEL"
    End If
    Call DelChg
End Sub

Private Sub DelChg()
    With vsGrid1
        If SyoriKB = "DEL" Then
           .BackColor = &HC0C0FF
        Else
           .BackColor = &HC0FFFF
        End If
    End With
End Sub

'*****************************************
'*                                       *
'*      グリッド内容によりＰＦキー　表示   *
'*                                       *
'*****************************************
Private Sub PFDsip()
    
    With vsGrid1
        PB_ADD.Visible = True
        If Trim(.TextMatrix(.Row, c080_VsGrid.FixedGet(1))) = "" Then
           PB_MNT.Visible = False
           PB_DEL.Visible = False
        Else
           PB_MNT.Visible = True
           PB_DEL.Visible = True
        End If
    End With
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   Call CISFrm.TimeSet("1")
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    
    If SyoriKB = "DEL" Then
        If Not Body1Chk Then Exit Sub
        If Not DBDelete Then Exit Sub
        Exit Sub
    End If
    
    
    With vsGrid1
        If Trim(.TextMatrix(.Row, c080_VsGrid.FixedGet(1))) = "" Then
           Call PB_ADD_Click
        Else
           Call PB_MNT_Click
        End If
    End With
    
ReturnPress_Ed:
End Sub
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    If SyoriKB = "DEL" And Trim(Trim(vsGrid1.TextMatrix(vsGrid1.Row, c080_VsGrid.FixedGet(1)))) = gTanto Then
        With CisFun
            .MB_Lines = 3
            .MB_MSG(2) = "       入力担当者は削除できません。            "
            .MB_Button = Error
            Call .MBOX
        End With
       Exit Function
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
Private Function DBDelete() As Boolean
    DBDelete = False
    
    If Trim(vsGrid1.TextMatrix(vsGrid1.Row, c080_VsGrid.FixedGet(1))) = "" Then
       DBDelete = True
       Exit Function
    End If
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "         選択行を削除します。                  "
        .MB_MSG(3) = "          よろしいですか？                "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
' ( 削除処理 )
    With CisDB
        .SQL = "DELETE Irnaksic "
        .SQL = .SQL & " Where bk = '" & Trim(vsGrid1.TextMatrix(vsGrid1.Row, c080_VsGrid.FixedGet(1))) & "'"
        .SQL = .SQL & "   And rno = 99"
        .DBExec
    End With

' 内容再表示
    SyoriKB = ""
    
    Call DelChg
    
    Call GridSet
    
    DBDelete = True
DBDelete_ED:
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With c080_VsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容)
         Call .FixedSet(中中, 右中, あり, 2, 0, "№")
         Call .FixedSet(中中, 左中, あり, 15, 1, "担当者ｺｰﾄﾞ")
         Call .FixedSet(中中, 左中, あり, 20, 2, "担当者名")
         Call .FixedSet(中中, 中中, あり, 14, 3, "ﾊﾟｽﾜｰﾄﾞﾚﾍﾞﾙ")
         Call .FixedSet(中中, 左中, あり, 0, 4, "PLEBEL")
         Call .FixedSet(中中, 左中, あり, 0, 5, "INFO")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
         Call .FixedSet(中中, 左中, なし, 0, 6, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3")
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
    
    Call c080_VsGrid.Clear
    
' 表示内容読込み
    gSL_Select = "SELECT * FROM Irnaksic "
    gSL_Select = gSL_Select & " WHERE rno = 99"
    gSL_Select = gSL_Select & " ORDER BY bk"
    If Not IRNRead(gSL_Select) Then
        Call IRNClose
        Exit Sub
    End If
' 抽出処理 初期ｾｯﾄ
    Call CISFrm.MousePT(11)  ' ﾏｳｽﾎﾟｲﾝﾀ (砂時計)
'
    PNL_PF.Enabled = False      ' ﾌｧﾝｸｼｮﾝ使用禁止
    

    Ttl_Kensu = CisDB.RecordCount     ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
        .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = c080_VsGrid.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < c080_VsGrid.DispMax Then
            .Rows = c080_VsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 1
    Do Until Not IRN_RDSTS
        With vsGrid1
            .TextMatrix(ID, GridCol(0)) = Format(ID, "###")
            .TextMatrix(ID, GridCol(1)) = IRN.bk
            .TextMatrix(ID, GridCol(2)) = IRN.oyian2
            gStr = CisDB.PassWord(Mid(IRN.oyian, 1, 15))
            .TextMatrix(ID, GridCol(4)) = Mid(gStr, 1, 1)
        '+--------------------------------+
        '+   ｺｰﾄﾞよりｺﾝﾎﾞﾎﾞｯｸｽ内容を求める
        '+--------------------------------+
            If gStr <> "" Then
                With c080_Conbo
                    .Code = Trim(Format(gStr, "0"))
                    If .NameGet Then
                        vsGrid1.TextMatrix(ID, GridCol(3)) = .CombText
                    End If
                End With
            End If
            .TextMatrix(ID, GridCol(5)) = IRN.oyian
            ID = ID + 1
            Call IRNReadNext
        End With
    Loop
    
    Call IRNClose
    
    With vsGrid1
        .Row = 1: .Col = 1
    End With
    Call c080_VsGrid.VSSort(True)
    
    PNL_PF.Enabled = True      ' ﾌｧﾝｸｼｮﾝ使用禁止 解除
    Call CISFrm.MousePT(0)  ' ﾏｳｽﾎﾟｲﾝﾀ (砂時計)
'
End Sub
'******************************************'
'*****        DblClick                *****
'******************************************'
Private Sub vsGrid1_DblClick()
    If vsGrid1.MouseRow < 1 Then Exit Sub
    Call ReturnPress
End Sub
'******************************************'
'*****        RowColChange            *****
'******************************************'
Private Sub vsGrid1_RowColChange()
    If Not FormAct Then Exit Sub
    Call PFDsip
End Sub
'******************************************'
'*****       移動選択項目選択          *****
'******************************************'
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
' Button : 1 = 左ﾎﾞﾀﾝ ,2 = 右ﾎﾞﾀﾝ
    With vsGrid1
        .Tag = ""
        If Button <> 1 Then Exit Sub
    ' 固定行以外は行わない
        If .MouseRow <> 0 Then Exit Sub
    ' 固定列は行わない
        If .MouseCol = 0 Then Exit Sub
    ' 移動元列をタグへ退避
        .Tag = Str(.MouseCol)
        .Drag 1
    End With
End Sub
'******************************************'
'*****        表 示 項 目 移 動        *****
'******************************************'
Private Sub VSGrid1_DragDrop(Source As Control, x As Single, Y As Single)
    With vsGrid1
    ' 固定列は行わない
        If .MouseCol = 0 Then Exit Sub
        If .Tag = "" Then Exit Sub
' ｸﾞﾘｯﾄﾞ表示 禁止
        .Redraw = False
' 列移動
        .ColPosition(CisFun.Val2(.Tag)) = .MouseCol
        .Col = 1: .ColSel = .Cols - 1
' ｸﾞﾘｯﾄﾞ表示 解除
        .Redraw = True
'
        Call c080_VsGrid.VSSort(True)
    End With
End Sub

