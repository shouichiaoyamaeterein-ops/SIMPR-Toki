VERSION 5.00
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CXP0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "パスワード管理メンテナンス"
   ClientHeight    =   11010
   ClientLeft      =   -90
   ClientTop       =   345
   ClientWidth     =   14985
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   14985
   WindowState     =   2  '最大化
   Begin VB.PictureBox B1_Area1 
      Height          =   9345
      Left            =   1260
      ScaleHeight     =   9285
      ScaleWidth      =   12585
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1500
      Width           =   12645
      Begin vsFlexLib.vsFlexArray VSGrid1 
         DragIcon        =   "CXP0010.frx":0000
         Height          =   9105
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   12405
         _Version        =   196608
         _ExtentX        =   21881
         _ExtentY        =   16060
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
         FormatString    =   "№|業務名          |管理名           |名称                          |ﾊﾟｽﾜｰﾄﾞﾚﾍﾞﾙ| ﾒﾆｭｰ連携 |Exe名     |PLEBEL|PMENU|INNO|Select"
         Rows            =   26
         Cols            =   11
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   350
         FillStyle       =   1
         SelectionMode   =   1
         MergeCells      =   2
         AllowUserResizing=   1
         BackColorAlternate=   12648447
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   1800
      Top             =   975
   End
   Begin VB.TextBox Dummy 
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   150
      Width           =   225
   End
   Begin VB.PictureBox PF_Back 
      BackColor       =   &H00808080&
      Height          =   765
      Left            =   7770
      ScaleHeight     =   705
      ScaleWidth      =   6045
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   6105
      Begin CisBtn_60.CisBtn PB_Can 
         Height          =   585
         Left            =   3360
         Top             =   60
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   3
         BtnCaption      =   "取消"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
      Begin CisBtn_60.CisBtn PB_Pass 
         Height          =   585
         Left            =   45
         Top             =   60
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   1
         BtnCaption      =   "ﾊﾟｽﾜｰﾄﾞ変更"
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
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   4245
         Top             =   60
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "更新"
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
         Left            =   5160
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
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
      Begin CisBtn_60.CisBtn PB_Set 
         Height          =   585
         Left            =   1710
         Top             =   60
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   1
         BtnCaption      =   "ﾊﾟｽﾜｰﾄﾞ設定"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   0
         BtnShadowColor  =   16777215
         PFCaption       =   "( F5 )"
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
         PFBackColor     =   0
      End
   End
   Begin VB.Timer Timer1 
      Left            =   2640
      Top             =   915
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
      Caption         =   "【 パスワード管理メンテナンス  】"
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
   Begin VB.Menu Pop_MenuX 
      Caption         =   "PopMenu"
      Visible         =   0   'False
      Begin VB.Menu Pop_Menu 
         Caption         =   "Pop0"
         Index           =   0
      End
      Begin VB.Menu Pop_Menu 
         Caption         =   "Pop1"
         Index           =   1
      End
      Begin VB.Menu Pop_Menu 
         Caption         =   "Pop2"
         Index           =   2
      End
   End
End
Attribute VB_Name = "CXP0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  パスワード管理メンテナンス
'**       フォームID    :  CXP0010
'**       処理概要      :
'**
'**       作  成  日    :  2004/11/24  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim SyoriKB         As String       ' 処理区分(追加･修正･削除)
    Dim SyoriNM         As String       ' 処理名称(追加･修正･削除)
    
    Dim CISFrm          As CISFormContorl

    Dim CisVsGrid       As New CisVsGrid3
    
    Dim mFNo            As Integer
    Dim mF_Text         As String
    Dim mF_TextStr()    As String
    Dim mf_Eof          As Boolean
    
    Dim mSystem         As String
    Dim mKanriNm        As String
    Dim mGyoumuNm       As String


    Dim mKanri(2)       As String
    
    Dim gMMax           As Integer
    Dim gMM_Inf(15)     As String       'ﾒﾆｭｰ タイトル
    Dim gMM_InfKey(15)  As String       'ﾒﾆｭｰ タイトル(ｷｰ)
    Dim ReadFlg         As Boolean
    
    Dim MenuFlg         As Boolean
    
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   # 二 重 起 動 防 止  #
'   #-------------------#
    If Not Dupli_Start Then End

'    Call GetTemp
    
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
    With CisVsGrid
        Set .GridObj = VSGrid1              ' ｸﾞﾘｯﾄﾞｵﾌﾞｼﾞｪｸﾄ名 ｾｯﾄ
        .FixedHeight = 380                  ' 固定行(見出し)高さ設定(省略時は300)
        .PatanMax = 1                       ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ最大数
        .PatanInit = 0                      ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ初期値
        .SelectCol = 10                     ' セル選択処理ﾜｰｸ列
        .NoSpaceCol = 1
        .NumberSet = NoSpace
'        Set .SelectDispObj = B1lb_SKensu    ' セル選択件数表示ｵﾌﾞｼﾞｪｸﾄｾｯﾄ
' ﾌﾟﾛｸﾞﾗﾑ作成時、以下のﾌﾗｸﾞをｾｯﾄし、ｸﾘｯﾌﾟﾎﾞｰﾄﾞよりｺｰﾄﾞの貼り付けを行う。(GridInit)
        .InitGet = False                    ' ｸﾞﾘｯﾄﾞ内容初期設定 ｸﾘｯﾌﾟﾎﾞｰﾄﾞ作成ﾌﾗｸﾞ

        .Init                               ' ｸﾞﾘｯﾄﾞ初期処理 実行
    End With
    Call GridInit                           ' ｸﾞﾘｯﾄﾞ初期表示
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CISFrm.Frm_Center(True, x1024y768)

    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    Call Kanri
    
    Call GridSet
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    If Not FormAct Then Exit Function
    
    If UCase(IVENT) <> "DOWN" Then Exit Function
    
    Key_Acc = False
    
    Select Case Key_Code
        Case vbKeyF1:       If PB_Pass.Visible Then Call PB_Pass_Click  '【ﾊﾟｽﾜｰﾄﾞ変更】
        Case vbKeyF5:       If PB_Set.Visible Then Call PB_Set_Click    '【ﾊﾟｽﾜｰﾄﾞ設定】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
        Case vbKeyEscape:   If PB_Can.Visible Then Call PB_Can_Click    '【取消】
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
'+    取　消  キ ー(Eacape)   +
'+----------------------------+
Private Sub PB_Can_Click()
    Call GridColorSet("F")
End Sub
'+----------------------------+
'+    パスワード変更(F1) 　   +
'+----------------------------+
Private Sub PB_Pass_Click()
    CXP0030.Show vbModal
    Unload CXP0030
    Set CXP0030 = Nothing
End Sub
'+----------------------------+
'+    パスワード設定(F5)　　   +
'+----------------------------+
Private Sub PB_Set_Click()
    CXP0020.Show vbModal
    Unload CXP0020
    Set CXP0020 = Nothing
End Sub

'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
   Call CISFrm.MousePT(11)
   Call TxtPut
ReturnPress_Ed:
   Call CISFrm.MousePT(1)
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   Call CISFrm.TimeSet("1")
End Sub
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 0, 0, "№")
        Call .FixedSet(中中, 中中, あり, 24, 1, "業務名")
        Call .FixedSet(中中, 中中, あり, 20, 2, "管理名")
        Call .FixedSet(中中, 左中, あり, 40, 3, "名称")
        Call .FixedSet(中中, 中中, あり, 0, 4, "ﾊﾟｽﾜｰﾄﾞﾚﾍﾞﾙ")
        Call .FixedSet(中中, 左中, あり, 0, 5, "ﾒﾆｭｰ連携")
        Call .FixedSet(中中, 左中, あり, 0, 6, "Exe名")
        Call .FixedSet(中中, 左中, あり, 0, 7, "PLEBEL")
        Call .FixedSet(中中, 左中, あり, 0, 8, "PMENU")
        Call .FixedSet(中中, 左中, あり, 0, 9, "INNO")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 10, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'+-----------------------------------------+
'+          レベル選択
'+-----------------------------------------+
Private Sub Pop_Menu_Click(Index As Integer)
    Dim wPG     As String
    Dim wSt     As Long
    Dim wEd     As Long
    With VSGrid1
        .Redraw = False
        gStr = Trim(.TextMatrix(.Row, .Col))
        wPG = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(6)))
        
        wSt = 1
        wEd = .Rows - 1
        
        gCnt1 = CisVsGrid.FixedGetNM(.Col)
        If gCnt1 = 3 Or gCnt1 = 4 Then
            wSt = .Row
            wEd = .Row
        End If
        
        
        For gCnt1 = .Row To wSt Step -1
            If Trim(.TextMatrix(gCnt1, .Col)) = gStr Then
               .TextMatrix(gCnt1, CisVsGrid.FixedGet(7)) = Index
               .TextMatrix(gCnt1, CisVsGrid.FixedGet(4)) = mKanri(Index)
            End If
        Next gCnt1
        For gCnt1 = .Row + 1 To wEd
            If Trim(.TextMatrix(gCnt1, .Col)) = gStr Then
               .TextMatrix(gCnt1, CisVsGrid.FixedGet(7)) = Index
               .TextMatrix(gCnt1, CisVsGrid.FixedGet(4)) = mKanri(Index)
            End If
        Next gCnt1
                        
        For gCnt1 = 1 To .Rows - 1
            If Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(6))) = wPG Then
               .TextMatrix(gCnt1, CisVsGrid.FixedGet(7)) = Index
               .TextMatrix(gCnt1, CisVsGrid.FixedGet(4)) = mKanri(Index)
            End If
        Next gCnt1
                        
        If .Row = 0 Then
            For gCnt1 = .Row + 1 To .Rows - 1
                .TextMatrix(gCnt1, CisVsGrid.FixedGet(7)) = Index
                .TextMatrix(gCnt1, CisVsGrid.FixedGet(4)) = mKanri(Index)
            Next gCnt1
        End If
        
        Call GridColorSet("")

        .Redraw = True
    End With
End Sub
Private Sub GridColorSet(SetType As String)
        
    With VSGrid1
        .Redraw = False
        gCnt3 = .Row
        If SetType = "F" Then
            For gCnt1 = 1 To .Rows - 1
                .TextMatrix(gCnt1, CisVsGrid.FixedGet(7)) = .TextMatrix(gCnt1, CisVsGrid.FixedGet(9))
            Next gCnt1
        End If
        If SetType = "T" Then
            For gCnt1 = 1 To .Rows - 1
                .TextMatrix(gCnt1, CisVsGrid.FixedGet(9)) = .TextMatrix(gCnt1, CisVsGrid.FixedGet(7))
            Next gCnt1
        End If
        gCnt2 = 0
        For gCnt1 = 1 To .Rows - 1
            .Row = gCnt1
            If .TextMatrix(gCnt1, CisVsGrid.FixedGet(7)) <> .TextMatrix(gCnt1, CisVsGrid.FixedGet(9)) Then
               .Col = 0: .CellBackColor = &HC0FFC0
               .Col = CisVsGrid.FixedGet(3): .CellBackColor = &HC0FFC0
               .Col = CisVsGrid.FixedGet(4): .CellBackColor = &HC0FFC0
               gCnt2 = gCnt2 + 1
            Else
               .Col = 0: .CellBackColor = &H8000000F
               .Col = CisVsGrid.FixedGet(3): .CellBackColor = &HC0FFFF
               .Col = CisVsGrid.FixedGet(4): .CellBackColor = &HC0FFFF
            End If
        Next gCnt1
        .Row = gCnt3
        .Redraw = True
        If gCnt2 <> 0 Then
           PB_ENT.Visible = True
           PB_Can.Visible = True
        Else
           PB_ENT.Visible = False
           PB_Can.Visible = False
        End If
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
    Dim TextStr()   As String
    Dim eCnt        As Long

    If Not TextProc("O") Then Exit Sub
    
    mSystem = "SIMPR-A"
    mKanriNm = ""
    mGyoumuNm = ""
    
    With VSGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(ID)
        Next ID
        
        ID = 0
'        '№
'        .TextMatrix(1, GridCol(0)) = 1
'        '業務名
'        .TextMatrix(1, GridCol(1)) = "メニュー"
'        '管理名
'        .TextMatrix(1, GridCol(2)) = "メニュー"
'        '名称
'        .TextMatrix(1, GridCol(3)) = mSystem
'        'Exe
'        .TextMatrix(ID, GridCol(6)) = "SysMenu"
'
'        gCnt2 = getPgmLabel("SIMPR-A")
'        .TextMatrix(ID, GridCol(7)) = gCnt2
'        .TextMatrix(ID, GridCol(9)) = gCnt2
'        .TextMatrix(ID, GridCol(4)) = mKanri(gCnt2)
        
        eCnt = 1
        Do
            With CisFun
                .INI_Section = "subject"
                .INI_Key = "S" & Format(eCnt, "000")
                .INI_Default = ""
                .INI_FileID = mF_Text
                
                .GetIni
                If Trim(.INI_String) = "" Then Exit Do
                gStr = Trim(.INI_String)
                mF_TextStr = Split(gStr, ",")
            End With
            ID = ID + 1
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            '業務名
            .TextMatrix(ID, GridCol(1)) = mF_TextStr(0)
            '管理名
            .TextMatrix(ID, GridCol(2)) = mF_TextStr(1)
            '名称
            .TextMatrix(ID, GridCol(3)) = mF_TextStr(2)
            'Exe
            gCnt1 = InStr(mF_TextStr(3), ".")
            If gCnt1 <> 0 Then
               gStr = Mid(mF_TextStr(3), 1, gCnt1 - 1)
               .TextMatrix(ID, GridCol(6)) = gStr
               gCnt2 = getPgmLabel(gStr)
               .TextMatrix(ID, GridCol(7)) = gCnt2
               .TextMatrix(ID, GridCol(9)) = gCnt2
               .TextMatrix(ID, GridCol(4)) = mKanri(gCnt2)
            End If
            
            eCnt = eCnt + 1
        Loop
        
        
        
        
        Do Until mf_Eof
            If Not TextProc("GET") Then Exit Sub
            If mf_Eof Then Exit Do
            
            If ID = 1 Then
               '名称
               .TextMatrix(1, GridCol(3)) = mSystem
            End If
            
            ID = ID + 1
            
            If ID >= CisVsGrid.DispMax Then
                .Rows = ID + 1
            End If
                        
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            '業務名
            .TextMatrix(ID, GridCol(1)) = mGyoumuNm
            '管理名
            .TextMatrix(ID, GridCol(2)) = mKanriNm
            '名称
            .TextMatrix(ID, GridCol(3)) = mF_TextStr(4)
            'Exe
            gCnt1 = InStr(mF_TextStr(6), ".")
            If gCnt1 <> 0 Then
               gStr = Mid(mF_TextStr(6), 1, gCnt1 - 1)
               .TextMatrix(ID, GridCol(6)) = gStr
               gCnt2 = getPgmLabel(gStr)
               .TextMatrix(ID, GridCol(7)) = gCnt2
               .TextMatrix(ID, GridCol(9)) = gCnt2
               .TextMatrix(ID, GridCol(4)) = mKanri(gCnt2)
            End If
                 
        Loop
       .Row = 1: .Col = 1: .ColSel = .Cols - 1
       .Redraw = True
    End With
End Sub
'+-------------------------------------+
Private Function TextProc(ProcType As String) As Boolean
    TextProc = False
    On Error GoTo TextProc_Er
    
    mf_Eof = False
    If mF_Text = "" Then
    '   < メニューテキスト取得 >
        With CisFun
            .INI_Section = "MenuText"
            .INI_Key = .WSID
            .INI_Default = ""
            .INI_FileID = gIniExe & gIniName
            
            .GetIni
            mF_Text = .INI_String
        
            If RTrim$(mF_Text) = "" Then
                mF_Text = gIniExe & "Menu.Txt"
                
                .INI_String = mF_Text
                .PutIni
            End If
            .INI_Section = "Menu"
            .INI_Key = "HCNT"
            .INI_Default = ""
            .INI_FileID = mF_Text
            
            .GetIni
            gMMax = CisFun.Val2(.INI_String)
            If gMMax = 0 Then
               MenuFlg = False
            Else
               MenuFlg = True
            End If
        End With
    End If
    
    If MenuFlg Then
       If Not TextProcIni(ProcType) Then Exit Function
       TextProc = True
       Exit Function
    End If
    ' テキストファイル オープン
    If mFNo = 0 Or ProcType = "O" Then
       mFNo = FreeFile

       Open mF_Text For Input Access Read As mFNo
    End If
    ' テキストファイル クローズ
    If ProcType = "C" Then
       Close #mFNo
       mFNo = 0
    End If
    
    ' テキストファイル 読み込み
    If ProcType = "GET" Then
       Do While Not EOF(mFNo)
            
          Line Input #mFNo, gStr
        
          If Mid(gStr, 1, 1) = "S" Then
             mF_TextStr = Split(gStr, ",")
             
             mSystem = mF_TextStr(2)
          End If
        
          If Mid(gStr, 1, 1) = "H" Then
             mF_TextStr = Split(gStr, ",")
             
             If mF_TextStr(2) <> "END" Then
                mGyoumuNm = mF_TextStr(2)
             End If
          End If
          If Mid(gStr, 1, 1) = "M" Then
             mF_TextStr = Split(gStr, ",")
             
             If mF_TextStr(3) <> "" Then
                mKanriNm = mF_TextStr(3)
             End If
          End If
          If Mid(gStr, 1, 1) = "B" Then
             mF_TextStr = Split(gStr, ",")
             If mF_TextStr(6) Like "PXP*" Then
             Else
                If mF_TextStr(5) = "" Or mF_TextStr(5) = "S" Or mF_TextStr(5) = "D" Then
                   Exit Do
                End If
             End If
          End If
        Loop
        
        If EOF(mFNo) Then
           mf_Eof = True
        End If
    End If
    
    TextProc = True
    Exit Function
TextProc_Er:
    If Err = 53 Then
        gStr = "初期導入作業が正しく行われていません。" & vbCr & vbCr & _
                 Space(10) & "契約処理を行ってから、再度実行して下さい。"
        MsgBox gStr, vbCritical, ""
        Exit Function
    End If
    
    Call CisFun.ErrorBox("テキスト処理")
End Function
Private Function TextProcIni(ProcType As String) As Boolean
    Dim mIdx        As Long
    Dim mIdx2       As Long
    Dim mStr        As String
    Dim wFT_RText   As String
    Dim wFT_Split() As String
    Dim wFT_I       As Integer
    
    TextProcIni = False
    On Error GoTo TextProcIni_Er
    
    mf_Eof = False
    If mF_Text = "" Then
    '   < メニューテキスト取得 >
        With CisFun
            .INI_Section = "MenuText"
            .INI_Key = .WSID
            .INI_Default = ""
            .INI_FileID = gIniExe & gIniName
            
            .GetIni
            mF_Text = .INI_String
        
            If RTrim$(mF_Text) = "" Then
                mF_Text = gIniExe & "Menu.Txt"
                
                .INI_String = mF_Text
                .PutIni
            End If
        End With
    End If
    
    ' テキストファイル オープン
    If mFNo = 0 Or ProcType = "O" Then
       mFNo = FreeFile

       Open mF_Text For Input Access Read As mFNo
    
        With CisFun
            .INI_Section = "Menu"
            .INI_Key = "HCNT"
            .INI_Default = ""
            .INI_FileID = mF_Text
            
            .GetIni
            gMMax = CisFun.Val2(.INI_String)
            
            Erase gMM_Inf
            Erase gMM_InfKey
    
            ' ﾒﾆｭｰﾀｲﾄﾙ取得
            For wFT_I = 0 To gMMax
                .INI_Section = "Menu"
                .INI_Key = Format(wFT_I, "0")
                .INI_Default = ""
                .INI_FileID = mF_Text
                
                .GetIni
                mStr = Trim(.INI_String)
                If mStr <> "" Then
                   wFT_Split = Split(mStr, ",")
                   If UBound(wFT_Split) > 0 Then
                      gMM_Inf(wFT_I) = wFT_Split(0)
                      gMM_InfKey(wFT_I) = wFT_Split(1)
                   End If
                End If
            Next wFT_I
        End With
        ReadFlg = False
    End If
    ' テキストファイル クローズ
    If ProcType = "C" Then
       Close #mFNo
       mFNo = 0
    End If
    
    ' テキストファイル 読み込み
    If ProcType = "GET" Then
       Do While Not EOF(mFNo)
            
          Line Input #mFNo, wFT_RText
        
          If Mid(wFT_RText, 1, 1) <> ";" And Trim(wFT_RText) <> "" Then

             If Mid(wFT_RText, 1, 3) = "99=" Then
                ReadFlg = True
             End If
            
             If ReadFlg Then
                If Mid(wFT_RText, 1, 1) = "[" Then
                   mStr = Replace(Trim(wFT_RText), "[", "")
                   mStr = Replace(mStr, "]", "")
                   For mIdx = 0 To UBound(gMM_Inf) - 1
                       If gMM_InfKey(mIdx) = mStr Then
                          mGyoumuNm = gMM_Inf(mIdx)
                          Exit For
                       End If
                   Next mIdx
                End If
                If Mid(wFT_RText, 1, 1) = "M" Then
                   mKanriNm = Mid(wFT_RText, 4)
                End If
                If Mid(wFT_RText, 1, 1) = "B" Then
                   ReDim mF_TextStr(6)
                   mStr = Mid(wFT_RText, 6)
                   wFT_Split = Split(mStr, ",")
                   If UBound(wFT_Split) > 1 Then
                      If wFT_Split(2) Like "PXP*" Then
                      Else
                         If wFT_Split(1) = "" Or wFT_Split(1) = "S" Or wFT_Split(1) = "D" Then
                            mF_TextStr(4) = wFT_Split(0)
                            mF_TextStr(6) = wFT_Split(2)
                            Exit Do
                         End If
                      End If
                   End If
                End If
             End If
          End If
        
        Loop
        
        If EOF(mFNo) Then
           mf_Eof = True
        End If
    End If
    
    TextProcIni = True
    Exit Function
TextProcIni_Er:
    If Err = 53 Then
        gStr = "初期導入作業が正しく行われていません。" & vbCr & vbCr & _
                 Space(10) & "契約処理を行ってから、再度実行して下さい。"
        MsgBox gStr, vbCritical, ""
        Exit Function
    End If
    
    Call CisFun.ErrorBox("テキスト処理")
End Function

Private Function getPgmLabel(ATitle As String) As Long
    Dim pStr        As String
    Dim pStr2       As String
    getPgmLabel = 2         ' 未登録時は管理責任者
    pStr = "Lebelp" & ATitle
    gSL_Select = "Select bk,rno,oyian,oyian2 From Irnaksic "
    gSL_Select = gSL_Select & " Where bk = '" & pStr & "'"
    gSL_Select = gSL_Select & "   And rno = 0"
    If IRNRead(gSL_Select, 1) Then
        pStr2 = RTrim(IRN.oyian2)
        pStr2 = Trim(CisDB.PassWord(pStr2))
        getPgmLabel = CisFun.Val2(Mid(pStr2, Len(pStr) + 1, 1))
    End If
    If getPgmLabel < 0 Then getPgmLabel = 0

End Function

'******************************************'
'*****       移動選択項目選択          *****
'******************************************'
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Button : 1 = 左ﾎﾞﾀﾝ ,2 = 右ﾎﾞﾀﾝ
    With VSGrid1
        .Tag = ""
        If Button = 2 Then  ' 右ボタン
           If .MouseRow = 0 Then
              If .MouseCol <> 0 Then Exit Sub
           End If
           .Row = .MouseRow
           .Col = .MouseCol
           PopupMenu Pop_MenuX
        End If
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
'*****                                *****
'******************************************'
Private Sub VSGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
End Sub
'******************************************'
'*****        表 示 項 目 移 動        *****
'******************************************'
Private Sub VSGrid1_DragDrop(Source As Control, X As Single, Y As Single)
    With VSGrid1
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
        Call CisVsGrid.VSSort(True)
    End With
End Sub
'+------------------------------------+
'+      管理情報出力
'+------------------------------------+
Private Sub Kanri()
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "管理名"
        .区分桁数 = 1
        .値 = "0"
        .値名称 = "担当者"
        .制御区分 = "X"
        .作成者 = gTanto
        If Not SYMInitCreate(True) Then
            .値 = "1"
            .値名称 = "管理者"
            Call SYMInitCreate
            .値 = "2"
            .値名称 = "管理責任者"
            Call SYMInitCreate
        End If
    End With
    
    Erase mKanri
    
    gSL_Select = "Select * From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '管理名'"
    gSL_Select = gSL_Select & "   And 値 <> ''"
    gSL_Select = gSL_Select & " Order By 値 "
    If SYMRead(gSL_Select) Then
       Do Until Not SYM_RDSTS
          gCnt1 = CisFun.Val2(SYM.値)
          If gCnt1 > 2 Then Exit Do
            
          mKanri(gCnt1) = Trim(SYM.値名称)
          Pop_Menu(gCnt1).Caption = mKanri(gCnt1)
          Call SYMReadNext
       Loop
    End If
    Call SYMClose
    
End Sub
'**************************************
'*                                    *
'*       ファイル入力 ＆ 表示       *
'*                                    *
'**************************************
Public Function TxtPut() As Boolean
    
    On Error GoTo TxtPut_er:
    With CisFun
        .MB_Lines = 3
        .MB_MSG(1) = "      更新処理を開始しますか ？      "
        .MB_Title = "処理確認"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    
    Call CisDB.DBTran(TransBegin)
    
'    gSL_Select = "Delete Irnaksic "
'    gSL_Select = gSL_Select & " Where bk Like 'Lebelp%'"
'    gSL_Select = gSL_Select & "   And rno = 0"
'    CisDB.SQL = gSL_Select
'    Call CisDB.DBExec
    
    With VSGrid1
        For gCnt1 = 1 To .Rows - 1
            If Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(6))) = "" Then Exit For
                        
            gStr = Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(6)))
'            If UCase(gStr) = "SYSMENU" Then
'               gStr = "SIMPR-A"
'            End If
            gSL_Select = "Select * From Irnaksic "
            gSL_Select = gSL_Select & " Where bk = '" & "Lebelp" & gStr & "'"
            gSL_Select = gSL_Select & "   And rno = 0"
            If Not IRNRead(gSL_Select, 1) Then
               Call ItemsClearIRN
               IRN.bk = "Lebelp" & gStr
               gStr = IRN.bk & Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(7)))
               gStr = Trim(CisDB.PassWord(gStr))
               IRN.oyian2 = gStr
               Call IRNInsert
            Else
               If Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(7))) <> _
                  Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(9))) Then
                  gStr = IRN.bk & Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(7)))
                  gStr = Trim(CisDB.PassWord(gStr))
                  IRN.oyian2 = gStr
                  Call IRNUpdate
               End If
            End If
        Next gCnt1
    End With
    
'    gSL_Select = "Select * From Irnaksic "
'    gSL_Select = gSL_Select & " Where bk = '" & "Lebelp" & Trim(App.Title) & "'"
'    gSL_Select = gSL_Select & "   And rno = 0"
'    If Not IRNRead(gSL_Select, 1) Then
'       Call ItemsClearIRN
'       IRN.bk = "Lebelp" & Trim(App.Title)
'       gStr = Trim(App.Title) & "2"
'       gStr = Trim(CisDB.PassWord(gStr))
'       IRN.oyian2 = gStr
'       Call IRNInsert
'    Else
'       gStr = Trim(App.Title) & "2"
'       gStr = Trim(CisDB.PassWord(gStr))
'       IRN.oyian = gStr
'       Call IRNUpdate
'    End If
    Call CisDB.DBTran(TransCommit)
    
    Call GridColorSet("T")
    
    Exit Function
TxtPut_er:
    Call CisFun.ErrorBox("更新ｴﾗｰ")
    Call CisDB.DBTran(TransRollback)
End Function
