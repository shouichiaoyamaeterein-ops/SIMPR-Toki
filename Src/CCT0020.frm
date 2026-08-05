VERSION 5.00
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Begin VB.Form CCT0020 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "TEXT変換処理"
   ClientHeight    =   10845
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   15135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10845
   ScaleWidth      =   15135
   WindowState     =   2  '最大化
   Begin VB.Timer Timer1 
      Left            =   7050
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Left            =   7110
      Top             =   90
   End
   Begin VB.TextBox Dummy 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Left            =   0
      TabIndex        =   6
      Text            =   "Dummy"
      Top             =   30
      Width           =   345
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   855
      Left            =   13005
      ScaleHeight     =   795
      ScaleWidth      =   2115
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   420
      Width           =   2175
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   210
         Top             =   120
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
         Left            =   1170
         Top             =   120
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
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   9645
      Left            =   225
      ScaleHeight     =   9585
      ScaleWidth      =   14910
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   14970
      Begin vsFlexLib.vsFlexArray VSGrid1 
         DragIcon        =   "CCT0020.frx":0000
         Height          =   9345
         Left            =   75
         TabIndex        =   0
         Top             =   120
         Width           =   14715
         _Version        =   196608
         _ExtentX        =   25956
         _ExtentY        =   16484
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
         FormatString    =   $"CCT0020.frx":030A
         Rows            =   26
         Cols            =   20
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   0
         RowHeightMin    =   350
         FillStyle       =   1
         SelectionMode   =   1
         MergeCells      =   4
         BackColorAlternate=   12648447
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   630
      Left            =   255
      ScaleHeight     =   593.265
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   3015
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   495
      Width           =   3075
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   465
         Left            =   90
         Top             =   60
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   820
         ForeColor       =   16711680
         Caption         =   "処　理 単 位"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cBoderWidth     =   2
         cFont3D         =   2
         cAlingnment     =   6
         cPositionX      =   50
         cPositionY      =   20
         Begin VB.OptionButton H1Op_Jyushin 
            BackColor       =   &H00FF8080&
            Caption         =   "個　　別"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Index           =   1
            Left            =   1410
            Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
            TabIndex        =   2
            Top             =   15
            Width           =   1395
         End
         Begin VB.OptionButton H1Op_Jyushin 
            BackColor       =   &H00FFFF00&
            Caption         =   "グループ"
            BeginProperty Font 
               Name            =   "ＭＳ 明朝"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   15
            Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
            TabIndex        =   1
            Top             =   15
            Value           =   -1  'True
            Width           =   1395
         End
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
      Caption         =   "【 テキスト変換　処理  】"
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
   Begin VB.Menu PopMnu 
      Caption         =   "Dummu"
      Visible         =   0   'False
      Begin VB.Menu PopSend 
         Caption         =   "受信～累積"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu PopSend 
         Caption         =   "変換～累積"
         Index           =   1
      End
      Begin VB.Menu PopSend 
         Caption         =   "累積"
         Index           =   2
      End
   End
End
Attribute VB_Name = "CCT0020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   テキスト変換 処理
'**       フォームID    :   CCT0020
'**       処理概要      :
'**
'**       作  成  日    :   2003/12/09  By CIS
'**       変  更  日    :
'**
'**********************************************************'

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    
    Dim mGrpCount       As Integer
    Dim mGrpDisp        As String
    Dim mGrpNow         As Integer      '
    
    Dim wOpObj          As OptionButton
    
    Dim OpGotForeColor  As Long         ' ｵﾌﾟｼｮﾝﾎﾞﾀﾝ　GotFocus色
    Dim OpLostForeColor As Long         ' ｵﾌﾟｼｮﾝﾎﾞﾀﾝ　LostFocus色
    Dim OpGotBackColor  As Long         ' ｵﾌﾟｼｮﾝﾎﾞﾀﾝ　GotFocus色
    Dim OpLostBackColor As Long         ' ｵﾌﾟｼｮﾝﾎﾞﾀﾝ　LostFocus色
    
    Dim mSelCol         As Integer
    Dim mProc           As Boolean
    Dim mInit           As String


    Dim CISFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
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
' ｵﾌﾟｼｮﾝﾎﾞﾀﾝ　色設定
    OpGotForeColor = &H0&
    OpLostForeColor = &HFFFFFF
    
    OpGotBackColor = &HFFFF00
    OpLostBackColor = &HFF8080
    
    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
'   #------------------#
'   #  ｸﾞﾘｯﾄﾞ初期処理   #
'   #------------------#
    With CisVsGrid
        Set .GridObj = VSGrid1              ' ｸﾞﾘｯﾄﾞｵﾌﾞｼﾞｪｸﾄ名 ｾｯﾄ
        .FixedHeight = 380                  ' 固定行(見出し)高さ設定(省略時は300)
        .PatanMax = 2                       ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ最大数
        .PatanInit = 0                      ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ初期値
        .SelectCol = 17                     ' セル選択処理ﾜｰｸ列
        .NoSpaceCol = 2
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
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    Call InitAdd
'
' 初期内容表示
    mInit = "INIT"
    H1Op_Jyushin(0).Value = True
'    Call BodySetGrp("INIT")
    FormAct = Not FormAct
    
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》                 +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Cancel = 1
    Call PB_END_Click
End Sub
'+-------------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》         +
'+-------------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    Select Case Key_Code
           Case vbKeyF1:      H1Op_Jyushin(0).Value = True ': Call H1Op_Jyushin_GotFocus(0)   '【ｸﾞﾙｰﾌﾟ】
           Case vbKeyF2:      H1Op_Jyushin(1).Value = True ': Call H1Op_Jyushin_GotFocus(1)   '【個別】
'           Case vbKeyF8:      VSGrid1.Col = cisvsgrid.FixedGet(16): Call GridErrDisp  '【ｸﾞﾘｯﾄﾞｴﾗｰ内容表示】
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
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Call CisDB.DBDISConnect
    Call CISFrm.UnLoadDisp
    End
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+------------------------------+
'+    受信区分(H1Op_Jyushin)      +
'+------------------------------+
Private Sub H1Op_Jyushin_GotFocus(Index As Integer)
    
    For Each wOpObj In H1Op_Jyushin
        wOpObj.ForeColor = OpLostForeColor
        wOpObj.BackColor = OpLostBackColor
    Next wOpObj
    
    H1Op_Jyushin(Index).ForeColor = OpGotForeColor
    H1Op_Jyushin(Index).BackColor = OpGotBackColor
    
    If Index = 0 Then
        VSGrid1.MergeCells = flexMergeRestrictAll
        VSGrid1.HighLight = flexHighlightNever
        Call CisVsGrid.LayOut(0)
    Else
        VSGrid1.MergeCells = flexMergeNever
        VSGrid1.HighLight = flexHighlightWithFocus
        Call CisVsGrid.LayOut(1)
    End If
    
    mGrpNow = Index
    
    If Index = 0 Then
        Call BodySetGrp
    Else
        Call BodySetKob
    End If
    VSGrid1.SetFocus
    
End Sub
Private Sub H1Op_Jyushin_Click(Index As Integer)
    SendKeys "{TAB}"
End Sub

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CISFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
'    Call CISFrm.MousePT(11)
    
    If Not HeadChk Then GoTo ReturnPress_Ed
    If Not JyushinProc Then GoTo ReturnPress_Ed
   
ReturnPress_Ed:
'   Call CISFrm.MousePT(1)
End Sub
'****************************
'*      Head部 入力ﾁｪｯｸ
'****************************
Private Function HeadChk()
    Dim wSelFlg     As Boolean
    
    HeadChk = False
'
    wSelFlg = False
    With VSGrid1
        .Redraw = False
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(CisVsGrid.SelectCol))) = "X" Then
               wSelFlg = True
               Exit For
            End If
        Next gLong
        .Redraw = True
    End With
    If Not wSelFlg Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    処理するファイルを選択して下さい            "
            .MB_Button = Error
            .MBOX
        End With
        GoTo HeadChk_ED
    End If
    
    HeadChk = True
    
HeadChk_ED:
    PNL_PF.Enabled = True      ' ﾌｧﾝｸｼｮﾝ使用禁止 解除
End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 0, 0, "№")
        Call .FixedSet(中中, 左中, あり, 26, 1, "グループ名")
        Call .FixedSet(中中, 左中, あり, 10, 2, "ﾌｧｲﾙｺｰﾄﾞ")
        Call .FixedSet(中中, 左中, あり, 36, 3, "業務名(ﾌｧｲﾙ名称)")
        Call .FixedSet(中中, 中中, あり, 4, 4, "結果")
        Call .FixedSet(中中, 中中, あり, 17, 5, "処 理 日 時")
        Call .FixedSet(中中, 右中, あり, 0, 6, "処理件数")
        Call .FixedSet(中中, 左中, あり, 20, 7, "変換元ﾌｧｲﾙ名")
        Call .FixedSet(中中, 左中, あり, 20, 8, "累積元ﾌｧｲﾙ名")
        Call .FixedSet(中中, 左中, あり, 0, 9, "発信元")
        Call .FixedSet(中中, 左中, あり, 0, 10, "Fｺｰﾄﾞ")
        Call .FixedSet(中中, 左中, あり, 0, 11, "処理")
        Call .FixedSet(中中, 左中, あり, 0, 12, "GRP")
        Call .FixedSet(中中, 左中, あり, 0, 13, "TNS区分")
        Call .FixedSet(中中, 左中, あり, 0, 14, "受信処理")
        Call .FixedSet(中中, 左中, あり, 0, 15, "受信")
        Call .FixedSet(中中, 左中, あり, 0, 16, "結果内容")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 17, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7")
        Call .PatanSet(1, "0,2,3,4,5,6,7,1")
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim id          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
    Call CISFrm.MousePT(0)   ' ﾏｳｽﾎﾟｲﾝﾀ (砂時計)
'
    PNL_PF.Enabled = False      ' ﾌｧﾝｸｼｮﾝ使用禁止
    Ttl_Kensu = CisDB.RecordCount      ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VSGrid1
        ReDim GridCol(.Cols)
        For id = 0 To .Cols - 1
            GridCol(id) = CisVsGrid.FixedGet(id)
        Next id
        .Redraw = True
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    id = 1
    Do Until Not TKT_RDSTS
        With VSGrid1
            If id > CisVsGrid.DispMax Then
                .Redraw = False
            End If
            .TextMatrix(id, GridCol(0)) = Format(id, "###")
            If mGrpDisp <> Trim(TKT.実行GRP) Then
               mGrpDisp = Trim(TKT.実行GRP)
               mGrpCount = mGrpCount + 1
            End If
            
            If Trim(TKT.GRP名) <> "" Then
               .TextMatrix(id, GridCol(1)) = Trim(TKT.実行GRP) & ":" & Trim(TKT.GRP名)
            Else
               .TextMatrix(id, GridCol(1)) = Trim(TKT.実行GRP)
            End If
            .TextMatrix(id, GridCol(12)) = Trim(TKT.実行GRP)
            
            .TextMatrix(id, GridCol(2)) = Trim(TKT.FILECD)
            .TextMatrix(id, GridCol(3)) = Trim(TKT.F名称)
            If Trim(TKT.処理結果) <> "" Then
               If TKT.処理結果 = "H6" Then
                  .TextMatrix(id, GridCol(4)) = "中止"
               Else
                  .TextMatrix(id, GridCol(4)) = "ｴﾗｰ"
               End If
               .TextMatrix(id, GridCol(16)) = TKT.処理結果
            End If
            If Format(TKT.最終処理日時, "YYYYMMDD") > "20000101" Then
               .TextMatrix(id, GridCol(5)) = Format(TKT.最終処理日時, "YYYY/MM/DD hh:mm")
               .TextMatrix(id, GridCol(6)) = Format(TKT.処理件数, "#,###")
            End If
            .TextMatrix(id, GridCol(7)) = TKT.変換元ファイル名
            .TextMatrix(id, GridCol(8)) = TKT.累積元ファイル名
            
            .TextMatrix(id, GridCol(9)) = Trim(TKT.発信元)
            .TextMatrix(id, GridCol(10)) = Trim(TKT.FCODE)
            
            .TextMatrix(id, GridCol(13)) = TKT.TNS区分
            .TextMatrix(id, GridCol(14)) = TKT.受信処理
            
            ' 受信済みﾌｧｲﾙ数を求める
            .TextMatrix(id, GridCol(15)) = SendKensu(TKT.FCODE, TKT.発信元)
            
            id = id + 1
            Call TKTReadNext
        End With
    Loop
    
    Call TKTClose

    With VSGrid1
        .Row = 1: .Col = 1
    End With

    PNL_PF.Enabled = True      ' ﾌｧﾝｸｼｮﾝ使用禁止 解除
    Call CISFrm.MousePT(0)         ' ﾏｳｽﾎﾟｲﾝﾀ (砂時計)解除
'
End Sub
'******************************************'
'*****        行選択（クリック）       *****
'******************************************'
Private Sub VSGrid1_Click()
    If H1Op_Jyushin(0).Value Then
       gStr = Trim(VSGrid1.TextMatrix(VSGrid1.Row, CisVsGrid.FixedGet(12)))
       Call GrpSelColor(gStr)
    Else
        Call CisVsGrid.vsColor
    End If
End Sub
'******************************************'
'*****        行選択（スペースｷｰ)      *****
'******************************************'
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    If H1Op_Jyushin(0).Value Then
       gStr = Trim(VSGrid1.TextMatrix(VSGrid1.Row, CisVsGrid.FixedGet(12)))
       Call GrpSelColor(gStr)
    Else
        Call CisVsGrid.vsColor
    End If
End Sub
'******************************************'
'*****       移動選択項目選択          *****
'******************************************'
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
' Button : 1 = 左ﾎﾞﾀﾝ ,2 = 右ﾎﾞﾀﾝ
    With VSGrid1
        .Tag = ""
        If Button = 2 Then
            .Col = 1: .ColSel = .Cols - 1
            .Row = .MouseRow
            If .MouseCol = CisVsGrid.FixedGet(4) Then
                ' エラー内容表示
                .Col = .MouseCol
                If GridErrDisp Then Exit Sub
            End If
            Exit Sub
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
'*****        表 示 項 目 移 動        *****
'******************************************'
Private Sub VSGrid1_DragDrop(Source As Control, x As Single, Y As Single)
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

'****************************
'*      ボディ内容セット
'****************************
'+---------------------------+
'+  グループ内容 読み込み      +
'+---------------------------+
Private Function BodySetGrp(Optional LInit As String)
    
    CisVsGrid.Clear
    
    mGrpCount = 0: mGrpDisp = ""
' SQL Server
    gSL_Select = "SELECT TK.*,oyian2 AS GRP名 FROM TEXT管理テーブル TK "
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN Irnaksic"
    gSL_Select = gSL_Select & "         ON  bk = 'Group'"
    gSL_Select = gSL_Select & "         AND oyian = TK.実行GRP "
    gSL_Select = gSL_Select & "  WHERE 連続実行 = 1"
    gSL_Select = gSL_Select & "    AND 実行GRP <> ''"
    gSL_Select = gSL_Select & " ORDER BY 実行GRP,Filecd"
    
    If Not TKTRead(gSL_Select) Then
        Call TKTClose
        Call Timer1_Timer
        DoEvents
'        If LInit <> "INIT" Then
        If mInit <> "INIT" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(1) = "    該当するファイル情報はありません  "
                .MB_MSG(2) = "       個別内容を表示します。        "
                .MB_Button = OK
                .MBOX
            End With
        End If
        mInit = ""
        
        H1Op_Jyushin(1).SetFocus
        
'        Call BodySetKob
        Exit Function
    Else
       Call GridSet
    End If
    If mGrpCount = 1 Then
       With VSGrid1
           Call GrpSelColor("")
           .Row = 1: .RowSel = 1
       End With
    End If
End Function
'+---------------------------+
'+     個別内容 読み込み      +
'+---------------------------+
Private Function BodySetKob()
    CisVsGrid.Clear
    
    gSL_Select = "SELECT * FROM TEXT管理テーブル  "
    gSL_Select = gSL_Select & " ORDER BY Filecd"

    If Not TKTRead(gSL_Select) Then
        Call TKTClose
        With CisFun
            .MB_Lines = 3
            .MB_MSG(1) = "    該当するファイル情報はありません  "
            .MB_Button = OK
            .MBOX
        End With
        Exit Function
    Else
       Call GridSet
'       VSGrid1.HighLight = flexHighlightWithFocus
    End If
End Function
'+------------------------------+
'+ グループ単位　色設定          +
'+------------------------------+
Private Sub GrpSelColor(GrpStr As String)
    Dim sRow        As Long
    With VSGrid1
        If Trim(.TextMatrix(1, CisVsGrid.FixedGet(2))) = "" Then Exit Sub
        If GrpStr = "" Then
           GrpStr = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(12)))
        End If
        
        .Redraw = False
        For gLong = 1 To .Rows - 1
            .Row = gLong
            If Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(12))) = GrpStr Then
                Call CisVsGrid.vsColor
            End If
        Next gLong
        .Redraw = True
    End With
End Sub
'+-----------------------------+
'+      受 信 処 理 開 始       +
'+-----------------------------+
Private Function JyushinProc() As Boolean
    
    JyushinProc = False
    
    With CisFun
        .MB_Lines = 3
        .MB_MSG(1) = "      処理を開始しますか ？      "
        .MB_Title = "処理確認"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    tx_ProcErr = False
    
    Call typTextClear
    
    gInt = 0
    With VSGrid1
        For gLong = 1 To .Rows - 1
           If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(2))) = "" Then
              Exit For
           End If
           If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(CisVsGrid.SelectCol))) = "X" Then
                ReDim Preserve typText.FILECD(gInt + 1)
                ReDim Preserve typText.ProcRow(gInt + 1)
                ReDim Preserve typText.LogDelDay(gInt + 1)
                
                typText.FILECD(gInt) = Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(2)))
                typText.ProcRow(gInt) = gLong
                With CisFun
                    ' 変換元フォルダ
                    .INI_Section = "TextConvert"
                    .INI_Default = gIniExe & "Text\"
                    .INI_Key = "InputDir_" & typText.FILECD(gInt)
                    .GetIni
                    If Trim(.INI_String) <> "" Then
                        ' 変換ログ削除日数
                        .INI_Section = "TextConvert"
                        .INI_Default = "60"
                        .INI_Key = "LogDelDay_" & typText.FILECD(gInt)
                        .GetIni
                        typText.LogDelDay(gInt) = .Val2(.INI_String)
                    Else
                        .INI_Section = "TextConvert"
                        .INI_Default = "60"
                        .INI_Key = "LogDelDay"
                        .GetIni
                        typText.LogDelDay(gInt) = .Val2(.INI_String)
                    End If
                End With
                gInt = gInt + 1
            End If
        Next gLong
    End With
    
    Set tx_GridObj = VSGrid1
    With CisVsGrid
         tx_Result = .FixedGet(16)
         tx_ProcDate = .FixedGet(5)
         tx_Kensu = .FixedGet(6)
         tx_ErMsg = .FixedGet(4)
    End With
    
    CCT0030.Show vbModal
    Unload CCT0030
    Set CCT0030 = Nothing
    
    If tx_ProcErr Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "            処 理 異 常 終 了                          "
            .MB_MSG(3) = "   エラーがあります。エラー内容を確認して下さい。        "
            .MB_Title = "異常終了"
            .MB_Button = OK
            .MBOX
        End With
    Else
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "            処 理 正 常 終 了                          "
            .MB_Title = "正常終了"
            .MB_Button = OK
            .MBOX
        End With
    End If

End Function
'+--------------------------------+
'+  データなし時、基本情報出力      +
'+--------------------------------+
Private Sub InitAdd()
' テキスト管理テーブル　初期データセット
    Call TEXTInitAdd
'-------------------------- 名称マスタ初期設定
' 内示工区変換
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "内示工区変換"
        .区分桁数 = 8
        .値 = ""
        .英数字1 = "値=契約先(7)+工区(1),空白なし"
        .英数字2 = "英数字3(1桁目)に変換工区"
        .値名称 = "内示工区変換"
        .作成者 = gTanto
    End With
    If Not SYMInitCreate(True) Then
    End If
' 内示年月
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "内示年月"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "内示対象年月"
        .英数字2 = "定常"
        .作成者 = gTanto
        .制御区分 = "X"
    End With
    If Not SYMInitCreate(False) Then
    End If

End Sub
'+--------------------------------+
'+  エラー内容表示（グリッド処理）
'+--------------------------------+
Private Function GridErrDisp() As Boolean
    GridErrDisp = False
    
    mSelCol = CisVsGrid.FixedGet(4)    ' 処理結果
    With VSGrid1
        If .Col = mSelCol Then
            Call ErrDisp(.Row)
            GridErrDisp = True
        End If
    End With
End Function
'+--------------------------------+
'+  エラー内容表示
'+--------------------------------+
Private Sub ErrDisp(ERow As Long)
    Dim wErrCd  As String
    Dim wFName  As String

    With VSGrid1
        wErrCd = .TextMatrix(ERow, RTrim(CisVsGrid.FixedGet(16)))
        wFName = .TextMatrix(ERow, RTrim(CisVsGrid.FixedGet(3)))
    End With

    gSL_Select = "SELECT * FROM TEXTERRMSG "
    gSL_Select = gSL_Select & " Where ERRCODE = '" & RTrim(wErrCd) & "'"
    If Not TERRead(gSL_Select, 1) Then Exit Sub

    If Trim(TER.ERRLEBEL) <> "" Then
       With CisFun
            .MB_Lines = 7
            .MB_MSG(1) = "   テキスト変換処理エラー(" & Trim(wFName) & ")     "
            .MB_MSG(3) = "  理由ｺｰﾄﾞ = " & Mid(TER.ERRLEBEL, 1, 3) & _
                         "  内部ｺｰﾄﾞ = " & Mid(TER.ERRLEBEL, 4, 3) & _
                         "  ｴﾗｰﾚﾍﾞﾙ = " & TER.ERRLEBEL & "     "
            .MB_MSG(5) = "  ｴﾗｰ内容  = " & Trim(TER.ERRMSG1) & "      "
            .MB_MSG(6) = "  ｴﾗｰ原因  = " & Trim(TER.ERRMSG2) & "      "
            .MB_Title = "TEXT変換処理ｴﾗｰ表示"
            .MB_Button = Error
            .MBOX
       End With
       Exit Sub
    End If

    With CisFun
         .MB_Lines = 6
         .MB_MSG(1) = "   【 " & wFName & " 】   "
         .MB_MSG(3) = TER.ERRMSG1
         .MB_MSG(5) = TER.ERRMSG2
         .MB_Title = "ｴﾗｰ内容"
         .MB_Button = Error
         .MBOX
    End With
End Sub
'+--------------------------------+
'+  受信済みファイル数を求める
'+--------------------------------+
Private Function SendKensu(wFCd As String, wHMoto As String) As String
'    Dim wFName      As String
'    Dim RPath       As String
'    Dim wFCnt       As Long
'
'    SendKensu = ""
'    wFCnt = 0
'
'    wFName = wFCd & ".*." & wHMoto & "*.Dat"
'
'    Call GetIni("PATH", "RCVPATH", "", RPath, gIniWinDir, "Tnssys.ini")
'
'    gStr = Dir(RPath & wFName) ' 最初のファイル名を返します。
''    '
'    Do While gStr <> ""                ' ループを開始します。
'       wFCnt = wFCnt + 1
'       gStr = Dir                     ' 次のファイル名を返します。
'    Loop
'
'    SendKensu = Format(wFCnt, "###")
    
End Function
