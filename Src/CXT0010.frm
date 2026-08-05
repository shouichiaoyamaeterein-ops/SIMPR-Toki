VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Begin VB.Form CXT0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ＴＮＳ送信データ作成処理"
   ClientHeight    =   10845
   ClientLeft      =   1110
   ClientTop       =   3345
   ClientWidth     =   15135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10845
   ScaleWidth      =   15135
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D TNS_Start 
      Height          =   615
      Left            =   5100
      Top             =   510
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1085
      BackColor       =   12632319
      Caption         =   "通信パッケージ システム停止"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   24
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   8421504
      cBoderColor2    =   -2147483643
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      cAlingnment     =   7
      cBoderStyle     =   1
   End
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   14970
      Begin vsFlexLib.vsFlexArray VSGrid1 
         DragIcon        =   "CXT0010.frx":0000
         Height          =   9435
         Left            =   75
         TabIndex        =   0
         Top             =   60
         Width           =   14715
         _Version        =   196608
         _ExtentX        =   25956
         _ExtentY        =   16642
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
         FormatString    =   " №| 契約先(送信先)  |ﾌｧｲﾙ|業務名(ﾌｧｲﾙ名称)|処理結果| 処 理 日 時 |送信件数|ﾃﾞｰﾀ件数|契約先|結果内容|処理結果|履歴管理NO|Select"
         Rows            =   27
         Cols            =   16
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
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 ＴＮＳ送信データ作成　処理  】"
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
Attribute VB_Name = "CXT0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   ＴＮＳ送信データ作成一覧
'**       フォームID    :   CXT0010
'**       処理概要      :
'**
'**       作  成  日    :   2005/09/01  By CIS
'**       変  更  日    :   2006/02/08  By CIS  自動化対応
'**
'**********************************************************'

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'----------------------------------------------- 2006/02/08 Delete
''   #-------------------#
''   # 二 重 起 動 防 止  #
''   #-------------------#
'    If Not Dupli_Start Then End
''   #------------------#
''   # 初 期 内 容 取 得 #
''   #------------------#
'    If Not IniGet Then End
'----------------------------------------------- 2006/02/08 Delete End
    
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
    
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
        .PatanMax = 1                       ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ最大数
        .PatanInit = 0                      ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ初期値
        .SelectCol = 14                     ' セル選択処理ﾜｰｸ列
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
    Call CisFrm.Frm_Center(True, x1024y768)

    Me.Show
    
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    Call InitAdd
    
    If Not DataInput Then
       Call PB_END_Click
    End If
'
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
    Call CisFrm.UnLoadDisp
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
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
   If TNS_Start.Visible Then
       If CisFun.TackCheck("EdiSys01.EXE", 終了しない) = 9 Then
       Else
          TNS_Start.Visible = False
       End If
   End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
'    Call CISFrm.MousePT(11)
    
    If Not HeadChk Then GoTo ReturnPress_Ed
    If Not SendProc Then GoTo ReturnPress_Ed
   
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
        Call .FixedSet(中中, 左中, あり, 24, 1, "契 約 先")
        Call .FixedSet(中中, 中中, あり, 5, 2, "ﾌｧｲﾙ")
        Call .FixedSet(中中, 左中, あり, 36, 3, "業務名(ﾌｧｲﾙ名称)")
        Call .FixedSet(中中, 中中, あり, 4, 4, "結果")
        Call .FixedSet(中中, 中中, あり, 17, 5, "処 理 日 時")
        Call .FixedSet(中中, 右中, あり, 0, 6, "送信件数")
        Call .FixedSet(中中, 右中, あり, 0, 7, "ﾃﾞｰﾀ件数")
        
        Call .FixedSet(中中, 左中, あり, 0, 8, "S契約先")
        Call .FixedSet(中中, 左中, あり, 0, 9, "結果内容")
        Call .FixedSet(中中, 左中, あり, 0, 10, "処理結果")
        Call .FixedSet(中中, 左中, あり, 0, 11, "履歴管理NO")
        Call .FixedSet(中中, 中中, あり, 8, 12, "送信先")
        Call .FixedSet(中中, 中中, あり, 8, 13, "送信先区分")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 14, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,12,3,4,5,6,7")
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
    Dim TNSSendFlg  As Boolean
' 抽出処理 初期ｾｯﾄ
    Call CisFrm.MousePT(0)   ' ﾏｳｽﾎﾟｲﾝﾀ (砂時計)
'
    PNL_PF.Enabled = False      ' ﾌｧﾝｸｼｮﾝ使用禁止
    Ttl_Kensu = CisDB.RecordCount      ' 抽出件数(合計)退避
    TNSSendFlg = False
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VSGrid1
        .Redraw = False
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
        id = 1
        Do Until Not TSR_RDSTS
            .TextMatrix(id, GridCol(0)) = Format(id, "###")
            
            .TextMatrix(id, GridCol(1)) = Trim(TSR.送信先名)
            
            .TextMatrix(id, GridCol(2)) = Trim(TSR.FCODE)
            .TextMatrix(id, GridCol(3)) = Trim(TSR.ファイル名称)
            If Trim(TSR.処理結果) > "10" Then
               .TextMatrix(id, GridCol(4)) = "ｴﾗｰ"
            End If
            .TextMatrix(id, GridCol(10)) = TSR.処理結果
            If Format(TSR.開始日時, "YYYYMMDD") > "20000101" Then
               .TextMatrix(id, GridCol(5)) = Format(TSR.開始日時, "YYYY/MM/DD hh:mm")
               .TextMatrix(id, GridCol(6)) = Format(TSR.送信件数, "#,###")
               .TextMatrix(id, GridCol(7)) = Format(TSR.データ件数, "#,###")
            End If
            
            .TextMatrix(id, GridCol(8)) = Trim(TSR.送信先)
            .TextMatrix(id, GridCol(9)) = Trim(TSR.ERR)
            .TextMatrix(id, GridCol(11)) = Trim(TSR.履歴管理NO)
            
            If TSR.送信先区分 = 0 Then
                .TextMatrix(id, GridCol(12)) = Trim(TSR.送信先)
            Else
                .TextMatrix(id, GridCol(12)) = Trim(TSR.代行送信先)
            End If
           .TextMatrix(id, GridCol(13)) = TSR.送信先区分
            
            If TSR.件数 = 1 Then        ' TNS自動送信ありの時
               TNSSendFlg = True
            End If
            
            If Ttl_Kensu = 1 Then
                Call CisVsGrid.vsColor(個別)
            End If
            
            id = id + 1
            Call TSRReadNext
        Loop
        
        Call TSRClose
        
        .Row = 1: .Col = 1
        .Redraw = True
    End With
    
    If TNSSendFlg Then
       If CisFun.TackCheck("EdiSys01.EXE", 終了しない) = 9 Then
          TNS_Start.Visible = True
       End If
    End If
    
    PNL_PF.Enabled = True      ' ﾌｧﾝｸｼｮﾝ使用禁止 解除
    Call CisFrm.MousePT(0)         ' ﾏｳｽﾎﾟｲﾝﾀ (砂時計)解除
'
End Sub
'******************************************'
'*****        行選択（クリック）       *****
'******************************************'
Private Sub VSGrid1_Click()
    Call CisVsGrid.vsColor
End Sub
'******************************************'
'*****        行選択（スペースｷｰ)      *****
'******************************************'
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    Call CisVsGrid.vsColor
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
            ' エラー内容表示
            .Col = .MouseCol
            If GridErrDisp Then Exit Sub
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
'+     送信設定内容 読み込み  +
'+---------------------------+
Private Function DataInput() As Boolean
    DataInput = False

    CisVsGrid.Clear
    
    gSL_Select = "Select TS.送信先,ISNULL(TR2.略称,'') 送信先名,TS.送信 件数,"
    gSL_Select = gSL_Select & " TS.代行送信先,ISNULL(TR3.略称,'') 代行送信先名,TS.送信先区分,"
    gSL_Select = gSL_Select & " TS.FCODE,TS.ファイル名称,SR.履歴管理NO,"
    gSL_Select = gSL_Select & " SR.開始日時 , SR.送信件数, SR.データ件数, SR.ERR,SR.処理結果"
    gSL_Select = gSL_Select & "  From TNS送信設定テーブル TS"
    gSL_Select = gSL_Select & "   Left Outer Join TNS送信履歴テーブル SR"
    gSL_Select = gSL_Select & "    ON  SR.送信先 = TS.送信先"
    gSL_Select = gSL_Select & "    And SR.FCODE = TS.FCODE"
    gSL_Select = gSL_Select & "    And SR.履歴管理NO = ( Select MAX(履歴管理NO) From TNS送信履歴テーブル"
    gSL_Select = gSL_Select & "                              Where 送信先 = TS.送信先"
    gSL_Select = gSL_Select & "                                  AND  FCODE = TS.FCODE )"
'''    gSL_Select = gSL_Select & "   Left Outer Join 取引先マスタ TR"
'''    gSL_Select = gSL_Select & "    On  TR.取引先CD = TS.送信先"
'''    gSL_Select = gSL_Select & "    And TR.取引先区分 = 0"
    gSL_Select = gSL_Select & "   Left Outer Join 取引先マスタ TR2"
    gSL_Select = gSL_Select & "    On  TR2.取引先CD = TS.送信先"
    gSL_Select = gSL_Select & "    And TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "   Left Outer Join 取引先マスタ TR3"
    gSL_Select = gSL_Select & "    On  TR3.取引先CD = TS.代行送信先"
    gSL_Select = gSL_Select & "    And TR3.取引先区分 = 1"
    
    If Not TSRRead(gSL_Select) Then
        Call TSRClose
        With CisFun
            .MB_Lines = 3
            .MB_MSG(1) = "     ＴＮＳ送信設定情報がありません   "
            .MB_MSG(3) = "          処理を中止します。          "
            .MB_Button = OK
            .MBOX
        End With
        Exit Function
    Else
       Call GridSet
    End If

    DataInput = True
End Function
'+-----------------------------+
'+      送 信 処 理 開 始       +
'+-----------------------------+
Private Function SendProc() As Boolean
    
    SendProc = False
    
    With CisFun
        .MB_Lines = 3
        .MB_MSG(1) = "      処理を開始しますか ？      "
        .MB_Title = "処理確認"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    gInt = 0
    With VSGrid1
        For gLong = 1 To .Rows - 1
           If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(2))) = "" Then
              Exit For
           End If
           If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(CisVsGrid.SelectCol))) = "X" Then
                ReDim Preserve Ts_SndTorcd(gInt + 1)
                ReDim Preserve Ts_SndFCode(gInt + 1)
                ReDim Preserve Ts_SndKTorcd(gInt + 1)
                ReDim Preserve Ts_SndSendKb(gInt + 1)
                ReDim Preserve Ts_SndRow(gInt + 1)

                Ts_SndTorcd(gInt) = Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(12)))
                Ts_SndFCode(gInt) = Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(2)))
                Ts_SndKTorcd(gInt) = Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(8)))
                Ts_SndSendKb(gInt) = Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(13)))
                Ts_SndRow(gInt) = gLong
                gInt = gInt + 1
            End If
        Next gLong
    End With
'
    Ts_ProcType = 0
    Ts_ReSendNo = 0
    Set Ts_Grid = VSGrid1
    With CisVsGrid
         Ts_ColResult = .FixedGet(4)
         Ts_ColResultCD = .FixedGet(10)
         Ts_ColProcData = .FixedGet(5)
         Ts_ColSendCnt = .FixedGet(6)
         Ts_ColDataCnt = .FixedGet(7)
         Ts_ColErrCode = .FixedGet(9)
         Ts_ColRNo = .FixedGet(11)
         Ts_ColSelect = .SelectCol
    End With
'
    Ts_ProcErr = False
    Ts_SeqNoAdd = 0
    
    CXT0020.Show vbModal
    Unload CXT0020
    Set CXT0020 = Nothing
'
'    '------------------------------------------------------------------------------- 2007/03/09 Insert
    '+-----------> 品番ﾁｪｯｸｴﾗｰﾘｽﾄ発行
    If TSS.TNSタイプ = 1 Or TSS.TNSタイプ = 2 Then
       If TSS.品番CHK = 1 And TSS.品番ERRLIST = 1 And C04_HinErrCnt <> 0 Then
          Call C04_ErrList
       End If
    End If
'    '------------------------------------------------------------------------------- 2007/03/09 Insert End

    If Ts_ProcErr Then
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
            .MB_Lines = 3
            .MB_MSG(1) = "            処 理 正 常 終 了                          "
            .MB_Title = "正常終了"
            .MB_Button = OK
            .MBOX
        End With
    End If

End Function
'+--------------------------------+
'+  名称マスタ設定                +
'+--------------------------------+
Private Sub InitAdd()
' 前工程発行JAMA 送信対象契約先
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "TNS前工程送信対象契約先"
        .区分桁数 = 7
        .値 = ""
        .英数字1 = "送信開始日"
        .値名称 = "前工程発行JAMA"
        .作成者 = gTanto
    End With
    If Not SYMInitCreate(True) Then
    End If

'' テキスト管理テーブル　初期データセット
'    Call TEXTInitAdd
''-------------------------- 名称マスタ初期設定
'' 内示工区変換
'    Call ItemsClearSYMCreate
'    With SYMCreate
'        .区分名称 = "内示工区変換"
'        .区分桁数 = 8
'        .値 = ""
'        .英数字1 = "値=契約先(7)+工区(1),空白なし"
'        .英数字2 = "英数字3(1桁目)に変換工区"
'        .値名称 = "内示工区変換"
'        .作成者 = gTanto
'    End With
'    If Not SYMInitCreate(True) Then
'    End If
'' 内示年月
'    Call ItemsClearSYMCreate
'    With SYMCreate
'        .区分名称 = "内示年月"
'        .区分桁数 = 1
'        .値 = "1"
'        .値名称 = "内示対象年月"
'        .英数字2 = "定常"
'        .作成者 = gTanto
'        .制御区分 = "X"
'    End With
'    If Not SYMInitCreate(False) Then
'    End If

End Sub
'+--------------------------------+
'+  エラー内容表示（グリッド処理）
'+--------------------------------+
Private Function GridErrDisp() As Boolean
    GridErrDisp = False
    
    
    With VSGrid1
         If Trim(.TextMatrix(.Row, RTrim(CisVsGrid.FixedGet(10)))) = "" Then Exit Function
          Call ErrDisp(.Row)
          GridErrDisp = True
    End With
End Function
'+--------------------------------+
'+  エラー内容表示
'+--------------------------------+
Private Sub ErrDisp(ERow As Long)
    Dim wErrCd      As String
    Dim wResultCd   As String
    Dim wFName      As String

    With VSGrid1
        wResultCd = .TextMatrix(ERow, RTrim(CisVsGrid.FixedGet(10)))
        wErrCd = .TextMatrix(ERow, RTrim(CisVsGrid.FixedGet(9)))
        wFName = .TextMatrix(ERow, RTrim(CisVsGrid.FixedGet(3)))
    End With
'
    If Trim(wErrCd) <> "" Then
        gSL_Select = "SELECT * FROM TNSERRMSG "
        gSL_Select = gSL_Select & " Where ERRCODE = '" & RTrim(wErrCd) & "'"
        If Not TETRead(gSL_Select, 1) Then Exit Sub
    
        If Trim(TET.ERRLEBEL) <> "" Then
           With CisFun
                .MB_Lines = 7
                .MB_MSG(1) = "   ＴＮＳ送信処理エラー(" & Trim(wFName) & ")     "
                .MB_MSG(3) = "  理由ｺｰﾄﾞ = " & Mid(TET.ERRLEBEL, 1, 3) & _
                             "  内部ｺｰﾄﾞ = " & Mid(TET.ERRLEBEL, 4, 3) & _
                             "  ｴﾗｰﾚﾍﾞﾙ = " & TET.ERRLEBEL & "     "
                .MB_MSG(5) = "  ｴﾗｰ内容  = " & Trim(TET.ERRMSG1) & "      "
                .MB_MSG(7) = "  ｴﾗｰ原因  = " & Trim(TET.ERRMSG2) & "      "
                .MB_Title = "TNS送信処理ｴﾗｰ表示"
                .MB_Button = Error
                .MBOX
           End With
           Exit Sub
        End If
    End If
    gSL_Select = "SELECT * FROM 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '送受信処理結果'"
    gSL_Select = gSL_Select & "   And 値 = '" & wResultCd & "'"
    If Not SYMRead(gSL_Select, 1) Then
       SYM.値名称 = "その他エラー"
       SYM.値 = wResultCd
    End If
    With CisFun
         .MB_Lines = 4
         .MB_MSG(1) = "   【 " & wFName & " 】                  "
         .MB_MSG(3) = SYM.値名称
         .MB_Title = "結果内容"
         If wResultCd > "10" Then
            .MB_Button = Error
         Else
            .MB_Button = OK
         End If
         .MBOX
    End With
End Sub
