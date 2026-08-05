VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form frmTantoSelect 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "担当者選択"
   ClientHeight    =   7815
   ClientLeft      =   5190
   ClientTop       =   2625
   ClientWidth     =   5370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Back_Now 
      BorderStyle     =   0  'なし
      Height          =   735
      Left            =   150
      ScaleHeight     =   735
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   420
      Width           =   2955
      Begin VB.Label lb_TanName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   660
         TabIndex        =   5
         Top             =   300
         Width           =   2160
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "現在設定担当者"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   570
         TabIndex        =   4
         Top             =   30
         Width           =   1680
      End
      Begin VB.Label lb_TanCD 
         BackColor       =   &H00FFFFFF&
         Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         TabIndex        =   3
         Top             =   300
         Width           =   540
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  '不透明
         Height          =   705
         Left            =   0
         Shape           =   4  '丸みのある長方形
         Top             =   0
         Width           =   2925
      End
   End
   Begin VB.PictureBox PNL_PF 
      Height          =   705
      Left            =   3270
      ScaleHeight     =   645
      ScaleWidth      =   1875
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   450
      Width           =   1935
      Begin CisBtn_60.CisBtn PB_Ent 
         Height          =   585
         Left            =   990
         Top             =   30
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
         BtnCaption      =   "選択"
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   90
         Top             =   30
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
         BtnCaption      =   "未選択"
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
   End
   Begin vsFlexLib.vsFlexArray vsGridS 
      Height          =   6435
      Left            =   120
      TabIndex        =   0
      Top             =   1230
      Width           =   5085
      _Version        =   196608
      _ExtentX        =   8969
      _ExtentY        =   11351
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
      FormatString    =   "№|ｺｰﾄﾞ|担当者名|配列NO|Select|"
      Rows            =   21
      Cols            =   7
      BackColor       =   12648447
      BackColorBkg    =   14737632
      FocusRect       =   0
      RowHeightMin    =   300
      SelectionMode   =   1
      BackColorAlternate=   12648447
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 担当者選択  】"
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
      cPositionX      =   30
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   4320
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
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
End
Attribute VB_Name = "frmTantoSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   担当者選択
'**       フォームID    :   frmTantoSelect
'**       処理概要      :
'**
'**       作  成  日    :   2005/03/03  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim FormCap         As String
    
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim mCount          As Long

'/////////////////////////////////////////////
''   #-------------------#
''   #   担 当 者 選 択   #
''   #-------------------#
'    frmTantoSelect.Show vbModal
'    Unload frmTantoSelect
'    Set frmTantoSelect = Nothing
'/////////////////////////////////////////////


'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    mCount = CisDB.RecordCount
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGridS
        .NoSpaceCol = 2
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 8
        .InitGet = False
        .Init
    End With
    Call GridInit
    
    
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Me.Left = -10000
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    FormAct = Not FormAct
'
    If Not GridSet Then
       Call PB_CAN_Click
       Exit Sub
    End If
    
    FormCap = Me.Caption
    Call CisFrm.Frm_Center(False, x1024y768)
    
    
    PB_Ent.Visible = True
    Back_Now.Visible = True
    
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
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click     '【取消】
           Case vbKeyReturn
                               If PB_Ent.Visible Then                      '【入力】
'                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
'                                  Dummy.Enabled = False
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
Private Sub PB_CAN_Click()
    Call CisDB.DBDISConnect
    Call CisFrm.UnLoadDisp
    End

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
        Call .FixedSet(中中, 中中, あり, 5, 1, "ｺｰﾄﾞ", 英字)
        Call .FixedSet(中中, 左中, あり, 20, 2, "担当者名", 英字)
        Call .FixedSet(中中, 左中, あり, 10, 3, "ｶﾅ(印字用)", 英字)
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3")
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Function GridSet() As Boolean
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
    Dim wSetFlg     As Boolean
    
    GridSet = False

    gSL_Select = "Select * From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '担当管理'"
    gSL_Select = gSL_Select & "   And 区分ﾀｲﾌﾟ = ''"
    gSL_Select = gSL_Select & "   And 値  = '" & gHTanto & "'"
    If SYMRead(gSL_Select, 1) Then
       lb_TanCD = gHTanto
       lb_TanName = Trim(SYM.値名称)
       gTanto = gHTanto
       gTantoName = Trim(SYM.値名称)
    Else
       gHTanto = ""
       gTanto = ""
       gTantoName = ""
    End If
' 抽出処理 初期ｾｯﾄ
    
    gSL_Select = "Select * From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '担当管理'"
    gSL_Select = gSL_Select & "   And 区分ﾀｲﾌﾟ = ''"
    gSL_Select = gSL_Select & " Order By 値 "
    
    If Not SYMRead(gSL_Select) Then
       Call SYMClose
       Exit Function
    End If
    
    Ttl_Kensu = CisDB.RecordCount
'' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGridS
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If

        gCnt1 = 0
        Do Until Not SYM_RDSTS
            gCnt1 = gCnt1 + 1
            .TextMatrix(gCnt1, GridCol(0)) = Format(gCnt1, "####")
            .TextMatrix(gCnt1, GridCol(1)) = SYM.値
            .TextMatrix(gCnt1, GridCol(2)) = SYM.値名称
            .TextMatrix(gCnt1, GridCol(3)) = SYM.英数字1
            
            Call SYMReadNext
        Loop
        
        Call SYMClose

        .Row = 1: .Col = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With
    
    Call CisVsGrid.VSSort
    
    With vsGridS
        wSetFlg = False
        For gCnt1 = 1 To .Rows - 1
            If Trim(.TextMatrix(gCnt1, GridCol(1))) = Trim(gHTanto) Then
               .Row = gCnt1
               wSetFlg = True
               Exit For
            End If
        Next gCnt1
    
        If Not wSetFlg Then
           .Row = 1
        End If
    End With
    
    
    GridSet = True
    
End Function
'+----------------------------+
'+  印刷先選択
'+----------------------------+
Private Sub ReturnPress()
    With vsGridS

        If Trim(.TextMatrix(.Row, 1)) = "" Then Exit Sub

        If Trim(.TextMatrix(.Row, 1)) <> "" Then
            gHTanto = RTrim(.TextMatrix(.Row, 1))
            gTanto = gHTanto
            gTantoName = RTrim(.TextMatrix(.Row, 2))
        End If
    ' 担当情報
        With CisFun
            .INI_FileID = gIniExe & gIniName
            .INI_Section = "担当"
    '       ( 端末単位 )
            .INI_Key = CisFun.WSID: .INI_Default = ""
            .INI_String = Trim(gHTanto)
            .PutIni
        End With
'
    End With
    
    Unload frmTantoSelect
End Sub

Private Sub vsGridS_DblClick()
    If PB_Ent.Visible Then
       Call ReturnPress
    End If
End Sub

