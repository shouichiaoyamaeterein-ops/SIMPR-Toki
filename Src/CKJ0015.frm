VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CKJ0015 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "品番工順選択"
   ClientHeight    =   3285
   ClientLeft      =   2850
   ClientTop       =   3975
   ClientWidth     =   7440
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PNL_PF 
      Height          =   705
      Left            =   5370
      ScaleHeight     =   645
      ScaleWidth      =   1875
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   420
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
         BtnCaption      =   "戻る"
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
      Height          =   1905
      Left            =   120
      TabIndex        =   0
      Top             =   1230
      Width           =   7185
      _Version        =   196608
      _ExtentX        =   12674
      _ExtentY        =   3360
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
      FormatString    =   "№|工順|    機 械 NO            |    機  械              |||"
      Rows            =   6
      Cols            =   8
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
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 品番工順選択  】"
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
         Left            =   8790
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
   Begin VB.Label lb_Hinbn 
      BackStyle       =   0  '透明
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXX5"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1020
      TabIndex        =   2
      Top             =   690
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "品番"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   390
      TabIndex        =   1
      Top             =   690
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  '不透明
      Height          =   495
      Left            =   150
      Shape           =   4  '丸みのある長方形
      Top             =   570
      Width           =   4065
   End
End
Attribute VB_Name = "CKJ0015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   品番取引先選択
'**       フォームID    :   CKJ0015 (Based CKN0215)
'**       処理概要      :
'**
'**       作  成  日    :   2009/12/25  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim FormCap         As String
    
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim mCount          As Long
  ''Const DebugFlg      As Boolean = False  ' 試験時True
  
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
        .NoSpaceCol = 1
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
    If RV_Left <> 0 Or RV_Top <> 0 Then
        Me.Move RV_Left, RV_Top
    Else
        Call CisFrm.Frm_Center(False, x1024y768)
    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    FormCap = Me.Caption
    
    RV_Rtn = False
    
    FormAct = Not FormAct
'
    Rv_CallRtn = 0          ' 2009.12.25 add
    Call GridSet
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
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
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
    Call PB_CAN_Click
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
    Unload CKJ0015          ' 2009.12.24 upd : Unload CKN0215
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
        Call .FixedSet(中中, 左中, あり, 4, 1, "工順")
        Call .FixedSet(中中, 左中, あり, 24, 2, "機  械  №")
        Call .FixedSet(中中, 左中, あり, 24, 3, "機  械")
        Call .FixedSet(中中, 左中, あり, 10, 4, "記憶機械№")
        Call .FixedSet(中中, 左中, あり, 10, 5, "記憶機械")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3")
      ''If DebugFlg Then Call .PatanSet(0, "0,1,2,3,4,5")   ' 試験用
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
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = mCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
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
    End With
    
    lb_Hinbn.Caption = HKT.品番
        
    ID = 0
    Do Until Not HKT_RDSTS
        With vsGridS
            ID = ID + 1
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            '
            .TextMatrix(ID, GridCol(1)) = CStr(HKT.工順)
            If Trim(HKT.機械NO名) <> "" Then
               .TextMatrix(ID, GridCol(2)) = HKT.機械NO & "：" & HKT.機械NO名
            Else
               .TextMatrix(ID, GridCol(2)) = HKT.機械NO
            End If
            If Trim(HKT.機械名) <> "" Then
               .TextMatrix(ID, GridCol(3)) = HKT.機械 & "：" & HKT.機械名
            Else
               .TextMatrix(ID, GridCol(3)) = HKT.機械
            End If
            .TextMatrix(ID, GridCol(4)) = HKT.機械NO
            .TextMatrix(ID, GridCol(5)) = HKT.機械
       End With
        Call HKTReadNext '(1)
    Loop
    Call HKTClose(1)
    With vsGridS
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'+----------------------------+
'+  品番取引先選択
'+----------------------------+
Private Sub ReturnPress()
    With vsGridS
        If Trim(.TextMatrix(.Row, 1)) = "" Then Exit Sub
        
        Rv_CallRtn = CisFun.Val2(.TextMatrix(.Row, 1))  ' 工順
        RV_Code = .TextMatrix(.Row, 4)                  ' 機械NO
    End With
    RV_Rtn = True
    Call PB_CAN_Click
End Sub
