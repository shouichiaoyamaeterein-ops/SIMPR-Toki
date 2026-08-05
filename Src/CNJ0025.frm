VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CNJ0025 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "手配先選択"
   ClientHeight    =   2475
   ClientLeft      =   5295
   ClientTop       =   1155
   ClientWidth     =   7020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Dummy 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   780
      Width           =   195
   End
   Begin CisBtn_60.CisBtn PB_END 
      Height          =   585
      Left            =   6135
      Top             =   1050
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
   Begin CisBtn_60.CisBtn PB_ENT 
      Height          =   585
      Left            =   6135
      Top             =   450
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
   Begin vsFlexLib.vsFlexArray VsGrid1 
      Height          =   1905
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   5970
      _Version        =   196608
      _ExtentX        =   10530
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
      FormatString    =   "№| 手 配 先 |   手   配   先   名   |受入|比率|"
      Rows            =   6
      Cols            =   6
      BackColor       =   12648447
      BackColorBkg    =   16744576
      FocusRect       =   0
      RowHeightMin    =   300
      SelectionMode   =   1
      BackColorAlternate=   12648447
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 手配先選択  】"
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
         Left            =   6015
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
Attribute VB_Name = "CNJ0025"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   生産計画手配先検索
'**       フォームID    :   CNJ0025
'**       処理概要      :
'**
'**       作  成  日    :   2003/11/19  By CIS
'**       変  更  日    :
'**
'**********************************************************'

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim mSL_OrderA      As String       ' ソート順(Order BY)昇順
    Dim mSL_OrderD      As String       ' ソート順(Order BY)降順
    Dim mSL_SelectW     As String       ' SQL SELECT(Where)条件文
    Dim FormCap         As String
    Dim TorWhere        As String
    
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    Dim SCol()          As Long
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
    
    Me.Caption = FormCap
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
'    ProcHB = "H1"
'    Call DispChange(ProcHB)

    '+--------------------------------+
    '+  グリッド初期設定(取引先情報)
    '+--------------------------------+
    With CisVsGrid
        Set .GridObj = VsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 5
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

    Call HeadBodyClear("H")
    RV_Rtn = False
    FormAct = Not FormAct
    
    Call BodySet
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
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
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
    Call PB_END_Click
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'   If RV_Call = "Z" Then
'      Unload CKK0025
'   Else
'      Me.Hide
'   End If
    Unload CNJ0025
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Call ReturnPress
End Sub
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
                .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    Call RtnSet
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      ボディ内容セット
'****************************
Private Function BodySet() As Boolean
    
    BodySet = False
    
    Select Case RTrim$(RV_Kubun)
        Case ""
'                   < 生産計画テーブルより >
                    gSL_Select = "SELECT SK.手配先 取引先,SK.受入,"
                    gSL_Select = gSL_Select & " ISNULL(HT.発注比率,0) 発注比率,"
                    gSL_Select = gSL_Select & " ISNULL(TR.略称,'') 取引先名 "
                    gSL_Select = gSL_Select & " FROM 生産計画テーブル SK "
                    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT "
                    gSL_Select = gSL_Select & "   ON HT.品番 = SK.品番"
                    gSL_Select = gSL_Select & "  AND HT.取引先 = SK.手配先 "
                    gSL_Select = gSL_Select & "  AND HT.受入 = SK.受入 "
                    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR "
                    gSL_Select = gSL_Select & "   ON TR.取引先CD = SK.手配先 "
                    gSL_Select = gSL_Select & "  AND TR.取引先区分 = 1 "
                    gSL_Select = gSL_Select & " WHERE SK.内示年月 = '" & RV_NJYM & "'"
                    gSL_Select = gSL_Select & "   AND SK.計画年月 = '" & RV_YM & "'"
                    gSL_Select = gSL_Select & "   AND SK.品番 = '" & RV_Hinbn & "'"
                    gSL_Select = gSL_Select & " GROUP BY SK.手配先,SK.受入,HT.発注比率,TR.略称"
                    gSL_Select = gSL_Select & " ORDER BY SK.手配先,SK.受入"
                    If Not HTMRead(gSL_Select) Then GoTo BodySet_Err
        Case "HNJ"
'                   < 発注内示テーブルより >
                    gSL_Select = "SELECT HJ.手配先 取引先, HJ.受入,"
                    gSL_Select = gSL_Select & " ISNULL(HT.発注比率,0) 発注比率,"
                    gSL_Select = gSL_Select & " ISNULL(TR.略称,'') 取引先名 "
                    gSL_Select = gSL_Select & " FROM 発注内示テーブル HJ "
                    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT "
                    gSL_Select = gSL_Select & "   ON HT.品番 = HJ.品番"
                    gSL_Select = gSL_Select & "  AND HT.取引先 = HJ.手配先 "
                    gSL_Select = gSL_Select & "  AND HT.受入 = HJ.受入 "
                    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR "
                    gSL_Select = gSL_Select & "   ON TR.取引先CD = HJ.手配先 "
                    gSL_Select = gSL_Select & "  AND TR.取引先区分 = 1 "
                    gSL_Select = gSL_Select & " WHERE HJ.内示年月 = '" & RV_NJYM & "'"
                    gSL_Select = gSL_Select & "   AND HJ.計画年月 = '" & RV_YM & "'"
                    gSL_Select = gSL_Select & "   AND HJ.品番 = '" & RV_Hinbn & "'"
                    gSL_Select = gSL_Select & " GROUP BY HJ.手配先, HJ.受入, HT.発注比率, TR.略称"
                    gSL_Select = gSL_Select & " ORDER BY HJ.手配先, HJ.受入"
                    If Not HTMRead(gSL_Select) Then GoTo BodySet_Err
        Case Else
'                   < 品番取引先マスタより >
                    gSL_Select = "Select HT.取引先, HT.受入, HT.発注比率, ISNULL( TR.略称, '' ) 取引先名"
                    gSL_Select = gSL_Select & " From 品番取引先マスタ HT"
                    gSL_Select = gSL_Select & " Left Outer Join 取引先マスタ TR"
                    gSL_Select = gSL_Select & " On  TR.取引先CD = HT.取引先"
                    gSL_Select = gSL_Select & " And TR.取引先区分 = 1 "
                    gSL_Select = gSL_Select & " Where 品番 = '" & RV_Hinbn & "'"
                    gSL_Select = gSL_Select & " And   品目 = 1"
                    If Not HTMRead(gSL_Select) Then GoTo BodySet_Err
    
    End Select

    gInt = 0
    If VsGrid1.Rows - 1 < CisDB.RecordCount Then VsGrid1.Rows = CisDB.RecordCount + 1
    Do Until Not HTM_RDSTS
    
        With VsGrid1
            gInt = gInt + 1
            .TextMatrix(gInt, 0) = Format(gInt, "##")
            .TextMatrix(gInt, SCol(1)) = RTrim(HTM.取引先)
            .TextMatrix(gInt, SCol(2)) = RTrim(HTM.取引先名)
            .TextMatrix(gInt, SCol(3)) = RTrim(HTM.受入)
            .TextMatrix(gInt, SCol(4)) = RTrim(HTM.発注比率)

        End With
        Call HTMReadNext
    Loop

    Call HTMClose
    BodySet = True
    Exit Function

BodySet_Err:
    With CisFun
        .MB_Lines = 3
        .MB_MSG(1) = "    該当するデータはありません  "
        .MB_Button = OK
        .MBOX
    End With
    Call PB_END_Click
End Function
'+-------------------------------------+
'+      選択内容ｾｯﾄ & 復帰ｽﾃｰﾀｽｾｯﾄ      +
'+-------------------------------------+
Private Sub RtnSet()
    With VsGrid1
        RV_Torcd = .TextMatrix(.Row, SCol(1))
        RV_Ukeir = .TextMatrix(.Row, SCol(3))
    End With
    RV_Rtn = True
    Call PB_END_Click
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
        ReDim SCol(6)
        ' No|取引先|受|背番号|品目|手配|比率|SK区分|ｻｲｸﾙ|収容数1|変更日-便|収容数2|アドレス|契約先|伝票発行|納品書形態|有効期限|受領書形態|受領書発行|明細書発行|か係数|平準化|係数1|係数2|係数3|LT|ﾛｯﾄ①|ﾛｯﾄ②|曜日|日|工数|備考|納入先|納受|収容器1|収容器2|置場|最終発行NO|取引先CD|品目CD|Select
        Call .FixedSet(中中, 右中, あり, 0, 0, "No")
        Call .FixedSet(中中, 左中, あり, 0, 1, "手配先")
        Call .FixedSet(中中, 左中, あり, 0, 2, "手配先名")
        Call .FixedSet(中中, 左中, あり, 0, 3, "受入")
        Call .FixedSet(中中, 中中, あり, 0, 4, "比率")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, あり, 0, 5, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
        
        SCol(1) = .FixedGet(1)  ' 手配先
        SCol(2) = .FixedGet(2)  ' 手配先名
        SCol(3) = .FixedGet(3)  ' 受入
        SCol(4) = .FixedGet(4)  ' 比率
    End With
    
End Sub
'*----------------------*
'*      vsGrid1         *
'*----------------------*
Private Sub VsGrid1_DblClick()
    Call RtnSet
End Sub
