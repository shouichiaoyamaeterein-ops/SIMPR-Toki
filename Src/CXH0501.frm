VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXH0501 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "入出庫入力データ検索"
   ClientHeight    =   7740
   ClientLeft      =   90
   ClientTop       =   1530
   ClientWidth     =   10515
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
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
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   7740
   ScaleWidth      =   10515
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1020
      Left            =   105
      ScaleHeight     =   960
      ScaleWidth      =   5010
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   780
      Width           =   5070
      Begin Cis3D_v60.CIS3D UC_3D 
         Height          =   885
         Index           =   0
         Left            =   45
         Top             =   45
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   1561
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin Cis3D_v60.CIS3D CIS3D15 
            Height          =   795
            Index           =   0
            Left            =   45
            Top             =   45
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "日　付"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionX      =   45
            cPositionY      =   45
            Begin CisYMD_With_Btn.CisYMDwB iH1_Ymd 
               Height          =   360
               Left            =   90
               TabIndex        =   8
               Top             =   330
               Width           =   1950
               _ExtentX        =   2990
               _ExtentY        =   582
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Object.Height          =   360
               Object.Width           =   1950
               cSize           =   1
               cChkResult      =   0   'False
               cTypeIntval     =   0
               BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin Cis3D_v60.CIS3D Back_Jkbn 
            Height          =   795
            Left            =   2190
            Top             =   45
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "実 績 区 分"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionX      =   45
            cPositionY      =   45
            Begin VB.CheckBox H1Ck_JKbn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "出庫"
               Height          =   285
               Index           =   2
               Left            =   960
               TabIndex        =   11
               Top             =   390
               Width           =   825
            End
            Begin VB.CheckBox H1Ck_JKbn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "支給"
               Height          =   285
               Index           =   3
               Left            =   1800
               TabIndex        =   10
               Top             =   390
               Width           =   825
            End
            Begin VB.CheckBox H1Ck_JKbn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "入庫"
               Height          =   285
               Index           =   1
               Left            =   120
               TabIndex        =   9
               Top             =   390
               Width           =   825
            End
         End
      End
   End
   Begin Cis3D_v60.CIS3D HLB_Title 
      Height          =   675
      Left            =   120
      Top             =   60
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   1191
      BackColor       =   12648384
      Caption         =   "入出庫入力データ検索"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      IMEMode         =   3  'ｵﾌ固定
      Left            =   180
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   150
      Width           =   855
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   6840
      ScaleHeight     =   630
      ScaleWidth      =   3465
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   60
      Width           =   3525
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   1740
         Top             =   15
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Left            =   2595
         Top             =   15
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   885
         Top             =   15
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   30
         Top             =   15
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
         BtnType         =   11
         BtnCaption      =   "検索"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   33023
         BtnShadowColor  =   16777215
         PFCaption       =   "( F7 )"
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
         PFBackColor     =   33023
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5745
      Left            =   105
      ScaleHeight     =   5685
      ScaleWidth      =   10215
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1875
      Width           =   10275
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CXH0501.frx":0000
         Height          =   5550
         Left            =   60
         OleObjectBlob   =   "CXH0501.frx":0015
         TabIndex        =   0
         Top             =   60
         Width           =   10080
      End
   End
   Begin VB.PictureBox Back_Sort 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4980
      ScaleHeight     =   615
      ScaleWidth      =   1695
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   75
      Width           =   1755
      Begin VB.OptionButton Op_Sort 
         BackColor       =   &H00808000&
         Caption         =   "降順"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   900
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   675
      End
      Begin VB.OptionButton Op_Sort 
         BackColor       =   &H00808000&
         Caption         =   "昇順"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   0
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   105
         Value           =   -1  'True
         Width           =   705
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   105
      Top             =   150
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "CXH0501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   入出庫入力データ検索
'**       フォームID    :   CXH0501
'**       処理概要      :
'**
'**       作  成  日    :   2004/06/23  By CIS
'**       変  更  日    :   2004/06/29  By CIS - バグ対応（データ検索時、エラー発生）
'**       変  更  日    :   2004/12/29  By CIS - 条件表示不具合
'**       変  更  日    :   2014/02/04  By CIS WIN7対応
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim Sl_Where        As String

    Dim FormCap         As String

    Dim CisFrm          As CISFormContorl
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

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
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
    Op_Sort(0).Value = True
'   ADOｺﾝﾄﾛｰﾙ　接続情報　セット
    Adodc1.ConnectionString = CisDB.ConnectString
    '-----2014/02/04 EDIT START-----
    gStr = CisDB.ConnectString
    gConnect = "server=" & CisDB.Server & ";database=" & CisDB.DBName & ";User ID=" & CisDB.User & ";password=" & CisDB.PassWord(GetSetting(gRegKey, "Environment", "DrowsSap", ""))
    
    gStr = Replace(gStr, "Use Procedure", "Extended Properties=" & gConnect & ";Use Procedure")
    
    Adodc1.ConnectionString = gStr
    '-----2014/02/04 EDIT END-------
    Call HeadBodyClear("H")
    RV_Rtn = False
'
    FormAct = Not FormAct
'
    If RV_Call <> "" Then Call ReturnPress
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
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
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
    Call PB_END_Click
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Unload CXH0501
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
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        Call HeadBodyClear("H")
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Ymd" Then iH1_Ymd.ShowCalender: Exit Sub
End Sub
'+-----------------------------+
'+      iH1_Ymd
'+-----------------------------+
Private Sub iH1_Ymd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ymd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
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

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1": H1Mode = True: H1Color = gPGotSel: PB_ENT.BtnCaption = "開始"
       Case "B1": B1Mode = True: B1Color = gPGotSel: PB_ENT.BtnCaption = "選択"
    End Select
    H1_Area1.Enabled = H1Mode: H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode: B1_Area1.BackColor = B1Color

    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode

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
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
                .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
    If HeadBody = "B" Then
       DBGrid1.Close True
       Adodc1.Recordset.Close
       Adodc1.RecordSource = ""
       Me.Caption = FormCap
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
       If Not Head1Chk Then GoTo ReturnPress_Ed
       ProcHB = "B1"
       Call DispChange(ProcHB)
       GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then Call RtnSet
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False

    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'   >> 日付
    If RTrim(iH1_Ymd) <> "" Then
        If CisFun.Val2(iH1_Ymd.cDay) = 0 Then
        Else
            If Not iH1_Ymd.cChkResult Then
                CisFun.MB_MSG(2) = "    日付を正しく入力して下さい            "
                CisFun.MBOX
                iH1_Ymd.SetFocus
                Exit Function
            End If
        End If
    End If

    If Not BodySet Then Exit Function

    Head1Chk = True
End Function
'****************************
'*      分類切替
'****************************
Private Sub Op_Sort_Click(Index As Integer)
    If ProcHB = "H1" Then Exit Sub
    If Adodc1.Recordset.RecordCount <> 0 Then Call BodySet
    SendKeys "{TAB}"
End Sub
'****************************
'*      ボディ内容セット
'****************************
Private Function BodySet() As Boolean
    BodySet = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  jkt.納入番号"
    gSL_Select = gSL_Select & ", jkt.受入"
    gSL_Select = gSL_Select & ", case"
    gSL_Select = gSL_Select & "    when jkt.実績区分 = 1 then '入庫'"
    gSL_Select = gSL_Select & "    when jkt.実績区分 = 2 then '出庫'"
    gSL_Select = gSL_Select & "    when jkt.実績区分 = 3 then '支給'"
    gSL_Select = gSL_Select & "    else ''"
    gSL_Select = gSL_Select & "  end 区分名"
    gSL_Select = gSL_Select & ", case"
    gSL_Select = gSL_Select & "    when jkt.実績日 <> '' and jkt.実績便 <> '' then"
    gSL_Select = gSL_Select & "         substring(jkt.実績日,1,4) + '/'"
    gSL_Select = gSL_Select & "       + substring(jkt.実績日,5,2) + '/'"
    gSL_Select = gSL_Select & "       + substring(jkt.実績日,7,2) + '-'"
    gSL_Select = gSL_Select & "       + jkt.実績便"
    gSL_Select = gSL_Select & "    when jkt.実績日 <> '' and jkt.実績便 = '' then"
    gSL_Select = gSL_Select & "         substring(jkt.実績日,1,4) + '/'"
    gSL_Select = gSL_Select & "       + substring(jkt.実績日,5,2) + '/'"
    gSL_Select = gSL_Select & "       + substring(jkt.実績日,7,2) + '   '"
    gSL_Select = gSL_Select & "    else ''"
    gSL_Select = gSL_Select & "  end 編集実績日"
    gSL_Select = gSL_Select & ", case when jkt.実績先 = '' then '' else jkt.実績先 + ':' + isnull(tr1.略称,'') end 実績先名"
    gSL_Select = gSL_Select & ", case when jkt.実績元 = '' then '' else jkt.実績元 + ':' + isnull(tr2.略称,'') end 実績元名"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  実績管理テーブル jkt"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ tr1"
    gSL_Select = gSL_Select & "    on tr1.取引先CD = jkt.実績先"
    gSL_Select = gSL_Select & "   and tr1.取引先区分 ="
    gSL_Select = gSL_Select & "    ("
    gSL_Select = gSL_Select & "select max(tt1.取引先区分) from 取引先マスタ tt1 where tt1.取引先CD = tr1.取引先CD"
    gSL_Select = gSL_Select & "    )"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ tr2"
    gSL_Select = gSL_Select & "    on tr2.取引先CD = jkt.実績元"
    gSL_Select = gSL_Select & "   and tr2.取引先区分 ="
    gSL_Select = gSL_Select & "    ("
    gSL_Select = gSL_Select & "select max(tt2.取引先区分) from 取引先マスタ tt2 where tt2.取引先CD = tr2.取引先CD"
    gSL_Select = gSL_Select & "    )"
'*--------------------------------------------------<< 2004/06/29 >> Delete Start
'   gSL_Select = gSL_Select & " where jkt.データ区分 = 0"
'*--------------------------------------------------<< 2004/06/29 >> Delete End
'   >> 選択条件セット
'*--------------------------------------------------<< 2004/12/29 >> Update Start
    Sl_Where = " where 納入番号 <> ''"
    If RTrim(iH1_Ymd) <> "" Then
        If CisFun.Val2(iH1_Ymd.cDay) = 0 Then
            Sl_Where = Sl_Where & "   and substring(jkt.実績日,1,6) = '" & iH1_Ymd & "'"
        Else
            Sl_Where = Sl_Where & "   and jkt.実績日 = '" & iH1_Ymd & "'"
        End If
    End If
    If H1Ck_JKbn(1) = 1 Or _
       H1Ck_JKbn(2) = 1 Or _
       H1Ck_JKbn(3) = 1 Then
        Sl_Where = Sl_Where & " and  jkt.実績区分 in ("
        gStr = ""
        If H1Ck_JKbn(1) = 1 Then Sl_Where = Sl_Where & gStr & "1": gStr = ","
        If H1Ck_JKbn(2) = 1 Then Sl_Where = Sl_Where & gStr & "2": gStr = ","
        If H1Ck_JKbn(3) = 1 Then Sl_Where = Sl_Where & gStr & "3": gStr = ","
        Sl_Where = Sl_Where & ")"
    End If
    gSL_Select = gSL_Select & Sl_Where
'*--------------------------------------------------<< 2004/12/29 >> Update End
'   >> 集計
    gSL_Select = gSL_Select & " group by"
    gSL_Select = gSL_Select & "  jkt.納入番号"
    gSL_Select = gSL_Select & ", jkt.実績区分"
    gSL_Select = gSL_Select & ", jkt.実績日"
    gSL_Select = gSL_Select & ", jkt.実績便"
    gSL_Select = gSL_Select & ", jkt.実績先"
    gSL_Select = gSL_Select & ", jkt.受入"
    gSL_Select = gSL_Select & ", jkt.実績元"
    gSL_Select = gSL_Select & ", tr1.略称"
    gSL_Select = gSL_Select & ", tr2.略称"
'   >> 表示順
    If Op_Sort(0) Then
        gSL_Select = gSL_Select & " order by"
        gSL_Select = gSL_Select & "  jkt.実績区分"
        gSL_Select = gSL_Select & ", jkt.実績日"
        gSL_Select = gSL_Select & ", jkt.実績便"
        gSL_Select = gSL_Select & ", jkt.実績先"
    Else
        gSL_Select = gSL_Select & " order by"
        gSL_Select = gSL_Select & "  jkt.実績区分 desc"
        gSL_Select = gSL_Select & ", jkt.実績日 desc"
        gSL_Select = gSL_Select & ", jkt.実績便 desc"
        gSL_Select = gSL_Select & ", jkt.実績先 desc"
    End If

    Adodc1.RecordSource = gSL_Select
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount = 0 Then
        With CisFun
            .MB_Lines = 3
            .MB_MSG(1) = "    該当するデータはありません  "
            .MB_Button = OK
            .MBOX
        End With
        Exit Function
    End If
    Call TotalDisp

    BodySet = True
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      グリッド処理(TDBGridPro)                             +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'+----------------------------------------------------------+
'+      ADODC
'+----------------------------------------------------------+
Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Adodc1.RecordSource <> "" Then Call TotalDisp
End Sub
'+----------------------------------------------------------+
'+      ADODC1(件数表示)                                     +
'+----------------------------------------------------------+
Private Sub TotalDisp()
    Static NowCnt       As Long
    Static MaxCnt       As Long
    NowCnt = Adodc1.Recordset.AbsolutePosition
    MaxCnt = Adodc1.Recordset.RecordCount
    Me.Caption = FormCap & Space(5)
    Me.Caption = Me.Caption & "抽出件数 = " & Format(NowCnt, "##,##0")
    Me.Caption = Me.Caption & "/" & Format(MaxCnt, "##,##0") & " 件"
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseUp  イベント   》                +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SendKeys "{RIGHT}"
    If DBGrid1.RowContaining(Y) < 0 Then Exit Sub
    If DBGrid1.ColContaining(X) < 0 Then Exit Sub
    DBGrid1.Row = DBGrid1.RowContaining(Y)
    DoEvents
    Call RtnSet
End Sub
'+-------------------------------------+
'+      選択内容ｾｯﾄ & 復帰ｽﾃｰﾀｽｾｯﾄ      +
'+-------------------------------------+
Private Sub RtnSet()
    With Adodc1
        RV_Code = .Recordset("納入番号")
    End With
    Adodc1.Recordset.Close
    RV_Rtn = True
    Call PB_END_Click
End Sub
