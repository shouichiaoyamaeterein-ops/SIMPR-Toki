VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form CKK0079 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   11445
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   11520
      Top             =   210
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
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Dummy 
      Height          =   285
      Left            =   11400
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   780
      Width           =   150
   End
   Begin VB.PictureBox Back_Sort 
      BackColor       =   &H00808000&
      Height          =   675
      Left            =   6210
      ScaleHeight     =   615
      ScaleWidth      =   1710
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   60
      Width           =   1770
      Begin VB.OptionButton Op_Sort 
         BackColor       =   &H00808000&
         Caption         =   "昇順"
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   0
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   105
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton Op_Sort 
         BackColor       =   &H00808000&
         Caption         =   "降順"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   870
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00404040&
      Height          =   675
      Left            =   8010
      ScaleHeight     =   615
      ScaleWidth      =   3330
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   60
      Width           =   3390
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   1665
         Top             =   0
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
         Left            =   2490
         Top             =   0
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
         Left            =   840
         Top             =   0
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
         Left            =   15
         Top             =   0
         Visible         =   0   'False
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
      Height          =   4245
      Left            =   60
      ScaleHeight     =   4185
      ScaleWidth      =   11265
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1860
      Width           =   11325
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CKK0079.frx":0000
         Height          =   4050
         Left            =   60
         OleObjectBlob   =   "CKK0079.frx":0015
         TabIndex        =   7
         Top             =   60
         Width           =   11145
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1035
      Left            =   60
      ScaleHeight     =   975
      ScaleWidth      =   11280
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   780
      Width           =   11340
      Begin Cis3D_v60.CIS3D UC_3D 
         Height          =   855
         Index           =   0
         Left            =   45
         Top             =   60
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   1508
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
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   795
            Index           =   5
            Left            =   30
            Top             =   30
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "材　　　質"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
            cPositionY      =   30
            Begin CisText_V60.CisText iH1_Zai 
               Height          =   375
               Left            =   195
               TabIndex        =   1
               Top             =   330
               Width           =   3210
               _ExtentX        =   5662
               _ExtentY        =   661
               cWildeCard      =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
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
               cDataReplace    =   1
               cFaZero         =   0
               cFbComma        =   0
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   795
            Left            =   3585
            Top             =   30
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "取　引　先"
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
            cPositionY      =   30
            Begin CisText_V60.CisText iH1_Torcd 
               Height          =   360
               Left            =   150
               TabIndex        =   2
               Top             =   330
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   635
               cWildeCard      =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXX"
               MaxLength       =   7
            End
            Begin Cis3D_v60.CIS3D H1lb_Tornm 
               Height          =   360
               Left            =   1125
               Top             =   315
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   635
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hinbn 
            Height          =   795
            Left            =   7290
            Top             =   30
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "品　　番                  "
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
            cPositionY      =   60
            Begin VB.OptionButton H1Op_Hinku 
               BackColor       =   &H00C0C0C0&
               Caption         =   "受給品番"
               Height          =   240
               Index           =   1
               Left            =   2535
               TabIndex        =   4
               Top             =   75
               Width           =   1020
            End
            Begin VB.OptionButton H1Op_Hinku 
               BackColor       =   &H00C0FFC0&
               Caption         =   "発注品番"
               Height          =   240
               Index           =   0
               Left            =   1485
               TabIndex        =   3
               Top             =   75
               Value           =   -1  'True
               Width           =   1020
            End
            Begin CisText_V60.CisText iH1_Hinbn 
               Height          =   375
               Left            =   360
               TabIndex        =   5
               Top             =   330
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   661
               cWildeCard      =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   25
               cAutoEject      =   0   'False
            End
         End
      End
   End
   Begin Cis3D_v60.CIS3D HLB_Title 
      Height          =   675
      Left            =   90
      Top             =   60
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   1191
      BackColor       =   12648384
      Caption         =   "材料取引先検索(購入)"
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
End
Attribute VB_Name = "CKK0079"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   材料取引先マスタ検索
'**       フォームID    :   CKK0079
'**       処理概要      :
'**
'**       作  成  日    :   2004/02/18  By CIS
'**       変  更  日    :   2008/03/01  By CIS 材料取引先対応
'**       変  更  日    :   2008/03/07  By CIS 材料購入単価マスタの材料契約先追加
'**       変  更  日    :   2014/02/04  By CIS WIN7対応
'**
'**********************************************************'

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim mSL_Order       As String       ' ソート順(Order BY)
    Dim mSL_OrderA      As String       ' ソート順(Order BY)昇順
    Dim mSL_OrderD      As String       ' ソート順(Order BY)降順
    Dim mSL_SelectW     As String       ' SQL SELECT(Where)条件文
    Dim FormCap         As String
    
    Dim CisFrm          As CISFormContorl
    Dim CisComboTehai   As CisCombo
    Dim DBSpCols        As Columns      ' Columnsｵﾌﾞｼﾞｪｸﾄﾜｰｸ1
    Dim DBSpCol         As Column       ' Columnｵﾌﾞｼﾞｪｸﾄﾜｰｸ1
    Dim mColMoveFlg     As Boolean
    Dim SortCnt         As Long
    Dim mOrder()        As String
    
    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
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
    
    mGotColor = H1Op_Hinku(0).BackColor
    mLostColor = H1Op_Hinku(1).BackColor
    
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
' ADOｺﾝﾄﾛｰﾙ　接続情報　セット
    Adodc1.ConnectionString = CisDB.ConnectString
    '-----2014/02/04 EDIT START-----
    gStr = CisDB.ConnectString
    gConnect = "server=" & CisDB.Server & ";database=" & CisDB.DBName & ";User ID=" & CisDB.User & ";password=" & CisDB.PassWord(GetSetting(gRegKey, "Environment", "DrowsSap", ""))
    
    gStr = Replace(gStr, "Use Procedure", "Extended Properties=" & gConnect & ";Use Procedure")
    
    Adodc1.ConnectionString = gStr
    '-----2014/02/04 EDIT END-------
    Call HeadBodyClear("H")
    RV_Rtn = False
    
    H1Op_Hinku(0).Value = True
        
    If RV_Call <> "" Then
       Call ReturnPress
    End If
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
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    Select Case Key_Code
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click   '【検索】
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
    Call PB_END_Click
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Unload CKK0079
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
        PB_ENT.BtnCaption = "開始"
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
End Sub
'+----------------------------+
'+       仕入先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 0
    RV_Call = ""
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+-----------------------------+
'+  iH1_Torcd
'+-----------------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_Tornm = ""
    If TorNmGet(iH1_Torcd, 9) Then
       H1lb_Tornm = TRM_RName
    End If
End Sub

'+--------------------+
'+     H1Op_Hinku     +
'+--------------------+
Private Sub H1Op_Hinku_Click(Index As Integer)
    H1Op_Hinku(0).BackColor = mLostColor
    H1Op_Hinku(1).BackColor = mLostColor
    H1Op_Hinku(Index).BackColor = mGotColor
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
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1"
          H1Mode = True
          H1Color = gPGotSel
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
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
       If Not BodySet("") Then GoTo ReturnPress_Ed
       ProcHB = "B1"
       PB_ENT.BtnCaption = "選択"
       Call DispChange(ProcHB)
       GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then Call RtnSet
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      分類切替
'****************************
Private Sub Op_Sort_Click(Index As Integer)
    If ProcHB = "H1" Then Exit Sub
    If Adodc1.Recordset.RecordCount <> 0 Then Call BodySet("S")
    SendKeys "{TAB}"
End Sub
'****************************
'*      ボディ内容セット
'****************************
Private Function BodySet(SetType As String) As Boolean
    
    BodySet = False
    
    If FormAct Then
       DoEvents
    End If
    
    gSL_Select = ""
    gSL_Select = gSL_Select & " Select ZT.背番号, ZT.取引先, ZM.材質, ZT.材料管理番号," & vbCrLf
    gSL_Select = gSL_Select & "        ZM.板厚, ZM.幅, ZM.長さ, ISNULL(ZT.発注品番,'') 発注品番, ISNULL(ZT.受給品番,'') 受給品番," & vbCrLf
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(TR.略称,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "             Then ISNULL(ZT.取引先,'') + '：' + TR.略称" & vbCrLf
    gSL_Select = gSL_Select & "             Else ISNULL(ZT.取引先,'') End 表示取引先," & vbCrLf
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(TR1.略称,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "             Then ISNULL(ZK.契約先,'') + '：' + TR1.略称" & vbCrLf
    gSL_Select = gSL_Select & "             Else ISNULL(ZK.契約先,'') End 表示契約先," & vbCrLf
    gSL_Select = gSL_Select & "        ZT.材質 + ' ' +" & vbCrLf
    gSL_Select = gSL_Select & "        CAST(ZM.板厚 AS VARCHAR) + 'X' +" & vbCrLf
    gSL_Select = gSL_Select & "        CAST(ZM.幅   AS VARCHAR) + 'X' +" & vbCrLf
    gSL_Select = gSL_Select & "        CAST(ZM.長さ AS VARCHAR) 材料," & vbCrLf
    gSL_Select = gSL_Select & "        ISNULL(ZK.契約先, '') 契約先"
    gSL_Select = gSL_Select & "   From 材料取引先マスタ ZT" & vbCrLf
    gSL_Select = gSL_Select & "        Left Outer join 材料マスタ ZM" & vbCrLf
    gSL_Select = gSL_Select & "             On ZM.材料管理番号 = ZT.材料管理番号" & vbCrLf
    gSL_Select = gSL_Select & "        Left Outer Join 取引先マスタ TR" & vbCrLf
    gSL_Select = gSL_Select & "             ON TR.取引先CD = ZT.取引先" & vbCrLf
    gSL_Select = gSL_Select & "            AND TR.取引先区分 = 1" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料購入単価マスタ ZK" & vbCrLf
    gSL_Select = gSL_Select & "             ON ZK.材料管理番号 = ZT.材料管理番号" & vbCrLf
    gSL_Select = gSL_Select & "            AND ZK.取引先 = ZT.取引先" & vbCrLf
    gSL_Select = gSL_Select & "            AND ZK.履歴NO = ( SELECT MAX(履歴NO)" & vbCrLf
    gSL_Select = gSL_Select & "                                FROM 材料購入単価マスタ ZK1" & vbCrLf
    gSL_Select = gSL_Select & "                               WHERE ZK1.材料管理番号 = ZK.材料管理番号" & vbCrLf
    gSL_Select = gSL_Select & "                                 AND ZK1.取引先 = ZK.取引先" & vbCrLf
    gSL_Select = gSL_Select & "                                 AND ZK1.契約先 = ZK.契約先" & vbCrLf
    If RTrim(RV_SYmd) <> "" Then
       gSL_Select = gSL_Select & "                                 AND ZK1.適用開始日付 >= '" & RV_SYmd & "'" & vbCrLf
    End If
    If RTrim(RV_EYmd) <> "" Then
       gSL_Select = gSL_Select & "                                 AND ZK1.適用終了日付 <= '" & RV_EYmd & "'" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                            )"
    If RTrim(RV_SYmd) <> "" Then
       gSL_Select = gSL_Select & "            AND ZK.適用開始日付 >= '" & RV_SYmd & "'" & vbCrLf
    End If
    If RTrim(RV_EYmd) <> "" Then
       gSL_Select = gSL_Select & "            AND ZK.適用終了日付 <= '" & RV_EYmd & "'" & vbCrLf
    End If
    gSL_Select = gSL_Select & "        Left Outer Join 取引先マスタ TR1"
    gSL_Select = gSL_Select & "             ON TR1.取引先CD = ZK.契約先"
    gSL_Select = gSL_Select & "            AND TR1.取引先区分 = 1" & vbCrLf

    mSL_OrderA = " Order By ZM.材質, ZM.板厚, ZM.幅, ZM.長さ, ZT.取引先, ISNULL(ZT.発注品番,''), ISNULL(ZT.受給品番,'')"
    mSL_OrderD = " Order By ZM.材質 DESC, ZM.板厚 DESC, ZM.幅 DESC, ZM.長さ DESC, ZT.取引先 DESC, ISNULL(ZT.発注品番,'') DESC, ISNULl(ZT.受給品番,'') DESC"
' 選択条件セット
    mSL_SelectW = ""
    With CisFun
        .Where_CreateKB = NewDocument
        .Where_Create AlphaMe, "ISNULL(ZM.材質,'')", iH1_Zai
        .Where_Create AlphaMe, "ISNULL(ZT.取引先,'')", iH1_Torcd
        If H1Op_Hinku(0).Value Then
            .Where_Create AlphaMe, "ISNULL(ZT.発注品番,'')", iH1_Hinbn
        Else
            .Where_Create AlphaMe, "ISNULL(ZT.受給品番,'')", iH1_Hinbn
        End If
'        .Where_Create AlphaMe, "ZT.品目", 1
        
        mSL_SelectW = .Where_Phrase
    End With
    
    If Op_Sort(0).Value Then
       Adodc1.RecordSource = gSL_Select & mSL_SelectW & mSL_OrderA
    Else
       Adodc1.RecordSource = gSL_Select & mSL_SelectW & mSL_OrderD
    End If
    
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount = 0 Then
        DoEvents
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
    If DBGrid1.RowContaining(Y) < 0 Then Exit Sub
    If DBGrid1.ColContaining(X) < 0 Then Exit Sub
    SendKeys "{RIGHT}"
    DBGrid1.Row = DBGrid1.RowContaining(Y)
    DoEvents
    Call RtnSet
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 ColMove イベント   》            +
'+----------------------------------------------------------+
Private Sub DBGrid1_ColMove(ByVal Position As Integer, Cancel As Integer)
    mColMoveFlg = True
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 Paint イベント   》              +
'+----------------------------------------------------------+
Private Sub DBGrid1_Paint()
    If Not FormAct Then Exit Sub
    If mColMoveFlg Then  ' 再表示(分類)
        mColMoveFlg = False
        Call BodySet("S")
        SendKeys "{RIGHT}"
    End If
End Sub
'+-------------------------------------+
'+      選択内容ｾｯﾄ & 復帰ｽﾃｰﾀｽｾｯﾄ      +
'+-------------------------------------+
Private Sub RtnSet()
    With Adodc1
        RV_Zaist = .Recordset("材質")
        RV_TorcdK = .Recordset("取引先")
        RV_Itatu = .Recordset("板厚")
        RV_Width = .Recordset("幅")
        RV_Long = .Recordset("長さ")
        RV_JHinbn = .Recordset("受給品番")
        RV_HHinbn = .Recordset("発注品番")
        Rv_KanriNo = .Recordset("材料管理番号")
        RV_KTorcd = .Recordset("契約先")
    End With
    Adodc1.Recordset.Close
    RV_Rtn = True
    Call PB_END_Click
End Sub

