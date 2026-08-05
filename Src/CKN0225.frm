VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Begin VB.Form CKN0225 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "かんばん発行"
   ClientHeight    =   10980
   ClientLeft      =   0
   ClientTop       =   345
   ClientWidth     =   15330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   15330
   Begin VB.PictureBox Picture1 
      Height          =   855
      Index           =   0
      Left            =   870
      ScaleHeight     =   795
      ScaleWidth      =   4485
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   720
      Width           =   4545
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   705
         Left            =   60
         Top             =   60
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   1244
         ForeColor       =   16711680
         Caption         =   "表示区分"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin VB.OptionButton H1Op_DispKB 
            BackColor       =   &H00C0C0C0&
            Caption         =   "発行済"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   3210
            TabIndex        =   2
            Top             =   240
            Width           =   1125
         End
         Begin VB.OptionButton H1Op_DispKB 
            BackColor       =   &H00C0C0C0&
            Caption         =   "未発行"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   2040
            TabIndex        =   1
            Top             =   240
            Width           =   1125
         End
         Begin VB.OptionButton H1Op_DispKB 
            BackColor       =   &H00C0FFC0&
            Caption         =   "全て"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1110
            TabIndex        =   0
            Top             =   240
            Value           =   -1  'True
            Width           =   945
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   450
      Top             =   780
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      Left            =   150
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   870
      Visible         =   0   'False
      Width           =   225
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
      Left            =   11430
      ScaleHeight     =   660
      ScaleWidth      =   2655
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   660
      Width           =   2715
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   915
         Top             =   30
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
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   1755
         Top             =   30
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
      Begin CisBtn_60.CisBtn PB_Prt 
         Height          =   585
         Left            =   60
         Top             =   30
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
         BtnType         =   13
         BtnCaption      =   "印刷"
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
         PFCaption       =   "( F8 )"
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
      Height          =   8595
      Left            =   870
      ScaleHeight     =   8535
      ScaleWidth      =   13245
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2010
      Width           =   13305
      Begin vsFlexLib.vsFlexArray vsGrid1 
         Height          =   8310
         Left            =   180
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   12975
         _Version        =   196608
         _ExtentX        =   22886
         _ExtentY        =   14658
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
         FormatString    =   "№ |仕入先 |仕入先名　　　　　　|受入|品　番　　　　　　　　　|背番号|収容器   |収容数  |発行連番|発行|編集前品番|Select|"
         Rows            =   26
         Cols            =   13
         BackColor       =   12648447
         BackColorBkg    =   12632256
         RowHeightMin    =   320
         GridLines       =   2
      End
   End
   Begin Cis3D_v60.CIS3D Head 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 発注かんばん　発行品番参照 】"
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14130
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
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
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   300
         Left            =   11625
         Top             =   30
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   529
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   0
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
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   495
      Index           =   1
      Left            =   5730
      ScaleHeight     =   495
      ScaleWidth      =   8460
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1530
      Width           =   8460
      Begin Cis3D_v60.CIS3D UC_3D6 
         Height          =   435
         Left            =   30
         Top             =   30
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   767
         ForeColor       =   0
         Caption         =   "発行総枚数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D iT1_Addsu 
            Height          =   315
            Left            =   1200
            Top             =   60
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            BackColor       =   16761024
            ForeColor       =   0
            Caption         =   "ZZ,ZZ9"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   435
         Left            =   2130
         Top             =   30
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   767
         ForeColor       =   0
         Caption         =   "発行済枚数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D iT1_Sumisu 
            Height          =   315
            Left            =   1200
            Top             =   60
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            BackColor       =   16777152
            ForeColor       =   0
            Caption         =   "ZZ,ZZ9"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   435
         Left            =   4230
         Top             =   30
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   767
         ForeColor       =   0
         Caption         =   "未発行枚数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D iT1_HSu 
            Height          =   315
            Left            =   1200
            Top             =   60
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            BackColor       =   12648384
            ForeColor       =   0
            Caption         =   "ZZ,ZZ9"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   435
         Left            =   6330
         Top             =   30
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   767
         BackColor       =   16761087
         ForeColor       =   0
         Caption         =   "発行指定数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D iT1_SSu 
            Height          =   315
            Left            =   1200
            Top             =   60
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            BackColor       =   0
            ForeColor       =   65535
            Caption         =   "ZZ,ZZ9"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cFont3D         =   2
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
   End
End
Attribute VB_Name = "CKN0225"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   かんばん発行品番参照
'**       フォームID    :   CKN0215
'**       処理概要      :
'**
'**       作  成  日    :   2003/12/24  By CIS SIMPR-A
'**       変  更  日    :   2008/04/11  By CIS 背番号８桁対応
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim CisVsGrid       As New CisVsGrid3
    
    Dim mGotOPColor     As Long
    Dim mLostOPColor    As Long
    Dim mGotGDColor     As Long
    Dim mLostGDColor    As Long
    Dim mLostFXColor    As Long
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
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
    
    mGotOPColor = H1Op_DispKB(0).BackColor
    mLostOPColor = H1Op_DispKB(1).BackColor
    
    mGotGDColor = iT1_Addsu.BackColor
    mLostGDColor = vsGrid1.BackColor
    mLostFXColor = vsGrid1.GridColorFixed
    
    Select Case kn_PrtKB
        Case 1: Head = "【 工程かんばん　発行品番参照 】"
        Case Else
                Head = "【 発注かんばん　発行品番参照 】"
    End Select
    
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 10
        .InitGet = False
        .Init
    End With
    Call GridInit
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    FormAct = Not FormAct
'
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
           Case vbKeyF8:       If PB_Prt.Visible Then Call PB_Prt_Click    '【印刷】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
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
    Call PB_END_Click
End Sub

'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Unload CKN0225
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Call GridColorSet
End Sub
'+----------------------------+
'+    印　刷  キ ー(F8)        +
'+----------------------------+
Private Sub PB_Prt_Click()
    If CisFun.Val2(iT1_SSu) = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　かんばん発行選択をして下さい       "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Sub
    End If
    
    gSL_Select = "Select * From かんばん発行テーブル "
    gSL_Select = gSL_Select & " Where かんばん種類 = " & kn_PrtKB
    gSL_Select = gSL_Select & "   And 発行FLG = 'X'"
    gSL_Select = gSL_Select & " Order By 発行種類,取引先,受入,品番,背番号,収容区分,発行NO"
    If Not KHTRead(gSL_Select, , 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　かんばん発行データがありません       "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Sub
    End If

    Kn_PCnt = CisDB.RecordCount

    CKN0240.Show vbModal
    Unload CKN0240
    Set CKN0240 = Nothing
    
    Call GridSet
    
    vsGrid1.SetFocus
    
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'+---------------------+
'+    表示区分          +
'+---------------------+
Private Sub H1Op_DispKB_Click(Index As Integer)
    For Each gObj In H1Op_DispKB
        If gObj.Index = Index Then
           gObj.BackColor = mGotOPColor
        Else
           gObj.BackColor = mLostOPColor
        End If
    Next gObj
    
    Call GridSet
    vsGrid1.SetFocus
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
        Call .FixedSet(中中, 右中, なし, 6, 0, "印刷頁")
        Select Case kn_PrtKB
            Case 1
                Call .FixedSet(中中, 左中, あり, 7, 1, "手配先")
                Call .FixedSet(中中, 左中, あり, 20, 2, "手配先名(社内)")
            Case Else
                Call .FixedSet(中中, 左中, あり, 7, 1, "仕入先")
                Call .FixedSet(中中, 左中, あり, 20, 2, "仕入先名")
        End Select
        Call .FixedSet(中中, 左中, あり, 2, 3, "受")
        Call .FixedSet(中中, 左中, あり, 25, 4, "品　番")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
        'Call .FixedSet(中中, 左中, あり, 6, 5, "背番号")
        Call .FixedSet(中中, 左中, あり, 8, 5, "背番号")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
        Call .FixedSet(中中, 左中, あり, 10, 6, "収容器")
        Call .FixedSet(中中, 右中, あり, 7, 7, "収容数")
        Call .FixedSet(中中, 中中, なし, 0, 8, "発行連番")
        Call .FixedSet(中中, 中中, あり, 0, 9, "発行")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 10, "編集前品番")
        Call .FixedSet(中中, 左中, なし, 0, 11, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9")
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
    Dim Ttl_SumiSu  As Long
    Dim Ttl_HSu     As Long

    CisVsGrid.Clear
    iT1_Addsu = ""
    iT1_Sumisu = ""
    iT1_HSu = ""
    iT1_SSu = ""
    With CisDB
        .SQL = "Update かんばん発行テーブル Set 発行FLG = '' "
        .SQL = .SQL & " Where かんばん種類 = " & kn_PrtKB
        .DBExec
    End With
' 抽出処理 初期ｾｯﾄ
    gSL_Select = "select KT.取引先,KT.受入,KT.表示品番,KT.背番号,KT.収容器,KT.品番,"
    gSL_Select = gSL_Select & "KT.収容数,KT.発行NO,KT.発行済,"
    gSL_Select = gSL_Select & "    TR.略称 取引先名1"
    gSL_Select = gSL_Select & "       from かんばん発行テーブル KT"
    gSL_Select = gSL_Select & "    Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "    ON TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "    AND TR.取引先CD = KT.取引先"
    gSL_Select = gSL_Select & "    Where KT.かんばん種類 = " & kn_PrtKB
    If H1Op_DispKB(1).Value Then
       gSL_Select = gSL_Select & "   And KT.発行済 = 0"
    End If
    If H1Op_DispKB(2).Value Then
       gSL_Select = gSL_Select & "   And KT.発行済 = 1"
    End If
    gSL_Select = gSL_Select & "    Order by KT.発行種類,KT.取引先,KT.受入,KT.品番,KT.背番号,KT.収容区分,KT.発行NO"
    If Not KHTRead(gSL_Select) Then GoTo GridSet_Ed
    
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        If CisDB.RecordCount + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = CisDB.RecordCount + 1
        End If
    End With
    ID = 0: Ttl_Kensu = 0
    Do Until Not KHT_RDSTS
        With vsGrid1
            ID = ID + 1
            '№
            .TextMatrix(ID, 0) = Format(ID, "####")
            .TextMatrix(ID, 1) = KHT.取引先
            .TextMatrix(ID, 2) = KHT.取引先名1
            .TextMatrix(ID, 3) = KHT.受入 & "     " & KHT.取引先 & KHT.品番
            .TextMatrix(ID, 4) = KHT.表示品番
            .TextMatrix(ID, 5) = KHT.背番号 & "    " & KHT.取引先 & KHT.品番
            .TextMatrix(ID, 6) = KHT.収容器
            .TextMatrix(ID, 7) = Format(KHT.収容数, "#,###")
            .TextMatrix(ID, 8) = Format(KHT.発行NO, "000")
            If KHT.発行済 = 1 Then
               .TextMatrix(ID, 9) = "済"
               Ttl_SumiSu = Ttl_SumiSu + 1
            Else
               Ttl_HSu = Ttl_HSu + 1
            End If
            .TextMatrix(ID, 10) = KHT.品番
            Ttl_Kensu = Ttl_Kensu + 1
       End With
       Call KHTReadNext
    Loop
    With vsGrid1
         iT1_Addsu = Format(Ttl_Kensu, "#,###")
         iT1_Sumisu = Format(Ttl_SumiSu, "#,###")
         iT1_HSu = Format(Ttl_HSu, "#,###")
         .Row = 1: .Col = 8
         .Redraw = True
         .MergeCells = flexMergeRestrictColumns
    End With
    
GridSet_Ed:
    Call KHTClose
End Sub
Private Sub vsGrid1_Click()
    If vsGrid1.MouseCol = 0 Or vsGrid1.MouseRow = 0 Then
        vsGrid1.Col = vsGrid1.MouseCol
        vsGrid1.Row = vsGrid1.MouseRow
    End If
    Call GridColorSet
End Sub
Private Sub GridColorSet()
    Dim wCol        As Long
    Dim wRow        As Long
    Dim wKey        As String
    
    With vsGrid1
        .Redraw = False
        wCol = .Col
        wRow = .Row
        
        If wCol = 0 And wRow = 0 Then
           For gLong = 1 To .Rows - 1
               If Trim(.TextMatrix(gLong, 1)) = "" Then Exit For
               Call ColorSet(gLong)
           Next gLong
           wCol = 1
           wRow = 1
           GoTo GridColorSet_Ed
        End If
        If wCol = 0 Or wCol = 8 Then
           For gLong = .Row To .RowSel
               If Trim(.TextMatrix(gLong, 1)) = "" Then Exit For
               Call ColorSet(gLong, 1)
           Next gLong
           GoTo GridColorSet_Ed
        End If
        
        wKey = .TextMatrix(.Row, .Col)
        For gLong = .Row To 1 Step -1
            If Trim(.TextMatrix(gLong, 1)) = "" Then Exit For
            If .TextMatrix(gLong, wCol) <> wKey Then Exit For
            Call ColorSet(gLong)
        Next gLong
        For gLong = wRow + 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, 1)) = "" Then Exit For
            If .TextMatrix(gLong, wCol) <> wKey Then Exit For
            Call ColorSet(gLong)
        Next gLong
        
    End With
GridColorSet_Ed:
    With vsGrid1
        .Row = wRow: .Col = wCol
        .Redraw = True
    End With
End Sub
Private Sub ColorSet(cRow As Long, Optional cColSel As Integer = 0)
    Dim wColFlg         As Boolean
    
    With vsGrid1
        .Row = cRow
        For gInt = 0 To .Cols - 1
            .Col = gInt
            
            wColFlg = False
            If cColSel = 0 Then
                wColFlg = True
            Else
                Select Case gInt
                    Case 0, 8, 9
                        wColFlg = True
                End Select
            End If
            If wColFlg Then
                If .TextMatrix(cRow, 11) = "X" Then
                   If gInt = 0 Then
                      .CellBackColor = mLostFXColor
                   Else
                      .CellBackColor = mLostGDColor
                   End If
                Else
                   .CellBackColor = mGotGDColor
                End If
            Else
                If .TextMatrix(cRow, 11) = "X" Then
                   If gInt = 0 Then
                      .CellBackColor = mLostFXColor
                   Else
                      .CellBackColor = mLostGDColor
                   End If
                End If
            End If
        Next gInt
                
        If .TextMatrix(cRow, 11) = "X" Then
           .TextMatrix(cRow, 11) = ""
           iT1_SSu = Format(CisFun.Val2(iT1_SSu) - 1, "#,###")
        Else
           .TextMatrix(cRow, 11) = "X"
           iT1_SSu = Format(CisFun.Val2(iT1_SSu) + 1, "#,###")
        End If
        With CisDB
            .SQL = "Update かんばん発行テーブル Set 発行FLG = '" & vsGrid1.TextMatrix(cRow, 11) & "'"
            .SQL = .SQL & " Where かんばん種類 = " & kn_PrtKB
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
            '.SQL = .SQL & " And   背番号 = '" & RTrim(Mid(vsGrid1.TextMatrix(cRow, 5), 1, 5)) & "'"
            .SQL = .SQL & " And   背番号 = '" & RTrim(Mid(vsGrid1.TextMatrix(cRow, 5), 1, 8)) & "'"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
            .SQL = .SQL & " And   品番 = '" & RTrim(vsGrid1.TextMatrix(cRow, 10)) & "'"
            .SQL = .SQL & " And   取引先 = '" & RTrim(vsGrid1.TextMatrix(cRow, 1)) & "'"
            .SQL = .SQL & " And   受入 = '" & RTrim(Mid(vsGrid1.TextMatrix(cRow, 3), 1, 2)) & "'"
            .SQL = .SQL & " And   発行NO = " & CisFun.Val2(vsGrid1.TextMatrix(cRow, 8))
            .DBExec
        End With
    
    End With
End Sub

Private Sub vsGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeySpace Then Exit Sub
    
    Call GridColorSet
    
End Sub
