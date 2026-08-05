VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CNJ0510 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "かんばん増減リスト　発行"
   ClientHeight    =   11235
   ClientLeft      =   75
   ClientTop       =   1485
   ClientWidth     =   15360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin VB.PictureBox Csv_Bar 
      Height          =   345
      Left            =   4440
      ScaleHeight     =   285
      ScaleWidth      =   6555
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2460
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   4365
      Left            =   4410
      ScaleHeight     =   4480.719
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6600
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2850
      Width           =   6660
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   4200
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   7408
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin VB.CheckBox H1Chk_HKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "回収"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   3270
            TabIndex        =   3
            Top             =   1680
            Width           =   915
         End
         Begin VB.CheckBox H1Chk_HKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "追加"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   2040
            TabIndex        =   2
            Top             =   1680
            Width           =   915
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "　出　力　先"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   75
            TabIndex        =   12
            Top             =   3060
            Width           =   6345
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "プリンター"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   1065
               TabIndex        =   5
               Top             =   330
               Value           =   -1  'True
               Width           =   1260
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "画面(ﾌﾟﾚﾋﾞｭｰ)"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   2505
               TabIndex        =   6
               Top             =   330
               Width           =   1875
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "CSV"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   4380
               TabIndex        =   7
               Top             =   330
               Width           =   975
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   315
            Left            =   660
            Top             =   2220
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "仕入先"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
         End
         Begin CisYM_V60.CisYM iH1_Ymd 
            Height          =   375
            Left            =   1980
            TabIndex        =   1
            Top             =   1200
            Width           =   1440
            _ExtentX        =   2328
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
            cSize           =   -1  'True
            cChkResult      =   0   'False
         End
         Begin Cis3D_v60.CIS3D UC_3D28 
            Height          =   615
            Index           =   1
            Left            =   60
            Top             =   60
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   1085
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
            Begin VB.PictureBox B1_Area 
               BackColor       =   &H00FFC0FF&
               Height          =   555
               Left            =   45
               ScaleHeight     =   495
               ScaleWidth      =   6210
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   30
               Width           =   6270
               Begin Cis3D_v60.CIS3D lb_msg 
                  Height          =   405
                  Left            =   60
                  Top             =   60
                  Width           =   6090
                  _ExtentX        =   10742
                  _ExtentY        =   714
                  BackColor       =   16761087
                  Caption         =   "かんばん増減リスト　発行を行います"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   14.25
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cBoderColor1    =   16761087
                  cBoderColor2    =   16761087
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   6
                  cBoderStyle     =   2
               End
            End
         End
         Begin Cis3D_v60.CIS3D lb_PrtDev 
            Height          =   375
            Left            =   75
            Top             =   3750
            Width           =   5970
            _ExtentX        =   10530
            _ExtentY        =   661
            BackColor       =   16777152
            Caption         =   ""
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D PB_Print 
            Height          =   345
            Left            =   6075
            Top             =   3780
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   609
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cPositionX      =   30
            cPositionY      =   -90
         End
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   345
            Index           =   1
            Left            =   570
            Top             =   1200
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            ForeColor       =   16711680
            Caption         =   "対象年月"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   315
            Left            =   660
            Top             =   1680
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "発行区分"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
         End
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   1980
            TabIndex        =   4
            Top             =   2190
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   661
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
            Left            =   3345
            Top             =   2190
            Width           =   2535
            _ExtentX        =   4471
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D PB_Torcd 
            Height          =   360
            Left            =   3000
            Top             =   2190
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            ForeColor       =   65280
            Caption         =   "▼"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cButton         =   -1  'True
            cPositionX      =   20
         End
      End
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   585
      Left            =   5535
      Top             =   7740
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1032
      Caption         =   "終  了"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   4
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin VB.Timer Timer2 
      Left            =   945
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Left            =   510
      Top             =   840
   End
   Begin VB.TextBox Dummy 
      Height          =   375
      Left            =   105
      TabIndex        =   0
      Text            =   "Dummy"
      Top             =   855
      Width           =   345
   End
   Begin Cis3D_v60.CIS3D PB_Ent 
      Height          =   585
      Left            =   8400
      Top             =   7740
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1032
      Caption         =   "開  始"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   4
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3150
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 かんばん増減リスト 発行 】"
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
         Left            =   14370
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
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
         Height          =   285
         Left            =   11850
         Top             =   30
         Width           =   2490
         _ExtentX        =   4392
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
   End
   Begin Cis3D_v60.CIS3D PB_Ok 
      Height          =   585
      Left            =   6975
      Top             =   7740
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1032
      Caption         =   "確　認"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   4
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin VB.Label PNL_Ok 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6960
      TabIndex        =   15
      Top             =   8415
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8400
      TabIndex        =   10
      Top             =   8415
      Width           =   1485
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5520
      TabIndex        =   9
      Top             =   8415
      Width           =   1485
   End
   Begin VB.Label BG_Logo 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  '透明
      Caption         =   "SIMPR-A"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   8
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CNJ0510"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  かんばん増減リスト発行
'**       フォームID    :  CNJ0510
'**       処理概要      :
'**
'**       作  成  日    :  2004/03/16
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
'
    Dim mSaveMsg        As String       ' ﾒｯｾｰｼﾞ
    Dim mSavePrint      As String
    Dim mPrtDev         As String       ' ﾌﾟﾘﾝﾀ出力先
    Dim mCsvDir         As String
    Dim mCsvKey         As String
    Dim mCsvName        As String
'   背景色
    Dim mSelColor       As Long
    Dim mUnSelColor     As Long
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   # 二 重 起 動 防 止 #
'   #------------------#
    If Not Dupli_Start Then End
'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
    If Not IniGet Then End
    mSaveMsg = lb_msg
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
    Timer2.Interval = 1000   ' 時刻表示間隔
    Timer2.Enabled = True    ' 時刻表示開始

    With CisFrm
        Set .MSGObject = lb_msg
        Set .MSGTimer = Timer2
        .MSGStopCount = 10
        .MSGType = Left_To_Right
        .MSGText = mSaveMsg
        .MSGInterval = 100
        .MSGRelate
    End With

'年月初期表示
    Call GetNJProc_YM(3)
    iH1_Ymd = gn_HNYM
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

'   背景色の退避
    mSelColor = Op_Out(0).BackColor
    mUnSelColor = Op_Out(1).BackColor

    If gLogoDisp Then
       BG_Logo(0).Caption = gLogo
       CisFrm.BGLDisp BG_Logo, 10
    End If
    
    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

'CSVﾌｧｲﾙ出力先
    mCsvName = Mid(CisFun.WinRoot, 1, 3) & "かんばん増減リスト.csv"
    mCsvKey = "CNJ0510"
    mCsvDir = GetSetting(appname:=gRegKey, section:="CSV", Key:=mCsvKey, Default:=mCsvName)
'   ﾌﾟﾘﾝﾀ出力先
    Op_Out(0).Value = True
    mPrtDev = Printer.DeviceName
    lb_PrtDev = mPrtDev
    mSavePrint = mPrtDev

    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
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
    Key_Acc = False

    If Not FormAct Then Exit Function
    If UCase(IVENT) <> "DOWN" Then Exit Function

    Select Case Key_Code
        Case vbKeyF7:       If Dummy.Tag <> "" Then Call PB_Look_Click      '【検索】
        Case vbKeyF12
                            If PB_END.Visible Then Call PB_END_Click        '【終了】
                            If PB_Ok.Visible Then Call PB_END_Click         '【終了】
        Case vbKeyReturn
                            If PB_ENT.Visible Then                          '【入力】
   '                            Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                Call PB_ENT_Click
   '                            Dummy.Enabled = False
                            End If
        Case Else:          Exit Function
    End Select

    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      フォーム    《Ｐａｉｎｔ イベント》                   +
'+----------------------------------------------------------+
Private Sub Form_Paint()
'ｸﾞﾗﾃﾞｰｼｮﾝ表示
    With CisFrm
        .Gradation_Direction = Landscape
        .Gradation_Shading = Lighten
       '.Gradation 100, 160, 250, 4, 2, 3
        .Gradation 0, 0, 250, 4, 4, 3
    End With
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Call CisDB.DBDISConnect
    Call CisFrm.UnLoadDisp

    CisFrm.PrintSet mSavePrint

    End
End Sub
'+----------------------------+
'+    確　認  キ ー(F12)       +
'+----------------------------+
Private Sub PB_OK_Click()
    Call PB_END_Click
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+      検         索         +
'+----------------------------+
Private Sub PB_Look_Click()
    If Dummy.Tag = "iH1_Torcd" Then Call PB_Torcd_Click: Exit Sub
End Sub
'+---------------------+
'+    取引先検索        +
'+---------------------+
Private Sub PB_Torcd_Click()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 9
'
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    iH1_Torcd.SetFocus

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+----------------------------+
'+    iH1_Torcd(GotFocus)    +
'+----------------------------+
Private Sub iH1_Torcd_GotFocus()
    Dummy.Tag = ActiveControl.Name
End Sub
'+----------------------------+
'+    iH1_Torcd(LostFocus)    +
'+----------------------------+
Private Sub iH1_Torcd_LostFocus()
    Dummy.Tag = ""
    H1lb_Tornm = ""
    If RTrim(iH1_Torcd) <> "" Then
        If TorNmGet(iH1_Torcd, 9) Then
            H1lb_Tornm = TRM_RName
        End If
    End If
End Sub
'+----------------------------------------+
'+  Op_Out   出力先
'+----------------------------------------+
Private Sub Op_Out_GotFocus(Index As Integer)
    Op_Out(0).Tag = Index
    For Each gObj In Op_Out
        If gObj.Index = Index Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj
    Select Case Index
        Case 0: lb_PrtDev = mPrtDev
        Case 1: lb_PrtDev = mPrtDev
        Case 2: lb_PrtDev = mCsvDir
        Case Else
            lb_PrtDev = ""
    End Select
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*      メッセージ表示        *
'*******************************
Private Sub Timer2_Timer()
    CisFrm.MSGRelate
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
    Dim B2Mode      As Boolean

    H1Mode = False
    B2Mode = False
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1":   H1Mode = True
       Case "B2":   B2Mode = True
    End Select
    H1_Area1.Enabled = H1Mode

    PB_ENT.Visible = H1Mode
    PNL_ENT.Visible = H1Mode
    PB_END.Visible = H1Mode
    PNL_END.Visible = H1Mode
    Timer2.Enabled = H1Mode Or B2Mode
    PB_Ok.Visible = B2Mode
    PNL_Ok.Visible = B2Mode

    Csv_Bar.Visible = False

' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'*********************************
'*          入力処理(Enter)      *
'*********************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)

    If Not Head1Chk Then GoTo ReturnPress_Ed

    Call DispChange("B1")

    If Not PrintStart Then GoTo ReturnPress_Ee

    lb_msg.Caption = "【  出  力  終  了  】"
    DoEvents

    Call CisFrm.MousePT(1)
    With CisFrm
        .MSGType = FSize_Changes
        .MSGText = "【  出  力  終  了  】"
    End With
    ProcHB = "B2"
    Call DispChange(ProcHB)
    Exit Sub
ReturnPress_Ee:
    lb_msg = mSaveMsg
    DoEvents
    ProcHB = "H1"
    Call DispChange(ProcHB)
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False

'   対象年月
    If Trim(iH1_Ymd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    対象年月を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ymd.SetFocus
        Exit Function
    Else
        If Not iH1_Ymd.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    対象年月を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Ymd.SetFocus
            Exit Function
        End If
    End If
'   仕入先
    If Trim(iH1_Torcd) <> "" Then
        If Not TorNmGet(iH1_Torcd, 9) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    仕入先　未登録！（取引先マスタ）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
    End If

    lb_msg.Caption = "【 かんばん増減リスト 抽出中 】"
    DoEvents

    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  kkt.年月"
    gSL_Select = gSL_Select & ", kkt.仕入先"
    gSL_Select = gSL_Select & ", kkt.背番号"
    gSL_Select = gSL_Select & ", kkt.回転枚数"
    gSL_Select = gSL_Select & ", kkt.前月回転枚数"
    gSL_Select = gSL_Select & ", kkt.増発枚数"
    gSL_Select = gSL_Select & ", kkt.抜取枚数"
    gSL_Select = gSL_Select & ", kkt.増減区分"
    gSL_Select = gSL_Select & ", isnull(trm.略称, '') 仕入先名"
    gSL_Select = gSL_Select & ", isnull(hnm.表示品番, kkt.品番) 表示品番"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  かんばん回転枚数テーブル kkt"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = kkt.仕入先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm"
    gSL_Select = gSL_Select & "    on hnm.品番 = kkt.品番"
    gSL_Select = gSL_Select & " where kkt.年月 = '" & iH1_Ymd & "'"
'   >> 仕入先
    If RTrim(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & "   and kkt.仕入先 = '" & RTrim(iH1_Torcd) & "'"
    End If
'   >> 発行区分
    If H1Chk_HKbn(0).Value = 1 And H1Chk_HKbn(1).Value = 0 Then
    '   > 回収
        gSL_Select = gSL_Select & "   and kkt.増減区分 = 0"
    End If
    If H1Chk_HKbn(0).Value = 0 And H1Chk_HKbn(1).Value = 1 Then
    '   > 追加
        gSL_Select = gSL_Select & "   and kkt.増減区分 = 1"
    End If
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  kkt.年月"
    gSL_Select = gSL_Select & ", kkt.仕入先"
    gSL_Select = gSL_Select & ", kkt.増減区分 desc"
    gSL_Select = gSL_Select & ", kkt.増発枚数 desc"
    gSL_Select = gSL_Select & ", kkt.抜取枚数 desc"
    gSL_Select = gSL_Select & ", kkt.背番号"
    gSL_Select = gSL_Select & ", kkt.品番"

    If Not KKTRead(gSL_Select, , 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    該当データがありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If

    Head1Chk = True
End Function

'+-------------------------------+
'+     プリンタ・CSV出力先 変更    +
'+-------------------------------+
Private Sub PB_Print_Click()
    If Op_Out(2).Value Then
        ' CSV出力先変更
        On Error Resume Next
        CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly
        CommonDialog1.FileName = mCsvDir
        CommonDialog1.Filter = "ＣＳＶ ファイル (*.csv)|*.csv"
        ' 以下のｿｰｽｺｰﾄﾞは、<ｷｬﾝｾﾙ>ﾎﾞﾀﾝがｸﾘｯｸされたときに実行されます。
        CommonDialog1.CancelError = True
        CommonDialog1.ShowOpen
        If Err = 32755 Then
            '<ｷｬﾝｾﾙ>ﾎﾞﾀﾝが押されたときにｴﾗｰｺｰﾄﾞを返すようにする設定です。
            '必要なｿｰｽｺｰﾄﾞを入力してください。
        End If
        If Trim(CommonDialog1.FileName) <> "" Then
            mCsvDir = CommonDialog1.FileName
            lb_PrtDev.Caption = mCsvDir
        End If
        On Error GoTo 0
    Else
    'If Op_Out(0).Value Or Op_Out(1).Value Then
        ' プリンタ種類変更
         On Error Resume Next
         CommonDialog1.Flags = 64 + cdlPDHidePrintToFile
         CommonDialog1.Max = 0
         CommonDialog1.Min = 0
         CommonDialog1.Copies = 1
         CommonDialog1.ToPage = 1
         CommonDialog1.FromPage = 1
        ' 以下のｿｰｽｺｰﾄﾞは、<ｷｬﾝｾﾙ>ﾎﾞﾀﾝがｸﾘｯｸされたときに実行されます。
         CommonDialog1.CancelError = True
         CommonDialog1.ShowPrinter
         If Err = 32755 Then
        '<ｷｬﾝｾﾙ>ﾎﾞﾀﾝが押されたときにｴﾗｰｺｰﾄﾞを返すようにする設定です。
        '必要なｿｰｽｺｰﾄﾞを入力してください。
         Else
             mPrtDev = Printer.DeviceName
             lb_PrtDev.Caption = mPrtDev
         End If
         On Error GoTo 0
    End If
End Sub
'+---------------------------------------------+
'+      印刷処理                               +
'+---------------------------------------------+
Private Function PrintStart() As Boolean
    PrintStart = False
    If Op_Out(0).Value Or Op_Out(1).Value Then
        lb_msg.Caption = "【 かんばん増減リスト 印刷中 】"
        DoEvents
        With CisPrtCNJ0510
            .cDataMaxCnt = CisDB.RecordCount
            Set .iPrtForm = CNJ0510L
            .dpMSG = "かんばん増減リスト　印刷中"
            If Op_Out(0).Value Then
                .iPreview = pv_Print
            Else
                .iPreview = pv_Preview
            End If
            .PrintStart
        End With
    Else
        If Not CsvPut Then Exit Function
    End If
    PrintStart = True
End Function
'****************************
'*       CSV出力処理         *
'****************************
Private Function CsvPut() As Boolean
    Dim OutFNo      As Long
    Dim OutStr      As String
    Dim OutCnt      As Long

    CsvPut = False

    On Error GoTo CsvPut_Err

'   出力ﾌｧｲﾙ存在確認
    If Dir(mCsvDir) <> "" Then
        DoEvents
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "   　" & mCsvDir & "　ﾌｧｲﾙは既に存在します。            "
            .MB_MSG(3) = "   　" & "上書きしますか？            "
            .MB_Title = "上書き確認"
            .MB_Button = Yes_No
            If Not .MBOX Then Exit Function
        End With
    End If

    lb_msg.Caption = "【 かんばん増減リスト 出力中 】"
    Csv_Bar.Visible = True

    DoEvents

    Call CisFrm.MousePT(11)

    With CisFrm
        Set .MeterPic = Csv_Bar
        .MeterMax = CisDB.RecordCount
    End With

'   >> 出力ﾃｷｽﾄ Open
    OutFNo = FreeFile
    Open mCsvDir For Output As #OutFNo
'   >> Head部 編集＆出力
    Call CsvEdit_Head1(OutStr)
    Print #OutFNo, OutStr
    Call CsvEdit_Head2(OutStr)
    Print #OutFNo, OutStr
'   >> ｸﾘｱ
    OutCnt = 0
    Do Until Not KKT_RDSTS
    '   >> ｶｳﾝﾄ
        OutCnt = OutCnt + 1
    '   >> Body部 編集＆出力
        Call CsvEdit_Body1(OutStr, OutCnt)
        Print #OutFNo, OutStr

        CisFrm.MeterUpdate OutCnt

    '   >> 次のﾚｺｰﾄﾞを読み込む
        Call KKTReadNext(1)
    Loop
    Call KKTClose(1)
'   >> 出力ﾃｷｽﾄ Close
    Close #OutFNo

    SaveSetting appname:=gRegKey, section:="CSV", Key:=mCsvKey, setting:=RTrim(mCsvDir)

    Call CisFrm.MousePT(1)

    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "    【  CSV出力 正常終了 】            "
        .MB_MSG(3) = "    出 力 先  = " & RTrim(mCsvDir) & "　　　　　"
        .MB_MSG(4) = "    出力件数 = " & Format(OutCnt, "#,##0") & " 件"
        .MB_Title = "【確認】"
        .MB_Button = OK
        .MBOX
    End With

    CsvPut = True
    Exit Function
CsvPut_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------------+
'+    Head部　編集処理
'+----------------------------+
Private Sub CsvEdit_Head1(CsvStr As String)
    CsvStr = ""
    CsvStr = CsvStr & CisFun.Mid2(KKT.年月, 1, 4)
    CsvStr = CsvStr & "年"
    CsvStr = CsvStr & CisFun.Mid2(KKT.年月, 5, 2)
    CsvStr = CsvStr & "月度"
    CsvStr = CsvStr & "かんばん増減リスト"
End Sub
Private Sub CsvEdit_Head2(CsvStr As String)
    CsvStr = ""
    CsvStr = CsvStr & "№"
    CsvStr = CsvStr & ",仕入先"
    CsvStr = CsvStr & ",仕入先名"
    CsvStr = CsvStr & ",背番号"
    CsvStr = CsvStr & ",品番"
    CsvStr = CsvStr & ",当月枚数"
    CsvStr = CsvStr & ",前月枚数"
    CsvStr = CsvStr & ",追加枚数"
    CsvStr = CsvStr & ",回収枚数"
End Sub
'+----------------------------+
'+    Body部　編集処理
'+----------------------------+
Private Sub CsvEdit_Body1(CsvStr As String, CsvCnt As Long)
    With KKT
        CsvStr = ""
        CsvStr = CsvStr & CsvCnt
        CsvStr = CsvStr & "," & RTrim(.仕入先)
        CsvStr = CsvStr & "," & RTrim(.仕入先名)
        CsvStr = CsvStr & "," & RTrim(.背番号)
        CsvStr = CsvStr & "," & RTrim(.表示品番)
        CsvStr = CsvStr & "," & .回転枚数
        CsvStr = CsvStr & "," & .前月回転枚数
        Select Case .増減区分
            Case 1
                CsvStr = CsvStr & "," & .増発枚数
                CsvStr = CsvStr & ","
            Case 0
                CsvStr = CsvStr & ","
                CsvStr = CsvStr & "," & .抜取枚数
        End Select
    End With
End Sub

