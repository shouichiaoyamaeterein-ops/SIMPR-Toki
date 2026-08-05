VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CNJ3110 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "内示表　発行"
   ClientHeight    =   11235
   ClientLeft      =   1815
   ClientTop       =   2220
   ClientWidth     =   15345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15345
   Begin VB.PictureBox Csv_Bar 
      Height          =   345
      Left            =   4275
      ScaleHeight     =   285
      ScaleWidth      =   6750
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2805
      Visible         =   0   'False
      Width           =   6810
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1605
      Top             =   4590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   585
      Left            =   5625
      Top             =   8520
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
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   4470
      Left            =   4275
      ScaleHeight     =   4590.005
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6750
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3210
      Width           =   6810
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   4140
         Index           =   1
         Left            =   135
         Top             =   135
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   7303
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
         Begin VB.PictureBox B1_Area 
            BackColor       =   &H00FFC0FF&
            Height          =   555
            Left            =   45
            ScaleHeight     =   495
            ScaleWidth      =   6330
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   45
            Width           =   6390
            Begin Cis3D_v60.CIS3D lb_msg 
               Height          =   405
               Left            =   15
               Top             =   60
               Width           =   6300
               _ExtentX        =   11113
               _ExtentY        =   714
               BackColor       =   16761087
               Caption         =   "内示表　発行を行います"
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
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   315
            Left            =   660
            Top             =   1170
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "契約先"
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
         Begin VB.OptionButton Op_OutKB 
            BackColor       =   &H00C0C0C0&
            Caption         =   "全て＜計画変更含む＞"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   2010
            TabIndex        =   6
            Top             =   2760
            Width           =   2820
         End
         Begin VB.OptionButton Op_OutKB 
            BackColor       =   &H00C0C0C0&
            Caption         =   "計画変更＜最新変更分＞"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   2010
            TabIndex        =   5
            Top             =   2430
            Width           =   2820
         End
         Begin VB.OptionButton Op_OutKB 
            BackColor       =   &H00C0C0C0&
            Caption         =   "定常＜初回＞"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   2010
            TabIndex        =   4
            Top             =   2100
            Value           =   -1  'True
            Width           =   2820
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   315
            Left            =   660
            Top             =   2100
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "出力区分"
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
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   375
            Left            =   1980
            TabIndex        =   1
            Top             =   720
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
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "　出　力　先"
            Height          =   600
            Left            =   75
            TabIndex        =   14
            Top             =   3060
            Width           =   6345
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "CSV"
               Height          =   240
               Index           =   2
               Left            =   4380
               TabIndex        =   9
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "画面(ﾌﾟﾚﾋﾞｭｰ)"
               Height          =   240
               Index           =   1
               Left            =   2505
               TabIndex        =   8
               Top             =   270
               Width           =   1875
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "プリンター"
               Height          =   240
               Index           =   0
               Left            =   1065
               TabIndex        =   7
               Top             =   270
               Value           =   -1  'True
               Width           =   1260
            End
         End
         Begin Cis3D_v60.CIS3D lb_PrtDev 
            Height          =   375
            Left            =   75
            Top             =   3690
            Width           =   6000
            _ExtentX        =   10583
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
            Left            =   6105
            Top             =   3720
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
            Height          =   330
            Index           =   1
            Left            =   570
            Top             =   720
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
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
            Top             =   1620
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "車　型"
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
            TabIndex        =   2
            Top             =   1125
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
         Begin Cis3D_v60.CIS3D H1lb_Torcd 
            Height          =   360
            Left            =   3345
            Top             =   1140
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
            Top             =   1140
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
         Begin CisText_V60.CisText iH1_Syake 
            Height          =   375
            Left            =   1980
            TabIndex        =   3
            Top             =   1590
            Width           =   765
            _ExtentX        =   1349
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
            Text            =   "XXXXX"
            MaxLength       =   5
         End
      End
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
      Caption         =   "【 内示表 発行 】"
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
         Left            =   14400
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
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
   Begin VB.Timer Timer2 
      Left            =   705
      Top             =   -120
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   -150
   End
   Begin VB.TextBox Dummy 
      Height          =   375
      Left            =   -15
      TabIndex        =   0
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
   Begin Cis3D_v60.CIS3D PB_Ent 
      Height          =   585
      Left            =   8280
      Top             =   8520
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
   Begin Cis3D_v60.CIS3D PB_Ok 
      Height          =   585
      Left            =   6915
      Top             =   8520
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
      Left            =   6900
      TabIndex        =   16
      Top             =   9195
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8280
      TabIndex        =   13
      Top             =   9195
      Width           =   1485
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5610
      TabIndex        =   12
      Top             =   9195
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
      TabIndex        =   11
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CNJ3110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  内示表 発行（キノックス向けレイアウト）
'**       フォームID    :  CNJ3110
'**       処理概要      :
'**
'**       作  成  日    :  2005/06/06
'**       変  更  日    :  2006/01/24  ”全て”選択時の異常終了バグ修正、ソート順の変更
'**       変  更  日    :  2006/01/30  帳票ヘッド部に『受入』追加(改ページ条件)
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim mPrtDev         As String       ' ﾌﾟﾘﾝﾀ出力先
    Dim mCsvDir         As String
    Dim mCsvKey         As String
    Dim mCsvName        As String

    Dim mMsgType        As String
    Dim SL_Select       As String

    Dim mSelColor       As Long
    Dim mUnSelColor     As Long
    Dim mSavePrint      As String

    Dim mSaveMsg        As String
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

    With CisFrm
        Set .MSGObject = lb_msg
        Set .MSGTimer = Timer2
        .MSGStopCount = 10
        .MSGTYPE = Left_To_Right
        .MSGText = mSaveMsg
        .MSGInterval = 100
        .MSGRelate
    End With

    Dummy.Left = -1000
    Dummy.Enabled = False

    mSelColor = Op_Out(0).BackColor
    mUnSelColor = Op_Out(1).BackColor

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

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

    ' 対象年月初期表示値を求める
    gSL_Select = "Select 英数字1,数字1,英数字2 From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '内示年月'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値 = '1'"
    If SYMRead(gSL_Select, 1) Then
        iH1_YM = SYM.英数字1
        CNJ3110_Naiji = Mid(SYM.英数字2, 1, 2)
    Else
        CNJ3110_Naiji = "定常"
    End If
    Op_OutKB(SYM.数字1).Value = True
    Op_OutKB(0).Caption = CNJ3110_Naiji & "＜初回＞"
'    gProcMsg = "差異リスト　発行を行います"
'    mMsgType = "L"

'CSVﾌｧｲﾙ出力先
    mCsvName = "内示表"
    mCsvKey = "CNJ0110"
    mCsvDir = GetSetting(appname:=gRegKey, section:="CSV", _
                       Key:=mCsvKey, Default:=Mid(CisFun.WinRoot, 1, 3) & mCsvName & ".csv")
''ﾌﾟﾘﾝﾀ出力先
    Op_Out(0).Value = True
    mPrtDev = Printer.DeviceName
    lb_PrtDev = mPrtDev
    mSavePrint = mPrtDev

    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 1000   ' 時刻表示間隔
    Timer2.Enabled = True    ' 時刻表示開始

    FormAct = Not FormAct

End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
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
        .Gradation 100, 160, 250, 4, 2, 3
    End With
End Sub

'+----------------------------------------+
'+  Op_OutKB   出力区分
'+----------------------------------------+
Private Sub Op_OutKB_Click(Index As Integer)

    Op_OutKB(0).Tag = Index
    For Each gObj In Op_OutKB
        If gObj.Index = Index Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj
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
'    gp_PrtKB = Index
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
'+    契約先検索        +
'+---------------------+
Private Sub PB_Torcd_Click()

    RV_TorKb = 2    ' 得意先
    RV_Call = "C"

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    iH1_Torcd.SetFocus

    If RV_Rtn Then
       iH1_Torcd = RV_TorcdK
       Call iH1_Torcd_LostFocus
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
    If TorNmGet(iH1_Torcd, 2) Then
        H1lb_Torcd = TRM_RName
        If Trim(H1lb_Torcd) = "" Then
           H1lb_Torcd = TRM_Name
        End If
    Else
        H1lb_Torcd = ""
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
       Case "H1"
          H1Mode = True
       Case "B2"
          B2Mode = True
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

    If Not Head1Chk Then GoTo ReturnPress_Ee

    Call DispChange("B1")

    If Not PrintStart Then GoTo ReturnPress_Ee

    gStr = "【  出  力  終  了  】"
    lb_msg.Caption = gStr
    DoEvents

ReturnPress_Ed:
    Call CisFrm.MousePT(1)
    With CisFrm
        .MSGTYPE = FSize_Changes
        .MSGText = gStr
    End With
    ProcHB = "B2"
    Call DispChange(ProcHB)
    Exit Sub
ReturnPress_Ee:
    lb_msg = mSaveMsg
    DoEvents
    ProcHB = "H1"
    Call DispChange(ProcHB)
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Dim wSl_Where   As String

    Head1Chk = False

    If Trim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　対象年月を入力して下さい    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If
    TGet_YM = RTrim$(iH1_YM)
    TGet_YMD = RTrim$(iH1_YM) & "01"

    If Trim(iH1_Torcd) <> "" Then
        If Trim(H1lb_Torcd) = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "   　得意先　未登録    "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
    End If
    
    lb_msg.Caption = "【 内示データ 抽出中 】"
    DoEvents

    If Op_OutKB(0).Value Or Op_OutKB(1).Value Then       ' 初回・計変
       gSL_Select = "Select TK.契約先,TR.正式名称,Cast(TK.オプション3 as Tinyint) 内示形態, "
       gSL_Select = gSL_Select & "  NJ.コメント,ISNULL(HN.品番編集,'') 品番編集,HN.車型,HN.品名,"
'       gSL_Select = gSL_Select & "  NT.*"                      ' 内示TEXTテーブルより発行
       gSL_Select = gSL_Select & "  NJ.*" ',NJ.工区 置換工区 "      ' 内示テーブルより発行
       gSL_Select = gSL_Select & "  From 内示TEXTテーブル NT"
       If Op_OutKB(0).Value Then
          gSL_Select = gSL_Select & "  Left Outer Join 内示初回ビュー NV"
       Else
          gSL_Select = gSL_Select & "  Left Outer Join 内示最新ビュー NV"
       End If
       gSL_Select = gSL_Select & "    On  NV.内示年月   = NT.内示年月"
       gSL_Select = gSL_Select & "    And NV.FILECODE = NT.FILECODE"
       gSL_Select = gSL_Select & "    And NV.HEAD作成日 = NT.HEAD作成日"
       gSL_Select = gSL_Select & "    And NV.HEADSEQ = NT.HEADSEQ"
       gSL_Select = gSL_Select & "  Left Outer Join 内示テーブル NJ"
       gSL_Select = gSL_Select & "    On  NJ.内示年月   = NT.内示年月"
       gSL_Select = gSL_Select & "    And NJ.品番 = NT.品番"
       gSL_Select = gSL_Select & "    And NJ.納入先 = NT.納入先"
       gSL_Select = gSL_Select & "    And NJ.受入 = NT.受入"
       gSL_Select = gSL_Select & "    And NJ.内示区分 = ''"
       gSL_Select = gSL_Select & "  Left Outer Join TEXT管理テーブル TK"
       gSL_Select = gSL_Select & "    On TK.FILECD = NT.FILECODE"
       gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR"
       gSL_Select = gSL_Select & "    On TR.取引先CD = TK.契約先"
       gSL_Select = gSL_Select & "    And TR.取引先区分 = 0"
       gSL_Select = gSL_Select & "  Left Outer Join 品番マスタ HN"
       gSL_Select = gSL_Select & "    On HN.品番 = NT.品番"
       With CisFun
            .Where_CreateKB = NewDocument
            .Where_Create AlphaMe, "NV.内示年月", "", NotNul, Create
       End With
    End If
    If Op_OutKB(2).Value Then       ' 全て
       gSL_Select = "Select TK.契約先,TR.正式名称,Cast(TK.オプション3 as Tinyint) 内示形態,"
       gSL_Select = gSL_Select & "  NJ.コメント,ISNULL(HN.品番編集,'') 品番編集,HN.車型,HN.品名,"
'       gSL_Select = gSL_Select & "  NT.*"                      ' 内示TEXTテーブルより発行
       gSL_Select = gSL_Select & "  NJ.*" ',NV.工区 置換工区 "      ' 内示テーブルより発行
       gSL_Select = gSL_Select & "  From 内示TEXTテーブル NT"
       gSL_Select = gSL_Select & "  Left Outer Join 内示テーブル NJ"
       gSL_Select = gSL_Select & "    On  NJ.内示年月   = NT.内示年月"
       gSL_Select = gSL_Select & "    And NJ.登録FCODE = NT.FILECODE"
       gSL_Select = gSL_Select & "    And NJ.登録H作成日 = NT.HEAD作成日"
       gSL_Select = gSL_Select & "    And NJ.登録HSEQ = NT.HEADSEQ"
       gSL_Select = gSL_Select & "    And NJ.登録連番 = NT.連番"
       gSL_Select = gSL_Select & "  Left Outer Join TEXT管理テーブル TK"
       gSL_Select = gSL_Select & "    On TK.FILECD = NT.FILECODE"
       gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR"
       gSL_Select = gSL_Select & "    On TR.取引先CD = TK.契約先"
       gSL_Select = gSL_Select & "    And TR.取引先区分 = 0"
       gSL_Select = gSL_Select & "  Left Outer Join 品番マスタ HN"
       gSL_Select = gSL_Select & "    On HN.品番 = NT.品番"
        With CisFun
            .Where_CreateKB = NewDocument
            .Where_Create AlphaMe, "NJ.内示年月", "", NotNul, Create
            .Where_Create AlphaMe, "ISNULL(NJ.内示年月,'')", "", 等しくない, Create
        End With
    End If
'
    With CisFun
        .Where_CreateKB = ExistDocument
        .Where_Create AlphaMe, "NT.内示年月", iH1_YM
        .Where_Create AlphaMe, "TK.契約先", iH1_Torcd
        .Where_Create AlphaMe, "HN.車型", iH1_Syake
        wSl_Where = .Where_Phrase
    End With

    gSL_Select = gSL_Select & wSl_Where
'    gSL_Select = gSL_Select & " Order By TK.契約先,HN.車型,NT.品番,NT.背番号,NJ.受入"

'< 2006/01/30   MNT START >--------------------------------------------------------------------------------
'
'    gSL_Select = gSL_Select & " Order By TK.契約先,HN.車型,NJ.受入,NT.品番,NT.背番号"           '2006/01/24
    gSL_Select = gSL_Select & " Order By TK.契約先, NJ.受入, HN.車型, NT.品番, NT.背番号"
'
'< 2006/01/30   MNT  END  >--------------------------------------------------------------------------------
    
    If Not NTTRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　該当データがありません       "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If


    Head1Chk = True
End Function

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   Call CisFrm.TimeSet("1")
End Sub
'*******************************
'*      メッセージ表示        *
'*******************************
Private Sub Timer2_Timer()
    CisFrm.MSGRelate
End Sub
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
        lb_msg.Caption = "【 内示データ 印刷中 】"
        DoEvents
        With CisPrt
             Set .iPrtForm = CNJ3110L
             .dpMSG = "内示表　印刷中"
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
    Dim OutStr      As String
    Dim OutFNo      As Long
    Dim OutCnt      As Long
    Dim wPutCnt     As Long
    Dim wYM         As String   ' 当月
    Dim wYMy        As String   ' 翌月
    Dim wYMyy       As String   ' 翌々月
    Dim wHead1      As String
    Dim wHead2      As String
    Dim wCnt        As Integer
    
    CsvPut = False
    
    PgBk_Torcd = ""
    PgBk_Syake = ""
    Erase wKadou
    
    ' 出力ﾌｧｲﾙ存在確認
    If Dir(mCsvDir) <> "" Then
        DoEvents
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　 " & mCsvDir & "　ﾌｧｲﾙは既に存在します。                 "
            .MB_MSG(4) = "   　 " & "上書きしますか？           "
            .MB_Title = "上書き確認"
            .MB_Button = Yes_No
            If Not .MBOX Then Exit Function
        End With
    End If

    lb_msg.Caption = "【 内示データ 出力中 】"
    Csv_Bar.Visible = True

    DoEvents

    OutFNo = FreeFile
    Open mCsvDir For Output As #OutFNo

    Call CisFrm.MousePT(11)

    With CisFrm
        Set .MeterPic = Csv_Bar
        .MeterMax = CisDB.RecordCount
    End With
    
    '翌月・翌々月を求める
    gStr = NTT.内示年月
    wNengetu(0) = NTT.内示年月
    wYM = Mid(gStr, 3, 2) & "年 " & CisFun.RSetFld(Mid(gStr, 5, 2), 2, "##") & "月"
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add gStr & "01", M, 1, LastDate
        wNengetu(1) = .Date_Result
        gStr = .Date_Result
        wYMy = Mid(gStr, 3, 2) & "年 " & CisFun.RSetFld(Mid(gStr, 5, 2), 2, "##") & "月"
        .Date_RsFormat = fYM
        .Date_Add gStr & "01", M, 1, LastDate
        wNengetu(2) = .Date_Result
        gStr = .Date_Result
        wYMyy = Mid(gStr, 3, 2) & "年 " & CisFun.RSetFld(Mid(gStr, 5, 2), 2, "##") & "月"
    End With
        
    '見出し編集
    wHead1 = """内示年月"",""契約先"",""契約先名"",""車型"""
    wHead2 = """№"",""背番号"",""品番"",""品名"",""収容数"","
    wHead2 = wHead2 & """" & wYM & """,""日当数"","
    wHead2 = wHead2 & """" & wYMy & """,""日当数"","
    wHead2 = wHead2 & """" & wYMyy & """,""日当数"""
    
    wPutCnt = 0
    With NTT
        '品番編集初期値を求める   品番未登録でも編集したい時はコメントアウトする
        'Call EditHinbnInit
        
        OutCnt = 0:         OutStr = ""
        PgBk_Torcd = "":    PgBk_Syake = ""
        
        Do Until Not NTT_RDSTS
            If RTrim$(.契約先) <> PgBk_Torcd Or RTrim$(.車型) <> PgBk_Syake Or _
               PrtK_Torcd <> RTrim$(.契約先) Or PrtK_Ukeir <> RTrim$(.受入) Then
                If PrtK_Torcd <> RTrim$(.契約先) Or PrtK_Ukeir <> RTrim$(.受入) Then
                    '月別稼働日数
                    For wCnt = 0 To 2
                        Call GetKadou(wNengetu(wCnt))   'カレンダマスタの稼働日数取得
                        wKadou(wCnt) = CLM.稼動日数
                    Next
                End If
                If RTrim$(.契約先) <> PgBk_Torcd Or RTrim$(.車型) <> PgBk_Syake Then
                    '見出し出力
                    If RTrim$(wHead1) <> "" Then Print #OutFNo, wHead1: wHead1 = ""
                    OutStr = """" & Mid(.内示年月, 1, 4) & " / " & Mid(.内示年月, 5, 2) & ""","
                    OutStr = OutStr & """" & .契約先 & ""","
                    OutStr = OutStr & """" & .正式名称 & ""","
                    OutStr = OutStr & """" & .車型 & ""","
                    If .内示形態 = 0 Then
                       OutStr = OutStr & """【" & CNJ3110_Naiji & "】"""
                    Else
                       OutStr = OutStr & """【計画変更】"""
                    End If
                    Print #OutFNo, OutStr
                    If RTrim$(wHead2) <> "" Then Print #OutFNo, wHead2: wHead2 = ""
                End If
                PrtK_Torcd = RTrim$(.契約先)
                PrtK_Syake = RTrim$(.車型)
                PrtK_Hinbn = RTrim$(.品番)
                PrtK_Seban = RTrim$(.背番号)
                PrtK_Ukeir = RTrim$(.受入)
                PrtK_Syuyo = RTrim$(.収容数)
                OutStr = "": OutCnt = 0
            End If
            
            PgBk_Torcd = RTrim$(.契約先)
            PgBk_Syake = RTrim$(.車型)
            
            OutCnt = OutCnt + 1
            OutStr = """" & OutCnt & ""","
            OutStr = OutStr & """" & RTrim(.背番号) & ""","
            OutStr = OutStr & """" & EditHinbn(.品番, .品番編集) & ""","
            OutStr = OutStr & """" & RTrim(.品名) & ""","
            OutStr = OutStr & """" & .収容数 & ""","

            OutStr = OutStr & """" & .内示数 & ""","
            If wKadou(0) > 0 Then
                gLong = CisFun.HasuProc((.内示数 / wKadou(0)), 切上げ, 0)
                OutStr = OutStr & """" & gLong & ""","
            End If
            
            OutStr = OutStr & """" & .翌月内示数 & ""","
            If wKadou(1) > 0 Then
                gLong = CisFun.HasuProc((.翌月内示数 / wKadou(1)), 切上げ, 0)
                OutStr = OutStr & """" & gLong & ""","
            End If
            
            OutStr = OutStr & """" & .翌々月内示数 & ""","
            If wKadou(2) > 0 Then
                gLong = CisFun.HasuProc((.翌々月内示数 / wKadou(2)), 切上げ, 0)
                OutStr = OutStr & """" & gLong & """"
            End If
            
            Print #OutFNo, OutStr
            wPutCnt = wPutCnt + 1

            CisFrm.MeterUpdate wPutCnt

            Call NTTReadNext
        Loop
    End With
CsvPut_Ed:
    Call NTTClose

    Close #OutFNo
    SaveSetting appname:=gRegKey, section:="CSV", _
                    Key:=mCsvKey, setting:=RTrim(lb_PrtDev)

    Call CisFrm.MousePT(1)

    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "   出力処理が終了しました。    "
        .MB_MSG(4) = "   　 " & "出力先    = " & mCsvDir
        .MB_MSG(5) = "   　 " & "出力件数 = " & StrConv(Format(wPutCnt, "##,##0"), vbWide) & " 件 "
        .MB_Title = "出力確認"
        .MB_Button = OK
        .MBOX
    End With
    CsvPut = True

End Function
