VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CNJ0220 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "生産計画表　発行"
   ClientHeight    =   11205
   ClientLeft      =   2070
   ClientTop       =   1680
   ClientWidth     =   15360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11205
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin VB.PictureBox Csv_Bar 
      Height          =   345
      Left            =   4320
      ScaleHeight     =   285
      ScaleWidth      =   6525
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   6585
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1605
      Top             =   4770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   585
      Left            =   5655
      Top             =   7650
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
      BackColor       =   &H00C0C0C0&
      Height          =   4695
      Left            =   4290
      ScaleHeight     =   4824.189
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2535
      Width           =   6660
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   4530
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   7990
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
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   465
            Left            =   2040
            Top             =   2520
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   820
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
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cBoderStyle     =   1
            Begin VB.OptionButton Op_OutKB 
               BackColor       =   &H00C0C0C0&
               Caption         =   "計画変更分のみ"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   1350
               TabIndex        =   4
               Top             =   120
               Width           =   1980
            End
            Begin VB.OptionButton Op_OutKB 
               BackColor       =   &H00C0C0C0&
               Caption         =   "全て"
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
               Left            =   330
               TabIndex        =   3
               Top             =   120
               Value           =   -1  'True
               Width           =   990
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   375
            Left            =   870
            Top             =   1890
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   661
            ForeColor       =   16711680
            Caption         =   "車型"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   465
            Left            =   870
            Top             =   2520
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   820
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
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   375
            Left            =   2040
            TabIndex        =   0
            Top             =   930
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
            TabIndex        =   14
            Top             =   3330
            Width           =   6345
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
               TabIndex        =   10
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
                  Caption         =   "生産計画表　発行を行います"
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
            Top             =   4080
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
            Top             =   4110
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
            Height          =   375
            Index           =   1
            Left            =   870
            Top             =   930
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   661
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
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   375
            Left            =   870
            Top             =   1410
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   661
            ForeColor       =   16711680
            Caption         =   "工  区"
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
            cAlingnment     =   7
         End
         Begin CisText_V60.CisText iH1_Shakei 
            Height          =   375
            Left            =   2040
            TabIndex        =   2
            Top             =   1890
            Width           =   780
            _ExtentX        =   1376
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
         Begin CisText_V60.CisText iH1_KOKbn 
            Height          =   375
            Left            =   2040
            TabIndex        =   1
            Top             =   1410
            Width           =   270
            _ExtentX        =   476
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
            Text            =   "X"
            MaxLength       =   1
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
      Caption         =   "【 生産計画表 発行 】"
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
         Width           =   885
         _ExtentX        =   1561
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
         Left            =   11820
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
      TabIndex        =   9
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
   Begin Cis3D_v60.CIS3D PB_Ent 
      Height          =   585
      Left            =   8310
      Top             =   7650
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
      Left            =   6945
      Top             =   7650
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
      Left            =   6930
      TabIndex        =   15
      Top             =   8325
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8310
      TabIndex        =   13
      Top             =   8325
      Width           =   1485
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5640
      TabIndex        =   12
      Top             =   8325
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
Attribute VB_Name = "CNJ0220"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  生産計画表 発行
'**       フォームID    :  CNJ0220
'**       処理概要      :
'**
'**       作  成  日    :  2003/12/16
'**       変  更  日    :  2007/08/31 By CIS - CSV作成時にｴﾗｰ('ｵｰﾊﾞｰﾌﾛｰ')が出る点を修正
'**                        2008/04/11 By CIS  背番号8桁,車型5桁
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
    gSL_Select = gSL_Select & " Where 区分名称 = '処理内容'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値 = 'PNJ0210'"
    If SYMRead(gSL_Select, 1) Then
        iH1_YM = SYM.英数字1
    End If
    
'CSVﾌｧｲﾙ出力先
    mCsvName = "生産計画表"
    mCsvKey = "CNJ0220"
    mCsvDir = GetSetting(appname:=gRegKey, section:="CSV", _
                       Key:=mCsvKey, Default:=Mid(CisFun.WinRoot, 1, 3) & mCsvName & ".csv")
''ﾌﾟﾘﾝﾀ出力先
    Op_Out(0).Value = True
    mPrtDev = Printer.DeviceName
    lb_PrtDev = mPrtDev
    mSavePrint = mPrtDev
    
    Op_OutKB(0).Value = True
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 1000   ' 時刻表示間隔
    Timer2.Enabled = True    ' 時刻表示開始
    
    wKadouFlg = False
    
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
'        Case vbKeyF7:       If Dummy.Tag <> "" Then Call PB_Look_Click      '【検索】
        Case vbKeyF12
                            If PB_END.Visible Then Call PB_END_Click        '【終了】
                            If PB_Ok.Visible Then Call PB_END_Click         '【終了】
        Case vbKeyReturn
                            If PB_Ent.Visible Then                          '【入力】
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
    PB_Ent.Visible = H1Mode
    PNL_ENT.Visible = H1Mode
    PB_END.Visible = H1Mode
    PNL_END.Visible = H1Mode
    Timer2.Enabled = H1Mode Or B2Mode
    
    PB_Ok.Visible = B2Mode
    PNL_Ok.Visible = B2Mode
    
'    Csv_Bar.Visible = False
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

    gSL_Select = "SELECT COUNT(*) 件数 "
    gSL_Select = gSL_Select & " FROM 生産計画テーブル SW"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN "
    gSL_Select = gSL_Select & "   ON HN.品番 = SW.品番 "
    gSL_Select = gSL_Select & " WHERE SUBSTRING(納期,1,6) = '" & RTrim(iH1_YM) & "'"
    If CisFun.Val2(Op_OutKB(0).Tag) <> 0 Then
        gSL_Select = gSL_Select & " AND SW.変更回数 <> 0 "
    End If
    If RTrim(iH1_KOKbn) <> "" Then
        gSL_Select = gSL_Select & " AND SW.工区 = '" & RTrim(iH1_KOKbn) & "'"
    End If
    If RTrim(iH1_Shakei) <> "" Then
        gSL_Select = gSL_Select & " AND HN.車型 = '" & RTrim(iH1_Shakei) & "'"
    End If
    
    If Not SKHRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　該当データが 存在しません           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If
    
    If SKH.件数 = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　該当データが 存在しません           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If
    
' 生産計画表データ 抽出
    Timer2.Enabled = False
    lb_msg.Caption = "【 生産計画表データ 抽出中 】"
    DoEvents
    
    With CisDB
        .SQL = "生産計画表抽出"
        .StoadoCount = 2
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
    
        .ParaNo = 1:    .ParaValue = RTrim(iH1_YM)
        .ParaNo = 2:    .ParaValue = CisFun.Val2(Op_OutKB(0).Tag)
        .DBStored
    
        .ParaNo = 0
        If .ParaValue <> 0 Then
            CisFun.MB_Lines = 5
            CisFun.MB_MSG(2) = "    生産計画表 抽出処理が異常終了しました      "
            CisFun.MB_Button = OK
            CisFun.MBOX
            Exit Function
        End If
    End With
    
    
' * 生産計画表ワーク
    gSL_Select = "SELECT SW.工区,ISNULL(HN.車型,'') 車型,SW.品番,ISNULL(HN.社内背番,'') 社内背番,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(HN.基本収容数,0) = 0 "
    gSL_Select = gSL_Select & " THEN 1 ELSE HN.基本収容数 END 収容数,"
    gSL_Select = gSL_Select & " HN.品番編集, "
    gSL_Select = gSL_Select & " SUM(SW.前月計画数) 前月計画数,"
    gSL_Select = gSL_Select & " SUM(SW.当月計画数) 当月計画数,"
    gSL_Select = gSL_Select & " SUM(SW.前月内示数) 前月内示数,"
    gSL_Select = gSL_Select & " SUM(SW.当月内示数) 当月内示数,"
    gSL_Select = gSL_Select & " SUM(SW.翌月内示数) 翌月内示数,"
    gSL_Select = gSL_Select & " SUM(SW.翌々月内示数) 翌々月内示数,SW.変更区分"
    gSL_Select = gSL_Select & " FROM 生産計画表発行ワーク SW "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN "
    gSL_Select = gSL_Select & "   ON HN.品番 = SW.品番 "
    gSL_Select = gSL_Select & " WHERE 処理端末 = HOST_NAME() "
    If RTrim(iH1_KOKbn) <> "" Then
        gSL_Select = gSL_Select & "   AND SW.工区 = '" & RTrim(iH1_KOKbn) & "'"
    End If
    If RTrim(iH1_Shakei) <> "" Then
        gSL_Select = gSL_Select & "   AND HN.車型 = '" & RTrim(iH1_Shakei) & "'"
    End If
    gSL_Select = gSL_Select & " GROUP BY SW.工区,HN.車型,SW.品番,HN.社内背番,HN.基本収容数,SW.変更区分,HN.品番編集"
    gSL_Select = gSL_Select & " ORDER BY SW.工区,HN.車型,SW.品番"
    Call SKHRead(gSL_Select)
    
    gNJYM = RTrim(iH1_YM)
    
    Head1Chk = True
End Function

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
        lb_msg.Caption = "【 生産計画データ 印刷中 】"
        DoEvents

        With CisPrt
             Set .iPrtForm = CNJ0220L
             .dpMSG = "生産計画表　印刷中"
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
    Dim End_Day     As Long
    Dim wYM         As String   ' 当月
    Dim wYMy        As String   ' 翌月
    Dim wYMyy       As String   ' 翌々月
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2007/08/31 UPDATE START
    Dim Save_RecordCount As Long
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2007/08/31 UPDATE END
    
    CsvPut = False
    
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
    
    Timer2.Enabled = False
    lb_msg.Caption = "【 生産計画データ 出力中 】"
    Csv_Bar.Visible = True
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2007/08/31 INSERT START
    Save_RecordCount = CisDB.RecordCount
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2007/08/31 INSERT END
    
    DoEvents
        
    OutFNo = FreeFile
    Open mCsvDir For Output As #OutFNo

    Call CisFrm.MousePT(11)
    
' 稼動状況(カレンダマスタ参照)
    Call GetKadou
    
    With CisFrm
        Set .MeterPic = Csv_Bar
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2007/08/31 UPDATE START
        '.MeterMax = CisDB.RecordCount
        .MeterMax = Save_RecordCount
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2007/08/31 UPDATE END
    End With
        
    OutStr = "対象年月,工区,車型,出力区分,,"
    OutStr = OutStr & wMM(0) & ","
    OutStr = OutStr & wMM(1) & ","
    OutStr = OutStr & wMM(2) & ","
    OutStr = OutStr & wMM(3) & ","
    
    Print #OutFNo, OutStr
    With SKH
        OutStr = Mid(iH1_YM, 1, 4) & "/" & Mid(iH1_YM, 5, 2) & ","
        OutStr = OutStr & RTrim(iH1_KOKbn) & ","
        OutStr = OutStr & RTrim(iH1_Shakei) & ","
        If CisFun.Val2(Op_OutKB(0).Tag) = 0 Then
            OutStr = OutStr & "全て,"
        Else
            OutStr = OutStr & "変更分のみ,"
        End If
        OutStr = OutStr & ","
        OutStr = OutStr & wKad(0) & ","
        OutStr = OutStr & wKad(1) & ","
        OutStr = OutStr & wKad(2) & ","
        OutStr = OutStr & wKad(3) & ","
        
        Print #OutFNo, OutStr
        
        OutCnt = 0
        OutStr = "№,品番,背番号,収容数,"
        OutStr = OutStr & Mid(wMM(1), 1, 2) & "月 計画数,"
        OutStr = OutStr & Mid(wMM(1), 1, 2) & "月 計画日当数,"
        OutStr = OutStr & Mid(wMM(1), 1, 2) & "月 計画日当枚数,"
        OutStr = OutStr & Mid(wMM(0), 1, 2) & "月 計画数,"
        OutStr = OutStr & Mid(wMM(0), 1, 2) & "月 計画日当数,"
        OutStr = OutStr & Mid(wMM(0), 1, 2) & "月 計画日当枚数,"
        OutStr = OutStr & "計画日当枚数差,"
        OutStr = OutStr & Mid(wMM(1), 1, 2) & "月 内示数,"
        OutStr = OutStr & Mid(wMM(0), 1, 2) & "月 内示数,"
        OutStr = OutStr & "内示数差,"
        OutStr = OutStr & Mid(wMM(2), 1, 2) & "月 内示数,"
        OutStr = OutStr & Mid(wMM(2), 1, 2) & "月 内示日当枚数,"
        OutStr = OutStr & Mid(wMM(3), 1, 2) & "月 内示数,"
        OutStr = OutStr & Mid(wMM(3), 1, 2) & "月 内示日当枚数,"
        
        Print #OutFNo, OutStr
        
        Do Until Not SKH_RDSTS
            OutCnt = OutCnt + 1
            OutStr = OutCnt & ","
            OutStr = OutStr & EditHinbn(.品番, .品番編集) & ","
            OutStr = OutStr & RTrim(.社内背番) & ","
            OutStr = OutStr & .収容数 & ","
            OutStr = OutStr & .当月計画数 & ","
            OutStr = OutStr & CisFun.HasuProc(.当月計画数 / wKadsu(1), 切上げ) & ","
            OutStr = OutStr & CisFun.HasuProc(.当月計画数 / wKadsu(1) / .収容数, 切上げ) & ","
            OutStr = OutStr & .前月計画数 & ","
            OutStr = OutStr & CisFun.HasuProc(.前月計画数 / wKadsu(0), 切上げ) & ","
            OutStr = OutStr & CisFun.HasuProc(.前月計画数 / wKadsu(0) / .収容数, 切上げ) & ","
            OutStr = OutStr & CisFun.HasuProc(.当月計画数 / wKadsu(1) / .収容数, 切上げ) - _
                              CisFun.HasuProc(.前月計画数 / wKadsu(0) / .収容数, 切上げ) & ","
            OutStr = OutStr & .当月内示数 & ","
            OutStr = OutStr & .前月内示数 & ","
            OutStr = OutStr & .当月内示数 - .前月内示数 & ","
            OutStr = OutStr & .翌月内示数 & ","
            OutStr = OutStr & CisFun.HasuProc(.翌月内示数 / wKadsu(2) / .収容数, 切上げ) & ","
            OutStr = OutStr & .翌々月内示数 & ","
            OutStr = OutStr & CisFun.HasuProc(.翌々月内示数 / wKadsu(3) / .収容数, 切上げ) & ","
            
            Print #OutFNo, OutStr
    
            CisFrm.MeterUpdate OutCnt
            
            Call SKHReadNext
        Loop
    End With
CsvPut_Ed:
    Call SKHClose

    Close #OutFNo
    SaveSetting appname:=gRegKey, section:="CSV", _
                    Key:=mCsvKey, setting:=RTrim(lb_PrtDev)
    
    Call CisFrm.MousePT(1)
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "   出力処理が終了しました。    "
        .MB_MSG(4) = "   　 " & "出力先    = " & mCsvDir
        '-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2007/08/31 UPDATE START
        '.MB_MSG(5) = "   　 " & "出力件数 = " & StrConv(Format(CisDB.RecordCount, "##,##0"), vbWide) & " 件 "
        .MB_MSG(5) = "   　 " & "出力件数 = " & StrConv(Format(Save_RecordCount, "##,##0"), vbWide) & " 件 "
        '-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2007/08/31 UPDATE END
        .MB_Title = "出力確認"
        .MB_Button = OK
        .MBOX
    End With
    CsvPut = True
    
End Function


