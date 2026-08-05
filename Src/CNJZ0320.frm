VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CNJZ0320 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "材料発注内示表　発行"
   ClientHeight    =   11205
   ClientLeft      =   2340
   ClientTop       =   2295
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11205
   ScaleWidth      =   15360
   Begin VB.PictureBox Csv_Bar 
      Height          =   345
      Left            =   4350
      ScaleHeight     =   285
      ScaleWidth      =   6600
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2250
      Visible         =   0   'False
      Width           =   6660
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1695
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   4470
      Left            =   4350
      ScaleHeight     =   4590.005
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2715
      Width           =   6660
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   4275
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   7541
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
         Begin VB.OptionButton Op_OutKB 
            BackColor       =   &H00C0C0C0&
            Caption         =   "最新変更分のみ"
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
            Left            =   2055
            TabIndex        =   4
            Top             =   2715
            Width           =   1950
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
            Index           =   2
            Left            =   2055
            TabIndex        =   3
            Top             =   2355
            Value           =   -1  'True
            Width           =   1950
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   285
            Left            =   645
            Top             =   1815
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
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
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   375
            Left            =   2055
            TabIndex        =   1
            Top             =   960
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
            TabIndex        =   11
            Top             =   3090
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
               TabIndex        =   9
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
                  Caption         =   "材料発注内示表　発行を行います"
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
            Top             =   3780
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
            Top             =   3810
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
            Left            =   645
            Top             =   960
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
            ForeColor       =   16711680
            Caption         =   "内示年月"
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   2055
            TabIndex        =   2
            Top             =   1785
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
            Left            =   3420
            Top             =   1785
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
            Left            =   3075
            Top             =   1785
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
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   285
            Left            =   645
            Top             =   2325
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
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
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D H1lb_YM 
            Height          =   285
            Left            =   2040
            Top             =   1365
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   503
            ForeColor       =   128
            Caption         =   "( 対象年月 ： 9999年 99月 )"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
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
      End
   End
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   661
      BackColor       =   8438015
      ForeColor       =   0
      Caption         =   "【 材料発注内示表 発行 】"
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
         Left            =   14430
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   450
         BackColor       =   8438015
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
         cBoderColor1    =   8438015
         cBoderColor2    =   8438015
         cFont3DColor1   =   8454143
         cFont3DColor2   =   16777215
         cFont3D         =   1
         cAlingnment     =   7
         cBoderStyle     =   2
      End
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   11880
         Top             =   30
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   503
         BackColor       =   8438015
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
      IMEMode         =   3  'ｵﾌ固定
      Left            =   -15
      TabIndex        =   0
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   585
      Left            =   5625
      Top             =   7860
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
   Begin Cis3D_v60.CIS3D PB_Ent 
      Height          =   585
      Left            =   8280
      Top             =   7860
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
      Top             =   7860
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
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5610
      TabIndex        =   15
      Top             =   8535
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8280
      TabIndex        =   14
      Top             =   8535
      Width           =   1485
   End
   Begin VB.Label PNL_Ok 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6900
      TabIndex        =   13
      Top             =   8535
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
      TabIndex        =   10
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CNJZ0320"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  材料発注内示表 発行
'**       フォームID    :  CNJZ0320
'**       処理概要      :
'**
'**       作  成  日    :  2004/03/24
'**       変  更  日    :  2010/02/09  CSV出力部修正
'**       変  更  日    :
'**       変  更  日    :  2010/02/24  対象年月算出の変更→材料発注管理_内示対象データ(CISKANRI：'Option'_11)が当月(0)の場合
'**                                    CisFun.Date_Addを回避し内示年月をそのまま対象年月とする。
'**                                    YMAdd:内示対象年月(加算値)をPublic項目に変更 ⇒ CNJZ0320LB.bas へ移動
'**                                    CSV出力時の材料名にカンマが入っているのでﾄﾞｯﾄに変更
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
    'Dim YMAdd           As Byte         ' 内示対象年月(加算値)     'MOVE 2010/02/24

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
    
'< 材料発注内示年月 >
    Call GetNJProc_YM(3)
    If RTrim$(gn_HNYM) = "" Then gn_HNYM = Format(Now(), "YYYYMM")
    iH1_YM = gn_HNYM
    TGet_NJYM = gn_HNYM
    
    Call Zai_Nai_YM

'CSVﾌｧｲﾙ出力先
    mCsvName = "材料発注内示表"
    mCsvKey = "CNJZ0320"
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
    
'< 前回テーブル確認 >
    gInt = 2
    gSL_Select = "if exists("
    gSL_Select = gSL_Select & " select 内示年月"
    gSL_Select = gSL_Select & " From 材料発注内示前回テーブル"
    gSL_Select = gSL_Select & " Where 内示年月 = '" & RTrim$(iH1_YM) & "' )"
    gSL_Select = gSL_Select & " begin"
    gSL_Select = gSL_Select & " select 1 件数"
    gSL_Select = gSL_Select & " End"
    gSL_Select = gSL_Select & " Else"
    gSL_Select = gSL_Select & " begin"
    gSL_Select = gSL_Select & " select 2 件数"
    gSL_Select = gSL_Select & " End"
    If ZNJRead(gSL_Select, 1) Then gInt = ZNJ.件数
    
    Op_OutKB(gInt).Value = True
    
' 内示対象年月
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 11
    Call IRNRead("", 1)
    YMAdd = CisFun.Val2(IRN.oyian)
    
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
        .Gradation 160, 250, 100, 2, 3, 4
    End With
End Sub
'+----------------------------------------+
'+  Op_Out   出力先
'+----------------------------------------+
Private Sub Op_Out_GotFocus(INDEX As Integer)
    Op_Out(0).Tag = INDEX
    For Each gObj In Op_Out
        If gObj.INDEX = INDEX Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj
    Select Case INDEX
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
'+    仕入先検索        +
'+---------------------+
Private Sub PB_Torcd_Click()

    RV_TorKb = 8    ' 手配先
    RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    iH1_Torcd.SetFocus

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        Call iH1_Torcd_LostFocus
        H1lb_Torcd = RTrim$(RV_TorRName)
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
    If TorNmGet(iH1_Torcd, 8) Then
        H1lb_Torcd = TRM_RName
        If Trim(H1lb_Torcd) = "" Then
           H1lb_Torcd = TRM_Name
        End If
    Else
        H1lb_Torcd = ""
    End If
End Sub
'+-------------------------------------+
'+    出力区分  Op_OutKB    (Click)    +
'+-------------------------------------+
Private Sub Op_OutKB_Click(INDEX As Integer)
    opOutKB = INDEX
    For Each gObj In Op_OutKB
        If gObj.INDEX = INDEX Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj
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
    
'( 内示年月 )
    If Trim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　内示年月を入力して下さい    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If
    TGet_NJYM = RTrim$(iH1_YM)
    
'( 対象年月 )
    Call Zai_Nai_YM
    
'( 仕入先 )
    If Trim(iH1_Torcd) <> "" Then
        If Trim(H1lb_Torcd) = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "   　仕入先　未登録    "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
    End If

    lb_msg.Caption = "【 材料発注内示データ 抽出中 】"
    DoEvents

'[ 印刷データ読込 ]

'   ( ３ヶ月　月量データ ) < 納期:｢00000000｣/当月数量、 ｢00000001｣/翌月数量、 ｢00000002｣/翌々月数量 >
    gSL_Select = "SELECT  HJ.材料管理番号, HJ.手配先, HJ.受入, HJ.直送先, '0000000' +" & vbCrLf
    gSL_Select = gSL_Select & " CAST( DATEDIFF( M, HJ.内示年月 + '01', HJ.計画年月 + '01' ) AS VARCHAR(1) ) 納期," & vbCrLf
    gSL_Select = gSL_Select & " SUM( HJ.数量 ) 数量," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( MAX( TR.正式名称 ), '' ) 手配先名," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( MAX( HN.材質 ), '' ) 材質," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( MAX( HN.寸法区分 ), 1 ) 寸法区分," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( MAX( HN.板厚 ), 0 ) 板厚," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( MAX( HN.幅 ), 0 ) 幅," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( MAX( HN.長さ ), 0 ) 長さ," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( MAX( HN.材種コード ), '' ) 材種コード," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( MAX( SY.値名称 ), '' ) 手配単位名," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( MAX( HN.材料名 ),     '' ) 材料名," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( MAX( HT.背番号 ),   '' ) 背番号," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( MAX( HT.サイクル ), '' ) サイクル," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( MAX( CASE WHEN ISNULL( HT.変更日, '9' ) > HJ.計画年月 + '01'" & vbCrLf
    gSL_Select = gSL_Select & "                     OR ISNULL( HT.変更日,  '' ) = ''" & vbCrLf
    gSL_Select = gSL_Select & "                   THEN ISNULL( HT.収容数1, 0 )" & vbCrLf
    gSL_Select = gSL_Select & "                   ELSE ISNULL( HT.収容数2, 0 )" & vbCrLf
    gSL_Select = gSL_Select & "              END ), 0 ) 収容数" & vbCrLf
    gSL_Select = gSL_Select & " FROM 材料発注内示テーブル HJ" & vbCrLf
    
    If opOutKB = 1 Then
        gSL_Select = gSL_Select & " INNER JOIN ( SELECT DISTINCT" & vbCrLf
        gSL_Select = gSL_Select & "                     ISNULL( HZ.内示年月, HN.内示年月 ) 内示年月," & vbCrLf
        gSL_Select = gSL_Select & "                     ISNULL( HZ.材料管理番号, HN.材料管理番号 ) 材料管理番号," & vbCrLf
        gSL_Select = gSL_Select & "                     ISNULL( HZ.手配先, HN.手配先 ) 手配先," & vbCrLf
        gSL_Select = gSL_Select & "                     ISNULL( HZ.受入, HN.受入 ) 受入," & vbCrLf
        gSL_Select = gSL_Select & "                     ISNULL( HZ.直送先, HN.直送先 ) 直送先" & vbCrLf
        gSL_Select = gSL_Select & "                     FROM 材料発注内示前回テーブル HZ" & vbCrLf
        gSL_Select = gSL_Select & "                     FULL OUTER JOIN 材料発注内示テーブル HN" & vbCrLf
        gSL_Select = gSL_Select & "                     ON    HN.内示年月 = HZ.内示年月" & vbCrLf
        gSL_Select = gSL_Select & "                       AND HN.計画年月 = HZ.計画年月" & vbCrLf
        gSL_Select = gSL_Select & "                       AND HN.材料管理番号 = HZ.材料管理番号" & vbCrLf
        gSL_Select = gSL_Select & "                       AND HN.手配先 = HZ.手配先" & vbCrLf
        gSL_Select = gSL_Select & "                       AND HN.受入 = HZ.受入" & vbCrLf
        gSL_Select = gSL_Select & "                       AND HN.直送先 = HZ.直送先" & vbCrLf
        gSL_Select = gSL_Select & "                       AND HN.納期 = HZ.納期" & vbCrLf
        gSL_Select = gSL_Select & "                       AND HN.数量 = HZ.数量" & vbCrLf
        gSL_Select = gSL_Select & "                     WHERE (HN.材料管理番号 Is Null Or HZ.材料管理番号 Is Null)" & vbCrLf
        gSL_Select = gSL_Select & "                       AND ISNULL( HZ.内示年月, '' )  = CASE WHEN HZ.材料管理番号 IS NULL THEN '' ELSE '" & RTrim$(TGet_NJYM) & "' END" & vbCrLf
        gSL_Select = gSL_Select & "                       AND ISNULL( HZ.計画年月, '' ) >= CASE WHEN HZ.材料管理番号 IS NULL THEN '' ELSE '" & RTrim$(TGet_YM) & "' END" & vbCrLf
        gSL_Select = gSL_Select & "                       AND ISNULL( HN.内示年月, '' )  = CASE WHEN HN.材料管理番号 IS NULL THEN '' ELSE '" & RTrim$(TGet_NJYM) & "' END" & vbCrLf
        gSL_Select = gSL_Select & "                       AND ISNULL( HN.計画年月, '' ) >= CASE WHEN HN.材料管理番号 IS NULL THEN '' ELSE '" & RTrim$(TGet_YM) & "' END" & vbCrLf
        gSL_Select = gSL_Select & "            ) BD" & vbCrLf
        gSL_Select = gSL_Select & " ON  BD.内示年月 = HJ.内示年月" & vbCrLf
        gSL_Select = gSL_Select & " And BD.材料管理番号 = HJ.材料管理番号" & vbCrLf
        gSL_Select = gSL_Select & " And BD.手配先 = HJ.手配先" & vbCrLf
        gSL_Select = gSL_Select & " And BD.受入 = HJ.受入" & vbCrLf
        gSL_Select = gSL_Select & " And BD.直送先 = HJ.直送先" & vbCrLf
    End If
    
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR" & vbCrLf
    gSL_Select = gSL_Select & " ON  TR.取引先CD = HJ.手配先" & vbCrLf
    gSL_Select = gSL_Select & " AND TR.取引先区分 = 1" & vbCrLf
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料マスタ HN" & vbCrLf
    gSL_Select = gSL_Select & " ON  HN.材料管理番号 = HJ.材料管理番号" & vbCrLf
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料取引先マスタ HT" & vbCrLf
    gSL_Select = gSL_Select & " ON  HT.材料管理番号 = HJ.材料管理番号" & vbCrLf
    gSL_Select = gSL_Select & " AND HT.取引先 = HJ.手配先" & vbCrLf
    gSL_Select = gSL_Select & " AND HT.受入 = HJ.受入" & vbCrLf
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY" & vbCrLf
    gSL_Select = gSL_Select & " ON  SY.区分名称 = '単位区分'" & vbCrLf
    gSL_Select = gSL_Select & " AND SY.値 = HN.手配単位" & vbCrLf
    gSL_Select = gSL_Select & " AND SY.区分タイプ = ''" & vbCrLf
    gSL_Select = gSL_Select & " WHERE HJ.内示年月 = '" & RTrim$(TGet_NJYM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " AND   HJ.計画年月 >= '" & RTrim$(TGet_YM) & "'" & vbCrLf
    
    If RTrim$(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & " AND   TR.取引先CD = '" & RTrim$(iH1_Torcd) & "'" & vbCrLf
    End If
    
'    gSL_Select = gSL_Select & " AND   TR.取引先区分 = 1 AND TR.契約先区分 = 0 AND TR.社内区分 = 0" & vbCrLf
    gSL_Select = gSL_Select & " AND   TR.取引先区分 = 1 AND TR.社内区分 = 0" & vbCrLf
    
    gSL_Select = gSL_Select & " GROUP BY HJ.材料管理番号, HJ.手配先, HJ.受入, HJ.直送先, HJ.計画年月, HJ.内示年月" & vbCrLf
    
'   ( 日程データ )
    gSL_Select = gSL_Select & " UNION" & vbCrLf
    
    gSL_Select = gSL_Select & " SELECT  HJ.材料管理番号, HJ.手配先, HJ.受入, HJ.直送先, HJ.納期, HJ.数量," & vbCrLf
    gSL_Select = gSL_Select & " '' 手配先名, " & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( HN.材質,       '' ) 材質," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( HN.寸法区分,    1 ) 寸法区分," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( HN.板厚,        0 ) 板厚," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( HN.幅,          0 ) 幅," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( HN.長さ,        0 ) 長さ," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( HN.材種コード, '' ) 材種コード," & vbCrLf
    gSL_Select = gSL_Select & " '' 手配単位名, '' 材料名, '' 背番号, '' サイクル, 0 収容数" & vbCrLf
    gSL_Select = gSL_Select & " FROM 材料発注内示テーブル HJ" & vbCrLf
    
    If opOutKB = 1 Then
        gSL_Select = gSL_Select & " INNER JOIN ( SELECT DISTINCT" & vbCrLf
        gSL_Select = gSL_Select & "                     ISNULL( HZ.内示年月, HN.内示年月 ) 内示年月," & vbCrLf
        gSL_Select = gSL_Select & "                     ISNULL( HZ.材料管理番号, HN.材料管理番号 ) 材料管理番号," & vbCrLf
        gSL_Select = gSL_Select & "                     ISNULL( HZ.手配先, HN.手配先 ) 手配先," & vbCrLf
        gSL_Select = gSL_Select & "                     ISNULL( HZ.受入, HN.受入 ) 受入," & vbCrLf
        gSL_Select = gSL_Select & "                     ISNULL( HZ.直送先, HN.直送先 ) 直送先" & vbCrLf
        gSL_Select = gSL_Select & "                     FROM 材料発注内示前回テーブル HZ" & vbCrLf
        gSL_Select = gSL_Select & "                     FULL OUTER JOIN 材料発注内示テーブル HN" & vbCrLf
        gSL_Select = gSL_Select & "                     ON    HN.内示年月 = HZ.内示年月" & vbCrLf
        gSL_Select = gSL_Select & "                       AND HN.計画年月 = HZ.計画年月" & vbCrLf
        gSL_Select = gSL_Select & "                       AND HN.材料管理番号 = HZ.材料管理番号" & vbCrLf
        gSL_Select = gSL_Select & "                       AND HN.手配先 = HZ.手配先" & vbCrLf
        gSL_Select = gSL_Select & "                       AND HN.受入 = HZ.受入" & vbCrLf
        gSL_Select = gSL_Select & "                       AND HN.直送先 = HZ.直送先" & vbCrLf
        gSL_Select = gSL_Select & "                       AND HN.納期 = HZ.納期" & vbCrLf
        gSL_Select = gSL_Select & "                       AND HN.数量 = HZ.数量" & vbCrLf
        gSL_Select = gSL_Select & "                     WHERE (HN.材料管理番号 Is Null Or HZ.材料管理番号 Is Null)" & vbCrLf
        gSL_Select = gSL_Select & "                       AND ISNULL( HZ.内示年月, '' )  = CASE WHEN HZ.材料管理番号 IS NULL THEN '' ELSE '" & RTrim$(TGet_NJYM) & "' END" & vbCrLf
        gSL_Select = gSL_Select & "                       AND ISNULL( HZ.計画年月, '' ) >= CASE WHEN HZ.材料管理番号 IS NULL THEN '' ELSE '" & RTrim$(TGet_YM) & "' END" & vbCrLf
        gSL_Select = gSL_Select & "                       AND ISNULL( HN.内示年月, '' )  = CASE WHEN HN.材料管理番号 IS NULL THEN '' ELSE '" & RTrim$(TGet_NJYM) & "' END" & vbCrLf
        gSL_Select = gSL_Select & "                       AND ISNULL( HN.計画年月, '' ) >= CASE WHEN HN.材料管理番号 IS NULL THEN '' ELSE '" & RTrim$(TGet_YM) & "' END" & vbCrLf
        gSL_Select = gSL_Select & "            ) BD" & vbCrLf
        gSL_Select = gSL_Select & " ON  BD.内示年月 = HJ.内示年月" & vbCrLf
        gSL_Select = gSL_Select & " And BD.材料管理番号 = HJ.材料管理番号" & vbCrLf
        gSL_Select = gSL_Select & " And BD.手配先 = HJ.手配先" & vbCrLf
        gSL_Select = gSL_Select & " And BD.受入 = HJ.受入" & vbCrLf
        gSL_Select = gSL_Select & " And BD.直送先 = HJ.直送先" & vbCrLf
    End If
    
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料マスタ HN" & vbCrLf
    gSL_Select = gSL_Select & " ON  HN.材料管理番号 = HJ.材料管理番号" & vbCrLf
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR" & vbCrLf
    gSL_Select = gSL_Select & " ON  TR.取引先CD = HJ.手配先" & vbCrLf
    gSL_Select = gSL_Select & " AND TR.取引先区分 = 1" & vbCrLf
    gSL_Select = gSL_Select & " WHERE HJ.内示年月 = '" & RTrim$(TGet_NJYM) & "'" & vbCrLf
    gSL_Select = gSL_Select & " AND   HJ.計画年月 = '" & RTrim$(TGet_YM) & "'" & vbCrLf
    
    If RTrim$(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & " AND   TR.取引先CD = '" & RTrim$(iH1_Torcd) & "'" & vbCrLf
    End If
    
'    gSL_Select = gSL_Select & " AND   TR.取引先区分 = 1 AND TR.契約先区分 = 0 AND TR.社内区分 = 0" & vbCrLf
    gSL_Select = gSL_Select & " AND   TR.取引先区分 = 1 AND TR.社内区分 = 0" & vbCrLf

'   ( 読み込み順 )
    gSL_Select = gSL_Select & " ORDER BY HJ.手配先, HJ.受入, 材質, 板厚, 幅, 長さ, 納期" & vbCrLf
    
    If Not ZNJRead(gSL_Select) Then
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
   CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*      メッセージ表示        *
'*******************************
Private Sub Timer2_Timer()
    CisFrm.MSGRelate
End Sub
'*------------------------------------------------*
'*      材　料　内　示　対　象　年　月　算　出      *
'*------------------------------------------------*
Private Function Zai_Nai_YM()
    
'## CHG-S 2010/02/24 ##
'    With CisFun
'        .Date_RsFormat = fYM
'        .Date_Add TGet_NJYM, M, CInt(YMAdd)
'
'        TGet_YM = .Date_Result
'        H1lb_YM = "( 対象年月 ： "
'        H1lb_YM = H1lb_YM & Mid$(TGet_YM, 1, 4) & "年 " & Mid$(TGet_YM, 5, 2) & "月 )"
'    End With

    If YMAdd = 0 Then
        TGet_YM = TGet_NJYM
    Else
        With CisFun
            .Date_RsFormat = fYM
            .Date_Add TGet_NJYM, M, CInt(YMAdd)
        
            TGet_YM = .Date_Result
        End With
    End If
    
    H1lb_YM = "( 対象年月 ： "
    H1lb_YM = H1lb_YM & Mid$(TGet_YM, 1, 4) & "年 " & Mid$(TGet_YM, 5, 2) & "月 )"
'## CHG-S 2010/02/24 ##

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
        lb_msg.Caption = "【 材料発注内示データ 印刷中 】"
        DoEvents
        With CisPrt
             Set .iPrtForm = CNJZ0320L
             .dpMSG = "材料発注内示表　印刷中"
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

    Dim MD1         As String
    Dim MD2         As String
    
    Dim OutStr      As String
    Dim OutFNo      As Long
    Dim OutCnt      As Long
    Dim NoCnt       As Long
    Dim RecCnt      As Long
    Dim End_Day     As Long
    Dim wYM         As String   ' 当月
    Dim wYMy        As String   ' 翌月
    Dim wYMyy       As String   ' 翌々月
    Dim w_MSu(2)    As Long     ' 月量
    Dim w_DSu(31)   As Long     ' 日量
    Dim w_DFlg      As Byte     ' 日量データフラグ
    
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
    
    lb_msg.Caption = "【 材料発注内示データ 出力中 】"
    Csv_Bar.Visible = True
    
    DoEvents
        
    OutFNo = FreeFile
    Open mCsvDir For Output As #OutFNo

    Call CisFrm.MousePT(11)
    
    With CisFrm
        Set .MeterPic = Csv_Bar
        .MeterMax = CisDB.RecordCount
        RecCnt = 0
    End With
        
'< 対象年月(内示年月)の末日を求める >
    With CisFun
        .Date_RsFormat = fDD
        .Date_End TGet_YM
        End_Day = .Val2(.Date_Result)
    End With
    
    gStr = RTrim$(TGet_YM)
    wYM = Mid(gStr, 3, 2) & "年 " & CisFun.RSetFld(Mid(gStr, 5, 2), 2, "##") & "月"
    With CisFun
'       < 翌月 >
        .Date_RsFormat = fYM
        .Date_Add gStr & "01", M, 1, LastDate
        gStr = .Date_Result
        wYMy = Mid(gStr, 3, 2) & "年 " & CisFun.RSetFld(Mid(gStr, 5, 2), 2, "##") & "月"
        
'       < 翌々月 >
        .Date_RsFormat = fYM
        .Date_Add gStr & "01", M, 1, LastDate
        gStr = .Date_Result
        wYMyy = Mid(gStr, 3, 2) & "年 " & CisFun.RSetFld(Mid(gStr, 5, 2), 2, "##") & "月"
    End With
        

'[ 大見出し ]       (見出し)
    MD1 = "内示年月,仕入先,仕入先名"

'[ 明細見出し ]     (見出し)
'-------------------------------------------------------------2010/02/09 Update Start
'    MD2 = "№,材質,厚さ,幅,長さ,材種コード,納入先,受入,背番号,納入回数,収容数,"
'    MD2 = MD2 & "材料名,"
'    MD2 = MD2 & wYM & "仕掛内示,"
'    MD2 = MD2 & wYMy & "仕掛内示,"
'    MD2 = MD2 & wYMyy & "仕掛内示,"
    MD2 = "№,材料名,材種コード,納入先,受入,背番号,納入回数,収容数,"
    MD2 = MD2 & wYM & "仕掛内示,"
    MD2 = MD2 & wYMy & "仕掛内示,"
    MD2 = MD2 & wYMyy & "仕掛内示,"
    MD2 = MD2 & "単位,"
'-------------------------------------------------------------2010/02/09 Update End
    For gInt = 1 To End_Day
        MD2 = MD2 & gInt & "日"
        If gInt <> End_Day Then
           MD2 = MD2 & ","
        End If
    Next gInt
    
    
    With ZNJ
        OutCnt = 0:     PgBk_Torcd = "":    OutStr = ""
        Erase w_MSu:    Erase w_DSu:        NoCnt = 0
        
        Do Until Not ZNJ_RDSTS
            
            If RTrim$(.手配先) <> PgBk_Torcd Then
                If RTrim$(MD1) <> "" Then Print #OutFNo, MD1: MD1 = ""
                OutStr = Format(RTrim$(TGet_YM), "@@@@/@@") & ","
                OutStr = OutStr & .手配先 & ","
                OutStr = OutStr & .手配先名
                Print #OutFNo, OutStr
                If RTrim$(MD2) <> "" Then Print #OutFNo, MD2: MD2 = ""
                
                PrtK_Hinbn = RTrim$(.材料管理番号)
                PrtK_Torcd = RTrim$(.手配先)
                PrtK_Ukeir = RTrim$(.受入)
                PgBk_Torcd = RTrim$(.手配先)
                OutStr = "": NoCnt = 0
            End If
            
            If CisFun.Val2(.納期) <= 2 Then
'               < 月量情報 >
                If RTrim$(OutStr) = "" Then
                    
                    OutCnt = OutCnt + 1
                    NoCnt = NoCnt + 1
                    
                    OutStr = NoCnt & ","
'-------------------------------------------------------2010/02/09 Delete Start
'                    OutStr = OutStr & RTrim$(.材質) & ","
'
'                    If .寸法区分 = 0 Then
'                        OutStr = OutStr & RTrim$(.板厚) & ","
'                        OutStr = OutStr & RTrim$(.幅) & ","
'                        If .長さ = 0 Then
'                            OutStr = OutStr & "C,"
'                        Else
'                            OutStr = OutStr & RTrim$(.長さ) & ","
'                        End If
'                    Else
'                        OutStr = OutStr & ",,,"
'                    End If
'-------------------------------------------------------2010/02/09 Delete End
                    '-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* CHG-S 2010/02/24
                    'OutStr = OutStr & .材料名 & ","             '2010/02/09 Insert
                    OutStr = OutStr & Replace(.材料名, ",", ".") & ","
                    '-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* CHG-E 2010/02/24
                    OutStr = OutStr & RTrim$(.材種コード) & ","
                    
                    If RTrim$(.直送先) = "" Then
                        OutStr = OutStr & ","
                        OutStr = OutStr & ","
                    Else
                        OutStr = OutStr & RTrim(.直送先) & ","
                        OutStr = OutStr & ","
                    End If
                    
                    OutStr = OutStr & RTrim(.背番号) & ","
                    If Trim(.サイクル) <> "" Then
                        OutStr = OutStr & "'" & Mid(.サイクル, 1, 1) & "-" & _
                                                Mid(.サイクル, 2, 2) & "-" & _
                                                Mid(.サイクル, 4, 2) & ","
                    Else
                        OutStr = OutStr & ","
                    End If
                    OutStr = OutStr & .収容数 & ","
'                    OutStr = OutStr & .材料名 & ","            '2010/02/09 Delete
                End If
            
'               ( 月量退避 )
               '## CHG-S 2010/02/24 ##
'                w_MSu(CisFun.Val2(.納期)) = .数量
                w_MSu(CisFun.Val2(.納期 - YMAdd)) = .数量
               '## CHG-S 2010/02/24 ##
            Else
'               < 日量情報 >
                gInt = CisFun.Val2(Mid$(.納期, 7, 2))
                w_DSu(gInt) = .数量
            End If
            
            RecCnt = RecCnt + 1
            CisFrm.MeterUpdate RecCnt
            
            Call ZNJReadNext
        
            If Not ZNJ_RDSTS Or _
              PrtK_Hinbn <> .材料管理番号 Or _
              PrtK_Torcd <> RTrim$(.手配先) Or _
              PrtK_Ukeir <> RTrim$(.受入) Then

'               ( 月量 )
                For gInt = 0 To 2
                    OutStr = OutStr & w_MSu(gInt) & ","
                Next gInt

                OutStr = OutStr & .手配単位名 & ","         '2010/02/09 Insert

'               ( 日量 )
                For gInt = 1 To End_Day - 1
                    OutStr = OutStr & w_DSu(gInt) & ","
                Next gInt
                OutStr = OutStr & w_DSu(End_Day)
                
'               ( 出力 )
                Print #OutFNo, OutStr
                
                PrtK_Hinbn = RTrim$(.材料管理番号)
                PrtK_Torcd = RTrim$(.手配先)
                PrtK_Ukeir = RTrim$(.受入)
                Erase w_MSu: Erase w_DSu
                OutStr = ""
            End If
        Loop
    End With

CsvPut_Ed:
    Call ZNJReadNext

    Close #OutFNo
    SaveSetting appname:=gRegKey, section:="CSV", _
                    Key:=mCsvKey, setting:=RTrim(lb_PrtDev)
    
    Call CisFrm.MousePT(1)
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "   出力処理が終了しました。    "
        .MB_MSG(4) = "   　 " & "出力先    = " & mCsvDir
        .MB_MSG(5) = "   　 " & "出力件数 = " & StrConv(Format(OutCnt, "##,##0"), vbWide) & " 件 "
        .MB_Title = "出力確認"
        .MB_Button = OK
        .MBOX
    End With
    CsvPut = True
    
End Function


