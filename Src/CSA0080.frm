VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CSA0080 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "在庫マスタ(部品)"
   ClientHeight    =   11190
   ClientLeft      =   90
   ClientTop       =   1500
   ClientWidth     =   15360
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
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin VB.PictureBox B2_Area3 
      BackColor       =   &H00808080&
      Height          =   1800
      Left            =   4440
      ScaleHeight     =   1740
      ScaleWidth      =   6900
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   6960
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'なし
         Height          =   1710
         Left            =   0
         ScaleHeight     =   1710
         ScaleWidth      =   6870
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   6870
         Begin Cis3D_v60.CIS3D PB_CsvEnt 
            Height          =   450
            Left            =   1560
            Top             =   1740
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   794
            BackColor       =   -2147483633
            Caption         =   "開始"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cFont3D         =   2
            cAlingnment     =   7
         End
         Begin VB.PictureBox Csv_Bar 
            Height          =   345
            Left            =   120
            ScaleHeight     =   285
            ScaleWidth      =   6555
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   810
            Width           =   6615
         End
         Begin Cis3D_v60.CIS3D UC_3D28 
            Height          =   615
            Index           =   1
            Left            =   240
            Top             =   120
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
               TabIndex        =   19
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
                  Caption         =   "【データ出力　処理中】"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   12
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
            Left            =   240
            Top             =   1260
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
            Left            =   6240
            Top             =   1260
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
         Begin Cis3D_v60.CIS3D PB_CsvCan 
            Height          =   450
            Left            =   4080
            Top             =   1740
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   794
            BackColor       =   -2147483633
            Caption         =   "取消"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cFont3D         =   2
            cAlingnment     =   7
         End
      End
   End
   Begin VB.PictureBox B2_Area1 
      Height          =   8970
      Left            =   240
      ScaleHeight     =   8910
      ScaleWidth      =   14820
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2070
      Width           =   14880
      Begin VB.PictureBox B2_Area2 
         BackColor       =   &H00808080&
         Height          =   1710
         Left            =   5130
         ScaleHeight     =   1650
         ScaleWidth      =   9180
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1290
         Width           =   9240
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'なし
            Height          =   1620
            Left            =   15
            ScaleHeight     =   1620
            ScaleWidth      =   1065
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   30
            Width           =   1065
            Begin Cis3D_v60.CIS3D B2lb_KjJikan 
               Height          =   330
               Left            =   0
               Top             =   990
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   582
               BackColor       =   16777152
               Caption         =   "999.999"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cAlingnment     =   4
            End
            Begin Cis3D_v60.CIS3D CIS3D2 
               Height          =   285
               Left            =   0
               Top             =   1335
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   503
               BackColor       =   -2147483633
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
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cAlingnment     =   1
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   330
               Left            =   0
               Top             =   660
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   582
               BackColor       =   -2147483633
               Caption         =   "基準時間"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   60
            End
            Begin Cis3D_v60.CIS3D B2lb_KoteiSu 
               Height          =   330
               Left            =   0
               Top             =   330
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   582
               BackColor       =   16777152
               Caption         =   "9999.0"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cAlingnment     =   4
            End
            Begin Cis3D_v60.CIS3D CIS3D7 
               Height          =   330
               Left            =   0
               Top             =   0
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   582
               BackColor       =   -2147483633
               Caption         =   "工程数"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   60
            End
         End
         Begin vsFlexLib.vsFlexArray VsGrid2 
            Height          =   1680
            Left            =   0
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   9210
            _Version        =   196608
            _ExtentX        =   16245
            _ExtentY        =   2963
            _StockProps     =   228
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "基準時間|内示　|09月内示数|09月工数|09月ﾊﾟﾝﾁ数|10月内示数|10月工数|10月ﾊﾟﾝﾁ数|11月内示数|11月工数|11月ﾊﾟﾝﾁ数|"
            Rows            =   4
            Cols            =   12
            FixedCols       =   2
            BackColor       =   16777152
            BackColorFixed  =   12640511
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   1
            RowHeightMin    =   330
            SelectionMode   =   1
         End
      End
      Begin vsFlexLib.vsFlexArray VsGrid1 
         Height          =   8880
         Left            =   0
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   14820
         _Version        =   196608
         _ExtentX        =   26141
         _ExtentY        =   15663
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   $"CSA0080.frx":0000
         Rows            =   26
         Cols            =   25
         FixedCols       =   6
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   330
         GridLines       =   2
         SelectionMode   =   1
         BackColorAlternate=   12648447
      End
      Begin Cis3D_v60.CIS3D CIS3D_Color 
         Height          =   375
         Left            =   120
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BackColor       =   12648384
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
      End
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   12510
      ScaleHeight     =   1320
      ScaleWidth      =   2640
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   540
      Width           =   2700
      Begin CisBtn_60.CisBtn PB_DTL 
         Height          =   585
         Left            =   900
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   12
         BtnCaption      =   "詳細"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16711935
         BtnShadowColor  =   16777215
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
         PFForeColor     =   12648447
         PFShadowColor   =   0
         PFBackColor     =   16744703
      End
      Begin CisBtn_60.CisBtn PB_CSV 
         Height          =   585
         Left            =   1770
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   14
         BtnCaption      =   "出力"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   192
         BtnShadowColor  =   16777215
         PFCaption       =   "( F10 )"
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
         PFBackColor     =   192
      End
      Begin CisBtn_60.CisBtn PB_LOOK 
         Height          =   585
         Left            =   30
         Top             =   60
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   30
         Top             =   690
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
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   900
         Top             =   690
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
         Left            =   1770
         Top             =   690
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
         BtnType         =   1
         BtnCaption      =   "終了"
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
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   -30
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   420
      Width           =   855
   End
   Begin VB.PictureBox B1_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9120
      Left            =   120
      ScaleHeight     =   9060
      ScaleWidth      =   15030
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1980
      Width           =   15090
   End
   Begin VB.PictureBox H1_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   120
      ScaleHeight     =   1320
      ScaleWidth      =   10425
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   540
      Width           =   10485
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   780
         Index           =   5
         Left            =   60
         Top             =   30
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1376
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   60
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   360
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   1380
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
      End
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   510
         Left            =   4275
         Top             =   810
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "品　　番"
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
         cAlingnment     =   1
         cPositionX      =   100
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   1215
            TabIndex        =   4
            Top             =   75
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
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   780
         Left            =   1680
         Top             =   30
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "部　　門"
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   90
            TabIndex        =   1
            Top             =   360
            Width           =   960
            _ExtentX        =   1693
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
            Text            =   "XXXXXXX"
            MaxLength       =   7
            IMEMode         =   2
         End
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1065
            Top             =   360
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   635
            BackColor       =   16777152
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   510
         Left            =   60
         Top             =   810
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "機　　械"
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
         cAlingnment     =   1
         cPositionX      =   100
         Begin CisText_V60.CisText iH1_Kikai 
            Height          =   375
            Left            =   1125
            TabIndex        =   3
            Top             =   75
            Width           =   510
            _ExtentX        =   900
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
            Text            =   "XXX"
            MaxLength       =   3
            IMEMode         =   2
         End
         Begin Cis3D_v60.CIS3D H1lb_Kikai 
            Height          =   360
            Left            =   1635
            Top             =   75
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   635
            BackColor       =   16777152
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   1305
         Index           =   1
         Left            =   8805
         Top             =   15
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   2302
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
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   60
         Begin VB.OptionButton OpH1_Kousu 
            BackColor       =   &H00C0C0C0&
            Caption         =   "基準工数"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   495
            Width           =   1290
         End
         Begin VB.OptionButton OpH1_Kousu 
            BackColor       =   &H00C0C0C0&
            Caption         =   "目標工数"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   885
            Width           =   1290
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Left            =   5520
         Top             =   30
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "機　械  №"
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
         Begin CisText_V60.CisText iH1_KikaiNo 
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   375
            Width           =   510
            _ExtentX        =   900
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
            Text            =   "XXX"
            MaxLength       =   3
            IMEMode         =   2
         End
         Begin Cis3D_v60.CIS3D H1lb_KikaiNo 
            Height          =   360
            Left            =   630
            Top             =   375
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   635
            BackColor       =   16777152
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   420
   End
   Begin VB.Timer Timer2 
      Left            =   690
      Top             =   420
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 内示・生産負荷一覧  】"
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
         Left            =   14220
         Top             =   90
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
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   11670
         Top             =   60
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
      Begin Cis3D_v60.CIS3D UC_3D32 
         Height          =   255
         Left            =   8640
         Top             =   60
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   450
         BackColor       =   14737632
         Caption         =   "抽出件数[              ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   14737632
         cBoderColor2    =   14737632
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cBoderStyle     =   2
         cPositionX      =   50
         Begin Cis3D_v60.CIS3D lb_Kensu 
            Height          =   285
            Left            =   1095
            Top             =   0
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "##,### / ##,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   14737632
            cBoderColor2    =   14737632
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "CSA0080"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   内示・生産負荷一覧
'**       フォームID    :   CSA0080
'**       処理概要      :
'**
'**       作  成  日    :   2010/01/19  By CIS SIMPR-A
'**       変  更  日    :   2013/05/10  By CIS 1.表示順を機械NO(機械NO名)順
'**                                            2.基準工数/目標工数選択時にｸﾞﾘｯﾄﾞﾀｲﾄﾙを変更
'**                                            3.条件に機械NOを追加
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替

    Dim CisVsGrid       As New CisVsGrid3
    Dim CisVsGrid2      As New CisVsGrid3
   
    Dim mYM_P1          As String       ' 翌月
    Dim mYM_P2          As String       ' 翌々月
    
    Dim mCsvDir         As String
    Dim mCsvKey         As String
    Dim mCsvName        As String
    
    Dim mTtl_Kensu      As Long
    
    Dim ActObj          As Object
    Dim ActRow          As Integer
    
    Dim wsDate          As String
    Dim SCol()          As Long
    
    Dim wDate           As String
    Dim wEndDay         As Byte
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub form_load()
'   #------------------#
'   # 二 重 起 動 防 止 #
'   #------------------#
    If Not Dupli_Start Then End
'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
    If Not IniGet Then End
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
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
    
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = VsGrid1
        .NoSpaceCol = 5
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 15
        Set .SelectDispObj = CIS3D_Color
        .InitGet = False
        .Init
    End With
    Call GridInit
    '+-----------------------+
    '+  グリッド2 初期設定
    '+-----------------------+
    With CisVsGrid2
        Set .GridObj = VsGrid2
        .NoSpaceCol = 0
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
      ''.SelectCol = 11
        .InitGet = False
        .Init
    End With
    Call Grid2Init

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    OpH1_Kousu(0).Value = True
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
' 年月初期値
    wsDate = Format(Now(), "yyyymmdd")
    iH1_YM = Mid(wsDate, 1, 6)
   
' .CSVファイル出力先
    mCsvName = "内示・生産負荷一覧"
    mCsvKey = "CSA0080"
    mCsvDir = GetSetting(appname:=gRegKey, section:="CSV", _
                       Key:=mCsvKey, Default:=Mid(CisFun.WinRoot, 1, 3) & mCsvName & ".csv")
    
    FormAct = Not FormAct
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
           Case vbKeyF7:       If PB_LOOK.Visible Then Call PB_Look_Click   '【検索】
           Case vbKeyF8:       If PB_DTL.Visible Then Call PB_DTL_Click     '【詳細】
           Case vbKeyF10:      If PB_CSV.Visible Then Call PB_CSV_Click     '【出力】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click     '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click     '【終了】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                       '【入力】
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
    Cancel = 1
    Call PB_END_Click
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'+---------------------+
'+   ﾃﾞｰﾀﾍﾞｰｽ切断
'+---------------------+
    Call CisDB.DBDISConnect
'+---------------------+
'+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
'+---------------------+
    Call CisFrm.UnLoadDisp
    End
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
    
    If ProcHB = "B3" Then
        Call PB_CsvCan_Click
        GoTo PB_CAN_Ed
    End If
    
   If ProcHB = "B2" Then
      ProcHB = "B1"
      Call DispChange(ProcHB)
      Call vsGrid1_Select
      Call VsGrid1.SetFocus
      GoTo PB_CAN_Ed
   End If
   
   If ProcHB <> "H1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      iH1_YM.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+------------------------------+
'+    詳  細  キ ー(F8)         +
'+------------------------------+
Private Sub PB_DTL_Click()
    Dim wCnt    As Integer
    Dim wKousu  As Double
    Dim wJikan  As Double
    Dim wAtai   As Double
    
    wCnt = VsGrid1.Row
    wCnt = VsGrid1.RowSel
    VsGrid1.RowSel = VsGrid1.Row
    
    Call CisFrm.MousePT(11)
    
    If ProcHB = "B1" Then
        If VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid.NoSpaceCol) = "" Then GoTo PB_DTL_Ed
        Call vsGrid1_Select("ON")
        ProcHB = "B2"
        Call DispChange(ProcHB)
   
        VsGrid2.Clear
        With VsGrid2
           '.Left = 1050
           '.Width = .Width - .Left
            .TextMatrix(1, 0) = "内示"
            .TextMatrix(2, 0) = "実績"
            .TextMatrix(3, 0) = "差"
           '.FixedCols = 1
           '.Cols = 11
        End With
   
        If VsGrid1.CellTop > (CisVsGrid.FixedHeight * 21) Then
           B2_Area2.Top = VsGrid1.CellTop - (CisVsGrid.FixedHeight * 4) - 30
        Else
           B2_Area2.Top = VsGrid1.CellTop - 30
        End If
          
        For wCnt = 1 To 9
            VsGrid2.TextMatrix(0, wCnt) = VsGrid1.TextMatrix(0, wCnt + 5)
            VsGrid2.TextMatrix(1, wCnt) = VsGrid1.TextMatrix(VsGrid1.Row, wCnt + 5)
        Next
        B2lb_KoteiSu = VsGrid1.TextMatrix(VsGrid1.Row, 4): wKousu = CisFun.Val2(B2lb_KoteiSu)
        B2lb_KjJikan = VsGrid1.TextMatrix(VsGrid1.Row, 5): wJikan = CisFun.Val2(B2lb_KjJikan)
   
        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT SUBSTRING(JK0.実績日,1,6) AS 実績日, JK0.実績先, JK0.受入, JK0.品番, JK0.工順" & Chr(13)
        gSL_Select = gSL_Select & "      ,ISNULL(JK1.実績数,0) AS 実績数, ISNULL(JK2.実績数,0) AS 実績枚数, ISNULL(JK3.実績数,0) AS 実績端数" & Chr(13)
        gSL_Select = gSL_Select & " FROM  実績管理テーブル JK0" & Chr(13)
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUM(J1.実績数) AS 実績数, SUBSTRING(J1.実績日,1,6) AS 実績日, J1.実績先, J1.受入, J1.品番, J1.工順" & Chr(13)
        gSL_Select = gSL_Select & "                         FROM  実績管理テーブル J1" & Chr(13)
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J1.実績日,1,6), J1.実績先, J1.受入, J1.品番, J1.工順) JK1" & Chr(13)
        gSL_Select = gSL_Select & "        ON  SUBSTRING(JK1.実績日,1,6) = '" & Trim(iH1_YM) & "'" & Chr(13)
        gSL_Select = gSL_Select & "        AND JK1.実績先 = JK0.実績先" & Chr(13)
        gSL_Select = gSL_Select & "        AND JK1.受入   = JK0.受入" & Chr(13)
        gSL_Select = gSL_Select & "        AND JK1.品番   = JK0.品番" & Chr(13)
        gSL_Select = gSL_Select & "        AND JK1.工順   = JK0.工順" & Chr(13)
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUM(J2.実績数) AS 実績数, SUBSTRING(J2.実績日,1,6) AS 実績日, J2.実績先, J2.受入, J2.品番, J2.工順" & Chr(13)
        gSL_Select = gSL_Select & "                         FROM  実績管理テーブル J2" & Chr(13)
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J2.実績日,1,6), J2.実績先, J2.受入, J2.品番, J2.工順) JK2" & Chr(13)
        gSL_Select = gSL_Select & "        ON  SUBSTRING(JK2.実績日,1,6) = '" & Trim(mYM_P1) & "'" & Chr(13)
        gSL_Select = gSL_Select & "        AND JK2.実績先 = JK0.実績先" & Chr(13)
        gSL_Select = gSL_Select & "        AND JK2.受入   = JK0.受入" & Chr(13)
        gSL_Select = gSL_Select & "        AND JK2.品番   = JK0.品番" & Chr(13)
        gSL_Select = gSL_Select & "        AND JK2.工順   = JK0.工順" & Chr(13)
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUM(J3.実績数) AS 実績数, SUBSTRING(J3.実績日,1,6) AS 実績日, J3.実績先, J3.受入, J3.品番, J3.工順" & Chr(13)
        gSL_Select = gSL_Select & "                         FROM  実績管理テーブル J3" & Chr(13)
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J3.実績日,1,6), J3.実績先, J3.受入, J3.品番, J3.工順) JK3" & Chr(13)
        gSL_Select = gSL_Select & "        ON  SUBSTRING(JK1.実績日,3,6) = '" & Trim(mYM_P2) & "'" & Chr(13)
        gSL_Select = gSL_Select & "        AND JK3.実績先 = JK0.実績先" & Chr(13)
        gSL_Select = gSL_Select & "        AND JK3.受入   = JK0.受入" & Chr(13)
        gSL_Select = gSL_Select & "        AND JK3.品番   = JK0.品番" & Chr(13)
        gSL_Select = gSL_Select & "        AND JK3.工順   = JK0.工順" & Chr(13)
        gSL_Select = gSL_Select & " WHERE (SUBSTRING(JK0.実績日,1,6) = '" & Trim(iH1_YM) & "' OR SUBSTRING(JK0.実績日,1,6) = '" & Trim(mYM_P1) & "' OR SUBSTRING(JK0.実績日,1,6) = '" & Trim(mYM_P2) & "')" & Chr(13)
        gSL_Select = gSL_Select & "   AND JK0.実績先 = '" & Trim(VsGrid1.TextMatrix(VsGrid1.Row, 16)) & "'" & Chr(13)
        gSL_Select = gSL_Select & "   AND JK0.受入   = '" & Trim(VsGrid1.TextMatrix(VsGrid1.Row, 20)) & "'" & Chr(13)
        gSL_Select = gSL_Select & "   AND JK0.品番   = '" & Trim(VsGrid1.TextMatrix(VsGrid1.Row, 19)) & "'" & Chr(13)
        gSL_Select = gSL_Select & "   AND JK0.工順   =  " & Trim(VsGrid1.TextMatrix(VsGrid1.Row, 21)) & " " & Chr(13)
        gSL_Select = gSL_Select & " GROUP BY SUBSTRING(JK0.実績日,1,6), JK0.実績先, JK0.受入, JK0.品番, JK0.工順, JK1.実績数, JK2.実績数, JK3.実績数" & Chr(13)
        
        With VsGrid2
            If JKTRead(gSL_Select) Then
                .TextMatrix(2, 1) = Format(JKT.実績数, "########")                                                  ' 実績：当月
                wAtai = JKT.実績数 * wJikan: If wAtai <> 0 Then .TextMatrix(2, 2) = Format(wAtai, "#####0.0")
                .TextMatrix(2, 3) = Format(JKT.実績数 * wKousu, "#########")
                .TextMatrix(2, 4) = Format(JKT.実績枚数, "########")                                                ' 実績：翌月
                wAtai = JKT.実績枚数 * wJikan: If wAtai <> 0 Then .TextMatrix(2, 5) = Format(wAtai, "#####0.0")
                .TextMatrix(2, 6) = Format(JKT.実績枚数 * wKousu, "#########")
                .TextMatrix(2, 7) = Format(JKT.実績端数, "########")                                                ' 実績：翌々月
                wAtai = JKT.実績端数 * wJikan: If wAtai <> 0 Then .TextMatrix(2, 8) = Format(wAtai, "#####0.0")
                .TextMatrix(2, 9) = Format(JKT.実績端数 * wKousu, "#########")
            End If
            .TextMatrix(3, 1) = Format(CisFun.Val2(.TextMatrix(1, 1)) - CisFun.Val2(.TextMatrix(2, 1)), "########") ' 差：当月
            wAtai = CisFun.Val2(.TextMatrix(1, 2)) - CisFun.Val2(.TextMatrix(2, 2))
            If wAtai <> 0 Then .TextMatrix(3, 2) = Format(wAtai, "#####0.0")
            .TextMatrix(3, 3) = Format(CisFun.Val2(.TextMatrix(1, 3)) - CisFun.Val2(.TextMatrix(2, 3)), "#########")
            .TextMatrix(3, 4) = Format(CisFun.Val2(.TextMatrix(1, 4)) - CisFun.Val2(.TextMatrix(2, 4)), "########") ' 差：翌月
            wAtai = CisFun.Val2(.TextMatrix(1, 5)) - CisFun.Val2(.TextMatrix(2, 5))
            If wAtai <> 0 Then .TextMatrix(3, 5) = Format(wAtai, "#####0.0")
            .TextMatrix(3, 6) = Format(CisFun.Val2(.TextMatrix(1, 6)) - CisFun.Val2(.TextMatrix(2, 6)), "#########")
            .TextMatrix(3, 7) = Format(CisFun.Val2(.TextMatrix(1, 7)) - CisFun.Val2(.TextMatrix(2, 7)), "########") ' 差：翌々月
            wAtai = CisFun.Val2(.TextMatrix(1, 8)) - CisFun.Val2(.TextMatrix(2, 8))
            If wAtai <> 0 Then .TextMatrix(3, 8) = Format(wAtai, "#####0.0")
            .TextMatrix(3, 9) = Format(CisFun.Val2(.TextMatrix(1, 9)) - CisFun.Val2(.TextMatrix(2, 9)), "#########")
        End With
        GoTo PB_DTL_Ed
    End If

PB_DTL_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+------------------------------+
'+    出  力  キ ー(F10)        +
'+------------------------------+
Private Sub PB_CSV_Click()
    Call CisFrm.MousePT(11)
    
    ProcHB = "B3"
    VsGrid1.Enabled = False
    PB_ENT.Visible = True: PB_ENT.BtnCaption = "開始"
  ''PB_CAN.Visible = False
    PB_DTL.Visible = False
    PB_CSV.Visible = False
    
    Csv_Bar.Cls: Csv_Bar.BackColor = PB_CsvEnt.BackColor
    B2_Area3.Visible = True
    lb_msg = "内示・生産負荷一覧 CSV出力を行ないます"
    lb_PrtDev = mCsvDir
PB_CSV_Ed:
    Call CisFrm.MousePT(1)
End Sub
Private Sub PB_CsvEnt_Click()
    If Not CsvPut Then Exit Sub
    Call PB_CsvCan_Click
End Sub
Private Sub PB_CsvCan_Click()
    B2_Area3.Visible = False: PB_ENT.BtnCaption = "入力"
    ProcHB = "B1"
    PB_ENT.Visible = False
    PB_CAN.Visible = True
    PB_DTL.Visible = True
    PB_CSV.Visible = True
    VsGrid1.Enabled = True
    Call VsGrid1.SetFocus
End Sub

'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_LOOK.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_LOOK.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_LOOK.Tag = "iH1_Kikai" Then Call Look_Kikai: Exit Sub
    If PB_LOOK.Tag = "iH1_KikaiNo" Then Call Look_KikaiNO: Exit Sub             '2013/05/10
End Sub
'+----------------------------+
'+       品番 検索              +
'+----------------------------+
Private Sub Look_Hinbn()
''''    RV_Left = 0
''''    RV_Top = 0
''''
''''    CKK0030.Show vbModal
''''    Unload CKK0030
''''    Set CKK0030 = Nothing
''''
''''    If RV_Rtn Then
''''       iH1_Hinbn = RV_Hinbn
''''    End If
    
    RV_Left = 0
    RV_Top = 0
        
    RVI_Torcd = Trim(iH1_Torcd)
    RVI_Hinmoku = 1 ' 0:受注／1:発注／9:他
    RVI_Tehai = 2   ' 2:社内／9:空白
    RVI_SK = 9      ' 0:かんばん／1:指示
    RV_Call = ""
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Hinbn = RV_Hinbn
        iH1_Torcd = RV_TorcdK: Call iH1_Torcd_LostFocus
    End If
End Sub
'+---------------------+
'+    取引先検索        +
'+---------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 1                '   発注
    RV_TorKb2 = 1               '   社内

    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
       iH1_Torcd = RV_Torcd
       Call iH1_Torcd_LostFocus
    End If
End Sub
'+----------------------------+
'+       機械検索             +
'+----------------------------+
Private Sub Look_Kikai()
    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "機械"
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
       iH1_Kikai = RV_Code          ' 値
       H1lb_Kikai = RV_Kubun        ' 値名称
    End If
End Sub
'============================= 2013/05/10 Start
'+----------------------------+
'+       機械NO検索             +
'+----------------------------+
Private Sub Look_KikaiNO()
    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "機械NO"
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
       iH1_KikaiNo = RV_Code          ' 値
       H1lb_KikaiNo = RV_Kubun        ' 値名称
    End If
End Sub
'============================= 2013/05/10 End

'+-----------------------------+
'+  iH1_YM
'+-----------------------------+
Private Sub iH1_YM_GotFocus()
    PB_LOOK.Visible = False
' *ADD* 2004/04/14
    DoEvents
End Sub
'+-----------------------------+
'+  iH1_Hinbn
'+-----------------------------+
Private Sub iH1_Hinbn_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_Torcd
'+-----------------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    If RTrim(iH1_Torcd) <> "" Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT 略称 FROM  取引先マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先CD = '" & Trim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 取引先区分 = 1 AND 社内区分 = 1"          ' 発注＋社内
        If TRMRead(gSL_Select) Then H1lb_Tornm = Trim(TRM.略称) Else H1lb_Tornm = ""
    Else
        H1lb_Tornm = ""
    End If
End Sub
'+-----------------------------+
'+  iH1_Kikai
'+-----------------------------+
Private Sub iH1_Kikai_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_Kikai_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    
    If Trim(iH1_Kikai) <> "" Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT 値名称 FROM  名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = '機械'"
        gSL_Select = gSL_Select & "   AND 値 = '" & Trim(iH1_Kikai) & "'"
        If SYMRead(gSL_Select) Then H1lb_Kikai = Trim(SYM.値名称) Else H1lb_Kikai = ""
    Else
        H1lb_Kikai = ""
    End If
End Sub
'================================= 2013/05/10 Start
'+-----------------------------+
'+  iH1_KikaiNo
'+-----------------------------+
Private Sub iH1_KikaiNo_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_KikaiNo_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    
    If Trim(iH1_KikaiNo) <> "" Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT 値名称 FROM  名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = '機械NO'"
        gSL_Select = gSL_Select & "   AND 値 = '" & Trim(iH1_KikaiNo) & "'"
        If SYMRead(gSL_Select) Then H1lb_KikaiNo = Trim(SYM.値名称) Else H1lb_KikaiNo = ""
    Else
        H1lb_KikaiNo = ""
    End If
End Sub
'================================= 2013/05/10 End

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
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
    Dim B2Mode      As Boolean
    Dim B2Color     As Long
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
    B2Mode = False
    B2Color = gPLostSel
    
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
       Case "B2"
          B2Mode = True
          B2Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B1_Area1.Enabled = B1Mode
    VsGrid1.Enabled = B1Mode
    B2_Area1.BackColor = B2Color
    B2_Area2.Visible = B2Mode
    
    PB_LOOK.Visible = H1Mode
    PB_DTL.Visible = B1Mode
    PB_CSV.Visible = B1Mode
    PB_CAN.Visible = Not H1Mode
    PB_ENT.Visible = H1Mode
    PB_END.Visible = H1Mode
    
    If H1Mode Then lb_Kensu.Caption = ""
    
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
'                .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        CisVsGrid.Clear
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "B3" Then
        Call PB_CsvEnt_Click
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
            
      ''DoEvents
        ProcHB = "B1"
        Call DispChange(ProcHB)
        
        ActRow = 1
        Call VsGrid1_Click
      'VsGrid1.SetFocus
        DoEvents
        GoTo ReturnPress_Ed
    End If

ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
' 年月度
    If RTrim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    年月度を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_YM.SetFocus
       Exit Function
    End If
' 部署
    If Trim(iH1_Torcd) <> "" And Trim(H1lb_Tornm) = "" Then
        With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "    部門  未登録               "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
        End With
        iH1_Torcd.SetFocus
        Exit Function
    End If
' 機械
    If Trim(iH1_Kikai) <> "" And RTrim(H1lb_Kikai) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    機械  未登録               "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Kikai.SetFocus
       Exit Function
    End If
' 品番
    If RTrim(iH1_Hinbn) <> "" Then
        gSL_Select = "SELECT * FROM 品番取引先マスタ"
        gSL_Select = gSL_Select & " WHERE 品番   = '" & Trim(Replace(iH1_Hinbn, "-", "")) & "'"
        If Trim(iH1_Torcd) <> "" Then _
        gSL_Select = gSL_Select & "   AND 取引先 = '" & Trim(iH1_Torcd) & "'"
        
        If Not HNMRead(gSL_Select, 1) Then
            With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    品番マスタ  未登録            "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
            End With
            iH1_Hinbn.SetFocus
            Exit Function
        End If
    End If
    
' BODY部 セット
    If Not BodySet("") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     該当データがありません                "
            .MB_Title = "抽出結果"
            .MB_Button = OK
            .MBOX
        End With
        Exit Function
    End If
    
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
'
    Body1Chk = True
End Function
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Function BodySet(SetType As String) As Boolean
        BodySet = False
        Static wsBDFlg  As Boolean

    '======================== 2013/05/10 Start
    If OpH1_Kousu(0) Then
        VsGrid1.TextMatrix(0, 5) = "基準工数"
    Else
        VsGrid1.TextMatrix(0, 5) = "目標工数"
    End If
    '======================== 2013/05/10 End

'
    With CisFun
        .Date_RsFormat = fYM
      ''.Date_Add RTrim(iH1_YM), M, 2, ResultDate   ' .Date_Add RTrim(iH1_YM), M, -12, ResultDate
      ''gStr = .Date_Result
        .Date_Add RTrim(iH1_YM), M, 1, ResultDate
        mYM_P1 = .Date_Result
        .Date_Add RTrim(iH1_YM), M, 2, ResultDate
        mYM_P2 = .Date_Result
    End With

'---------------------------------------------------------------------------------------------------------------' 2010.01.29 del start
'' 発注内示テーブル検索
'    gSL_Select = ""
'    gSL_Select = gSL_Select & "SELECT HN0.内示年月, HN0.品番, HN0.手配先, HN0.受入"     '--, HN0.計画年月
'    gSL_Select = gSL_Select & "      ,HK1.機械 AS 機械, HK2.機械NO AS 機械NO, ISNULL(HK2.工順,0) AS 工順"
'    gSL_Select = gSL_Select & "      ,TR1.略称 AS 手配先名"
'    gSL_Select = gSL_Select & "      ,MM1.値名称 AS 機械名, MM2.値名称 AS 機械NO名"
'    gSL_Select = gSL_Select & "      ,HB1.表示品番 AS 表示品番"
'    gSL_Select = gSL_Select & "      ,HK2.工程数, HK2.基準工数1, HK2.基準工数2"
'    gSL_Select = gSL_Select & "      ,ISNULL(HK2.工程数,0) AS 工程数"
'    gSL_Select = gSL_Select & "      ,ISNULL(HK2.基準工数1,0) AS 基準工数1, ISNULL(HK2.基準工数2,0) AS 基準工数2"
'    gSL_Select = gSL_Select & "      ,ISNULL(HN1.数量,0) AS 当月内示数"
'    gSL_Select = gSL_Select & "      ,ISNULL(HN2.数量,0) AS 翌月内示数"
'    gSL_Select = gSL_Select & "      ,ISNULL(HN3.数量,0) AS 翌々月内示数"
'  ''gSL_Select = gSL_Select & "--      ,HN0.*"
'    gSL_Select = gSL_Select & " FROM  発注内示テーブル HN0"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUM(H1.数量) AS 数量, H1.内示年月, H1.品番, H1.手配先"    '--, H1.計画年月
'    gSL_Select = gSL_Select & "                         FROM  発注内示テーブル H1"
'    gSL_Select = gSL_Select & "                         WHERE H1.内示年月 = '" & Trim(iH1_YM) & "'"
'    gSL_Select = gSL_Select & "                           AND H1.計画年月 = '" & Trim(iH1_YM) & "'"     '--(H1.計画年月 BETWEEN '201001' AND '201003')"
'    gSL_Select = gSL_Select & "                         GROUP BY H1.内示年月, H1.計画年月, H1.品番, H1.手配先) HN1"     '--, H1.計画年月
'    gSL_Select = gSL_Select & "        ON  HN1.内示年月 = HN0.内示年月"
'  ''gSL_Select = gSL_Select & "      --AND HN1.計画年月 = HN0.計画年月"
'    gSL_Select = gSL_Select & "        AND HN1.品番     = HN0.品番"
'    gSL_Select = gSL_Select & "        AND HN1.手配先   = HN0.手配先"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUM(H2.数量) AS 数量, H2.内示年月, H2.品番, H2.手配先"    '--, H2.計画年月"
'    gSL_Select = gSL_Select & "                         FROM  発注内示テーブル H2"
'    gSL_Select = gSL_Select & "                         WHERE H2.内示年月 = '" & Trim(iH1_YM) & "'"
'    gSL_Select = gSL_Select & "                           AND H2.計画年月 = '" & Trim(mYM_P1) & "'"     '--(H1.計画年月 BETWEEN '201001' AND '201003')"
'    gSL_Select = gSL_Select & "                         GROUP BY H2.内示年月, H2.計画年月, H2.品番, H2.手配先) HN2"     '--, H2.計画年月"
'    gSL_Select = gSL_Select & "        ON  HN2.内示年月 = HN0.内示年月"
'  ''gSL_Select = gSL_Select & "      --AND HN2.計画年月 = HN0.計画年月"
'    gSL_Select = gSL_Select & "        AND HN2.品番     = HN0.品番"
'    gSL_Select = gSL_Select & "        AND HN2.手配先   = HN0.手配先"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUM(H3.数量) AS 数量, H3.内示年月, H3.品番, H3.手配先"    '--, H3.計画年月"
'    gSL_Select = gSL_Select & "                         FROM  発注内示テーブル H3"
'    gSL_Select = gSL_Select & "                         WHERE H3.内示年月 = '" & Trim(iH1_YM) & "'"
'    gSL_Select = gSL_Select & "                           AND H3.計画年月 = '" & Trim(mYM_P2) & "'"     '--(H1.計画年月 BETWEEN '201001' AND '201003')"
'    gSL_Select = gSL_Select & "                         GROUP BY H3.内示年月, H3.計画年月, H3.品番, H3.手配先) HN3"     '--, H3.計画年月"
'    gSL_Select = gSL_Select & "        ON  HN3.内示年月 = HN0.内示年月"
'  ''gSL_Select = gSL_Select & "      --AND HN3.計画年月 = HN0.計画年月"
'    gSL_Select = gSL_Select & "        AND HN3.品番     = HN0.品番"
'    gSL_Select = gSL_Select & "        AND HN3.手配先   = HN0.手配先"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT    KJ1.品番, KJ1.取引先, KJ1.受入, MAX(KJ1.機械) AS 機械"
'    gSL_Select = gSL_Select & "                         FROM  品番工順マスタ KJ1"
'    gSL_Select = gSL_Select & "                         GROUP BY KJ1.品番, KJ1.取引先, KJ1.受入) HK1"
'    gSL_Select = gSL_Select & "        ON  HK1.品番   = HN0.品番"
'    gSL_Select = gSL_Select & "        AND HK1.取引先 = HN0.手配先"
'    gSL_Select = gSL_Select & "        AND HK1.受入   = HN0.受入"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT KJ2.品番,   KJ2.取引先, KJ2.受入, KJ2.機械, KJ2.機械NO, MAX(KJ2.工順) AS 工順"
'    gSL_Select = gSL_Select & "                              ,KJ2.工程数, KJ2.基準工数1, KJ2.基準工数2"
'    gSL_Select = gSL_Select & "                         FROM  品番工順マスタ KJ2"
'    gSL_Select = gSL_Select & "                         GROUP BY KJ2.品番, KJ2.取引先, KJ2.受入, KJ2.機械, KJ2.機械NO"
'    gSL_Select = gSL_Select & "                                 ,KJ2.工程数, KJ2.基準工数1, KJ2.基準工数2) HK2"
'    gSL_Select = gSL_Select & "        ON  HK2.品番   = HN0.品番"
'    gSL_Select = gSL_Select & "        AND HK2.取引先 = HN0.手配先"
'    gSL_Select = gSL_Select & "        AND HK2.受入   = HN0.受入"
'    gSL_Select = gSL_Select & "        AND HK2.機械   = HK1.機械"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TR1"
'    gSL_Select = gSL_Select & "        ON  TR1.取引先     = HN0.手配先"
'    gSL_Select = gSL_Select & "        AND TR1.取引先区分 = 1"
'    gSL_Select = gSL_Select & "        AND TR1.社内区分   = 1"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ MM1"
'    gSL_Select = gSL_Select & "        ON  MM1.区分名称 = '機械'"
'    gSL_Select = gSL_Select & "        AND MM1.値       = HK1.機械"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ MM2"
'    gSL_Select = gSL_Select & "        ON  MM2.区分名称 = '機械NO'"
'    gSL_Select = gSL_Select & "        AND MM2.値       = HK2.機械NO"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 品番マスタ HB1"
'    gSL_Select = gSL_Select & "        ON  HB1.品番     = HN0.品番"
'    gSL_Select = gSL_Select & " WHERE HN0.内示年月 = '" & iH1_YM & "'"
'    If Trim(iH1_Torcd) <> "" Then _
'    gSL_Select = gSL_Select & "   AND HN0.手配先   = '" & Trim(iH1_Torcd) & "'"
'    If Trim(iH1_Kikai) <> "" Then _
'    gSL_Select = gSL_Select & "   AND HK1.機械     = '" & Trim(iH1_Kikai) & "'"
'    If Trim(iH1_Hinbn) <> "" Then _
'    gSL_Select = gSL_Select & "   AND HN0.品番     = '" & Trim(iH1_Hinbn) & "'"
'    gSL_Select = gSL_Select & " GROUP BY HN0.内示年月, HN0.品番, HN0.手配先, HN0.受入, HN1.数量, HN2.数量, HN3.数量"    '--, HN0.計画年月"
'    gSL_Select = gSL_Select & "         ,HK1.機械, HK2.機械NO, HK2.工順"
'    gSL_Select = gSL_Select & "         ,MM1.値名称,MM2.値名称"
'    gSL_Select = gSL_Select & "         ,TR1.略称, HB1.表示品番"
'    gSL_Select = gSL_Select & "         ,HK2.工程数, HK2.基準工数1, HK2.基準工数2"
'    gSL_Select = gSL_Select & " ORDER BY HN0.手配先  , HN0.受入, HN0.品番  , HN0.内示年月"              '--, HN0.計画年月"
'    If Not HNJRead(gSL_Select) Then
'---------------------------------------------------------------------------------------------------------------' 2010.01.29 del end

'---上記と差替--------------------------------------------------------------------------------------------------' 2010.01.29 add start
' 発注内示テーブル検索
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT HN1.内示年月 AS 納期," & Chr(13)
'    gSL_Select = gSL_Select & "       HK0.品番, HK0.取引先, HK0.受入, HK1.機械, HK2.機械NO, ISNULL(HK2.工順,0) AS 工順" & Chr(13)
'    gSL_Select = gSL_Select & "      ,TR1.略称 AS 得意先名, HB1.表示品番 AS 表示品番, MM1.値名称 AS 機械名, MM2.値名称 AS 機械NO名" & Chr(13)
'    gSL_Select = gSL_Select & "      ,ISNULL(HK2.工程数   ,0) AS 工程数" & Chr(13)
'    gSL_Select = gSL_Select & "      ,ISNULL(HK2.基準工数1,0) AS 基準工数1, ISNULL(HK2.基準工数2,0) AS 基準工数2" & Chr(13)
    
    gSL_Select = gSL_Select & "       HK0.品番, HK0.取引先, HK0.受入, HK0.機械, HK0.機械NO, HK0.工順" & Chr(13)
    gSL_Select = gSL_Select & "      ,TR1.略称 AS 得意先名, HB1.表示品番 AS 表示品番, MM1.値名称 AS 機械名, MM2.値名称 AS 機械NO名" & Chr(13)
    gSL_Select = gSL_Select & "      ,ISNULL(HK0.工程数   ,0) AS 工程数" & Chr(13)
    gSL_Select = gSL_Select & "      ,ISNULL(HK0.基準工数1,0) AS 基準工数1, ISNULL(HK0.基準工数2,0) AS 基準工数2" & Chr(13)
    
    gSL_Select = gSL_Select & "      ,ISNULL(HN1.数量,0) AS 内示当月" & Chr(13)
    gSL_Select = gSL_Select & "      ,ISNULL(HN2.数量,0) AS 内示翌月" & Chr(13)
    gSL_Select = gSL_Select & "      ,ISNULL(HN3.数量,0) AS 内示翌々月" & Chr(13)
    gSL_Select = gSL_Select & " FROM  品番工順マスタ HK0" & Chr(13)
    
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT KJ1.品番, KJ1.取引先, KJ1.受入, MAX(KJ1.機械) AS 機械" & Chr(13)
'    gSL_Select = gSL_Select & "                         FROM  品番工順マスタ KJ1" & Chr(13)
'    gSL_Select = gSL_Select & "                         GROUP BY KJ1.品番, KJ1.取引先, KJ1.受入) HK1" & Chr(13)
'    gSL_Select = gSL_Select & "        ON  HK1.品番   = HK0.品番" & Chr(13)
'    gSL_Select = gSL_Select & "        AND HK1.取引先 = HK0.取引先" & Chr(13)
'    gSL_Select = gSL_Select & "        AND HK1.受入   = HK0.受入" & Chr(13)
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT KJ2.品番,   KJ2.取引先, KJ2.受入, KJ2.機械, KJ2.機械NO, MAX(KJ2.工順) AS 工順" & Chr(13)
'    gSL_Select = gSL_Select & "                              ,KJ2.工程数, KJ2.基準工数1, KJ2.基準工数2" & Chr(13)
'    gSL_Select = gSL_Select & "                         FROM  品番工順マスタ KJ2" & Chr(13)
'    gSL_Select = gSL_Select & "                         GROUP BY KJ2.品番, KJ2.取引先, KJ2.受入, KJ2.機械, KJ2.機械NO" & Chr(13)
'    gSL_Select = gSL_Select & "                                 ,KJ2.工程数, KJ2.基準工数1, KJ2.基準工数2) HK2" & Chr(13)
'    gSL_Select = gSL_Select & "        ON  HK2.品番   = HK0.品番" & Chr(13)
'    gSL_Select = gSL_Select & "        AND HK2.取引先 = HK0.取引先" & Chr(13)
'    gSL_Select = gSL_Select & "        AND HK2.受入   = HK0.受入" & Chr(13)
'    gSL_Select = gSL_Select & "        AND HK2.機械   = HK1.機械" & Chr(13)
    
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUM(H1.数量) AS 数量, H1.内示年月, H1.品番, H1.手配先, H1.計画年月, H1.受入" & Chr(13)
    gSL_Select = gSL_Select & "                         FROM  発注内示テーブル H1" & Chr(13)
    gSL_Select = gSL_Select & "                         WHERE H1.内示年月 = '" & iH1_YM & "'" & Chr(13)
    gSL_Select = gSL_Select & "                           AND H1.計画年月 = '" & iH1_YM & "'" & Chr(13)
    gSL_Select = gSL_Select & "                         GROUP BY H1.内示年月, H1.計画年月, H1.品番, H1.手配先, H1.受入) HN1" & Chr(13)
    gSL_Select = gSL_Select & "        ON  HN1.品番   = HK0.品番" & Chr(13)
    gSL_Select = gSL_Select & "        AND HN1.手配先 = HK0.取引先 AND HN1.受入 = HK0.受入" & Chr(13)
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUM(H2.数量) AS 数量, H2.内示年月, H2.品番, H2.手配先, H2.計画年月, H2.受入" & Chr(13)
    gSL_Select = gSL_Select & "                         FROM  発注内示テーブル H2" & Chr(13)
    gSL_Select = gSL_Select & "                         WHERE H2.内示年月 = '" & iH1_YM & "'" & Chr(13)
    gSL_Select = gSL_Select & "                           AND H2.計画年月 = '" & mYM_P1 & "'" & Chr(13)
    gSL_Select = gSL_Select & "                         GROUP BY H2.内示年月, H2.計画年月, H2.品番, H2.手配先, H2.受入) HN2" & Chr(13)
    gSL_Select = gSL_Select & "        ON  HN2.品番   = HK0.品番" & Chr(13)
    gSL_Select = gSL_Select & "        AND HN2.手配先 = HK0.取引先 AND HN2.受入 = HK0.受入" & Chr(13)
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUM(H3.数量) AS 数量, H3.内示年月, H3.品番, H3.手配先, H3.計画年月, H3.受入" & Chr(13)
    gSL_Select = gSL_Select & "                         FROM  発注内示テーブル H3" & Chr(13)
    gSL_Select = gSL_Select & "                         WHERE H3.内示年月 = '" & iH1_YM & "'" & Chr(13)
    gSL_Select = gSL_Select & "                           AND H3.計画年月 = '" & mYM_P2 & "'" & Chr(13)
    gSL_Select = gSL_Select & "                         GROUP BY H3.内示年月, H3.計画年月, H3.品番, H3.手配先, H3.受入) HN3" & Chr(13)
    gSL_Select = gSL_Select & "        ON  HN3.品番   = HK0.品番" & Chr(13)
    gSL_Select = gSL_Select & "        AND HN3.手配先 = HK0.取引先 AND HN3.受入 = HK0.受入" & Chr(13)
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TR1" & Chr(13)
    gSL_Select = gSL_Select & "        ON  TR1.取引先CD = HK0.取引先 AND TR1.取引先区分 = 1 " & Chr(13)  ''''AND TR1.社内区分 = 1" & Chr(13)
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ MM1  ON  MM1.区分名称 = '機械'   AND MM1.値 = HK1.機械" & Chr(13)
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ MM2  ON  MM2.区分名称 = '機械NO' AND MM2.値 = HK2.機械NO" & Chr(13)
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ MM1  ON  MM1.区分名称 = '機械'   AND MM1.値 = HK0.機械" & Chr(13)
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ MM2  ON  MM2.区分名称 = '機械NO' AND MM2.値 = HK0.機械NO" & Chr(13)
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 品番マスタ HB1  ON  HB1.品番     = HK0.品番" & Chr(13)
    gSL_Select = gSL_Select & " WHERE HN1.内示年月 = '" & iH1_YM & "'" & Chr(13)
    If Trim(iH1_Torcd) <> "" Then _
    gSL_Select = gSL_Select & "   AND HK0.取引先   = '" & Trim(iH1_Torcd) & "'" & Chr(13)
    If Trim(iH1_Kikai) <> "" Then _
    gSL_Select = gSL_Select & "   AND HK0.機械     = '" & Trim(iH1_Kikai) & "'" & Chr(13)
    If Trim(iH1_Hinbn) <> "" Then _
    gSL_Select = gSL_Select & "   AND HK0.品番     = '" & Trim(iH1_Hinbn) & "'" & Chr(13)

    '================== 2013/05/10 Start
    If Trim(iH1_KikaiNo) <> "" Then _
    gSL_Select = gSL_Select & "   AND HK0.機械NO     = '" & Trim(iH1_KikaiNo) & "'" & Chr(13)
    '================== 2013/05/10 End
    
'    gSL_Select = gSL_Select & " GROUP BY HN1.内示年月," & Chr(13)
'    gSL_Select = gSL_Select & "          HK0.品番, HK0.取引先, HK0.受入, HK1.機械, HK2.機械NO, HK2.工順" & Chr(13)
'    gSL_Select = gSL_Select & "         ,TR1.略称, HB1.表示品番, MM1.値名称, MM2.値名称" & Chr(13)
'    gSL_Select = gSL_Select & "         ,HK2.工程数, HK2.基準工数1, HK2.基準工数2" & Chr(13)
'    gSL_Select = gSL_Select & "         ,HN1.数量, HN2.数量, HN3.数量" & Chr(13)
    gSL_Select = gSL_Select & " GROUP BY HN1.内示年月," & Chr(13)
    gSL_Select = gSL_Select & "          HK0.品番, HK0.取引先, HK0.受入, HK0.機械, HK0.機械NO, HK0.工順" & Chr(13)
    gSL_Select = gSL_Select & "         ,TR1.略称, HB1.表示品番, MM1.値名称, MM2.値名称" & Chr(13)
    gSL_Select = gSL_Select & "         ,HK0.工程数, HK0.基準工数1, HK0.基準工数2" & Chr(13)
    gSL_Select = gSL_Select & "         ,HN1.数量, HN2.数量, HN3.数量" & Chr(13)
'    gSL_Select = gSL_Select & " ORDER BY HK0.取引先  , HK0.受入, HK0.品番, HN1.内示年月" & Chr(13)
    gSL_Select = gSL_Select & " ORDER BY  HK0.取引先  ,MM2.値名称,  HK0.受入,  HK0.品番, HN1.内示年月" & Chr(13)      '2013/05/10(機械NO順)
    
    If Not HKTRead(gSL_Select) Then
'---------------------------------------------------------------------------------------------------------------' 2010.01.29 add end
       Exit Function
    End If
    Call GridSet
    
    Call HKTClose
    BodySet = True
End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    ReDim SCol(10)
    
    Dim wYM0    As String
    Dim wYM1    As String
    Dim wYM2    As String
    
    wYM0 = Mid(Format(Now(), "yyyymmdd"), 1, 6)
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add wYM0, M, 1, ResultDate
        wYM1 = .Date_Result
        .Date_Add wYM0, M, 2, ResultDate
        wYM2 = .Date_Result
    End With
    
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
'       Call .FixedSet(左中, 左中, あり, 0, 0, "")
        Call .FixedSet(左中, 左中, あり, 8, 0, "部門")
        Call .FixedSet(左中, 左中, あり, 10, 1, "機械")
        Call .FixedSet(左中, 左中, あり, 6, 2, "機械№")
        Call .FixedSet(左中, 左中, あり, 16, 3, "品番")
        Call .FixedSet(左中, 右中, あり, 6, 4, "工程数")
        Call .FixedSet(左中, 右中, あり, 8, 5, "基準時間")
        
        Call .FixedSet(中中, 右中, あり, 10, 6, Mid(wYM0, 5, 2) & "月内示数")
        Call .FixedSet(中中, 右中, あり, 8, 7, Mid(wYM0, 5, 2) & "月工数")
        Call .FixedSet(中中, 右中, あり, 10, 8, Mid(wYM0, 5, 2) & "月ﾊﾟﾝﾁ数")
        Call .FixedSet(中中, 右中, あり, 10, 9, Mid(wYM1, 5, 2) & "月内示数")
        Call .FixedSet(中中, 右中, あり, 8, 10, Mid(wYM1, 5, 2) & "月工数")
        Call .FixedSet(中中, 右中, あり, 10, 11, Mid(wYM1, 5, 2) & "月ﾊﾟﾝﾁ数")
        Call .FixedSet(中中, 右中, あり, 10, 12, Mid(wYM2, 5, 2) & "月内示数")
        Call .FixedSet(中中, 右中, あり, 8, 13, Mid(wYM2, 5, 2) & "月工数")
        Call .FixedSet(中中, 右中, あり, 10, 14, Mid(wYM2, 5, 2) & "月ﾊﾟﾝﾁ数")

        '-- データ退避エリア
        Call .FixedSet(左中, 左中, あり, 10, 16, "退避部門")
        Call .FixedSet(左中, 左中, あり, 10, 17, "退避機械")
        Call .FixedSet(左中, 左中, あり, 10, 18, "退避機械NO")
        Call .FixedSet(左中, 左中, あり, 25, 19, "退避品番")
        Call .FixedSet(左中, 左中, あり, 10, 20, "退避受入")
        Call .FixedSet(左中, 右中, あり, 10, 21, "退避工順")

'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(左中, 中中, なし, 3, 15, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14")                         ' 運用時は14まで
      ''Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21")    ' 試験用
        .InitDisp
    
'---------------------------------------------------------------------------' 2010.01.19 del start
''''        SCol(1) = .FixedGet(1)  ' 区分
''''        SCol(2) = .FixedGet(2)  ' 年月度
''''        SCol(3) = .FixedGet(3)  ' 在庫数
''''        SCol(4) = .FixedGet(4)  ' 実棚日
''''        SCol(5) = .FixedGet(5)  ' 変更事由
''''        SCol(6) = .FixedGet(6)  ' 担当
''''        SCol(7) = .FixedGet(7)  ' 変更日
'---------------------------------------------------------------------------' 2010.01.19 del end
    End With
'    KCol = 8   ' 非表示項目開始ｶﾗﾑ
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
  ''Dim Ttl_Kensu   As Long
    Dim wKjnKousu   As Double
    
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    mTtl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
    
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid1
       .Redraw = False
'---------------------------------------------------------------------------' 2010.01.19 del start
''''        ReDim GridCol(.Cols)
''''        For ID = 0 To .Cols - 1
''''            GridCol(ID) = CisVsGrid.FixedGet(CisFun.Val2(ID))
''''        Next ID
'---------------------------------------------------------------------------' 2010.01.19 del end
        If mTtl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = mTtl_Kensu + 1
        End If
    
        '-- 見出し部作成
        .TextMatrix(0, 6) = Mid(iH1_YM, 5, 2) & "月内示数"
        .TextMatrix(0, 7) = Mid(iH1_YM, 5, 2) & "月工数"
        .TextMatrix(0, 8) = Mid(iH1_YM, 5, 2) & "月ﾊﾟﾝﾁ数"
        .TextMatrix(0, 9) = Mid(mYM_P1, 5, 2) & "月内示数"
        .TextMatrix(0, 10) = Mid(mYM_P1, 5, 2) & "月工数"
        .TextMatrix(0, 11) = Mid(mYM_P1, 5, 2) & "月ﾊﾟﾝﾁ数"
        .TextMatrix(0, 12) = Mid(mYM_P2, 5, 2) & "月内示数"
        .TextMatrix(0, 13) = Mid(mYM_P2, 5, 2) & "月工数"
        .TextMatrix(0, 14) = Mid(mYM_P2, 5, 2) & "月ﾊﾟﾝﾁ数"
    End With
    
    ID = 0
    Do Until Not HKT_RDSTS
        With VsGrid1
            ID = ID + 1
            ' №
          ''.TextMatrix(ID, 0) = Format(ID, "###")
           ' 部門
            .TextMatrix(ID, 0) = Trim(HKT.得意先名)
            ' 機械
            .TextMatrix(ID, 1) = Trim(HKT.機械名)
            ' 機械№
            .TextMatrix(ID, 2) = Trim(HKT.機械NO名)
            ' 品番
            .TextMatrix(ID, 3) = HKT.表示品番
            ' 工程数
            .TextMatrix(ID, 4) = Format(HKT.工程数, "###0.0")
            ' 基準時間
            If OpH1_Kousu(0).Value = True Then
               wKjnKousu = HKT.基準工数1
            Else
               wKjnKousu = HKT.基準工数2
            End If
            .TextMatrix(ID, 5) = Format(wKjnKousu, "#0.000")
'
            ' 当月内示数
            .TextMatrix(ID, 6) = Format(HKT.内示当月, "########")
            ' 当月工数
            If wKjnKousu <> 0 Then _
            .TextMatrix(ID, 7) = Format(HKT.内示当月 * wKjnKousu, "#####0.0")
            ' 当月ﾊﾟﾝﾁ数
            .TextMatrix(ID, 8) = Format(HKT.内示当月 * HKT.工程数, "########")
            ' 翌月内示数
            .TextMatrix(ID, 9) = Format(HKT.内示翌月, "########")
            ' 翌月工数
            If wKjnKousu <> 0 Then _
            .TextMatrix(ID, 10) = Format(HKT.内示翌月 * wKjnKousu, "#####0.0")
            ' 翌月ﾊﾟﾝﾁ数
            .TextMatrix(ID, 11) = Format(HKT.内示翌月 * HKT.工程数, "########")
            ' 翌々月内示数
            .TextMatrix(ID, 12) = Format(HKT.内示翌々月, "########")
            ' 翌々月工数
            If wKjnKousu <> 0 Then _
            .TextMatrix(ID, 13) = Format(HKT.内示翌々月 * wKjnKousu, "#####0.0")
            ' 翌々月ﾊﾟﾝﾁ数
            .TextMatrix(ID, 14) = Format(HKT.内示翌々月 * HKT.工程数, "########")
            
            ' 退避部門
            .TextMatrix(ID, 16) = Trim(HKT.取引先)
            ' 退避機械
            .TextMatrix(ID, 17) = Trim(HKT.機械)
            ' 退避機械NO
            .TextMatrix(ID, 18) = Trim(HKT.機械NO)
            ' 退避品番
            .TextMatrix(ID, 19) = Trim(HKT.品番)
            ' 退避受入
            .TextMatrix(ID, 20) = Trim(HKT.受入)
            ' 退避工順
            .TextMatrix(ID, 21) = Format(HKT.工順, "##0")
        End With
        Call HKTReadNext
    Loop
    
'-- TEST Start ---------------------------------------------------------------------------------' test data set start
''GoTo test_exit
'test1:
'    With VsGrid1
'       '.Left = .Left - 1080
'        .TextMatrix(1, 5) = "3.000"
'        .TextMatrix(2, 5) = "0.200"
'
'        Dim wCnt As Integer
'        For wCnt = 3 To .Rows - 1
'            .TextMatrix(wCnt, 5) = "1.234"
'        Next
'        mTtl_Kensu = .Rows - 1
'    End With
'test_exit:
'-- TEST End -----------------------------------------------------------------------------------' test data set end
    
    With VsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub

Private Sub VsGrid1_Click()
    Dim wCnt    As Integer
    Dim wRow    As Integer
    
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
    
    With VsGrid1
        If .Row > mTtl_Kensu Then .Row = mTtl_Kensu
        
        wRow = .Row
        .Row = ActRow
        For wCnt = 0 To 5
            .Col = wCnt
            .CellBackColor = .BackColorFixed
            .CellForeColor = .ForeColor
        Next
        .Row = wRow
    
        VsGrid1.RowSel = VsGrid1.Row
        ActRow = VsGrid1.Row

        For wCnt = 0 To 5
            .Col = wCnt
            .CellBackColor = .BackColorSel
            .CellForeColor = .ForeColorSel
        Next
    End With

    wRow = ActRow
    If ActRow > mTtl_Kensu Then wRow = mTtl_Kensu
    lb_Kensu.Caption = Format(wRow, "###,##0") & " / " & _
                       Format(mTtl_Kensu, "###,##0")
                       
    Call VsGrid1.SetFocus

' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
End Sub
'+----------------------------------------------------------+
'+ キー入力にて行移動 《 KeyUp イベント   》                  +
'+----------------------------------------------------------+
Private Sub vsGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or _
        KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then
            VsGrid1.RowSel = VsGrid1.Row
            Call VsGrid1_Click
    End If
End Sub

'+==========================================================+
'+      グリッド処理 (2)                                    +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub Grid2Init()
    Dim wCnt1   As Integer
    Dim wCnt2   As Integer
    
    ReDim SCol(10)
    
    With CisVsGrid2
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 8, 0, "")
        Call .FixedSet(中中, 中中, あり, 6, 1, "　")
        
        Call .FixedSet(中中, 右中, あり, 10, 2, "09月内示数")
        Call .FixedSet(中中, 右中, あり, 8, 3, "09月工数")
        Call .FixedSet(中中, 右中, あり, 10, 4, "09月ﾊﾟﾝﾁ数")
        Call .FixedSet(中中, 右中, あり, 10, 5, "10月内示数")
        Call .FixedSet(中中, 右中, あり, 8, 6, "10月工数")
        Call .FixedSet(中中, 右中, あり, 10, 7, "10月ﾊﾟﾝﾁ数")
        Call .FixedSet(中中, 右中, あり, 10, 8, "11月内示数")
        Call .FixedSet(中中, 右中, あり, 8, 9, "11月工数")
        Call .FixedSet(中中, 右中, あり, 10, 10, "11月ﾊﾟﾝﾁ数")

'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
      ''Call .FixedSet(中中, 左中, なし, 0, 11, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "1,2,3,4,5,6,7,8,9,10")
        .InitDisp
    End With

    With VsGrid2
        .Left = 1050
        .Width = .Width - .Left
        .TextMatrix(1, 0) = "内示"
        .TextMatrix(2, 0) = "実績"
        .TextMatrix(3, 0) = "差"
        .FixedCols = 1
        .Cols = 11
    End With
End Sub
'******************************************'
'*****      行選択  (詳細／取消ボタン)
'******************************************'
Private Sub vsGrid1_Select(Optional wSel As String = "")
    Dim wCnt As Integer

    With VsGrid1
        Call CisVsGrid.vsColor
        If wSel = "" Then
            Call VsGrid1_Click
        Else
            For wCnt = 0 To 5
                .Col = wCnt
                .CellForeColor = .ForeColor
            Next
        End If
    End With
End Sub

'****************************
'*       CSV出力処理        *   ' 参考：PNJ1110
'****************************
Private Function CsvPut() As Boolean

    Dim OutStr      As String
    Dim OutFNo      As Long
    Dim OutCnt      As Long
    
    Dim wCnt        As Integer
''''Dim Moji_F      As String
''''Dim Moji_B      As String
''''Moji_F = "=" & Chr(34): Moji_B = Chr(34) & ","
    
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
    
    On Error GoTo CsvPut_Err
    OutFNo = FreeFile
    Open mCsvDir For Output As #OutFNo
    On Error GoTo 0

    Call CisFrm.MousePT(11)
    
    lb_PrtDev = mCsvDir
    lb_msg.Caption = "【 内示・生産負荷一覧 出力中 】"
  ''Csv_Bar.Visible = True
    DoEvents
    
    With CisFrm
        Set .MeterPic = Csv_Bar
        .MeterMax = mTtl_Kensu
    End With
        
    ' 表題
    OutStr = ""
    OutStr = OutStr & "内示・生産負荷一覧" & vbCrLf
    OutStr = OutStr & "対象年月 = " & Mid(iH1_YM, 1, 4) & "年" & Format(Mid(iH1_YM, 5, 2), "00") & "月　：　"
    If OpH1_Kousu(0).Value = True Then OutStr = OutStr & "工数 = 基準工数" Else OutStr = OutStr & "工数 = 目標工数"
    Print #OutFNo, OutStr
    
    ' 見出し
'=========================== 2013/05/10 Start
'    OutStr = "部門,機械,機械№,品番,工程数,基準時間"
    If OpH1_Kousu(0) Then
       OutStr = "部門,機械,機械№,品番,工程数,基準時間"
    Else
       OutStr = "部門,機械,機械№,品番,工程数,目標時間"
    End If
'=========================== 2013/05/10 End
    
    OutStr = OutStr & "," & Mid(iH1_YM, 5, 2) & "月内示数"
    OutStr = OutStr & "," & Mid(iH1_YM, 5, 2) & "月工数"
    OutStr = OutStr & "," & Mid(iH1_YM, 5, 2) & "月ﾊﾟﾝﾁ数"
    OutStr = OutStr & "," & Mid(mYM_P1, 5, 2) & "月内示数"
    OutStr = OutStr & "," & Mid(mYM_P1, 5, 2) & "月工数"
    OutStr = OutStr & "," & Mid(mYM_P1, 5, 2) & "月ﾊﾟﾝﾁ数"
    OutStr = OutStr & "," & Mid(mYM_P2, 5, 2) & "月内示数"
    OutStr = OutStr & "," & Mid(mYM_P2, 5, 2) & "月工数"
    OutStr = OutStr & "," & Mid(mYM_P2, 5, 2) & "月ﾊﾟﾝﾁ数"
    Print #OutFNo, OutStr
    
    ' 内容
    OutCnt = 0
    Do While OutCnt < mTtl_Kensu
        With VsGrid1
            OutCnt = OutCnt + 1
            OutStr = ""
            For wCnt = 0 To 3
                OutStr = OutStr & .TextMatrix(OutCnt, wCnt) & ","
            Next
            OutStr = OutStr & Format(CisFun.Val2(.TextMatrix(OutCnt, 4)), "###0.0") & ","
            OutStr = OutStr & Format(CisFun.Val2(.TextMatrix(OutCnt, 5)), "###0.000") & ","
            For wCnt = 6 To 14
                If Trim(.TextMatrix(OutCnt, wCnt)) <> "" Then
                    OutStr = OutStr & .TextMatrix(OutCnt, wCnt) & ","
                Else
                    If (wCnt - 6) Mod 3 = 1 Then OutStr = OutStr & "0.0," _
                                            Else OutStr = OutStr & "0,"
                End If
            Next
        End With
        Print #OutFNo, OutStr
        CisFrm.MeterUpdate OutCnt
    Loop

CsvPut_Ed:
    Close #OutFNo
    SaveSetting appname:=gRegKey, section:="CSV", _
                    Key:=mCsvKey, setting:=RTrim(mCsvDir)
    
    Call CisFrm.MousePT(1)
    
    B2_Area3.Visible = False    ' バー消去
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "   出力処理が終了しました。    "
        .MB_MSG(4) = "   　 " & "出力先    = " & mCsvDir
        .MB_MSG(5) = "   　 " & "出力件数 = " & StrConv(Format(mTtl_Kensu, "##,##0"), vbWide) & " 件 "
        .MB_Title = "出力確認"
        .MB_Button = OK
        .MBOX
    End With
    
    CsvPut = True
    Exit Function

CsvPut_Err:
    On Error GoTo 0
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "    出力先ファイルを開けませんでした。                        "
        .MB_MSG(4) = "    ファイルを使用中のときはファイルを閉じてください。        "
        .MB_Title = "エラー"
        .MB_Button = Error
        .MBOX
    End With
  ''B2_Area3.Visible = False    ' バー消去
    Exit Function
End Function
'+------------------------------+
'+     CSV出力先 変更           +
'+------------------------------+
Private Sub PB_Print_Click()
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
End Sub

