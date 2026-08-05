VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CXM0050 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "構成マスタメンテナンス"
   ClientHeight    =   11190
   ClientLeft      =   90
   ClientTop       =   1515
   ClientWidth     =   15345
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11190
   ScaleWidth      =   15345
   WindowState     =   2  '最大化
   Begin VB.PictureBox Picture3 
      Height          =   9435
      Left            =   150
      ScaleHeight     =   9375
      ScaleWidth      =   5565
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1620
      Width           =   5625
      Begin MSComctlLib.TreeView TView1 
         Height          =   8955
         Left            =   0
         TabIndex        =   2
         Top             =   390
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   15796
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   405
         Left            =   0
         Top             =   0
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   714
         BackColor       =   8421504
         ForeColor       =   16744703
         Caption         =   "構   成   情   報"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
      End
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   3165
      Left            =   5880
      ScaleHeight     =   3105
      ScaleWidth      =   9255
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1590
      Width           =   9315
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   2985
         Left            =   60
         Top             =   60
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5265
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   765
            Index           =   1
            Left            =   7830
            Top             =   30
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   1349
            ForeColor       =   16711680
            Caption         =   "背番号"
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
            Begin Cis3D_v60.CIS3D B1lb_SSeban 
               Height          =   345
               Left            =   300
               Top             =   360
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   609
               BackColor       =   16777152
               Caption         =   "XXXXX"
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
         Begin vsFlexLib.vsFlexArray VsGridT 
            Height          =   1365
            Left            =   30
            TabIndex        =   11
            Top             =   1590
            Width           =   9075
            _Version        =   196608
            _ExtentX        =   16007
            _ExtentY        =   2408
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
            FormatString    =   "№ |取引先               |受|背番号|手配|比率|SK|ｻｲｸﾙ   |有効期限               |係1 |係2 |係3 |LT|ロッ ト|月 火 水 木 金|"
            Rows            =   3
            Cols            =   16
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   330
            SelectionMode   =   1
            BackColorAlternate=   12648447
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   765
            Left            =   5190
            Top             =   30
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   1349
            ForeColor       =   16711680
            Caption         =   "使用数"
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
            Begin CisText_V60.CisText iB1_Siyou 
               Height          =   360
               Left            =   180
               TabIndex        =   4
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   635
               cFormat         =   "#,##0.0000"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#,##0.0000"
               MaxLength       =   10
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   4
               cFdAutoFormat   =   1
               cGFormat        =   "###0.0000"
               cILength        =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   765
            Index           =   0
            Left            =   5190
            Top             =   810
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   1349
            ForeColor       =   16711680
            Caption         =   "適 用 範 囲"
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
            Begin CisYMD_With_Btn.CisYMDwB iB1_SDate 
               Height          =   360
               Left            =   180
               TabIndex        =   9
               Top             =   330
               Width           =   1575
               _ExtentX        =   2249
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
               Object.Width           =   1575
               cSize           =   1
               cChkResult      =   0   'False
               cCalenderPosition=   1
               cType           =   1
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
            Begin CisYMD_With_Btn.CisYMDwB iB1_EDate 
               Height          =   360
               Left            =   2190
               TabIndex        =   10
               Top             =   330
               Width           =   1575
               _ExtentX        =   2249
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
               Object.Width           =   1575
               cSize           =   1
               cChkResult      =   0   'False
               cCalenderPosition=   1
               cType           =   1
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
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   375
               Index           =   4
               Left            =   1800
               Top             =   330
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   661
               ForeColor       =   16711680
               Caption         =   "～"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
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
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   765
            Index           =   2
            Left            =   6840
            Top             =   30
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   1349
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D B1lb_Syakei 
               Height          =   345
               Left            =   210
               Top             =   360
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   609
               BackColor       =   16777152
               Caption         =   "XXXX"
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
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   765
            Index           =   6
            Left            =   30
            Top             =   810
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   1349
            ForeColor       =   16711680
            Caption         =   "展開区分"
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
            cPositionX      =   360
            cPositionY      =   30
            Begin VB.PictureBox Picture2 
               Height          =   675
               Left            =   2670
               ScaleHeight     =   615
               ScaleWidth      =   2235
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   30
               Width           =   2295
               Begin VB.CheckBox B1Chk_Tenkb 
                  Caption         =   " 行  う"
                  Height          =   255
                  Left            =   210
                  TabIndex        =   7
                  Top             =   330
                  Width           =   1785
               End
               Begin Cis3D_v60.CIS3D UC_3D15 
                  Height          =   285
                  Index           =   3
                  Left            =   0
                  Top             =   0
                  Width           =   2250
                  _ExtentX        =   3969
                  _ExtentY        =   503
                  ForeColor       =   16711680
                  Caption         =   "下層展開区分"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   6
               End
            End
            Begin VB.OptionButton B1Op_Tenkb 
               BackColor       =   &H00C0C0C0&
               Caption         =   "対象外"
               Height          =   225
               Index           =   1
               Left            =   1350
               TabIndex        =   6
               Top             =   390
               Width           =   1050
            End
            Begin VB.OptionButton B1Op_Tenkb 
               BackColor       =   &H00C0C0C0&
               Caption         =   "対象"
               Height          =   225
               Index           =   0
               Left            =   360
               TabIndex        =   5
               Top             =   390
               Value           =   -1  'True
               Width           =   810
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   765
            Left            =   1440
            Top             =   30
            Width           =   3750
            _ExtentX        =   6615
            _ExtentY        =   1349
            ForeColor       =   16711680
            Caption         =   "構成品番"
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
            Begin CisText_V60.CisText iB1_KHinbn 
               Height          =   360
               Left            =   240
               TabIndex        =   3
               Top             =   330
               Width           =   3240
               _ExtentX        =   5715
               _ExtentY        =   635
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
         Begin Cis3D_v60.CIS3D K_Sykbnm2 
            Height          =   765
            Left            =   15
            Top             =   30
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   1349
            ForeColor       =   16711680
            Caption         =   "追加"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   765
            Left            =   5190
            Top             =   810
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1349
            ForeColor       =   16711680
            Caption         =   "ＬＴ"
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
            Begin CisText_V60.CisText iB1_LT 
               Height          =   360
               Left            =   180
               TabIndex        =   8
               Top             =   330
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   635
               cFormat         =   "##"
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "##"
               MaxLength       =   2
               cDataType       =   1
               cDataReplace    =   1
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   -2
            End
         End
      End
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
      Height          =   735
      Left            =   4830
      ScaleHeight     =   675
      ScaleWidth      =   10305
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   420
      Width           =   10365
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   7725
         Top             =   30
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
         Left            =   8580
         Top             =   30
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
         Left            =   9435
         Top             =   30
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1755
         Top             =   30
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
         BtnType         =   4
         BtnCaption      =   "削除"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   255
         BtnShadowColor  =   16777215
         PFCaption       =   "( F3 )"
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
         PFBackColor     =   255
      End
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   45
         Top             =   30
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
         BtnCaption      =   "追加"
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
         PFCaption       =   "( F1 )"
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
      Begin CisBtn_60.CisBtn PB_MNT 
         Height          =   585
         Left            =   900
         Top             =   30
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
         BtnType         =   8
         BtnCaption      =   "修正"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16776960
         BtnShadowColor  =   0
         PFCaption       =   "( F2 )"
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
         PFBackColor     =   16776960
      End
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   3465
         Top             =   30
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   6870
         Top             =   30
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
         BtnType         =   5
         BtnCaption      =   "前進"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8421376
         BtnShadowColor  =   16777215
         PFCaption       =   "( Down )"
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
         PFBackColor     =   8421376
      End
      Begin CisBtn_60.CisBtn PB_BACK 
         Height          =   585
         Left            =   6015
         Top             =   30
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
         BtnType         =   6
         BtnCaption      =   "後退"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8421376
         BtnShadowColor  =   16777215
         PFCaption       =   "( Up )"
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
         PFBackColor     =   8421376
      End
      Begin CisBtn_60.CisBtn PB_CPY 
         Height          =   585
         Left            =   2610
         Top             =   30
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
         BtnType         =   9
         BtnCaption      =   "複写"
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
         PFCaption       =   "( F5 )"
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
         PFBackColor     =   16711935
      End
      Begin CisBtn_60.CisBtn PB_Zai 
         Height          =   585
         Left            =   5160
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
         BtnCaption      =   "関連付"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   0
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
         PFForeColor     =   65535
         PFShadowColor   =   0
         PFBackColor     =   0
      End
      Begin CisBtn_60.CisBtn PB_Hin 
         Height          =   585
         Left            =   4320
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
         BtnCaption      =   "マスタ"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8388736
         BtnShadowColor  =   16777215
         PFCaption       =   "( F9 )"
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
         PFBackColor     =   8388736
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   150
      ScaleHeight     =   1065
      ScaleWidth      =   4245
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   450
      Width           =   4305
      Begin Cis3D_v60.CIS3D HLB_Hinbn 
         Height          =   465
         Left            =   60
         Top             =   60
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   820
         ForeColor       =   16711680
         Caption         =   "品  番"
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
         cPositionX      =   90
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   330
            Left            =   930
            TabIndex        =   0
            Top             =   60
            Width           =   3180
            _ExtentX        =   5609
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D HLB_CHinbn 
         Height          =   465
         Left            =   60
         Top             =   540
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   820
         ForeColor       =   16711680
         Caption         =   "複写先"
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
         cPositionX      =   90
         Begin CisText_V60.CisText iH1_CHinbn 
            Height          =   330
            Left            =   930
            TabIndex        =   1
            Top             =   60
            Width           =   3180
            _ExtentX        =   5609
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
            IMEMode         =   2
         End
      End
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1050
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   1020
      Top             =   1050
   End
   Begin VB.Timer Timer2 
      Left            =   1050
      Top             =   990
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
      Caption         =   "【 構成マスタメンテナンス  】"
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
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'なし
         Height          =   345
         Index           =   1
         Left            =   10710
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   30
         Width           =   3675
         Begin Cis3D_v60.CIS3D CIS3D15 
            Height          =   345
            Left            =   2460
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   8421631
            ForeColor       =   0
            Caption         =   "削除"
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
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Delsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   8421631
               ForeColor       =   0
               Caption         =   "ZZ9"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
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
         Begin Cis3D_v60.CIS3D UC_3D6 
            Height          =   345
            Left            =   0
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   8454143
            ForeColor       =   0
            Caption         =   "追加"
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
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Addsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   8454143
               ForeColor       =   0
               Caption         =   "ZZ9"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
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
         Begin Cis3D_v60.CIS3D UC_3D7 
            Height          =   345
            Left            =   1230
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   16777088
            ForeColor       =   0
            Caption         =   "修正"
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
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Mntsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   16777088
               ForeColor       =   0
               Caption         =   "ZZ9"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
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
      End
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14400
         Top             =   90
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   8160
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
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   16711935
         Caption         =   "XXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
      End
   End
   Begin VB.PictureBox B2_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   6210
      Left            =   5880
      ScaleHeight     =   6150
      ScaleWidth      =   9270
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4860
      Width           =   9330
      Begin vsFlexLib.vsFlexArray VsGrid1 
         DragIcon        =   "CXM0050.frx":0000
         Height          =   6060
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   9150
         _Version        =   196608
         _ExtentX        =   16140
         _ExtentY        =   10689
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
         FormatString    =   "No |レベル|品番| 品名|使用数|単位|手配|手配先|受|有無|展|LT|適用範囲|材料区分|"
         Rows            =   21
         Cols            =   26
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   240
         SelectionMode   =   1
         BackColorAlternate=   12648447
      End
   End
   Begin VB.Timer Timer3 
      Left            =   1620
      Top             =   1020
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "CXM0050"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   構成マスタメンテナンス
'**       フォームID    :   CXM0050
'**       処理概要      :
'**
'**       作  成  日    :   2003/11/25  By CIS SIMPR-A
'**       変  更  日    :   2004/04/05  不具合修正(切替ボタンをボディ部のみ有効)
'**       変  更  日    :   2004/04/16  バグ修正
'**       変  更  日    :   2004/07/01  バグ修正
'**       　  　  　                    > 1.品番・複写先で検索ﾎﾞﾀﾝ押下でエラー発生
'**       　  　  　                    > 2.明細の複写時、そのまま実行でエラー発生（キーの重複）
'**       　  　  　                    > 3.明細の修正時、構成品番修正でエラー発生（キーの重複）
'**       変  更  日    :   2005/01/18  削除の時、品番マスタに存在しなくても削除可能とする
'**       変  更  日    :   2007/02/08  ツリー表示時に材質が18桁を越えるとエラーになるのを修正
'**       変  更  日    :   2008/07/31  同じ構成品を違う親品番に連続で追加すると重複エラーとなる
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    Dim SyoriKB2        As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM2        As String       ' 処理名称(新規･修正･削除 等)
    
    Dim mCHK            As Boolean
    Dim RowColFlg       As Boolean
    Dim GridBack        As String
    
    Dim DayString()     As String
    
    Dim mEdit           As String

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
    Dim CisVsGrid1      As New CisVsGrid3
    Dim CisVsGridT      As New CisVsGrid3
    Dim ID              As Long
    Dim mIx             As Long
    Dim mIx2            As Long
    Dim mStr            As String
    
    Dim ActObj          As Object
    Dim ActRow          As Integer
    
    Dim KCol            As Long         ' 非表示項目開始ｶﾗﾑ
    
    Dim SCol()          As Long         ' 表示項目ｶﾗﾑ(品番)
    Dim SColT()         As Long         ' 表示項目ｶﾗﾑ(材料)
    
    Dim wsOHinbn        As String
    Dim wsKHinbn        As String
    Dim CPYKB           As Byte
'   < TreeView 系 >
    Dim oNod            As Node
    Dim indrag          As Boolean      ' ドラッグ アンド ドロップ操作中かどうかを示すフラグを作成します。
    Dim oImg            As ListImage
    Dim wkNod           As Node
    Dim wStr            As String

    Dim gIdx1           As Long
    Dim gIdx2           As Long
    Dim gIdx3           As Long
    Dim wFLG            As Boolean
    Dim wTIdx           As Long
    
    Dim sSvOHinbn       As String
    Dim sSvKHinbn       As String
    Dim sSvRow          As Long
    Dim MouseButton     As Integer
    
' マスタ用
    Dim wsv_HinKB       As String
    Dim wsv_ROW         As Long
    Dim wsv_Hinbn       As String
    Dim wsv_Itatu       As Double
    Dim wsv_Width       As Double
    Dim wsv_Long        As Double
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
'
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
'
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
'
    Dummy.Left = -1000
    Dummy.Enabled = False
'    Dummy2.Left = -1000
'    Dummy3.Left = -1000
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔(基本)
    Timer2.Enabled = False   ' 処理区分点滅禁止(基本)
'    Timer3.Interval = 500    ' 処理区分点滅間隔(取引先)
'    Timer3.Enabled = False   ' 処理区分点滅禁止(取引先)
'
    '+--------------------------------+
    '+  グリッド初期設定(構成)
    '+--------------------------------+
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 25
        .InitGet = False
        .Init
'        .FixedHeight = 260
    End With
    Call GridInit1

    '+--------------------------------+
    '+  グリッド初期設定(手配)
    '+--------------------------------+
    With CisVsGridT
        Set .GridObj = VsGridT
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 16
        .InitGet = False
        .Init
    End With
    Call GridInitT

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB2 = ""
    Call SyoriName2(SyoriKB2)
    SyoriKB = "ADD"
    Call SyoriName(SyoriKB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
'
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

    sSvOHinbn = ""
    sSvKHinbn = ""
    sSvRow = 0

    With ImageList1.ListImages
        Set oImg = .Add(1, , LoadPicture(gIniExe & "BMP\" & "LV1.ICO"))
        Set oImg = .Add(2, , LoadPicture(gIniExe & "BMP\" & "LV2.ICO"))
        Set oImg = .Add(3, , LoadPicture(gIniExe & "BMP\" & "LV3.ICO"))
        Set oImg = .Add(4, , LoadPicture(gIniExe & "BMP\" & "LV4.ICO"))
        Set oImg = .Add(5, , LoadPicture(gIniExe & "BMP\" & "LV5.ICO"))
        Set oImg = .Add(6, , LoadPicture(gIniExe & "BMP\" & "LV6.ICO"))
        Set oImg = .Add(7, , LoadPicture(gIniExe & "BMP\" & "LV7.ICO"))
        Set oImg = .Add(8, , LoadPicture(gIniExe & "BMP\" & "LV8.ICO"))
        Set oImg = .Add(9, , LoadPicture(gIniExe & "BMP\" & "LV9.ICO"))
        Set oImg = .Add(10, , LoadPicture(gIniExe & "BMP\" & "LV0.ICO"))
        
        Set oImg = .Add(11, , LoadPicture(gIniExe & "BMP\" & "ZLV1.ICO"))
        Set oImg = .Add(12, , LoadPicture(gIniExe & "BMP\" & "ZLV2.ICO"))
        Set oImg = .Add(13, , LoadPicture(gIniExe & "BMP\" & "ZLV3.ICO"))
        Set oImg = .Add(14, , LoadPicture(gIniExe & "BMP\" & "ZLV4.ICO"))
        Set oImg = .Add(15, , LoadPicture(gIniExe & "BMP\" & "ZLV5.ICO"))
        Set oImg = .Add(16, , LoadPicture(gIniExe & "BMP\" & "ZLV6.ICO"))
        Set oImg = .Add(17, , LoadPicture(gIniExe & "BMP\" & "ZLV7.ICO"))
        Set oImg = .Add(18, , LoadPicture(gIniExe & "BMP\" & "ZLV8.ICO"))
        Set oImg = .Add(19, , LoadPicture(gIniExe & "BMP\" & "ZLV9.ICO"))
        Set oImg = .Add(20, , LoadPicture(gIniExe & "BMP\" & "ZLV0.ICO"))
    
    End With
    TView1.ImageList = ImageList1
    TView1.Style = tvwTreelinesPlusMinusPictureText     ' スタイルを 7 に設定します。
' *ADD* 2004/04/16
    mGotColor = &HC0FFC0
    mLostColor = &HC0C0C0

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
           Case vbKeyF1:
                If PB_ADD.Visible Then
                    Call PB_ADD_Click    '【新規】
                End If
'                If ProcHB = "B3" And iB3_No.Visible Then
'                   On Error Resume Next
'                   If ActiveControl.Name = "iB3_No" Then
'                      iB3_Seban.SetFocus
'                   Else
'                      iB3_No.SetFocus
'                   End If
'                   On Error GoTo 0
'                End If
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF5:       If PB_CPY.Visible Then Call PB_CPY_Click    '【複写】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF9:       If PB_Hin.Visible Then Call PB_Hin_Click    '【マスタ】
           Case vbKeyF10:      If PB_Zai.Visible Then Call PB_Zai_Click    '【関連付】
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
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
   
    If ProcHB = "B1" Then
        If SyoriKB2 <> "ADD" Then
            ProcHB = "B2"
            Call DispChange(ProcHB)
            Call HeadBodyClear("B1")
    '        Call SyoriName("")
'            VsGrid1.SetFocus
            vsGrid1.SetFocus
            Call VSGrid1_Click
            GoTo PB_CAN_Ed
        Else
            ProcHB = "H1"
            Call DispChange(ProcHB)
            SyoriKB2 = ""
            Call SyoriName2(SyoriKB2)
            Call HeadBodyClear("B")
            Call SyoriName(SyoriKB)
            GoTo PB_CAN_Ed
        End If
    End If
            
    If ProcHB = "B2" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        SyoriKB2 = ""
        Call SyoriName2(SyoriKB2)
        Call HeadBodyClear("B")
        Call SyoriName(SyoriKB)
        GoTo PB_CAN_Ed
    End If

PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新　規  キ ー(F1)        +
'+----------------------------+
Private Sub PB_ADD_Click()
    
    If ProcHB = "H1" Then
        If SyoriKB <> "ADD" Then
            SyoriKB = "ADD"
            Call SyoriName(SyoriKB)
        End If
    Else
        If SyoriKB2 <> "ADD" Then
            
            ProcHB = "B1"
            Call DispChange(ProcHB)
            Call HeadBodyClear("B1")
            B1Op_Tenkb(0).Value = True
            SyoriKB2 = "ADD"
            Call SyoriName2(SyoriKB2)
'            Call CisVsGrid1.vsColor(消去全部)
            iB1_KHinbn.SetFocus
        End If
    End If
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)        +
'+----------------------------+
Private Sub PB_MNT_Click()
    
    If ProcHB = "H1" Then
        If SyoriKB <> "MNT" Then
            SyoriKB = "MNT"
            Call SyoriName(SyoriKB)
        End If
    Else
        If SyoriKB2 <> "MNT" Then
            ProcHB = "B2"
            Call DispChange(ProcHB)
            SyoriKB2 = "MNT"
            Call SyoriName2(SyoriKB2)
'            Call CisVsGrid1.vsColor(消去全部)
            vsGrid1.SetFocus
            Call VSGrid1_Click
        End If
    End If
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)        +
'+----------------------------+
Private Sub PB_DEL_Click()
    If ProcHB = "H1" Then
        If SyoriKB <> "DEL" Then
            SyoriKB = "DEL"
            Call SyoriName(SyoriKB)
        End If
    Else
        If SyoriKB2 <> "DEL" Then
            ProcHB = "B2"
            SyoriKB2 = "DEL"
            Call DispChange(ProcHB)
            Call SyoriName2(SyoriKB2)
'            Call CisVsGrid1.vsColor(消去全部)
            
            vsGrid1.SetFocus
            Call VSGrid1_Click
        End If
    End If

End Sub
'+----------------------------+
'+    複  写  キ ー(F5)        +
'+----------------------------+
Private Sub PB_CPY_Click()
    If ProcHB = "H1" Then
        If SyoriKB <> "CPY" Then
            SyoriKB = "CPY"
            Call SyoriName(SyoriKB)
        End If
    Else
        If SyoriKB2 <> "CPY" Then
            
            ProcHB = "B2"
            SyoriKB2 = "CPY"
            Call DispChange(ProcHB)
            Call SyoriName2(SyoriKB2)
'            Call CisVsGrid1.vsColor(消去全部)
            
            vsGrid1.SetFocus
            Call VSGrid1_Click
        End If
    End If
    
    
End Sub
'+----------------------------+
'+    品番・材料マスタ(F9)
'+----------------------------+
Private Sub PB_Hin_Click()
    RV_Left = 0
    RV_Top = 0
    Rv_Call3 = "KOM"
    
    wsv_ROW = vsGrid1.ROW
    If RTrim(vsGrid1.TextMatrix(wsv_ROW, SCol(13))) = "" Then
        RV_Hinbn = RTrim(vsGrid1.TextMatrix(wsv_ROW, SCol(12)))
    
        CXM0030.Show vbModal
        Unload CXM0030
        Set CXM0030 = Nothing
    
        Call GridSet_TEH(wsv_ROW)
    Else
        RV_Zaist = RTrim(vsGrid1.TextMatrix(wsv_ROW, SCol(15)))
        RV_Itatu = CisFun.Val2(vsGrid1.TextMatrix(wsv_ROW, SCol(16)))
        RV_Width = CisFun.Val2(vsGrid1.TextMatrix(wsv_ROW, SCol(17)))
        RV_Long = CisFun.Val2(vsGrid1.TextMatrix(wsv_ROW, SCol(18)))
    
        CXM0070.Show vbModal
        Unload CXM0070
        Set CXM0070 = Nothing
    
        Call GridSet_TEH(wsv_ROW)
    End If
    
End Sub
'+----------------------------+
'+    品番・材料関連付け(F10)
'+----------------------------+
Private Sub PB_Zai_Click()
    RV_Left = 0
    RV_Top = 0
    
    wsv_ROW = vsGrid1.ROW
    If RTrim(vsGrid1.TextMatrix(wsv_ROW, SCol(13))) = "" Then
        Rv_Call2 = "HIN"
        RV_Hinbn = RTrim(vsGrid1.TextMatrix(wsv_ROW, SCol(12)))
    Else
        Rv_Call2 = "ZAI"
        RV_Zaist = RTrim(vsGrid1.TextMatrix(wsv_ROW, SCol(15)))
        RV_Itatu = CisFun.Val2(vsGrid1.TextMatrix(wsv_ROW, SCol(16)))
        RV_Width = CisFun.Val2(vsGrid1.TextMatrix(wsv_ROW, SCol(17)))
        RV_Long = CisFun.Val2(vsGrid1.TextMatrix(wsv_ROW, SCol(18)))
    End If
        
    CXM0110.Show vbModal
    Unload CXM0110
    Set CXM0110 = Nothing
    
    Call GridSet_KOM(0)
End Sub


'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = " SELECT 親品番"
    gSL_Select = gSL_Select & " FROM 構成マスタ"
    gSL_Select = gSL_Select & " WHERE 親品番 < '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " GROUP BY 親品番 "
    gSL_Select = gSL_Select & " ORDER BY 親品番 DESC"
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Hinbn.SetFocus
    Dummy.Enabled = False
    
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = " SELECT 親品番"
    gSL_Select = gSL_Select & " FROM 構成マスタ"
    gSL_Select = gSL_Select & " WHERE 親品番 > '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " GROUP BY 親品番 "
    gSL_Select = gSL_Select & " ORDER BY 親品番 "
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Hinbn.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn(0): Exit Sub
    If PB_Look.Tag = "iH1_CHinbn" Then Call Look_Hinbn(2): Exit Sub
    If PB_Look.Tag = "iB1_KHinbn" Then Call Look_Hinbn(1): Exit Sub
'
    If PB_Look.Tag = "iB1_SDate" Then
        iB1_SDate.ShowCalender
        Exit Sub
    End If
    If PB_Look.Tag = "iB1_EDate" Then
        iB1_EDate.ShowCalender
        Exit Sub
    End If

End Sub
'+----------------------------+
'+       品番検索              +
'+----------------------------+
Private Sub Look_Hinbn(KB As Byte)

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
'*--------------------------------------------------<< 2004/07/01 Update Start >>
' *ADD* 2004/04/16
'   If KB = 0 Then
'       RVI_Hinmoku = 9
'   Else
'       RVI_Hinmoku = KB
'   End If
'   >> 品目の設定値 .. 0/1/9
    Select Case KB
        Case 0: RVI_Hinmoku = 9         '>品目：ﾁｪｯｸ無
        Case 1: RVI_Hinmoku = 1         '>品目：ﾁｪｯｸ有（発注）
        Case 2: RVI_Hinmoku = 9         '>品目：ﾁｪｯｸ無
    End Select
'*--------------------------------------------------<< 2004/07/01 Update Start >>
    
    CKK0030.Show vbModal
    Unload CKK0030
    Set CKK0030 = Nothing

    If RV_Rtn Then
        Select Case KB
            Case 0
                iH1_Hinbn = RV_Hinbn
            Case 1
                iB1_KHinbn = RV_Hinbn
                Call iB1_KHinbn_LostFocus
            Case 2
                iH1_CHinbn = RV_Hinbn
        End Select
    End If
End Sub
'+---( H1 )----------------------------------------------------------------------------
'+-----------------------------+
'+  iH1_Hinbn
'+-----------------------------+
Private Sub iH1_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_CHinbn
'+-----------------------------+
Private Sub iH1_CHinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_CHinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+---( B1 )----------------------------------------------------------------------------
'+-----------------------------+
'+  iB1_KHinbn
'+-----------------------------+
Private Sub iB1_KHinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_KHinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    B1lb_Syakei = "": B1lb_SSeban = ""
    gSL_Select = "SELECT * FROM 品番マスタ "
    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iB1_KHinbn) & "'"
    If HNMRead(gSL_Select, 1) Then
        B1lb_Syakei = RTrim(HNM.車型)
        B1lb_SSeban = RTrim(HNM.社内背番)
    End If
End Sub
' *ADD* 2004/04/16
'+-----------------------------+
'+  B1Op_Tenkb
'+-----------------------------+
Private Sub B1Op_Tenkb_Click(Index As Integer)
    B1Op_Tenkb(0).BackColor = mLostColor
    B1Op_Tenkb(1).BackColor = mLostColor
    
    B1Op_Tenkb(Index).BackColor = mGotColor
End Sub
'+-----------------------------+
'+  B1Chk_Tenkb
'+-----------------------------+
Private Sub B1Chk_Tenkb_Click()
    If B1Chk_Tenkb.Value = 0 Then
        B1Chk_Tenkb.Caption = " 行 う"
    Else
        B1Chk_Tenkb.Caption = " 行わない"
    End If
End Sub
'+-----------------------------+
'+  iB1_SDate
'+-----------------------------+
Private Sub iB1_SDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_SDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_EDate
'+-----------------------------+
Private Sub iB1_EDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_EDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
'+------------------------------------------------------------------+
'+      基本情報
'+------------------------------------------------------------------+
Private Sub SyoriName(SYKB As String)
    Dim W_Color     As Long
    W_Color = &H80000008 '初期色ｾｯﾄ(黒)
   
'   H2LB_Size_B.Enabled = False
    H1_Area1.Height = 645
    HLB_CHinbn.Visible = False
    HLB_Hinbn.Caption = "品  番"
    
    Select Case SyoriKB
      Case "ADD"
           SyoriNM = "追加"
           K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
               
'          H2LB_Size_B.Enabled = True
      Case "MNT"
           SyoriNM = "修正"
           K_Sykbnm.ForeColor = &HFF0000  '(青色)
      
      Case "DEL"
           SyoriNM = "削除"
           K_Sykbnm.ForeColor = &HFF&     '(赤色)
           W_Color = &HFF&
      Case "LOK"
           SyoriNM = "照会"
           K_Sykbnm.ForeColor = &HFF00FF   '(ピンク)
           PB_ADD.Visible = False
           PB_MNT.Visible = False
           PB_DEL.Visible = False
      Case "CPY"
           SyoriNM = "複写"
           K_Sykbnm.ForeColor = &HFF00FF   '(ピンク)
            H1_Area1.Height = 1125
            HLB_CHinbn.Visible = True
            HLB_Hinbn.Caption = "複写元"
      Case Else
           SyoriNM = ""
    End Select
    For Each gObj In Me.Controls
        If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
        If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
'      * PF-Key 表示設定 *
        If SyoriKB <> "LOC" Then
           If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
              gObj.Name = "PB_DEL" Or gObj.Name = "PB_CPY" Then
              gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
           End If
        End If
    Next gObj
   
'    If RTrim(PB_Chg.Tag) <> "Hin" Then: PB_DEL.Visible = False
    Timer2.Enabled = False
    If SyoriNM = "" Then
       K_Sykbnm.Visible = False
       K_Sykbnm.Caption = ""
       Exit Sub
    End If
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    Timer2.Enabled = True
End Sub
'+------------------------------------------------------------------+
'+      詳細情報
'+------------------------------------------------------------------+
Private Sub SyoriName2(SYKB As String)
   Dim W_Color     As Long
   W_Color = &H80000008 '初期色ｾｯﾄ(黒)
   
'   H2LB_Size_B.Enabled = False
'    B1LB_Back.Enabled = True
   
   Select Case SyoriKB2
     Case "ADD"
          SyoriNM2 = "追加"
          K_Sykbnm2.ForeColor = &HFFFF&   '(黄色)
          PB_ADD.Visible = False
              
'          H2LB_Size_B.Enabled = True
     Case "MNT"
          SyoriNM2 = "修正"
          K_Sykbnm2.ForeColor = &HFF0000  '(青色)
     
'          B1LB_Back.Enabled = False
     Case "DEL"
          SyoriNM2 = "削除"
          K_Sykbnm2.ForeColor = &HFF&     '(赤色)
          W_Color = &HFF&
          
     Case "LOK"
          SyoriNM2 = "照会"
          K_Sykbnm2.ForeColor = &HFF00FF   '(ピンク)
          PB_ADD.Visible = False
          PB_MNT.Visible = False
          PB_DEL.Visible = False
     Case "CPY"
          SyoriNM2 = "複写"
          K_Sykbnm2.ForeColor = &HFF00FF   '(ピンク)
     Case Else
          SyoriNM2 = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
'      * PF-Key 表示設定 *
       If SyoriKB2 <> "LOC" Then
          If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
             gObj.Name = "PB_DEL" Or gObj.Name = "PB_CPY" Then
             gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB2)
          End If
       End If
   Next gObj
   Timer3.Enabled = False
   If SyoriNM2 = "" Then
      K_Sykbnm2.Visible = False
      K_Sykbnm2.Caption = ""
      Exit Sub
   End If
   K_Sykbnm2.Caption = SyoriNM2
   K_Sykbnm2.Visible = True
   Timer3.Enabled = True
End Sub

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*      処理区分名点滅制御      *
'*******************************
Private Sub Timer2_Timer()
   If K_Sykbnm.Caption <> "" Then
      K_Sykbnm.Caption = ""
   Else
      K_Sykbnm.Caption = SyoriNM
   End If
End Sub
'*******************************
'*      処理区分名点滅制御      *
'*******************************
Private Sub Timer3_Timer()
   If K_Sykbnm2.Caption <> "" Then
      K_Sykbnm2.Caption = ""
   Else
      K_Sykbnm2.Caption = SyoriNM2
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
    Dim B2Mode      As Boolean
    Dim B2Color     As Long
    Dim B3Mode      As Boolean
    Dim B3Color     As Long
    
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
    B2_Area1.Enabled = Not H1Mode
    B2_Area1.BackColor = B2Color
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
' *CHG* 2004/04/05
    PB_Hin.Visible = Not H1Mode
    PB_Hin.Visible = Not H1Mode
'    PB_Chg.Visible = H1Mode
'    PB_ZLook.Visible = B1Mode
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
    RowColFlg = True
    
    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    
    On Error GoTo HeadBodyClear_ER
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
              '  .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
'            If .Name Like wbc_Enm & "Op_*" Then
'                .Value = False
'                .BackColor = mLostColor
'            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    
    
    If HeadBody = "B2" Or HeadBody = "B" Or HeadBody = "ALL" Then
        CisVsGrid1.Clear
        CisVsGridT.Clear
    End If
    If HeadBody = "B" Or HeadBody = "ALL" Then
        TView1.Nodes.Clear
    End If
    RowColFlg = False
    Exit Sub
HeadBodyClear_ER:
    RowColFlg = False
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        
' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            GoTo ReturnPress_Ed
        End If
        
' << 削除処理 >>
        If SyoriKB = "CPY" Then
            If Not DBCopy Then GoTo ReturnPress_Ed
        End If
        
        ProcHB = "B1"
        Call DispChange(ProcHB)
' *ADD* 2004/04/16
        B1Op_Tenkb(0).Value = True
        
        SyoriKB2 = "ADD"
        Call SyoriName2(SyoriKB2)
        Call VSGrid1_Click
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed

        Call HeadBodyClear("B1")
        
        Call GridSet_KOM(CPYKB)
        If SyoriKB2 <> "ADD" Then
            ProcHB = "B2"
            Call DispChange(ProcHB)
            vsGrid1.SetFocus
            Call VSGrid1_Click
        End If
        
        RowColFlg = True
        Call CisVsGrid1.vsColor(消去全部)
        Call CisVsGrid1.vsColor(個別)
        RowColFlg = False
        
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B2" Then
'        If Not Body2Chk Then GoTo ReturnPress_Ed
' << 削除処理 >>
        If SyoriKB2 = "DEL" Then
            If Not DBDeleteM Then
               RowColFlg = True
               Call CisVsGrid1.vsColor(消去全部)
               RowColFlg = False
               GoTo ReturnPress_Ed
            End If
            Call HeadBodyClear("B2")
            Call GridSet_KOM(CPYKB)
            Call DispChange(ProcHB)
            Call SyoriName2(SyoriKB2)
        Else
            GridBack = ProcHB
            Call VSGrid1_Click
            Call vsGrid1_DblClick
'            ProcHB = "B1"
'            Call DispChange(ProcHB)
        End If
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
    
    CPYKB = 0
    ' 品番
    If RTrim(iH1_Hinbn) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    品番を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Hinbn.SetFocus
       Exit Function
    Else
'------<< 削除の時、品番マスタに存在しなくても削除可能とする >>--------2005/01/18 (START)-----'
        If SyoriKB <> "DEL" Then
            gSL_Select = "SELECT * FROM 品番マスタ "
            gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
            If Not HNMRead(gSL_Select, 1) Then
                With CisFun
                     .MB_Lines = 4
                     .MB_MSG(2) = "    品番 未登録！          "
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                End With
                iH1_Hinbn.SetFocus
                Exit Function
            End If
        End If
    End If
'
    gSL_Select = "SELECT * "
    gSL_Select = gSL_Select & " FROM 構成マスタ "
    gSL_Select = gSL_Select & " WHERE 親品番 = '" & RTrim(iH1_Hinbn) & "'"
    If Not DBInput("INV") Then
        If SyoriKB <> "ADD" Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    構成マスタ 未登録            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Hinbn.SetFocus
            Exit Function
        End If
    Else
        If SyoriKB = "ADD" Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    構成マスタ 登録済            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Hinbn.SetFocus
            Exit Function
        End If
    End If

' 複写先チェック
    If SyoriKB <> "CPY" Then GoTo Head1Chk_ED
    
    ' 品番
    If RTrim(iH1_CHinbn) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    複写先・品番を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_CHinbn.SetFocus
       Exit Function
    Else
        gSL_Select = "SELECT * FROM 品番マスタ "
        gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_CHinbn) & "'"
        If Not HNMRead(gSL_Select, 1) Then
            With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    複写先・品番 未登録！          "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
            End With
            iH1_CHinbn.SetFocus
            Exit Function
        End If
    End If
    
    gSL_Select = "SELECT * FROM 構成マスタ "
    gSL_Select = gSL_Select & " WHERE 親品番 = '" & RTrim(iH1_CHinbn) & "'"
    If KOMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "    複写先・構成マスタ 登録済み           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_CHinbn.SetFocus
        Exit Function
    End If

Head1Chk_ED:
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    
    Body1Chk = False
' 構成品番
    If Trim(iB1_KHinbn) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    構成品番を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_KHinbn.SetFocus
       Exit Function
    Else
        gSL_Select = "SELECT * FROM 品番マスタ "
        gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iB1_KHinbn) & "'"
        
        If Not HNMRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    構成品番 未登録！            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB1_KHinbn.SetFocus
            Exit Function
        End If
    
    End If
' 使用数
    If CisFun.Val2(iB1_Siyou) = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    使用数を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_Siyou.SetFocus
       Exit Function
    End If
    ' 有効期限
    If RTrim(iB1_SDate) <> "" Then
        If Not iB1_SDate.cChkResult Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    有効期限 開始日エラー              "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB1_SDate.SetFocus
           Exit Function
        End If
    End If
    
    If RTrim(iB1_EDate) <> "" Then
        If Not iB1_EDate.cChkResult Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    有効期限 終了日エラー              "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB1_EDate.SetFocus
           Exit Function
        End If
        
        If RTrim(iB1_SDate) > RTrim(iB1_EDate) Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    有効期限エラー              "
                .MB_MSG(4) = "    【開始＞終了】"
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB1_SDate.SetFocus
           Exit Function
        End If
    End If

    
'*--------------------------------------------------<< 2004/07/01 Update Start >>
'   If SyoriKB2 = "ADD" Then
    If (SyoriKB2 = "ADD") Or _
       (SyoriKB2 = "CPY") Or _
       (SyoriKB2 = "MNT" And wsKHinbn <> iB1_KHinbn) Then
'*--------------------------------------------------<< 2004/07/01 Update End >>
        wsOHinbn = vsGrid1.TextMatrix(vsGrid1.ROW, SCol(12))            '2008/07/31修正
        
        gSL_Select = "SELECT * FROM 構成マスタ "
        gSL_Select = gSL_Select & " WHERE 親品番 = '" & RTrim(wsOHinbn) & "'"
        gSL_Select = gSL_Select & "   AND 子品番 = '" & RTrim(iB1_KHinbn) & "'"
        If KOMRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     既に同一の構成品番が登録されています。      "
                .MB_Title = "構成確認"
                .MB_Button = Error
                .MBOX
            End With
            iB1_KHinbn.SetFocus
            Exit Function
        End If
    End If
    
    With CisDB
        If SyoriKB2 = "ADD" Then
            If RTrim(vsGrid1.TextMatrix(vsGrid1.ROW, SCol(13))) <> "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "     材料に構成品番を登録する事はできません      "
                    .MB_Title = "構成確認"
                    .MB_Button = Error
                    .MBOX
                End With
                iB1_KHinbn.SetFocus
                Exit Function
            End If
            
            wsOHinbn = vsGrid1.TextMatrix(vsGrid1.ROW, SCol(12))
            wsKHinbn = ""
'        Else
'            wsOHinbn = VsGrid1.TextMatrix(VsGrid1.Row, SCol(11))
'            wsKHinbn = VsGrid1.TextMatrix(VsGrid1.Row, SCol(12))
        End If
        
        
        .SQL = "構成マスタメンテ入力チェック"
        .StoadoCount = 3
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = Input用
        .ParaNo = 3: .ParaIO = Input用
    
        .ParaNo = 3: .ParaValue = ""
        .ParaNo = 1: .ParaValue = RTrim(wsOHinbn)
        If RTrim(wsKHinbn) <> "" Then
            .ParaNo = 2: .ParaValue = RTrim(wsKHinbn)
            If RTrim(wsKHinbn) <> RTrim(iB1_KHinbn) Then
                .ParaNo = 3: .ParaValue = RTrim(iB1_KHinbn)
            End If
        Else
            .ParaNo = 2: .ParaValue = RTrim(iB1_KHinbn)
        End If
        If Not .DBStored Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     構成入力チェック時にエラーが発生しました。    "
                .MB_Title = "構成確認"
                .MB_Button = Error
                .MBOX
                
                Exit Function
            End With
        End If
    
        .ParaNo = 0
        If .ParaValue <> 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     構成情報に矛盾が発生しました。       "
                .MB_Title = "構成確認"
                .MB_Button = Error
                .MBOX
            End With
            iB1_KHinbn.SetFocus
            Exit Function
        End If
    End With
    
    Body1Chk = True
End Function
'***********************************************************************************
'*
'*       データベース入力 ＆ 表示
'*
'***********************************************************************************
'+---------------------------------+
'+      基本情報
'+---------------------------------+
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("B")
    
    If Not KOMRead(gSL_Select, 1) Then
        If InputKey = "Back" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     先頭レコードです            "
                .MB_Title = "【後退】"
                .MB_Button = OK
                .MBOX
            End With
            GoTo DBInput_Ed
        End If
        If InputKey = "Next" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     最終レコードです            "
                .MB_Title = "【前進】"
                .MB_Button = OK
                .MBOX
            End With
            GoTo DBInput_Ed
        End If
    Else
        iH1_Hinbn = RTrim(KOM.親品番)
        Call iH1_Hinbn_LostFocus
        
        DBInput = True
    End If
    
    Call GridSet_KOM(CPYKB)
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------------+
'+      削 除 処 理            +
'+----------------------------+
Private Function DBDelete()
    DBDelete = False
'  ( 確認メッセージ )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     削除処理を行います      "
        .MB_Title = "削除処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
        .SQL = "DELETE FROM 構成マスタ "
        .SQL = .SQL & " WHERE 親品番 = '" & RTrim(iH1_Hinbn) & "'"
        .DBExec

        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + CisDB.RecordCount, "###")
        
        On Error GoTo 0
    End With
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------------+
'+      削 除 処 理            +
'+----------------------------+
Private Function DBDeleteM()
    DBDeleteM = False
'  ( 確認メッセージ )
    K_Sykbnm2.Caption = SyoriNM2
    K_Sykbnm2.Visible = True
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     削除処理を行います      "
        .MB_Title = "削除処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    With CisDB
        On Error GoTo DBDeleteM_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
        .SQL = "DELETE FROM 構成マスタ "
        .SQL = .SQL & " WHERE 親品番 = '" & RTrim(wsOHinbn) & "'"
        .SQL = .SQL & "   AND 子品番 = '" & RTrim(iB1_KHinbn) & "'"
        .DBExec

        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")
        
        On Error GoTo 0
    End With
    
    sSvRow = vsGrid1.ROW
    
    DBDeleteM = True
    Exit Function
DBDeleteM_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Dim AddFlg  As Boolean

    K_Sykbnm2.Caption = SyoriNM2
    K_Sykbnm2.Visible = True
' 複写処理
    If SyoriKB2 = "CPY" Then
        If Not DBCopyM Then Exit Function
        GoTo DBPut_ED
    End If
    
    With CisFun
        .MB_Lines = 8
        If SyoriKB2 = "ADD" Then
            .MB_MSG(2) = "     追加処理を行います                     "
            .MB_Title = "追加処理"
            wsOHinbn = RTrim(vsGrid1.TextMatrix(vsGrid1.ROW, SCol(12)))
        Else
            .MB_MSG(2) = "     修正処理を行います                     "
            .MB_Title = "修正処理"
        End If
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_MSG(5) = "     親品番   = " & RTrim(wsOHinbn)
        .MB_MSG(6) = "     構成品番 = " & RTrim(iB1_KHinbn)
        .MB_Button = OK_CAN
        If Not .MBOX Then
            iB1_KHinbn.SetFocus
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err

'   マスタ項目移送
    gSL_Select = "SELECT * From 構成マスタ"
    gSL_Select = gSL_Select & " WHERE 親品番 = '" & RTrim(wsOHinbn) & "'"
    gSL_Select = gSL_Select & "   AND 子品番 = '" & RTrim(wsKHinbn) & "'"
    If KOMRead(gSL_Select, 1) Then
        AddFlg = False
        KOM.変更前親品番 = RTrim(wsOHinbn)
        KOM.変更前子品番 = RTrim(wsKHinbn)
    Else
        Call ItemsClearKOM
        DoEvents
        AddFlg = True
        KOM.親品番 = RTrim(wsOHinbn)
    End If
    
    With KOM
        .子品番 = RTrim(iB1_KHinbn)
        .構成数 = CisFun.Val2(iB1_Siyou)
        
        If B1Op_Tenkb(0).Value = True Then
            If B1Chk_Tenkb.Value = 0 Then
                .展開区分 = 0
            Else
                .展開区分 = 1
            End If
        Else
            If B1Chk_Tenkb.Value = 0 Then
                .展開区分 = 3
            Else
                .展開区分 = 2
            End If
        End If
        .LT = CisFun.Val2(iB1_LT)
        .適用開始日 = RTrim(iB1_SDate)
        .適用終了日 = RTrim(iB1_EDate)
    End With

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 追加処理 )
    If AddFlg Then
        KOM.作成者 = gTanto
        Call KOMInsert
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
    Else
' ( 更新処理 )
        KOM.更新者 = gTanto
        Call KOMUpdate
        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
        
        If RTrim(wsKHinbn) <> RTrim(iB1_KHinbn) Then
            Call DBPut2(wsKHinbn, iB1_KHinbn)
        End If
    End If
'
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
DBPut_ED:
    sSvOHinbn = vsGrid1.TextMatrix(vsGrid1.ROW, SCol(11))
    If SyoriKB2 = "ADD" Then
        sSvKHinbn = vsGrid1.TextMatrix(vsGrid1.ROW, SCol(12))
    Else
        sSvKHinbn = RTrim(iB1_KHinbn)
    End If
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
Private Function DBPut2(Hinbn As String, CHinbn As String, Optional DBNo As Integer)
    DBPut2 = False

    On Error GoTo DBPut2_Err

    gSL_Select = "SELECT * FROM 構成マスタ"
    gSL_Select = gSL_Select & " WHERE 親品番 = '" & RTrim(CHinbn) & "'"
    If KOMRead(gSL_Select, 1) Then GoTo DBPut2_ED
        
    gSL_Select = "INSERT INTO 構成マスタ"
    gSL_Select = gSL_Select & "(親品番,子品番,構成数,"
    gSL_Select = gSL_Select & " 展開区分,適用開始日,適用終了日,"
    gSL_Select = gSL_Select & " 作成日,作成者,作成端末) "
    gSL_Select = gSL_Select & " SELECT '" & RTrim(CHinbn) & "',子品番,構成数,"
    gSL_Select = gSL_Select & "     展開区分,適用開始日,適用終了日, "
    gSL_Select = gSL_Select & "     GETDATE(),'" & gTanto & "',SUBSTRING(HOST_NAME(),1,20) "
    gSL_Select = gSL_Select & "     FROM 構成マスタ "
    gSL_Select = gSL_Select & "     WHERE 親品番 = '" & RTrim(Hinbn) & "'"

    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With

DBPut2_ED:
    DBPut2 = True
    Exit Function
DBPut2_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------------+
'+      複  写 処 理           +
'+----------------------------+
Private Function DBCopy()
    DBCopy = False
'  ( 確認メッセージ )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     複写処理を行います      "
        .MB_Title = "複写処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    With CisDB
        On Error GoTo DBCopy_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 複写処理 )
        Call DBPut2(iH1_Hinbn, iH1_CHinbn)
        
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
        iT1_Addsu = Format(CisFun.Val2(iT1_Delsu) + CisDB.RecordCount, "###")
        
        On Error GoTo 0
    End With
    
    CPYKB = 1
    Call GridSet_KOM(CPYKB)
    
    DBCopy = True
    Exit Function
DBCopy_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------------+
'+      複  写 処 理 (明細)    +
'+----------------------------+
Private Function DBCopyM()
    DBCopyM = False
'  ( 確認メッセージ )
    K_Sykbnm2.Caption = SyoriNM2
    K_Sykbnm2.Visible = True
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     複写処理を行います      "
        .MB_Title = "複写処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    With CisDB
        On Error GoTo DBCopyM_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    
        Call ItemsClearKOM
        KOM.親品番 = wsOHinbn
        KOM.子品番 = RTrim(iB1_KHinbn)
        KOM.構成数 = CisFun.Val2(iB1_Siyou)
        
        If B1Op_Tenkb(0).Value = True Then
            If B1Chk_Tenkb.Value = 0 Then
                KOM.展開区分 = 0
            Else
                KOM.展開区分 = 1
            End If
        Else
            If B1Chk_Tenkb.Value = 0 Then
                KOM.展開区分 = 3
            Else
                KOM.展開区分 = 2
            End If
        End If
        KOM.適用開始日 = RTrim(iB1_SDate)
        KOM.適用終了日 = RTrim(iB1_EDate)
        KOM.作成者 = gTanto
        Call KOMInsert
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")

' ( 複写処理 )
        Call DBPut2(wsKHinbn, iB1_KHinbn)
        
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
        iT1_Addsu = Format(CisFun.Val2(iT1_Delsu) + CisDB.RecordCount, "###")
        
        On Error GoTo 0
    End With
    
    sSvOHinbn = wsOHinbn
    sSvKHinbn = iB1_KHinbn
    Call GridSet_KOM(CPYKB)
    
    DBCopyM = True
    Exit Function
DBCopyM_Err:
    CisFun.ErrorBox
    End
End Function
'***************************************************************************************************************
'*      グリッド処理
'***************************************************************************************************************
'+---------------------------------------------------------------------------------------+
'+   構成情報   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+---------------------------------------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        ReDim SCol(25)

        Call .FixedSet(中中, 右中, あり, 3, 0, "No")
        Call .FixedSet(中中, 左中, あり, 7, 1, "レベル")
        Call .FixedSet(中中, 左中, あり, 20, 2, "品  番")
        Call .FixedSet(中中, 左中, あり, 21, 3, "品  名")
        Call .FixedSet(中中, 右中, あり, 10, 4, "使用数")
        Call .FixedSet(中中, 中中, あり, 5, 5, "車型")
        Call .FixedSet(中中, 左中, あり, 5, 6, "背番")
        Call .FixedSet(中中, 左中, あり, 10, 7, "展開区分")
        Call .FixedSet(中中, 左中, あり, 2, 8, "LT")
        Call .FixedSet(中中, 左中, あり, 21, 9, "適用範囲")
        Call .FixedSet(中中, 左中, あり, 0, 10, "順位")
        Call .FixedSet(中中, 左中, あり, 0, 11, "親品番")
        Call .FixedSet(中中, 左中, あり, 0, 12, "子品番")
        Call .FixedSet(中中, 左中, あり, 0, 13, "部品区分")
        Call .FixedSet(中中, 左中, あり, 0, 14, "材料管理番号")
        Call .FixedSet(中中, 左中, あり, 0, 15, "材種")
        Call .FixedSet(中中, 左中, あり, 0, 16, "形状")
        Call .FixedSet(中中, 左中, あり, 0, 17, "材質")
        Call .FixedSet(中中, 左中, あり, 0, 18, "板厚")
        Call .FixedSet(中中, 左中, あり, 0, 19, "幅")
        Call .FixedSet(中中, 左中, あり, 0, 20, "長さ")
        Call .FixedSet(中中, 左中, あり, 0, 22, "展開区")
        Call .FixedSet(中中, 左中, あり, 0, 23, "有効期限開始")
        Call .FixedSet(中中, 左中, あり, 0, 24, "有効期限終了")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, あり, 0, 25, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "1,2,3,4,7,9")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp

        SCol(1) = .FixedGet(1)     ' レベル
        SCol(2) = .FixedGet(2)     ' 品番
        SCol(3) = .FixedGet(3)     ' 品名
        SCol(4) = .FixedGet(4)     ' 使用数
        SCol(5) = .FixedGet(5)     ' 単位
        SCol(6) = .FixedGet(6)     ' 手配
        SCol(7) = .FixedGet(7)     ' 展開区分
        SCol(8) = .FixedGet(8)     ' 適用範囲
        SCol(9) = .FixedGet(9)     ' 順位
        SCol(10) = .FixedGet(10)   ' 親品番
        SCol(11) = .FixedGet(11)   ' 子品番
        SCol(12) = .FixedGet(12)   ' 部品区分
        SCol(13) = .FixedGet(13)   ' 材料管理番号
        SCol(14) = .FixedGet(14)   ' 材種
        SCol(15) = .FixedGet(15)   ' 形状
        SCol(16) = .FixedGet(16)   ' 材質
        SCol(17) = .FixedGet(17)   ' 板厚
        SCol(18) = .FixedGet(18)   ' 幅
        SCol(19) = .FixedGet(19)   ' 長さ
        SCol(20) = .FixedGet(20)   ' 展開区分
        SCol(21) = .FixedGet(21)   ' 有効期限開始
        SCol(22) = .FixedGet(22)   ' 有効期限終了
        SCol(23) = .FixedGet(23)   ' 材料区分
        SCol(24) = .FixedGet(24)   ' 材料区分
        
    End With
    KCol = 10   ' 非表示項目開始ｶﾗﾑ

End Sub
'+---------------------------------------------------------------------------------------+
'+      グリッド内容セット   0:通常、1:複写
'+---------------------------------------------------------------------------------------+
Private Sub GridSet_KOM(KB As Byte)
    
    Dim Ttl_Kensu       As Long
    Static wsROW        As Long

    wsROW = 1

    With CisDB
        .SQL = "構成マスタメンテ情報抽出"
        .StoadoCount = 1
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
    
        If KB = 0 Then
            .ParaNo = 1: .ParaValue = RTrim(iH1_Hinbn)
        Else
            .ParaNo = 1: .ParaValue = RTrim(iH1_CHinbn)
        End If
    
        If Not .DBStored Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     構成マスタ情報の抽出に失敗しました。     "
                .MB_Title = "抽出処理"
                .MB_Button = Error
                If Not .MBOX Then Exit Sub
            End With
        End If
    End With

    CisVsGrid1.Clear
    
    gSL_Select = "SELECT KW.*,"
    gSL_Select = gSL_Select & " HN.車型,HN.社内背番,HN.品名,HN.品番編集 "
    gSL_Select = gSL_Select & " FROM 構成メンテワーク KW "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN "
    gSL_Select = gSL_Select & "   ON HN.品番 = KW.子品番 "
    gSL_Select = gSL_Select & " WHERE KW.処理端末 = HOST_NAME() "
    gSL_Select = gSL_Select & " ORDER BY KW.順位 "
    If KOWRead(gSL_Select) Then

    ' 抽出処理 初期ｾｯﾄ
    '    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
        Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
    ' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
        With vsGrid1
           .Redraw = False
'            If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
'                .Rows = CisVsGrid1.DispMax
'            Else
                .Rows = Ttl_Kensu + 1
'            End If
        End With
        ID = 0
        Do Until Not KOW_RDSTS
            ID = ID + 1
                
            If sSvRow = 0 Then
                If sSvOHinbn = KOW.親品番 And sSvKHinbn = KOW.子品番 Then
                    If wsROW = 1 Then: wsROW = ID
                End If
            End If
                
            Call KOMInfoToGrid(ID)
            
            Call KOWReadNext
        Loop
    Else
        Call KOMInfoToGrid_ADD
    End If
    Call KOWClose
    
    With vsGrid1
        If Not .RowIsVisible(wsROW) Then: .TopRow = wsROW
         .ROW = wsROW: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  構成情報よりグリッドへセット
'+---------------------------------------------------------------------------------------+
Private Sub KOMInfoToGrid(SetRow As Long)
    With vsGrid1
        For mIx = 0 To .Cols - 1
            .TextMatrix(SetRow, mIx) = ""
        Next mIx
        
        ' No
        .TextMatrix(SetRow, 0) = SetRow
        
        ' レベル
        Select Case KOW.レベル
            Case 0: .TextMatrix(SetRow, SCol(1)) = "0"
            Case 1: .TextMatrix(SetRow, SCol(1)) = ".1"
            Case 2: .TextMatrix(SetRow, SCol(1)) = "..2"
            Case 3: .TextMatrix(SetRow, SCol(1)) = "...3"
            Case 4: .TextMatrix(SetRow, SCol(1)) = "....4"
            Case Else
                .TextMatrix(SetRow, SCol(1)) = "....." & Format(KOW.レベル, "##")
        End Select
        
        ' 品番
        .TextMatrix(SetRow, SCol(2)) = EditHinbn(KOW.子品番, KOW.品番編集)
        .TextMatrix(SetRow, SCol(11)) = RTrim(KOW.親品番)
        .TextMatrix(SetRow, SCol(12)) = RTrim(KOW.子品番)
        ' 品名
        .TextMatrix(SetRow, SCol(3)) = RTrim(KOW.品名)
        ' 使用数
        .TextMatrix(SetRow, SCol(4)) = CisFun.RSetFld(KOW.構成数, 10, "#,##0.0000")
        ' 車型
        .TextMatrix(SetRow, SCol(5)) = RTrim(KOW.車型)
        ' 社内背番
        .TextMatrix(SetRow, SCol(6)) = RTrim(KOW.社内背番)
        ' 展開区分
        Select Case KOW.展開区分
            Case 0: .TextMatrix(SetRow, SCol(7)) = ""
            Case 1: .TextMatrix(SetRow, SCol(7)) = "まで対象"
            Case 2: .TextMatrix(SetRow, SCol(7)) = "以降対象外"
            Case 3: .TextMatrix(SetRow, SCol(7)) = "のみ対象外"
        End Select
        .TextMatrix(SetRow, SCol(21)) = RTrim(KOW.展開区分)
        ' LT
        .TextMatrix(SetRow, SCol(8)) = CisFun.RSetFld(KOW.LT, 2, "##")
        ' 有効期限
        .TextMatrix(SetRow, SCol(22)) = KOW.適用開始日
        .TextMatrix(SetRow, SCol(23)) = KOW.適用終了日
        gStr = ""
        If Trim(KOW.適用開始日) <> "" Then
            gStr = Mid(KOW.適用開始日, 1, 4) & "/" & _
                   Mid(KOW.適用開始日, 5, 2) & "/" & _
                   Mid(KOW.適用開始日, 7, 2)
        End If
        If Trim(KOW.適用終了日) <> "" Then
           If gStr = "" Then gStr = Space(10)
           gStr = gStr & "～"
           gStr = gStr & Mid(KOW.適用終了日, 1, 4) & "/" & _
                         Mid(KOW.適用終了日, 5, 2) & "/" & _
                         Mid(KOW.適用終了日, 7, 2)
        Else
           If Trim(KOW.適用開始日) <> "" Then
              gStr = gStr & "～" & Space(10)
           End If
        End If
        .TextMatrix(SetRow, SCol(9)) = gStr
        ' 順位
        .TextMatrix(SetRow, SCol(10)) = RTrim(KOW.順位)
    End With
 
        
        With TView1.Nodes
'|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|~|
'
'               品番    基材品番    子品番              sw_Relative     sw_key
'       ( 親 )  A                                                      "親品" & A_
'               A       "" (A)      B   ( 基材 )        "親品" & A_    "基材" & AB
'               A       B           C                   "基材" & AB    "基子" & BC
'               A       B           D                   "基材" & AB    "基子" & BD
'               A       "" (A)      E   ( 子品番 )      "親品" & A_    "子品" & AE
'
'                                                                       _ : 空白(Space(18))
'
'|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|

'            gStr = RTrim$(HKM.基材品番)
'            If RTrim$(gStr) = "" Then gStr = RTrim$(HKM.品番)
'
'            Select Case sv_LVL
'                Case "基材":    sw_Relative = "親品"
'                Case "基子":    sw_Relative = "基材"
'                Case "子品":    sw_Relative = "親品"
'            End Select
'            sw_Relative = sw_Relative & RTrim$(HKM.品番) & Space(18 - Len(RTrim$(HKM.品番)))
'            sw_Relative = sw_Relative & RTrim$(HKM.基材品番) & Space(18 - Len(RTrim$(HKM.基材品番)))
'
'            sw_key = sv_LVL
'            sw_key = sw_key & RTrim$(gStr) & Space(18 - Len(RTrim$(gStr)))
'            sw_key = sw_key & RTrim$(HKM.子品番) & Space(18 - Len(RTrim$(HKM.子品番)))

'            Debug.Print sw_Relative, sw_key
            
'            If KOW.レベル = 0 Then
'                wSTR = EditHinbn(KOW.子品番, KOW.品番編集) & Space(40 - Len(RTrim(EditHinbn(KOW.子品番, KOW.品番編集))) - (KOW.レベル * 4)) & KOW.品名
'            Else
'                wSTR = EditHinbn(KOW.子品番, KOW.品番編集) & Space(40 - Len(RTrim(EditHinbn(KOW.子品番, KOW.品番編集))) - (KOW.レベル * 4) + 1) & KOW.品名
'            End If
            If KOW.レベル = 0 Then
                .Clear
                gStr = "A" & KOW.順位
                Set oNod = .Add(, , gStr, EditHinbn(KOW.子品番, KOW.品番編集), 10)
'                Set oNod = .Add(, , gStr, wSTR, 1)
            Else
                gStr = "A"
                If Len(KOW.順位) > 3 Then
                    gStr = "A" & CisFun.Mid2(KOW.順位, 1, Len(KOW.順位) - 3)
                End If
                If KOW.レベル > 9 Then
                    Set oNod = .Add(gStr, tvwChild, "A" & KOW.順位, EditHinbn(KOW.子品番, KOW.品番編集), 10)
'                    Set oNod = .Add(gStr, tvwChild, "A" & KOW.順位, wSTR, 1)
                Else
                    Set oNod = .Add(gStr, tvwChild, "A" & KOW.順位, EditHinbn(KOW.子品番, KOW.品番編集), KOW.レベル)
'                    Set oNod = .Add(gStr, tvwChild, "A" & KOW.順位, wSTR, KOW.レベル)
                End If
            End If
'                       Add([relative], [relationship], [key], text, [image[, selectedimage]])
            oNod.Expanded = True
        End With
' 材料セット
        Call KOMInfoToGrid_ZAI(KOW.レベル, KOW.子品番, KOW.順位)

End Sub
'   (追加用)
Private Sub KOMInfoToGrid_ADD()
    With vsGrid1
        For mIx = 0 To .Cols - 1
            .TextMatrix(1, mIx) = ""
        Next mIx
        
        gSL_Select = "SELECT * FROM 品番マスタ "
        gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
        Call HNMRead(gSL_Select, 1, 1)
        ' No
        .TextMatrix(1, 0) = 1
        ' レベル
        .TextMatrix(1, SCol(1)) = "0"
        ' 品番
        .TextMatrix(1, SCol(2)) = EditHinbn(iH1_Hinbn, HNM.品番編集)
        .TextMatrix(1, SCol(11)) = ""
        .TextMatrix(1, SCol(12)) = RTrim(iH1_Hinbn)
        ' 品名
        .TextMatrix(1, SCol(3)) = RTrim(HNM.品名)
        ' 使用数
        .TextMatrix(1, SCol(4)) = CisFun.RSetFld(1, 10, "#,##0.0000")
        ' 車型
        .TextMatrix(1, SCol(5)) = RTrim(HNM.車型)
        ' 社内背番
        .TextMatrix(1, SCol(6)) = RTrim(HNM.社内背番)
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  材料情報よりグリッドへセット
'+---------------------------------------------------------------------------------------+
Private Sub KOMInfoToGrid_ZAI(LEVEL As Byte, Hinbn As String, JUN As String)

    gSL_Select = "SELECT ZH.*, "
    gSL_Select = gSL_Select & " ZR.材質,ZR.材料名 品名, "
    gSL_Select = gSL_Select & " ZR.板厚,ZR.幅,ZR.長さ "
    gSL_Select = gSL_Select & " FROM 材料品番マスタ ZH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料マスタ ZR "
    gSL_Select = gSL_Select & "   ON ZH.材料管理番号 = ZR.材料管理番号"
    gSL_Select = gSL_Select & " WHERE ZH.使用品番 = '" & RTrim(Hinbn) & "'"
    If ZHMRead(gSL_Select, , 1) Then
        
        vsGrid1.Rows = vsGrid1.Rows + CisDB.RecordCount
        gInt = 0
        Do Until Not ZHM_RDSTS

' グリッド
            With vsGrid1
                ID = ID + 1
                
                For mIx = 0 To .Cols - 1
                    .TextMatrix(ID, mIx) = ""
                Next mIx
                
                ' No
                .TextMatrix(ID, 0) = ID
                
                ' レベル
                Select Case LEVEL + 1
                    Case 0: .TextMatrix(ID, SCol(1)) = "0"
                    Case 1: .TextMatrix(ID, SCol(1)) = ".1"
                    Case 2: .TextMatrix(ID, SCol(1)) = "..2"
                    Case 3: .TextMatrix(ID, SCol(1)) = "...3"
                    Case 4: .TextMatrix(ID, SCol(1)) = "....4"
                    Case Else
                        .TextMatrix(ID, SCol(1)) = "....." & Format(LEVEL + 1, "##")
                End Select
                
                ' 品番
                .TextMatrix(ID, SCol(2)) = RTrim(ZHM.材質)
                .TextMatrix(ID, SCol(11)) = RTrim(Hinbn)
                .TextMatrix(ID, SCol(12)) = RTrim(ZHM.材質)
                ' 品名
                .TextMatrix(ID, SCol(3)) = RTrim(ZHM.品名)
                ' 使用数
                .TextMatrix(ID, SCol(4)) = CisFun.RSetFld(ZHM.使用量, 10, "#,##0.0000")
                ' 展開区分
                If ZHM.展開区分 <> 0 Then: .TextMatrix(ID, SCol(7)) = "行わない"
                .TextMatrix(ID, SCol(21)) = RTrim(ZHM.展開区分)
                ' LT
                .TextMatrix(ID, SCol(8)) = CisFun.RSetFld(ZHM.LT, 2, "##")
                ' 有効期限
                .TextMatrix(ID, SCol(22)) = ZHM.有効期限開始
                .TextMatrix(ID, SCol(23)) = ZHM.有効期限終了
                gStr = ""
                If Trim(ZHM.有効期限開始) <> "" Then
                    gStr = Mid(ZHM.有効期限開始, 1, 4) & "/" & _
                           Mid(ZHM.有効期限開始, 5, 2) & "/" & _
                           Mid(ZHM.有効期限開始, 7, 2)
                End If
                If Trim(ZHM.有効期限終了) <> "" Then
                   If gStr = "" Then gStr = Space(10)
                   gStr = gStr & "～"
                   gStr = gStr & Mid(ZHM.有効期限終了, 1, 4) & "/" & _
                                 Mid(ZHM.有効期限終了, 5, 2) & "/" & _
                                 Mid(ZHM.有効期限終了, 7, 2)
                Else
                   If Trim(ZHM.有効期限開始) <> "" Then
                      gStr = gStr & "～" & Space(10)
                   End If
                End If
                .TextMatrix(ID, SCol(9)) = gStr
                ' 順位
                gInt = gInt + 1
                .TextMatrix(ID, SCol(10)) = RTrim(JUN) & Format(gInt, "00")
                            
                ' 部品区分
                .TextMatrix(ID, SCol(13)) = "X"
                ' 材料管理番号
                .TextMatrix(ID, SCol(14)) = RTrim(ZHM.材料管理番号)
                ' 材質
                .TextMatrix(ID, SCol(15)) = RTrim(ZHM.材質)
                ' 板厚
                .TextMatrix(ID, SCol(16)) = RTrim(ZHM.板厚)
                ' 幅
                .TextMatrix(ID, SCol(17)) = RTrim(ZHM.幅)
                ' 長さ
                .TextMatrix(ID, SCol(18)) = RTrim(ZHM.長さ)
            
            End With
         
' ツリー
            With TView1.Nodes
'------------------------------------------------------------------------------------ 2007/02/08 UPDATE START
'                gStr = RTrim(ZHM.材質) & Space(18 - Len(ZHM.材質)) & _
'                       Format(ZHM.板厚, "###0.000") & "x" & _
'                       Format(ZHM.幅, "###0.000") & "x" & _
'                       Format(ZHM.長さ, "###0.000")
                gStr = RTrim(Mid(ZHM.材質, 1, 18)) & Space(18 - Len(Mid(ZHM.材質, 1, 18))) & _
                       Format(ZHM.板厚, "###0.000") & "x" & _
                       Format(ZHM.幅, "###0.000") & "x" & _
                       Format(ZHM.長さ, "###0.000")
'------------------------------------------------------------------------------------ 2007/02/08 UPDATE END
                If LEVEL + 1 > 9 Then
                    Set oNod = .Add("A" & JUN, tvwChild, "A" & RTrim(JUN) & Format(gInt, "00"), gStr, 20)
                Else
                    Set oNod = .Add("A" & JUN, tvwChild, "A" & RTrim(JUN) & Format(gInt, "00"), gStr, LEVEL + 11)
                End If
                
                oNod.Expanded = True
            End With

            Call ZHMReadNext(1)
        Loop

    End If
    
    Call ZHMClose(1)

End Sub
'+---------------------------------------------------------------------------------------+
'+  構成情報グリッドより入力域へセット
'+---------------------------------------------------------------------------------------+
Private Sub GridToText_KOM()
    
    With vsGrid1
'        If Trim(.TextMatrix(.ROW, SCol(11))) <> "" Then
            
        iB1_KHinbn = RTrim(.TextMatrix(.ROW, SCol(12)))
        Call iB1_KHinbn_LostFocus
        If CisFun.Val2(.TextMatrix(.ROW, SCol(4))) <> 0 Then
            iB1_Siyou = CisFun.RSetFld(.TextMatrix(.ROW, SCol(4)), 10, "#,##0.0000")
        End If
        Select Case CisFun.Val2(.TextMatrix(.ROW, SCol(21)))
            Case 0
                B1Op_Tenkb(0).Value = True: B1Chk_Tenkb.Value = 0
            Case 1
                B1Op_Tenkb(0).Value = True: B1Chk_Tenkb.Value = 1
            Case 2
                B1Op_Tenkb(1).Value = True: B1Chk_Tenkb.Value = 1
            Case 3
                B1Op_Tenkb(1).Value = True: B1Chk_Tenkb.Value = 0
        End Select
        iB1_SDate = RTrim(.TextMatrix(.ROW, SCol(22)))
        iB1_EDate = RTrim(.TextMatrix(.ROW, SCol(23)))
        
        If CisFun.Val2(.TextMatrix(.ROW, SCol(8))) <> 0 Then
            iB1_LT = CisFun.RSetFld(.TextMatrix(.ROW, SCol(8)), 2, "##")
        End If
        
        wsv_HinKB = RTrim(.TextMatrix(.ROW, SCol(13)))
        wsv_ROW = .ROW
'        End If
        ' 手配先グリッド
        Call GridSet_TEH(.ROW)
    End With

End Sub
'+---------------------------------------------------------------------------------------+
'+  グリッド (Click)
'+---------------------------------------------------------------------------------------+
Private Sub VSGrid1_Click()
    With vsGrid1
        
        If ProcHB = "B1" And SyoriKB2 = "ADD" Then
            RowColFlg = True
            Call CisVsGrid1.vsColor(消去全部)
            Call CisVsGrid1.vsColor(個別)
            If RTrim(.TextMatrix(.ROW, SCol(12))) <> "" Then
                Call TView1_NodeClick(TView1.Nodes(.ROW))
            End If
            RowColFlg = False
            
            iB1_KHinbn.SetFocus
        End If
        
        If ProcHB = "B2" Then

            GridBack = ProcHB
            Call HeadBodyClear("B1")
             
            If SyoriKB2 <> "ADD" Then
                If Trim(.TextMatrix(.ROW, SCol(12))) <> "" Then
                   Call GridToText_KOM
                End If
                wsOHinbn = RTrim(.TextMatrix(.ROW, SCol(11)))
                wsKHinbn = RTrim(.TextMatrix(.ROW, SCol(12)))
            Else
                wsOHinbn = RTrim(.TextMatrix(.ROW, SCol(12)))
            End If
            
            RowColFlg = True
            Call CisVsGrid1.vsColor(消去全部)
            Call CisVsGrid1.vsColor(個別)
            
            Call TView1_NodeClick(TView1.Nodes(.ROW))
            
            RowColFlg = False
        End If
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  グリッド (DblClick)
'+---------------------------------------------------------------------------------------+
Private Sub vsGrid1_DblClick()
'    If MouseButton = 2 Then
        With vsGrid1
            If ProcHB = "B2" Then
'                If Trim(.TextMatrix(.ROW, SCol(11))) <> "" And _
'                   Trim(.TextMatrix(.ROW, SCol(12))) <> "" And _
'                   Trim(.TextMatrix(.ROW, SCol(13))) = "" Then
                If Trim(.TextMatrix(.ROW, SCol(11))) <> "" And _
                   Trim(.TextMatrix(.ROW, SCol(13))) = "" Then
                    If SyoriKB <> "DEL" Then
                        RowColFlg = True
                        Call CisVsGrid1.vsColor(消去全部)
                        Call CisVsGrid1.vsColor(個別)
                        RowColFlg = False
    
                        ProcHB = "B1"
                        Call DispChange(ProcHB)
                        
                        If SyoriKB2 = "ADD" Then
                            wsOHinbn = RTrim(.TextMatrix(.ROW, SCol(12)))
                            
                            B1Op_Tenkb(0).Value = True
                            B1Chk_Tenkb.Value = 0
                        Else
                            wsOHinbn = RTrim(.TextMatrix(.ROW, SCol(11)))
                            wsKHinbn = RTrim(.TextMatrix(.ROW, SCol(12)))
                        End If
                        iB1_KHinbn.SetFocus
                    End If
                End If
            End If
        End With
'    End If
End Sub
'+---------------------------------------------------------------------------------------+
''+  グリッド (GotFocus)
'+---------------------------------------------------------------------------------------+
'Private Sub vsGrid1_GotFocus()
'    If ProcHB = "B2" Then
'        iB2_Torcd.SetFocus
'    End If
'    If ProcHB = "B3" Then
'        iB3_Seban.SetFocus
'    End If
'End Sub
'+---------------------------------------------------------------------------------------+
'+  グリッド (GotFocus)
'+---------------------------------------------------------------------------------------+
Private Sub vsGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
       Call VSGrid1_Click
 '      MouseButton = 2
       Call vsGrid1_DblClick
    End If
End Sub
'+----------------------------------------+'
'+           移動選択項目選択             +
'+----------------------------------------+'
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Button : 1 = 左ﾎﾞﾀﾝ ,2 = 右ﾎﾞﾀﾝ
    With vsGrid1
        .Tag = ""
        If Button <> 1 Then Exit Sub
    ' 固定列以外は行わない
        If .MouseCol <> 0 Then Exit Sub
    ' 固定行は行わない
        If .MouseRow = 0 Then Exit Sub
    ' 移動元列をタグへ退避
        .Tag = str(.MouseRow)
        .Drag 1
    End With
End Sub
'+----------------------------------------+'
'+            表 示 項 目 移 動           +
'+----------------------------------------+'
Private Sub VSGrid1_DragDrop(Source As Control, X As Single, Y As Single)
    With vsGrid1
    ' 固定列は行わない
        If .MouseRow = 0 Then Exit Sub
        If .Tag = "" Then Exit Sub
        If CisFun.Val2(.Tag) = .MouseRow Then GoTo VsGrid1_DragDrop_ED
' ｸﾞﾘｯﾄﾞ表示 禁止
        .Redraw = False

        If RTrim(.TextMatrix(.MouseRow, SCol(13))) <> "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     材料品番の移動処理は行えません         "
                .MB_Title = "構成確認"
                .MB_Button = Error
                .MBOX
                GoTo VsGrid1_DragDrop_ED
            End With
        End If

'       < 構成チェック >
        With CisDB
            .SQL = "構成マスタメンテ入力チェック"
            .StoadoCount = 3
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = Input用
            .ParaNo = 3: .ParaIO = Input用
        
            .ParaNo = 3: .ParaValue = ""
            .ParaNo = 1: .ParaValue = RTrim(vsGrid1.TextMatrix(vsGrid1.MouseRow, SCol(12)))
            .ParaNo = 2: .ParaValue = RTrim(vsGrid1.TextMatrix(CisFun.Val2(vsGrid1.Tag), SCol(12)))
            If Not .DBStored Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "     構成入力チェック時にエラーが発生しました。    "
                    .MB_Title = "構成確認"
                    .MB_Button = Error
                    .MBOX
                    
                    GoTo VsGrid1_DragDrop_ED
                End With
            End If
        
            .ParaNo = 0
            
            If .ParaValue <> 0 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "     構成情報に矛盾が発生しました。       "
                    .MB_Title = "構成確認"
                    .MB_Button = Error
                    .MBOX
                End With
                vsGrid1.SetFocus
                GoTo VsGrid1_DragDrop_ED
            End If
        End With

'       < 移動先構成(同レベル)内の同一子品番を確認(エラー) >
        gStr1 = .TextMatrix(CisFun.Val2(.Tag), 1)
        gIdx1 = CisFun.Val2(Replace(.TextMatrix(.MouseRow, 0), ".", ""))
        
        For gInt = .MouseRow + 1 To .Rows - 1
            If gIdx1 >= CisFun.Val2(Replace(.TextMatrix(gInt, 0), ".", "")) Then Exit For
            If gIdx1 + 1 = CisFun.Val2(Replace(.TextMatrix(gInt, 0), ".", "")) Then
                gStr2 = .TextMatrix(gInt, 1)
                If gStr1 = gStr2 Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "     構成情報に矛盾が発生しました。       "
                        .MB_MSG(4) = "     同一レベルに同一構成品番の設定不可        "
                        .MB_Title = "構成確認"
                        .MB_Button = Error
                        .MBOX
                    End With
                    vsGrid1.SetFocus
                    GoTo VsGrid1_DragDrop_ED
                End If
            End If
        Next gInt
    
'   < DB 更新>
        Call DBMove(CisFun.Val2(.Tag), .MouseRow)
    
        Call GridSet_KOM(0)
VsGrid1_DragDrop_ED:
        Set oNod = Nothing
      
' ｸﾞﾘｯﾄﾞ表示 解除
    vsGrid1.Redraw = True
'
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  グリッド (RowColChange)
'+---------------------------------------------------------------------------------------+
Private Sub vsGrid1_RowColChange()
    If Not FormAct Then Exit Sub
    If RowColFlg Then Exit Sub
    
    Call VSGrid1_Click
End Sub
'+----------------------------------------------------------+
'+   手配情報   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInitT()
    With CisVsGridT
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        ReDim SColT(16)

        Call .FixedSet(中中, 右中, あり, 3, 0, "No")
        Call .FixedSet(中中, 左中, あり, 20, 1, "手配先")
        Call .FixedSet(中中, 左中, あり, 2, 2, "受")
        Call .FixedSet(中中, 左中, あり, 6, 3, "背番号")
        Call .FixedSet(中中, 左中, あり, 5, 4, "手配")
        Call .FixedSet(中中, 右中, あり, 4, 5, "比率")
        Call .FixedSet(中中, 中中, あり, 3, 6, "SK")
        Call .FixedSet(中中, 左中, あり, 8, 7, "ｻｲｸﾙ")
        Call .FixedSet(中中, 左中, あり, 21, 8, "有効期限")
        Call .FixedSet(中中, 左中, あり, 4, 9, "係1")
        Call .FixedSet(中中, 左中, あり, 4, 10, "係2")
        Call .FixedSet(中中, 左中, あり, 4, 11, "係3")
        Call .FixedSet(中中, 左中, あり, 2, 12, "LT")
        Call .FixedSet(中中, 左中, あり, 6, 13, "ロット")
        Call .FixedSet(中中, 左中, あり, 15, 14, "曜日(日付)")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, あり, 0, 15, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp

        SColT(1) = .FixedGet(1)     ' 手配先
        SColT(2) = .FixedGet(2)     ' 受入
        SColT(3) = .FixedGet(3)     ' 背番号
        SColT(4) = .FixedGet(4)     ' 手配区分
        SColT(5) = .FixedGet(5)     ' 比率
        SColT(6) = .FixedGet(6)     ' ＳＫ
        SColT(7) = .FixedGet(7)     ' サイクル
        SColT(8) = .FixedGet(8)     ' 有効期限
        SColT(9) = .FixedGet(9)     ' 係数１
        SColT(10) = .FixedGet(10)   ' 係数２
        SColT(11) = .FixedGet(11)   ' 係数３
        SColT(12) = .FixedGet(12)   ' ＬＴ
        SColT(13) = .FixedGet(13)   ' ロット
        SColT(14) = .FixedGet(14)   ' 曜日or日付
    
    End With
    KCol = 15   ' 非表示項目開始ｶﾗﾑ

End Sub
'+---------------------------------------------------------------------------------------+
'+      グリッド内容セット
'+---------------------------------------------------------------------------------------+
Private Sub GridSet_TEH(ROW As Long)
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
    Static wsKB     As String
    Call CisVsGridT.Clear
    
    With vsGrid1
        wsKB = RTrim(.TextMatrix(ROW, SCol(13)))
        If wsKB = "" Then
            gSL_Select = "SELECT HT.*,"
            gSL_Select = gSL_Select & " TR.略称 取引先名,"
            gSL_Select = gSL_Select & " SY.値名称 手配区分名,"
            gSL_Select = gSL_Select & " HN.品名"
            gSL_Select = gSL_Select & " FROM 品番取引先マスタ HT "
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN "
            gSL_Select = gSL_Select & "   ON HN.品番 = HT.品番 "
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR "
            gSL_Select = gSL_Select & "   ON TR.取引先CD = HT.取引先 "
            gSL_Select = gSL_Select & "  AND TR.取引先区分 = HT.品目 "
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY "
            gSL_Select = gSL_Select & "   ON SY.区分名称 = '手配区分' "
            gSL_Select = gSL_Select & "  AND SY.値 = HT.手配区分"
            gSL_Select = gSL_Select & "  AND SY.区分タイプ = ''"
            gSL_Select = gSL_Select & " WHERE HT.品番 = '" & RTrim(.TextMatrix(ROW, SCol(12))) & "'"
            gSL_Select = gSL_Select & " ORDER BY HT.取引先"
            Call HTMRead(gSL_Select)
        Else
            gSL_Select = "SELECT ZT.*,"
            gSL_Select = gSL_Select & " TR.略称 手配先名,"
            gSL_Select = gSL_Select & " SY.値名称 手配区分名,"
            gSL_Select = gSL_Select & " ZR.材料名"
            gSL_Select = gSL_Select & " FROM 材料マスタ ZR "
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料取引先マスタ ZT "
            gSL_Select = gSL_Select & "   ON ZR.材料管理番号 = ZT.材料管理番号 "
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR "
            gSL_Select = gSL_Select & "   ON TR.取引先CD = ZT.取引先 "
            gSL_Select = gSL_Select & "  AND TR.取引先区分 = 1 "
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY "
            gSL_Select = gSL_Select & "   ON SY.区分名称 = '手配区分' "
            gSL_Select = gSL_Select & "  AND SY.値 = ZT.手配区分"
            gSL_Select = gSL_Select & "  AND SY.区分タイプ = ''"
            gSL_Select = gSL_Select & " WHERE ZR.材料管理番号 = " & .TextMatrix(ROW, SCol(14))
            gSL_Select = gSL_Select & " ORDER BY ZT.取引先"
            Call ZTMRead(gSL_Select)
        End If
    End With
    
    If CisDB.RecordCount <> 0 Then

    ' 抽出処理 初期ｾｯﾄ
    '    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
        Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
    ' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
        With VsGridT
           .Redraw = False
            If Ttl_Kensu + 1 < CisVsGridT.DispMax Then
                .Rows = CisVsGridT.DispMax
            Else
                .Rows = Ttl_Kensu + 1
            End If
        End With
        ID = 0
        
        If wsKB = "" Then
            Do Until Not HTM_RDSTS
                ID = ID + 1
                    
                Call HTMInfoToGrid(ID, wsKB)
                Call HTMReadNext
            Loop
        Else
            Do Until Not ZTM_RDSTS
                ID = ID + 1
                    
                Call HTMInfoToGrid(ID, wsKB)
                Call ZTMReadNext
            Loop
        End If
    
    End If
    If wsKB = "" Then
        Call HTMClose
    Else
        Call ZTMClose
    End If
    With VsGridT
         .ROW = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  構成情報よりグリッドへセット
'+---------------------------------------------------------------------------------------+
Private Sub HTMInfoToGrid(SetRow As Long, KB As String)
    With VsGridT
        For mIx = 0 To .Cols - 1
            .TextMatrix(SetRow, mIx) = ""
        Next mIx
        
        ' No
        .TextMatrix(SetRow, 0) = SetRow
        If RTrim(KB) = "" Then
            If RTrim(HTM.取引先) <> "" Then
                ' 取引先
                .TextMatrix(SetRow, SColT(1)) = RTrim(HTM.取引先) & ":" & RTrim(HTM.取引先名)
                ' 受入
                .TextMatrix(SetRow, SColT(2)) = RTrim(HTM.受入)
                ' 背番号
                .TextMatrix(SetRow, SColT(3)) = RTrim(HTM.背番号)
                ' 手配区分
                .TextMatrix(SetRow, SColT(4)) = RTrim(HTM.手配区分名)
                ' 比率
                .TextMatrix(SetRow, SColT(5)) = CisFun.RSetFld(HTM.発注比率, 3, "###")
                ' SK
                If HTM.SK区分 = 0 Then
                    .TextMatrix(SetRow, SColT(6)) = "K"
                Else
                    .TextMatrix(SetRow, SColT(6)) = "S"
                End If
                ' サイクル
                If RTrim(HTM.サイクル) <> "" Then
                    .TextMatrix(SetRow, SColT(7)) = Mid(HTM.サイクル, 1, 1) & "-" & _
                                                    Mid(HTM.サイクル, 2, 2) & "-" & _
                                                    Mid(HTM.サイクル, 4, 2)
                End If
                ' 有効期限
                gStr = ""
                If Trim(HTM.有効期限開始) <> "" Then
                    gStr = Mid(HTM.有効期限開始, 1, 4) & "/" & _
                           Mid(HTM.有効期限開始, 5, 2) & "/" & _
                           Mid(HTM.有効期限開始, 7, 2)
                End If
                If Trim(HTM.有効期限終了) <> "" Then
                   If gStr = "" Then gStr = Space(10)
                   gStr = gStr & "～"
                   gStr = gStr & Mid(HTM.有効期限終了, 1, 4) & "/" & _
                                 Mid(HTM.有効期限終了, 5, 2) & "/" & _
                                 Mid(HTM.有効期限終了, 7, 2)
                Else
                   If Trim(HTM.有効期限開始) <> "" Then
                      gStr = gStr & "～" & Space(10)
                   End If
                End If
                .TextMatrix(SetRow, SColT(8)) = gStr
                ' 係数1
                If HTM.係数1 <> 0 Then
                    .TextMatrix(SetRow, SColT(9)) = CisFun.RSetFld(HTM.係数1, 4, "0.00")
                End If
                ' 係数2
                If HTM.係数2 <> 0 Then
                    .TextMatrix(SetRow, SColT(10)) = CisFun.RSetFld(HTM.係数2, 4, "0.00")
                End If
                ' 係数3
                If HTM.係数3 <> 0 Then
                    .TextMatrix(SetRow, SColT(11)) = CisFun.RSetFld(HTM.係数3, 4, "0.00")
                End If
                ' ＬＴ
                If HTM.LT <> 0 Then
                    .TextMatrix(SetRow, SColT(12)) = CisFun.RSetFld(HTM.LT, 2, "##")
                End If
                ' ロット
                If HTM.ロット1 <> 0 Then
                    .TextMatrix(SetRow, SColT(13)) = CisFun.RSetFld(HTM.ロット1, 7, "###,###")
                End If
                ' 有効期限
                gStr = ""
                If HTM.Day1 = 0 And HTM.Day2 = 0 And HTM.Day3 = 0 And HTM.Day4 = 0 And HTM.Day5 = 0 Then
                    If HTM.Week1 <> 0 Then: gStr = gStr & "月 "
                    If HTM.Week2 <> 0 Then: gStr = gStr & "火 "
                    If HTM.Week3 <> 0 Then: gStr = gStr & "水 "
                    If HTM.Week4 <> 0 Then: gStr = gStr & "木 "
                    If HTM.Week5 <> 0 Then: gStr = gStr & "金 "
                Else
                    If HTM.Day1 <> 0 Then: gStr = gStr & Format(HTM.Day1, "00") & Space(1)
                    If HTM.Day2 <> 0 Then: gStr = gStr & Format(HTM.Day2, "00") & Space(1)
                    If HTM.Day3 <> 0 Then: gStr = gStr & Format(HTM.Day3, "00") & Space(1)
                    If HTM.Day4 <> 0 Then: gStr = gStr & Format(HTM.Day4, "00") & Space(1)
                    If HTM.Day5 <> 0 Then: gStr = gStr & Format(HTM.Day5, "00") & Space(1)
                End If
                .TextMatrix(SetRow, SColT(14)) = gStr
            End If
        Else
            If RTrim(ZTM.取引先) <> "" Then
                ' 手配先
                .TextMatrix(SetRow, SColT(1)) = RTrim(ZTM.取引先) & ":" & RTrim(ZTM.手配先名)
                ' 受入
                .TextMatrix(SetRow, SColT(2)) = RTrim(ZTM.受入)
                ' 背番号
                .TextMatrix(SetRow, SColT(3)) = RTrim(ZTM.背番号)
                ' 手配区分
                .TextMatrix(SetRow, SColT(4)) = RTrim(ZTM.手配区分名)
                ' 比率
                .TextMatrix(SetRow, SColT(5)) = CisFun.RSetFld(ZTM.発注比率, 3, "###")
                ' SK
                If ZTM.SK区分 = 0 Then
                    .TextMatrix(SetRow, SColT(6)) = "K"
                Else
                    .TextMatrix(SetRow, SColT(6)) = "S"
                End If
                ' サイクル
                If RTrim(ZTM.サイクル) <> "" Then
                    .TextMatrix(SetRow, SColT(7)) = Mid(HTM.サイクル, 1, 1) & "-" & _
                                                    Mid(HTM.サイクル, 2, 2) & "-" & _
                                                    Mid(HTM.サイクル, 4, 2)
                End If
                ' 有効期限
                gStr = ""
                If Trim(ZTM.有効期限開始) <> "" Then
                    gStr = Mid(ZTM.有効期限開始, 1, 4) & "/" & _
                           Mid(ZTM.有効期限開始, 5, 2) & "/" & _
                           Mid(ZTM.有効期限開始, 7, 2)
                End If
                If Trim(ZTM.有効期限終了) <> "" Then
                   If gStr = "" Then gStr = Space(10)
                   gStr = gStr & "～"
                   gStr = gStr & Mid(ZTM.有効期限終了, 1, 4) & "/" & _
                                 Mid(ZTM.有効期限終了, 5, 2) & "/" & _
                                 Mid(ZTM.有効期限終了, 7, 2)
                Else
                   If Trim(ZTM.有効期限開始) <> "" Then
                      gStr = gStr & "～" & Space(10)
                   End If
                End If
                .TextMatrix(SetRow, SColT(8)) = gStr
                ' 係数1
                If ZTM.係数1 <> 0 Then
                    .TextMatrix(SetRow, SColT(9)) = CisFun.RSetFld(HTM.係数1, 4, "0.00")
                End If
                ' 係数2
                If ZTM.係数2 <> 0 Then
                    .TextMatrix(SetRow, SColT(10)) = CisFun.RSetFld(HTM.係数2, 4, "0.00")
                End If
                ' 係数3
                If ZTM.係数3 <> 0 Then
                    .TextMatrix(SetRow, SColT(11)) = CisFun.RSetFld(HTM.係数3, 4, "0.00")
                End If
                ' ＬＴ
                If ZTM.LT <> 0 Then
                    .TextMatrix(SetRow, SColT(12)) = CisFun.RSetFld(HTM.LT, 2, "##")
                End If
                ' ロット
                If ZTM.ロット1 <> 0 Then
                    .TextMatrix(SetRow, SColT(13)) = CisFun.RSetFld(HTM.ロット1, 7, "###,###")
                End If
                ' 有効期限
                gStr = ""
                If ZTM.Day1 = 0 And ZTM.Day2 = 0 And ZTM.Day3 = 0 And ZTM.Day4 = 0 And ZTM.Day5 = 0 Then
                    If ZTM.Week1 <> 0 Then: gStr = gStr & "月 "
                    If ZTM.Week2 <> 0 Then: gStr = gStr & "火 "
                    If ZTM.Week3 <> 0 Then: gStr = gStr & "水 "
                    If ZTM.Week4 <> 0 Then: gStr = gStr & "木 "
                    If ZTM.Week5 <> 0 Then: gStr = gStr & "金 "
                Else
                    If ZTM.Day1 <> 0 Then: gStr = gStr & Format(ZTM.Day1, "00") & Space(1)
                    If ZTM.Day2 <> 0 Then: gStr = gStr & Format(ZTM.Day2, "00") & Space(1)
                    If ZTM.Day3 <> 0 Then: gStr = gStr & Format(ZTM.Day3, "00") & Space(1)
                    If ZTM.Day4 <> 0 Then: gStr = gStr & Format(ZTM.Day4, "00") & Space(1)
                    If ZTM.Day5 <> 0 Then: gStr = gStr & Format(ZTM.Day5, "00") & Space(1)
                End If
                .TextMatrix(SetRow, SColT(14)) = gStr
            End If
        End If
    End With
End Sub

'☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆
'☆
'☆                     ツリー ビュー (TreeView) コントロール　処理
'☆
'☆  ( Event 発生順 )
'☆
'☆         TView1_MouseDown → TView1_MouseMove → TView1_DragOver → TView1_DragDrop
'☆
'☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆

'*----------------------------------*
'*      表示・非表示                 *
'*----------------------------------*
Private Sub TView1_Collapse(ByVal Node As MSComctlLib.Node)
    Call DispChg_TView1(Node)
End Sub
Private Sub TView1_Expand(ByVal Node As MSComctlLib.Node)
    Call DispChg_TView1(Node)
End Sub
Private Sub DispChg_TView1(ByVal Node As MSComctlLib.Node)
    Call CisFrm.MousePT(11)
    
    With vsGrid1
        For gInt = 1 To TView1.Nodes.Count
'            TView1.SelectedItem.Index = gInt
            If Not TView1.Nodes(gInt) Is Nothing Then
'                TView1.DropHighlight = TView1.Nodes(gInt)
'                If TView1.DropHighlight.Visible Then
                If TView1.Nodes(gInt).Visible Then
                    .RowHeight(gInt) = 285
                Else
                    .RowHeight(gInt) = 0
                End If
            End If
        Next gInt
        
        Dummy.Enabled = True
        Dummy.SetFocus
        Call TView1_NodeClick(Node)
        Dummy.Enabled = False
        DoEvents
        
    End With

    Call CisFrm.MousePT(1)
End Sub
'*----------------------------------*
'*      選択                         *
'*----------------------------------*
Private Sub TView1_DblClick()
    If MouseButton = 2 Then
        Call vsGrid1_DblClick
    End If
End Sub
'*----------------------------------*
'*      カーソル制御                 *
'*----------------------------------*
Private Sub TView1_NodeClick(ByVal Node As MSComctlLib.Node)
    
    If Not indrag Then

        TView1.DropHighlight = Node
        
        If Not TView1.DropHighlight.Visible Then
            TView1.DropHighlight.Selected = True
        End If
        If ActiveControl.Name = "TView1" And ProcHB <> "H1" Then
            RowColFlg = True
            vsGrid1.ROW = TView1.DropHighlight.Index
            Call CisVsGrid1.vsColor(消去全部)
            Call CisVsGrid1.vsColor(個別)
            If Not vsGrid1.RowIsVisible(TView1.DropHighlight.Index) Then: vsGrid1.TopRow = TView1.DropHighlight.Index
            DoEvents
            RowColFlg = False
        End If
    End If
End Sub
'*----------------------------------*
'* ①   行 選 択 ( Mouse Down )     *
'*----------------------------------*
Private Sub TView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseButton = Button
    
    If TView1.HitTest(X, Y) Is Nothing Then Exit Sub

    If wFLG Then Exit Sub
    wFLG = True
    
    TView1.HitTest(X, Y).Selected = True                ' Focus
    Set TView1.DropHighlight = TView1.HitTest(X, Y)     ' 反転表示
    
'   < 親品番・基材品番の移動は不可 >
    If RTrim(TView1.HitTest(X, Y).Key) = "A" Then
        Set oNod = Nothing
    Else
        Set oNod = TView1.HitTest(X, Y)                 ' ドラッグされる項目を代入します。
    End If
    
    If ActiveControl.Name = "TView1" And ProcHB <> "H1" Then
        vsGrid1.ROW = TView1.SelectedItem.Index
        RowColFlg = True
        Call CisVsGrid1.vsColor(消去全部)
        Call CisVsGrid1.vsColor(個別)
        RowColFlg = False
        DoEvents
    End If
    
    wFLG = False
    DoEvents
End Sub
'*----------------------------------*
'* ②   行 移 動 ( Mouse Move )     *
'*----------------------------------*
Private Sub TView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then                               ' ドラッグ操作の場合、
        If ProcHB = "H1" Then Exit Sub
        If oNod Is Nothing Then Exit Sub
        If TView1.HitTest(X, Y) Is Nothing Then Exit Sub
        If oNod = TView1.HitTest(X, Y).Root Then Exit Sub
        
        
        wTIdx = 0
        indrag = True                                           ' フラグを真に設定します。
        TView1.DragIcon = TView1.SelectedItem.CreateDragImage   ' CreateDragImage メソッドを使用してドラッグ中のアイコンを設定します。
        TView1.Drag vbBeginDrag                                 ' ドラッグ操作を開始します。
        
    End If
End Sub
'*-----------------------------------------*
'* ③   行 移 動 開 始 ( Drag Over )        *
'*-----------------------------------------*
Private Sub TView1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If indrag = True And wTIdx = 0 Then
        If Not TView1.HitTest(X, Y) Is Nothing Then
            Set TView1.DropHighlight = TView1.HitTest(X, Y)     ' DropHighlight プロパティにマウスの座標位置にあったオブジェクトを設定します。
            wTIdx = TView1.HitTest(X, Y).Index
        End If
    End If
End Sub
'*-----------------------------------------*
'* ④   行 移 動 終 了 ( Drag Drop )        *
'*-----------------------------------------*
Private Sub TView1_DragDrop(Source As Control, X As Single, Y As Single)
    Static sw_I         As Integer
    Static sw_UpdKB     As Byte
    Static swsv_Key     As Integer
    Static swsv_Hinbn   As Integer

    sw_UpdKB = 0
    If Not (TView1.HitTest(X, Y) Is Nothing) Then
        If oNod = TView1.HitTest(X, Y) Then GoTo TView1_DragDrop_ED
        Cls

'       < 移動元親インデックス退避 >
        sw_I = wTIdx

    ' 材料
        If RTrim(vsGrid1.TextMatrix(TView1.HitTest(X, Y).Index, SCol(13))) <> "" Or _
           RTrim(vsGrid1.TextMatrix(sw_I, SCol(13))) <> "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     材料品番の移動処理は行えません         "
                .MB_Title = "構成確認"
                .MB_Button = Error
                .MBOX
                GoTo TView1_DragDrop_ED
            End With
        End If

'       < 構成チェック >
        With CisDB
            .SQL = "構成マスタメンテ入力チェック"
            .StoadoCount = 3
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = Input用
            .ParaNo = 3: .ParaIO = Input用
        
            .ParaNo = 3: .ParaValue = ""
            .ParaNo = 1: .ParaValue = RTrim(vsGrid1.TextMatrix(TView1.HitTest(X, Y).Index, SCol(12)))
            .ParaNo = 2: .ParaValue = RTrim(vsGrid1.TextMatrix(sw_I, SCol(12)))
            If Not .DBStored Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "     構成入力チェック時にエラーが発生しました。    "
                    .MB_Title = "構成確認"
                    .MB_Button = Error
                    .MBOX
                    
                    GoTo TView1_DragDrop_ED
                End With
            End If
        
            .ParaNo = 0
            
            If .ParaValue <> 0 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "     構成情報に矛盾が発生しました。       "
                    .MB_Title = "構成確認"
                    .MB_Button = Error
                    .MBOX
                End With
                TView1.SetFocus
                GoTo TView1_DragDrop_ED
            End If
        End With

'       < 移動先構成(同レベル)内の同一子品番を確認(エラー) >
        gIdx1 = TView1.HitTest(X, Y).Children
        If gIdx1 <> 0 Then
            gIdx2 = TView1.HitTest(X, Y).Child.FirstSibling.Index
            gIdx3 = gIdx2

            While gIdx2 <> TView1.HitTest(X, Y).Child.LastSibling.Index

                If gIdx2 = gIdx3 Then
                    gStr = TView1.HitTest(X, Y).Child.FirstSibling.Text
                    gIdx3 = -99
                Else
                    gStr = TView1.Nodes(gIdx2).Next.Text
                    gIdx2 = TView1.Nodes(gIdx2).Next.Index
                End If

                If oNod = gStr Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "     構成情報に矛盾が発生しました。       "
                        .MB_MSG(4) = "     同一レベルに同一構成品番の設定不可        "
                        .MB_Title = "構成確認"
                        .MB_Button = Error
                        .MBOX
                    End With
                    TView1.SetFocus
                    GoTo TView1_DragDrop_ED
                End If
            Wend
        End If

        swsv_Key = sw_I
        swsv_Hinbn = TView1.HitTest(X, Y).Index

'   < DB 更新>
        Call DBMove(swsv_Key, swsv_Hinbn)
        sw_UpdKB = 1
    End If

    Call GridSet_KOM(0)
TView1_DragDrop_ED:
    Set oNod = Nothing
    Set TView1.DropHighlight = Nothing
    
    indrag = False
End Sub
'*------------------------------------------------*
'*      ツリービュー移動に伴うデータ更新処理        *
'*------------------------------------------------*
Private Function DBMove(pKey As Integer, pHinbn As Integer)
    DBMove = False

    On Error GoTo DBMove_Err

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    With KOM
'       < 移動元更新 >

        .親品番 = vsGrid1.TextMatrix(pHinbn, SCol(12))
        .変更前親品番 = vsGrid1.TextMatrix(pKey, SCol(11))
        .変更前子品番 = vsGrid1.TextMatrix(pKey, SCol(12))
        .更新者 = gTanto
        Call KOMUpdate2
    End With
'
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    On Error GoTo 0
    
    DBMove = True
    Exit Function
DBMove_Err:
    CisFun.ErrorBox
    End
End Function


