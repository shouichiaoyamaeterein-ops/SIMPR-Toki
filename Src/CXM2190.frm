VERSION 5.00
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXM2190 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "期首受注単価　一括メンテナンス"
   ClientHeight    =   11235
   ClientLeft      =   270
   ClientTop       =   3480
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   Begin Cis3D_v60.CIS3D CIS3D10 
      Height          =   285
      Left            =   13680
      Top             =   3690
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   503
      BackColor       =   14737632
      Caption         =   "有無"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
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
      cAlingnment     =   7
      cBoderStyle     =   2
   End
   Begin Cis3D_v60.CIS3D Back_ZHinbn 
      Height          =   255
      Left            =   3195
      Top             =   3690
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   450
      BackColor       =   14737632
      Caption         =   "材質/寸法"
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
      cAlingnment     =   7
      cBoderStyle     =   2
   End
   Begin VB.PictureBox H0_Area1 
      Height          =   825
      Left            =   300
      ScaleHeight     =   765
      ScaleWidth      =   2235
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   600
      Width           =   2295
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   660
         Left            =   60
         Top             =   60
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1164
         ForeColor       =   16711680
         Caption         =   "単価区分"
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
         Begin VB.OptionButton H0Op_Tanku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "部品"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton H0Op_Tanku 
            BackColor       =   &H00C0FFC0&
            Caption         =   "材料"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1050
            TabIndex        =   0
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
   End
   Begin Cis3D_v60.CIS3D Message 
      Height          =   870
      Left            =   4065
      Top             =   5145
      Visible         =   0   'False
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   1535
      BackColor       =   16744576
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   720
         Left            =   75
         Top             =   75
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   1270
         BackColor       =   16777152
         Caption         =   "【 処 理 中 】 しばらくお待ち下さい"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HG創英角ﾎﾟｯﾌﾟ体"
            Size            =   18
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
         cAlingnment     =   7
         cBoderStyle     =   1
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   345
      Index           =   1
      Left            =   7695
      ScaleHeight     =   345
      ScaleWidth      =   3900
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   30
      Width           =   3900
      Begin Cis3D_v60.CIS3D UC_3D6 
         Height          =   345
         Left            =   0
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         BackColor       =   12648384
         ForeColor       =   0
         Caption         =   "表示件数"
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
         Begin Cis3D_v60.CIS3D B1lb_Kensu 
            Height          =   285
            Left            =   990
            Top             =   30
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            BackColor       =   12648384
            ForeColor       =   0
            Caption         =   "##,###"
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
         Left            =   1935
         Top             =   0
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   609
         BackColor       =   16777152
         ForeColor       =   0
         Caption         =   "変更件数"
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
            Height          =   285
            Left            =   990
            Top             =   45
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            BackColor       =   16777152
            ForeColor       =   0
            Caption         =   "##,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
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
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   825
      Left            =   11400
      ScaleHeight     =   765
      ScaleWidth      =   3510
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   600
      Width           =   3570
      Begin CisBtn_60.CisBtn PB_LOOK 
         Height          =   585
         Left            =   60
         Top             =   90
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
         Left            =   915
         Top             =   90
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
         Left            =   1770
         Top             =   90
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
         Left            =   2625
         Top             =   90
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
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   7620
      Left            =   300
      ScaleHeight     =   7560
      ScaleWidth      =   14640
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3480
      Width           =   14700
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   7395
         Left            =   75
         Top             =   90
         Width           =   14460
         _ExtentX        =   25506
         _ExtentY        =   13044
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin vsFlexLib.vsFlexArray VsGrid 
            Height          =   7305
            Left            =   45
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   30
            Width           =   14340
            _Version        =   196608
            _ExtentX        =   25294
            _ExtentY        =   12885
            _StockProps     =   228
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "№ |支給先            |品番                               |前回開始  |前回終了  |仮|前回単価     |単価         |有 |仮 |"
            Rows            =   23
            Cols            =   28
            BackColor       =   12648384
            BackColorFixed  =   14737632
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   2
            ScrollBars      =   2
            RowHeightMin    =   330
            GridLines       =   2
            SelectionMode   =   1
            Begin Cis3D_v60.CIS3D Input_Area 
               Height          =   345
               Left            =   11550
               Top             =   960
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   609
               BackColor       =   4210752
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               Begin VB.TextBox F_TabDummy 
                  Height          =   330
                  Left            =   2505
                  TabIndex        =   18
                  Text            =   "Text1"
                  Top             =   0
                  Width           =   240
               End
               Begin VB.TextBox B_TabDummy 
                  Height          =   330
                  Left            =   -240
                  TabIndex        =   19
                  Text            =   "Text1"
                  Top             =   0
                  Width           =   240
               End
               Begin Cis3D_v60.CIS3D B1Btn_Umu 
                  Height          =   345
                  Left            =   1650
                  Top             =   0
                  Width           =   450
                  _ExtentX        =   794
                  _ExtentY        =   609
                  BackColor       =   14737632
                  Caption         =   "有"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   5
                  cAlingnment     =   7
               End
               Begin CisText_V60.CisText iB1_Tan 
                  Height          =   375
                  Left            =   0
                  TabIndex        =   21
                  Top             =   0
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   661
                  cFormat         =   "##,###,##0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##,###,##0.00"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######0.00"
                  cILength        =   8
               End
               Begin Cis3D_v60.CIS3D B1Btn_Kari 
                  Height          =   345
                  Left            =   2085
                  Top             =   0
                  Width           =   420
                  _ExtentX        =   741
                  _ExtentY        =   609
                  BackColor       =   14737632
                  Caption         =   "正"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   5
                  cAlingnment     =   7
               End
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   5985
      Top             =   15
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      IMEMode         =   3  'ｵﾌ固定
      Left            =   0
      TabIndex        =   22
      Text            =   "Dummy"
      Top             =   0
      Width           =   210
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
      Caption         =   "【 支給単価マスタ　一括入力  】"
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
         Left            =   11640
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
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   16711935
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
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   1875
      Left            =   300
      ScaleHeight     =   1815
      ScaleWidth      =   14640
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1500
      Width           =   14700
      Begin Cis3D_v60.CIS3D CIS3D16 
         Height          =   1710
         Left            =   2745
         Top             =   60
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   3016
         ForeColor       =   16711680
         Caption         =   "支 給 先"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
            Left            =   150
            TabIndex        =   4
            Top             =   525
            Width           =   990
            _ExtentX        =   1746
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
            Left            =   150
            Top             =   885
            Width           =   2505
            _ExtentX        =   4419
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
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   1710
         Left            =   105
         Top             =   60
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   3016
         BackColor       =   16761087
         ForeColor       =   16711680
         Caption         =   "対象期間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   330
            Left            =   150
            Top             =   570
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
            Caption         =   "開始"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cPositionX      =   30
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_EYmd 
            Height          =   330
            Left            =   750
            TabIndex        =   3
            Top             =   900
            Width           =   1695
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_SYmd 
            Height          =   330
            Left            =   750
            TabIndex        =   2
            Top             =   570
            Width           =   1695
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
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   330
            Left            =   150
            Top             =   900
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
            Caption         =   "終了"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1710
         Left            =   11955
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   3016
         ForeColor       =   16711680
         Caption         =   "品番表示"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Begin VB.OptionButton OpH1_HinKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "品番マスタ"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   15
            Top             =   750
            Width           =   1140
         End
         Begin VB.OptionButton OpH1_HinKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "単価マスタ"
            Height          =   240
            Index           =   0
            Left            =   60
            TabIndex        =   14
            Top             =   465
            Value           =   -1  'True
            Width           =   1140
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   1710
         Left            =   13260
         Top             =   60
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   3016
         ForeColor       =   16711680
         Caption         =   "単価表示"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Begin VB.OptionButton OpH1_TanKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "無"
            Height          =   240
            Index           =   0
            Left            =   60
            TabIndex        =   16
            Top             =   465
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OpH1_TanKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "前回単価"
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   17
            Top             =   750
            Width           =   1035
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hinbn 
         Height          =   1710
         Left            =   5565
         Top             =   60
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   3016
         ForeColor       =   16711680
         Caption         =   " 品　番           "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cPositionY      =   60
         Begin Cis3D_v60.CIS3D H1_ZaiBox 
            Height          =   375
            Left            =   90
            Top             =   660
            Visible         =   0   'False
            Width           =   6150
            _ExtentX        =   10848
            _ExtentY        =   661
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
            cBoderColor2    =   16777215
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cFont3D         =   2
            cBoderStyle     =   1
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   345
               Index           =   27
               Left            =   5130
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   609
               ForeColor       =   16711680
               Caption         =   "長"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
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
            End
            Begin Cis3D_v60.CIS3D DH1_Long 
               Height          =   90
               Left            =   5310
               Top             =   15
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   159
               BackColor       =   12640511
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
               cBoderColor1    =   12632256
               cBoderColor2    =   12640511
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cBoderStyle     =   2
               cPositionX      =   90
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   345
               Index           =   28
               Left            =   4140
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   609
               ForeColor       =   16711680
               Caption         =   "幅"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
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
            End
            Begin Cis3D_v60.CIS3D DH1_Width 
               Height          =   90
               Left            =   4320
               Top             =   15
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   159
               BackColor       =   12640511
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
               cBoderColor1    =   12632256
               cBoderColor2    =   12640511
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cBoderStyle     =   2
               cPositionX      =   90
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   345
               Index           =   29
               Left            =   3015
               Top             =   15
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   609
               ForeColor       =   16711680
               Caption         =   "板厚"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   3
               cPositionY      =   60
            End
            Begin Cis3D_v60.CIS3D DH1_Itatu 
               Height          =   90
               Left            =   3330
               Top             =   15
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   159
               BackColor       =   12640511
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
               cBoderColor1    =   12632256
               cBoderColor2    =   12640511
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cBoderStyle     =   2
               cPositionX      =   90
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   345
               Index           =   26
               Left            =   15
               Top             =   15
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   609
               ForeColor       =   16711680
               Caption         =   "材質"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   3
               cPositionY      =   60
            End
            Begin Cis3D_v60.CIS3D DH1_Zaist 
               Height          =   90
               Left            =   330
               Top             =   15
               Width           =   2820
               _ExtentX        =   4974
               _ExtentY        =   159
               BackColor       =   12640511
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
               cBoderColor1    =   12632256
               cBoderColor2    =   12640511
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   90
            End
            Begin CisText_V60.CisText iH1_Long 
               Height          =   330
               Left            =   5340
               TabIndex        =   13
               Top             =   60
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               cFormat         =   "###0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###0.00"
               MaxLength       =   7
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "###0.00"
               cILength        =   4
            End
            Begin CisText_V60.CisText iH1_Width 
               Height          =   300
               Left            =   4350
               TabIndex        =   12
               Top             =   60
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   529
               cFormat         =   "###0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###0.00"
               MaxLength       =   7
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "###0.00"
               cILength        =   4
            End
            Begin CisText_V60.CisText iH1_Itatu 
               Height          =   300
               Left            =   3360
               TabIndex        =   11
               Top             =   60
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   529
               cFormat         =   "###0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###0.00"
               MaxLength       =   7
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "###0.00"
               cILength        =   4
            End
            Begin CisText_V60.CisText iH1_Zaist 
               Height          =   300
               Left            =   360
               TabIndex        =   10
               Top             =   60
               Width           =   2730
               _ExtentX        =   4815
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX1XXXXXXXXX2XXXX5"
               MaxLength       =   25
            End
         End
         Begin VB.OptionButton H1Op_Hinku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "自社品番"
            Height          =   240
            Index           =   2
            Left            =   3390
            TabIndex        =   7
            Top             =   120
            Width           =   1140
         End
         Begin VB.OptionButton H1Op_Hinku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "材料/寸法"
            Height          =   240
            Index           =   3
            Left            =   4590
            TabIndex        =   8
            Top             =   120
            Width           =   1140
         End
         Begin VB.OptionButton H1Op_Hinku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "発注品番"
            Height          =   240
            Index           =   0
            Left            =   960
            TabIndex        =   5
            Top             =   120
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton H1Op_Hinku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "受給品番"
            Height          =   240
            Index           =   1
            Left            =   2160
            TabIndex        =   6
            Top             =   120
            Width           =   1035
         End
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   90
            TabIndex        =   9
            Top             =   525
            Width           =   6255
            _ExtentX        =   11033
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
         Begin Cis3D_v60.CIS3D H1lb_HinBn 
            Height          =   225
            Index           =   3
            Left            =   840
            Top             =   1380
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   397
            Caption         =   "SPH440-0D2.9Ｘ238Ｘ1219 5XXXX3XXXXXXXXX4XXXXXXXXX5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
            cPositionX      =   60
         End
         Begin Cis3D_v60.CIS3D H1lb_HinBn 
            Height          =   225
            Index           =   1
            Left            =   840
            Top             =   960
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   397
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXX5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
            cPositionX      =   60
         End
         Begin Cis3D_v60.CIS3D b_HinNM 
            Height          =   225
            Index           =   1
            Left            =   120
            Top             =   960
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            Caption         =   "受給品番"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D b_HinNM 
            Height          =   225
            Index           =   3
            Left            =   120
            Top             =   1380
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            Caption         =   "材　　　料"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D H1lb_HinBn 
            Height          =   225
            Index           =   2
            Left            =   840
            Top             =   1170
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   397
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXX5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
            cPositionX      =   60
         End
         Begin Cis3D_v60.CIS3D b_HinNM 
            Height          =   225
            Index           =   2
            Left            =   120
            Top             =   1170
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            Caption         =   "自社品番"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D H1lb_HinBn 
            Height          =   225
            Index           =   0
            Left            =   840
            Top             =   840
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   397
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXX5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
            cPositionX      =   60
         End
         Begin Cis3D_v60.CIS3D b_HinNM 
            Height          =   225
            Index           =   0
            Left            =   120
            Top             =   840
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            Caption         =   "発注品番"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
         End
      End
   End
   Begin Cis3D_v60.CIS3D H1_ZaiDsp 
      Height          =   300
      Left            =   2820
      Top             =   1110
      Visible         =   0   'False
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   529
      BackColor       =   16744576
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
      cBoderColor1    =   16744576
      cBoderColor2    =   16744576
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cFont3D         =   2
      cBoderStyle     =   2
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   330
         Index           =   30
         Left            =   0
         Top             =   -30
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
         BackColor       =   4210752
         ForeColor       =   16777215
         Caption         =   "材料/寸法"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   4210752
         cBoderColor2    =   4210752
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cFont3D         =   2
         cAlingnment     =   6
         cBoderStyle     =   2
         cPositionY      =   60
      End
      Begin Cis3D_v60.CIS3D H1_ZaiRyo 
         Height          =   300
         Left            =   960
         Top             =   0
         Width           =   5910
         _ExtentX        =   10425
         _ExtentY        =   529
         BackColor       =   16744576
         ForeColor       =   16777215
         Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXX5/ZZZZ.ZZx0000.00xZZZZ.ZZ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16744576
         cBoderColor2    =   16744576
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   30
      End
      Begin Cis3D_v60.CIS3D H1lb_ZaiKan 
         Height          =   300
         Left            =   7050
         Top             =   0
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "9999999999"
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cFont3D         =   2
         cAlingnment     =   1
         cBoderStyle     =   2
      End
   End
End
Attribute VB_Name = "CXM2190"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   支給単価マスタ　一括入力
'**       フォームID    :   CXM2190
'**       処理概要      :
'**
'**       作  成  日    :   2009/10/05  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKB                 As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM                 As String       ' 処理名称(新規･修正･削除 等)
    Dim CisVsGrid               As New CisVsGrid3
    
    Dim ActObj                  As Object
    Dim ErrObj                  As Object
    Dim ActRow                  As Integer
    Dim SaveSYmd                As String
    Dim SaveEYmd                As String
    Dim SaveRow                 As Integer
    
    Dim KamiKI                  As Integer
    Dim SimoKI                  As Integer
    Dim wEtcTan                 As Currency
    
    Dim mGotColor               As Long
    Dim mLostColor              As Long

    Dim Hinku_top(3)            As Integer      ' 品番表示位置      ' 2009.10.05 add
    Dim gMaeIndex               As Integer      ' １つ前の選択品番  ' 2009.10.07 add
    Dim cnt                     As Integer      ' ワークカウンタ    ' 2009.10.05 add
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   # 二 重 起 動 防 止  #
'   #-------------------#
    If Not Dupli_Start Then End
'   #-------------------#
'   # 初 期 内 容 取 得  #
'   #-------------------#
    If Not IniGet Then End
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    
    mGotColor = H0Op_Tanku(1).BackColor
    mLostColor = H0Op_Tanku(0).BackColor
    Hinku_top(0) = b_HinNM(1).Top       ' 2009.10.05 add
    Hinku_top(1) = b_HinNM(1).Top       ' 2009.10.05 add
    Hinku_top(2) = b_HinNM(2).Top       ' 2009.10.05 add
    Hinku_top(3) = b_HinNM(3).Top       ' 2009.10.05 add
    
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = VsGrid
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 18
        .InitGet = False
        .Init
    End With
    Call GridInit
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Call CisFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
'-----------------------------------------------' 2009.10.07 add start
    Hinku_top(0) = b_HinNM(1).Top
    Hinku_top(1) = b_HinNM(1).Top
    Hinku_top(2) = b_HinNM(2).Top
    Hinku_top(3) = b_HinNM(3).Top
    H0Op_Tanku(0).Value = True
    Call H0Op_Tanku_Click(1)
    H1_ZaiRyo.Caption = ""
    H1_ZaiDsp.Visible = False
'-----------------------------------------------' 2009.10.07 add end
'-------------------------------------------------------------------------------' 2009.10.07 試験用 start
'H1lb_HinBn(0).Caption = "HC-HINBAN-発注品番"
'H1lb_HinBn(1).Caption = "JK-HINBAN-受給品番"
'H1lb_HinBn(2).Caption = "JS-HINBAN-自社品番"
'H1lb_HinBn(3).Caption = "SPH440-0D2.9Ｘ238Ｘ1219 5XXXX3XXXXXXXXX4XXXXXXXXX5"
'-------------------------------------------------------------------------------' 2009.10.07 試験用 end

'   #------------------#
'   #    初期入力設定  #
'   #------------------#
    ProcHB = "H0"
    Call DispChange(ProcHB)
    
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = Not FormAct
    
    H0Op_Tanku(1).Value = True
   
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
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
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF11: MsgBox (ActiveControl.Name)
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      '【入力】
                                    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                    Dummy.Enabled = False
                                End If
           Case Else:           Exit Function
    End Select
    
    Key_Acc = True
End Function
'---------------------------------------------------------------' 2009.10.06 add start
'+---------------------+
'+    材質入力エリア   +
'+---------------------+
Private Sub iH1_Zaist_GotFocus()
    DH1_Zaist.BackColor = &HC0FFFF
    DH1_Zaist.cBoderColor2 = &HC0FFFF
    PB_Look.Tag = "iH1_Hinbn"   ' ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Zaist_LostFocus()
    Call HinbnToLabel
    DH1_Zaist.BackColor = &HC0E0FF
    DH1_Zaist.cBoderColor2 = &HC0E0FF
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
Private Sub iH1_Itatu_GotFocus()
    DH1_Itatu.BackColor = &HC0FFFF
    DH1_Itatu.cBoderColor2 = &HC0FFFF
    PB_Look.Tag = "iH1_Hinbn"   ' ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Itatu_LostFocus()
    Call HinbnToLabel
    DH1_Itatu.BackColor = &HC0E0FF
    DH1_Itatu.cBoderColor2 = &HC0E0FF
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
Private Sub iH1_Width_GotFocus()
    DH1_Width.BackColor = &HC0FFFF
    DH1_Width.cBoderColor2 = &HC0FFFF
    PB_Look.Tag = "iH1_Hinbn"   ' ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Width_LostFocus()
    Call HinbnToLabel
    DH1_Width.BackColor = &HC0E0FF
    DH1_Width.cBoderColor2 = &HC0E0FF
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
Private Sub iH1_Long_GotFocus()
    DH1_Long.BackColor = &HC0FFFF
    DH1_Long.cBoderColor2 = &HC0FFFF
    PB_Look.Tag = "iH1_Hinbn"   ' ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Long_LostFocus()
    Call HinbnToLabel
    DH1_Long.BackColor = &HC0E0FF
    DH1_Long.cBoderColor2 = &HC0E0FF
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
Private Sub H1_ZaiBox_Set()
    Dim Atai As String
    iH1_Zaist.Text = RTrim(Mid(H1_ZaiRyo.Caption, 1, 25))
    Atai = RTrim(Mid(H1_ZaiRyo.Caption, 27, 7))
    If Atai <> "" Then iH1_Itatu.Text = Format(CisFun.Val2(Atai), "#,##0.00")
    Atai = RTrim(Mid(H1_ZaiRyo.Caption, 35, 7))
    If Atai <> "" Then iH1_Width.Text = Format(CisFun.Val2(Atai), "#,##0.00")
    Atai = RTrim(Mid(H1_ZaiRyo.Caption, 43, 7))
    If Atai <> "" Then iH1_Long.Text = Format(CisFun.Val2(Atai), "#,##0.00")
End Sub
'---------------------------------------------------------------' 2009.10.06 add end

'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'    Call CisDB.DBDISConnect                            ' ﾃﾞｰﾀﾍﾞｰｽ切断
    Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
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
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   
    If ProcHB = "H1" Then
       ProcHB = "H0"
       Call DispChange(ProcHB)
       Call HeadBodyClear("H1")
'-----------------------------------------------' 2009.10.07 add start
      Call HeadBodyClear("B")
      H1_ZaiBox.Visible = False
      H1_ZaiRyo.Caption = ""
      iH1_Hinbn.Text = ""
      iH1_Hinbn.Visible = True
      If H0Op_Tanku(0).Value Then
         H0Op_Tanku(0).SetFocus
      Else
         H0Op_Tanku(1).SetFocus
      End If
'-----------------------------------------------' 2009.10.07 add end
       GoTo PB_CAN_Ed
    End If
    
    If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      iH1_Torcd.SetFocus
      iB1_Tan.Tag = ""
      iT1_Mntsu = ""
      For cnt = 0 To 3: H1lb_HinBn(cnt) = "": Next              ' 2009.10.08 add
      iH1_Hinbn.Text = "": iH1_Zaist.Text = ""                  ' 2009.10.08 add
      iH1_Itatu.Text = "": iH1_Width = "": iH1_Long.Text = ""   ' 2009.10.08 add
      GoTo PB_CAN_Ed
    End If
    
PB_CAN_Ed:
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
'-----------------------------------------------------------------------' 2009.10.28 add start
    If H0Op_Tanku(0).Value Then
       If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
       If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    Else
       If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Zairyo: Exit Sub
       If PB_Look.Tag = "iH1_Torcd" Then Call Look_ZaiTor: Exit Sub
    End If
'-----------------------------------------------------------------------' 2009.10.28 upd end
   'If PB_LOOK.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub         ' 2009.10.28 del
    If PB_Look.Tag = "iH1_SYmd" Then iH1_SYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_EYmd" Then iH1_EYmd.ShowCalender: Exit Sub
End Sub
'+----------------------------+
'+       支給先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 8            ' 2009.10.28 upd : 3 --> 8
    RV_Call = ""
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'-----------------------------------------------------------------------' 2009.10.28 add start
'+----------------------------+
'+       検　　　　 索         +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RVI_Hinmoku = 1
    RVI_SK = 9
    
    CKK0030.Show vbModal
    Unload CKK0030
    Set CKK0030 = Nothing

    If RV_Rtn Then
       iH1_Hinbn = RV_Hinbn
'       iH1_Torcd = RV_TorcdK
'       Call iH1_Torcd_LostFocus
    End If

End Sub
'+------------------+
'+     材料検索     +
'+------------------+
Private Sub Look_Zairyo()
    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RV_Hinmoku = 1
    RV_Zaist = ""

    CKK0070.Show vbModal
    Unload CKK0070
    Set CKK0070 = Nothing

    If RV_Rtn Then
        H1lb_HinBn(0) = RV_HHinbn
        H1lb_HinBn(1) = RV_JHinbn
        H1lb_HinBn(3) = RV_Zaist & " " & _
                      Format(RV_Itatu, "#,##0.00") & "X" & _
                      Format(RV_Width, "#,##0.00") & "X" & _
                      Format(RV_Long, "#,##0.00")
        H1_ZaiRyo.Caption = RTrim(RV_Zaist) & Space(25 - Len(RTrim(RV_Zaist))) & "/" & _
                        Format(RV_Itatu, "0000.00") & "x" & _
                        Format(RV_Width, "0000.00") & "x" & _
                        Format(RV_Long, "0000.00")
                        Call H1_ZaiBox_Set
        H1lb_ZaiKan = Rv_KanriNo
        
        gSL_Select = "SELECT * FROM 材料取引先マスタ WHERE 材料管理番号 = " & Rv_KanriNo
        Call ZTMRead(gSL_Select)
        H1lb_HinBn(2) = ZTM.自社品番
    
        For cnt = 0 To 3
            If H1Op_Hinku(cnt).Value = True Then
               iH1_Hinbn = H1lb_HinBn(cnt)
               cnt = 4
            End If
        Next
    End If
End Sub

'+------------------------+
'+     材料取引先検索     +
'+------------------------+
Private Sub Look_ZaiTor()
    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 1
    RVI_SK = 9
    RVI_Tehai = 4

    CKK0077.Show vbModal
    Unload CKK0077
    Set CKK0077 = Nothing

    If RV_Rtn Then
       iH1_Torcd = RV_TorcdK
'      iH1_Torcd = RV_Torcd
       H1lb_Tornm = RV_TorRName
       
       H1lb_HinBn(0) = RV_HHinbn
       H1lb_HinBn(1) = RV_JHinbn
       H1lb_HinBn(3) = RV_Zaist & " " & _
                      Format(RV_Itatu, "#,##0.00") & "X" & _
                      Format(RV_Width, "#,##0.00") & "X" & _
                      Format(RV_Long, "#,##0.00")
        H1_ZaiRyo.Caption = RTrim(RV_Zaist) & Space(25 - Len(RTrim(RV_Zaist))) & "/" & _
                            Format(RV_Itatu, "0000.00") & "x" & _
                            Format(RV_Width, "0000.00") & "x" & _
                            Format(RV_Long, "0000.00")
       Call H1_ZaiBox_Set
       H1lb_ZaiKan = Rv_KanriNo
       
       gSL_Select = "SELECT * FROM 材料取引先マスタ WHERE 材料管理番号 = " & Rv_KanriNo
       Call ZTMRead(gSL_Select)
       H1lb_HinBn(2) = ZTM.自社品番
    
       For cnt = 0 To 3
           If H1Op_Hinku(cnt).Value = True Then
              iH1_Hinbn = H1lb_HinBn(cnt)
              cnt = 4
           End If
       Next
       
       Call iH1_Torcd_LostFocus
    End If
End Sub
'-----------------------------------------------------------------------' 2009.10.28 add end

'+--------------------+
'+     H0Op_Tanku     +
'+--------------------+
Private Sub H0Op_Tanku_Click(Index As Integer)
    H0Op_Tanku(0).BackColor = mLostColor
    H0Op_Tanku(1).BackColor = mLostColor
    H0Op_Tanku(Index).BackColor = mGotColor
    
    If Index = 0 Then
        OpH1_HinKb(1).Caption = "品番マスタ"
        Back_Hinbn.Caption = " 品　番"
        Back_ZHinbn.Visible = False
        H1Op_Hinku(0).Visible = False
        H1Op_Hinku(1).Visible = False
    Else
        OpH1_HinKb(1).Caption = "材料マスタ"
        Back_Hinbn.Caption = " 品　番"
        Back_ZHinbn.Visible = True
        H1Op_Hinku(0).Visible = True
        H1Op_Hinku(1).Visible = True
    End If
    
    H1Op_Hinku(0).Visible = H0Op_Tanku(1).Value
    b_HinNM(0).Visible = False
    H1lb_HinBn(0).Visible = False
    
    For cnt = 1 To 3
        H1Op_Hinku(cnt).Visible = H0Op_Tanku(1).Value
        b_HinNM(cnt).Visible = H0Op_Tanku(1).Value
        H1lb_HinBn(cnt).Visible = H0Op_Tanku(1).Value
        b_HinNM(cnt).Top = Hinku_top(cnt)
        H1lb_HinBn(cnt).Top = Hinku_top(cnt)
    Next
End Sub

'''''+--------------------+
'''''+     H1Op_Hinku     +
'''''+--------------------+
''''Private Sub H1Op_Hinku_Click(INDEX As Integer)
''''    H1Op_Hinku(0).BackColor = mLostColor
''''    H1Op_Hinku(1).BackColor = mLostColor
''''    H1Op_Hinku(2).BackColor = mLostColor
''''    H1Op_Hinku(3).BackColor = mLostColor
''''    H1Op_Hinku(INDEX).BackColor = mGotColor
''''
''''    H1Op_Hinku(INDEX).Value = True
'''''    If Index = 0 Then                      ' 2009.10.05 del start
'''''       Back_ZHinbn.Caption = "発注品番"
'''''    Else
'''''       Back_ZHinbn.Caption = "受給品番"
'''''    End If                                 ' 2009.10.05 del end
''''
''''End Sub
'+---------------------------+
'+      H1Op_Hinku           +
'+---------------------------+
Private Sub H1Op_Hinku_Click(Index As Integer)
    H1Op_Hinku(0).BackColor = mLostColor
    H1Op_Hinku(1).BackColor = mLostColor
    H1Op_Hinku(2).BackColor = mLostColor
    H1Op_Hinku(3).BackColor = mLostColor
    H1Op_Hinku(Index).BackColor = mGotColor
    
    If H0Op_Tanku(1).Value = True Then
       For cnt = 0 To 3
           b_HinNM(cnt).Visible = True
           H1lb_HinBn(cnt).Visible = True
       Next
    End If
    
    b_HinNM(Index).Visible = False
    H1lb_HinBn(Index).Visible = False

    gStr = RTrim(iH1_Hinbn.Text)
    iH1_Hinbn.Text = RTrim(H1lb_HinBn(Index))
    H1lb_HinBn(gMaeIndex) = gStr
   'Call HinbnSet(0)

    If H0Op_Tanku(1).Value Then
       Call HinbnToLabel
       Select Case Index
          Case 0:
            iH1_Hinbn.Text = H1lb_HinBn(0).Caption
            b_HinNM(1).Top = Hinku_top(1)
            H1lb_HinBn(1).Top = Hinku_top(1)
            b_HinNM(2).Top = Hinku_top(2)
            H1lb_HinBn(2).Top = Hinku_top(2)
            b_HinNM(3).Top = Hinku_top(3)
            H1lb_HinBn(3).Top = Hinku_top(3)
         
         Case 1:
            iH1_Hinbn.Text = H1lb_HinBn(1).Caption
            b_HinNM(0).Top = Hinku_top(1)
            H1lb_HinBn(0).Top = Hinku_top(1)
            b_HinNM(2).Top = Hinku_top(2)
            H1lb_HinBn(2).Top = Hinku_top(2)
            b_HinNM(3).Top = Hinku_top(3)
            H1lb_HinBn(3).Top = Hinku_top(3)
          
          Case 2:
            iH1_Hinbn.Text = H1lb_HinBn(2).Caption
            b_HinNM(0).Top = Hinku_top(1)
            H1lb_HinBn(0).Top = Hinku_top(1)
            b_HinNM(1).Top = Hinku_top(2)
            H1lb_HinBn(1).Top = Hinku_top(2)
            b_HinNM(3).Top = Hinku_top(3)
            H1lb_HinBn(3).Top = Hinku_top(3)

          Case 3:
            iH1_Hinbn.Text = H1lb_HinBn(3).Caption
            b_HinNM(0).Top = Hinku_top(1)
            H1lb_HinBn(0).Top = Hinku_top(1)
            b_HinNM(1).Top = Hinku_top(2)
            H1lb_HinBn(1).Top = Hinku_top(2)
            b_HinNM(2).Top = Hinku_top(3)
            H1lb_HinBn(2).Top = Hinku_top(3)
            Call H1_ZaiBox_Set
          Case Else:
       End Select
    End If
    gMaeIndex = Index
    
    If H1Op_Hinku(3).Value = True Then
       iH1_Hinbn.Visible = False
       H1_ZaiBox.Top = 525 '525 '660 : 660:試験用 / 525:正しい位置
       H1_ZaiBox.Visible = True
    Else
       H1_ZaiBox.Visible = False
       iH1_Hinbn.Visible = True
    End If
    
    If H1Op_Hinku(3).Value = True Then
       Call iH1_Zaist.SetFocus
    Else
       Call iH1_Hinbn.SetFocus
    End If
End Sub
'---------------------------------------' 2009.10.28 add start
'+-----------------------------+        '
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
'---------------------------------------' 2009.10.28 add end

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
    If TorNmGet(iH1_Torcd, 9, 1) Then
       H1lb_Tornm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iH1_SYmd
'+-----------------------------+
Private Sub iH1_SYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_SYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_EYmd
'+-----------------------------+
Private Sub iH1_EYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_EYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Tan
'+-----------------------------+
Private Sub iB1_Tan_GotFocus()
    Set ActObj = iB1_Tan
    iB1_Tan.Tag = "S"
End Sub
Private Sub iB1_Tan_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Tan_LostFocus()
    If CisFun.Val2(iB1_Tan) <> 0 Then
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 7) = iB1_Tan
    Else
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 7) = ""
    End If
End Sub

'+-----------------------------+
'+  B1Btn_Umu
'+-----------------------------+
Private Sub B1Btn_Umu_GotFocus()
    Set ActObj = B1Btn_Umu
End Sub
Private Sub B1Btn_Umu_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub B1Btn_Umu_LostFocus()
    If CisFun.Val2(iB1_Tan) <> 0 Then
        If B1Btn_Umu = "無" Then
            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 8) = "無"
        Else
            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 8) = "有"
        End If
    End If
End Sub
Private Sub B1Btn_Umu_Click()
    If B1Btn_Umu = "有" Then
        B1Btn_Umu.Caption = "無"
    Else
        B1Btn_Umu.Caption = "有"
    End If
End Sub

'+-----------------------------+
'+  B1Btn_Kari
'+-----------------------------+
Private Sub B1Btn_Kari_GotFocus()
    Set ActObj = B1Btn_Kari
End Sub
Private Sub B1Btn_Kari_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub B1Btn_Kari_LostFocus()
    If CisFun.Val2(iB1_Tan) <> 0 Then
        If B1Btn_Kari = "仮" Then
            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 9) = "仮"
        Else
            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 9) = "正"
        End If
    End If
End Sub
Private Sub B1Btn_Kari_Click()
    If B1Btn_Kari = "正" Then
        B1Btn_Kari.Caption = "仮"
    Else
        B1Btn_Kari.Caption = "正"
    End If
End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If iB1_Tan.Tag <> "" Then
       Call TextToVsGrid(Val(Input_Area.Tag))
       Call Grid_RowChange(vbKeyUp)
    End If
    iB1_Tan.SetFocus
End Sub

'+-----------------------------+
'+  F_TabDummy
'+-----------------------------+
Private Sub F_TabDummy_GotFocus()
    If iB1_Tan.Tag <> "" Then
       Call TextToVsGrid(Val(Input_Area.Tag))
       Call Grid_RowChange(vbKeyDown)
    End If
    iB1_Tan.SetFocus
End Sub

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
'  Call CisFrm.TimeSet(1)                   ' 2009.10.08 del
   CisFrm.TimeSet CisFun.INI_FileID, 110    ' 2009.10.08 add
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
'****************************
'*                          *
'*      HEAD/BODY 処理      *
'*                          *
'****************************
'****************************
'*      入力切替 処理        *
'****************************
Private Function DispChange(Pro As String)
    Dim H0Mode      As Boolean
    Dim H0Color     As Long
    Dim H1Mode      As Boolean
    Dim H1Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    H0Mode = False
    H0Color = gPLostSel
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
       Case "H0"
          H0Mode = True
          H0Color = gPGotSel
       Case "H1"
          H1Mode = True
          H1Color = gPGotSel
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
    End Select
    
    H0_Area1.Enabled = H0Mode
    H0_Area1.BackColor = H0Color
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    PB_Look.Visible = H1Mode
    PB_CAN.Visible = Not H0Mode
    PB_END.Visible = H0Mode
    
    Call GridTextEnabled(B1Mode)
    
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
            If .Name Like "Chk" & wbc_Enm & "_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
                .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        CisVsGrid.Clear
    End If
    If HeadBody = "B1" Then
        CisVsGrid.Clear
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H0" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        If H0Op_Tanku(1).Value = True Then
           H1Op_Hinku(0).Value = True
        Else
            For cnt = 0 To 3
                H1Op_Hinku(cnt).Visible = False
                H1lb_HinBn(cnt).Visible = False
                b_HinNM(cnt).Visible = False
            Next
        End If
        Call H1Op_Hinku_Click(0)
        gMaeIndex = 0
        Call iH1_SYmd.SetFocus          ' 2009.10.05 add
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        VsGrid.Row = 1
        Call VsGridToText(VsGrid.Row)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If H0Op_Tanku(0).Value Then
           If Not DBPut Then GoTo ReturnPress_Ed
        Else
           If Not DBPutZ Then GoTo ReturnPress_Ed
        End If
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        '-------------------------------' 2009.10.08 add start
        If H0Op_Tanku(1).Value = True Then
           For cnt = 0 To 3: H1lb_HinBn(cnt).Caption = "": Next
           H1lb_ZaiKan.Caption = ""
           H1_ZaiRyo.Caption = ""
        End If
        '-------------------------------' 2009.10.08 add end
        iH1_Torcd.SetFocus
        GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
        
'---------------------------------------------------------------' 2009.10.06 add start
Private Sub HinbnToLabel()
    If H0Op_Tanku(1).Value = False Then Exit Sub
    
    If H1Op_Hinku(0).Value = True Then
       H1lb_HinBn(0).Caption = Trim(iH1_Hinbn.Text)
    ElseIf H1Op_Hinku(1).Value = True Then
       H1lb_HinBn(1).Caption = Trim(iH1_Hinbn.Text)
    ElseIf H1Op_Hinku(2).Value = True Then
       H1lb_HinBn(2).Caption = Trim(iH1_Hinbn.Text)
    Else
       H1lb_HinBn(3).Caption = "": H1_ZaiRyo.Caption = "": iH1_Hinbn.Text = ""
       If Trim(iH1_Zaist.Text) <> "" Then
          H1lb_HinBn(3).Caption = Trim(iH1_Zaist.Text) & " " & _
                                  Format(CisFun.Val2(iH1_Itatu.Text), "###0.00") & "X" & _
                                  Format(CisFun.Val2(iH1_Width.Text), "###0.00") & "X" & _
                                  Format(CisFun.Val2(iH1_Long.Text), "###0.00")
          H1_ZaiRyo.Caption = Trim(iH1_Zaist.Text) & Space(25 - Len(Trim(iH1_Zaist.Text))) & "/" & _
                              Format(CisFun.Val2(iH1_Itatu.Text), "0000.00") & "x" & _
                              Format(CisFun.Val2(iH1_Width.Text), "0000.00") & "x" & _
                              Format(CisFun.Val2(iH1_Long.Text), "0000.00")
          iH1_Hinbn.Text = H1lb_HinBn(3).Caption
       End If
    End If

    For cnt = 0 To 3
        If Trim(H1lb_HinBn(cnt)) <> "" Then Exit For
    Next
    If cnt > 3 Then
        H1lb_ZaiKan.Caption = ""        ' 全ての品番が空白ならば材料管理番号を無しにする
    End If
End Sub
'---------------------------------------------------------------' 2009.10.06 add end

'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
    Call HinbnToLabel       ' 2009.10.07 add
    
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
    
    '============= 対象期間(開始必須) ==============
    If RTrim(iH1_SYmd) = "" Then
        CisFun.MB_MSG(1) = "     　【  対象期間 必須！ 】     "
        CisFun.MB_MSG(3) = "     対象期間開始 を入力して下さい。      "
        CisFun.MBOX
        iH1_SYmd.SetFocus
        Exit Function
    End If
    '============= 対象期間(開始異常) ==============
    If Not iH1_SYmd.cChkResult Then
        CisFun.MB_MSG(1) = "     　【  対象期間 異常！ 】     "
        CisFun.MB_MSG(3) = "     対象期間開始 を再入力して下さい。      "
        CisFun.MBOX
        iH1_SYmd.SetFocus
        Exit Function
    End If
    '============= 対象期間(終了必須) ==============
    If RTrim(iH1_EYmd) = "" Then
        CisFun.MB_MSG(1) = "     　【  対象期間 必須！ 】     "
        CisFun.MB_MSG(3) = "     対象期間終了 を入力して下さい。      "
        CisFun.MBOX
        iH1_EYmd.SetFocus
        Exit Function
    End If
    '============= 対象期間(終了異常) ==============
    If Not iH1_EYmd.cChkResult Then
        CisFun.MB_MSG(1) = "     　【  対象期間 異常！ 】     "
        CisFun.MB_MSG(3) = "     対象期間終了 を再入力して下さい。      "
        CisFun.MBOX
        iH1_EYmd.SetFocus
        Exit Function
    End If
    
    If RTrim(iH1_SYmd) > RTrim(iH1_EYmd) Then
        CisFun.MB_MSG(1) = "     　【  対象期間 異常！ 】     "
        CisFun.MB_MSG(3) = "     対象期間 開始 ＜　終了にして下さい。      "
        CisFun.MBOX
        iH1_EYmd.SetFocus
        Exit Function
    End If

'    '================ 支給先(必須) =================
'    If Trim(iH1_Torcd) = "" Then
'        CisFun.MB_MSG(2) = "    支給先を入力して下さい            "
'        CisFun.MBOX
'        iH1_Torcd.SetFocus
'        Exit Function
'    End If
'    '=========== 支給先(マスタチェック) ============
'    If Not TorNmGet(iH1_Torcd, 3) Then
'        CisFun.MB_MSG(2) = "    支給先　未登録！（取引先マスタ）            "
'        CisFun.MBOX
'        iH1_Torcd.SetFocus
'        Exit Function
'    End If
    
' マスタ入力 SELECT文 ｾｯﾄ
    
    If Not DBInput("INV") Then
        If H0Op_Tanku(0).Value Then
           CisFun.MB_MSG(2) = "    支給単価マスタ　未登録！           "
        Else
           CisFun.MB_MSG(2) = "    材料支給単価マスタ　未登録！　　　　"
        End If
        CisFun.MBOX
        iH1_Torcd.SetFocus
        Exit Function
    End If
       
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call GridTextEnabled(False)
'ﾒｯｾｰｼﾞﾎﾞｯｸｽ設定
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'
    With VsGrid
        For gCnt1 = 1 To .Rows - 1
            If RTrim(.TextMatrix(gCnt1, CisVsGrid.SelectCol)) <> "" Then
            '<< 修正の時 >>
                If CisFun.Val2(.TextMatrix(gCnt1, 13)) <> 0 Then
                    If CisFun.Val2(.TextMatrix(gCnt1, 7)) = 0 Then
                        CisFun.MB_MSG(1) = "    　修正の時、単価＝０は不可！            "
                        CisFun.MB_MSG(3) = "    　　" & gCnt1 & "行目            "
                        CisFun.MBOX
                        Set ErrObj = iB1_Tan
                        GoTo Body1Chk_Err
                    End If
                End If
            End If
        Next gCnt1
    End With
'
'
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    VsGrid.Row = gCnt1
    Call VsGridToText(VsGrid.Row)
    Call GridTextEnabled(True)
    ErrObj.SetFocus
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("B1")
    
    If H0Op_Tanku(0).Value Then
       '=*=*=*=*=*=*=*=*=*=*=*=*=*=*
       '=*     単価区分(部品)     *=
       '=*=*=*=*=*=*=*=*=*=*=*=*=*=*
       If Not STWCreate Then GoTo DBInput_Ed
        
       gSL_Select = ""
       gSL_Select = gSL_Select & " SELECT JT.*,ISNULL(HN.表示品番,JT.品番) 表示品番,"
       gSL_Select = gSL_Select & "        CASE WHEN ISNULL(TR.略称, '') <> ''"
       gSL_Select = gSL_Select & "             THEN TR.取引先CD + ':' + TR.略称"
       gSL_Select = gSL_Select & "             ELSE TR.取引先 END 取引先名,"
       gSL_Select = gSL_Select & "        '' 材料管理番号"
       gSL_Select = gSL_Select & "   FROM 支給単価ワーク JT"
       gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番マスタ HN"
       gSL_Select = gSL_Select & "             ON HN.品番 = JT.品番"
       gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR"
       gSL_Select = gSL_Select & "             ON TR.取引先CD = JT.取引先"
       gSL_Select = gSL_Select & "            AND TR.取引先区分 = 1"
       gSL_Select = gSL_Select & "        LEFT OUTER JOIN (SELECT 品番, 取引先"
       gSL_Select = gSL_Select & "                           FROM 品番取引先マスタ"
       gSL_Select = gSL_Select & "                          WHERE 品目 = 1"
'       gSL_Select = gSL_Select & "                          WHERE (品目 = 0"
'       gSL_Select = gSL_Select & "                             OR (品目 = 1 AND 手配区分 = '1'))"
       gSL_Select = gSL_Select & "                          GROUP BY 品番, 取引先) HT"
       gSL_Select = gSL_Select & "             ON HT.品番 = JT.品番"
       gSL_Select = gSL_Select & "            AND HT.取引先 = JT.取引先"
       gSL_Select = gSL_Select & "  WHERE 処理端末 = HOST_NAME()"
       gSL_Select = gSL_Select & "    AND ISNULL(TR.取引先CD,'') <> ''"
       gSL_Select = gSL_Select & "  ORDER BY JT.取引先,JT.品番"
        
       If Not STMRead(gSL_Select, 0, 1) Then
              Call STMClose
          Else
              Call GridSet
              DBInput = True
          End If
          
    Else    ' ========  材料側  ====================
       If Not ZTWCreate Then GoTo DBInput_Ed
        
       gSL_Select = ""
       gSL_Select = gSL_Select & " SELECT ZW.*, ZA.表示寸法 表示品番,"                  ' 2009.10.07 upd : ZW.品番 --> ZA.表示寸法
       gSL_Select = gSL_Select & "        CASE WHEN ISNULL(TR.略称, '') <> ''"
       gSL_Select = gSL_Select & "             THEN TR.取引先CD + ':' + TR.略称"
       gSL_Select = gSL_Select & "             ELSE TR.取引先 END 取引先名"
       gSL_Select = gSL_Select & "      , ZM.発注品番, ZM.受給品番, ZM.自社品番"        ' 2009.10.07 add
       gSL_Select = gSL_Select & "      , ZA.材質, ZA.板厚, ZA.幅, ZA.長さ "            ' 2009.10.07 add
       gSL_Select = gSL_Select & "   FROM 材料支給単価ワーク ZW"
'       gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ ZM"
'       If H1Op_Hinku(0).Value Then
'          gSL_Select = gSL_Select & "             ON ZM.発注品番 = ZW.品番"
'       Else
'          gSL_Select = gSL_Select & "             ON ZM.受給品番 = ZW.品番"
'       End If
       gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ ZA"                ' 2009.10.07 add
       gSL_Select = gSL_Select & "             ON ZA.材料管理番号 = ZW.材料管理番号"    ' 2009.10.07 add
       gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料取引先マスタ ZM"
       gSL_Select = gSL_Select & "             ON ZM.材料管理番号 = ZW.材料管理番号"
       gSL_Select = gSL_Select & "            AND ZM.取引先 = ZW.取引先"
       gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR"
       gSL_Select = gSL_Select & "             ON TR.取引先CD = ZW.取引先"
       gSL_Select = gSL_Select & "            AND TR.取引先区分 = 1"
       gSL_Select = gSL_Select & "  WHERE 処理端末 = HOST_NAME()"
       gSL_Select = gSL_Select & "    AND ISNULL(TR.取引先CD,'') <> ''"
       gSL_Select = gSL_Select & "  ORDER BY ZW.取引先,ZW.品番"
        
       If Not STMRead(gSL_Select, 0, 1) Then
           Call STMClose
       Else
           Call GridSet
           DBInput = True
       End If
    End If
    
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function

'+----------------------------+
'+     支給単価ワーク作成     +
'+----------------------------+
Private Function STWCreate() As Boolean
       '=*=*=*=*=*=*=*=*=*=*=*=*=*=*
       '=*     単価区分(部品)     *=
       '=*=*=*=*=*=*=*=*=*=*=*=*=*=*
    STWCreate = False

    gSL_Select = ""
    gSL_Select = gSL_Select & " DELETE FROM 支給単価ワーク WHERE 処理端末 = HOST_NAME()"
    CisDB.SQL = gSL_Select
    If Not CisDB.DBExec Then Exit Function
    
'-----------------------------------------------------------' 2009.10.29 del start
'    With CisFun
'        .Where_CreateKB = NewDocument
'
'        If Not OpH1_HinKb(0).Value Then
'           .Where_Create Numeric, "品目", 1
'        End If
'        '=========== 取引先 ============
'        .Where_Create AlphaMe, "取引先", RTrim(iH1_Torcd)
'        '============ 品番 =============
'        .Where_Create AlphaMe, "品番", RTrim(iH1_Hinbn)
'
'    End With
'-----------------------------------------------------------' 2009.10.29 del end

    If OpH1_HinKb(0).Value Then
        '-----------------------------------------------------------------------------------' 2009.10.29 add start
        With CisFun
            .Where_CreateKB = NewDocument
            '=========== 取引先 ============
            .Where_Create AlphaMe, "ST.取引先", RTrim(iH1_Torcd)
            '============ 品番 =============
            .Where_Create AlphaMe, "ST.品番", RTrim(iH1_Hinbn)
        End With
        '-----------------------------------------------------------------------------------' 2009.10.29 add end
        '<< 支給単価マスタより品番抽出  >>
        gSL_Select = ""
        gSL_Select = gSL_Select & " INSERT INTO 支給単価ワーク"
        gSL_Select = gSL_Select & "             (処理端末,品番,取引先,前回履歴NO,履歴NO)"
        gSL_Select = gSL_Select & " SELECT HOST_NAME(), ST.品番, ST.取引先, 0, 0"           ' 2009.10.29 upd : ST. 追加
        gSL_Select = gSL_Select & "   FROM 支給単価マスタ ST"                               ' 2009.10.29 upd : ST. 追加
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番取引先マスタ HT "            ' 2009.10.29 add
        gSL_Select = gSL_Select & "               ON  HT.品番 = ST.品番 "                   ' 2009.10.29 add
        gSL_Select = gSL_Select & CisFun.Where_Phrase
        '-----------------------------------------------------------------------------------' 2009.10.29 add start
        If CisFun.Where_Phrase = "" Then
           gSL_Select = gSL_Select & "  WHERE HT.手配区分 = 4 "
        Else
           gSL_Select = gSL_Select & "    AND HT.手配区分 = 4 "
        End If
        '-----------------------------------------------------------------------------------' 2009.10.29 add end
        gSL_Select = gSL_Select & "  GROUP BY ST.品番,ST.取引先"                            ' 2009.10.29 upd : ST. 追加
        CisDB.SQL = gSL_Select
        If Not CisDB.DBExec Then Exit Function
    Else
        '-----------------------------------------------------------------------------------' 2009.10.29 add start
        With CisFun
            .Where_CreateKB = NewDocument
            .Where_Create Numeric, "品目", 1
            '=========== 取引先 ============
            .Where_Create AlphaMe, "取引先", RTrim(iH1_Torcd)
            '============ 品番 =============
            .Where_Create AlphaMe, "品番", RTrim(iH1_Hinbn)
        End With
        '-----------------------------------------------------------------------------------' 2009.10.29 add end
        '<< 品番取引先マスタより品番抽出  >>
        gSL_Select = ""
        gSL_Select = gSL_Select & " INSERT INTO 支給単価ワーク"
        gSL_Select = gSL_Select & "             (処理端末,品番,取引先,前回履歴NO,履歴NO)"
        gSL_Select = gSL_Select & " SELECT HOST_NAME(), 品番, 取引先, 0, 0"
        gSL_Select = gSL_Select & "   FROM 品番取引先マスタ"
        gSL_Select = gSL_Select & CisFun.Where_Phrase
        'gSL_Select = gSL_Select & " AND  品目 = 0"
        gSL_Select = gSL_Select & "    AND 手配区分 = 4 "       ' 2009.10.29 add
        gSL_Select = gSL_Select & " GROUP BY 品番, 取引先"
        CisDB.SQL = gSL_Select
        If Not CisDB.DBExec Then Exit Function
    End If
    '============ ワーク作成 ===============
    With CisFun
        .Where_CreateKB = NewDocument
        '=========== 取引先 ============
        .Where_Create AlphaMe, "ST.取引先", RTrim(iH1_Torcd)
        '============ 品番 =============
        .Where_Create AlphaMe, "ST.品番", RTrim(iH1_Hinbn)
    End With

    gSL_Select = ""
    gSL_Select = gSL_Select & " UPDATE 支給単価ワーク SET"
    gSL_Select = gSL_Select & "        履歴NO = ISNULL(ST1.履歴NO, 0),"
    gSL_Select = gSL_Select & "        単価 = 0,"
    gSL_Select = gSL_Select & "        仮単価区分 = 0,"
    gSL_Select = gSL_Select & "        有無償区分 = 0,"
    gSL_Select = gSL_Select & "        適用開始日付 = ISNULL(ST1.適用開始日付, ''),"
    gSL_Select = gSL_Select & "        適用終了日付 = ISNULL(ST1.適用終了日付, ''),"
    gSL_Select = gSL_Select & "        前回履歴NO = ISNULL(ST1.履歴NO, 0),"
    gSL_Select = gSL_Select & "        前回単価 =  ISNULL(ST1.単価, 0),"
    gSL_Select = gSL_Select & "        前回仮単価区分 = ISNULL(ST1.仮単価区分, ''),"
    gSL_Select = gSL_Select & "        前回有無賞区分 = ISNULL(ST1.有無償区分, '')"
    gSL_Select = gSL_Select & "   FROM 支給単価ワーク ST"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 支給単価マスタ ST1"
    gSL_Select = gSL_Select & "             ON ST1.取引先 = ST.取引先"
    gSL_Select = gSL_Select & "            AND ST1.品番 = ST.品番"
    gSL_Select = gSL_Select & "            AND ST1.履歴NO = (SELECT MAX(ST1S.履歴NO)履歴NO FROM 支給単価マスタ ST1S"
    gSL_Select = gSL_Select & "                               WHERE ST1S.取引先 = ST.取引先"
    gSL_Select = gSL_Select & "                                 AND ST1S.品番 = ST.品番"
    gSL_Select = gSL_Select & "                               GROUP BY ST1S.取引先,ST1S.品番)"
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    CisDB.SQL = gSL_Select
    If Not CisDB.DBExec Then Exit Function
    
    STWCreate = True

End Function

'+--------------------------------+
'+     材料支給単価ワーク作成     +
'+--------------------------------+
Private Function ZTWCreate() As Boolean
    ZTWCreate = False

    gSL_Select = ""
    gSL_Select = gSL_Select & " DELETE FROM 材料支給単価ワーク WHERE 処理端末 = HOST_NAME()"
    CisDB.SQL = gSL_Select
    If Not CisDB.DBExec Then Exit Function
    
    With CisFun
        .Where_CreateKB = NewDocument
        '=========== 取引先 ============
        .Where_Create AlphaMe, "ZT.取引先", RTrim(iH1_Torcd)
        '============ 品番 =============
        If OpH1_HinKb(0).Value Then                                    ' 2009.10.05 upd start : 品区に自社品番、材質を追加
           If H1Op_Hinku(0).Value Then
               .Where_Create AlphaMe, "ZM.発注品番", RTrim(iH1_Hinbn)
           ElseIf H1Op_Hinku(1).Value Then
               .Where_Create AlphaMe, "ZM.受給品番", RTrim(iH1_Hinbn)
           ElseIf H1Op_Hinku(2).Value Then
               .Where_Create AlphaMe, "ZM.自社品番", RTrim(iH1_Hinbn)
           End If
        Else
           If H1Op_Hinku(0).Value Then
               .Where_Create AlphaMe, "ZT.発注品番", RTrim(iH1_Hinbn)
           ElseIf H1Op_Hinku(1).Value Then
               .Where_Create AlphaMe, "ZT.受給品番", RTrim(iH1_Hinbn)
           ElseIf H1Op_Hinku(2).Value Then
               .Where_Create AlphaMe, "ZT.自社品番", RTrim(iH1_Hinbn)
           End If
        End If
        If H1Op_Hinku(3).Value = True Then
           If CisFun.Val2(iH1_Itatu) = 0 Then iH1_Itatu = ""
           If CisFun.Val2(iH1_Width) = 0 Then iH1_Width = ""
           If CisFun.Val2(iH1_Long) = 0 Then iH1_Long = ""
           If Not (Trim(iH1_Zaist) = "" And Trim(iH1_Itatu) = "" And Trim(iH1_Width) = "" And Trim(iH1_Long) = "") Then
              If Trim(iH1_Itatu) = "" And Trim(iH1_Width) = "" And Trim(iH1_Long) = "" Then
                 If .Where_Phrase = "" Then
                    .Where_Phrase = .Where_Phrase & "   WHERE ZA.材質 LIKE '" & Trim(iH1_Zaist) & "%'"
                 Else
                    .Where_Phrase = .Where_Phrase & "     AND ZA.材質 LIKE '" & Trim(iH1_Zaist) & "%'"
                 End If
              Else
                 .Where_Create AlphaMe, "ZA.材質", RTrim(iH1_Zaist)
                 If RTrim(iH1_Itatu) <> "" Then .Where_Create AlphaMe, "ZA.板厚", RTrim(iH1_Itatu)
                 If RTrim(iH1_Width) <> "" Then .Where_Create AlphaMe, "ZA.幅", RTrim(iH1_Width)
                 If RTrim(iH1_Long) <> "" Then .Where_Create AlphaMe, "ZA.長さ", RTrim(iH1_Long)
              End If
           End If
        End If
    End With                                                            ' 2009.10.05 upd end

    If OpH1_HinKb(0).Value Then
        '<< 支給単価マスタより品番抽出  >>
        gSL_Select = ""
        gSL_Select = gSL_Select & " INSERT INTO 材料支給単価ワーク"
        gSL_Select = gSL_Select & "             (処理端末, 材料管理番号, 品番, 取引先, 前回履歴NO, 履歴NO)"
'
'        If H1Op_Hinku(0).Value Then                                                                                     ' 2009.10.05 upd start : 品区別作成
'           gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZM.発注品番,''), ZT.取引先, 0, 0"
'        ElseIf H1Op_Hinku(1).Value Then
'           gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZM.受給品番,''), ZT.取引先, 0, 0"
'        ElseIf H1Op_Hinku(2).Value Then
'           gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZM.自社品番,''), ZT.取引先, 0, 0"
'        Else
'           gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZA.表示寸法,''), ZT.取引先, 0, 0"
'        End If                                                                                                          ' 2009.10.05 upd end
        gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZA.材質,''), ZT.取引先, 0, 0"   ' 2009.10.09 upd : 上記品番別から材質に変更
'
        gSL_Select = gSL_Select & "   FROM 材料支給単価マスタ ZT"
'        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ ZM"
'        gSL_Select = gSL_Select & "             ON ZM.材料管理番号 = ZT.材料管理番号"
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料取引先マスタ ZM"
        gSL_Select = gSL_Select & "             ON ZM.材料管理番号 = ZT.材料管理番号"
        ''gSL_Select = gSL_Select & "            AND ZM.品目 = 1"
        gSL_Select = gSL_Select & "            AND ZM.取引先 = ZT.取引先"
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ ZA"               ' 2009.10.05 add
        gSL_Select = gSL_Select & "             ON ZA.材料管理番号 = ZT.材料管理番号"   ' 2009.10.05 add
        gSL_Select = gSL_Select & CisFun.Where_Phrase
        If CisFun.Where_Phrase = "" Then                                                                ' 2009.10.08 add start
           gSL_Select = gSL_Select & "  WHERE ZT.材料管理番号 != '' AND ZA.材料管理番号 != '' "
        Else
           gSL_Select = gSL_Select & "    AND ZT.材料管理番号 != '' AND ZA.材料管理番号 != '' "
        End If                                                                                          ' 2009.10.08 add end
        gSL_Select = gSL_Select & "    AND ZM.手配区分 = 4 "                            ' 2009.10.29 add
'
'        If H1Op_Hinku(0).Value Then                                                                     ' 2009.10.05 upd start : 品区別作成
'           gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZM.発注品番,''), ZT.取引先, ZT.材料管理番号"
'        ElseIf H1Op_Hinku(1).Value Then
'           gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZM.受給品番,''), ZT.取引先, ZT.材料管理番号"
'        ElseIf H1Op_Hinku(2).Value Then
'           gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZM.自社品番,''), ZT.取引先, ZT.材料管理番号"
'        Else
'           gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZA.表示寸法,''), ZT.取引先, ZT.材料管理番号"
'        End If                                                                                          ' 2009.10.05 upd end
        gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZA.材質,''), ZT.取引先, ZT.材料管理番号"   ' 2009.10.09 upd : 上記品番別から材質に変更
'
        CisDB.SQL = gSL_Select
        If Not CisDB.DBExec Then Exit Function
    Else
        '<< 材料取引先マスタより品番抽出  >>
        gSL_Select = ""
        gSL_Select = gSL_Select & " INSERT INTO 材料支給単価ワーク"
        gSL_Select = gSL_Select & "             (処理端末, 材料管理番号, 品番, 取引先, 前回履歴NO, 履歴NO)"
'
'        If H1Op_Hinku(0).Value Then                                                                     ' 2009.10.05 upd start : 品区別作成
'           gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZT.発注品番,''), ZT.取引先, 0, 0"
'        ElseIf H1Op_Hinku(1).Value Then
'           gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZT.受給品番,''), ZT.取引先, 0, 0"
'        ElseIf H1Op_Hinku(2).Value Then
'           gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZT.自社品番,''), ZT.取引先, 0, 0"
'        Else
'           gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZA.表示寸法,''), ZT.取引先, 0, 0"
'        End If                                                                                          ' 2009.10.05 upd end
        gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZA.材質,''), ZT.取引先, 0, 0"   ' 2009.10.09 upd : 上記品番別から材質に変更
'
        gSL_Select = gSL_Select & "   FROM 材料取引先マスタ ZT"
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ ZA"
        gSL_Select = gSL_Select & "             ON ZA.材料管理番号 = ZT.材料管理番号"
        gSL_Select = gSL_Select & CisFun.Where_Phrase
        If CisFun.Where_Phrase = "" Then                                                                ' 2009.10.08 add start
           gSL_Select = gSL_Select & "  WHERE ZT.材料管理番号 != '' AND ZA.材料管理番号 != '' "
        Else
           gSL_Select = gSL_Select & "    AND ZT.材料管理番号 != '' AND ZA.材料管理番号 != '' "
        End If                                                                                          ' 2009.10.08 add end
        gSL_Select = gSL_Select & "    AND ZT.手配区分 = 4 "                            ' 2009.10.29 add
        'gSL_Select = gSL_Select & "   AND (ZT.品目 = 0 OR (ZT.品目 = 1 AND ZT.手配区分 = '1'))"
        ''gSL_Select = gSL_Select & "   AND ZT.品目 = 1 "
'
'        If H1Op_Hinku(0).Value Then                                                                                     ' 2009.10.05 upd start : 品区別作成
'           gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZT.発注品番,''), ZT.取引先, ZT.材料管理番号"
'        ElseIf H1Op_Hinku(1).Value Then
'           gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZT.受給品番,''), ZT.取引先, ZT.材料管理番号"
'        ElseIf H1Op_Hinku(2).Value Then
'           gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZT.自社品番,''), ZT.取引先, ZT.材料管理番号"
'        Else
'           gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZA.表示寸法,''), ZT.取引先, ZT.材料管理番号"
'        End If                                                                                                          ' 2009.10.05 add end
        gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZA.材質,''), ZT.取引先, ZT.材料管理番号"   ' 2009.10.09 upd : 上記品番別から材質に変更
'
        CisDB.SQL = gSL_Select
        If Not CisDB.DBExec Then Exit Function
    End If
'Call Test_LogOutput(gSL_Select, 1)  ' デバッグ用
    '============ ワーク作成 ===============
    With CisFun
        .Where_CreateKB = NewDocument
        '=========== 取引先 ============
        .Where_Create AlphaMe, "ST.取引先", RTrim(iH1_Torcd)
        '============ 品番 =============
        .Where_Create AlphaMe, "ST.品番", RTrim(iH1_Hinbn)
    End With

    gSL_Select = ""
    gSL_Select = gSL_Select & " UPDATE 材料支給単価ワーク SET"
    gSL_Select = gSL_Select & "        履歴NO = ISNULL(ST1.履歴NO,0),"
    gSL_Select = gSL_Select & "        単価 = 0,"
    gSL_Select = gSL_Select & "        仮単価区分 = 0,"
    gSL_Select = gSL_Select & "        適用開始日付 = ISNULL(ST1.適用開始日付, ''),"
    gSL_Select = gSL_Select & "        適用終了日付 = ISNULL(ST1.適用終了日付, ''),"
    gSL_Select = gSL_Select & "        有無償区分 = 0,"
    gSL_Select = gSL_Select & "        前回履歴NO = ISNULL(ST1.履歴NO, 0),"
    gSL_Select = gSL_Select & "        前回単価 =  ISNULL(ST1.単価,0),"
    gSL_Select = gSL_Select & "        前回仮単価区分 = ISNULL(ST1.仮単価区分, ''),"
    gSL_Select = gSL_Select & "        前回有無賞区分 = ISNULL(ST1.有無償区分, '')"
    gSL_Select = gSL_Select & "   FROM 材料支給単価ワーク ST"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料支給単価マスタ ST1"
    gSL_Select = gSL_Select & "             ON ST1.取引先 = ST.取引先"
    gSL_Select = gSL_Select & "            AND ST1.材料管理番号 = ST.材料管理番号"
    gSL_Select = gSL_Select & "            AND ST1.履歴NO = (SELECT MAX(ST1S.履歴NO)履歴NO"
    gSL_Select = gSL_Select & "                                FROM 材料支給単価マスタ ST1S"
    gSL_Select = gSL_Select & "                                     LEFT OUTER JOIN 材料マスタ ZM"
    gSL_Select = gSL_Select & "                                          ON ZM.材料管理番号 = ST1S.材料管理番号"
    gSL_Select = gSL_Select & "                               WHERE ST1S.取引先 = ST.取引先"
    gSL_Select = gSL_Select & "                                 AND ST1S.材料管理番号 = ST.材料管理番号"
    gSL_Select = gSL_Select & "                               GROUP BY ST1S.取引先, ST1S.材料管理番号)"
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    CisDB.SQL = gSL_Select
    If Not CisDB.DBExec Then Exit Function
    
    ZTWCreate = True

End Function

'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------------------+
'+                                  +
'+          追加･更新 処理          +
'+                                  +
'+----------------------------------+
Private Function DBPut()
    DBPut = False
    Dim RirekiNo  As Integer

    K_Sykbnm.Caption = SyoriNM
'    K_Sykbnm.Visible = True        ' 2009.10.13 del
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     追加・更新処理を行います      "
        .MB_Title = "追加・更新処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call VsGridToText(Val(Input_Area.Tag))
            Call GridTextEnabled(True)
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err
    Message.Visible = True
    DoEvents
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
'   マスタ項目移送
    For gIndex = 1 To VsGrid.Rows - 1
        If RTrim(VsGrid.TextMatrix(gIndex, CisVsGrid.SelectCol)) <> "" Then
           SyoriKB = "ADD"
           gSL_Select = "SELECT MAX(履歴NO)履歴NO FROM 支給単価マスタ"
           gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(VsGrid.TextMatrix(gIndex, 11)) & "'"
           gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(RTrim(VsGrid.TextMatrix(gIndex, 12))) & "'"
           If STMRead(gSL_Select, 1) Then
              RirekiNo = STM.履歴NO + 1
           Else
              RirekiNo = 1
           End If
            
           Call ItemsClearSTM
           With STM
               '===== 取引先 =====
               .取引先 = RTrim(VsGrid.TextMatrix(gIndex, 11))
               '====== 品番 ======
               .品番 = RTrim(VsGrid.TextMatrix(gIndex, 12))
               '===== 履歴NO =====
               .履歴NO = RirekiNo
               '=== 適用期間 ===
               .適用開始日付 = RTrim(iH1_SYmd)
               .適用終了日付 = RTrim(iH1_EYmd)
               '====== 単価 ======
               .単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 7))
               '===== 作成者 =====
               .作成者 = gTanto
               '=== 有無償区分 ===
               If RTrim(VsGrid.TextMatrix(gIndex, 8)) = "有" Then
                  .有無償区分 = 0
               Else
                  .有無償区分 = 1
               End If
               '=== 仮単価区分 ===
               If RTrim(VsGrid.TextMatrix(gIndex, 9)) = "正" Then
                  .仮単価区分 = ""
               Else
                  .仮単価区分 = "1"
               End If
           End With
           
       ' ( 追加処理 )
           Call STMInsert
'           iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")       ' 2009.10.09 del
               
        End If
    Next gIndex
'
' ( 追加処理 )
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    Message.Visible = False
    DoEvents
    DBPut = True
    Exit Function
DBPut_Err:
    Message.Visible = False
    DoEvents
    CisFun.ErrorBox
    End
End Function

'+--------------------------------------+
'+                                      +
'+         追加･更新 処理(材料)         +
'+                                      +
'+--------------------------------------+
Private Function DBPutZ()
    DBPutZ = False
    Dim RirekiNo  As Integer

    K_Sykbnm.Caption = SyoriNM
'    K_Sykbnm.Visible = True        ' 2009.10.13 del
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     追加・更新処理を行います      "
        .MB_Title = "追加・更新処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call VsGridToText(Val(Input_Area.Tag))
            Call GridTextEnabled(True)
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err
    Message.Visible = True
    DoEvents
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
'   マスタ項目移送
    For gIndex = 1 To VsGrid.Rows - 1
        If RTrim(VsGrid.TextMatrix(gIndex, CisVsGrid.SelectCol)) <> "" Then
           SyoriKB = "ADD"
           gSL_Select = "SELECT MAX(履歴NO)履歴NO FROM 材料支給単価マスタ"
           gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(VsGrid.TextMatrix(gIndex, 11)) & "'"
           gSL_Select = gSL_Select & "   AND 材料管理番号 = '" & RTrim(VsGrid.TextMatrix(gIndex, 17)) & "'"
           If ZSMRead(gSL_Select, 1) Then
              RirekiNo = ZSM.履歴NO + 1
           Else
              RirekiNo = 1
           End If
            
           Call ItemsClearZSM
           With ZSM
               '===== 取引先 =====
               .取引先 = RTrim(VsGrid.TextMatrix(gIndex, 11))
               '===== 履歴NO =====
               .履歴NO = RirekiNo
               If SyoriKB = "ADD" Then
                  '=== 適用期間 ===
                  .適用開始日付 = RTrim(iH1_SYmd)
                  .適用終了日付 = RTrim(iH1_EYmd)
               End If
               '====== 単価 ======
               .単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 7))
               '===== 作成者 =====
               .作成者 = gTanto
               '=== 有無償区分 ===
               If RTrim(VsGrid.TextMatrix(gIndex, 8)) = "有" Then
                  .有無償区分 = 0
               Else
                  .有無償区分 = 1
               End If
               '=== 仮単価区分 ===
               If RTrim(VsGrid.TextMatrix(gIndex, 9)) = "正" Then
                  .仮単価区分 = ""
               Else
                  .仮単価区分 = "1"
               End If
               '==== 材料管理番号 ====
               .材料管理番号 = RTrim(VsGrid.TextMatrix(gIndex, 17))
           End With
           
       ' ( 追加処理 )
           Call ZSMInsert
        
'       ' ( 影響データ自動更新 )
'           Call ZTanAutoUpdate(gIndex)     ' del 2009.10.05
            
        End If
        
    Next gIndex
'
' ( 追加処理 )
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    Message.Visible = False
    DoEvents
    DBPutZ = True
    Exit Function
DBPut_Err:
    Message.Visible = False
    DoEvents
    CisFun.ErrorBox
    End
End Function

'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 左中, あり, 18, 1, "支給先")
        Call .FixedSet(中中, 左中, あり, 35, 2, "品  番")
        Call .FixedSet(中中, 右中, あり, 10, 3, "前回開始")
        Call .FixedSet(中中, 右中, あり, 10, 4, "前回終了")
        Call .FixedSet(中中, 中中, あり, 2, 5, "仮")
        Call .FixedSet(中中, 右中, あり, 13, 6, "前回単価")
        Call .FixedSet(中中, 右中, あり, 13, 7, "単価")
        Call .FixedSet(中中, 中中, あり, 3, 8, "有")
        Call .FixedSet(中中, 中中, あり, 3, 9, " 仮 ")
        Call .FixedSet(中中, 中中, あり, 0, 10, "処理")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 11, "x取引先")
        Call .FixedSet(中中, 左中, なし, 0, 12, "x品番")
        Call .FixedSet(中中, 左中, なし, 0, 13, "x履歴NO")
        Call .FixedSet(中中, 左中, なし, 0, 14, "x単価")
        Call .FixedSet(中中, 右中, あり, 4, 15, "x有無")
        Call .FixedSet(中中, 左中, なし, 0, 16, "x仮単価区分")
        Call .FixedSet(中中, 左中, なし, 0, 17, "x材料管理番号")
        Call .FixedSet(中中, 左中, なし, 0, 18, "Select")
        Call .FixedSet(中中, 左中, なし, 0, 19, "x発注品番")    ' 2009.10.07 add
        Call .FixedSet(中中, 左中, なし, 0, 20, "x受給品番")    ' 2009.10.07 add
        Call .FixedSet(中中, 左中, なし, 0, 21, "x自社品番")    ' 2009.10.07 add
        Call .FixedSet(中中, 左中, なし, 0, 22, "x材質寸法")    ' 2009.10.07 add
        
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9") ',19,20,21,22")
        .InitDisp
    End With
    
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(CisFun.Val2(ID))
        Next ID
        If Ttl_Kensu + 1 > .Rows Then
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not STM_RDSTS
        'Grid Row,Col
        ID = ID + 1
        With VsGrid
            .TextMatrix(ID, GridCol(0)) = ID
            
            '============= 支給先 ==============
            .TextMatrix(ID, GridCol(1)) = RTrim(STM.取引先名)
            .TextMatrix(ID, GridCol(11)) = RTrim(STM.取引先)
            
            '============== 品番 ===============
            If H0Op_Tanku(0).Value = True Then                      ' 2009.10.07 add
               .TextMatrix(ID, GridCol(2)) = RTrim(STM.品番)  'RTrim(STM.表示品番)
               .TextMatrix(ID, GridCol(12)) = RTrim(STM.品番)
            Else                                                    ' 2009.10.07 add
              '.TextMatrix(ID, GridCol(2)) = Trim(STM.表示品番)     ' 2009.10.09 del    ' 2009.10.07 add
              '.TextMatrix(ID, GridCol(12)) = ""                    ' 2009.10.09 del    ' 2009.10.07 add
               .TextMatrix(ID, GridCol(2)) = Trim(STM.材質) & _
                        " " & Format(STM.板厚, "#,##0.00") & _
                        "x" & Format(STM.幅, "#,##0.00") & _
                        "x" & Format(STM.長さ, "#,##0.00")          ' 2009.10.09 add
            End If                                                  ' 2009.10.07 add
            '========== 前回適用期間 ===========
            If STM.適用開始日付 <> "" Then
               gStr = ""
               gStr = gStr & CisFun.Mid2(STM.適用開始日付, 1, 4) & "/"
               gStr = gStr & CisFun.Mid2(STM.適用開始日付, 5, 2) & "/"
               gStr = gStr & CisFun.Mid2(STM.適用開始日付, 7, 2)
               .TextMatrix(ID, GridCol(3)) = gStr
            End If
            
            If STM.適用終了日付 <> "" Then
               gStr = ""
               gStr = gStr & CisFun.Mid2(STM.適用終了日付, 1, 4) & "/"
               gStr = gStr & CisFun.Mid2(STM.適用終了日付, 5, 2) & "/"
               gStr = gStr & CisFun.Mid2(STM.適用終了日付, 7, 2)
               .TextMatrix(ID, GridCol(4)) = gStr
            End If
            
            '========= 前回仮単価区分 ==========
            If STM.前回単価 <> 0 Then
                If STM.前回仮単価区分 = "" Then
                    .TextMatrix(ID, GridCol(5)) = "正"
                Else
                    .TextMatrix(ID, GridCol(5)) = "仮"
                End If
            End If
            
            '============ 前回単価 =============
            If STM.前回単価 <> 0 Then
                .TextMatrix(ID, GridCol(6)) = Format(STM.前回単価, "#,##0.00")
            End If
            
            If OpH1_TanKb(1).Value Then
                '============ 単価 =============
                If STM.前回単価 <> 0 Then
                    .TextMatrix(ID, GridCol(7)) = Format(STM.前回単価, "#,##0.00")
                    .TextMatrix(ID, GridCol(14)) = STM.前回単価
                    '====== 有無償区分 ======
                    If STM.有無償区分 = 0 Then
                        .TextMatrix(ID, GridCol(8)) = "有"
                        .TextMatrix(ID, GridCol(15)) = "有"
                    Else
                        .TextMatrix(ID, GridCol(8)) = "無"
                        .TextMatrix(ID, GridCol(15)) = "無"
                    End If
                    '====== 仮単価区分 ======
                    If STM.仮単価区分 = "" Then
                        .TextMatrix(ID, GridCol(9)) = "正"
                        .TextMatrix(ID, GridCol(16)) = "正"
                    Else
                        .TextMatrix(ID, GridCol(9)) = "仮"
                        .TextMatrix(ID, GridCol(16)) = "仮"
                    End If
                End If
            Else
                .TextMatrix(ID, GridCol(14)) = STM.単価
            End If
            
            .TextMatrix(ID, GridCol(13)) = RTrim(STM.履歴NO)
            
            '========== 材料管理番号 ===========
            .TextMatrix(ID, GridCol(17)) = RTrim(STM.材料管理番号)
            
            '========== 材料品番情報 ===========                                    ' 2009.10.07 add start
            .TextMatrix(ID, GridCol(19)) = RTrim(STM.発注品番)
            .TextMatrix(ID, GridCol(20)) = RTrim(STM.受給品番)
            .TextMatrix(ID, GridCol(21)) = RTrim(STM.自社品番)
            .TextMatrix(ID, GridCol(22)) = RTrim(STM.材質) & Space(25 - Len(Trim(STM.材質))) & "/" & _
                                           Format(STM.板厚, "0000.00") & "x" & _
                                           Format(STM.幅, "0000.00") & "x" & _
                                           Format(STM.長さ, "000.00")              ' 2009.10.07 add end
            
        End With
        Call STMReadNext(1)
    Loop

    With VsGrid
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    Call STMClose(1)
    
    Call VsGridToHnbn(1)    ' 2009.10.08 add
End Sub
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    Input_Area.Visible = EnabledFlg
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'テキストからグリッドに退避
    Call TextToVsGrid(Val(Input_Area.Tag))
    gInt = VsGrid.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > VsGrid.Rows - 1 Then
            GoTo Grid_RowChange_End
        Else
            If RTrim(VsGrid.TextMatrix(gInt + 1, 1)) = "" Then
                GoTo Grid_RowChange_End
            End If
        End If
        gInt = gInt + 1
    End If
    'PageUp
    If KeyCode = vbKeyPageUp Then
        If gInt - CisVsGrid.DispMax < 1 Then
            gInt = 1
        Else
            gInt = gInt - CisVsGrid.DispMax
        End If
    End If
    'PageDown
    If KeyCode = vbKeyPageDown Then
        If gInt + CisVsGrid.DispMax > VsGrid.Rows - 1 Then
            gInt = VsGrid.Rows - 1
        Else
            gInt = gInt + CisVsGrid.DispMax
        End If
    End If
    VsGrid.Row = gInt
'グリッドからテキストに表示
    Call VsGridToText(VsGrid.Row)
    Call VsGridToHnbn(VsGrid.Row)       ' 2009.10.08
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With VsGrid
        '======== 単価 ========
        If CisFun.Val2(iB1_Tan) <> 0 Then
            .TextMatrix(tRow, 7) = Format(CisFun.Val2(iB1_Tan), "#,##0.00")
        Else
            .TextMatrix(tRow, 7) = ""
        End If
        
        '===== 有無償区分 =====
        If CisFun.Val2(iB1_Tan) <> 0 Then
            If B1Btn_Umu = "有" Then
                .TextMatrix(tRow, 8) = "有"
            Else
                .TextMatrix(tRow, 8) = "無"
            End If
        Else
                .TextMatrix(tRow, 8) = ""
        End If
        
        '===== 仮単価区分 =====
        If CisFun.Val2(iB1_Tan) <> 0 Then
            If B1Btn_Kari = "正" Then
                .TextMatrix(tRow, 9) = "正"
            Else
                .TextMatrix(tRow, 9) = "仮"
            End If
        Else
                .TextMatrix(tRow, 9) = ""
        End If
        
    End With
    SaveRow = VsGrid.Row
    Call UpdateCheck(tRow)
    VsGrid.Row = SaveRow
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With VsGrid
        Input_Area.Tag = tRow
        Input_Area.Move Input_Area.Left, VsGrid.CellTop, Input_Area.Width, VsGrid.CellHeight
        '単価
        If CisFun.Val2(.TextMatrix(tRow, 7)) <> 0 Then
            iB1_Tan = Format(CisFun.Val2(.TextMatrix(tRow, 7)), "#,##0.00")
        Else
            iB1_Tan = ""
        End If
        
        ' 有無償区分
        If RTrim(.TextMatrix(tRow, 8)) = "無" Then
            B1Btn_Umu = "無"
        Else
            B1Btn_Umu = "有"
        End If
        
        '仮区分
        If RTrim(.TextMatrix(tRow, 9)) <> "仮" Then
            B1Btn_Kari = "正"
        Else
            B1Btn_Kari = "仮"
        End If
    End With
End Sub
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call VsGridToText(VsGrid.Row)
End Sub
Private Sub VSGrid_Click()
    Call TextToVsGrid(Val(Input_Area.Tag))
    If VsGrid.TextMatrix(VsGrid.Row, 0) = "" Then Exit Sub
    Call GridTextEnabled(True)
    Call VsGridToText(VsGrid.Row)
    Call VsGridToHnbn(VsGrid.Row)       ' 2009.10.13
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
Private Sub VSGrid_DblClick()
    If RTrim(VsGrid.TextMatrix(VsGrid.Row, 1)) = "" Then Exit Sub
    If VsGrid.Col = 1 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     " & VsGrid.Row & " 行目をクリアしますか？      "
            .MB_Title = ""
            .MB_Button = OK_CAN
            If Not .MBOX Then Exit Sub
        End With
        iB1_Tan = ""
        B1Btn_Kari = "正"
        Call TextToVsGrid(Val(Input_Area.Tag))
    End If
    If VsGrid.Col = 5 And CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 6)) <> 0 Then   ' 2009.10.13 upd : VsGrid.Col = 4 --> 5
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     " & VsGrid.Row & " 行目 前回単価を単価にコピーしますか？      "
            .MB_Title = ""
            .MB_Button = OK_CAN
            If Not .MBOX Then Exit Sub
        End With
        iB1_Tan = VsGrid.TextMatrix(VsGrid.Row, 6)
        If VsGrid.TextMatrix(VsGrid.Row, 5) = "正" Then
            B1Btn_Kari = "正"
        Else
            B1Btn_Kari = "仮"
        End If
        Call TextToVsGrid(Val(Input_Area.Tag))
    End If
End Sub

'+---------------------------------------------+
'+      変更チェック
'+---------------------------------------------+
Private Sub UpdateCheck(sRow As Integer)
    VsGrid.Row = sRow
    If CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 13)) = 0 Then
        If CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 7)) <> 0 Then
            If RTrim(VsGrid.TextMatrix(VsGrid.Row, CisVsGrid.SelectCol)) = "" Then
                VsGrid.TextMatrix(VsGrid.Row, 10) = "追加"
                iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
                Call CisVsGrid.vsColor(個別, &HFFFFC0)
            End If
        Else
            If RTrim(VsGrid.TextMatrix(VsGrid.Row, CisVsGrid.SelectCol)) <> "" Then
                VsGrid.TextMatrix(VsGrid.Row, 10) = ""
                iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) - 1, "###")
                Call CisVsGrid.vsColor(消去個別)
            End If
        End If
    Else
        If CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 7)) <> CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 14)) Or _
           RTrim(VsGrid.TextMatrix(VsGrid.Row, 8)) <> RTrim(VsGrid.TextMatrix(VsGrid.Row, 15)) Or _
           RTrim(VsGrid.TextMatrix(VsGrid.Row, 9)) <> RTrim(VsGrid.TextMatrix(VsGrid.Row, 16)) Then
            If RTrim(VsGrid.TextMatrix(VsGrid.Row, CisVsGrid.SelectCol)) = "" Then
                VsGrid.TextMatrix(VsGrid.Row, 10) = "修正"
                iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
                Call CisVsGrid.vsColor(個別, &HFFFFC0)
            End If
        Else
            If RTrim(VsGrid.TextMatrix(VsGrid.Row, CisVsGrid.SelectCol)) <> "" Then
                VsGrid.TextMatrix(VsGrid.Row, 10) = "*"
                iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) - 1, "###")
                Call CisVsGrid.vsColor(消去個別)
            End If
        End If
    End If
End Sub

Private Sub VsGridToHnbn(iRow As Integer)
    If H0Op_Tanku(1).Value = False Then Exit Sub
    'Dim c As Integer: For c = 0 To VsGrid.Cols - 1: H1lb_HinBn(1) = VsGrid.TextMatrix(iRow, c): Next
    
    H1lb_HinBn(0) = VsGrid.TextMatrix(iRow, 19)
    H1lb_HinBn(1) = VsGrid.TextMatrix(iRow, 20)
    H1lb_HinBn(2) = VsGrid.TextMatrix(iRow, 21)
    H1_ZaiRyo.Caption = VsGrid.TextMatrix(iRow, 22)
    H1lb_HinBn(3) = Trim(Mid(H1_ZaiRyo.Caption, 1, 25)) & " " & _
                    Format(Trim(Mid(H1_ZaiRyo.Caption, 27, 7)), "###0.00") & "x" & _
                    Format(Trim(Mid(H1_ZaiRyo.Caption, 35, 7)), "###0.00") & "x" & _
                    Format(Trim(Mid(H1_ZaiRyo.Caption, 43, 7)), "###0.00")

'    Call HinbnToLabel
    If H1Op_Hinku(0).Value = True Then
       iH1_Hinbn.Text = H1lb_HinBn(0).Caption
       b_HinNM(1).Top = Hinku_top(1)
       H1lb_HinBn(1).Top = Hinku_top(1)
       b_HinNM(2).Top = Hinku_top(2)
       H1lb_HinBn(2).Top = Hinku_top(2)
       b_HinNM(3).Top = Hinku_top(3)
       H1lb_HinBn(3).Top = Hinku_top(3)
         
    ElseIf H1Op_Hinku(1).Value = True Then
       iH1_Hinbn.Text = H1lb_HinBn(1).Caption
       b_HinNM(0).Top = Hinku_top(1)
       H1lb_HinBn(0).Top = Hinku_top(1)
       b_HinNM(2).Top = Hinku_top(2)
       H1lb_HinBn(2).Top = Hinku_top(2)
       b_HinNM(3).Top = Hinku_top(3)
       H1lb_HinBn(3).Top = Hinku_top(3)
          
    ElseIf H1Op_Hinku(2).Value = True Then
       iH1_Hinbn.Text = H1lb_HinBn(2).Caption
       b_HinNM(0).Top = Hinku_top(1)
       H1lb_HinBn(0).Top = Hinku_top(1)
       b_HinNM(1).Top = Hinku_top(2)
       H1lb_HinBn(1).Top = Hinku_top(2)
       b_HinNM(3).Top = Hinku_top(3)
       H1lb_HinBn(3).Top = Hinku_top(3)

    Else
       iH1_Hinbn.Text = H1lb_HinBn(3).Caption
       b_HinNM(0).Top = Hinku_top(1)
       H1lb_HinBn(0).Top = Hinku_top(1)
       b_HinNM(1).Top = Hinku_top(2)
       H1lb_HinBn(1).Top = Hinku_top(2)
       b_HinNM(2).Top = Hinku_top(3)
       H1lb_HinBn(2).Top = Hinku_top(3)
       Call H1_ZaiBox_Set
    End If
    
    If H1Op_Hinku(3).Value = True Then
      'iH1_Hinbn.Visible = False
       H1_ZaiBox.Top = 525 '525 '660 : 660:試験用 / 525:正しい位置
       H1_ZaiBox.Visible = True
       iH1_Zaist.Text = Trim(Mid(H1_ZaiRyo.Caption, 1, 25))
       iH1_Itatu.Text = Trim(Mid(H1_ZaiRyo.Caption, 27, 7))
       iH1_Width.Text = Trim(Mid(H1_ZaiRyo.Caption, 35, 7))
       iH1_Long.Text = Trim(Mid(H1_ZaiRyo.Caption, 43, 7))
    Else
       H1_ZaiBox.Visible = False
       iH1_Hinbn.Visible = True
       For cnt = 0 To 2
           If H1Op_Hinku(cnt).Value = True Then
              iH1_Hinbn.Text = H1lb_HinBn(cnt)
              Exit For
           End If
       Next
    End If

    H1lb_ZaiKan.Caption = VsGrid.TextMatrix(iRow, 17)

End Sub

''+-------------------------------------+
''+      Debug用 SQL記録                +
''+-------------------------------------+
Private Sub Test_LogOutput(str As String, Msg As Integer)
'Exit Sub   '=========================================== 実行しないときはこの行を生かすこと =========================
    Dim dt01 As String
    Dim dt02 As String
    Dim f_No As Integer
    
    If Msg <> 0 Then MsgBox (str)
    dt01 = CStr(DateTime.Now)
    dt02 = Mid(dt01, 1, 4) & "." & _
           Mid(dt01, 6, 2) & "." & _
           Mid(dt01, 9, 2)

    f_No = FreeFile                             'ファイル番号の取得
    Open "D:\TestLog\SQL_" & dt02 & ".log" _
          For Append As #f_No                   'ファイルを追加モードで開く
    Print #f_No, "[" & Me.Name & "] " & dt01 & _
                 " ----------------------------------------" & Chr(13) & Chr(10) & _
                 str & Chr(13) & Chr(10)        'ファイルへ書き込む"
    Close #f_No                                 'ファイルを閉じる
End Sub
