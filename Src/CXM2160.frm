VERSION 5.00
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXM2160 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "期首受注単価　一括メンテナンス"
   ClientHeight    =   11235
   ClientLeft      =   120
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
      Left            =   630
      Top             =   3240
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   503
      BackColor       =   14737632
      Caption         =   "納　 入 　先"
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
   Begin Cis3D_v60.CIS3D CIS3D8 
      Height          =   285
      Left            =   5640
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BackColor       =   14737632
      Caption         =   "前回開始"
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
   Begin Cis3D_v60.CIS3D CIS3D9 
      Height          =   285
      Left            =   3450
      Top             =   3240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      BackColor       =   14737632
      Caption         =   "品  番"
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
   Begin Cis3D_v60.CIS3D CIS3D11 
      Height          =   285
      Left            =   6990
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BackColor       =   14737632
      Caption         =   "前回終了"
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
   Begin Cis3D_v60.CIS3D CIS3D12 
      Height          =   285
      Left            =   8340
      Top             =   3240
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      BackColor       =   14737632
      Caption         =   "仮"
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
   Begin Cis3D_v60.CIS3D CIS3D13 
      Height          =   285
      Left            =   10110
      Top             =   3240
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   503
      BackColor       =   14737632
      Caption         =   "単　 価"
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
   Begin Cis3D_v60.CIS3D CIS3D14 
      Height          =   285
      Left            =   8640
      Top             =   3240
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   503
      BackColor       =   14737632
      Caption         =   "前回単価"
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
   Begin Cis3D_v60.CIS3D CIS3D15 
      Height          =   285
      Left            =   11550
      Top             =   3240
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   503
      BackColor       =   14737632
      Caption         =   "部品単価"
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
   Begin Cis3D_v60.CIS3D CIS3D17 
      Height          =   285
      Left            =   13020
      Top             =   3240
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   503
      BackColor       =   14737632
      Caption         =   "梱包単価"
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
   Begin Cis3D_v60.CIS3D CIS3D18 
      Height          =   285
      Left            =   14460
      Top             =   3240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      BackColor       =   14737632
      Caption         =   "仮"
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
      TabIndex        =   16
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
      Left            =   11685
      ScaleHeight     =   765
      ScaleWidth      =   3540
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   600
      Width           =   3600
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
      Height          =   7860
      Left            =   90
      ScaleHeight     =   7800
      ScaleWidth      =   15135
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3105
      Width           =   15195
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   7725
         Left            =   45
         Top             =   30
         Width           =   15060
         _ExtentX        =   26564
         _ExtentY        =   13626
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
            Height          =   7665
            Left            =   30
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   30
            Width           =   15015
            _Version        =   196608
            _ExtentX        =   26485
            _ExtentY        =   13520
            _StockProps     =   228
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   $"CXM2160.frx":0000
            Rows            =   23
            Cols            =   24
            BackColor       =   12648384
            BackColorFixed  =   14737632
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   330
            GridLines       =   2
            SelectionMode   =   1
            Begin Cis3D_v60.CIS3D Input_Area 
               Height          =   345
               Left            =   9840
               Top             =   960
               Width           =   4800
               _ExtentX        =   8467
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
               Begin CisText_V60.CisText iB1_Tan 
                  Height          =   375
                  Left            =   0
                  TabIndex        =   8
                  Top             =   0
                  Width           =   1500
                  _ExtentX        =   2646
                  _ExtentY        =   661
                  cFormat         =   "#,###,##0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#,###,##0.00"
                  MaxLength       =   12
                  cDataType       =   2
                  cDataReplace    =   1
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "######0.00"
                  cILength        =   7
               End
               Begin CisText_V60.CisText iB1_BTan 
                  Height          =   375
                  Left            =   1470
                  TabIndex        =   9
                  Top             =   0
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   661
                  cFormat         =   "#,###,##0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#,###,##0.00"
                  MaxLength       =   12
                  cDataType       =   2
                  cDataReplace    =   1
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "######0.00"
                  cILength        =   7
               End
               Begin CisText_V60.CisText iB1_KTan 
                  Height          =   375
                  Left            =   2925
                  TabIndex        =   10
                  Top             =   0
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   661
                  cFormat         =   "#,###,##0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#,###,##0.00"
                  MaxLength       =   12
                  cDataType       =   2
                  cDataReplace    =   1
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "######0.00"
                  cILength        =   7
               End
               Begin Cis3D_v60.CIS3D B1Btn_Kari 
                  Height          =   315
                  Left            =   4395
                  Top             =   30
                  Width           =   390
                  _ExtentX        =   688
                  _ExtentY        =   556
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
         Begin VB.TextBox B_TabDummy 
            Height          =   330
            Left            =   225
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   120
            Width           =   240
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
      TabIndex        =   13
      Text            =   "Dummy"
      Top             =   0
      Width           =   210
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   1320
      Left            =   90
      ScaleHeight     =   1260
      ScaleWidth      =   13455
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1740
      Width           =   13515
      Begin Cis3D_v60.CIS3D CIS3D16 
         Height          =   1110
         Left            =   2595
         Top             =   90
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "納 入 先"
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
            Left            =   105
            TabIndex        =   2
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
            Left            =   1080
            Top             =   525
            Width           =   2595
            _ExtentX        =   4577
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
         Height          =   1110
         Left            =   75
         Top             =   90
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1958
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
            Height          =   315
            Left            =   105
            Top             =   390
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cPositionX      =   30
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_EYmd 
            Height          =   330
            Left            =   720
            TabIndex        =   1
            Top             =   720
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
            Left            =   720
            TabIndex        =   0
            Top             =   390
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
            Height          =   315
            Left            =   105
            Top             =   720
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            Caption         =   "終了"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
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
         Height          =   1110
         Left            =   9945
         Top             =   90
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1958
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
            Height          =   225
            Index           =   1
            Left            =   315
            TabIndex        =   5
            Top             =   735
            Width           =   1350
         End
         Begin VB.OptionButton OpH1_HinKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "単価マスタ"
            Height          =   225
            Index           =   0
            Left            =   315
            TabIndex        =   4
            Top             =   465
            Value           =   -1  'True
            Width           =   1350
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   1110
         Left            =   11790
         Top             =   90
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1958
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
            Height          =   225
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   465
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton OpH1_TanKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "前回単価"
            Height          =   225
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   735
            Width           =   1125
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   1110
         Left            =   6375
         Top             =   90
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "品　　　番"
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
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   195
            TabIndex        =   3
            Top             =   525
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
         End
      End
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
      Caption         =   "【 受注単価マスタ　一括入力  】"
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
End
Attribute VB_Name = "CXM2160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   売上単価マスタ　一括入力
'**       フォームID    :   CXM2160
'**       処理概要      :
'**
'**       作  成  日    :   2008/02/25  By CIS
'**       変  更  日    :   2009/10/28  By CIS : 土岐津産業対応
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

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub form_load()
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
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = VsGrid
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 21
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
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
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
    
    If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      iH1_Torcd.SetFocus
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
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub             ' 2009.10.28 add
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_SYmd" Then iH1_SYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_EYmd" Then iH1_EYmd.ShowCalender: Exit Sub
End Sub

'---------------------------------------------------------------------------' 2009.10.28 add start
'+----------------------------+
'+       品番検索             +
'+----------------------------+
Private Sub Look_Hinbn()
'--<< 品番検索 >>--*
    RV_Left = 0
    RV_Top = 0
    
    CKK0030.Show vbModal
    Unload CKK0030
    Set CKK0030 = Nothing

    If RV_Rtn Then
        iH1_Hinbn = RV_Hinbn
    End If
End Sub
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
'---------------------------------------------------------------------------' 2009.10.28 add end

'+----------------------------+
'+       納入先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 3
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
    If TorNmGet(iH1_Torcd, 3, 1) Then
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
    B_TabDummy.Tag = "S"
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
    Call TankaEdit
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
        If B1Btn_Kari <> "仮" Then
            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 10) = "仮"
        Else
            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 10) = "正"
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
'+  iB1_BTan
'+-----------------------------+
Private Sub iB1_BTan_GotFocus()
    Set ActObj = iB1_BTan
End Sub
Private Sub iB1_BTan_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_BTan_LostFocus()
    If RTrim(iB1_BTan) <> "" Then
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 8) = iB1_BTan
    Else
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 8) = ""
    End If
    Call TankaEdit
End Sub
'+-----------------------------+
'+  iB1_KTan
'+-----------------------------+
Private Sub iB1_KTan_GotFocus()
    Set ActObj = iB1_KTan
    B_TabDummy.Tag = "E"
End Sub
Private Sub iB1_KTan_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_KTan_LostFocus()
    If RTrim(iB1_KTan) <> "" Then
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 9) = iB1_KTan
    Else
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 9) = ""
    End If
    Call TankaEdit
End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If B_TabDummy.Tag = "S" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyUp)
    End If
    If B_TabDummy.Tag = "E" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyDown)
    End If
    SendKeys "{Tab}"
End Sub

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   Call CisFrm.TimeSet(1)
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
    PB_Look.Visible = H1Mode
'    PB_ADD.Visible = H1Mode
'    PB_MNT.Visible = H1Mode
'    PB_DEL.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
'    PB_DSP.Visible = H1Mode
    PB_END.Visible = H1Mode
    
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
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        iH1_Torcd.SetFocus
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
    
    If Not DBInput("INV") Then
        CisFun.MB_MSG(2) = "    受注単価マスタ　未登録！           "
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
                If CisFun.Val2(.TextMatrix(gCnt1, 14)) <> 0 Then
                    If CisFun.Val2(.TextMatrix(gCnt1, 7)) = 0 Then
                        CisFun.MB_MSG(1) = "    　修正の時、単価＝０は不可！            "
                        CisFun.MB_MSG(3) = "    　　" & gCnt1 & "行目            "
                        CisFun.MBOX
                        Set ErrObj = iB1_Tan
                        GoTo Body1Chk_Err
                    End If
                End If
            ' 単価 = 部品単価 + 梱包単価 + その他費
                If CisFun.Val2(.TextMatrix(gCnt1, 7)) <> CisFun.Val2(.TextMatrix(gCnt1, 8)) + CisFun.Val2(.TextMatrix(gCnt1, 9)) And _
                    CisFun.Val2(.TextMatrix(gCnt1, 8)) + CisFun.Val2(.TextMatrix(gCnt1, 9)) <> 0 Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = "    " & gCnt1 & "行目  単価≠部品単価+梱包単価+その他です。           "
                        .MB_MSG(4) = "    　　 よろしいですか？                      "
                        .MB_Title = ""
                        .MB_Button = OK_CAN
                        If Not .MBOX Then
                            Set ErrObj = iB1_Tan
                            GoTo Body1Chk_Err
                        End If
                    End With
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
    
    If Not JTWCreate Then GoTo DBInput_Ed
        
    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT JT.*,ISNULL(HN.表示品番,JT.品番) 表示品番,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(TR.略称, '') <> ''"
    gSL_Select = gSL_Select & "             THEN TR.取引先CD + ':' + TR.略称"
    gSL_Select = gSL_Select & "             ELSE TR.取引先CD END 取引先名"
    gSL_Select = gSL_Select & "   FROM 受注単価ワーク JT"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "             ON HN.品番 = JT.品番"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "             ON TR.取引先CD = JT.取引先"
    gSL_Select = gSL_Select & "            AND TR.取引先区分 = 0"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN (SELECT 品番, 取引先"
    gSL_Select = gSL_Select & "                           FROM 品番取引先マスタ"
    gSL_Select = gSL_Select & "                          WHERE 品目 = 0"
    gSL_Select = gSL_Select & "                          GROUP BY 品番, 取引先) HT"
    gSL_Select = gSL_Select & "             ON HT.品番 = JT.品番"
    gSL_Select = gSL_Select & "            AND HT.取引先 = JT.取引先"
    gSL_Select = gSL_Select & "  WHERE 処理端末 = HOST_NAME()"
    gSL_Select = gSL_Select & "  AND   ISNULL(HT.品番,'') <> ''"
    gSL_Select = gSL_Select & "  AND   ISNULL(TR.取引先CD,'') <> ''"
    gSL_Select = gSL_Select & "  ORDER BY JT.取引先,JT.品番"
    
    If Not JTMRead(gSL_Select, 0, 1) Then
        Call JTMClose
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
        Call GridSet
        DBInput = True
    End If
    
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function

'+----------------------------+
'+     受注単価ワーク作成     +
'+----------------------------+
Private Function JTWCreate() As Boolean
    JTWCreate = False

    gSL_Select = ""
    gSL_Select = gSL_Select & " DELETE FROM 受注単価ワーク WHERE 処理端末 = HOST_NAME()"
    CisDB.SQL = gSL_Select
    If Not CisDB.DBExec Then Exit Function
    
    With CisFun
        .Where_CreateKB = NewDocument
        .Where_Create Numeric, "品目", 0
        '=========== 取引先 ============
        .Where_Create AlphaMe, "取引先", RTrim(iH1_Torcd)
        '============ 品番 =============
        .Where_Create AlphaMe, "品番", RTrim(iH1_Hinbn)
        
    End With

    If OpH1_HinKb(0).Value Then
        '<< 受注単価マスタより品番抽出  >>
        gSL_Select = ""
        gSL_Select = gSL_Select & " INSERT INTO 受注単価ワーク"
        gSL_Select = gSL_Select & "             (処理端末,品番,取引先,前回履歴NO,履歴NO)"
        gSL_Select = gSL_Select & " SELECT HOST_NAME(), 品番, 取引先, 0, 0 FROM 受注単価マスタ "
        gSL_Select = gSL_Select & CisFun.Where_Phrase
        gSL_Select = gSL_Select & "  GROUP BY 品番,取引先"
        CisDB.SQL = gSL_Select
        If Not CisDB.DBExec Then Exit Function
    Else
        '<< 品番取引先マスタより品番抽出  >>
        gSL_Select = ""
        gSL_Select = gSL_Select & " INSERT INTO 受注単価ワーク"
        gSL_Select = gSL_Select & "             (処理端末,品番,取引先,前回履歴NO,履歴NO)"
        gSL_Select = gSL_Select & " SELECT HOST_NAME(),品番,取引先,0,0 FROM 品番取引先マスタ "
        gSL_Select = gSL_Select & CisFun.Where_Phrase
        'gSL_Select = gSL_Select & "   AND 品目 = 0"
        gSL_Select = gSL_Select & " GROUP BY 品番,取引先"
        CisDB.SQL = gSL_Select
        If Not CisDB.DBExec Then Exit Function
    End If
    '============ ワーク作成 ===============
    With CisFun
        .Where_CreateKB = NewDocument
        '=========== 取引先 ============
        .Where_Create AlphaMe, "JT.取引先", RTrim(iH1_Torcd)
        '============ 品番 =============
        .Where_Create AlphaMe, "JT.品番", RTrim(iH1_Hinbn)
    End With

    gSL_Select = ""
    gSL_Select = gSL_Select & " UPDATE 受注単価ワーク SET"
    gSL_Select = gSL_Select & "        履歴NO = ISNULL(JT1.履歴NO,0),"
    gSL_Select = gSL_Select & "        単価 = 0,"
    gSL_Select = gSL_Select & "        仮単価区分 = 0,"
    gSL_Select = gSL_Select & "        部品単価 = 0,"
    gSL_Select = gSL_Select & "        梱包単価 = 0,"
    gSL_Select = gSL_Select & "        適用開始日付 = ISNULL(JT1.適用開始日付,''),"
    gSL_Select = gSL_Select & "        適用終了日付 = ISNULL(JT1.適用終了日付,''),"
    gSL_Select = gSL_Select & "        前回履歴NO = ISNULL(JT1.履歴NO,0),"
    gSL_Select = gSL_Select & "        前回単価 =  ISNULL(JT1.単価,0),"
    gSL_Select = gSL_Select & "        前回仮単価区分 = ISNULL(JT1.仮単価区分,''),"
    gSL_Select = gSL_Select & "        前回部品単価 = ISNULL(JT1.部品単価,0),"
    gSL_Select = gSL_Select & "        前回梱包単価 = IsNull(JT1.梱包単価, 0)"
    gSL_Select = gSL_Select & "   FROM 受注単価ワーク JT"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 受注単価マスタ JT1"
    gSL_Select = gSL_Select & "             ON JT1.取引先 = JT.取引先"
    gSL_Select = gSL_Select & "            AND JT1.品番 = JT.品番"
    gSL_Select = gSL_Select & "            AND JT1.履歴NO = (SELECT MAX(JT1S.履歴NO)履歴NO FROM 受注単価マスタ JT1S"
    gSL_Select = gSL_Select & "                               WHERE JT1S.取引先 = JT.取引先"
    gSL_Select = gSL_Select & "                                 AND JT1S.品番 = JT.品番"
    gSL_Select = gSL_Select & "                               GROUP BY JT1S.取引先,JT1S.品番)"
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    CisDB.SQL = gSL_Select
    If Not CisDB.DBExec Then Exit Function
    
    JTWCreate = True

End Function

'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Dim RirekiNo  As Integer

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
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
           gSL_Select = "SELECT MAX(履歴NO)履歴NO FROM 受注単価マスタ"
           gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(RTrim(VsGrid.TextMatrix(gIndex, 12))) & "'"
           gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(RTrim(VsGrid.TextMatrix(gIndex, 13))) & "'"
           If JTMRead(gSL_Select, 1) Then
              RirekiNo = JTM.履歴NO + 1
           Else
              RirekiNo = 1
           End If
            
           Call ItemsClearJTM
           With JTM
               '===== 取引先 =====
               .取引先 = RTrim(RTrim(VsGrid.TextMatrix(gIndex, 12)))
               '====== 品番 ======
               .品番 = RTrim(VsGrid.TextMatrix(gIndex, 13))
               '===== 履歴NO =====
               .履歴NO = RirekiNo
               '=== 適用期間 ===
               .適用開始日付 = RTrim(iH1_SYmd)
               .適用終了日付 = RTrim(iH1_EYmd)
               '====== 単価 ======
               .単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 7))
               '==== 部品単価 ====
               .部品単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 8))
               '==== 梱包単価 ====
               .梱包単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 9))
               '===== 作成者 =====
               .作成者 = gTanto
               If RTrim(VsGrid.TextMatrix(gIndex, 10)) = "正" Then      ' 2009.10.28 upd : "正式" --> "正"
                  .仮単価区分 = ""
               Else
                  .仮単価区分 = "1"
               End If
           End With
           
       ' ( 追加処理 )
           Call JTMInsert
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
        Call .FixedSet(中中, 左中, あり, 26, 1, "納入先")
        Call .FixedSet(中中, 左中, あり, 20, 2, "品  番")
        Call .FixedSet(中中, 右中, あり, 12, 3, "前回適用開始")
        Call .FixedSet(中中, 右中, あり, 12, 4, "前回適用終了")
        Call .FixedSet(中中, 中中, あり, 2, 5, "仮")
        Call .FixedSet(中中, 右中, あり, 13, 6, "前回単価")
        Call .FixedSet(中中, 右中, あり, 13, 7, "単価")
        Call .FixedSet(中中, 右中, あり, 13, 8, "部品単価")
        Call .FixedSet(中中, 右中, あり, 13, 9, "梱包単価")
        Call .FixedSet(中中, 中中, あり, 3, 10, " 仮 ")
        Call .FixedSet(中中, 中中, あり, 0, 11, "処理")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 12, "x取引先")
        Call .FixedSet(中中, 左中, なし, 0, 13, "x品番")
        Call .FixedSet(中中, 左中, なし, 0, 14, "x履歴NO")
        Call .FixedSet(中中, 左中, なし, 0, 15, "x前回部品単価")
        Call .FixedSet(中中, 左中, なし, 0, 16, "x前回梱包単価")
        Call .FixedSet(中中, 左中, なし, 0, 17, "x単価")
        Call .FixedSet(中中, 左中, なし, 0, 18, "x仮単価区分")
        Call .FixedSet(中中, 左中, なし, 0, 19, "x部品単価")
        Call .FixedSet(中中, 左中, なし, 0, 20, "x梱包単価")
        Call .FixedSet(中中, 左中, なし, 0, 21, "Select")
        
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10")
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
    Do Until Not JTM_RDSTS
        'Grid Row,Col
        ID = ID + 1
        With VsGrid
            .TextMatrix(ID, GridCol(0)) = ID
            '============= 納入先 ==============
            .TextMatrix(ID, GridCol(1)) = RTrim(JTM.取引先名)
            .TextMatrix(ID, GridCol(12)) = RTrim(JTM.取引先)
            
            '============== 品番 ===============
            .TextMatrix(ID, GridCol(2)) = RTrim(JTM.表示品番)
            .TextMatrix(ID, GridCol(13)) = RTrim(JTM.品番)
            
            '========== 前回適用期間 ===========
            If JTM.適用開始日付 <> "" Then
               gStr = ""
               gStr = gStr & CisFun.Mid2(JTM.適用開始日付, 1, 4) & "/"
               gStr = gStr & CisFun.Mid2(JTM.適用開始日付, 5, 2) & "/"
               gStr = gStr & CisFun.Mid2(JTM.適用開始日付, 7, 2)
               .TextMatrix(ID, GridCol(3)) = gStr
            End If
            
            If JTM.適用終了日付 <> "" Then
               gStr = ""
               gStr = gStr & CisFun.Mid2(JTM.適用終了日付, 1, 4) & "/"
               gStr = gStr & CisFun.Mid2(JTM.適用終了日付, 5, 2) & "/"
               gStr = gStr & CisFun.Mid2(JTM.適用終了日付, 7, 2)
               .TextMatrix(ID, GridCol(4)) = gStr
            End If
            
            '========= 前回仮単価区分 ==========
            If JTM.前回単価 <> 0 Then
                If JTM.前回仮単価区分 = "" Then
                    .TextMatrix(ID, GridCol(5)) = "正"
                Else
                    .TextMatrix(ID, GridCol(5)) = "仮"
                End If
            End If
            
            '============ 前回単価 =============
            If JTM.前回単価 <> 0 Then
                .TextMatrix(ID, GridCol(6)) = Format(JTM.前回単価, "#,##0.00")
            End If
            
            If OpH1_TanKb(1).Value Then
                '============ 単価 =============
                If JTM.前回単価 <> 0 Then
                    .TextMatrix(ID, GridCol(7)) = Format(JTM.前回単価, "#,##0.00")
                    .TextMatrix(ID, GridCol(17)) = JTM.前回単価
                End If
                
                '========= 仮単価区分 ==========
                If JTM.前回単価 <> 0 Then
                    If JTM.前回仮単価区分 = "" Then
                        .TextMatrix(ID, GridCol(10)) = "正式"
                        .TextMatrix(ID, GridCol(18)) = "正式"
                    Else
                        .TextMatrix(ID, GridCol(10)) = "仮"
                        .TextMatrix(ID, GridCol(18)) = "仮"
                    End If
                End If
            Else
                .TextMatrix(ID, GridCol(17)) = JTM.単価
                '============ 単価 =============
                If JTM.単価 <> 0 Then
                    .TextMatrix(ID, GridCol(7)) = Format(JTM.単価, "#,##0.00")
                End If
                
                '========= 仮単価区分 ==========
                If JTM.単価 <> 0 Then
                    If JTM.仮単価区分 = "" Then
                        .TextMatrix(ID, GridCol(10)) = "正式"
                        .TextMatrix(ID, GridCol(18)) = "正式"
                    Else
                        .TextMatrix(ID, GridCol(10)) = "仮"
                        .TextMatrix(ID, GridCol(18)) = "仮"
                    End If
                End If
            End If
            
            .TextMatrix(ID, GridCol(14)) = RTrim(JTM.履歴NO)
            .TextMatrix(ID, GridCol(15)) = JTM.前回部品単価
            .TextMatrix(ID, GridCol(16)) = JTM.前回梱包単価
            
'            If JTM.単価 <> 0 Then
'                If JTM.仮単価区分 = "" Then
'                    .TextMatrix(ID, GridCol(18)) = "正式"
'                Else
'                    .TextMatrix(ID, GridCol(18)) = "仮"
'                End If
'            End If
'            .TextMatrix(ID, GridCol(19)) = JTM.部品単価
'            .TextMatrix(ID, GridCol(20)) = JTM.梱包単価
        End With
        Call JTMReadNext(1)
    Loop

    With VsGrid
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    Call JTMClose(1)
    
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
        
        '===== 仮単価区分 =====
        If CisFun.Val2(iB1_Tan) <> 0 Then
            If B1Btn_Kari = "正" Then
                .TextMatrix(tRow, 10) = "正"
            Else
                .TextMatrix(tRow, 10) = "仮"
            End If
        Else
                .TextMatrix(tRow, 10) = ""
        End If
        
        '====== 部品単価 ======
        If CisFun.Val2(iB1_BTan) <> 0 Then
            .TextMatrix(tRow, 8) = Format(CisFun.Val2(iB1_BTan), "#,##0.00")
        Else
            .TextMatrix(tRow, 8) = ""
        End If
        
        Call TankaEdit
        
        '====== 梱包単価 ======
        If CisFun.Val2(iB1_KTan) <> 0 Then
            .TextMatrix(tRow, 9) = Format(CisFun.Val2(iB1_KTan), "#,##0.00")
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
        '仮区分
        If RTrim(.TextMatrix(tRow, 10)) <> "仮" Then
            B1Btn_Kari = "正"
        Else
            B1Btn_Kari = "仮"
        End If
        '部品単価
        If CisFun.Val2(.TextMatrix(tRow, 8)) <> 0 Then
            iB1_BTan = Format(CisFun.Val2(.TextMatrix(tRow, 8)), "#,##0.00")
        Else
            iB1_BTan = ""
        End If
        '梱包単価
        If CisFun.Val2(.TextMatrix(tRow, 9)) <> 0 Then
            iB1_KTan = Format(CisFun.Val2(.TextMatrix(tRow, 9)), "#,##0.00")
        Else
            iB1_KTan = ""
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
        iB1_Tan = "": iB1_KTan = "": iB1_BTan = ""
        B1Btn_Kari = "正"
        Call TextToVsGrid(Val(Input_Area.Tag))
    End If
    If VsGrid.Col = 4 And CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 6)) <> 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     " & VsGrid.Row & " 行目 前回単価を単価にコピーしますか？      "
            .MB_Title = ""
            .MB_Button = OK_CAN
            If Not .MBOX Then Exit Sub
        End With
        '=========== 単価コピー ==============
        iB1_Tan = VsGrid.TextMatrix(VsGrid.Row, 6)
        '=========== 部品コピー ==============
        If CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 15)) <> 0 Then
            iB1_KTan = VsGrid.TextMatrix(VsGrid.Row, 15)
        End If
        '=========== 梱包コピー ==============
        If CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 16)) <> 0 Then
            iB1_BTan = VsGrid.TextMatrix(VsGrid.Row, 16)
        End If
        '========== 仮区分コピー =============
        If VsGrid.TextMatrix(VsGrid.Row, 6) = "正" Then
            B1Btn_Kari = "正"
        Else
            B1Btn_Kari = "仮"
        End If
        Call TextToVsGrid(Val(Input_Area.Tag))
    End If
End Sub
Private Sub vsGrid_Scroll()
    'スクロール時はテキストボックスを非表示にします。
'    Call GridTextEnabled(False)
End Sub
'+---------------------------------------------+
'+      単価 = 部品単価 + 梱包単価 + その他費　計算
'+---------------------------------------------+
Private Sub TankaEdit()
    wEtcTan = CisFun.Val2(iB1_Tan) - CisFun.Val2(iB1_BTan)
    If wEtcTan >= 0 And CisFun.Val2(iB1_BTan) <> 0 Then
        iB1_KTan = Format(wEtcTan, "#,##0.00")
    Else
        iB1_BTan = ""
        iB1_KTan = ""
    End If
End Sub

'+---------------------------------------------+
'+      変更チェック
'+---------------------------------------------+
Private Sub UpdateCheck(sRow As Integer)
    VsGrid.Row = sRow
    If CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 14)) = 0 Then
        If CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 7)) <> 0 Then
            If RTrim(VsGrid.TextMatrix(VsGrid.Row, CisVsGrid.SelectCol)) = "" Then
                VsGrid.TextMatrix(VsGrid.Row, 11) = "追加"
                iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
                Call CisVsGrid.vsColor(個別, &HFFFFC0)
            End If
        Else
            If RTrim(VsGrid.TextMatrix(VsGrid.Row, CisVsGrid.SelectCol)) <> "" Then
                VsGrid.TextMatrix(VsGrid.Row, 11) = ""
                iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) - 1, "###")
                Call CisVsGrid.vsColor(消去個別)
            End If
        End If
    Else
        If CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 7)) <> CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 17)) Or _
           CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 8)) <> CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 19)) Or _
           CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 9)) <> CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 20)) Or _
           RTrim(VsGrid.TextMatrix(VsGrid.Row, 10)) <> RTrim(VsGrid.TextMatrix(VsGrid.Row, 18)) Then
           If RTrim(VsGrid.TextMatrix(VsGrid.Row, CisVsGrid.SelectCol)) = "" Then
               VsGrid.TextMatrix(VsGrid.Row, 11) = "修正"
                iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
               Call CisVsGrid.vsColor(個別, &HFFFFC0)
           End If
        Else
           If RTrim(VsGrid.TextMatrix(VsGrid.Row, CisVsGrid.SelectCol)) <> "" Then
               VsGrid.TextMatrix(VsGrid.Row, 11) = "*"
                iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) - 1, "###")
               Call CisVsGrid.vsColor(消去個別)
           End If
        End If
    End If
End Sub

