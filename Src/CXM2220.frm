VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CXM2220 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "受注工区振分メンテ"
   ClientHeight    =   11190
   ClientLeft      =   1845
   ClientTop       =   615
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D SelectColor 
      Height          =   285
      Left            =   4290
      Top             =   1020
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      BackColor       =   12648384
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
   End
   Begin VB.PictureBox B2_Area1 
      Height          =   645
      Left            =   2490
      ScaleHeight     =   585
      ScaleWidth      =   5445
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3120
      Width           =   5505
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   465
         Left            =   90
         Top             =   60
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   820
         ForeColor       =   16711680
         Caption         =   "現工区　　 　　　　　　新工区"
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
         cPositionX      =   120
         Begin Cis3D_v60.CIS3D B2lb_MKoku 
            Height          =   345
            Left            =   930
            Top             =   60
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            BackColor       =   16777152
            ForeColor       =   128
            Caption         =   "本社工場"
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   60
         End
         Begin VB.ComboBox B2Comb_SKoku 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3720
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   5
            Top             =   45
            Width           =   1440
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   285
            Left            =   2490
            Top             =   90
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
            ForeColor       =   16711680
            Caption         =   "⇒"
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
            cBoderStyle     =   2
         End
      End
   End
   Begin VB.Timer Timer2 
      Left            =   2190
      Top             =   8970
   End
   Begin VB.Timer Timer1 
      Left            =   1500
      Top             =   8970
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      Left            =   2100
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1680
      Width           =   555
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   8595
      ScaleHeight     =   675
      ScaleWidth      =   4320
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   945
      Width           =   4380
      Begin CisBtn_60.CisBtn PB_Chg 
         Height          =   585
         Left            =   60
         Top             =   30
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
         BtnCaption      =   "個別"
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
         PFCaption       =   "( F6 )"
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
      Begin CisBtn_60.CisBtn PB_LOOK 
         Height          =   585
         Left            =   900
         Top             =   30
         Visible         =   0   'False
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
         Left            =   2580
         Top             =   30
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
         Left            =   1740
         Top             =   30
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
         Left            =   3420
         Top             =   30
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
      Height          =   6045
      Left            =   2490
      ScaleHeight     =   5985
      ScaleWidth      =   10425
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3810
      Width           =   10485
      Begin vsFlexLib.vsFlexArray vsGrid1 
         Height          =   5835
         Left            =   90
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   60
         Width           =   10230
         _Version        =   196608
         _ExtentX        =   18045
         _ExtentY        =   10292
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
         FormatString    =   "№ |車型 |品　　　　　番　　　　　      |現工区 　    |新工区      |"
         Rows            =   16
         Cols            =   11
         BackColor       =   12648447
         BackColorBkg    =   14737632
         FocusRect       =   0
         HighLight       =   0
         RowHeightMin    =   360
         GridLines       =   2
         SelectionMode   =   1
         BackColorAlternate=   12648447
         Begin Cis3D_v60.CIS3D Input_Area 
            Height          =   345
            Left            =   8340
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
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
            Begin VB.ComboBox B1Comb_NKoku 
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   0
               Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
               TabIndex        =   7
               Top             =   0
               Width           =   1560
            End
            Begin VB.TextBox B_TabDummyB 
               Height          =   345
               Left            =   1560
               TabIndex        =   8
               Text            =   "Text1"
               Top             =   0
               Width           =   315
            End
            Begin VB.TextBox B_TabDummyF 
               Height          =   345
               Left            =   0
               TabIndex        =   6
               Text            =   "Text1"
               Top             =   0
               Width           =   255
            End
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   1335
      Left            =   2490
      ScaleHeight     =   1275
      ScaleWidth      =   10425
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1725
      Width           =   10485
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1125
         Left            =   1785
         Top             =   75
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1984
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Syakei 
            Height          =   345
            Left            =   120
            TabIndex        =   1
            Top             =   555
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   609
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
            Text            =   "XXXXX"
            MaxLength       =   5
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Left            =   90
         Top             =   75
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1984
         BackColor       =   12632319
         ForeColor       =   16711680
         Caption         =   "自社工区"
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
         cAlingnment     =   6
         cPositionY      =   60
         Begin VB.ComboBox H1Comb_Koku 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   0
            Top             =   555
            Width           =   1440
         End
      End
      Begin Cis3D_v60.CIS3D H1_TorMsg 
         Height          =   1125
         Left            =   7020
         Top             =   75
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "品       番"
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
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   345
            Left            =   60
            TabIndex        =   4
            Top             =   555
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   609
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
      Begin Cis3D_v60.CIS3D H1_UkeMsg 
         Height          =   1125
         Left            =   6375
         Top             =   75
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "受入"
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
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   345
            Left            =   150
            TabIndex        =   3
            Top             =   555
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   609
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XX"
            MaxLength       =   2
         End
      End
      Begin Cis3D_v60.CIS3D H1_NorMsg 
         Height          =   1125
         Left            =   2790
         Top             =   75
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "納 入 先"
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
            Height          =   345
            Left            =   60
            TabIndex        =   2
            Top             =   555
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   609
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
            Height          =   345
            Left            =   1035
            Top             =   555
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   609
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
      Caption         =   "【 受注工区振分メンテ  】"
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
         Left            =   11595
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
         Left            =   60
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   12583104
         Caption         =   "XXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
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
   Begin Cis3D_v60.CIS3D CIS3D10 
      Height          =   375
      Left            =   11415
      Top             =   3390
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   661
      Caption         =   "件 数"
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
      cPositionX      =   60
   End
   Begin Cis3D_v60.CIS3D B1lb_DispSu 
      Height          =   375
      Left            =   12135
      Top             =   3390
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   661
      Caption         =   "#,###"
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
      cAlingnment     =   4
      cBoderStyle     =   1
      cPositionX      =   -60
   End
   Begin Cis3D_v60.CIS3D CIS3D4 
      Height          =   375
      Left            =   9480
      Top             =   3390
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   661
      Caption         =   "変更件数"
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
      cPositionX      =   60
   End
   Begin Cis3D_v60.CIS3D B1lb_ChgKensu 
      Height          =   375
      Left            =   10560
      Top             =   3390
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   661
      Caption         =   "B1lb_ChgSu"
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
      cAlingnment     =   4
      cBoderStyle     =   1
      cPositionX      =   -60
   End
End
Attribute VB_Name = "CXM2220"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   受注工区振分メンテ
'**       フォームID    :   CXM2220
'**       処理概要      :
'**
'**       作  成  日    :   2008/04/03  By CIS
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    Dim mTanKb          As Byte
    
    Dim CisComb1        As CisCombo
    Dim CisComb2        As CisCombo
    Dim CisComb3        As CisCombo
    
    Dim GSelColor       As Long
    Dim GUnSelColor     As Long

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
    
    mGotColor = H1_Area1.BackColor
    mLostColor = B1_Area1.BackColor
    GSelColor = SelectColor.BackColor
    GUnSelColor = vsGrid1.BackColor
    
    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 800
    Timer2.Enabled = True
    
    '+------------------------+
    '+     ｺﾝﾎﾞﾎﾞｯｸｽのｾｯﾄ     +
    '+------------------------+
    Set CisComb1 = New CisCombo
    With CisComb1
        Set .Connect = CisDB
        Set .CombObj = H1Comb_Koku
        .Control = "名称マスタ"
        .ComboName = "自社工区"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With
    
    Set CisComb2 = New CisCombo
    With CisComb2
        Set .Connect = CisDB
        Set .CombObj = B1Comb_NKoku
        .Control = "名称マスタ"
        .ComboName = "自社工区"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With
    
    Set CisComb3 = New CisCombo
    With CisComb3
        Set .Connect = CisDB
        Set .CombObj = B2Comb_SKoku
        .Control = "名称マスタ"
        .ComboName = "自社工区"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With
'
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 360
        .PatanMax = 1
        .SelectCol = 10
        .InitGet = False
        .Init
    End With
    Call GridInit
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = "SGL"
    Call SyoriName(SyoriKB)
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
    Call CisComb1.ComboSet
          
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
           Case vbKeyF6
                               If PB_Chg.Visible Then
                                  If SyoriKB = "SGL" Then
                                     SyoriKB = "GRP"
                                  Else
                                     SyoriKB = "SGL"
                                  End If
                                  Call SyoriName(SyoriKB)
                               End If
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_Can_Click    '【取消】
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
Private Sub PB_Can_Click()
   Call CisFrm.MousePT(11)
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    切  替  キ ー  (F6)     +
'+----------------------------+
Private Sub PB_Chg_Click()
    If SyoriKB = "SGL" Then
       SyoriKB = "GRP"
    Else
       SyoriKB = "SGL"
    End If
    Call SyoriName(SyoriKB)
    H1Comb_Koku.SetFocus
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_UTorcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd: Exit Sub
End Sub
'+----------------------------+
'+       検　　　　 索         +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RVI_Hinmoku = 0
    RVI_SK = 9
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       iH1_Hinbn = RV_Hinbn
    End If

End Sub
'+-----------------------------+
'+       取引先受入検索         +
'+-----------------------------+
Private Sub Look_UTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 3
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
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
    If TorNmGet(iH1_Torcd, 3) Then
       H1lb_Tornm = TRM_RName
    End If
End Sub

'+-----------------------------+
'+  iH1_Ukeir
'+-----------------------------+
Private Sub iH1_Ukeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub

'+-----------------------------+
'+  B_TabDummyF
'+-----------------------------+
Private Sub B_TabDummyF_GotFocus()
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call Grid_RowChange(vbKeyUp)
    B1Comb_NKoku.SetFocus
End Sub
'+-----------------------------+
'+  B_TabDummyB
'+-----------------------------+
Private Sub B_TabDummyB_GotFocus()
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call Grid_RowChange(vbKeyDown)
    B1Comb_NKoku.SetFocus
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

'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim w_Color     As Long
   w_Color = &H80000008 '初期色ｾｯﾄ(黒)
   Select Case SyoriKB
     Case "SGL"
          SyoriNM = "個別"
          PB_Chg.BtnCaption = "一括"
          iH1_Hinbn.Enabled = True
          iH1_Torcd.Enabled = True
          iH1_Ukeir.Enabled = True
          B2_Area1.Visible = False
     Case "GRP"
          SyoriNM = "一括"
          PB_Chg.BtnCaption = "個別"
          iH1_Hinbn = ""
          iH1_Hinbn.Enabled = False
          iH1_Torcd.Enabled = False
          iH1_Ukeir.Enabled = False
          B2_Area1.Visible = True
     Case Else
          SyoriNM = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
       If gObj.Name Like "*vsGrid*" Then gObj.ForeColor = w_Color
   Next gObj
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
    B2_Area1.Enabled = B1Mode
    B2_Area1.BackColor = B1Color
    PB_Look.Visible = False
    PB_CAN.Visible = Not H1Mode
    PB_Chg.Visible = H1Mode
'    PB_END.Visible = H1Mode
    
    Call GridTextEnabled(B1Mode)
    
    '------------------------------ 2005/06/14  Insert
    If Rv_Call3 <> "" Then
       If Rv_Call3 = "X" Then
          PB_CAN.Visible = False
       End If
       PB_END.BtnCaption = "戻る"
    End If
    '------------------------------ 2005/06/14  Insert End
    
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
            If .Name Like wbc_Enm & "*Ch_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
                .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
       On Error Resume Next
       Call CisVsGrid.Clear
       On Error GoTo 0
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
' << 削除処理 >>
        ProcHB = "B1"
        Call DispChange(ProcHB)
        vsGrid1.Row = 1
        Call VsGridToText(vsGrid1.Row)
        Call CisComb2.ComboSet
        Call CisComb1.CodeGet
        CisComb2.Code = CisComb1.Code
        Call CisComb2.TextGet
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
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
    
' マスタ入力 SELECT文 ｾｯﾄ
    If Not DBInput("INV") Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "    該当するデータがありません。       "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
       End With
       H1Comb_Koku.SetFocus
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
    
    If SyoriKB = "SGL" Then
        If RTrim(B1lb_ChgKensu) = "" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "工区の変更が行われたデータがありません。"
               .MB_Title = ""
               .MBOX
               Call GridTextEnabled(True)
               Exit Function
           End With
        End If
    Else
       Call CisComb1.CodeGet
       Call CisComb3.CodeGet
       If CisComb1.Code = CisComb3.Code Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "変更する工区を選択してください。"
               .MB_Title = ""
               .MBOX
               Exit Function
          End With
       End If
    End If
    
    Body1Chk = True
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    
    Call HeadBodyClear("B")
    Call CisComb1.CodeGet
        
    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT HTM.品番, HNM.車型, HTM.取引先, HTM.受入, HTM.自社工区,"
    gSL_Select = gSL_Select & "        ISNULL(SY1.値名称, '') 自社工区名, HNM.表示品番,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(TM.略称, '') <> ''"
    gSL_Select = gSL_Select & "             THEN HTM.取引先 + ':' + TM.略称"
    gSL_Select = gSL_Select & "             ELSE HTM.取引先 END 取引先名"
    gSL_Select = gSL_Select & "   FROM 品番取引先マスタ HTM"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "             ON SY1.区分名称 = '自社工区'"
    gSL_Select = gSL_Select & "            AND SY1.値 = '" & CisComb1.Code & "'"
    gSL_Select = gSL_Select & "            AND SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番マスタ HNM"
    gSL_Select = gSL_Select & "             ON HNM.品番 = HTM.品番"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TM"
    gSL_Select = gSL_Select & "             ON TM.取引先CD = HTM.取引先"
    gSL_Select = gSL_Select & "            AND TM.取引先区分 = 0"
    
    With CisFun
         .Where_CreateKB = NewDocument
         Call CisComb1.CodeGet
         Call .Where_Create(AlphaMe, "HTM.自社工区", CisComb1.Code)
         Call .Where_Create(AlphaMe, "HNM.車型", iH1_Syakei)
         Call .Where_Create(AlphaMe, "HTM.品番", iH1_Hinbn)
         Call .Where_Create(AlphaMe, "HTM.取引先", iH1_Torcd)
         Call .Where_Create(AlphaMe, "HTM.受入", iH1_Ukeir)
         Call .Where_Create(AlphaMe, "品目", 0)
         gSL_Select = gSL_Select & .Where_Phrase
    End With
    
    gSL_Select = gSL_Select & "  ORDER BY HNM.車型, HTM.取引先, HTM.受入, HTM.品番"
    
    If Not HTMRead(gSL_Select, , 1) Then
        Call HTMClose(1)
        GoTo DBInput_Ed
    Else
        Call GridSet
        DBInput = True
    End If
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'+------------------+
'+    更新 処理     +
'+------------------+
Private Function DBPut()
    DBPut = False
    Dim wCnt  As Integer

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     更新処理を行います      "
        .MB_MSG(3) = "        よろしいですか？"
        .MB_Title = "更新処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then
           GridTextEnabled (True)
           Exit Function
        End If
    End With
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    On Error GoTo DBPut_Err
    
    If SyoriKB = "SGL" Then
        '+++++++++++++  個別更新  +++++++++++++++
        For wCnt = 1 To vsGrid1.Rows - 1
            If vsGrid1.TextMatrix(wCnt, CisVsGrid.SelectCol) <> "" Then
                Call CisComb2.CodeGet
            
                gSL_Select = ""
                gSL_Select = gSL_Select & " UPDATE 品番取引先マスタ SET"
                gSL_Select = gSL_Select & "  自社工区 = '" & RTrim(vsGrid1.TextMatrix(wCnt, 6)) & "'"
                gSL_Select = gSL_Select & "  WHERE 品番 = '" & RTrim(vsGrid1.TextMatrix(wCnt, 8)) & "'"
                gSL_Select = gSL_Select & "    AND 取引先 = '" & RTrim(vsGrid1.TextMatrix(wCnt, 7)) & "'"
                gSL_Select = gSL_Select & "    AND 受入 = '" & RTrim(vsGrid1.TextMatrix(wCnt, 3)) & "'"
                gSL_Select = gSL_Select & "    AND 品目 = 0"
                
                CisDB.SQL = gSL_Select
                
                Call CisDB.DBExec
            End If
        Next wCnt
    Else
        '+++++++++++++  一括更新  +++++++++++++++
        Call CisComb3.CodeGet
        gSL_Select = ""
        gSL_Select = gSL_Select & " UPDATE 品番取引先マスタ"
        gSL_Select = gSL_Select & "    SET 自社工区 = '" & CisComb3.Code & "'"
        gSL_Select = gSL_Select & "   FROM 品番取引先マスタ HTM"
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番マスタ HNM"
        gSL_Select = gSL_Select & "             ON HNM.品番 = HTM.品番"
        
        With CisFun
             .Where_CreateKB = NewDocument
             Call CisComb1.CodeGet
             Call .Where_Create(AlphaMe, "HNM.車型", iH1_Syakei)
             Call .Where_Create(AlphaMe, "品目", 0)
             Call .Where_Create(AlphaMe, "自社工区", CisComb1.Code)
             gSL_Select = gSL_Select & .Where_Phrase
        End With
        
        CisDB.SQL = gSL_Select
        
        Call CisDB.DBExec
    End If

    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      グリッド処理(vsGrid)                                 +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設
'定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 左中, あり, 6, 1, "車型")
        Call .FixedSet(中中, 左中, あり, 25, 2, "取 引 先")
        Call .FixedSet(中中, 左中, あり, 2, 3, "受")
        Call .FixedSet(中中, 左中, あり, 30, 4, "品     番")
        Call .FixedSet(中中, 左中, あり, 12, 5, "新工区")
        
        Call .FixedSet(中中, 左中, あり, 0, 6, "工区新")
        Call .FixedSet(中中, 左中, あり, 0, 7, "取引先")
        Call .FixedSet(中中, 左中, あり, 0, 8, "品番")
        Call .FixedSet(中中, 左中, あり, 0, 9, "変更前自社工区")
        
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 10, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        'Call .PatanSet(0, "0,1,2,3,4,5,6")
        Call .PatanSet(0, "0,1,2,3,4,5")
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
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    B1lb_DispSu = CisDB.RecordCount       ' 抽出件数(合計)退避
    If SyoriKB = "GRP" Then
       B1lb_ChgKensu = CisDB.RecordCount
    End If
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(CisFun.Val2(ID))
        Next ID
        If CisFun.Val2(B1lb_DispSu) + 1 > .Rows Then
            .Rows = CisFun.Val2(B1lb_DispSu) + 1
        End If
    End With
    ID = 0
    Do Until Not HTM_RDSTS
        ID = ID + 1
        With vsGrid1
            .TextMatrix(ID, GridCol(0)) = ID
            .TextMatrix(ID, GridCol(1)) = RTrim(HTM.車型)
            .TextMatrix(ID, GridCol(2)) = RTrim(HTM.取引先名)
            .TextMatrix(ID, GridCol(3)) = RTrim(HTM.受入)
            .TextMatrix(ID, GridCol(4)) = RTrim(HTM.表示品番)
            .TextMatrix(ID, GridCol(5)) = RTrim(HTM.自社工区名)
            .TextMatrix(ID, GridCol(6)) = RTrim(HTM.自社工区)
            .TextMatrix(ID, GridCol(7)) = RTrim(HTM.取引先)
            .TextMatrix(ID, GridCol(8)) = RTrim(HTM.品番)
            .TextMatrix(ID, GridCol(9)) = RTrim(HTM.自社工区)
        End With
        Call HTMReadNext(1)
    Loop

    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    
    If SyoriKB = "GRP" Then
       B2lb_MKoku = H1Comb_Koku
       Call CisComb3.ComboSet
    End If
    
    Call HTMClose(1)
    
End Sub
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    If SyoriKB = "SGL" Then
       Input_Area.Visible = EnabledFlg
    End If
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'テキストからグリッドに退避
    Call TextToVsGrid(Val(Input_Area.Tag))
    gInt = vsGrid1.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > vsGrid1.Rows - 1 Then
            If RTrim(vsGrid1.TextMatrix(gInt, 2)) <> "" Then
                vsGrid1.Rows = gInt + 1 + 1
                vsGrid1.TextMatrix(gInt + 1, 0) = gInt + 1
            Else
                GoTo Grid_RowChange_End
            End If
        End If
        If vsGrid1.TextMatrix(gInt + 1, 2) = "" Then Exit Sub
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
        If gInt + CisVsGrid.DispMax > vsGrid1.Rows - 1 Then
            gInt = vsGrid1.Rows - 1
        Else
            gInt = gInt + CisVsGrid.DispMax
        End If
    End If
    vsGrid1.Row = gInt
'グリッドからテキストに表示
    Call VsGridToText(vsGrid1.Row)
    'Call CisComb2.ComboSet
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With vsGrid1
        Call CisComb2.CodeGet
        .TextMatrix(tRow, 6) = CisComb2.Code
        .TextMatrix(tRow, 5) = RTrim(B1Comb_NKoku)
        
        If .TextMatrix(tRow, 6) <> .TextMatrix(tRow, 9) Then
           For gInt = 1 To 6
               .Col = gInt
               .CellBackColor = GSelColor
           Next gInt
           If .TextMatrix(tRow, CisVsGrid.SelectCol) = "" Then
              .TextMatrix(tRow, CisVsGrid.SelectCol) = "X"
              B1lb_ChgKensu = CisFun.Val2(B1lb_ChgKensu) + 1
           End If
        Else
           For gInt = 1 To 6
               .Col = gInt
               .CellBackColor = vsGrid1.BackColor
           Next gInt
           If .TextMatrix(tRow, CisVsGrid.SelectCol) <> "" Then
              .TextMatrix(tRow, CisVsGrid.SelectCol) = ""
              B1lb_ChgKensu = CisFun.Val2(B1lb_ChgKensu) - 1
           End If
        End If
        
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With vsGrid1
        Input_Area.Tag = tRow
        Input_Area.Move Input_Area.Left, vsGrid1.CellTop, Input_Area.Width, vsGrid1.CellHeight
        '発注数
        
        CisComb2.Code = .TextMatrix(tRow, 6)
        Call CisComb2.TextGet
        
    End With
End Sub
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call VsGridToText(vsGrid1.Row)
End Sub
Private Sub VSGrid1_Click()
    Call TextToVsGrid(Val(Input_Area.Tag))
    If vsGrid1.TextMatrix(vsGrid1.Row, 0) = "" Then Exit Sub
    Call GridTextEnabled(True)
    Call VsGridToText(vsGrid1.Row)
    On Error Resume Next
    B1Comb_NKoku.SetFocus
    On Error GoTo 0
End Sub

