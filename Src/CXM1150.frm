VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "Cis3D_6.0.ocx"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CisBtn_6.0.ocx"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CisText_6.0.ocx"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CXM1150 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "構成マスタ照会"
   ClientHeight    =   11235
   ClientLeft      =   1905
   ClientTop       =   1800
   ClientWidth     =   15360
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D Proc_MSG 
      Height          =   1410
      Left            =   4305
      Top             =   4260
      Visible         =   0   'False
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   2487
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
      cBoderColor1    =   -2147483643
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   1155
         Left            =   135
         Top             =   120
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   2037
         BackColor       =   14737632
         ForeColor       =   0
         Caption         =   "【処理中】しばらくお待ち下さい"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   20.25
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
         cBoderWidth     =   5
         cAlingnment     =   7
         cBoderStyle     =   1
      End
   End
   Begin VB.PictureBox B1_Area2 
      Height          =   2280
      Left            =   1410
      ScaleHeight     =   2220
      ScaleWidth      =   12525
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   8550
      Width           =   12585
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   2055
         Left            =   90
         Top             =   90
         Width           =   12345
         _ExtentX        =   21775
         _ExtentY        =   3625
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
         Begin vsFlexLib.vsFlexArray VsGridT 
            Height          =   2025
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   12375
            _Version        =   196608
            _ExtentX        =   21828
            _ExtentY        =   3572
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
            ConvInfo        =   1413783674
            FormatString    =   "№ |取引先               |受|背番号|手配|比率|SK|ｻｲｸﾙ   |有効期限               |係1 |係2 |係3 |LT|ロッ ト|月 火 水 木 金|"
            Rows            =   5
            Cols            =   16
            BackColor       =   16777152
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   330
            SelectionMode   =   1
            AllowBigSelection=   0   'False
            BackColorAlternate=   16777152
            AllowSelection  =   0   'False
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
      Left            =   9645
      ScaleHeight     =   675
      ScaleWidth      =   4275
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   990
      Width           =   4335
      Begin CisBtn_60.CisBtn PB_GetB 
         Height          =   585
         Left            =   881
         Top             =   585
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
         BtnCaption      =   "戻す"
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
         PFBackColor     =   16711680
      End
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1710
         Top             =   45
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
         Left            =   2550
         Top             =   45
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
         Left            =   3390
         Top             =   45
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   45
         Top             =   45
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
      Begin CisBtn_60.CisBtn PB_PRT 
         Height          =   585
         Left            =   881
         Top             =   45
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
      Begin CisBtn_60.CisBtn PB_CSV 
         Height          =   585
         Left            =   1717
         Top             =   45
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
      Height          =   975
      Left            =   1410
      ScaleHeight     =   915
      ScaleWidth      =   6045
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   750
      Width           =   6110
      Begin Cis3D_v60.CIS3D HLB_Hinbn 
         Height          =   795
         Left            =   2705
         Top             =   60
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "品   番"
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
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   330
            Left            =   90
            TabIndex        =   0
            Top             =   360
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   582
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
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   795
         Left            =   75
         Top             =   60
         Width           =   1700
         _ExtentX        =   2990
         _ExtentY        =   1402
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin VB.OptionButton H1Op_Tenkb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "正"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   390
            Value           =   -1  'True
            Width           =   585
         End
         Begin VB.OptionButton H1Op_Tenkb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "逆"
            Height          =   315
            Index           =   1
            Left            =   920
            TabIndex        =   10
            Top             =   390
            Width           =   585
         End
         Begin Cis3D_v60.CIS3D HLB_Tenkb 
            Height          =   345
            Left            =   1380
            Top             =   360
            Visible         =   0   'False
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   609
            BackColor       =   14737632
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
            Begin VB.OptionButton H1Op_Oyakb 
               BackColor       =   &H00E0E0E0&
               Caption         =   "直前"
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   9
               Top             =   60
               Width           =   795
            End
            Begin VB.OptionButton H1Op_Oyakb 
               BackColor       =   &H00E0E0E0&
               Caption         =   "製品"
               Enabled         =   0   'False
               Height          =   255
               Index           =   1
               Left            =   1020
               TabIndex        =   8
               Top             =   60
               Width           =   795
            End
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   795
         Left            =   1775
         Top             =   60
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1402
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
         Begin CisText_V60.CisText iH1_Syakei 
            Height          =   330
            Left            =   90
            TabIndex        =   12
            Top             =   360
            Width           =   750
            _ExtentX        =   1323
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
            Text            =   "XXXXX"
            MaxLength       =   5
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   795
         Left            =   75
         Top             =   60
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1402
         ForeColor       =   16711680
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin VB.OptionButton H1Op_Hinkb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "材料"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   180
            TabIndex        =   14
            Top             =   420
            Width           =   855
         End
         Begin VB.OptionButton H1Op_Hinkb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "品番"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   13
            Top             =   75
            Value           =   -1  'True
            Width           =   855
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
      Left            =   1620
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1260
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   1800
      Top             =   1260
   End
   Begin VB.Timer Timer2 
      Left            =   1830
      Top             =   1200
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
      Caption         =   "【 構成マスタ照会 】"
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
         Left            =   11730
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
         Left            =   8655
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
            Top             =   15
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
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   6315
      Left            =   1410
      ScaleHeight     =   6255
      ScaleWidth      =   12510
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2130
      Width           =   12570
      Begin vsFlexLib.vsFlexArray VsGrid1 
         Height          =   6105
         Left            =   90
         TabIndex        =   1
         Top             =   75
         Width           =   12330
         _Version        =   196608
         _ExtentX        =   21749
         _ExtentY        =   10769
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
         ConvInfo        =   1413783674
         FormatString    =   "No |レベル|区|XXXXXXXXX1XXXXXXXXX2XXXXX|XXXXXXXXX1XXXXXXXXX2|使用数|展開区分  |9999/99/99-9999/99/99|寸法|"
         Rows            =   20
         Cols            =   13
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   240
         SelectionMode   =   1
         AllowBigSelection=   0   'False
         BackColorAlternate=   12648447
         AllowSelection  =   0   'False
      End
   End
   Begin VB.Timer Timer3 
      Left            =   2400
      Top             =   1230
   End
   Begin Cis3D_v60.CIS3D H1lb_Hinbn 
      Height          =   360
      Left            =   1410
      Top             =   1760
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   635
      BackColor       =   16777152
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
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
   Begin Cis3D_v60.CIS3D H1lb_Syake 
      Height          =   360
      Left            =   4500
      Top             =   1755
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   635
      BackColor       =   16777152
      Caption         =   "XXXXXXXXX1"
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
   Begin Cis3D_v60.CIS3D lb_JyutyuBiko 
      Height          =   645
      Left            =   120
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1138
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
   Begin VB.Menu mn_Menu 
      Caption         =   "左クリック"
      NegotiatePosition=   1  '左
      Visible         =   0   'False
      Begin VB.Menu mn_Copy 
         Caption         =   "品番/材質をクリップボードにコピー"
      End
      Begin VB.Menu mn_Dmy1 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Tenk 
         Caption         =   "再表示( 正 )"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_Gyaku 
         Caption         =   "再表示( 逆・直前 )"
         Index           =   0
      End
      Begin VB.Menu mn_Gyaku 
         Caption         =   "再表示( 逆・製品 )"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mn_Dmy2 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Can 
         Caption         =   "キャンセル"
      End
   End
End
Attribute VB_Name = "CXM1150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   構成マスタメンテナンス
'**       フォームID    :   CXM1150 (Based CXM0150)
'**       処理概要      :
'**
'**       作  成  日    :   2019/07/18  By CIS SIMPR-A
'**       変  更  日    :
'**
'**********************************************************'
    
    Const i_Dragico     As String = "BMP\DRAG1PG.ICO"
    
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    
    Dim mCHK            As Boolean

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
    Dim CisVsGrid1      As New CisVsGrid3
    Dim CisVsGridT      As New CisVsGrid3
    Dim ID              As Long
    
    Dim ActObj          As Object
    Dim ActRow          As Integer
    
    Dim KCol            As Long         ' 非表示項目開始ｶﾗﾑ
    
    Dim SCol()          As Long         ' 表示項目ｶﾗﾑ(品番)
    Dim SColT()         As Long         ' 表示項目ｶﾗﾑ(材料)

    Dim wsHKbn          As Long
    Dim RecCnt          As Long

    Dim RowColKB        As Byte
    
    Dim w_ZaiKB         As Byte
    Dim GetB_Cnt        As Byte
    Dim sv_Hinkb()      As Byte
    Dim sv_Tenkb()      As Byte
    Dim sv_Oyakb()      As Byte
    Dim sv_Hinbn()      As String
    Dim sv_TpRow()      As Long
    Dim sv_FcRow()      As Long
    Dim sv_Zaigy()      As String
    
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
    gConnectCount = 4                   ' 2019/07/26 ADD
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
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔(基本)
    Timer2.Enabled = False   ' 処理区分点滅禁止(基本)
    
    Proc_MSG.Visible = False
    Proc_MSG.Move 4305, 4260
'
    '+--------------------------------+
    '+  グリッド初期設定(構成)
    '+--------------------------------+
    With CisVsGrid1
        Set .GridObj = VsGrid1
        .NoSpaceCol = 3
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 4
        .SelectCol = 10
        .InitGet = False
        .Init
    End With
    wsHKbn = 0
    Call GridInit1

'   [ Drag Icon Set ]
    On Error Resume Next
    Set VsGrid1.DragIcon = LoadPicture(gIniExe & i_Dragico)
    On Error GoTo 0

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

    H1Op_Hinkb(0).Value = True
    H1Op_Tenkb(0).Value = True
    GetB_Cnt = 0
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
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF8:       If PB_PRT.Visible Then Call PB_Prt_Click    '【印刷】    ' 2019/07/30 ADD
           Case vbKeyF9:       If PB_GetB.Visible Then Call PB_GetB_Click  '【戻す】
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
        
        GetB_Cnt = 0
        Erase sv_Hinkb
        Erase sv_Tenkb
        Erase sv_Oyakb
        Erase sv_Hinbn
        Erase sv_TpRow
        Erase sv_FcRow
        Erase sv_Zaigy
        
        ProcHB = "H1"
        H1lb_Hinbn.Caption = ""         ' 2019/07/22 ADD
        H1lb_Syake.Caption = ""         ' 2019/07/22 ADD
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        GoTo PB_CAN_Ed
    End If

PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
End Sub
'+----------------------------+
'+    戻　す  キ ー(F9)       +
'+----------------------------+
Private Sub PB_GetB_Click()
    Call GetBack_vsGrid1
End Sub

'===============================================================================' 2019/07/25 ADD START
'+----------------------------+
'+    印　刷　キ　ー(F8)      +
'+----------------------------+
Private Sub PB_Prt_Click()
    Dim PrtCnt      As Long
    Dim PrtOut      As Integer

    Dim GyakKB      As Byte
    Dim HZKB        As Byte
    Dim TSKB        As Byte
    
    
    bStaPara2 = CisFun.Val2(H1Op_Hinkb(0).Tag)      ' 2019/07/26 ADD
    bStaPara3 = CisFun.Val2(H1Op_Tenkb(0).Tag)      ' 2019/07/26 ADD
    bStaPara4 = CisFun.Val2(H1Op_Oyakb(0).Tag)      ' 2019/07/26 ADD
    wPrtHKbn = wsHKbn                               ' 2019/07/26 ADD
    
    With CisSimplePrint
         .PrintTitle = "構成マスタ照会一覧"
         .SaveKey = gRegKey
       ''.Preservation = "CXM00150"                 ' 2019/07/26 UPD
         .Preservation = "CXM1150"                  ' 2019/07/26 UPD
         
         .GetCsvFileName = mCSVDir & "11111.csv"
         
         
         If Not .CheckPrtCsv Then Exit Sub
    
         Select Case .PrintStatus
            Case 4: PrtOut = 0
            Case 5: PrtOut = 1
            Case 3: PrtOut = 2
            Case Else: GoTo PB_Prt_Ed
         End Select
         
    End With
    
    'If Not CXM1150LPrint(PrtOut) Then Exit Sub       2009/11/10
    '================================================ 2009/11/10 Start
    If H1Op_Tenkb(0) Then GyakKB = 0 Else GyakKB = 1
    If H1Op_Oyakb(0) Then TSKB = 0 Else TSKB = 1
    If H1Op_Hinkb(0) Then HZKB = 0 Else HZKB = 1
    If Not CXM1150LPrint(PrtOut, GyakKB, HZKB, TSKB) Then Exit Sub
    '================================================ 2009/11/10 End
    
    Call CisSimplePrint.CheckFormUnload
PB_Prt_Ed:
End Sub
'===============================================================================' 2019/07/25 ADD END

'+----------------------------+
'+       品番検索              +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Hinbn = ""
    RV_Zaist = ""
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"

'    If CisFun.Val2(H1Op_Hinkb(0).Tag) = 0 Then     ' 2019/07/18 DEL
        CKK0030.Show vbModal
        Unload CKK0030
        Set CKK0030 = Nothing
        If RV_Rtn Then
            iH1_Hinbn = RV_Hinbn
        End If
'===================================================' 2019/07/18 DEL START
'    Else
'        CKK0070.Show vbModal
'        Unload CKK0070
'        Set CKK0070 = Nothing
'        If RV_Rtn Then
'            iH1_Hinbn = RV_Zaist
'        End If
'    End If
'===================================================' 2019/07/18 DEL END
End Sub
'+---( H1 )----------------------------------------------------------------------------
'+--------------------------------------------------+
'+  H1Op_Hinkb  (品番/材料)
'+--------------------------------------------------+
Private Sub H1Op_Hinkb_Click(INDEX As Integer)

    INDEX = 0                                       ' 2019/07/18 ADD：品番のみとする
    
    H1Op_Hinkb(0).Tag = INDEX
    
'    If INDEX = 0 Then                              ' 2019/07/18 DEL
        H1Op_Tenkb(0).Enabled = True
        HLB_Hinbn.Caption = "品   番"
'    Else                                           ' 2019/07/18 DEL
'        H1Op_Tenkb(0).Enabled = False              ' 2019/07/18 DEL
'        H1Op_Tenkb(1).Value = True                 ' 2019/07/18 DEL
'        HLB_Hinbn.Caption = "材   質"              ' 2019/07/18 DEL
'    End If                                         ' 2019/07/18 DEL
End Sub
Private Sub H1Op_Hinkb_GotFocus(INDEX As Integer)
    H1Op_Hinkb(INDEX).BackColor = &HC0FFFF
End Sub
Private Sub H1Op_Hinkb_LostFocus(INDEX As Integer)
    H1Op_Hinkb(INDEX).BackColor = &HC0C0C0
End Sub
'+--------------------------------------------------+
'+  H1Op_Tenkb  (正/逆展開)
'+--------------------------------------------------+
Private Sub H1Op_Tenkb_Click(INDEX As Integer)
    H1Op_Tenkb(0).Tag = INDEX
    If INDEX = 0 Then
        H1Op_Oyakb(0).Value = False
        H1Op_Oyakb(1).Value = False
        HLB_Tenkb.Enabled = False
    Else
        HLB_Tenkb.Enabled = True
        H1Op_Oyakb(0).Value = True
    End If
End Sub
Private Sub H1Op_Tenkb_GotFocus(INDEX As Integer)
    H1Op_Tenkb(INDEX).BackColor = &HC0FFFF
End Sub
Private Sub H1Op_Tenkb_LostFocus(INDEX As Integer)
    H1Op_Tenkb(INDEX).BackColor = &HC0C0C0
End Sub
'+--------------------------------------------------+
'+  H1Op_Oyakb  (直前/製品)
'+--------------------------------------------------+
Private Sub H1Op_Oyakb_Click(INDEX As Integer)
    INDEX = 1                   ' 2019/07/18 ADD：製品のみとする
    H1Op_Oyakb(1).Value = True  ' 2019/07/18 ADD：製品のみとする
    H1Op_Oyakb(0).Tag = INDEX
End Sub
Private Sub H1Op_Oyakb_GotFocus(INDEX As Integer)
    H1Op_Oyakb(INDEX).BackColor = &HC0FFFF
End Sub
Private Sub H1Op_Oyakb_LostFocus(INDEX As Integer)
    H1Op_Oyakb(INDEX).BackColor = &HE0E0E0
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
    B1_Area2.Enabled = B1Mode
    B1_Area2.BackColor = B1Color
    PB_CAN.Visible = Not H1Mode
    PB_PRT.Visible = Not H1Mode             ' 2019/07/23 ADD
    PB_CSV.Visible = Not H1Mode             ' 2019/07/23 ADD
    PB_ENT.Visible = H1Mode                 ' 2019/07/22 ADD
    
    PB_GetB.Visible = (GetB_Cnt <> 0)
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
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    
    If HeadBody = "B" Then
        VsGrid1.Clear
        VsGridT.Clear
        wsHKbn = 0
        Call GridInit1
        Call GridInitT
        lb_Kensu = ""
    End If
    
    iH1_Hinbn.SetFocus          ' 2019/07/18 ADD
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
    
' 品番
    If RTrim(iH1_Hinbn) = "" Then
        With CisFun
            .MB_Lines = 4
'            If CisFun.Val2(H1Op_Hinkb(0).Tag) = 0 Then                 ’2019.07.22 DEL
                .MB_MSG(2) = "    品番を入力して下さい            "
'            Else                                                       ’2019.07.22 DEL
'                .MB_MSG(2) = "    材質を入力して下さい            "    ’2019.07.22 DEL
'            End If                                                     ’2019.07.22 DEL
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Hinbn.SetFocus
       Exit Function
    End If

' 表示情報
    wsHKbn = 0
    If CisFun.Val2(H1Op_Hinkb(0).Tag) <> 0 Then
        wsHKbn = 2
    Else
        If CisFun.Val2(H1Op_Tenkb(0).Tag) <> 0 Then
            wsHKbn = 1
        End If
    End If
    
' グリッドセット
    Proc_MSG.Visible = True: DoEvents
    If Not GridSet_KOM Then GoTo Head1Chk_ED
    
    RowColKB = 0
    Head1Chk = True
Head1Chk_ED:
    Proc_MSG.Visible = False: DoEvents
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
        ReDim SCol(11)

        Call .FixedSet(中中, 右中, なし, 4, 0, "No")
        Call .FixedSet(中中, 左中, あり, 7, 1, "レベル")
        Call .FixedSet(中中, 左中, あり, 2, 2, "区")
        Call .FixedSet(中中, 左中, あり, 25, 3, "品  番")
        Call .FixedSet(中中, 左中, あり, 20, 4, "品  名")
        Call .FixedSet(中中, 右中, あり, 10, 5, "使用数")
        Call .FixedSet(中中, 左中, あり, 10, 6, "展開区分")
        Call .FixedSet(中中, 左中, あり, 21, 7, "適用範囲")
        Call .FixedSet(中中, 左中, あり, 25, 8, "寸  法")
        Call .FixedSet(中中, 左中, あり, 0, 9, "編集前品番")
        Call .FixedSet(中中, 左中, あり, 0, 11, "検索元品番")       ' 2019/07/18 ADD
        Call .FixedSet(中中, 左中, あり, 0, 12, "検索元車型")       ' 2019/07/22 ADD
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, あり, 0, 10, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,8,4,5,6,7")
        Call .PatanSet(1, "0,3,4")
        Call .PatanSet(2, "0,8,3,4")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
        Call CisVsGrid1.LayOut(wsHKbn)
        
        SCol(1) = .FixedGet(1)     ' レベル
        SCol(2) = .FixedGet(2)     ' 区
        SCol(3) = .FixedGet(3)     ' 品番
        SCol(4) = .FixedGet(4)     ' 品名
        SCol(5) = .FixedGet(5)     ' 使用数
        SCol(6) = .FixedGet(6)     ' 展開区分
        SCol(7) = .FixedGet(7)     ' 適用範囲
        SCol(8) = .FixedGet(8)     ' 寸法
        SCol(9) = .FixedGet(9)     ' 編集前品番
    End With
    KCol = 9   ' 非表示項目開始ｶﾗﾑ
End Sub
'+---------------------------------------------------------------------------------------+
'+      グリッド内容セット
'+---------------------------------------------------------------------------------------+
Private Function GridSet_KOM()
    GridSet_KOM = False
    
'===========================================================================================' 2019/07/18 ADD START
    Dim iCnt        As Integer
    
    Dim iLineCnt    As Long
    Dim iHnbnCnt    As Long
    
    Dim wWhere As String
    wWhere = ""
''＜＜　検索条件作成　＞＞
    ' 品番
    iH1_Hinbn.Text = Trim(iH1_Hinbn.Text)
    If InStr(1, iH1_Hinbn.Text, "*") = 0 Then
        wWhere = " WHERE HM.品番 = '" & iH1_Hinbn.Text & "'"
    Else
        wWhere = " WHERE HM.品番 LIKE '" & Replace(iH1_Hinbn.Text, "*", "%") & "'"
    End If
   ' 車型
    iH1_Syakei.Text = Trim(iH1_Syakei.Text)
    If iH1_Syakei.Text <> "" Then
        wWhere = wWhere & vbCrLf & "   AND HM.車型 = '" & iH1_Syakei.Text & "'"
    End If
    If wsHKbn = 0 Then
        ' 品目
        wWhere = wWhere & vbCrLf & "   AND HT.品目 = 0"
    End If
    Call Str_Monitor(wWhere)
    
    Dim wSQL As String
    With CisDB
        wSQL = ""
        wSQL = wSQL & vbCrLf & "SELECT HM.品番, HM.車型"
        wSQL = wSQL & vbCrLf & "  FROM 品番マスタ HM"
        wSQL = wSQL & vbCrLf & "  LEFT OUTER JOIN 品番取引先マスタ HT"
        wSQL = wSQL & vbCrLf & "    ON HT.品番 = HM.品番"
        wSQL = wSQL & vbCrLf & "   AND HT.品目 = 0"
        wSQL = wSQL & vbCrLf & wWhere
        wSQL = wSQL & vbCrLf & " GROUP BY HM.品番, HM.車型"
        Call Str_Monitor(wSQL)
    End With

    If Not HTMRead(wSQL) Then
      ''MsgBox ("データが無いよ")
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     該当する品番データがありません。     "
            .MB_Title = "品番検索"
            .MB_Button = Error
            Call .MBOX
        End With
        Call iH1_Hinbn.SetFocus
        Exit Function
    End If
    
  ''MsgBox (CisDB.RecordCount)
    iRdCnt = CisDB.RecordCount
    
    ReDim wHinbn(iRdCnt)
    ReDim wSyake(iRdCnt)
    For iCnt = 1 To iRdCnt
        wHinbn(iCnt) = HTM.品番
        wSyake(iCnt) = HTM.車型
        Call HTMReadNext
    Next
    
    ReDim wRdKOL(15, iRdCnt)
    iLineCnt = 0
'===========================================================================================' 2019/07/18 ADD END
    
    Call GridInit1
        
    For iCnt = 1 To iRdCnt      ' 2019/07/19 ADD
        With CisDB
            .SQL = "構成問合せ情報抽出"
            .StoadoCount = 4
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = Input用
            .ParaNo = 3: .ParaIO = Input用
            .ParaNo = 4: .ParaIO = Input用
        
          ''.ParaNo = 1: .ParaValue = RTrim(iH1_Hinbn)                  ' 2019/07/19 DEL
            .ParaNo = 1: .ParaValue = wHinbn(iCnt)                      ' 2019/07/19 ADD
            .ParaNo = 2: .ParaValue = CisFun.Val2(H1Op_Hinkb(0).Tag)
            .ParaNo = 3: .ParaValue = CisFun.Val2(H1Op_Tenkb(0).Tag)
            .ParaNo = 4: .ParaValue = CisFun.Val2(H1Op_Oyakb(0).Tag)
        
            If Not .DBStored Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "     構成マスタ情報の抽出に失敗しました。     "
                    .MB_Title = "抽出処理"
                    .MB_Button = Error
                    If .MBOX Then Exit Function
                End With
            End If
        End With
    
        CisVsGrid1.Clear
        
        If wsHKbn = 0 Then
            gSL_Select = "SELECT KW.* FROM 構成問合せワーク KW "
            gSL_Select = gSL_Select & " WHERE KW.処理端末 = HOST_NAME() "
            gSL_Select = gSL_Select & " ORDER BY KW.順位,KW.子品番 "
        End If
        If wsHKbn = 1 Then
            gSL_Select = "SELECT KW.親品番,KW.品名,KW.表示品番 FROM 構成問合せワーク KW "
            gSL_Select = gSL_Select & " WHERE KW.処理端末 = HOST_NAME() "
            gSL_Select = gSL_Select & " GROUP BY KW.親品番,KW.品名,KW.表示品番"
            gSL_Select = gSL_Select & " ORDER BY KW.親品番"
        End If
        If wsHKbn = 2 Then
            gSL_Select = "SELECT KW.親品番,KW.品名,KW.表示寸法,KW.表示品番 FROM 構成問合せワーク KW "
            gSL_Select = gSL_Select & " WHERE KW.処理端末 = HOST_NAME() "
            gSL_Select = gSL_Select & " GROUP BY KW.親品番,KW.品名,KW.表示寸法,KW.表示品番"
            gSL_Select = gSL_Select & " ORDER BY KW.表示寸法,KW.親品番"
        End If
    '===========================================' 2019/07/19 DEL START
    '    If Not KOLRead(gSL_Select) Then
    '        With CisFun
    '            .MB_Lines = 4
    '            .MB_MSG(2) = "     構成マスタ情報が存在しません。     "
    '            .MB_Title = "抽出処理"
    '            .MB_Button = Error
    '            If .MBOX Then Exit Function
    '        End With
    '    End If
    '===========================================' 2019/07/19 DEL END
    '===========================================' 2019/07/19 ADD START
        If Not KOLRead(gSL_Select) Then
            GoTo SKIP0001
        End If
    '===========================================' 2019/07/19 ADD END
        
    ' 抽出処理 初期ｾｯﾄ
        RecCnt = CisDB.RecordCount       ' 抽出件数(合計)退避
    
        ' 配列サイズ変更
        ReDim Preserve wRdKOL(15, iLineCnt + CisDB.RecordCount)
        Do While KOL_RDSTS
            iLineCnt = iLineCnt + 1
            With KOL
                wRdKOL(0, iLineCnt) = wHinbn(iCnt)      ' 検索元品番
                wRdKOL(1, iLineCnt) = .順位
                wRdKOL(2, iLineCnt) = .レベル
                wRdKOL(3, iLineCnt) = .親品番
                wRdKOL(4, iLineCnt) = .子品番
                wRdKOL(5, iLineCnt) = .構成数
                wRdKOL(6, iLineCnt) = .展開区分
                wRdKOL(7, iLineCnt) = .LT
                wRdKOL(8, iLineCnt) = .適用開始日
                wRdKOL(9, iLineCnt) = .適用終了日
                wRdKOL(10, iLineCnt) = .区分
                wRdKOL(11, iLineCnt) = .表示品番
                wRdKOL(12, iLineCnt) = .表示寸法
                wRdKOL(13, iLineCnt) = .品名
                wRdKOL(14, iLineCnt) = .材料管理番号
                wRdKOL(15, iLineCnt) = wSyake(iCnt)     ' 検索元車型
            End With
            Call KOLReadNext
        Loop
    
SKIP0001:                                   ' 2019/07/19 ADD
    Next                                    ' 2019/07/19 ADD
    
    Call KOLClose                           ' 2019/07/19 ADD
    RecCnt = iLineCnt                       ' 2019/07/19 ADD


'===========================================' 2019/07/19 ADD START
    If iLineCnt = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     構成マスタ情報が存在しません。     "
            .MB_Title = "抽出処理"
            .MB_Button = Error
            If .MBOX Then Exit Function
        End With
    End If

    Call GridInit1
'===========================================' 2019/07/19 ADD END



' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid1
       .Redraw = False
        If RecCnt + 1 < CisVsGrid1.DispMax Then
            .Rows = CisVsGrid1.DispMax
        Else
            .Rows = RecCnt + 1
        End If
    End With
    ID = 0


'    Do Until Not KOL_RDSTS
'        ID = ID + 1
'
'            If wsHKbn = 0 Then: Call KOMInfoToGrid0(ID)     ' 正展開
'            If wsHKbn = 1 Then: Call KOMInfoToGrid1(ID)     ' 逆展開・品番
'            If wsHKbn = 2 Then: Call KOMInfoToGrid2(ID)     ' 逆展開・材料
'
'            Call KOLReadNext
'        Loop
'
'    Call KOLClose
''''    ↑↑↑↑    差し替え    ↓↓↓↓                    ' 2019/07/19 UPD
    For ID = 1 To iLineCnt
        If wsHKbn = 0 Then: Call KOMInfoToGrid0(ID)         ' 正展開
        If wsHKbn = 1 Then: Call KOMInfoToGrid1(ID)         ' 逆展開・品番
    Next


    With VsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    
        Call GridSet_TEH(.Row)
        lb_Kensu = Format(.Row, "#,###") & " / " & Format(RecCnt, "#,###")
    End With

    iTotalLineCnt = RecCnt       ' DEBUG

    GridSet_KOM = True
End Function
''wRdKOL(xx, Line)
''  00：検索元品番
''  01：順位
''  02：レベル
''  03：親品番
''  04：子品番
''  05：構成数
''  06：展開区分
''  07：LT
''  08：適用開始日
''  09：適用終了日
''  10：区分
''  11：表示品番
''  12：表示寸法
''  13：品名
''  14：材料管理番号
''  15：検索元車型
'+---------------------------------------------------------------------------------------+
'+      構成情報よりグリッドへセット(正展開)
'+---------------------------------------------------------------------------------------+
Private Sub KOMInfoToGrid0(SetRow As Long)
    With VsGrid1
        ' No
        .TextMatrix(SetRow, 0) = SetRow
        
        '　検索元品番
        .TextMatrix(SetRow, 11) = wRdKOL(0, SetRow)                                         ' 00：検索元品番
        ' 検索元車型
        .TextMatrix(SetRow, 12) = wRdKOL(15, SetRow)                                        ' 15：検索元車型
        
        ' レベル
        Select Case wRdKOL(2, SetRow)                                                       ' 02：レベル
            Case 0: .TextMatrix(SetRow, SCol(1)) = "0"
            Case 1: .TextMatrix(SetRow, SCol(1)) = ".1"
            Case 2: .TextMatrix(SetRow, SCol(1)) = "..2"
            Case 3: .TextMatrix(SetRow, SCol(1)) = "...3"
            Case 4: .TextMatrix(SetRow, SCol(1)) = "....4"
            Case Else
                .TextMatrix(SetRow, SCol(1)) = "....." & Format(wRdKOL(2, SetRow), "##")    ' 02：レベル
        End Select
        ' 区分/品番
        If wRdKOL(10, SetRow) = "" Then                                                     ' 10：区分
            .TextMatrix(SetRow, SCol(2)) = ""
            .TextMatrix(SetRow, SCol(9)) = wRdKOL(4, SetRow)                                ' 04：子品番
        Else
            .TextMatrix(SetRow, SCol(2)) = "材"
            .TextMatrix(SetRow, SCol(9)) = wRdKOL(14, SetRow)                               ' 14：材料管理番号
        End If
        ' 品番
        .TextMatrix(SetRow, SCol(3)) = wRdKOL(11, SetRow)                                   ' 11：表示品番
        
        ' 品名
        .TextMatrix(SetRow, SCol(4)) = wRdKOL(13, SetRow)                                   ' 13：品名
        ' 使用数
        .TextMatrix(SetRow, SCol(5)) = CisFun.RSetFld(wRdKOL(5, SetRow), 10, "#,##0.0000")  ' 05：構成数
        ' 展開区分
        Select Case wRdKOL(0, SetRow)                                                       ' 06：展開区分
            Case 0: .TextMatrix(SetRow, SCol(6)) = ""
            Case 1: .TextMatrix(SetRow, SCol(6)) = "まで対象"
            Case 2: .TextMatrix(SetRow, SCol(6)) = "以降対象外"
            Case 3: .TextMatrix(SetRow, SCol(6)) = "のみ対象外"
        End Select
        ' 有効期限
        gStr = ""
        If Trim(wRdKOL(8, SetRow)) <> "" Then                                               ' 08：適用開始日
            gStr = Format(wRdKOL(8, SetRow), "yyyy/MM/dd")
        End If
        If Trim(wRdKOL(9, SetRow)) <> "" Then                                               ' 09：適用終了日
            If gStr = "" Then gStr = Space(10)
            gStr = gStr & "～"
            gStr = Format(wRdKOL(9, SetRow), "yyyy/MM/dd")
        Else
           If Trim(wRdKOL(8, SetRow)) <> "" Then
              gStr = gStr & "～" & Space(10)
           End If
        End If
        .TextMatrix(SetRow, SCol(7)) = gStr
        ' 寸法
        .TextMatrix(SetRow, SCol(8)) = RTrim(wRdKOL(12, SetRow))                            ' 12：表示寸法
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+      構成情報よりグリッドへセット(逆展開・品番)
'+---------------------------------------------------------------------------------------+
Private Sub KOMInfoToGrid1(SetRow As Long)
    With VsGrid1
        
        ' No
        .TextMatrix(SetRow, 0) = SetRow
        
        
        ' 検索元品番
        .TextMatrix(SetRow, 11) = wRdKOL(0, SetRow)                                         ' 00：検索元品番
        ' 検索元車型
        .TextMatrix(SetRow, 12) = wRdKOL(15, SetRow)                                        ' 15：検索元車型
        
        
        ' 品番
        .TextMatrix(SetRow, SCol(3)) = wRdKOL(11, SetRow)                                   ' 11：表示品番
        .TextMatrix(SetRow, SCol(9)) = wRdKOL(3, SetRow)                                    ' 03：親品番
        ' 品名
        .TextMatrix(SetRow, SCol(4)) = wRdKOL(13, SetRow)                                   ' 13：品名
    End With
End Sub
''+---------------------------------------------------------------------------------------+
''+      構成情報よりグリッドへセット(正展開)
''+---------------------------------------------------------------------------------------+
'Private Sub KOMInfoToGrid0(SetRow As Long)
'    With VsGrid1
'
'        ' No
'        .TextMatrix(SetRow, 0) = SetRow
'
'        ' レベル
'        Select Case KOL.レベル
'            Case 0: .TextMatrix(SetRow, SCol(1)) = "0"
'            Case 1: .TextMatrix(SetRow, SCol(1)) = ".1"
'            Case 2: .TextMatrix(SetRow, SCol(1)) = "..2"
'            Case 3: .TextMatrix(SetRow, SCol(1)) = "...3"
'            Case 4: .TextMatrix(SetRow, SCol(1)) = "....4"
'            Case Else
'                .TextMatrix(SetRow, SCol(1)) = "....." & Format(KOL.レベル, "##")
'        End Select
'        ' 区分/品番
'        If KOL.区分 = "" Then
'            .TextMatrix(SetRow, SCol(2)) = ""
'            .TextMatrix(SetRow, SCol(9)) = KOL.子品番
'        Else
'            .TextMatrix(SetRow, SCol(2)) = "材"
'            .TextMatrix(SetRow, SCol(9)) = KOL.材料管理番号
'        End If
'        ' 品番
'        .TextMatrix(SetRow, SCol(3)) = KOL.表示品番
'
'        ' 品名
'        .TextMatrix(SetRow, SCol(4)) = KOL.品名
'        ' 使用数
'        .TextMatrix(SetRow, SCol(5)) = CisFun.RSetFld(KOL.構成数, 10, "#,##0.0000")
'        ' 展開区分
'        Select Case KOL.展開区分
'            Case 0: .TextMatrix(SetRow, SCol(6)) = ""
'            Case 1: .TextMatrix(SetRow, SCol(6)) = "まで対象"
'            Case 2: .TextMatrix(SetRow, SCol(6)) = "以降対象外"
'            Case 3: .TextMatrix(SetRow, SCol(6)) = "のみ対象外"
'        End Select
'        ' 有効期限
'        gStr = ""
'        If Trim(KOL.適用開始日) <> "" Then
'            gStr = Mid(KOL.適用開始日, 1, 4) & "/" & _
'                   Mid(KOL.適用開始日, 5, 2) & "/" & _
'                   Mid(KOL.適用開始日, 7, 2)
'        End If
'        If Trim(KOL.適用終了日) <> "" Then
'           If gStr = "" Then gStr = Space(10)
'           gStr = gStr & "～"
'           gStr = gStr & Mid(KOL.適用終了日, 1, 4) & "/" & _
'                         Mid(KOL.適用終了日, 5, 2) & "/" & _
'                         Mid(KOL.適用終了日, 7, 2)
'        Else
'           If Trim(KOL.適用開始日) <> "" Then
'              gStr = gStr & "～" & Space(10)
'           End If
'        End If
'        .TextMatrix(SetRow, SCol(7)) = gStr
'        ' 寸法
'        .TextMatrix(SetRow, SCol(8)) = RTrim(KOL.表示寸法)
'    End With
'End Sub
''+---------------------------------------------------------------------------------------+
''+      構成情報よりグリッドへセット(逆展開・品番)
''+---------------------------------------------------------------------------------------+
'Private Sub KOMInfoToGrid1(SetRow As Long)
'    With VsGrid1
'
'        ' No
'        .TextMatrix(SetRow, 0) = SetRow
'
'        ' 品番
'        .TextMatrix(SetRow, SCol(3)) = KOL.表示品番
'        .TextMatrix(SetRow, SCol(9)) = KOL.親品番
'        ' 品名
'        .TextMatrix(SetRow, SCol(4)) = KOL.品名
'    End With
'End Sub
''+---------------------------------------------------------------------------------------+
''+      構成情報よりグリッドへセット(逆展開・材料)
''+---------------------------------------------------------------------------------------+
'Private Sub KOMInfoToGrid2(SetRow As Long)
'    With VsGrid1
'
'        ' No
'        .TextMatrix(SetRow, 0) = SetRow
'
'        ' 品番
'        .TextMatrix(SetRow, SCol(3)) = KOL.表示品番
'        .TextMatrix(SetRow, SCol(9)) = KOL.親品番
'        ' 品名
'        .TextMatrix(SetRow, SCol(4)) = KOL.品名
'        ' 表示寸法
'        .TextMatrix(SetRow, SCol(8)) = KOL.表示寸法
'    End With
'End Sub
'+---------------------------------------------------------------------------------------+
'+  グリッド (RowColChange)
'+---------------------------------------------------------------------------------------+
Private Sub vsGrid1_RowColChange()
    If Not FormAct Then Exit Sub
    If ProcHB <> "B1" Then Exit Sub
    If RowColKB = 1 Then Exit Sub

    With VsGrid1
        If RecCnt >= .Row Then
            If .Tag = "" Then
                Call GridSet_TEH(.Row)
                lb_Kensu = Format(.Row, "#,###") & " / " & Format(RecCnt, "#,###")
            End If
        Else
            RowColKB = 1
            .Row = RecCnt
            .Col = 0
            RowColKB = 0
        End If
        
    End With
End Sub
'******************************************'
'*****       移動選択項目選択          *****
'******************************************'
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        VsGrid1.Row = VsGrid1.MouseRow
        Call Hin_Zai_Copy
        Exit Sub
    End If
    
'< 材料・逆展開 ( = 2 又は、4 ) >
    If CisVsGrid1.PatanNo < 2 Then Exit Sub

' Button : 1 = 左ﾎﾞﾀﾝ ,2 = 右ﾎﾞﾀﾝ
    With VsGrid1
        .Tag = ""
        If Button <> 1 Then Exit Sub
    ' 固定行以外は行わない
        If .MouseRow <> 0 Then Exit Sub
    ' 固定列は行わない
        If .MouseCol = 0 Then Exit Sub
    ' 移動元列をタグへ退避
        .Tag = Str(.MouseCol)
        .Drag 1
    End With
End Sub
'******************************************'
'*****        表 示 項 目 移 動        *****
'******************************************'
Private Sub VSGrid1_DragDrop(Source As Control, X As Single, Y As Single)
    With VsGrid1
    ' 固定列は行わない
        If .MouseCol = 0 Then Exit Sub
        If .Tag = "" Then Exit Sub
' ｸﾞﾘｯﾄﾞ表示 禁止
        .Redraw = False
' 列移動
        .ColPosition(CisFun.Val2(.Tag)) = .MouseCol

        Call CisVsGrid1.VSSort

        .TopRow = .Row

        SCol(3) = CisVsGrid1.FixedGet(3)     ' 品番
        SCol(4) = CisVsGrid1.FixedGet(4)     ' 品名
        SCol(8) = CisVsGrid1.FixedGet(8)     ' 寸法
        
' ｸﾞﾘｯﾄﾞ表示 解除
        .Redraw = True
        CisVsGrid1.ReNumber
            
'< 品番取引先情報セット >
        Call GridSet_TEH(.Row)
        lb_Kensu = Format(.Row, "#,###") & " / " & Format(RecCnt, "#,###")
        .Tag = ""
    End With
End Sub
'+---------------------------------------------------------
'+      グリッド(vsGrid1)上の右クリック　サブメニュー
'+---------------------------------------------------------
Private Function Hin_Zai_Copy()
    Static sw_Hin           As String
    
    With VsGrid1
'        If .MouseCol <> 0 And .MouseRow > 1 Then
        If .MouseCol <> 0 And .MouseRow <> 0 Then
            sw_Hin = RTrim$(.TextMatrix(.MouseRow, SCol(3)))

            If RTrim$(sw_Hin) <> "" Then

                If CisVsGrid1.PatanNo = 0 And _
                    .TextMatrix(.MouseRow, SCol(2)) = "材" Then
                    mn_Copy.Tag = RTrim$(.TextMatrix(.MouseRow, SCol(3)))
                    w_ZaiKB = 1
                Else
                    mn_Copy.Tag = RTrim$(.TextMatrix(.MouseRow, SCol(9)))
                    w_ZaiKB = 0
                End If
                
                If H1Op_Tenkb(0).Tag = "0" Then
                    mn_Gyaku(0).Visible = True
                    mn_Gyaku(1).Visible = True
                    mn_Tenk.Visible = False
                Else
                    mn_Gyaku(0).Visible = False
                    mn_Gyaku(1).Visible = False
                    mn_Tenk.Visible = True
                End If
                
                PopupMenu mn_Menu
            End If
        End If
    End With

End Function
'+---------------------------------------------
'+      品番／材質をクリップボードにコピー
'+---------------------------------------------
Private Sub mn_Copy_Click()
    CisFun.ClipPut RTrim$(mn_Copy.Tag), False
End Sub
'+--------------------------------------
'+      品　番　展　開　(正)
'+--------------------------------------
Private Sub mn_Tenk_Click()
    Call Now_Save
    
    ProcHB = "H1"
    Call DispChange(ProcHB)
    Call HeadBodyClear("B")
    
    H1Op_Hinkb(0).Value = True
    H1Op_Tenkb(0).Value = True
    iH1_Hinbn = RTrim$(mn_Copy.Tag)
    
    If Not Head1Chk Then
        Call GetBack_vsGrid1
        Exit Sub
    End If
    ProcHB = "B1"
    Call DispChange(ProcHB)
End Sub
'+------------------------------------------
'+      品　番　/　材　料　展　開　(逆)
'+------------------------------------------
Private Sub mn_Gyaku_Click(INDEX As Integer)
    Call Now_Save
    
    ProcHB = "H1"
    Call DispChange(ProcHB)
    Call HeadBodyClear("B")

    H1Op_Hinkb(w_ZaiKB).Value = True
    H1Op_Tenkb(1).Value = True
    H1Op_Oyakb(INDEX).Value = True
    iH1_Hinbn = RTrim$(mn_Copy.Tag)
    
    If Not Head1Chk Then
        Call GetBack_vsGrid1
        Exit Sub
    End If
    ProcHB = "B1"
    Call DispChange(ProcHB)
End Sub
'+-------------------------------------
'+      現　状　態　退　避
'+-------------------------------------
Private Function Now_Save()
    GetB_Cnt = GetB_Cnt + 1

    ReDim Preserve sv_Hinkb(GetB_Cnt)
    ReDim Preserve sv_Tenkb(GetB_Cnt)
    ReDim Preserve sv_Oyakb(GetB_Cnt)
    ReDim Preserve sv_Hinbn(GetB_Cnt)
    ReDim Preserve sv_TpRow(GetB_Cnt)
    ReDim Preserve sv_FcRow(GetB_Cnt)
    ReDim Preserve sv_Zaigy(GetB_Cnt)

    sv_Hinkb(GetB_Cnt) = CisFun.Val2(H1Op_Hinkb(0).Tag)
    sv_Tenkb(GetB_Cnt) = CisFun.Val2(H1Op_Tenkb(0).Tag)
    sv_Oyakb(GetB_Cnt) = CisFun.Val2(H1Op_Oyakb(0).Tag)
    sv_Hinbn(GetB_Cnt) = RTrim$(iH1_Hinbn)

    With VsGrid1
        sv_TpRow(GetB_Cnt) = .TopRow
        sv_FcRow(GetB_Cnt) = .Row
    End With
    
    With CisVsGrid1
'                              ( 寸法 )       ( 品番 )       ( 品名 )
        sv_Zaigy(GetB_Cnt) = .FixedGet(8) & .FixedGet(3) & .FixedGet(4)
    End With
    
End Function
'+--------------------------------------------
'+      退　避　し　た　状　態　に　戻　す
'+--------------------------------------------
Private Function GetBack_vsGrid1()
    If GetB_Cnt = 0 Then Exit Function

    ProcHB = "H1"
    Call DispChange(ProcHB)
    Call HeadBodyClear("B")

    H1Op_Hinkb(sv_Hinkb(GetB_Cnt)).Value = True
    H1Op_Tenkb(sv_Tenkb(GetB_Cnt)).Value = True
    H1Op_Oyakb(sv_Oyakb(GetB_Cnt)).Value = True
    iH1_Hinbn = RTrim$(sv_Hinbn(GetB_Cnt))

    Call ReturnPress
    
    On Error Resume Next
    
    With VsGrid1
        If H1Op_Hinkb(1).Value Then     ' 材料逆展開は、列の入替を・・・
            
            gStr = "834"
            
            gLong = CisFun.Val2(Mid$(sv_Zaigy(GetB_Cnt), 1, 1))
            Mid(gStr, gLong, 1) = "8"
            
            gLong = CisFun.Val2(Mid$(sv_Zaigy(GetB_Cnt), 2, 1))
            Mid(gStr, gLong, 1) = "3"
            
            gLong = CisFun.Val2(Mid$(sv_Zaigy(GetB_Cnt), 3, 1))
            Mid(gStr, gLong, 1) = "4"
            
            gStr1 = "0," & Mid$(gStr, 1, 1) & "," & Mid$(gStr, 2, 1) & "," & Mid$(gStr, 3, 1)
            Call CisVsGrid1.PatanSet(3, gStr1)
        
            Call CisVsGrid1.LayOut(3)
        End If
    
        VsGrid1.TopRow = sv_TpRow(GetB_Cnt)
        VsGrid1.Row = sv_FcRow(GetB_Cnt)
    End With
    
    GetB_Cnt = GetB_Cnt - 1
    ReDim Preserve sv_Hinkb(GetB_Cnt)
    ReDim Preserve sv_Tenkb(GetB_Cnt)
    ReDim Preserve sv_Oyakb(GetB_Cnt)
    ReDim Preserve sv_Hinbn(GetB_Cnt)
    ReDim Preserve sv_TpRow(GetB_Cnt)
    ReDim Preserve sv_FcRow(GetB_Cnt)
    ReDim Preserve sv_Zaigy(GetB_Cnt)
    
    PB_GetB.Visible = (GetB_Cnt <> 0)
    On Error GoTo 0
End Function
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
Private Sub GridSet_TEH(pRow As Long)
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
    Static wsKB     As String
        
        
    H1lb_Hinbn.Caption = VsGrid1.TextMatrix(pRow, 11)           ' 2019/07/19 ADD：仮表示
    H1lb_Syake.Caption = VsGrid1.TextMatrix(pRow, 12)           ' 2019/07/22 ADD：仮表示
            
        
    Call CisVsGridT.Clear
    Call GridInitT

    With VsGrid1
        wsKB = RTrim(.TextMatrix(pRow, SCol(2)))
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
            gSL_Select = gSL_Select & " WHERE HT.品番 = '" & RTrim(.TextMatrix(pRow, SCol(9))) & "'"
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
            gSL_Select = gSL_Select & " WHERE ZR.材料管理番号 = " & .TextMatrix(pRow, SCol(9))
            gSL_Select = gSL_Select & " ORDER BY ZT.取引先"
            Call ZTMRead(gSL_Select)
        End If
    End With
    Call Str_Monitor(gSL_Select)            ' 2019/07/22 ADD

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
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  構成情報よりグリッドへセット
'+---------------------------------------------------------------------------------------+
Private Sub HTMInfoToGrid(SetRow As Long, KB As String)
    With VsGridT
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
'
''+++++++++++++++++++++++++++++++++++++++++++++++++++    '++ デバッグ用です +++
''+                                                 +
''+  文字列をイミディエイトウィンドウに表示する     +
''+                                                 +
''+++++++++++++++++++++++++++++++++++++++++++++++++++
'Private Sub Str_Monitor(ByVal wStr As String)
'    Debug.Print vbCrLf
'    Debug.Print Format(Now, "yyyy/MM/dd HH:mm:ss") + _
'                " .................................................." + _
'                vbCrLf + wStr
'End Sub
'
''+++++++++++++++++++++++++++++++++++++++++++++++++++    '++ デバッグ用です +++
''+                                                 +
''+  KOL取得情報をCSVファイルに出力する             +
''+                                                 +
''+++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub H1lb_Syake_Click()
    Call TestDataToVSV
End Sub
'


