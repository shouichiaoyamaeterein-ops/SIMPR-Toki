VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "Cis3D_6.0.ocx"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CisBtn_6.0.ocx"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CisText_6.0.ocx"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CXM0150 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "構成マスタ照会"
   ClientHeight    =   11190
   ClientLeft      =   1890
   ClientTop       =   2910
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D lb_JyutyuBiko 
      Height          =   300
      Left            =   435
      Top             =   1785
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   529
      BackColor       =   16777152
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
   End
   Begin VB.PictureBox B1_Area3 
      Height          =   1605
      Left            =   465
      ScaleHeight     =   1545
      ScaleWidth      =   14475
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   9450
      Width           =   14535
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   1365
         Left            =   90
         Top             =   90
         Width           =   14355
         _ExtentX        =   25321
         _ExtentY        =   2408
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
         Begin vsFlexLib.vsFlexArray VsGridN 
            Height          =   1335
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   14325
            _Version        =   196608
            _ExtentX        =   25268
            _ExtentY        =   2355
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
            FormatString    =   "№ |納入先　　　  　　                  |受|背番号|契約先　　　    　 　                |SK|ｻｲｸﾙ   |収容数 |収容器   "
            Rows            =   3
            BackColor       =   12648384
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   330
            SelectionMode   =   1
            AllowBigSelection=   0   'False
            BackColorAlternate=   12648384
            AllowSelection  =   0   'False
         End
      End
   End
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
      Left            =   450
      ScaleHeight     =   2220
      ScaleWidth      =   14475
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7080
      Width           =   14535
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   2055
         Left            =   90
         Top             =   90
         Width           =   14325
         _ExtentX        =   25268
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
            TabIndex        =   11
            Top             =   -30
            Width           =   14295
            _Version        =   196608
            _ExtentX        =   25215
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
            ConvInfo        =   1418783674
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
      Height          =   1380
      Left            =   12285
      ScaleHeight     =   1320
      ScaleWidth      =   2640
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   570
      Width           =   2700
      Begin CisBtn_60.CisBtn PB_GetB 
         Height          =   585
         Left            =   1725
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
         Left            =   60
         Top             =   660
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
         Left            =   900
         Top             =   660
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
         Left            =   1740
         Top             =   660
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
         Left            =   885
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
      Begin CisBtn_60.CisBtn PB_Prt 
         Height          =   585
         Left            =   60
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
      Left            =   450
      ScaleHeight     =   915
      ScaleWidth      =   11175
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   750
      Width           =   11235
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
            Caption         =   "品番"
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   0
            Top             =   75
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton H1Op_Hinkb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "材料"
            Height          =   315
            Index           =   1
            Left            =   180
            TabIndex        =   1
            Top             =   420
            Width           =   855
         End
      End
      Begin Cis3D_v60.CIS3D HLB_Hinbn 
         Height          =   795
         Left            =   4650
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
            TabIndex        =   6
            Top             =   360
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   795
         Left            =   1245
         Top             =   60
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "表  示  区  分"
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
         Begin Cis3D_v60.CIS3D HLB_Tenkb 
            Height          =   345
            Left            =   1380
            Top             =   360
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
               Caption         =   "製品"
               Height          =   255
               Index           =   1
               Left            =   1020
               TabIndex        =   5
               Top             =   60
               Width           =   795
            End
            Begin VB.OptionButton H1Op_Oyakb 
               BackColor       =   &H00E0E0E0&
               Caption         =   "直前"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   4
               Top             =   60
               Width           =   795
            End
         End
         Begin VB.OptionButton H1Op_Tenkb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "逆"
            Height          =   315
            Index           =   1
            Left            =   780
            TabIndex        =   3
            Top             =   390
            Width           =   585
         End
         Begin VB.OptionButton H1Op_Tenkb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "正"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   390
            Value           =   -1  'True
            Width           =   585
         End
      End
      Begin Cis3D_v60.CIS3D HLB_Long 
         Height          =   795
         Left            =   10065
         Top             =   60
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "長さ"
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
         Begin CisText_V60.CisText iH1_Long 
            Height          =   360
            Left            =   45
            TabIndex        =   9
            Top             =   330
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   635
            cFormat         =   "###0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
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
      End
      Begin Cis3D_v60.CIS3D HLB_Itatu 
         Height          =   795
         Left            =   7965
         Top             =   60
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "板厚"
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
         Begin CisText_V60.CisText iH1_Itatu 
            Height          =   360
            Left            =   45
            TabIndex        =   7
            Top             =   345
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   635
            cFormat         =   "###0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
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
      End
      Begin Cis3D_v60.CIS3D HLB_width 
         Height          =   795
         Left            =   9015
         Top             =   60
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "幅"
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
         Begin CisText_V60.CisText iH1_Width 
            Height          =   360
            Left            =   45
            TabIndex        =   8
            Top             =   330
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   635
            cFormat         =   "###0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
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
      TabIndex        =   13
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
      Height          =   4875
      Left            =   450
      ScaleHeight     =   4815
      ScaleWidth      =   14460
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2130
      Width           =   14520
      Begin vsFlexLib.vsFlexArray VsGrid1 
         Height          =   4650
         Left            =   90
         TabIndex        =   10
         Top             =   75
         Width           =   14310
         _Version        =   196608
         _ExtentX        =   25241
         _ExtentY        =   8202
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
         FormatString    =   "No |レベル|手配先            |区|XXXXXXXXX1XXXXXXXXX2|XXXXXXXXX1XXXXXXXXX2|使用数|展開区分  |9999/99/99-9999/99/99|寸法||"
         Rows            =   15
         Cols            =   21
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
      End
      Begin VB.Menu mn_Gyaku 
         Caption         =   "再表示( 逆・直前 )"
         Index           =   0
      End
      Begin VB.Menu mn_Gyaku 
         Caption         =   "再表示( 逆・製品 )"
         Index           =   1
      End
      Begin VB.Menu mn_Dmy2 
         Caption         =   "-"
      End
      Begin VB.Menu mn_Can 
         Caption         =   "キャンセル"
      End
   End
End
Attribute VB_Name = "CXM0150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   構成マスタ問合せ
'**       フォームID    :   CXM0150
'**       処理概要      :
'**
'**       作  成  日    :   2003/11/25  By CIS  SIMPR-A
'**       変  更  日    :   2011/02/28  BY CIS  画面・帳票に背番号・収容数を追加
'**       変  更  日    :   2011/04/04  By CIS  支給は表示しない
'**       変  更  日    :   2011/07/19  By CIS  条件:材料時、板厚・幅・長さを表示条件に追加
'**       変  更  日    :   2011/07/21  BY CIS  逆展開時：画面・帳票に背番号・収容数を追加
'**                                             画面：納入先情報表示を追加
'**                                             画面表示条件：比率が最大のものに変更(但し、同一比率の場合は現行通り)
'**
'**       変  更  日    :   2012/07/03  By CIS  号補区分表示無
'**       変  更  日    :   2012/09/20  By CIS  材料[ﾋﾟｯﾁ]追加
'**       変  更  日    :   2013/03/15  By CIS  CIS8606参照[手配先]取得
'**       変  更  日    :   2013/09/16  By CIS  帳票(CXM0150L)一部修正
'**       変  更  日    :   2014/02/13  By CIS  受注備考欄追加
'**       変  更  日    :   2015/01/21  By CIS  背番号枠拡大(８桁）
'**       変  更  日    :   2015/08/17  By CIS  受注収容数変更日対応
'**       変  更  日    :   2016/11/23  By CIS  背番号表示変更
'**       変  更  日    :   2017/11/08  By CIS  ＣＳＶ初期値はディスクトップ
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
    '-----2011/07/21 ADD START-----
    Dim CisVsGridN      As New CisVsGrid3
    '-----2011/07/21 ADD END-------
    Dim ID              As Long
    
    Dim ActObj          As Object
    Dim ActRow          As Integer
    
    Dim KCol            As Long         ' 非表示項目開始ｶﾗﾑ
    
    Dim SCol()          As Long         ' 表示項目ｶﾗﾑ(品番)
    Dim SColT()         As Long         ' 表示項目ｶﾗﾑ(材料)
    Dim SColN()         As Long

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
    
    Dim wFlg            As Boolean        '2013/03/15

    Dim objWshShell                     '2017/11/08
    Dim mWKDir      As String
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
    'gConnectCount = 3 'Insert 2011/03/01
    gConnectCount = 4 'UPDATE 2013/03/15
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
        .SelectCol = 20
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
    
    '-----2011/07/21 ADD START-----
    '+--------------------------------+
    '+  グリッド初期設定(手配)
    '+--------------------------------+
    With CisVsGridN
        Set .GridObj = VsGridN
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 16
        .InitGet = False
        .Init
    End With
    Call GridInitN
    '-----2011/07/21 ADD END-------

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

    Set objWshShell = CreateObject("WScript.Shell")             '2017/11/08
    mWKDir = GetSetting(appname:=gRegKey, section:="CXM00150", _
                       Key:="Csv", Default:=Mid(CisFun.WinRoot, 1, 3) & "")
    
    mCSVDir = objWshShell.SpecialFolders("Desktop") & "\構成照会.csv"               '2017/11/08
    
        SaveSetting appname:=gRegKey, section:="CXM00150", _
                        Key:="Csv", setting:=RTrim(mCSVDir)
    
'    If Not mWKDir Like "*Desktop*" And Not mWKDir Like "*デスクトップ*" Then    '2017/11/08
'        SaveSetting appname:=gRegKey, section:="CXM00150", _
'                        Key:="Csv", setting:=RTrim(mCSVDir)
'
'    End If
    
    
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
           Case vbKeyF8:       If PB_Prt.Visible Then Call PB_Prt_Click    '【印刷】
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
    '-----2011/07/19 ADD START-----
    If PB_Look.Tag = "iH1_Itatu" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Width" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Long" Then Call Look_Hinbn: Exit Sub
    '-----2011/07/19 ADD END-------
End Sub
'+----------------------------+
'+    印　刷　キ　ー(F8)      +
'+----------------------------+
Private Sub PB_Prt_Click()

    Dim PrtCnt     As Long
    Dim PrtOut      As Integer
    
    Dim GyakKB     As Byte
    Dim HZKB       As Byte
    Dim TSKB       As Byte
    
    With CisSimplePrint
         .PrintTitle = "構成マスタ照会一覧"
         .SaveKey = gRegKey
         .Preservation = "CXM00150"
         
         .GetCsvFileName = mCSVDir & "11111.csv"
         
         
         If Not .CheckPrtCsv Then Exit Sub
    
         Select Case .PrintStatus
            Case 4: PrtOut = 0
            Case 5: PrtOut = 1
            Case 3: PrtOut = 2
            Case Else: GoTo PB_Prt_Ed
         End Select
         
    End With
    
    'If Not CXM0150LPrint(PrtOut) Then Exit Sub       2009/11/10
    '================================================ 2009/11/10 Start
    If H1Op_Tenkb(0) Then GyakKB = 0 Else GyakKB = 1
    If H1Op_Oyakb(0) Then TSKB = 0 Else TSKB = 1
    If H1Op_Hinkb(0) Then HZKB = 0 Else HZKB = 1
    If Not CXM0150LPrint(PrtOut, GyakKB, HZKB, TSKB) Then Exit Sub
    '================================================ 2009/11/10 End
    
    Call CisSimplePrint.CheckFormUnload
PB_Prt_Ed:
End Sub

'+----------------------------+
'+    戻　す  キ ー(F9)       +
'+----------------------------+
Private Sub PB_GetB_Click()
    Call GetBack_vsGrid1
End Sub
'+----------------------------+
'+       品番検索              +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Hinbn = ""
    RV_Zaist = ""
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"

    If CisFun.Val2(H1Op_Hinkb(0).Tag) = 0 Then
        CKK0030.Show vbModal
        Unload CKK0030
        Set CKK0030 = Nothing
        If RV_Rtn Then
            iH1_Hinbn = RV_Hinbn
        End If
    Else
        CKK0070.Show vbModal
        Unload CKK0070
        Set CKK0070 = Nothing
        If RV_Rtn Then
            iH1_Hinbn = RV_Zaist
            '-----2011/07/15 ADD START-----
            iH1_Itatu = RV_Itatu
            iH1_Width = RV_Width
            iH1_Long = RV_Long
            '-----2011/07/15 ADD START-----
        End If
    End If
End Sub
'+---( H1 )----------------------------------------------------------------------------
'+--------------------------------------------------+
'+  H1Op_Hinkb  (品番/材料)
'+--------------------------------------------------+
Private Sub H1Op_Hinkb_Click(Index As Integer)
    H1Op_Hinkb(0).Tag = Index
    
    If Index = 0 Then
        H1Op_Tenkb(0).Enabled = True
        HLB_Hinbn.Caption = "品   番"
        '-----2011/07/15 ADD START-----
        HLB_Itatu.Visible = False
        HLB_width.Visible = False
        HLB_Long.Visible = False
        H1_Area1.Width = 8055
        iH1_Itatu.Text = ""
        iH1_Width.Text = ""
        iH1_Long.Text = ""
        '-----2011/07/15 ADD END-------
    Else
        H1Op_Tenkb(0).Enabled = False
        H1Op_Tenkb(1).Value = True
        HLB_Hinbn.Caption = "材   質"
        '-----2011/07/15 ADD START-----
        HLB_Itatu.Visible = True
        HLB_width.Visible = True
        HLB_Long.Visible = True
        H1_Area1.Width = 11235
        iH1_Itatu.Text = ""
        iH1_Width.Text = ""
        iH1_Long.Text = ""
        '-----2011/07/15 ADD END-------
    End If
End Sub
Private Sub H1Op_Hinkb_GotFocus(Index As Integer)
    H1Op_Hinkb(Index).BackColor = &HC0FFFF
End Sub
Private Sub H1Op_Hinkb_LostFocus(Index As Integer)
    H1Op_Hinkb(Index).BackColor = &HC0C0C0
End Sub
'+--------------------------------------------------+
'+  H1Op_Tenkb  (正/逆展開)
'+--------------------------------------------------+
Private Sub H1Op_Tenkb_Click(Index As Integer)
    H1Op_Tenkb(0).Tag = Index
    If Index = 0 Then
        H1Op_Oyakb(0).Value = False
        H1Op_Oyakb(1).Value = False
        HLB_Tenkb.Enabled = False
    Else
        HLB_Tenkb.Enabled = True
        H1Op_Oyakb(0).Value = True
    End If
End Sub
Private Sub H1Op_Tenkb_GotFocus(Index As Integer)
    H1Op_Tenkb(Index).BackColor = &HC0FFFF
End Sub
Private Sub H1Op_Tenkb_LostFocus(Index As Integer)
    H1Op_Tenkb(Index).BackColor = &HC0C0C0
End Sub
'+--------------------------------------------------+
'+  H1Op_Oyakb  (直前/製品)
'+--------------------------------------------------+
Private Sub H1Op_Oyakb_Click(Index As Integer)
    H1Op_Oyakb(0).Tag = Index
End Sub
Private Sub H1Op_Oyakb_GotFocus(Index As Integer)
    H1Op_Oyakb(Index).BackColor = &HC0FFFF
End Sub
Private Sub H1Op_Oyakb_LostFocus(Index As Integer)
    H1Op_Oyakb(Index).BackColor = &HE0E0E0
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
'-----2011/07/19 ADD START-----
'+-----------------------------+
'+  iH1_Itatu
'+-----------------------------+
Private Sub iH1_Itatu_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Itatu_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_Width
'+-----------------------------+
Private Sub iH1_Width_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Width_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_Long
'+-----------------------------+
Private Sub iH1_Long_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Long_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'-----2011/07/19 ADD END-------
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
    '-----2011/07/21 ADD START-----
    B1_Area3.Enabled = B1Mode
    B1_Area3.BackColor = B1Color
    '-----2011/07/21 ADD END-------
    PB_CAN.Visible = Not H1Mode
    
'    PB_Prt.Visible = B1Mode And H1Op_Hinkb(0).Value And H1Op_Tenkb(0).Value        '2009/11/10
    PB_Prt.Visible = B1Mode                                                         '2009/11/10
    
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
        '-----2011/07/21 ADD START-----
        VsGridN.Clear
        '-----2011/07/21 ADD END-------
        wsHKbn = 0
        Call GridInit1
        Call GridInitT
        '-----2011/07/21 ADD START-----
        Call GridInitN
        '-----2011/07/21 ADD END-------
        lb_Kensu = ""
        
        lb_JyutyuBiko = ""                  '2014/02/13
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
        
        '-----2011/07/21 ADD START-----
        VsGrid1.SetFocus
        '-----2011/07/21 ADD END-------
        
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
            If CisFun.Val2(H1Op_Hinkb(0).Tag) = 0 Then
                .MB_MSG(2) = "    品番を入力して下さい            "
            Else
                .MB_MSG(2) = "    材質を入力して下さい            "
            End If
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
    
'============ 2014/02/13 Start
'受注備考セット
    If H1Op_Hinkb(0).Value = True And H1Op_Tenkb(0).Value = True Then
        gSL_Select = "SELECT MAX(備考) 備考 FROM 品番取引先マスタ "
        gSL_Select = gSL_Select & " WHERE 品目 = 0 "
        gSL_Select = gSL_Select & " AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
        If HTMRead(gSL_Select, 1, 1) Then
            lb_JyutyuBiko = RTrim(HTM.備考)
        Else
            lb_JyutyuBiko = ""
        End If
    Else
        lb_JyutyuBiko = ""
    End If
'============ 2014/02/13 End


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

        ReDim SCol(14) 'Update 2012/09/20
        'ReDim SCol(13) 'Update 2011/02/28
        'ReDim SCol(11)

        Call .FixedSet(中中, 右中, なし, 4, 0, "No")
        Call .FixedSet(中中, 左中, あり, 7, 1, "レベル")
        Call .FixedSet(中中, 左中, あり, 2, 2, "区")
'        Call .FixedSet(中中, 左中, あり, 25, 3, "品  番")                  '2015/01/21
        Call .FixedSet(中中, 左中, あり, 24, 3, "品  番")                   '2015/01/21
        Call .FixedSet(中中, 左中, あり, 20, 4, "品  名")
        
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert Start
'        Call .FixedSet(中中, 左中, あり, 7, 5, "背番号")                   '2015/01/21
        Call .FixedSet(中中, 左中, あり, 8, 5, "背番号")                    '2015/01/21
        Call .FixedSet(中中, 右中, あり, 10, 6, "収容数")
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert End
        
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Update Start
        Call .FixedSet(中中, 右中, あり, 10, 7, "使用数")
        Call .FixedSet(中中, 左中, あり, 4, 8, "号補")
        Call .FixedSet(中中, 左中, あり, 10, 9, "展開区分")
        Call .FixedSet(中中, 左中, あり, 20, 10, "適用範囲")
        Call .FixedSet(中中, 左中, あり, 25, 11, "寸  法")
        Call .FixedSet(中中, 左中, あり, 25, 12, "手配先")
        Call .FixedSet(中中, 左中, あり, 0, 14, "編集前品番")       '2012/09/20
''        Call .FixedSet(中中, 左中, あり, 0, 13, "編集前品番")     '2012/09/20
'        Call .FixedSet(中中, 右中, あり, 10, 5, "使用数")
'        Call .FixedSet(中中, 左中, あり, 4, 6, "号補")
'        Call .FixedSet(中中, 左中, あり, 10, 7, "展開区分")
'        Call .FixedSet(中中, 左中, あり, 21, 8, "適用範囲")
'        Call .FixedSet(中中, 左中, あり, 25, 9, "寸  法")
'        Call .FixedSet(中中, 左中, あり, 25, 10, "手配先")
'        Call .FixedSet(中中, 左中, あり, 0, 11, "編集前品番")
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Update Start
        
        Call .FixedSet(中中, 左中, あり, 9, 13, "ﾋﾟｯﾁ")             '2012/09/20
        
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, あり, 0, 20, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Update Start
'        Call .PatanSet(0, "0,1,12,2,3,10,4,5,6,7,8,9,10")          '2012/09/20
        Call .PatanSet(0, "0,1,12,2,3,10,4,13,5,6,7,8,9,10")        '2012/09/20
        '-----2011/07/21 EDIT START-----
'''        Call .PatanSet(1, "0,12,3,4")
'''        Call .PatanSet(2, "0,12,10,3,4")
'        Call .PatanSet(1, "0,12,3,4,5,6")                          '2012/09/20
'        Call .PatanSet(2, "0,12,10,3,4,5,6")                       '2012/09/20
        Call .PatanSet(1, "0,12,3,4,13,5,6")                        '2012/09/20
        Call .PatanSet(2, "0,12,10,3,4,13,5,6")                     '2012/09/20
        '-----2011/07/21 EDIT END-------
'        Call .PatanSet(0, "0,1,10,2,3,8,4,5,6,7,8")
'        Call .PatanSet(1, "0,10,3,4")
'        Call .PatanSet(2, "0,10,8,3,4")
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Update Start
        
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
        Call CisVsGrid1.LayOut(wsHKbn)
        
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Update Start
        SCol(1) = .FixedGet(1)     ' レベル
        SCol(2) = .FixedGet(2)     ' 区
        SCol(3) = .FixedGet(3)     ' 品番
        SCol(4) = .FixedGet(4)     ' 品名
        SCol(5) = .FixedGet(5)     ' 背番号
        SCol(6) = .FixedGet(6)     ' 収容数
        SCol(7) = .FixedGet(7)     ' 使用数
        SCol(8) = .FixedGet(8)     ' 号補区分
        SCol(9) = .FixedGet(9)     ' 展開区分
        SCol(10) = .FixedGet(10)   ' 適用範囲
        SCol(11) = .FixedGet(11)   ' 寸法
        SCol(12) = .FixedGet(12)   ' 手配先
        SCol(14) = .FixedGet(14)   ' 編集前品番                 '2012/09/20
''        SCol(13) = .FixedGet(13)   ' 編集前品番               '2012/09/20
'        SCol(1) = .FixedGet(1)     ' レベル
'        SCol(2) = .FixedGet(2)     ' 区
'        SCol(3) = .FixedGet(3)     ' 品番
'        SCol(4) = .FixedGet(4)     ' 品名
'        SCol(5) = .FixedGet(5)     ' 使用数
'        SCol(6) = .FixedGet(6)     ' 展開区分
'        SCol(7) = .FixedGet(7)     ' 適用範囲
'        SCol(8) = .FixedGet(8)     ' 寸法
'        SCol(9) = .FixedGet(9)     ' 編集前品番
'        SCol(10) = .FixedGet(10)   ' 手配先
'        SCol(11) = .FixedGet(11)   ' 編集前品番
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Update Start
        SCol(13) = .FixedGet(13)    ' ピッチ                    '2012/09/20
    End With
    
    '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Update Start
    KCol = 11   ' 非表示項目開始ｶﾗﾑ
    'KCol = 9   ' 非表示項目開始ｶﾗﾑ
    '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Update Start
End Sub
'+---------------------------------------------------------------------------------------+
'+      グリッド内容セット
'+---------------------------------------------------------------------------------------+
Private Function GridSet_KOM()
    GridSet_KOM = False
    
    '-----2011/07/21 ADD START-----
    Dim TaihiRitsu As Integer
    Dim TaihiRitsuZ As Integer
    '-----2011/07/21 ADD END-------
    
    Call GridInit1
    With CisDB
        .SQL = "構成問合せ情報抽出"
        '-----2011/07/19 EDIT START-----
        '.StoadoCount = 4
        .StoadoCount = 7
        '-----2011/07/19 EDIT END-------
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = Input用
        .ParaNo = 3: .ParaIO = Input用
        .ParaNo = 4: .ParaIO = Input用
        '-----2011/07/19 ADD START-----
        .ParaNo = 5: .ParaIO = Input用
        .ParaNo = 6: .ParaIO = Input用
        .ParaNo = 7: .ParaIO = Input用
        '-----2011/07/19 ADD END-------
    
        .ParaNo = 1: .ParaValue = RTrim(iH1_Hinbn)
        .ParaNo = 2: .ParaValue = CisFun.Val2(H1Op_Hinkb(0).Tag)
        .ParaNo = 3: .ParaValue = CisFun.Val2(H1Op_Tenkb(0).Tag)
        .ParaNo = 4: .ParaValue = CisFun.Val2(H1Op_Oyakb(0).Tag)
        '-----2011/07/19 ADD START-----
        .ParaNo = 5: .ParaValue = Trim(Format(iH1_Itatu, "#.00"))
        .ParaNo = 6: .ParaValue = Trim(Format(iH1_Width, "#.00"))
        .ParaNo = 7: .ParaValue = Trim(Format(iH1_Long, "#.00"))
        '-----2011/07/19 ADD END-------
    
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
        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT KW.*," & vbCrLf
        gSL_Select = gSL_Select & "       CASE WHEN HTM.号補区分 = ''  THEN ''" & vbCrLf
        gSL_Select = gSL_Select & "            WHEN HTM.号補区分 = 'G' THEN '号口'" & vbCrLf
        gSL_Select = gSL_Select & "            WHEN HTM.号補区分 = 'H' THEN '補給' END 号補区分," & vbCrLf
        gSL_Select = gSL_Select & "       CASE WHEN ISNULL(HTM1.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(TRM1.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                      THEN HTM1.取引先 + ':' + TRM1.略称" & vbCrLf
        gSL_Select = gSL_Select & "                      ELSE HTM1.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "            WHEN ISNULL(HTM2.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(TRM2.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                      THEN HTM2.取引先 + ':' + TRM2.略称" & vbCrLf
        gSL_Select = gSL_Select & "                      ELSE HTM2.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "            WHEN ISNULL(ZTM1.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(TRM3.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                      THEN ZTM1.取引先 + ':' + TRM3.略称" & vbCrLf
        gSL_Select = gSL_Select & "                      ELSE ZTM1.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "            WHEN ISNULL(ZTM2.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(TRM4.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                      THEN ZTM2.取引先 + ':' + TRM4.略称" & vbCrLf
        gSL_Select = gSL_Select & "                      ELSE ZTM2.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "            ELSE '' END 手配先" & vbCrLf

        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert Start
        gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(HTM1.取引先, '') <> '' THEN HTM1.取引先" & vbCrLf
        gSL_Select = gSL_Select & "            WHEN ISNULL(HTM2.取引先, '') <> '' THEN HTM2.取引先" & vbCrLf
        gSL_Select = gSL_Select & "            WHEN ISNULL(ZTM1.取引先, '') <> '' THEN ZTM1.取引先" & vbCrLf
        gSL_Select = gSL_Select & "            WHEN ISNULL(ZTM2.取引先, '') <> '' THEN ZTM2.取引先" & vbCrLf
        gSL_Select = gSL_Select & "            ELSE '' END 取引先" & vbCrLf
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert End

        gSL_Select = gSL_Select & "  FROM 構成問合せワーク KW " & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(号補区分) 号補区分, 品番" & vbCrLf
        gSL_Select = gSL_Select & "                          FROM 品番取引先マスタ" & vbCrLf
        gSL_Select = gSL_Select & "                         GROUP BY 品番) HTM" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM.品番 = KW.子品番" & vbCrLf
        gSL_Select = gSL_Select & "           AND KW.区分 = ''" & vbCrLf
'**********************************************2011/07/21 EDIT START**********************************************
       gSL_Select = gSL_Select & "   LEFT OUTER JOIN (Select Cast(Substring(WK1.キー,1,3) As Integer) 発注比率, Substring(WK1.キー,4,7) 取引先, WK1.品番" & vbCrLf
       gSL_Select = gSL_Select & "                      From (" & vbCrLf
       gSL_Select = gSL_Select & "                            Select MAX(Right('000' + Cast(発注比率 As Varchar(3)),3) + 取引先) キー,品番" & vbCrLf
       gSL_Select = gSL_Select & "                              From 品番取引先マスタ" & vbCrLf
       gSL_Select = gSL_Select & "                             WHERE ( 手配区分 <> '2' AND 手配区分 <> '4' ) AND 品目 = 1" & vbCrLf
       gSL_Select = gSL_Select & "                             Group By 品番" & vbCrLf
       gSL_Select = gSL_Select & "                           ) WK1" & vbCrLf
       gSL_Select = gSL_Select & "                   ) HTM1" & vbCrLf
       gSL_Select = gSL_Select & "     ON HTM1.品番 = KW.子品番" & vbCrLf
         
       gSL_Select = gSL_Select & "   LEFT OUTER JOIN (Select Cast(Substring(WK2.キー,1,3) As Integer) 発注比率, Substring(WK2.キー,4,7) 取引先, WK2.品番" & vbCrLf
       gSL_Select = gSL_Select & "                      From (" & vbCrLf
       gSL_Select = gSL_Select & "                            Select MAX(Right('000' + Cast(発注比率 As Varchar(3)),3) + 取引先) キー,品番" & vbCrLf
       gSL_Select = gSL_Select & "                              From 品番取引先マスタ" & vbCrLf
       gSL_Select = gSL_Select & "                             WHERE 手配区分 = '2' AND 品目 = 1" & vbCrLf
       gSL_Select = gSL_Select & "                             Group By 品番" & vbCrLf
       gSL_Select = gSL_Select & "                           ) WK2" & vbCrLf
       gSL_Select = gSL_Select & "                   ) HTM2" & vbCrLf
       gSL_Select = gSL_Select & "     ON HTM2.品番 = KW.子品番" & vbCrLf
         
       gSL_Select = gSL_Select & "   LEFT OUTER JOIN (Select Cast(Substring(WK3.キー,1,3) As Integer) 発注比率, Substring(Wk3.キー,4,7) 取引先,WK3.材料管理番号" & vbCrLf
       gSL_Select = gSL_Select & "                      From (" & vbCrLf
       gSL_Select = gSL_Select & "                            Select MAX(Right('000' + Cast(発注比率 As Varchar(3)),3) + 取引先) キー,材料管理番号" & vbCrLf
       gSL_Select = gSL_Select & "                              From 材料取引先マスタ" & vbCrLf
       gSL_Select = gSL_Select & "                             Where ( 手配区分 <> '2' AND 手配区分 <> '4' )" & vbCrLf
       gSL_Select = gSL_Select & "                             Group By 材料管理番号" & vbCrLf
       gSL_Select = gSL_Select & "                           ) WK3" & vbCrLf
       gSL_Select = gSL_Select & "                   ) ZTM1" & vbCrLf
       gSL_Select = gSL_Select & "     ON CAST(ZTM1.材料管理番号 AS VARCHAR) = KW.材料管理番号" & vbCrLf
         
       gSL_Select = gSL_Select & "   LEFT OUTER JOIN (Select Cast(Substring(WK4.キー,1,3) As Integer) 発注比率, Substring(WK4.キー,4,7) 取引先,WK4.材料管理番号" & vbCrLf
       gSL_Select = gSL_Select & "                      From (" & vbCrLf
       gSL_Select = gSL_Select & "                            Select MAX(Right('000' + Cast(発注比率 As Varchar(3)),3) + 取引先) キー,材料管理番号" & vbCrLf
       gSL_Select = gSL_Select & "                              From 材料取引先マスタ" & vbCrLf
       gSL_Select = gSL_Select & "                             Where ( 手配区分 = '2' )" & vbCrLf
       gSL_Select = gSL_Select & "                             Group By 材料管理番号" & vbCrLf
       gSL_Select = gSL_Select & "                           ) WK4" & vbCrLf
       gSL_Select = gSL_Select & "                   ) ZTM2" & vbCrLf
       gSL_Select = gSL_Select & "     ON CAST(ZTM1.材料管理番号 AS VARCHAR) = KW.材料管理番号" & vbCrLf

'''        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 品番" & vbCrLf
'''        gSL_Select = gSL_Select & "                          FROM 品番取引先マスタ" & vbCrLf
''''======================= 2011/04/04 Start
''''        gSL_Select = gSL_Select & "                         WHERE 手配区分 <> '2' AND 品目 = 1 GROUP BY 品番) HTM1" & vbCrLf
'''        gSL_Select = gSL_Select & "                         WHERE ( 手配区分 <> '2' AND 手配区分 <> '4' ) AND 品目 = 1 GROUP BY 品番) HTM1" & vbCrLf
''''======================= 2011/04/04 End
'''        gSL_Select = gSL_Select & "            ON HTM1.品番 = KW.子品番" & vbCrLf
'''        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 品番" & vbCrLf
'''        gSL_Select = gSL_Select & "                          FROM 品番取引先マスタ" & vbCrLf
'''        gSL_Select = gSL_Select & "                         WHERE 手配区分 = '2' AND 品目 = 1 GROUP BY 品番) HTM2" & vbCrLf
'''        gSL_Select = gSL_Select & "            ON HTM2.品番 = KW.子品番" & vbCrLf
'''
'''        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 材料管理番号" & vbCrLf
'''        gSL_Select = gSL_Select & "                          FROM 材料取引先マスタ" & vbCrLf
''''======================= 2011/04/04 Start
''''        gSL_Select = gSL_Select & "                         WHERE 手配区分 <> '2' GROUP BY 材料管理番号) ZTM1" & vbCrLf
'''        gSL_Select = gSL_Select & "                         WHERE ( 手配区分 <> '2' AND 手配区分 <> '4' ) GROUP BY 材料管理番号) ZTM1" & vbCrLf
''''======================= 2011/04/04 End
'''        gSL_Select = gSL_Select & "            ON CAST(ZTM1.材料管理番号 AS VARCHAR) = KW.材料管理番号" & vbCrLf
'''        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 材料管理番号" & vbCrLf
'''        gSL_Select = gSL_Select & "                          FROM 材料取引先マスタ" & vbCrLf
'''        gSL_Select = gSL_Select & "                         WHERE 手配区分 = '2' GROUP BY 材料管理番号) ZTM2" & vbCrLf
'''        gSL_Select = gSL_Select & "            ON CAST(ZTM2.材料管理番号 AS VARCHAR) = KW.材料管理番号" & vbCrLf
'**********************************************2011/07/21 EDIT END**********************************************
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM1" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM1.取引先 = TRM1.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM1.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM2" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM2.取引先 = TRM2.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM2.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM3" & vbCrLf
        gSL_Select = gSL_Select & "            ON ZTM1.取引先 = TRM3.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM3.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM4" & vbCrLf
        gSL_Select = gSL_Select & "            ON ZTM2.取引先 = TRM4.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM4.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & "   WHERE KW.処理端末 = HOST_NAME()" & vbCrLf
        gSL_Select = gSL_Select & " ORDER BY KW.順位,KW.子品番" & vbCrLf
    End If
    If wsHKbn = 1 Then
        gSL_Select = "SELECT KW.親品番,KW.品名,KW.表示品番," & vbCrLf
        gSL_Select = gSL_Select & "       MAX(CASE WHEN ISNULL(HTM1.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                     CASE WHEN ISNULL(TRM1.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                          THEN HTM1.取引先 + ':' + TRM1.略称" & vbCrLf
        gSL_Select = gSL_Select & "                          ELSE HTM1.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "                WHEN ISNULL(HTM2.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                     CASE WHEN ISNULL(TRM2.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                          THEN HTM2.取引先 + ':' + TRM2.略称" & vbCrLf
        gSL_Select = gSL_Select & "                          ELSE HTM2.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "                ELSE '' END) 手配先" & vbCrLf

        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert Start
        gSL_Select = gSL_Select & "      ,MAX(CASE WHEN ISNULL(HTM1.取引先, '') <> '' THEN HTM1.取引先" & vbCrLf
        gSL_Select = gSL_Select & "                WHEN ISNULL(HTM2.取引先, '') <> '' THEN HTM2.取引先" & vbCrLf
        gSL_Select = gSL_Select & "                ELSE '' END) 取引先" & vbCrLf
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert End
        gSL_Select = gSL_Select & "  FROM 構成問合せワーク KW " & vbCrLf
        '**********************************************2011/07/21 EDIT START**********************************************
       gSL_Select = gSL_Select & "   LEFT OUTER JOIN (Select Cast(Substring(WK1.キー,1,3) As Integer) 発注比率, Substring(WK1.キー,4,7) 取引先, WK1.品番"
       gSL_Select = gSL_Select & "                      From ("
       gSL_Select = gSL_Select & "                            Select MAX(Right('000' + Cast(発注比率 as Varchar(3)),3) + 取引先) キー,品番"
       gSL_Select = gSL_Select & "                              From 品番取引先マスタ"
       gSL_Select = gSL_Select & "                             WHERE ( 手配区分 <> '2' AND 手配区分 <> '4' ) AND 品目 = 1"
       gSL_Select = gSL_Select & "                             Group By 品番"
       gSL_Select = gSL_Select & "                           ) WK1"
       gSL_Select = gSL_Select & "                   ) HTM1"
       gSL_Select = gSL_Select & "     ON HTM1.品番 = KW.親品番"
         
       gSL_Select = gSL_Select & "   LEFT OUTER JOIN (Select Cast(Substring(WK2.キー,1,3) As Integer) 発注比率, Substring(WK2.キー,4,7) 取引先, WK2.品番"
       gSL_Select = gSL_Select & "                      From ("
       gSL_Select = gSL_Select & "                            Select MAX(Right('000' + Cast(発注比率 as Varchar(3)),3) + 取引先) キー,品番"
       gSL_Select = gSL_Select & "                              From 品番取引先マスタ"
       gSL_Select = gSL_Select & "                             WHERE 手配区分 = '2' AND 品目 = 1"
       gSL_Select = gSL_Select & "                             Group By 品番"
       gSL_Select = gSL_Select & "                           ) WK2"
       gSL_Select = gSL_Select & "                   ) HTM2"
       gSL_Select = gSL_Select & "     ON HTM2.品番 = KW.親品番"
'''        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 品番" & vbCrLf
'''        gSL_Select = gSL_Select & "                          FROM 品番取引先マスタ" & vbCrLf
''''======================= 2011/04/04 Start
''''        gSL_Select = gSL_Select & "                         WHERE 手配区分 <> '2' AND 品目 = 1 GROUP BY 品番) HTM1" & vbCrLf
'''        gSL_Select = gSL_Select & "                         WHERE ( 手配区分 <> '2' AND 手配区分 <> '4' ) AND 品目 = 1 GROUP BY 品番) HTM1" & vbCrLf
''''======================= 2011/04/04 End
'''        gSL_Select = gSL_Select & "            ON HTM1.品番 = KW.親品番" & vbCrLf
'''        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 品番" & vbCrLf
'''        gSL_Select = gSL_Select & "                          FROM 品番取引先マスタ" & vbCrLf
'''        gSL_Select = gSL_Select & "                         WHERE 手配区分 = '2' AND 品目 = 1 GROUP BY 品番) HTM2" & vbCrLf
'''        gSL_Select = gSL_Select & "            ON HTM2.品番 = KW.親品番" & vbCrLf
        '**********************************************2011/07/21 EDIT END**********************************************
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM1" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM1.取引先 = TRM1.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM1.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM2" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM2.取引先 = TRM2.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM2.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & " WHERE KW.処理端末 = HOST_NAME()" & vbCrLf
        gSL_Select = gSL_Select & " GROUP BY KW.親品番,KW.品名,KW.表示品番" & vbCrLf
        gSL_Select = gSL_Select & " ORDER BY KW.親品番" & vbCrLf
    End If
    If wsHKbn = 2 Then
        gSL_Select = "SELECT KW.親品番,KW.品名,KW.表示寸法,KW.表示品番," & vbCrLf
        gSL_Select = gSL_Select & "       MAX(CASE WHEN ISNULL(HTM1.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                     CASE WHEN ISNULL(TRM1.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                          THEN HTM1.取引先 + ':' + TRM1.略称" & vbCrLf
        gSL_Select = gSL_Select & "                          ELSE HTM1.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "                WHEN ISNULL(HTM2.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                     CASE WHEN ISNULL(TRM2.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                          THEN HTM2.取引先 + ':' + TRM2.略称" & vbCrLf
        gSL_Select = gSL_Select & "                          ELSE HTM2.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "                ELSE '' END) 手配先" & vbCrLf

        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert Start
        gSL_Select = gSL_Select & "      ,MAX(CASE WHEN ISNULL(HTM1.取引先, '') <> '' THEN HTM1.取引先" & vbCrLf
        gSL_Select = gSL_Select & "                WHEN ISNULL(HTM2.取引先, '') <> '' THEN HTM2.取引先" & vbCrLf
        gSL_Select = gSL_Select & "                ELSE '' END) 取引先" & vbCrLf
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert End

        gSL_Select = gSL_Select & "  FROM 構成問合せワーク KW "
        '**********************************************2011/07/21 EDIT START**********************************************
       gSL_Select = gSL_Select & "   LEFT OUTER JOIN (Select Cast(Substring(WK1.キー,1,3) As Integer) 発注比率, Substring(WK1.キー,4,7) 取引先, WK1.品番"
       gSL_Select = gSL_Select & "                      From ("
       gSL_Select = gSL_Select & "                            Select MAX(Right('000' + Cast(発注比率 as Varchar(3)),3) + 取引先) キー,品番"
       gSL_Select = gSL_Select & "                              From 品番取引先マスタ"
       gSL_Select = gSL_Select & "                             WHERE ( 手配区分 <> '2' AND 手配区分 <> '4' ) AND 品目 = 1"
       gSL_Select = gSL_Select & "                             Group By 品番"
       gSL_Select = gSL_Select & "                           ) WK1"
       gSL_Select = gSL_Select & "                   ) HTM1"
       gSL_Select = gSL_Select & "     ON HTM1.品番 = KW.親品番"
         
       gSL_Select = gSL_Select & "   LEFT OUTER JOIN (Select Cast(Substring(WK2.キー,1,3) As Integer) 発注比率, Substring(WK2.キー,4,7) 取引先, WK2.品番"
       gSL_Select = gSL_Select & "                      From ("
       gSL_Select = gSL_Select & "                            Select MAX(Right('000' + Cast(発注比率 as Varchar(3)),3) + 取引先) キー,品番"
       gSL_Select = gSL_Select & "                              From 品番取引先マスタ"
       gSL_Select = gSL_Select & "                             WHERE 手配区分 = '2' AND 品目 = 1"
       gSL_Select = gSL_Select & "                             Group By 品番"
       gSL_Select = gSL_Select & "                           ) WK2"
       gSL_Select = gSL_Select & "                   ) HTM2"
       gSL_Select = gSL_Select & "     ON HTM2.品番 = KW.親品番"
'''        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 品番" & vbCrLf
'''        gSL_Select = gSL_Select & "                          FROM 品番取引先マスタ" & vbCrLf
''''======================= 2011/04/04 Start
''''        gSL_Select = gSL_Select & "                         WHERE 手配区分 <> '2' AND 品目 = 1 GROUP BY 品番) HTM1" & vbCrLf
'''        gSL_Select = gSL_Select & "                         WHERE ( 手配区分 <> '2' AND 手配区分 <> '4' )  AND 品目 = 1 GROUP BY 品番) HTM1" & vbCrLf
''''======================= 2011/04/04 End
'''        gSL_Select = gSL_Select & "            ON HTM1.品番 = KW.親品番" & vbCrLf
'''        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 品番" & vbCrLf
'''        gSL_Select = gSL_Select & "                          FROM 品番取引先マスタ" & vbCrLf
'''        gSL_Select = gSL_Select & "                         WHERE 手配区分 = '2' AND 品目 = 1 GROUP BY 品番) HTM2" & vbCrLf
'''        gSL_Select = gSL_Select & "            ON HTM2.品番 = KW.親品番" & vbCrLf
        '**********************************************2011/07/21 EDIT END**********************************************
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM1" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM1.取引先 = TRM1.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM1.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM2" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM2.取引先 = TRM2.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM2.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & " WHERE KW.処理端末 = HOST_NAME() " & vbCrLf
        gSL_Select = gSL_Select & " GROUP BY KW.親品番,KW.品名,KW.表示寸法,KW.表示品番" & vbCrLf
        gSL_Select = gSL_Select & " ORDER BY KW.表示寸法,KW.親品番" & vbCrLf
    End If

    If Not KOLRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     構成マスタ情報が存在しません。     "
            .MB_Title = "抽出処理"
            .MB_Button = Error
            If .MBOX Then Exit Function
        End With
    End If

' 抽出処理 初期ｾｯﾄ
    RecCnt = CisDB.RecordCount       ' 抽出件数(合計)退避
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
    Do Until Not KOL_RDSTS
        ID = ID + 1
                
            If wsHKbn = 0 Then: Call KOMInfoToGrid0(ID)     ' 正展開
            If wsHKbn = 1 Then: Call KOMInfoToGrid1(ID)     ' 逆展開・品番
            If wsHKbn = 2 Then: Call KOMInfoToGrid2(ID)     ' 逆展開・材料
            
            Call KOLReadNext
        Loop
    
    Call KOLClose
    
    With VsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    
        Call GridSet_TEH(.Row)
        '-----2011/07/21 ADD START-----
        Call GridSet_NOU(.Row)
        '-----2011/07/21 ADD END-------
        lb_Kensu = Format(.Row, "#,###") & " / " & Format(RecCnt, "#,###")
    End With

    GridSet_KOM = True
End Function
'+---------------------------------------------------------------------------------------+
'+      構成情報よりグリッドへセット(正展開)
'+---------------------------------------------------------------------------------------+
Private Sub KOMInfoToGrid0(SetRow As Long)
    With VsGrid1
        
        ' No
        .TextMatrix(SetRow, 0) = SetRow
        
        ' レベル
        Select Case KOL.レベル
            Case 0: .TextMatrix(SetRow, SCol(1)) = "0"
            Case 1: .TextMatrix(SetRow, SCol(1)) = ".1"
            Case 2: .TextMatrix(SetRow, SCol(1)) = "..2"
            Case 3: .TextMatrix(SetRow, SCol(1)) = "...3"
            Case 4: .TextMatrix(SetRow, SCol(1)) = "....4"
            Case Else
                .TextMatrix(SetRow, SCol(1)) = "....." & Format(KOL.レベル, "##")
        End Select
        ' 区分/品番
        If KOL.区分 = "" Then
            .TextMatrix(SetRow, SCol(2)) = ""
            .TextMatrix(SetRow, SCol(14)) = KOL.子品番 'Update 2012/09/20
            '.TextMatrix(SetRow, SCol(13)) = KOL.子品番 'Update 2011/02/28
            '.TextMatrix(SetRow, SCol(11)) = KOL.子品番
        Else
            .TextMatrix(SetRow, SCol(2)) = "材"
            .TextMatrix(SetRow, SCol(14)) = KOL.材料管理番号 'Update 2012/09/20
            '.TextMatrix(SetRow, SCol(13)) = KOL.材料管理番号 'Update 2011/02/28
            '.TextMatrix(SetRow, SCol(11)) = KOL.材料管理番号
        End If
        ' 品番
        .TextMatrix(SetRow, SCol(3)) = KOL.表示品番
        ' 品名
        .TextMatrix(SetRow, SCol(4)) = KOL.品名
        
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert Start
        Select Case wsHKbn
            Case 0
                If KOL.区分 = "" Then
                    gSL_Select = ""
                    gSL_Select = gSL_Select & "SELECT HT1.背番号"
                    gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(HT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= HT1.変更日 THEN HT1.収容数2 ELSE HT1.収容数1 END AS 収容数1"
                    gSL_Select = gSL_Select & "  FROM 品番取引先マスタ HT1"
                    gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                    gSL_Select = gSL_Select & "               FROM 品番取引先マスタ"
                    gSL_Select = gSL_Select & "              WHERE 取引先 = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "                AND 品番   = '" & KOL.子品番 & "') HT2"
                    gSL_Select = gSL_Select & "       ON  HT1.取引先 = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "       AND HT1.品番   = '" & KOL.子品番 & "'"
                    gSL_Select = gSL_Select & "       AND HT1.受入   = HT2.受入"
                Else
                    gSL_Select = ""
                    gSL_Select = gSL_Select & "SELECT ZT1.背番号"
                    gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(ZT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= ZT1.変更日 THEN ZT1.収容数2 ELSE ZT1.収容数1 END AS 収容数1"
                    gSL_Select = gSL_Select & "  FROM 材料取引先マスタ ZT1"
                    gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                    gSL_Select = gSL_Select & "               FROM 材料取引先マスタ"
                    gSL_Select = gSL_Select & "              WHERE 取引先       = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "                AND 材料管理番号 =  " & KOL.材料管理番号 & ") ZT2"
                    gSL_Select = gSL_Select & "       ON  ZT1.取引先       = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "       AND ZT1.材料管理番号 = '" & KOL.材料管理番号 & "'"
                    gSL_Select = gSL_Select & "       AND ZT1.受入         = ZT2.受入"
                End If
            Case 1, 2
                gSL_Select = ""
                gSL_Select = gSL_Select & "SELECT HT1.背番号"
                gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(HT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= HT1.変更日 THEN HT1.収容数2 ELSE HT1.収容数1 END AS 収容数1"
                gSL_Select = gSL_Select & "  FROM 品番取引先マスタ HT1"
                gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                gSL_Select = gSL_Select & "               FROM 品番取引先マスタ"
                gSL_Select = gSL_Select & "              WHERE 取引先 = '" & KOL.取引先 & "'"
                gSL_Select = gSL_Select & "                AND 品番   = '" & KOL.子品番 & "') HT2"
                gSL_Select = gSL_Select & "       ON  HT1.取引先 = '" & KOL.取引先 & "'"
                gSL_Select = gSL_Select & "       AND HT1.品番   = '" & KOL.子品番 & "'"
                gSL_Select = gSL_Select & "       AND HT1.受入   = HT2.受入"
        End Select
        
        HTM.背番号 = ""                             '2016/11/23
        Call HTMRead(gSL_Select, , 1)

        ' 背番号
        .TextMatrix(SetRow, SCol(5)) = HTM.背番号
        ' 収容数
        .TextMatrix(SetRow, SCol(6)) = Format(HTM.収容数1, "#,###")
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert End
        
        '=============== 2012/09/20 START
        If KOL.区分 = "1" Then
            gSL_Select = "SELECT * FROM 材料品番マスタ "
            gSL_Select = gSL_Select & " WHERE 材料管理番号 = " & KOL.材料管理番号
            gSL_Select = gSL_Select & " AND   使用品番 = '" & RTrim(KOL.親品番) & "'"
            If ZHMRead(gSL_Select, 1, 1) Then
                If ZHM.ピッチ <> 0 Then
                    .TextMatrix(SetRow, SCol(13)) = CisFun.RSetFld(ZHM.ピッチ, 10, "#,##0.0000")
                Else
                    .TextMatrix(SetRow, SCol(13)) = ""
                End If
            Else
                .TextMatrix(SetRow, SCol(13)) = ""
            End If
            
        End If
        '=============== 2012/09/20 END
        
        
        ' 使用数
        .TextMatrix(SetRow, SCol(7)) = CisFun.RSetFld(KOL.構成数, 10, "#,##0.0000") 'Update 2011/02/28
        '.TextMatrix(SetRow, SCol(5)) = CisFun.RSetFld(KOL.構成数, 10, "#,##0.0000")
        
        ' 号補区分
        .TextMatrix(SetRow, SCol(8)) = ""            'Update 2012/07/03
'        .TextMatrix(SetRow, SCol(8)) = KOL.号補区分 'Update 2011/02/28
        '.TextMatrix(SetRow, SCol(6)) = KOL.号補区分
        
        ' 展開区分 Update 2011/02/28
        Select Case KOL.展開区分
            Case 0: .TextMatrix(SetRow, SCol(9)) = ""
            Case 1: .TextMatrix(SetRow, SCol(9)) = "まで対象"
            Case 2: .TextMatrix(SetRow, SCol(9)) = "以降対象外"
            Case 3: .TextMatrix(SetRow, SCol(9)) = "のみ対象外"
        End Select
'        Select Case KOL.展開区分
'            Case 0: .TextMatrix(SetRow, SCol(7)) = ""
'            Case 1: .TextMatrix(SetRow, SCol(7)) = "まで対象"
'            Case 2: .TextMatrix(SetRow, SCol(7)) = "以降対象外"
'            Case 3: .TextMatrix(SetRow, SCol(7)) = "のみ対象外"
'        End Select

        ' 有効期限
        gStr = ""
        If Trim(KOL.適用開始日) <> "" Then
            gStr = Mid(KOL.適用開始日, 1, 4) & "/" & _
                   Mid(KOL.適用開始日, 5, 2) & "/" & _
                   Mid(KOL.適用開始日, 7, 2)
        End If
        If Trim(KOL.適用終了日) <> "" Then
           If gStr = "" Then gStr = Space(10)
           gStr = gStr & "～"
           gStr = gStr & Mid(KOL.適用終了日, 1, 4) & "/" & _
                         Mid(KOL.適用終了日, 5, 2) & "/" & _
                         Mid(KOL.適用終了日, 7, 2)
        Else
           If Trim(KOL.適用開始日) <> "" Then
              gStr = gStr & "～" & Space(10)
           End If
        End If
        .TextMatrix(SetRow, SCol(10)) = gStr 'Update 2011/02/28
        '.TextMatrix(SetRow, SCol(8)) = gStr
        
        ' 寸法
        .TextMatrix(SetRow, SCol(11)) = RTrim(KOL.表示寸法) 'Update 2011/02/28
        '.TextMatrix(SetRow, SCol(9)) = RTrim(KOL.表示寸法)

'=========================== 2013/03/15 Start
        ' 手配先
'        .TextMatrix(SetRow, SCol(12)) = RTrim(KOL.手配先) 'Update 2011/02/28
        '.TextMatrix(SetRow, SCol(10)) = RTrim(KOL.手配先)
    
        .TextMatrix(SetRow, SCol(5)) = HTM.背番号
        If KOL.区分 = "" Then
            gSL_Select = "SELECT 連番 FROM CIS8606"
            gSL_Select = gSL_Select & " WHERE REPLACE(品番,'-','') = '" & Replace(RTrim(iH1_Hinbn), "-", "") & "'"
            gSL_Select = gSL_Select & " ORDER BY CONVERT(INT,連番) DESC "
            If CISRead(gSL_Select, 1, 2) Then
                gSL_Select = "SELECT TOP 100 * FROM CIS8606"
                gSL_Select = gSL_Select & " WHERE CONVERT(INT,連番) >= " & CisFun.Val2(CIS.連番)
                gSL_Select = gSL_Select & " ORDER BY CONVERT(INT,連番) "
                If CISRead(gSL_Select, , 2) Then
                    wFlg = False
                    Do Until Not CIS_RDSTS
                        If Replace(CIS.品番, "-", "") = Replace(KOL.子品番, "-", "") Then
                            wFlg = True
                            Call TorNmGet(RTrim(CIS.手配先コード) & RTrim(CIS.手配先工場コード), 9, 1)
                            If RTrim(TRM.略称) <> "" Then
                                'gSL_Select = "SELECT 取引先 FROM 品番取引先マスタ "
                                gSL_Select = "SELECT 取引先,背番号,収容数1 FROM 品番取引先マスタ "
                                gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(KOL.子品番) & "'"
                                gSL_Select = gSL_Select & " AND 取引先 = '" & RTrim(CIS.手配先コード) & RTrim(CIS.手配先工場コード) & "'"
                                
                                '================= 2016/11/23 start
                                If RTrim(CIS.受入発注) <> "" Then
                                    gSL_Select = gSL_Select & " AND 受入 = '" & RTrim(CIS.受入発注) & "'"
                                Else
                                    gSL_Select = gSL_Select & " AND 受入 = '01'"
                                End If
                                '================= 2016/11/23 end
                                
                                gSL_Select = gSL_Select & " AND 品目 = 1 "
                                If HTMRead(gSL_Select, 1, 1) Then
                                    .TextMatrix(SetRow, SCol(12)) = RTrim(CIS.手配先コード) & RTrim(CIS.手配先工場コード) & ":" & RTrim(TRM.略称)
                                    
                                    '================= 2016/11/23 start
                                    If RTrim(CIS.SD発注) <> "" Then
                                        .TextMatrix(SetRow, SCol(5)) = RTrim(HTM.背番号)
                                    End If
                                    If RTrim(CIS.収容数発注) <> "" Then
                                        .TextMatrix(SetRow, SCol(6)) = Format(HTM.収容数1, "#,###")
                                    End If
                                    '================= 2016/11/23 end
                                
                                Else
                                    .TextMatrix(SetRow, SCol(12)) = RTrim(KOL.手配先)
                                End If
                            Else
                                .TextMatrix(SetRow, SCol(12)) = RTrim(KOL.手配先)
                            End If
                            Exit Do
                        End If
                        Call CISReadNext(2)
                    Loop
                    Call CISClose(2)
                    If wFlg = False Then
                        .TextMatrix(SetRow, SCol(12)) = RTrim(KOL.手配先)
                    End If
                Else
                    .TextMatrix(SetRow, SCol(12)) = RTrim(KOL.手配先)
                End If
            Else
                .TextMatrix(SetRow, SCol(12)) = RTrim(KOL.手配先)
            End If
        Else
            .TextMatrix(SetRow, SCol(12)) = RTrim(KOL.手配先)
        End If
'=========================== 2013/03/15 End
            
    
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+      構成情報よりグリッドへセット(逆展開・品番)
'+---------------------------------------------------------------------------------------+
Private Sub KOMInfoToGrid1(SetRow As Long)
    With VsGrid1
        
        ' No
        .TextMatrix(SetRow, 0) = SetRow
        
        ' 品番
        .TextMatrix(SetRow, SCol(3)) = KOL.表示品番
        .TextMatrix(SetRow, SCol(14)) = KOL.親品番 'Update 2012/09/20
        '.TextMatrix(SetRow, SCol(13)) = KOL.親品番 'Update 2011/02/28
        '.TextMatrix(SetRow, SCol(11)) = KOL.親品番
        ' 品名
        .TextMatrix(SetRow, SCol(4)) = KOL.品名
        ' 手配先
        .TextMatrix(SetRow, SCol(12)) = RTrim(KOL.手配先) 'Update 2011/02/28
        '.TextMatrix(SetRow, SCol(10)) = RTrim(KOL.手配先)
        
        '-----2011/07/21 ADD START-----
        Select Case wsHKbn
            Case 0
                If KOL.区分 = "" Then
                    gSL_Select = ""
                    gSL_Select = gSL_Select & "SELECT HT1.背番号"
                    gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(HT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= HT1.変更日 THEN HT1.収容数2 ELSE HT1.収容数1 END AS 収容数1"
                    gSL_Select = gSL_Select & "  FROM 品番取引先マスタ HT1"
                    gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                    gSL_Select = gSL_Select & "               FROM 品番取引先マスタ"
                    gSL_Select = gSL_Select & "              WHERE 取引先 = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "                AND 品番   = '" & KOL.子品番 & "') HT2"
                    gSL_Select = gSL_Select & "       ON  HT1.取引先 = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "       AND HT1.品番   = '" & KOL.子品番 & "'"
                    gSL_Select = gSL_Select & "       AND HT1.受入   = HT2.受入"
                Else
                    gSL_Select = ""
                    gSL_Select = gSL_Select & "SELECT ZT1.背番号"
                    gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(ZT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= ZT1.変更日 THEN ZT1.収容数2 ELSE ZT1.収容数1 END AS 収容数1"
                    gSL_Select = gSL_Select & "  FROM 材料取引先マスタ ZT1"
                    gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                    gSL_Select = gSL_Select & "               FROM 材料取引先マスタ"
                    gSL_Select = gSL_Select & "              WHERE 取引先       = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "                AND 材料管理番号 =  " & KOL.材料管理番号 & ") ZT2"
                    gSL_Select = gSL_Select & "       ON  ZT1.取引先       = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "       AND ZT1.材料管理番号 = '" & KOL.材料管理番号 & "'"
                    gSL_Select = gSL_Select & "       AND ZT1.受入         = ZT2.受入"
                End If
            Case 1, 2
                gSL_Select = ""
                gSL_Select = gSL_Select & "SELECT HT1.背番号"
                gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(HT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= HT1.変更日 THEN HT1.収容数2 ELSE HT1.収容数1 END AS 収容数1"
                gSL_Select = gSL_Select & "  FROM 品番取引先マスタ HT1"
                gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                gSL_Select = gSL_Select & "               FROM 品番取引先マスタ"
                gSL_Select = gSL_Select & "              WHERE 取引先 = '" & KOL.取引先 & "'"
                gSL_Select = gSL_Select & "                AND 品番   = '" & KOL.親品番 & "') HT2"
                gSL_Select = gSL_Select & "       ON  HT1.取引先 = '" & KOL.取引先 & "'"
                gSL_Select = gSL_Select & "       AND HT1.品番   = '" & KOL.親品番 & "'"
                gSL_Select = gSL_Select & "       AND HT1.受入   = HT2.受入"
        End Select
        
        Call HTMRead(gSL_Select, , 1)

        ' 背番号
        .TextMatrix(SetRow, SCol(5)) = HTM.背番号
        ' 収容数
        .TextMatrix(SetRow, SCol(6)) = Format(HTM.収容数1, "#,###")
        '-----2011/07/21 ADD END-------
        
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+      構成情報よりグリッドへセット(逆展開・材料)
'+---------------------------------------------------------------------------------------+
Private Sub KOMInfoToGrid2(SetRow As Long)
    With VsGrid1
        
        ' No
        .TextMatrix(SetRow, 0) = SetRow
        
        ' 品番
        .TextMatrix(SetRow, SCol(3)) = KOL.表示品番
        .TextMatrix(SetRow, SCol(14)) = KOL.親品番 'Update 2012/09/20
        '.TextMatrix(SetRow, SCol(13)) = KOL.親品番 'Update 2011/02/28
        '.TextMatrix(SetRow, SCol(11)) = KOL.親品番
        ' 品名
        .TextMatrix(SetRow, SCol(4)) = KOL.品名
        ' 表示寸法
        .TextMatrix(SetRow, SCol(11)) = KOL.表示寸法 'Update 2011/02/28
        '.TextMatrix(SetRow, SCol(9)) = KOL.表示寸法
        ' 手配先
        .TextMatrix(SetRow, SCol(12)) = RTrim(KOL.手配先) 'Update 2011/02/28
        '.TextMatrix(SetRow, SCol(10)) = RTrim(KOL.手配先)
        
        '-----2011/07/21 ADD START-----
        Select Case wsHKbn
            Case 0
                If KOL.区分 = "" Then
                    gSL_Select = ""
                    gSL_Select = gSL_Select & "SELECT HT1.背番号"
                    gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(HT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= HT1.変更日 THEN HT1.収容数2 ELSE HT1.収容数1 END AS 収容数1"
                    gSL_Select = gSL_Select & "  FROM 品番取引先マスタ HT1"
                    gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                    gSL_Select = gSL_Select & "               FROM 品番取引先マスタ"
                    gSL_Select = gSL_Select & "              WHERE 取引先 = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "                AND 品番   = '" & KOL.子品番 & "') HT2"
                    gSL_Select = gSL_Select & "       ON  HT1.取引先 = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "       AND HT1.品番   = '" & KOL.子品番 & "'"
                    gSL_Select = gSL_Select & "       AND HT1.受入   = HT2.受入"
                Else
                    gSL_Select = ""
                    gSL_Select = gSL_Select & "SELECT ZT1.背番号"
                    gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(ZT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= ZT1.変更日 THEN ZT1.収容数2 ELSE ZT1.収容数1 END AS 収容数1"
                    gSL_Select = gSL_Select & "  FROM 材料取引先マスタ ZT1"
                    gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                    gSL_Select = gSL_Select & "               FROM 材料取引先マスタ"
                    gSL_Select = gSL_Select & "              WHERE 取引先       = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "                AND 材料管理番号 =  " & KOL.材料管理番号 & ") ZT2"
                    gSL_Select = gSL_Select & "       ON  ZT1.取引先       = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "       AND ZT1.材料管理番号 = '" & KOL.材料管理番号 & "'"
                    gSL_Select = gSL_Select & "       AND ZT1.受入         = ZT2.受入"
                End If
            Case 1, 2
                gSL_Select = ""
                gSL_Select = gSL_Select & "SELECT HT1.背番号"
                gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(HT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= HT1.変更日 THEN HT1.収容数2 ELSE HT1.収容数1 END AS 収容数1"
                gSL_Select = gSL_Select & "  FROM 品番取引先マスタ HT1"
                gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                gSL_Select = gSL_Select & "               FROM 品番取引先マスタ"
                gSL_Select = gSL_Select & "              WHERE 取引先 = '" & KOL.取引先 & "'"
                gSL_Select = gSL_Select & "                AND 品番   = '" & KOL.親品番 & "') HT2"
                gSL_Select = gSL_Select & "       ON  HT1.取引先 = '" & KOL.取引先 & "'"
                gSL_Select = gSL_Select & "       AND HT1.品番   = '" & KOL.親品番 & "'"
                gSL_Select = gSL_Select & "       AND HT1.受入   = HT2.受入"
        End Select
        
        Call HTMRead(gSL_Select, , 1)

        ' 背番号
        .TextMatrix(SetRow, SCol(5)) = HTM.背番号
        ' 収容数
        .TextMatrix(SetRow, SCol(6)) = Format(HTM.収容数1, "#,###")
        '-----2011/07/21 ADD END-------
        
    End With
End Sub
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
                '-----2011/07/21 ADD START-----
                Call GridSet_NOU(.Row)
                '-----2011/07/21 ADD END-------
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
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

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
Private Sub VSGrid1_DragDrop(Source As Control, x As Single, Y As Single)
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
        
        SCol(10) = CisVsGrid1.FixedGet(11)     ' 寸法 Update 2011/02/28
        'SCol(8) = CisVsGrid1.FixedGet(9)    ' 寸法
        
' ｸﾞﾘｯﾄﾞ表示 解除
        .Redraw = True
        CisVsGrid1.ReNumber
            
'< 品番取引先情報セット >
        Call GridSet_TEH(.Row)
        '-----2011/07/21 ADD START-----
        Call GridSet_NOU(.Row)
        '-----2011/07/21 ADD END-------
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
                    mn_Copy.Tag = RTrim$(.TextMatrix(.MouseRow, SCol(14))) 'Update 2012/09/20
                    'mn_Copy.Tag = RTrim$(.TextMatrix(.MouseRow, SCol(13))) 'Update 2011/02/28
                    'mn_Copy.Tag = RTrim$(.TextMatrix(.MouseRow, SCol(11)))
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
Private Sub mn_Gyaku_Click(Index As Integer)
    Call Now_Save
    
    ProcHB = "H1"
    Call DispChange(ProcHB)
    Call HeadBodyClear("B")

    H1Op_Hinkb(w_ZaiKB).Value = True
    H1Op_Tenkb(1).Value = True
    H1Op_Oyakb(Index).Value = True
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
        sv_Zaigy(GetB_Cnt) = .FixedGet(11) & .FixedGet(3) & .FixedGet(4) 'Update 2011/02/28
        'sv_Zaigy(GetB_Cnt) = .FixedGet(9) & .FixedGet(3) & .FixedGet(4)
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
''        Call .FixedSet(中中, 左中, あり, 7, 3, "背番号")              '2015/01/21
        Call .FixedSet(中中, 左中, あり, 8, 3, "背番号")                '2015/01/21
        Call .FixedSet(中中, 左中, あり, 5, 4, "手配")
        Call .FixedSet(中中, 右中, あり, 4, 5, "比率")
        Call .FixedSet(中中, 中中, あり, 3, 6, "SK")
        Call .FixedSet(中中, 左中, あり, 8, 7, "ｻｲｸﾙ")
        Call .FixedSet(中中, 左中, あり, 20, 8, "有効期限")
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
'-----2011/07/21 ADD START-----
'+----------------------------------------------------------+
'+   納入先情報   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInitN()
    With CisVsGridN
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        ReDim SColN(10)

        Call .FixedSet(中中, 右中, あり, 3, 0, "No")
        Call .FixedSet(中中, 左中, あり, 32, 1, "納入先")
        Call .FixedSet(中中, 中中, あり, 2, 2, "受")
        Call .FixedSet(中中, 左中, あり, 8, 3, "背番号")
        Call .FixedSet(中中, 左中, あり, 32, 4, "契約先")
        Call .FixedSet(中中, 中中, あり, 3, 5, "SK")
        Call .FixedSet(中中, 左中, あり, 8, 6, "ｻｲｸﾙ")
        Call .FixedSet(中中, 右中, あり, 12, 7, "収容数")
        Call .FixedSet(中中, 左中, あり, 12, 8, "収容器")

'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, あり, 0, 9, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp

        SColN(1) = .FixedGet(1)     ' 納入先
        SColN(2) = .FixedGet(2)     ' 受入
        SColN(3) = .FixedGet(3)     ' 背番号
        SColN(4) = .FixedGet(4)     ' 契約先
        SColN(5) = .FixedGet(5)     ' ＳＫ
        SColN(6) = .FixedGet(6)     ' サイクル
        SColN(7) = .FixedGet(7)     ' 収容数
        SColN(8) = .FixedGet(8)     ' 収容器

    End With
    KCol = 9   ' 非表示項目開始ｶﾗﾑ

End Sub
'-----2011/07/21 ADD END-------
'+---------------------------------------------------------------------------------------+
'+      グリッド内容セット
'+---------------------------------------------------------------------------------------+
Private Sub GridSet_TEH(pRow As Long)
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
    Static wsKB     As String
        
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
            gSL_Select = gSL_Select & "  AND TR.取引先区分 = HT.品目"
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY "
            gSL_Select = gSL_Select & "   ON SY.区分名称 = '手配区分' "
            gSL_Select = gSL_Select & "  AND SY.値 = HT.手配区分"
            gSL_Select = gSL_Select & "  AND SY.区分タイプ = ''"
            gSL_Select = gSL_Select & " WHERE HT.品番 = '" & RTrim(.TextMatrix(pRow, SCol(14))) & "'" 'Update 2012/09/20
            'gSL_Select = gSL_Select & " WHERE HT.品番 = '" & RTrim(.TextMatrix(pRow, SCol(13))) & "'" 'Update 2011/02/28
            'gSL_Select = gSL_Select & " WHERE HT.品番 = '" & RTrim(.TextMatrix(pRow, SCol(11))) & "'"
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/06/18 INSERT START
            gSL_Select = gSL_Select & "  AND HT.品目 = 1"
'▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△▽△ 2009/06/18 INSERT END
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
            gSL_Select = gSL_Select & " WHERE ZR.材料管理番号 = " & .TextMatrix(pRow, SCol(14)) 'Update 2012/09/20
            'gSL_Select = gSL_Select & " WHERE ZR.材料管理番号 = " & .TextMatrix(pRow, SCol(13)) 'Update 2011/02/28
            'gSL_Select = gSL_Select & " WHERE ZR.材料管理番号 = " & .TextMatrix(pRow, SCol(11))
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
'============== 2013/03/15 Start
'                    .TextMatrix(SetRow, SColT(7)) = Mid(HTM.サイクル, 1, 1) & "-" & _
'                                                    Mid(HTM.サイクル, 2, 2) & "-" & _
'                                                    Mid(HTM.サイクル, 4, 2)
                    .TextMatrix(SetRow, SColT(7)) = Mid(ZTM.サイクル, 1, 1) & "-" & _
                                                    Mid(ZTM.サイクル, 2, 2) & "-" & _
                                                    Mid(ZTM.サイクル, 4, 2)
'============== 2013/03/15 End
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
'-----2011/07/21 ADD START-----
'+---------------------------------------------------------------------------------------+
'+      グリッド内容セット
'+---------------------------------------------------------------------------------------+
Private Sub GridSet_NOU(pRow As Long)
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
    Static wsKB     As String
        
    Call CisVsGridN.Clear
    Call GridInitN

    With VsGrid1
        wsKB = RTrim(.TextMatrix(pRow, SCol(2)))
        If wsKB = "" Then
            gSL_Select = ""
            gSL_Select = gSL_Select & "SELECT HT.*," & vbCrLf
            gSL_Select = gSL_Select & "       TR.略称 取引先名," & vbCrLf
            gSL_Select = gSL_Select & "       HN.品名," & vbCrLf
            gSL_Select = gSL_Select & "       TR2.略称 契約先名" & vbCrLf
            gSL_Select = gSL_Select & "  FROM 品番取引先マスタ HT " & vbCrLf
            gSL_Select = gSL_Select & "  LEFT OUTER JOIN 品番マスタ HN " & vbCrLf
            gSL_Select = gSL_Select & "    ON HN.品番 = HT.品番 " & vbCrLf
            gSL_Select = gSL_Select & "  LEFT OUTER JOIN 取引先マスタ TR " & vbCrLf
            gSL_Select = gSL_Select & "    ON TR.取引先CD = HT.取引先 " & vbCrLf
            gSL_Select = gSL_Select & "   AND TR.取引先区分 = 0" & vbCrLf
            gSL_Select = gSL_Select & "  LEFT OUTER JOIN 取引先マスタ TR2" & vbCrLf
            gSL_Select = gSL_Select & "    ON TR2.取引先CD = HT.契約先" & vbCrLf
            gSL_Select = gSL_Select & "   AND TR2.取引先区分 = 0" & vbCrLf
            'gSL_Select = gSL_Select & " WHERE HT.品番 = '" & RTrim(.TextMatrix(pRow, SCol(13))) & "'" & vbCrLf         '2012/09/20
            gSL_Select = gSL_Select & " WHERE HT.品番 = '" & RTrim(.TextMatrix(pRow, SCol(14))) & "'" & vbCrLf          '2012/09/20
            gSL_Select = gSL_Select & "   AND HT.品目 = 0" & vbCrLf
            gSL_Select = gSL_Select & " ORDER BY HT.取引先" & vbCrLf
            Call HTMRead(gSL_Select)
        Else
'''            gSL_Select = ""
'''            gSL_Select = gSL_Select & "SELECT ZT.*," & vbCrLf
'''            gSL_Select = gSL_Select & "       TR.略称 手配先名," & vbCrLf
'''            gSL_Select = gSL_Select & "       ZR.材料名," & vbCrLf
'''            gSL_Select = gSL_Select & "       TR2.略称 契約先名" & vbCrLf
'''            gSL_Select = gSL_Select & "  FROM 材料マスタ ZR " & vbCrLf
'''            gSL_Select = gSL_Select & "  LEFT OUTER JOIN 材料取引先マスタ ZT " & vbCrLf
'''            gSL_Select = gSL_Select & "    ON ZT.材料管理番号 = ZR.材料管理番号 " & vbCrLf
'''            gSL_Select = gSL_Select & "  LEFT OUTER JOIN 取引先マスタ TR " & vbCrLf
'''            gSL_Select = gSL_Select & "    ON TR.取引先CD = ZT.取引先 " & vbCrLf
'''            gSL_Select = gSL_Select & "   AND TR.取引先区分 = 0 " & vbCrLf
'''            gSL_Select = gSL_Select & "  LEFT OUTER JOIN 取引先マスタ TR2" & vbCrLf
'''            gSL_Select = gSL_Select & "    ON TR2.取引先CD = ZT.契約先" & vbCrLf
'''            gSL_Select = gSL_Select & "   AND TR2.取引先区分 = 0" & vbCrLf
'''            gSL_Select = gSL_Select & " WHERE ZR.材料管理番号 = " & .TextMatrix(pRow, SCol(13)) & vbCrLf
'''            gSL_Select = gSL_Select & " ORDER BY ZT.取引先" & vbCrLf
'''            Call ZTMRead(gSL_Select)
        End If
    End With

    If CisDB.RecordCount <> 0 Then

    ' 抽出処理 初期ｾｯﾄ
    '    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
        Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
    ' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
        With VsGridN
           .Redraw = False
            If Ttl_Kensu + 1 < CisVsGridN.DispMax Then
                .Rows = CisVsGridN.DispMax
            Else
                .Rows = Ttl_Kensu + 1
            End If
        End With
        ID = 0

        If wsKB = "" Then
            Do Until Not HTM_RDSTS
                ID = ID + 1

                Call HTMInfoToGrid2(ID, wsKB)
                Call HTMReadNext
            Loop
        Else
            Do Until Not ZTM_RDSTS
                ID = ID + 1

                Call HTMInfoToGrid2(ID, wsKB)
                Call ZTMReadNext
            Loop
        End If

    End If
    If wsKB = "" Then
        Call HTMClose
    Else
        Call ZTMClose
    End If
    With VsGridN
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    
End Sub
'+---------------------------------------------------------------------------------------+
'+  構成情報よりグリッドへセット（納入先情報）
'+---------------------------------------------------------------------------------------+
Private Sub HTMInfoToGrid2(SetRow As Long, KB As String)
    With VsGridN
        ' No
        .TextMatrix(SetRow, 0) = SetRow
        If RTrim(KB) = "" Then
            If RTrim(HTM.取引先) <> "" Then
                ' 取引先
                .TextMatrix(SetRow, SColN(1)) = RTrim(HTM.取引先) & ":" & RTrim(HTM.取引先名)
                ' 受入
                .TextMatrix(SetRow, SColN(2)) = RTrim(HTM.受入)
                ' 背番号
                .TextMatrix(SetRow, SColN(3)) = RTrim(HTM.背番号)
                ' 契約先
                If Trim(HTM.契約先) = "" Then
                    .TextMatrix(SetRow, SColN(4)) = ""
                Else
                    
                    .TextMatrix(SetRow, SColT(4)) = RTrim(HTM.契約先) & ":" & RTrim(HTM.契約先名)
                    
'''                    If Trim(HTM.契約先名) = "" Then
'''                        .TextMatrix(SetRow, SColT(4)) = RTrim(HTM.契約先)
'''                    Else
'''                        .TextMatrix(SetRow, SColT(4)) = RTrim(HTM.契約先) & ":" & RTrim(HTM.契約先名)
'''                    End If
                    
                End If
                ' SK
                If HTM.SK区分 = 0 Then
                    .TextMatrix(SetRow, SColN(5)) = "K"
                Else
                    .TextMatrix(SetRow, SColN(5)) = "S"
                End If
                ' サイクル
                If RTrim(HTM.サイクル) <> "" Then
                    .TextMatrix(SetRow, SColN(6)) = Mid(HTM.サイクル, 1, 1) & "-" & _
                                                    Mid(HTM.サイクル, 2, 2) & "-" & _
                                                    Mid(HTM.サイクル, 4, 2)
                End If
                '収容数
                .TextMatrix(SetRow, SColN(7)) = Format(HTM.収容数1, "#,##0")
                '==================== 2015/08/17 Start
                If HTM.変更日 <> "" And HTM.変更日 < Format(Now, "YYYYMMDD") Then
                    .TextMatrix(SetRow, SColN(7)) = Format(HTM.収容数2, "#,##0")
                Else
                    .TextMatrix(SetRow, SColN(7)) = Format(HTM.収容数1, "#,##0")
                End If
                '==================== 2015/08/17 End
                '収容器
                .TextMatrix(SetRow, SColN(8)) = HTM.収容器1
                '==================== 2015/08/17 Start
                If HTM.変更日 <> "" And HTM.変更日 < Format(Now, "YYYYMMDD") Then
                    .TextMatrix(SetRow, SColN(8)) = Format(HTM.収容器2, "#,##0")
                Else
                    .TextMatrix(SetRow, SColN(8)) = Format(HTM.収容器1, "#,##0")
                End If
                '==================== 2015/08/17 End
                
            End If
        Else
            If RTrim(ZTM.取引先) <> "" Then
                ' 手配先
                .TextMatrix(SetRow, SColN(1)) = RTrim(ZTM.取引先) & ":" & RTrim(ZTM.手配先名)
                ' 受入
                .TextMatrix(SetRow, SColN(2)) = RTrim(ZTM.受入)
                ' 背番号
                .TextMatrix(SetRow, SColN(3)) = RTrim(ZTM.背番号)
                ' 契約先
                If Trim(ZTM.契約先) = "" Then
                    .TextMatrix(SetRow, SColT(4)) = ""
                Else
                    
                    .TextMatrix(SetRow, SColN(4)) = RTrim(ZTM.契約先) & ":" & RTrim(ZTM.契約先名)
                    
'''                    If Trim(ZTM.契約先名) = "" Then
'''                        .TextMatrix(SetRow, SColT(4)) = RTrim(ZTM.契約先)
'''                    Else
'''                        .TextMatrix(SetRow, SColT(4)) = RTrim(ZTM.契約先) & ":" & RTrim(ZTM.契約先名)
'''                    End If
                    
                End If
                ' SK
                If ZTM.SK区分 = 0 Then
                    .TextMatrix(SetRow, SColN(5)) = "K"
                Else
                    .TextMatrix(SetRow, SColN(5)) = "S"
                End If
                ' サイクル
                If RTrim(ZTM.サイクル) <> "" Then
                    .TextMatrix(SetRow, SColN(6)) = Mid(ZTM.サイクル, 1, 1) & "-" & _
                                                    Mid(ZTM.サイクル, 2, 2) & "-" & _
                                                    Mid(ZTM.サイクル, 4, 2)
                End If
                '収容数
                .TextMatrix(SetRow, SColN(7)) = Format(ZTM.収容数1, "#,##0")
                '収容器
                .TextMatrix(SetRow, SColN(8)) = ZTM.収容器1
                
            End If
        End If
    End With
End Sub
