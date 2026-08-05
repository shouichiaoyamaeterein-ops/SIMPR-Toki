VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form CXM2540 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "スクラップ単価変動額問合せ(支払)"
   ClientHeight    =   11190
   ClientLeft      =   1350
   ClientTop       =   3615
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
   Begin VB.PictureBox H1_Area1 
      Height          =   1365
      Left            =   255
      ScaleHeight     =   1305
      ScaleWidth      =   9120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   585
      Width           =   9180
      Begin Cis3D_v60.CIS3D H1_TorMsg 
         Height          =   1185
         Left            =   3165
         Top             =   60
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   2090
         ForeColor       =   16711680
         Caption         =   "支　払　先"
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
         cPositionY      =   50
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   105
            TabIndex        =   3
            Top             =   585
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
            Top             =   585
            Width           =   2475
            _ExtentX        =   4366
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1185
         Left            =   6855
         Top             =   60
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   2090
         ForeColor       =   16711680
         Caption         =   "材　質"
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
         cPositionY      =   50
         Begin VB.ComboBox H1Comb_Zai 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   210
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   4
            Top             =   570
            Width           =   1800
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1185
         Left            =   60
         Top             =   60
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   2090
         ForeColor       =   255
         Caption         =   "対象期"
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
         cAlingnment     =   6
         cPositionY      =   60
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Height          =   480
            Left            =   1260
            TabIndex        =   9
            Top             =   330
            Width           =   1725
            Begin VB.OptionButton H1Op_Ki 
               BackColor       =   &H00C0C0C0&
               Caption         =   "上期"
               Height          =   285
               Index           =   0
               Left            =   90
               TabIndex        =   1
               Top             =   165
               Value           =   -1  'True
               Width           =   795
            End
            Begin VB.OptionButton H1Op_Ki 
               BackColor       =   &H00C0C0C0&
               Caption         =   "下期"
               Height          =   285
               Index           =   1
               Left            =   885
               TabIndex        =   2
               Top             =   165
               Width           =   795
            End
         End
         Begin Cis3D_v60.CIS3D H1lb_Kikan 
            Height          =   285
            Left            =   120
            Top             =   840
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   503
            BackColor       =   16777152
            Caption         =   "YYYY/MM/DD～YYYY/MM/DD"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
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
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   345
            Left            =   840
            Top             =   480
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   609
            ForeColor       =   16711680
            Caption         =   "年"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
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
            cBoderStyle     =   2
         End
         Begin CisText_V60.CisText iH1_Y 
            Height          =   345
            Left            =   120
            TabIndex        =   0
            Top             =   450
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   609
            cFormat         =   "####"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "####"
            MaxLength       =   4
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "####"
            cILength        =   4
         End
         Begin Cis3D_v60.CIS3D H1lb_Zenki 
            Height          =   285
            Left            =   120
            Top             =   1200
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   503
            BackColor       =   16777152
            Caption         =   "YYYY/MM/DD～YYYY/MM/DD"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
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
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   810
      Left            =   10620
      ScaleHeight     =   750
      ScaleWidth      =   4365
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   615
      Width           =   4425
      Begin CisBtn_60.CisBtn PB_LOOK 
         Height          =   585
         Left            =   30
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
         Left            =   3480
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
      Begin CisBtn_60.CisBtn PB_PRT 
         Height          =   585
         Left            =   900
         Top             =   90
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
         BtnType         =   7
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
   Begin VB.PictureBox B1_Area1 
      Height          =   8700
      Left            =   255
      ScaleHeight     =   8640
      ScaleWidth      =   14745
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2265
      Width           =   14805
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   8475
         Left            =   105
         Top             =   90
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   14949
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
         Begin TrueOleDBGrid60.TDBGrid DBGrid1 
            Bindings        =   "CXM2540.frx":0000
            Height          =   8370
            Left            =   30
            OleObjectBlob   =   "CXM2540.frx":0015
            TabIndex        =   10
            Top             =   30
            Width           =   14490
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
      Caption         =   "【 スクラップ単価変動額問合せ(支払)  】"
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
   End
   Begin Cis3D_v60.CIS3D CIS3D10 
      Height          =   345
      Left            =   13425
      Top             =   1605
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   609
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
      cPositionX      =   90
   End
   Begin Cis3D_v60.CIS3D B1lb_DispSu 
      Height          =   345
      Left            =   14295
      Top             =   1605
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   609
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
   Begin VB.TextBox Dummy 
      Height          =   390
      Left            =   30
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   420
      Width           =   195
   End
   Begin VB.Timer Timer1 
      Left            =   630
      Top             =   390
   End
   Begin VB.Timer Timer2 
      Left            =   210
      Top             =   390
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   270
      Top             =   1950
      Visible         =   0   'False
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "CXM2540"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   スクラップ単価変動額問合せ(支払)
'**       フォームID    :   CXM2540
'**       処理概要      :
'**
'**       作  成  日    :   2008/09/08  By CIS
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
'   Dim CisVsGrid       As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    Dim mTanKb          As Byte
    Dim wKi             As String
    Dim CisComb1        As CisCombo
    
    Dim SY_KamiM        As String       ' 上期開始月
    Dim SY_KamiSMD      As String       ' 上期開始日付
    Dim SY_KamiEMD      As String       ' 上期終了日付
    Dim SY_KamiZSMD     As String       ' 前期上期開始日付
    Dim SY_KamiZEMD     As String       ' 前期上期終了日付
    Dim SY_SimoM        As String       ' 下期開始月
    Dim SY_SimoSMD      As String       ' 下期開始日付
    Dim SY_SimoEMD      As String       ' 下期終了日付
    Dim SY_SimoZSMD     As String       ' 前期下期開始日付
    Dim SY_SimoZEMD     As String       ' 前期下期終了日付

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Public Sub form_load() 'Private Sub Form_Load()
    If Rv_Call3 = "" Then
    '   #------------------#
    '   # 二 重 起 動 防 止 #
    '   #------------------#
        If Not Dupli_Start Then End
    '   #------------------#
    '   # 初 期 内 容 取 得 #
    '   #------------------#
        If Not IniGet Then End
    End If
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
'
'   +------------------+
'   +   ｺﾝﾎﾞﾎﾞｯｸｽｾｯﾄ   +
'   +------------------+
    Set CisComb1 = New CisCombo
    With CisComb1
        Set .Connect = CisDB
        Set .CombObj = H1Comb_Zai
        .Control = "名称マスタ"
        .ComboName = "材種区分"
        .Ess = False
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    If Rv_Call3 = "" Then
       Me.Show
    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    FormAct = Not FormAct
'
    H1Op_Ki(RV_Ki).Value = True
    Adodc1.ConnectionString = CisDB.ConnectString
    
    Call CisComb1.ComboSet
    
    ' 適用月日抽出 & 退避
    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT * FROM 名称マスタ"
    gSL_Select = gSL_Select & "  WHERE 区分名称 = '自社情報'"
    gSL_Select = gSL_Select & "    AND 値 = '3'"
    gSL_Select = gSL_Select & "    AND 区分タイプ = ''"
    If SYMRead(gSL_Select) Then
       SY_KamiM = Format(SYM.数字1, "00")       ' 上期月
       SY_SimoM = Format(SYM.数字2, "00")       ' 下期月
    End If
    
    If Rv_Call3 = "X" Then
       iH1_Y = RV_Y
       Call H1Op_Ki_Click(CInt(RV_Ki))
       iH1_Torcd = RV_Torcd
       Call iH1_Torcd_LostFocus
       CisComb1.Code = RV_Zaisyu
       Call CisComb1.TextGet
                     
       Call ReturnPress
    End If
 
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
           Case vbKeyF7:       If PB_LOOK.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF8:       If PB_PRT.Visible Then Call PB_PRT_Click
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_Ent.Visible Then                      '【入力】
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
    Cancel = 1
    Call PB_END_Click
End Sub

'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    If Rv_Call3 = "" Then
    '+---------------------+
    '+   ﾃﾞｰﾀﾍﾞｰｽ切断
    '+---------------------+
        Call CisDB.DBDISConnect
    '+---------------------+
    '+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    '+---------------------+
        Call CisFrm.UnLoadDisp
        End
    Else
        Me.Hide
    End If
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
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      iH1_Y.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_LOOK.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
End Sub
'+----------------------------+
'+    印  刷  キ  ー (F8)     +
'+----------------------------+
Private Sub PB_PRT_Click()
    Call CisFrm.MousePT(11)
    
    PrtCnt = B1lb_DispSu                ' 印刷件数
    PrtNKi = iH1_Y & "年"               ' 対象期
    If H1Op_Ki(0) Then
        PrtNKi = PrtNKi & "上期"
    Else
        PrtNKi = PrtNKi & "下期"
    End If
    PrtKKn = H1lb_Kikan                 ' 期間

    CXM2545.Show vbModal
    Unload CXM2545
    Set CXM2545 = Nothing

PB_Prt_Click_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+       手配先検索            +
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
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    H1lb_Tornm = ""
    If TorNmGet(iH1_Torcd, 5, 1) Then
       H1lb_Tornm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iH1_Y
'+-----------------------------+
Private Sub iH1_Y_LostFocus()
    Call DateDSP(CisFun.Val2(wKi))
End Sub
'+-----------------+
'+     H1Op_Ki     +
'+-----------------+
Private Sub H1Op_Ki_Click(Index As Integer)
    wKi = Index
    
    Call DateDSP(Index)
    
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   Call CisFrm.TimeSet(1)
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
    PB_LOOK.Visible = False
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
'*---------------------------------------------- 2008.08.27 del start
'    If Rv_Call3 <> "" Then
'       If Rv_Call3 = "X" Then
'          PB_CAN.Visible = False
'       End If
'       PB_END.BtnCaption = "戻る"
'    End If
'*---------------------------------------------- 2008.08.27 del end
    PB_Ent.Visible = H1Mode                     '2008.09.08 add
    PB_PRT.Visible = Not H1Mode                 '2008.08.27 add
    
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
       DBGrid1.Close True
       Adodc1.Recordset.Close
       Adodc1.RecordSource = ""
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
    
    If Trim(iH1_Y) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    対象期を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Y.SetFocus
       Exit Function
    End If
    
    If Trim(iH1_Torcd) = "" Then
    Else
        If Not TorNmGet(iH1_Torcd, 5) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    支払先　未登録！            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Torcd.SetFocus
           Exit Function
        End If
    End If
' マスタ入力 SELECT文 ｾｯﾄ
    If Not GridSet Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "    支払先スクラップ単価マスタ 未登録！            "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
       End With
       iH1_Y.SetFocus
       Exit Function
    End If
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    Body1Chk = True
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      グリッド処理(vsGrid)                                 +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'******************************************'
'*****       グリッド内容セット       *****'
'******************************************'
Private Function GridSet() As Boolean
'   Dim wWhere      As String
    GridSet = False
    
''  Call Mk_SQL                         ' SQL作成
    Call Mk_SQL2                        ' SQL作成
    Adodc1.RecordSource = gSL_Select
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount = 0 Then Exit Function
    B1lb_DispSu = Format(Adodc1.Recordset.RecordCount, "#,###")
    
    GridSet = True
End Function

'+------------------+
'+     日付表示     +
'+------------------+
Private Sub DateDSP(Index As Integer)
    Dim wY As String
    
    If Len(Trim(iH1_Y)) < 4 And RTrim(iH1_Y) <> "" Then     ' 年が4桁未満ならば4桁にする
       With CisFun
            .Date_RsFormat = fYMD
            Call .Date_Add(iH1_Y & "0101", D, 1)
            wY = Mid(.Date_Result, 1, 4)
            If Mid(.Date_Result, 1, 1) = "0" Then
                 Mid(wY, 1, 1) = "2"
            End If
            iH1_Y = wY
        End With
    End If
    
    If RTrim(iH1_Y) = "" Then
       H1lb_Kikan = ""
    Else
        If Index = 0 Then
            '***** 上期開始月日抽出 *****
            SY_KamiSMD = iH1_Y & SY_KamiM & "01"
            '***** 上期終了月日抽出 *****
            SY_KamiEMD = iH1_Y & SY_SimoM & "01"
            With CisFun
                .Date_RsFormat = fYMD
                Call .Date_Add(SY_KamiEMD, D, -1)
                SY_KamiEMD = .Date_Result
            End With
            
            If SY_KamiSMD > SY_KamiEMD Then
               SY_KamiEMD = CisFun.Val2(iH1_Y) + 1 & SY_SimoM & "01"
               With CisFun
                   .Date_RsFormat = fYMD
                   Call .Date_Add(SY_KamiEMD, D, -1)
                   SY_KamiEMD = .Date_Result
               End With
            End If
            
            H1lb_Kikan = Mid(SY_KamiSMD, 1, 4) & "/" & _
                         Mid(SY_KamiSMD, 5, 2) & "/" & _
                         Mid(SY_KamiSMD, 7, 2) & "～" & _
                         Mid(SY_KamiEMD, 1, 4) & "/" & _
                         Mid(SY_KamiEMD, 5, 2) & "/" & _
                         Mid(SY_KamiEMD, 7, 2)
        Else
            '***** 下期開始月日抽出 *****
            SY_SimoSMD = iH1_Y & SY_SimoM & "01"
            '***** 下期終了月日抽出 *****
            SY_SimoEMD = iH1_Y & SY_KamiM & "01"
            With CisFun
                .Date_RsFormat = fYMD
                Call .Date_Add(SY_SimoEMD, D, -1)
                SY_SimoEMD = .Date_Result
            End With
            
            If SY_SimoSMD > SY_SimoEMD Then
               SY_SimoEMD = CisFun.Val2(iH1_Y) + 1 & SY_KamiM & "01"
               With CisFun
                   .Date_RsFormat = fYMD
                   Call .Date_Add(SY_SimoEMD, D, -1)
                   SY_SimoEMD = .Date_Result
               End With
            End If
            
            H1lb_Kikan = Mid(SY_SimoSMD, 1, 4) & "/" & _
                         Mid(SY_SimoSMD, 5, 2) & "/" & _
                         Mid(SY_SimoSMD, 7, 2) & "～" & _
                         Mid(SY_SimoEMD, 1, 4) & "/" & _
                         Mid(SY_SimoEMD, 5, 2) & "/" & _
                         Mid(SY_SimoEMD, 7, 2)
        End If
    End If
    
    '+------- 前期抽出 --------+
    Call ZenkiGet
    
End Sub

'+------------------+
'+     前期抽出     +
'+------------------+
Private Sub ZenkiGet()
    If RTrim(iH1_Y) = "" Then
       H1lb_Kikan = ""
    Else
        If H1Op_Ki(1).Value Then
            '***** 上期開始月日抽出 *****
            SY_KamiZSMD = iH1_Y & SY_KamiM & "01"
            '***** 上期終了月日抽出 *****
            With CisFun
                .Date_RsFormat = fYMD
                Call .Date_Add(SY_SimoSMD, D, -1)
                SY_KamiZEMD = .Date_Result
            End With
                        
            If SY_KamiZSMD > SY_KamiZEMD Then
               With CisFun
                   .Date_RsFormat = fYMD
                   Call .Date_Add(SY_KamiZSMD, Y, -1)
                   SY_KamiZSMD = .Date_Result
               End With
            End If
            
            H1lb_Zenki = Mid(SY_KamiZSMD, 1, 4) & "/" & _
                         Mid(SY_KamiZSMD, 5, 2) & "/" & _
                         Mid(SY_KamiZSMD, 7, 2) & "～" & _
                         Mid(SY_KamiZEMD, 1, 4) & "/" & _
                         Mid(SY_KamiZEMD, 5, 2) & "/" & _
                         Mid(SY_KamiZEMD, 7, 2)
        Else
            '***** 下期開始月日抽出 *****
            SY_SimoZSMD = iH1_Y & SY_SimoM & "01"
            '***** 下期終了月日抽出 *****
            With CisFun
                .Date_RsFormat = fYMD
                Call .Date_Add(SY_KamiSMD, D, -1)   ' 2008.08.27 upd : SY_SimoSMD
                SY_SimoZEMD = .Date_Result
            End With
            
            If SY_SimoZSMD > SY_SimoZEMD Then
               With CisFun
                   .Date_RsFormat = fYMD
                   Call .Date_Add(SY_SimoZSMD, Y, -1)
                   SY_SimoZSMD = .Date_Result
               End With
            End If
            
            H1lb_Zenki = Mid(SY_SimoZSMD, 1, 4) & "/" & _
                         Mid(SY_SimoZSMD, 5, 2) & "/" & _
                         Mid(SY_SimoZSMD, 7, 2) & "～" & _
                         Mid(SY_SimoZEMD, 1, 4) & "/" & _
                         Mid(SY_SimoZEMD, 5, 2) & "/" & _
                         Mid(SY_SimoZEMD, 7, 2)
        End If
    End If
End Sub

'+------------------+
'+     SQL 作成     +
'+------------------+
Private Function Mk_SQL() As Boolean
    Mk_SQL = False
    
    Dim wWhere      As String       ' WHERE句のみ
    wWhere = ""

    With CisFun
       .Where_CreateKB = NewDocument
       
       '''''.Where_Create AlphaMe, "ZTM.取引先", iH1_Torcd
       .Where_Create AlphaMe, "SSM1.支払先", iH1_Torcd
       
       Call CisComb1.CodeGet
       .Where_Create AlphaMe, "ZM.材種", CisComb1.Code
       
       wWhere = .Where_Phrase
    End With

    Call ZenkiGet                   ' 前期の日付を取得 2008.09.01 add
    gSL_Select = ""
    

'------------------------------------------------------------------ 2008.09.09 add ----------------------------------------
    gSL_Select = gSL_Select & "  SELECT " & vbCrLf
    gSL_Select = gSL_Select & "         CASE WHEN ISNULL(TRM.略称,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "              THEN TRM.取引先CD + ':' + TRM.略称 " & vbCrLf
    gSL_Select = gSL_Select & "              Else TRM.取引先CD " & vbCrLf
    gSL_Select = gSL_Select & "              END  編集支払先, " & vbCrLf
    '''''gSL_Select = gSL_Select & "         ZM.材質 +':' + SY2.値名称 材質, " & vbCrLf
    gSL_Select = gSL_Select & "         ISNULL(SY2.値名称,'') 材質," & vbCrLf
    gSL_Select = gSL_Select & "         CASE WHEN ISNULL(SSM2.履歴NO, 0) <> 0 " & vbCrLf
    gSL_Select = gSL_Select & "              THEN ISNULL(SSM2.単価  , 0) " & vbCrLf
    gSL_Select = gSL_Select & "              ELSE ISNULL(SSM3.単価  , 0) " & vbCrLf
    gSL_Select = gSL_Select & "              END  前回単価, " & vbCrLf
    gSL_Select = gSL_Select & "         ISNULL(SSM1.単価, 0) 今回単価, " & vbCrLf
    gSL_Select = gSL_Select & "         HNM.表示品番 品番, " & vbCrLf
    '''''gSL_Select = gSL_Select & "         ZTM.背番号, " & vbCrLf
    gSL_Select = gSL_Select & "         ISNULL(HTM.背番号,'') 背番号," & vbCrLf
    gSL_Select = gSL_Select & "         ZH.未使用量 スクラップ重量, " & vbCrLf
    gSL_Select = gSL_Select & "         CASE WHEN isnull(SSM2.履歴NO, 0) <> 0 " & vbCrLf
    gSL_Select = gSL_Select & "              THEN ISNULL(SSM2.単価  , 0)  * ZH.未使用量 " & vbCrLf
    gSL_Select = gSL_Select & "              ELSE ISNULL(SSM3.単価  , 0)  * ZH.未使用量 " & vbCrLf
    gSL_Select = gSL_Select & "              END  前回スクラップ単価, " & vbCrLf
    gSL_Select = gSL_Select & "         ISNULL(SSM1.単価, 0) * ZH.未使用量 今回スクラップ単価, " & vbCrLf
    gSL_Select = gSL_Select & "        (ISNULL(SSM1.単価, 0) - ISNULL(SSM2.単価, 0)) * ZH.未使用量 スクラップ単価変動額, " & vbCrLf
    gSL_Select = gSL_Select & "         CASE WHEN ISNULL( KT2.単価, 0) <> 0 " & vbCrLf
    gSL_Select = gSL_Select & "              THEN ISNULL( KT2.単価, 0) " & vbCrLf
    gSL_Select = gSL_Select & "              ELSE ISNULL(KT21.単価, 0) " & vbCrLf
    gSL_Select = gSL_Select & "              END  前回購入単価, " & vbCrLf
    gSL_Select = gSL_Select & "         CASE WHEN ISNULL( KT1.単価, 0) <> 0 " & vbCrLf
    gSL_Select = gSL_Select & "              THEN ISNULL( KT1.単価, 0) " & vbCrLf
    gSL_Select = gSL_Select & "              ELSE ISNULL(KT11.単価, 0) " & vbCrLf
    gSL_Select = gSL_Select & "              END  今回購入単価, " & vbCrLf
   'gSL_Select = gSL_Select & "         CASE WHEN ISNULL(SY1.数字1, 0) = 0 " & vbCrLf
   'gSL_Select = gSL_Select & "              THEN (ISNULL(KT1.単価, 0) - ISNULL(KT2.単価, 0)) * (1 + (       30 / 100)) " & vbCrLf
   'gSL_Select = gSL_Select & "              ELSE (ISNULL(KT1.単価, 0) - ISNULL(KT2.単価, 0)) * (1 + (SY1.数字1 / 100)) " & vbCrLf
   'gSL_Select = gSL_Select & "              END  購入単価変動額 " & vbCrLf
    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(SY1.数字1, 0) = 0 " & vbCrLf
    gSL_Select = gSL_Select & "         THEN CASE WHEN ISNULL(TRM.計算単位, 0) = 0 THEN ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                   WHEN ISNULL(TRM.計算単位, 0) = 1 THEN CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100))) " & vbCrLf
    gSL_Select = gSL_Select & "                   WHEN ISNULL(TRM.計算単位, 0) = 2 THEN ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 0) " & vbCrLf
    gSL_Select = gSL_Select & "                   WHEN ISNULL(TRM.計算単位, 0) = 3 THEN CASE WHEN ISNULL(TRM.計算位1, 0) = 0 " & vbCrLf
    gSL_Select = gSL_Select & "                                                              THEN CASE WHEN ISNULL(TRM.計算位2, 0) = 0 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 1, 1)     , 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                        WHEN ISNULL(TRM.計算位2, 0) = 1 THEN ROUND(CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100))  * 10) / 10, 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                        WHEN ISNULL(TRM.計算位2, 0) = 2 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 1   )     , 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                   End " & vbCrLf
    gSL_Select = gSL_Select & "                                                              WHEN IsNull(TRM.計算位1, 0) = 1 " & vbCrLf
    gSL_Select = gSL_Select & "                                                              THEN CASE WHEN ISNULL(TRM.計算位2, 0) = 0 THEN CEILING(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 1, 1)     ) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                        WHEN ISNULL(TRM.計算位2, 0) = 1 THEN CEILING(CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100))  * 10) / 10) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                        WHEN ISNULL(TRM.計算位2, 0) = 2 THEN CEILING(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 1   )     ) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                   End " & vbCrLf
    gSL_Select = gSL_Select & "                                                              WHEN IsNull(TRM.計算位1, 0) = 2 " & vbCrLf
    gSL_Select = gSL_Select & "                                                              THEN CASE WHEN ISNULL(TRM.計算位2, 0) = 0 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 1, 1)     , 0) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                        WHEN ISNULL(TRM.計算位2, 0) = 1 THEN ROUND(CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100))  * 10) / 10, 0) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                        WHEN ISNULL(TRM.計算位2, 0) = 2 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 1   )     , 0) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                   End " & vbCrLf
    gSL_Select = gSL_Select & "                                                         End " & vbCrLf
    gSL_Select = gSL_Select & "              End " & vbCrLf
    gSL_Select = gSL_Select & "         ELSE CASE WHEN ISNULL(TRM.計算単位, 0) = 0 THEN ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                   WHEN ISNULL(TRM.計算単位, 0) = 1 THEN CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100))) " & vbCrLf
    gSL_Select = gSL_Select & "                   WHEN ISNULL(TRM.計算単位, 0) = 2 THEN ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 0) " & vbCrLf
    gSL_Select = gSL_Select & "                   WHEN ISNULL(TRM.計算単位, 0) = 3 THEN CASE WHEN ISNULL(TRM.計算位1, 0) = 0 " & vbCrLf
    gSL_Select = gSL_Select & "                                                              THEN CASE WHEN ISNULL(TRM.計算位2, 0) = 0 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 1, 1)     , 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                        WHEN ISNULL(TRM.計算位2, 0) = 1 THEN ROUND(CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100))  * 10) / 10, 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                        WHEN ISNULL(TRM.計算位2, 0) = 2 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 1   )     , 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                   End " & vbCrLf
    gSL_Select = gSL_Select & "                                                              WHEN IsNull(TRM.計算位1, 0) = 1 " & vbCrLf
    gSL_Select = gSL_Select & "                                                              THEN CASE WHEN ISNULL(TRM.計算位2, 0) = 0 THEN CEILING(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 1, 1)     ) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                        WHEN ISNULL(TRM.計算位2, 0) = 1 THEN CEILING(CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100))  * 10) / 10) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                        WHEN ISNULL(TRM.計算位2, 0) = 2 THEN CEILING(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 1   )     ) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                   End " & vbCrLf
    gSL_Select = gSL_Select & "                                                              WHEN IsNull(TRM.計算位1, 0) = 2 " & vbCrLf
    gSL_Select = gSL_Select & "                                                              THEN CASE WHEN ISNULL(TRM.計算位2, 0) = 0 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 1, 1)     , 0) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                        WHEN ISNULL(TRM.計算位2, 0) = 1 THEN ROUND(CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100))  * 10) / 10, 0) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                        WHEN ISNULL(TRM.計算位2, 0) = 2 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 1   )     , 0) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                   End " & vbCrLf
    gSL_Select = gSL_Select & "                                                          End " & vbCrLf
    gSL_Select = gSL_Select & "              End " & vbCrLf
    gSL_Select = gSL_Select & "         END As 購入単価変動額 " & vbCrLf

    gSL_Select = gSL_Select & " FROM 材料品番マスタ ZH " & vbCrLf
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 材料マスタ ZM " & vbCrLf
    gSL_Select = gSL_Select & "                 ON   ZM.材料管理番号 = ZH.材料管理番号 " & vbCrLf
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 品番マスタ HNM " & vbCrLf
    gSL_Select = gSL_Select & "                 ON  HNM.品番 = ZH.使用品番 " & vbCrLf
    '''''gSL_Select = gSL_Select & "         LEFT OUTER JOIN 材料取引先マスタ ZTM " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 ON  ZTM.材料管理番号 = ZM.材料管理番号 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                AND  ZTM.材質         = ZM.材質 " & vbCrLf

    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 支払先スクラップ単価マスタ SSM1 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 ON SSM1.支払先       = ZTM.取引先 " & vbCrLf
    gSL_Select = gSL_Select & "                 ON SSM1.支払先       = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & "                AND SSM1.材質         = ZM.材種 " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
       gSL_Select = gSL_Select & "                AND SSM1.適用開始日付 = '" & SY_KamiSMD & "'" & vbCrLf
       gSL_Select = gSL_Select & "                AND SSM1.適用終了日付 = '" & SY_KamiEMD & "'" & vbCrLf
    Else
       '********* 今期下期 *********
       gSL_Select = gSL_Select & "                AND SSM1.適用開始日付 = '" & SY_SimoSMD & "'" & vbCrLf
       gSL_Select = gSL_Select & "                AND SSM1.適用終了日付 = '" & SY_SimoEMD & "'" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                AND SSM1.履歴NO       = (SELECT MAX(SSM1S.履歴NO) " & vbCrLf
    gSL_Select = gSL_Select & "                                         FROM   支払先スクラップ単価マスタ SSM1S " & vbCrLf
    gSL_Select = gSL_Select & "                                         Where  SSM1S.材質         = SSM1.材質 " & vbCrLf
    gSL_Select = gSL_Select & "                                           AND  SSM1S.支払先       = SSM1.支払先 " & vbCrLf
    gSL_Select = gSL_Select & "                                           AND  SSM1S.適用開始日付 = SSM1.適用開始日付" & vbCrLf
    gSL_Select = gSL_Select & "                                           AND  SSM1S.適用終了日付 = SSM1.適用終了日付) " & vbCrLf
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 支払先スクラップ単価マスタ SSM2 " & vbCrLf
    gSL_Select = gSL_Select & "                 ON SSM2.支払先       = SSM1.支払先 " & vbCrLf
    gSL_Select = gSL_Select & "                AND SSM2.材質         = SSM1.材質 " & vbCrLf
    '''''If H1Op_Ki(0).Value Then
    '''''   '********* 今期上期のとき、今期上期 *********
    '''''   gSL_Select = gSL_Select & "                AND SSM2.適用開始日付 = '" & SY_KamiSMD & "'" & vbCrLf
    '''''   gSL_Select = gSL_Select & "                AND SSM2.適用終了日付 = '" & SY_KamiEMD & "'" & vbCrLf
    '''''Else
    '''''   '********* 今期下期のとき、今期下期 *********
    '''''   gSL_Select = gSL_Select & "                AND SSM2.適用開始日付 = '" & SY_SimoSMD & "'" & vbCrLf
    '''''   gSL_Select = gSL_Select & "                AND SSM2.適用終了日付 = '" & SY_SimoEMD & "'" & vbCrLf
    '''''End If
    gSL_Select = gSL_Select & "                AND SSM2.履歴NO       = SSM1.履歴NO - 1 " & vbCrLf
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 支払先スクラップ単価マスタ SSM3 " & vbCrLf
    gSL_Select = gSL_Select & "                 ON SSM3.支払先       = SSM1.支払先 " & vbCrLf
    gSL_Select = gSL_Select & "                AND SSM3.材質         = SSM1.材質 " & vbCrLf
    '''''If H1Op_Ki(0).Value Then
    '''''   '********* 今期上期のとき、前期下期 *********
    '''''   gSL_Select = gSL_Select & "                AND SSM3.適用開始日付 = '" & SY_SimoZSMD & "'" & vbCrLf
    '''''   gSL_Select = gSL_Select & "                AND SSM3.適用終了日付 = '" & SY_SimoZEMD & "'" & vbCrLf
    '''''Else
    '''''   '********* 今期下期のとき、今期上期 *********
    '''''   gSL_Select = gSL_Select & "                AND SSM3.適用開始日付 = '" & SY_KamiSMD & "'" & vbCrLf
    '''''   gSL_Select = gSL_Select & "                AND SSM3.適用終了日付 = '" & SY_KamiEMD & "'" & vbCrLf
    '''''End If
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
       gSL_Select = gSL_Select & "             AND SSM3.適用開始日付 = '" & SY_KamiSMD & "'" & vbCrLf
       gSL_Select = gSL_Select & "             AND SSM3.適用終了日付 = '" & SY_KamiEMD & "'" & vbCrLf
    Else
       '********* 今期下期 *********
       gSL_Select = gSL_Select & "             AND SSM3.適用開始日付 = '" & SY_SimoSMD & "'" & vbCrLf
       gSL_Select = gSL_Select & "             AND SSM3.適用終了日付 = '" & SY_SimoEMD & "'" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                AND SSM3.履歴NO       = (SELECT MAX(SSM3S.履歴NO) " & vbCrLf
    gSL_Select = gSL_Select & "                                         FROM   支払先スクラップ単価マスタ SSM3S " & vbCrLf
    gSL_Select = gSL_Select & "                                         Where  SSM3S.材質         = SSM1.材質 " & vbCrLf
    gSL_Select = gSL_Select & "                                           AND  SSM3S.支払先       = SSM1.支払先 " & vbCrLf
    gSL_Select = gSL_Select & "                                           AND  SSM3S.適用開始日付 = SSM3.適用開始日付 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                                           AND  SSM3S.適用終了日付 = SSM3.適用終了日付) " & vbCrLf
    gSL_Select = gSL_Select & "                                           AND  SSM3S.適用終了日付 = SSM3.適用終了日付) - 1" & vbCrLf
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 品番取引先マスタ HTM " & vbCrLf
    gSL_Select = gSL_Select & "                 ON  HTM.品番    = HNM.品番 " & vbCrLf
    gSL_Select = gSL_Select & "                AND  HTM.品目    = 1 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                AND  HTM.契約先  = SSM1.支払先 " & vbCrLf 'AND  HTM.取引先  = SSM1.支払先"
    '''''gSL_Select = gSL_Select & "         LEFT OUTER JOIN 品番取引先マスタ HTM2 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 ON HTM2.品番    = HNM.品番 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                AND HTM2.品目    = 1 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                AND HTM2.契約先 <> SSM1.支払先 " & vbCrLf 'AND HTM2.取引先 <> SSM1.支払先"

    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 取引先マスタ TRM " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 ON (TRM.取引先CD   =  HTM.取引先 " & vbCrLf       'SSM1.支払先"
    '''''gSL_Select = gSL_Select & "                 OR  TRM.取引先CD   = HTM2.取引先) " & vbCrLf
    gSL_Select = gSL_Select & "                 ON  TRM.取引先CD = SSM1.支払先" & vbCrLf
    gSL_Select = gSL_Select & "                AND  TRM.取引先区分 = 1 " & vbCrLf

    gSL_Select = gSL_Select & "         LEFT OUTER JOIN dbo.購入単価マスタ AS KT1 " & vbCrLf
    gSL_Select = gSL_Select & "                 ON   KT1.品番       = ZH.使用品番 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                AND   KT1.取引先     = TRM.取引先CD " & vbCrLf
    '''''gSL_Select = gSL_Select & "                AND   KT1.契約先     = TRM.契約先  " & vbCrLf
    gSL_Select = gSL_Select & "                AND   KT1.取引先     = HTM.取引先 "
    gSL_Select = gSL_Select & "                AND   KT1.契約先     = HTM.契約先  " & vbCrLf
    gSL_Select = gSL_Select & "                AND   KT1.仮単価区分 = '' " & vbCrLf
    gSL_Select = gSL_Select & "                AND ((KT1.適用開始日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
      'gSL_Select = gSL_Select & "                AND   KT1.適用開始日付 <= '" & SY_KamiSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT1.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
      'gSL_Select = gSL_Select & "                AND   KT1.適用開始日付 <= '" & SY_SimoSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT1.適用開始日付 >= '" & SY_SimoSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                 OR   KT1.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                AND ((KT1.適用終了日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
      'gSL_Select = gSL_Select & "                AND   KT1.適用終了日付 >= '" & SY_KamiEMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT1.適用終了日付 <= '" & SY_KamiEMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
      'gSL_Select = gSL_Select & "                AND   KT1.適用終了日付 >= '" & SY_SimoEMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT1.適用終了日付 <= '" & SY_SimoEMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                 OR   KT1.適用終了日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                AND   KT1.履歴NO        = (SELECT MAX(履歴NO)        AS Expr1 " & vbCrLf
    gSL_Select = gSL_Select & "                                           FROM   dbo.購入単価マスタ AS KT1S " & vbCrLf
    gSL_Select = gSL_Select & "                                           Where (KT1S.品番 = KT1.品番) " & vbCrLf
    gSL_Select = gSL_Select & "                                             AND  (KT1S.取引先        = KT1.取引先) " & vbCrLf
    gSL_Select = gSL_Select & "                                             AND  (KT1S.契約先        = KT1.契約先) " & vbCrLf
    gSL_Select = gSL_Select & "                                             AND  (KT1S.仮単価区分    = ''        ) " & vbCrLf
    gSL_Select = gSL_Select & "                                             AND ((KT1S.適用開始日付 <> '' " & vbCrLf
   'gSL_Select = gSL_Select & "                                             AND   KT1S.適用開始日付 <= '20080401') " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
      'gSL_Select = gSL_Select & "                                             AND   KT1S.適用開始日付 <= '" & SY_KamiSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                             AND   KT1S.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
      'gSL_Select = gSL_Select & "                                             AND   KT1S.適用開始日付 <= '" & SY_SimoSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                             AND   KT1S.適用開始日付 >= '" & SY_SimoSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                              OR   KT1S.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                                             AND ((KT1S.適用終了日付 <> '' " & vbCrLf
   'gSL_Select = gSL_Select & "                                             AND   KT1S.適用終了日付 >= '20080930') " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
      'gSL_Select = gSL_Select & "                                             AND   KT1S.適用終了日付 >= '" & SY_KamiEMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                             AND   KT1S.適用終了日付 <= '" & SY_KamiEMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
      'gSL_Select = gSL_Select & "                                             AND   KT1S.適用終了日付 >= '" & SY_SimoEMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                             AND   KT1S.適用終了日付 <= '" & SY_SimoEMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                              OR   KT1S.適用終了日付  = '')) " & vbCrLf
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN dbo.購入単価マスタ AS KT11 " & vbCrLf
    gSL_Select = gSL_Select & "                 ON   KT11.品番       = ZH.使用品番 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                AND   KT11.取引先     = TRM.取引先CD " & vbCrLf
    gSL_Select = gSL_Select & "                AND   KT11.取引先     = HTM.取引先 "
    gSL_Select = gSL_Select & "                AND   KT11.契約先     = TRM.契約先  " & vbCrLf
    gSL_Select = gSL_Select & "                AND   KT11.仮単価区分 = '1' " & vbCrLf
    gSL_Select = gSL_Select & "                AND ((KT11.適用開始日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
      'gSL_Select = gSL_Select & "                AND   KT11.適用開始日付 <= '" & SY_KamiSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT11.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
      'gSL_Select = gSL_Select & "                AND   KT11.適用開始日付 <= '" & SY_SimoSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT11.適用開始日付 >= '" & SY_SimoSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                 OR   KT11.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                AND ((KT11.適用終了日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
      'gSL_Select = gSL_Select & "                AND   KT11.適用終了日付 >= '" & SY_KamiEMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT11.適用終了日付 <= '" & SY_KamiEMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
      'gSL_Select = gSL_Select & "                AND   KT11.適用終了日付 >= '" & SY_SimoEMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT11.適用終了日付 <= '" & SY_SimoEMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                 OR   KT11.適用終了日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                AND   KT11.履歴NO        = (SELECT MAX(履歴NO)        AS Expr1 " & vbCrLf
    gSL_Select = gSL_Select & "                                            FROM   dbo.購入単価マスタ AS KT11S " & vbCrLf
    gSL_Select = gSL_Select & "                                            Where  (KT11S.品番 = KT11.品番) " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND  (KT11S.取引先        = KT11.取引先) " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND  (KT11S.契約先        = KT11.契約先) " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND  (KT11S.仮単価区分    = '') " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND ((KT11S.適用開始日付 <> '' " & vbCrLf
   'gSL_Select = gSL_Select & "                                              AND   KT11S.適用開始日付 <= '20080401') " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
      'gSL_Select = gSL_Select & "                                              AND   KT11S.適用開始日付 <= '" & SY_KamiSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                              AND   KT11S.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
      'gSL_Select = gSL_Select & "                                              AND   KT11S.適用開始日付 <= '" & SY_SimoSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                              AND   KT11S.適用開始日付 >= '" & SY_SimoSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                               OR   KT11S.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND ((KT11S.適用終了日付 <> '' " & vbCrLf
   'gSL_Select = gSL_Select & "                                              AND   KT11S.適用終了日付 >= '20080930') " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
      'gSL_Select = gSL_Select & "                                              AND   KT11S.適用終了日付 >= '" & SY_KamiEMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                              AND   KT11S.適用終了日付 <= '" & SY_KamiEMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
      'gSL_Select = gSL_Select & "                                              AND   KT11S.適用終了日付 >= '" & SY_SimoEMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                              AND   KT11S.適用終了日付 <= '" & SY_SimoEMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                               OR   KT11S.適用終了日付  = '')) " & vbCrLf


    gSL_Select = gSL_Select & "         LEFT OUTER JOIN dbo.購入単価マスタ AS KT2 " & vbCrLf
    gSL_Select = gSL_Select & "                 ON   KT2.品番       = ZH.使用品番 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                AND   KT2.取引先     = TRM.取引先CD " & vbCrLf
    'gSL_Select = gSL_Select & "                AND   KT2.契約先     = TRM.契約先  " & vbCrLf
    gSL_Select = gSL_Select & "                AND   KT2.取引先     = HTM.取引先 "
    gSL_Select = gSL_Select & "                AND   KT2.仮単価区分 = '' " & vbCrLf
    gSL_Select = gSL_Select & "                AND ((KT2.適用開始日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
      'gSL_Select = gSL_Select & "                AND   KT2.適用開始日付 <= '" & SY_SimoZSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT2.適用開始日付 >= '" & SY_SimoZSMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、、今期上期 *********
      'gSL_Select = gSL_Select & "                AND   KTS.適用開始日付 <= '" & SY_KamiSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT2.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                OR    KT2.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                AND ((KT2.適用終了日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
      'gSL_Select = gSL_Select & "                AND   KT2.適用終了日付 >= '" & SY_SimoZEMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT2.適用終了日付 <= '" & SY_SimoZEMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、今期上期 *********
      'gSL_Select = gSL_Select & "                AND   KT2.適用終了日付 >= '" & SY_KamiSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT2.適用終了日付 <= '" & SY_KamiSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                OR    KT2.適用終了日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                AND KT2.履歴NO     = (SELECT MAX(履歴NO)      AS Expr1 " & vbCrLf
    gSL_Select = gSL_Select & "                                      FROM   dbo.購入単価マスタ AS KT2S " & vbCrLf
    gSL_Select = gSL_Select & "                                      Where  (KT2S.品番 = KT1.品番) " & vbCrLf
    gSL_Select = gSL_Select & "                                        AND  (KT2S.取引先        = KT2.取引先) " & vbCrLf
    gSL_Select = gSL_Select & "                                        AND  (KT2S.契約先        = KT2.契約先) " & vbCrLf
    gSL_Select = gSL_Select & "                                        AND  (KT2S.仮単価区分    = ''       ) " & vbCrLf
    gSL_Select = gSL_Select & "                                        AND ((KT2S.適用開始日付 <> '' " & vbCrLf
   'gSL_Select = gSL_Select & "                                        AND   KT2S.適用開始日付 <= '20071001') " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
      'gSL_Select = gSL_Select & "                                        AND   KT2S.適用開始日付 <= '" & SY_SimoZSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                        AND   KT2S.適用開始日付 >= '" & SY_SimoZSMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、、今期上期 *********
      'gSL_Select = gSL_Select & "                                        AND   KT2S.適用開始日付 <= '" & SY_KamiSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                        AND   KT2S.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                        OR    KT2S.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                                        AND ((KT2S.適用終了日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
      'gSL_Select = gSL_Select & "                                        AND   KT2S.適用終了日付 >= '" & SY_SimoZEMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                        AND   KT2S.適用終了日付 <= '" & SY_SimoZEMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、今期上期 *********
      'gSL_Select = gSL_Select & "                                        AND   KT2S.適用終了日付 >= '" & SY_KamiSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                        AND   KT2S.適用終了日付 <= '" & SY_KamiSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                        OR    KT2S.適用終了日付  = '')) " & vbCrLf
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN dbo.購入単価マスタ AS KT21 " & vbCrLf
    gSL_Select = gSL_Select & "                 ON   KT21.品番       = ZH.使用品番 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                AND   KT21.取引先     = TRM.取引先CD " & vbCrLf
    'gSL_Select = gSL_Select & "                AND   KT21.契約先     = TRM.契約先  " & vbCrLf
    gSL_Select = gSL_Select & "                AND   KT11.取引先     = HTM.取引先 "
    gSL_Select = gSL_Select & "                AND   KT21.仮単価区分 = '1' " & vbCrLf
    gSL_Select = gSL_Select & "                AND ((KT21.適用開始日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
      'gSL_Select = gSL_Select & "                AND   KT21.適用開始日付 <= '" & SY_SimoZSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT21.適用開始日付 >= '" & SY_SimoZSMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、、今期上期 *********
      'gSL_Select = gSL_Select & "                AND   KT21.適用開始日付 <= '" & SY_KamiSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT21.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                OR    KT21.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                AND ((KT21.適用終了日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
      'gSL_Select = gSL_Select & "                AND   KT21.適用終了日付 >= '" & SY_SimoZEMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT21.適用終了日付 <= '" & SY_SimoZEMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、今期上期 *********
       gSL_Select = gSL_Select & "                AND   KT21.適用終了日付 >= '" & SY_KamiSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                AND   KT21.適用終了日付 <= '" & SY_KamiSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                OR    KT21.適用終了日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                AND KT21.履歴NO     = (SELECT MAX(履歴NO)        AS Expr1 " & vbCrLf
    gSL_Select = gSL_Select & "                                       FROM   dbo.購入単価マスタ AS KT21S " & vbCrLf
    gSL_Select = gSL_Select & "                                       Where  (KT21S.品番 = KT21.品番) " & vbCrLf
    gSL_Select = gSL_Select & "                                         AND  (KT21S.取引先        = KT21.取引先) " & vbCrLf
    gSL_Select = gSL_Select & "                                         AND  (KT21S.契約先        = KT21.契約先) " & vbCrLf
    gSL_Select = gSL_Select & "                                         AND  (KT21S.仮単価区分    = '') " & vbCrLf
    gSL_Select = gSL_Select & "                                         AND ((KT21S.適用開始日付 <> '' " & vbCrLf
   'gSL_Select = gSL_Select & "                                         AND   KT21S.適用開始日付 <= '20071001') " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
      'gSL_Select = gSL_Select & "                                         AND   KT21S.適用開始日付 <= '" & SY_SimoZSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                         AND   KT21S.適用開始日付 >= '" & SY_SimoZSMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、、今期上期 *********
      'gSL_Select = gSL_Select & "                                         AND   KT21S.適用開始日付 <= '" & SY_KamiSMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                         AND   KT21S.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                         OR    KT21S.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                                         AND ((KT21S.適用終了日付 <> '' " & vbCrLf
   'gSL_Select = gSL_Select & "                                         AND   KT21S.適用終了日付 >= '20080331') " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
      'gSL_Select = gSL_Select & "                                         AND   KT21S.適用終了日付 >= '" & SY_SimoZEMD & "')" & vbCrLf
       gSL_Select = gSL_Select & "                                         AND   KT21S.適用終了日付 <= '" & SY_SimoZEMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、今期上期 *********
      'gSL_Select = gSL_Select & "                                         AND   KT21S.適用終了日付 >= '" & SY_KamiSMD & " ')" & vbCrLf
       gSL_Select = gSL_Select & "                                         AND   KT21S.適用終了日付 <= '" & SY_KamiSMD & " ')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                         OR    KT21S.適用終了日付  = '')) " & vbCrLf

    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 名称マスタ SY1 " & vbCrLf
    gSL_Select = gSL_Select & "                 ON SY1.区分名称   = 'スクラップ単価変動' " & vbCrLf
    gSL_Select = gSL_Select & "                AND SY1.値         = '1' " & vbCrLf
    gSL_Select = gSL_Select & "                AND SY1.区分タイプ = '' " & vbCrLf
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 名称マスタ SY2 " & vbCrLf
    gSL_Select = gSL_Select & "                 ON SY2.区分名称   = '材種区分' " & vbCrLf
    gSL_Select = gSL_Select & "                AND SY2.値         = ZM.材種 " & vbCrLf
    gSL_Select = gSL_Select & "                AND SY2.区分タイプ = '' " & vbCrLf
     
    If wWhere <> "" Then
       gSL_Select = gSL_Select & wWhere & " AND " & vbCrLf
    Else
       gSL_Select = gSL_Select & " WHERE " & vbCrLf
    End If
    gSL_Select = gSL_Select & "       ISNULL( ZM.材料管理番号, 0) <> 0 " & vbCrLf
    '''''gSL_Select = gSL_Select & "   AND ISNULL(ZTM.材料管理番号, 0) <> 0 " & vbCrLf
    
    '''''If H1Op_Ki(0).Value Then
    '''''   gSL_Select = gSL_Select & "   AND SSM1.適用開始日付 = '" & SY_KamiSMD & "' " & vbCrLf
    '''''   gSL_Select = gSL_Select & "   AND SSM1.適用終了日付 = '" & SY_KamiEMD & "' " & vbCrLf
    '''''Else
    '''''   gSL_Select = gSL_Select & "   AND SSM1.適用開始日付 = '" & SY_SimoSMD & "' " & vbCrLf
    '''''   gSL_Select = gSL_Select & "   AND SSM1.適用終了日付 = '" & SY_SimoEMD & "' " & vbCrLf
    '''''End If

    gSL_Select = gSL_Select & " GROUP BY ZH.未使用量 , ZH.使用品番, " & vbCrLf
    gSL_Select = gSL_Select & "          ZM.材質     , ZM.材種    , " & vbCrLf
    gSL_Select = gSL_Select & "          HNM.表示品番, " & vbCrLf
    '''''gSL_Select = gSL_Select & "          ZTM.背番号  , ZTM.取引先 , " & vbCrLf
    gSL_Select = gSL_Select & "          ISNULL(HTM.背番号,'')  , ZTM.取引先 , " & vbCrLf
    gSL_Select = gSL_Select & "          SSM1.単価   , SSM2.履歴NO, SSM2.単価, SSM3.単価, " & vbCrLf
    gSL_Select = gSL_Select & "          HTM.取引先  , " & vbCrLf
    gSL_Select = gSL_Select & "          TRM.取引先CD, TRM.略称   , " & vbCrLf
    gSL_Select = gSL_Select & "          KT1.単価    , KT11.単価  , KT2.単価 , KT21.単価, " & vbCrLf
    gSL_Select = gSL_Select & "          SY1.数字1   , SY2.値名称 , " & vbCrLf
    gSL_Select = gSL_Select & "          SSM1.適用開始日付, " & vbCrLf
    gSL_Select = gSL_Select & "          TRM.計算単位, TRM.計算位1, TRM.計算位2 " & vbCrLf
    
    gSL_Select = gSL_Select & " ORDER BY 編集支払先, 材質, ZH.使用品番 " & vbCrLf
'------------------------------------------------------------------ 2008.09.09 add ----------------------------------------
Call Test_LogOutput(gSL_Select, 0)
    
    Mk_SQL = True
End Function

'+------------------+
'+     SQL 作成 2   +
'+------------------+
Private Function Mk_SQL2() As Boolean
    Mk_SQL2 = False
    
    Dim wWhere      As String       ' WHERE句のみ
    wWhere = ""

    With CisFun
       .Where_CreateKB = NewDocument
       
       '''''.Where_Create AlphaMe, "ZTM.契約先", iH1_Torcd   ' ZTM.取引先
       .Where_Create AlphaMe, "SSM1.支払先", iH1_Torcd
       
       Call CisComb1.CodeGet
       .Where_Create AlphaMe, "ZM.材種", CisComb1.Code
       
       wWhere = .Where_Phrase
    End With

    Call ZenkiGet                   ' 前期の日付を取得 2008.09.01 add
    gSL_Select = ""
    

'------------------------------------------------------------------ 2008.09.10 add ----------------------------------------
    gSL_Select = gSL_Select & "   SELECT " & vbCrLf
    gSL_Select = gSL_Select & "          CASE WHEN ISNULL(TRM.略称,'') <> '' " & vbCrLf
    gSL_Select = gSL_Select & "               THEN TRM.取引先CD + ':' + TRM.略称 " & vbCrLf
    gSL_Select = gSL_Select & "               Else TRM.取引先CD " & vbCrLf
    gSL_Select = gSL_Select & "               END  編集支払先, " & vbCrLf
    '''''gSL_Select = gSL_Select & "          ZM.材質 +':' + SY2.値名称 材質, " & vbCrLf
    gSL_Select = gSL_Select & "          ISNULL(SY2.値名称,'') 材質," & vbCrLf
    gSL_Select = gSL_Select & "          CASE WHEN ISNULL(SSM2.履歴NO, 0) <> 0 " & vbCrLf
    gSL_Select = gSL_Select & "               THEN ISNULL(SSM2.単価  , 0) " & vbCrLf
    gSL_Select = gSL_Select & "               ELSE ISNULL(SSM3.単価  , 0) " & vbCrLf
    gSL_Select = gSL_Select & "               END  前回単価, " & vbCrLf
    gSL_Select = gSL_Select & "          ISNULL(SSM1.単価, 0) 今回単価, " & vbCrLf
    gSL_Select = gSL_Select & "          HNM.表示品番 品番, " & vbCrLf
    '''''gSL_Select = gSL_Select & "          ZTM.背番号, " & vbCrLf
    gSL_Select = gSL_Select & "          ISNULL(HTM.背番号,'') 背番号," & vbCrLf
    gSL_Select = gSL_Select & "          ZH.未使用量 スクラップ重量, " & vbCrLf
    gSL_Select = gSL_Select & "          CASE WHEN isnull(SSM2.履歴NO, 0) <> 0 " & vbCrLf
    gSL_Select = gSL_Select & "               THEN ISNULL(SSM2.単価  , 0)  * ZH.未使用量 " & vbCrLf
    gSL_Select = gSL_Select & "               ELSE ISNULL(SSM3.単価  , 0)  * ZH.未使用量 " & vbCrLf
    gSL_Select = gSL_Select & "               END  前回スクラップ単価, " & vbCrLf
    gSL_Select = gSL_Select & "          ISNULL(SSM1.単価, 0) * ZH.未使用量 今回スクラップ単価, " & vbCrLf
    
   'gSL_Select = gSL_Select & "         (ISNULL(SSM1.単価, 0) - ISNULL(SSM2.単価, 0)) * ZH.未使用量 スクラップ単価変動額, " & vbCrLf
    gSL_Select = gSL_Select & "          CASE WHEN ISNULL(SSM2.履歴NO, 0) <> 0 " & vbCrLf
    gSL_Select = gSL_Select & "               THEN (ISNULL(SSM1.単価, 0) - ISNULL(SSM2.単価, 0)) * ZH.未使用量 " & vbCrLf
    gSL_Select = gSL_Select & "               ELSE (ISNULL(SSM1.単価, 0) - ISNULL(SSM3.単価, 0)) * ZH.未使用量 " & vbCrLf
    gSL_Select = gSL_Select & "               END  スクラップ単価変動額, " & vbCrLf


'        ---- 前回分 ------------------------------------------------------------------------------------
    gSL_Select = gSL_Select & "         CASE WHEN ISNULL( KT2.単価, 0) <> 0 " & vbCrLf
    gSL_Select = gSL_Select & "               THEN ISNULL( KT2.単価, 0) " & vbCrLf
    gSL_Select = gSL_Select & "               ELSE ISNULL(KT21.単価, 0) " & vbCrLf
    gSL_Select = gSL_Select & "               END  前回購入単価, " & vbCrLf
'        ---- 今回分 ------------------------------------------------------------------------------------
    gSL_Select = gSL_Select & "          CASE WHEN ISNULL( KT1.単価, 0) <> 0 " & vbCrLf
    gSL_Select = gSL_Select & "               THEN ISNULL( KT1.単価, 0) " & vbCrLf
    gSL_Select = gSL_Select & "               ELSE ISNULL(KT11.単価, 0) " & vbCrLf
    gSL_Select = gSL_Select & "               END  今回購入単価, " & vbCrLf
'        ---- 変動額 ------------------------------------------------------------------------------------
'        -- CASE WHEN  ISNULL(SY1.数字1, 0) = 0
'        --      THEN (ISNULL(KT1.単価, 0) - ISNULL(KT2.単価, 0)) * (1 + (       30 / 100))
'        --      ELSE (ISNULL(KT1.単価, 0) - ISNULL(KT2.単価, 0)) * (1 + (SY1.数字1 / 100))
'        --      END  購入単価変動額,
    gSL_Select = gSL_Select & "     CASE WHEN ISNULL(SY1.数字1, 0) = 0 " & vbCrLf
    gSL_Select = gSL_Select & "          THEN CASE WHEN ISNULL(TRM.計算単位, 0) = 0 THEN ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                    WHEN ISNULL(TRM.計算単位, 0) = 1 THEN CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100))) " & vbCrLf
    gSL_Select = gSL_Select & "                    WHEN ISNULL(TRM.計算単位, 0) = 2 THEN ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 0) " & vbCrLf
    gSL_Select = gSL_Select & "                    WHEN ISNULL(TRM.計算単位, 0) = 3 THEN CASE WHEN ISNULL(TRM.計算位1, 0) = 0 " & vbCrLf
    gSL_Select = gSL_Select & "                                                               THEN CASE WHEN ISNULL(TRM.計算位2, 0) = 0 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 1, 1)     , 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                         WHEN ISNULL(TRM.計算位2, 0) = 1 THEN ROUND(CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100))  * 10) / 10, 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                         WHEN ISNULL(TRM.計算位2, 0) = 2 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 1   )     , 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                    End " & vbCrLf
    gSL_Select = gSL_Select & "                                                               WHEN IsNull(TRM.計算位1, 0) = 1 " & vbCrLf
    gSL_Select = gSL_Select & "                                                               THEN CASE WHEN ISNULL(TRM.計算位2, 0) = 0 THEN CEILING(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 1, 1)     ) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                         WHEN ISNULL(TRM.計算位2, 0) = 1 THEN CEILING(CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100))  * 10) / 10) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                         WHEN ISNULL(TRM.計算位2, 0) = 2 THEN CEILING(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 1   )     ) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                    End " & vbCrLf
    gSL_Select = gSL_Select & "                                                               WHEN IsNull(TRM.計算位1, 0) = 2 " & vbCrLf
    gSL_Select = gSL_Select & "                                                               THEN CASE WHEN ISNULL(TRM.計算位2, 0) = 0 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 1, 1)     , 0) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                         WHEN ISNULL(TRM.計算位2, 0) = 1 THEN ROUND(CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100))  * 10) / 10, 0) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                         WHEN ISNULL(TRM.計算位2, 0) = 2 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (       30 / 100)), 1   )     , 0) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                    End " & vbCrLf
    gSL_Select = gSL_Select & "                                                          End " & vbCrLf
    gSL_Select = gSL_Select & "               End " & vbCrLf
    gSL_Select = gSL_Select & "          ELSE CASE WHEN ISNULL(TRM.計算単位, 0) = 0 THEN ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                    WHEN ISNULL(TRM.計算単位, 0) = 1 THEN CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100))) " & vbCrLf
    gSL_Select = gSL_Select & "                    WHEN ISNULL(TRM.計算単位, 0) = 2 THEN ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 0) " & vbCrLf
    gSL_Select = gSL_Select & "                    WHEN ISNULL(TRM.計算単位, 0) = 3 THEN CASE WHEN ISNULL(TRM.計算位1, 0) = 0 " & vbCrLf
    gSL_Select = gSL_Select & "                                                               THEN CASE WHEN ISNULL(TRM.計算位2, 0) = 0 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 1, 1)     , 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                         WHEN ISNULL(TRM.計算位2, 0) = 1 THEN ROUND(CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100))  * 10) / 10, 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                         WHEN ISNULL(TRM.計算位2, 0) = 2 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 1   )     , 0, 1) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                    End " & vbCrLf
    gSL_Select = gSL_Select & "                                                               WHEN IsNull(TRM.計算位1, 0) = 1 " & vbCrLf
    gSL_Select = gSL_Select & "                                                               THEN CASE WHEN ISNULL(TRM.計算位2, 0) = 0 THEN CEILING(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 1, 1)     ) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                         WHEN ISNULL(TRM.計算位2, 0) = 1 THEN CEILING(CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100))  * 10) / 10) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                         WHEN ISNULL(TRM.計算位2, 0) = 2 THEN CEILING(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 1   )     ) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                    End " & vbCrLf
    gSL_Select = gSL_Select & "                                                               WHEN IsNull(TRM.計算位1, 0) = 2 " & vbCrLf
    gSL_Select = gSL_Select & "                                                               THEN CASE WHEN ISNULL(TRM.計算位2, 0) = 0 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 1, 1)     , 0) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                         WHEN ISNULL(TRM.計算位2, 0) = 1 THEN ROUND(CEILING((ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100))  * 10) / 10, 0) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                         WHEN ISNULL(TRM.計算位2, 0) = 2 THEN ROUND(ROUND(  (ISNULL(KT1.単価, ISNULL(KT11.単価, 0)) - ISNULL(KT2.単価, ISNULL(KT21.単価, 0))) * (1 + (SY1.数字1 / 100)), 1   )     , 0) " & vbCrLf
    gSL_Select = gSL_Select & "                                                                    End " & vbCrLf
    gSL_Select = gSL_Select & "                                                           End " & vbCrLf
    gSL_Select = gSL_Select & "               End " & vbCrLf
    gSL_Select = gSL_Select & "          END As 購入単価変動額 " & vbCrLf

    gSL_Select = gSL_Select & "  FROM 材料品番マスタ ZH " & vbCrLf
    gSL_Select = gSL_Select & "          LEFT OUTER JOIN 材料マスタ ZM " & vbCrLf
    gSL_Select = gSL_Select & "                  ON   ZM.材料管理番号 = ZH.材料管理番号 " & vbCrLf
    gSL_Select = gSL_Select & "          LEFT OUTER JOIN 品番マスタ HNM " & vbCrLf
    gSL_Select = gSL_Select & "                  ON  HNM.品番 = ZH.使用品番 " & vbCrLf
    '''''gSL_Select = gSL_Select & "          LEFT OUTER JOIN 材料取引先マスタ ZTM " & vbCrLf
    '''''gSL_Select = gSL_Select & "                  ON  ZTM.材料管理番号 = ZM.材料管理番号 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 AND  ZTM.材質         = ZM.材質 " & vbCrLf

    gSL_Select = gSL_Select & "          LEFT OUTER JOIN 支払先スクラップ単価マスタ SSM1 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                  ON SSM1.支払先       = ZTM.契約先  " & vbCrLf
    gSL_Select = gSL_Select & "                 ON SSM1.支払先       = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & "                 AND SSM1.材質         = ZM.材種 " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
       gSL_Select = gSL_Select & "                AND SSM1.適用開始日付 = '" & SY_KamiSMD & "'" & vbCrLf
       gSL_Select = gSL_Select & "                AND SSM1.適用終了日付 = '" & SY_KamiEMD & "'" & vbCrLf
    Else
       '********* 今期下期 *********
       gSL_Select = gSL_Select & "                AND SSM1.適用開始日付 = '" & SY_SimoSMD & "'" & vbCrLf
       gSL_Select = gSL_Select & "                AND SSM1.適用終了日付 = '" & SY_SimoEMD & "'" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                 AND SSM1.履歴NO       = (SELECT MAX(SSM1S.履歴NO) " & vbCrLf
    gSL_Select = gSL_Select & "                                          FROM   支払先スクラップ単価マスタ SSM1S " & vbCrLf
    gSL_Select = gSL_Select & "                                          Where SSM1S.材質 = SSM1.材質 " & vbCrLf
    gSL_Select = gSL_Select & "                                            AND  SSM1S.支払先       = SSM1.支払先 " & vbCrLf
    gSL_Select = gSL_Select & "                                            AND  SSM1S.適用開始日付 = SSM1.適用開始日付 " & vbCrLf
    gSL_Select = gSL_Select & "                                            AND  SSM1S.適用終了日付 = SSM1.適用終了日付) " & vbCrLf
    gSL_Select = gSL_Select & "          LEFT OUTER JOIN 支払先スクラップ単価マスタ SSM2 " & vbCrLf
    gSL_Select = gSL_Select & "                  ON SSM2.支払先       = SSM1.支払先 " & vbCrLf
    gSL_Select = gSL_Select & "                 AND SSM2.材質         = SSM1.材質 " & vbCrLf
    '''''If H1Op_Ki(0).Value Then
    '''''   '********* 今期上期のとき、今期上期 *********
    '''''   gSL_Select = gSL_Select & "                AND SSM2.適用開始日付 = '" & SY_KamiSMD & "'" & vbCrLf
    '''''   gSL_Select = gSL_Select & "                AND SSM2.適用終了日付 = '" & SY_KamiEMD & "'" & vbCrLf
    '''''Else
    '''''   '********* 今期下期のとき、今期下期 *********
    '''''   gSL_Select = gSL_Select & "                AND SSM2.適用開始日付 = '" & SY_SimoSMD & "'" & vbCrLf
    '''''   gSL_Select = gSL_Select & "                AND SSM2.適用終了日付 = '" & SY_SimoEMD & "'" & vbCrLf
    '''''End If
    gSL_Select = gSL_Select & "                 AND SSM2.履歴NO       = SSM1.履歴NO - 1 " & vbCrLf
    gSL_Select = gSL_Select & "          LEFT OUTER JOIN 支払先スクラップ単価マスタ SSM3 " & vbCrLf
    gSL_Select = gSL_Select & "                  ON SSM3.支払先       = SSM1.支払先 " & vbCrLf
    gSL_Select = gSL_Select & "                 AND SSM3.材質         = SSM1.材質 " & vbCrLf
    '''''If H1Op_Ki(0).Value Then
    '''''   '********* 今期上期のとき、前期下期 *********
    '''''   gSL_Select = gSL_Select & "                AND SSM3.適用開始日付 = '" & SY_SimoZSMD & "'" & vbCrLf
    '''''   gSL_Select = gSL_Select & "                AND SSM3.適用終了日付 = '" & SY_SimoZEMD & "'" & vbCrLf
    '''''Else
    '''''   '********* 今期下期のとき、今期上期 *********
    '''''   gSL_Select = gSL_Select & "                AND SSM3.適用開始日付 = '" & SY_KamiSMD & "'" & vbCrLf
    '''''   gSL_Select = gSL_Select & "                AND SSM3.適用終了日付 = '" & SY_KamiEMD & "'" & vbCrLf
    '''''End If
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
       gSL_Select = gSL_Select & "             AND SSM3.適用開始日付 = '" & SY_KamiSMD & "'" & vbCrLf
       gSL_Select = gSL_Select & "             AND SSM3.適用終了日付 = '" & SY_KamiEMD & "'" & vbCrLf
    Else
       '********* 今期下期 *********
       gSL_Select = gSL_Select & "             AND SSM3.適用開始日付 = '" & SY_SimoSMD & "'" & vbCrLf
       gSL_Select = gSL_Select & "             AND SSM3.適用終了日付 = '" & SY_SimoEMD & "'" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                 AND SSM3.履歴NO       = (SELECT MAX(SSM3S.履歴NO) " & vbCrLf
    gSL_Select = gSL_Select & "                                          FROM   支払先スクラップ単価マスタ SSM3S " & vbCrLf
    gSL_Select = gSL_Select & "                                          Where SSM3S.材質 = SSM1.材質 " & vbCrLf
    gSL_Select = gSL_Select & "                                            AND  SSM3S.支払先       = SSM1.支払先 " & vbCrLf
    gSL_Select = gSL_Select & "                                            AND  SSM3S.適用開始日付 = SSM3.適用開始日付 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                                            AND  SSM3S.適用終了日付 = SSM3.適用終了日付) " & vbCrLf
    gSL_Select = gSL_Select & "                                           AND  SSM3S.適用終了日付 = SSM3.適用終了日付) - 1" & vbCrLf

    gSL_Select = gSL_Select & "          LEFT OUTER JOIN 品番取引先マスタ HTM " & vbCrLf
    gSL_Select = gSL_Select & "                  ON  HTM.品番    = HNM.品番 " & vbCrLf
    gSL_Select = gSL_Select & "                 AND  HTM.品目    = 1 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 AND  HTM.取引先  = ZTM.取引先    " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 AND  HTM.契約先  = ZTM.契約先    " & vbCrLf
    '''''gSL_Select = gSL_Select & "          LEFT OUTER JOIN 品番取引先マスタ HTM2 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                  ON HTM2.品番    = HNM.品番 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 AND HTM2.品目    = 1 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 AND HTM2.契約先 <> ZTM.契約先    " & vbCrLf
    gSL_Select = gSL_Select & "          LEFT OUTER JOIN 取引先マスタ TRM " & vbCrLf
    '''''gSL_Select = gSL_Select & "                  ON (TRM.取引先CD   =  HTM.取引先   " & vbCrLf
    '''''gSL_Select = gSL_Select & "                  OR  TRM.取引先CD   = HTM2.取引先)  " & vbCrLf
    gSL_Select = gSL_Select & "                  ON  TRM.取引先CD = SSM1.支払先" & vbCrLf
    gSL_Select = gSL_Select & "                 AND  TRM.取引先区分 = 1 " & vbCrLf

'------- 今回分 ------------
    gSL_Select = gSL_Select & "          LEFT OUTER JOIN dbo.購入単価マスタ AS KT1 " & vbCrLf
    gSL_Select = gSL_Select & "                  ON   KT1.品番       = ZH.使用品番 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 AND   KT1.取引先     = TRM.取引先CD " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 AND   KT1.契約先     = TRM.契約先 " & vbCrLf
    gSL_Select = gSL_Select & "                 AND   KT1.取引先     = HTM.取引先 "
    'gSL_Select = gSL_Select & "                 AND   KT1.契約先     = HTM.契約先 "
    gSL_Select = gSL_Select & "                 AND   KT1.仮単価区分 = '' " & vbCrLf
    gSL_Select = gSL_Select & "                 AND ((KT1.適用開始日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
       gSL_Select = gSL_Select & "                AND   KT1.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
       gSL_Select = gSL_Select & "                AND   KT1.適用開始日付 >= '" & SY_SimoSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                  OR   KT1.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                 AND ((KT1.適用終了日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
       gSL_Select = gSL_Select & "                AND   KT1.適用終了日付 <= '" & SY_KamiEMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
       gSL_Select = gSL_Select & "                AND   KT1.適用終了日付 <= '" & SY_SimoEMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                  OR   KT1.適用終了日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                 AND   KT1.履歴NO        = (SELECT MAX(履歴NO)        AS Expr1 " & vbCrLf
    gSL_Select = gSL_Select & "                                            FROM   dbo.購入単価マスタ AS KT1S " & vbCrLf
    gSL_Select = gSL_Select & "                                            Where (KT1S.品番 = KT1.品番) " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND  (KT1S.取引先        = KT1.取引先) " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND  (KT1S.契約先        = KT1.契約先) " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND  (KT1S.仮単価区分    = ''        ) " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND ((KT1S.適用開始日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
       gSL_Select = gSL_Select & "                                             AND   KT1S.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
       gSL_Select = gSL_Select & "                                             AND   KT1S.適用開始日付 >= '" & SY_SimoSMD & "')" & vbCrLf
    End If

    gSL_Select = gSL_Select & "                                               OR   KT1S.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND ((KT1S.適用終了日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
       gSL_Select = gSL_Select & "                                             AND   KT1S.適用終了日付 <= '" & SY_KamiEMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
       gSL_Select = gSL_Select & "                                             AND   KT1S.適用終了日付 <= '" & SY_SimoEMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                               OR   KT1S.適用終了日付  = '')) " & vbCrLf
    gSL_Select = gSL_Select & "          LEFT OUTER JOIN dbo.購入単価マスタ AS KT11 " & vbCrLf
    gSL_Select = gSL_Select & "                  ON   KT11.品番       = ZH.使用品番 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 AND   KT11.取引先     = TRM.取引先CD " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 AND   KT11.契約先     = TRM.契約先 " & vbCrLf
    gSL_Select = gSL_Select & "                 AND   KT11.取引先     = HTM.取引先 "
    'gSL_Select = gSL_Select & "                 AND   KT11.契約先     = HTM.契約先 " & vbCrLf
    gSL_Select = gSL_Select & "                 AND   KT11.仮単価区分 = '1' " & vbCrLf
    gSL_Select = gSL_Select & "                 AND ((KT11.適用開始日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
       gSL_Select = gSL_Select & "                AND   KT11.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
       gSL_Select = gSL_Select & "                AND   KT11.適用開始日付 >= '" & SY_SimoSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                  OR   KT11.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                 AND ((KT11.適用終了日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
       gSL_Select = gSL_Select & "                AND   KT11.適用終了日付 <= '" & SY_KamiEMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
       gSL_Select = gSL_Select & "                AND   KT11.適用終了日付 <= '" & SY_SimoEMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                  OR   KT11.適用終了日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                 AND   KT11.履歴NO        = (SELECT MAX(履歴NO)        AS Expr1 " & vbCrLf
    gSL_Select = gSL_Select & "                                             FROM   dbo.購入単価マスタ AS KT11S " & vbCrLf
    gSL_Select = gSL_Select & "                                             Where (KT11S.品番 = KT11.品番) " & vbCrLf
    gSL_Select = gSL_Select & "                                               AND  (KT11S.取引先        = KT11.取引先) " & vbCrLf
    gSL_Select = gSL_Select & "                                               AND  (KT11S.契約先        = KT11.契約先) " & vbCrLf
    gSL_Select = gSL_Select & "                                               AND  (KT11S.仮単価区分    = '') " & vbCrLf
    gSL_Select = gSL_Select & "                                               AND ((KT11S.適用開始日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
       gSL_Select = gSL_Select & "                                              AND   KT11S.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
       gSL_Select = gSL_Select & "                                              AND   KT11S.適用開始日付 >= '" & SY_SimoSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                                OR   KT11S.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                                               AND ((KT11S.適用終了日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期 *********
       gSL_Select = gSL_Select & "                                              AND   KT11S.適用終了日付 <= '" & SY_KamiEMD & "')" & vbCrLf
    Else
       '********* 今期下期 *********
       gSL_Select = gSL_Select & "                                              AND   KT11S.適用終了日付 <= '" & SY_SimoEMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                                OR   KT11S.適用終了日付  = '')) " & vbCrLf


'------- 前回分 ------------
    gSL_Select = gSL_Select & "          LEFT OUTER JOIN dbo.購入単価マスタ AS KT2 " & vbCrLf
    gSL_Select = gSL_Select & "                  ON   KT2.品番       = ZH.使用品番 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 AND   KT2.取引先     = TRM.取引先CD " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 AND   KT2.契約先     = TRM.契約先 " & vbCrLf
    gSL_Select = gSL_Select & "                 AND   KT2.取引先     = HTM.取引先 "
    'gSL_Select = gSL_Select & "                 AND   KT2.契約先     = HTM.契約先 "
    gSL_Select = gSL_Select & "                 AND   KT2.仮単価区分 = '' " & vbCrLf
    gSL_Select = gSL_Select & "                 AND ((KT2.適用開始日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
       gSL_Select = gSL_Select & "                AND   KT2.適用開始日付 >= '" & SY_SimoZSMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、、今期上期 *********
       gSL_Select = gSL_Select & "                AND   KT2.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                 OR    KT2.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                 AND ((KT2.適用終了日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
       gSL_Select = gSL_Select & "                AND   KT2.適用終了日付 <= '" & SY_SimoZEMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、今期上期 *********
       gSL_Select = gSL_Select & "                AND   KT2.適用終了日付 <= '" & SY_KamiSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                 OR    KT2.適用終了日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                 AND   KT2.履歴NO        = (SELECT MAX(履歴NO)      AS Expr1 " & vbCrLf
    gSL_Select = gSL_Select & "                                            FROM dbo.購入単価マスタ AS KT2S " & vbCrLf
    gSL_Select = gSL_Select & "                                            Where (KT2S.品番 = KT1.品番) " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND  (KT2S.取引先        = KT2.取引先) " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND  (KT2S.契約先        = KT2.契約先) " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND  (KT2S.仮単価区分    = ''       ) " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND ((KT2S.適用開始日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
       gSL_Select = gSL_Select & "                                        AND   KT2S.適用開始日付 >= '" & SY_SimoZSMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、、今期上期 *********
       gSL_Select = gSL_Select & "                                        AND   KT2S.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                              OR    KT2S.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                                              AND ((KT2S.適用終了日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
       gSL_Select = gSL_Select & "                                        AND   KT2S.適用終了日付 <= '" & SY_SimoZEMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、今期上期 *********
       gSL_Select = gSL_Select & "                                        AND   KT2S.適用終了日付 <= '" & SY_KamiSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                              OR    KT2S.適用終了日付  = '')) " & vbCrLf
    
    gSL_Select = gSL_Select & "          LEFT OUTER JOIN dbo.購入単価マスタ AS KT21 " & vbCrLf
    gSL_Select = gSL_Select & "                  ON   KT21.品番       = ZH.使用品番 " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 AND   KT21.取引先     = TRM.取引先CD " & vbCrLf
    '''''gSL_Select = gSL_Select & "                 AND   KT21.契約先     = TRM.契約先 " & vbCrLf
    gSL_Select = gSL_Select & "                 AND   KT21.取引先     = HTM.取引先 "
    'gSL_Select = gSL_Select & "                 AND   KT21.契約先     = HTM.契約先 "
    gSL_Select = gSL_Select & "                 AND   KT21.仮単価区分 = '1' " & vbCrLf
    gSL_Select = gSL_Select & "                 AND ((KT21.適用開始日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
       gSL_Select = gSL_Select & "                AND   KT21.適用開始日付 >= '" & SY_SimoZSMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、、今期上期 *********
       gSL_Select = gSL_Select & "                AND   KT21.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                 OR    KT21.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                 AND ((KT21.適用終了日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
       gSL_Select = gSL_Select & "                AND   KT21.適用終了日付 <= '" & SY_SimoZEMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、今期上期 *********
       gSL_Select = gSL_Select & "                AND   KT21.適用終了日付 <= '" & SY_KamiSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                 OR    KT21.適用終了日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                 AND   KT21.履歴NO        = (SELECT MAX(履歴NO)        AS Expr1 " & vbCrLf
    gSL_Select = gSL_Select & "                                             FROM   dbo.購入単価マスタ AS KT21S " & vbCrLf
    gSL_Select = gSL_Select & "                                             Where (KT21S.品番 = KT21.品番) " & vbCrLf
    gSL_Select = gSL_Select & "                                               AND  (KT21S.取引先        = KT21.取引先) " & vbCrLf
    gSL_Select = gSL_Select & "                                               AND  (KT21S.契約先        = KT21.契約先) " & vbCrLf
    gSL_Select = gSL_Select & "                                               AND  (KT21S.仮単価区分    = '') " & vbCrLf
    gSL_Select = gSL_Select & "                                               AND ((KT21S.適用開始日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
       gSL_Select = gSL_Select & "                                         AND   KT21S.適用開始日付 >= '" & SY_SimoZSMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、、今期上期 *********
       gSL_Select = gSL_Select & "                                         AND   KT21S.適用開始日付 >= '" & SY_KamiSMD & "')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                               OR    KT21S.適用開始日付  = '') " & vbCrLf
    gSL_Select = gSL_Select & "                                               AND ((KT21S.適用終了日付 <> '' " & vbCrLf
    If H1Op_Ki(0).Value Then
       '********* 今期上期のとき、前期下期 *********
       gSL_Select = gSL_Select & "                                         AND   KT21S.適用終了日付 <= '" & SY_SimoZEMD & "')" & vbCrLf
    Else
       '********* 今期下期のとき、今期上期 *********
       gSL_Select = gSL_Select & "                                         AND   KT21S.適用終了日付 <= '" & SY_KamiSMD & " ')" & vbCrLf
    End If
    gSL_Select = gSL_Select & "                                               OR    KT21S.適用終了日付  = '')) " & vbCrLf

    gSL_Select = gSL_Select & "          LEFT OUTER JOIN 名称マスタ SY1 " & vbCrLf
    gSL_Select = gSL_Select & "                  ON SY1.区分名称   = 'スクラップ単価変動' " & vbCrLf
    gSL_Select = gSL_Select & "                 AND SY1.値         = '1' " & vbCrLf
    gSL_Select = gSL_Select & "                 AND SY1.区分タイプ = '' " & vbCrLf
    gSL_Select = gSL_Select & "          LEFT OUTER JOIN 名称マスタ SY2 " & vbCrLf
    gSL_Select = gSL_Select & "                  ON SY2.区分名称   = '材種区分' " & vbCrLf
    gSL_Select = gSL_Select & "                 AND SY2.値         = ZM.材種 " & vbCrLf
    gSL_Select = gSL_Select & "                 AND SY2.区分タイプ = '' " & vbCrLf
 
    If wWhere <> "" Then
       gSL_Select = gSL_Select & wWhere & " AND " & vbCrLf
    Else
       gSL_Select = gSL_Select & " WHERE " & vbCrLf
    End If
    gSL_Select = gSL_Select & "        IsNull( ZM.材料管理番号, 0) <> 0 " & vbCrLf
    '''''gSL_Select = gSL_Select & "    AND ISNULL(ZTM.材料管理番号, 0) <> 0 " & vbCrLf
    
    '''''If H1Op_Ki(0).Value Then
    '''''   gSL_Select = gSL_Select & "   AND SSM1.適用開始日付 = '" & SY_KamiSMD & "' " & vbCrLf
    '''''   gSL_Select = gSL_Select & "   AND SSM1.適用終了日付 = '" & SY_KamiEMD & "' " & vbCrLf
    '''''Else
    '''''   gSL_Select = gSL_Select & "   AND SSM1.適用開始日付 = '" & SY_SimoSMD & "' " & vbCrLf
    '''''   gSL_Select = gSL_Select & "   AND SSM1.適用終了日付 = '" & SY_SimoEMD & "' " & vbCrLf
    '''''End If
 
    gSL_Select = gSL_Select & "  GROUP BY ZH.未使用量 , ZH.使用品番, " & vbCrLf
    gSL_Select = gSL_Select & "           ZM.材質     , ZM.材種    , " & vbCrLf
    gSL_Select = gSL_Select & "           HNM.表示品番, " & vbCrLf
    '''''gSL_Select = gSL_Select & "           ZTM.背番号  , ZTM.取引先 , " & vbCrLf
    gSL_Select = gSL_Select & "           ISNULL(HTM.背番号,'')  , " & vbCrLf
    gSL_Select = gSL_Select & "           SSM1.単価   , SSM2.履歴NO, SSM2.単価, SSM3.単価, " & vbCrLf
    gSL_Select = gSL_Select & "           HTM.取引先  , " & vbCrLf
    gSL_Select = gSL_Select & "           TRM.取引先CD, TRM.略称   , " & vbCrLf
    gSL_Select = gSL_Select & "           KT1.単価    , KT11.単価  , KT2.単価 , KT21.単価, " & vbCrLf
    gSL_Select = gSL_Select & "           SY1.数字1   , SY2.値名称 , " & vbCrLf
    gSL_Select = gSL_Select & "           SSM1.適用開始日付, " & vbCrLf
    '''''gSL_Select = gSL_Select & "           TRM.計算単位 , TRM.計算位1, TRM.計算位2,ZTM.契約先 " & vbCrLf
    gSL_Select = gSL_Select & "           TRM.計算単位 , TRM.計算位1, TRM.計算位2" & vbCrLf

    gSL_Select = gSL_Select & "  ORDER BY 編集支払先, 材質, ZH.使用品番 " & vbCrLf
'------------------------------------------------------------------ 2008.09.10 add ----------------------------------------
Call Test_LogOutput(gSL_Select, 0)
    
    Mk_SQL2 = True
End Function

