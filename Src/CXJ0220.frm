VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Begin VB.Form CXJ0220 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "出荷実績入力 [個別]"
   ClientHeight    =   10530
   ClientLeft      =   900
   ClientTop       =   2385
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
   ScaleHeight     =   10530
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D H4_Area1 
      Height          =   1005
      Left            =   975
      Top             =   3060
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   1773
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
      Begin VB.TextBox QRDummy 
         Height          =   330
         Left            =   3660
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   30
         Width           =   240
      End
      Begin Cis3D_v60.CIS3D QRMsg 
         Height          =   855
         Left            =   90
         Top             =   75
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   1508
         BackColor       =   16761024
         ForeColor       =   16711680
         Caption         =   "QRﾀｯﾁﾘｰﾀﾞ読取"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   7
         cAlingnment     =   7
      End
   End
   Begin PdqcommLib.PDQComm PDQComm2 
      Height          =   480
      Left            =   120
      TabIndex        =   69
      Top             =   1635
      Width           =   480
      _Version        =   196610
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.26
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FastScroll      =   0   'False
      Object.Height          =   34
      Object.Width           =   84
   End
   Begin Cis3D_v60.CIS3D ErrMsg 
      Height          =   465
      Left            =   1140
      Top             =   795
      Visible         =   0   'False
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   820
      BackColor       =   16761087
      Caption         =   "エラー"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   14.25
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
   Begin Cis3D_v60.CIS3D Back_BunCount 
      Height          =   795
      Left            =   3585
      Top             =   540
      Visible         =   0   'False
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   1402
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
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   705
         Left            =   45
         Top             =   45
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1244
         ForeColor       =   128
         Caption         =   "分納処理件数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D Bun_Count 
            Height          =   630
            Left            =   1950
            Top             =   30
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   1111
            Caption         =   "ZZZ / ZZZ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
   End
   Begin Cis3D_v60.CIS3D CIS3D1 
      Height          =   375
      Left            =   9510
      Top             =   0
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   661
      BackColor       =   0
      ForeColor       =   65535
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
      cBoderColor1    =   8421504
      cBoderColor2    =   -2147483643
      cFont3DColor1   =   16777215
      cFont3DColor2   =   12632256
      cAlingnment     =   7
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   345
         Left            =   60
         Top             =   15
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
         Caption         =   "実績枚数"
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
         cAlingnment     =   1
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D lb_JMai 
            Height          =   270
            Left            =   975
            Top             =   30
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   476
            BackColor       =   12648384
            Caption         =   "#,###"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
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
      Height          =   990
      Left            =   1095
      ScaleHeight     =   930
      ScaleWidth      =   1860
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   1365
      Width           =   1920
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   840
         Left            =   60
         Top             =   45
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1482
         ForeColor       =   16711680
         Caption         =   "出荷指示書№"
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
         Begin CisText_V60.CisText iH1_SJNo 
            Height          =   360
            Left            =   195
            TabIndex        =   1
            Top             =   375
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   635
            cFormat         =   "0000000000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0000000000"
            MaxLength       =   10
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##########"
            cILength        =   8
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
      Height          =   720
      Left            =   8025
      ScaleHeight     =   660
      ScaleWidth      =   6045
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   495
      Width           =   6105
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   3465
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
         Left            =   5175
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
         Left            =   2610
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
         Left            =   1755
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
      Begin CisBtn_60.CisBtn PB_CHG 
         Height          =   585
         Left            =   900
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
         BtnType         =   10
         BtnCaption      =   "切替"
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
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   7815
      Left            =   1095
      ScaleHeight     =   7755
      ScaleWidth      =   12975
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   2535
      Width           =   13035
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   7665
         Left            =   60
         Top             =   45
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   13520
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
         Begin Cis3D_v60.CIS3D CIS3D17 
            Height          =   7035
            Left            =   7950
            Top             =   555
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   12409
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
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   19
               Left            =   15
               Top             =   4380
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   10
                  Left            =   15
                  TabIndex        =   41
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   18
               Left            =   15
               Top             =   3945
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   9
                  Left            =   15
                  TabIndex        =   37
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   17
               Left            =   15
               Top             =   3510
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   8
                  Left            =   15
                  TabIndex        =   33
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   16
               Left            =   15
               Top             =   3075
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   7
                  Left            =   15
                  TabIndex        =   29
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   15
               Left            =   15
               Top             =   2640
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   6
                  Left            =   15
                  TabIndex        =   25
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   14
               Left            =   15
               Top             =   2205
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   5
                  Left            =   15
                  TabIndex        =   21
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   13
               Left            =   15
               Top             =   1770
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   4
                  Left            =   15
                  TabIndex        =   17
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   12
               Left            =   15
               Top             =   1335
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   3
                  Left            =   15
                  TabIndex        =   13
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   11
               Left            =   15
               Top             =   900
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   2
                  Left            =   15
                  TabIndex        =   9
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D18 
               Height          =   450
               Left            =   1050
               Top             =   15
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   794
               ForeColor       =   16711680
               Caption         =   "分割納入数"
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
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D19 
               Height          =   450
               Left            =   2445
               Top             =   15
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   794
               ForeColor       =   16711680
               Caption         =   " 完納予定日  - 便"
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
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   0
               Left            =   1050
               Top             =   465
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   1
                  Left            =   285
                  TabIndex        =   6
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   1
               Left            =   2445
               Top             =   465
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   1
                  Left            =   45
                  TabIndex        =   7
                  Top             =   60
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   1
                  Left            =   1860
                  TabIndex        =   8
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   1
               Left            =   1050
               Top             =   900
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   2
                  Left            =   285
                  TabIndex        =   10
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   2
               Left            =   2445
               Top             =   900
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   2
                  Left            =   1860
                  TabIndex        =   12
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   2
                  Left            =   60
                  TabIndex        =   11
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   2
               Left            =   1050
               Top             =   1335
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   3
                  Left            =   285
                  TabIndex        =   14
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   3
               Left            =   2445
               Top             =   1335
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   3
                  Left            =   1860
                  TabIndex        =   16
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   3
                  Left            =   60
                  TabIndex        =   15
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   3
               Left            =   1050
               Top             =   1770
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   4
                  Left            =   285
                  TabIndex        =   18
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   4
               Left            =   2445
               Top             =   1770
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   4
                  Left            =   1860
                  TabIndex        =   20
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   4
                  Left            =   60
                  TabIndex        =   19
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   4
               Left            =   1050
               Top             =   2205
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   5
                  Left            =   285
                  TabIndex        =   22
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   5
               Left            =   2445
               Top             =   2205
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   5
                  Left            =   1860
                  TabIndex        =   24
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   5
                  Left            =   60
                  TabIndex        =   23
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   5
               Left            =   1050
               Top             =   2640
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   6
                  Left            =   285
                  TabIndex        =   26
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   6
               Left            =   2445
               Top             =   2640
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   6
                  Left            =   1860
                  TabIndex        =   28
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   6
                  Left            =   60
                  TabIndex        =   27
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   6
               Left            =   1050
               Top             =   3075
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   7
                  Left            =   285
                  TabIndex        =   30
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   7
               Left            =   2445
               Top             =   3075
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   7
                  Left            =   1860
                  TabIndex        =   32
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   7
                  Left            =   60
                  TabIndex        =   31
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   7
               Left            =   1050
               Top             =   3510
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   8
                  Left            =   285
                  TabIndex        =   34
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   8
               Left            =   2445
               Top             =   3510
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   8
                  Left            =   1860
                  TabIndex        =   36
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   8
                  Left            =   60
                  TabIndex        =   35
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   8
               Left            =   1050
               Top             =   3945
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   9
                  Left            =   285
                  TabIndex        =   38
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   9
               Left            =   2445
               Top             =   3945
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   9
                  Left            =   1860
                  TabIndex        =   40
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   9
                  Left            =   60
                  TabIndex        =   39
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   9
               Left            =   1050
               Top             =   4380
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   10
                  Left            =   285
                  TabIndex        =   42
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   10
               Left            =   2445
               Top             =   4380
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   10
                  Left            =   1860
                  TabIndex        =   44
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   10
                  Left            =   60
                  TabIndex        =   43
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   450
               Left            =   15
               Top             =   15
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   794
               ForeColor       =   16711680
               Caption         =   "出荷数"
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
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   10
               Left            =   15
               Top             =   465
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   1
                  Left            =   15
                  TabIndex        =   5
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   20
               Left            =   15
               Top             =   6555
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   15
                  Left            =   15
                  TabIndex        =   61
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   21
               Left            =   15
               Top             =   6120
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   14
                  Left            =   15
                  TabIndex        =   57
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   22
               Left            =   15
               Top             =   5685
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   13
                  Left            =   15
                  TabIndex        =   53
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   23
               Left            =   15
               Top             =   5250
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   12
                  Left            =   15
                  TabIndex        =   49
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   24
               Left            =   15
               Top             =   4815
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   11
                  Left            =   15
                  TabIndex        =   45
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   25
               Left            =   1050
               Top             =   4815
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   11
                  Left            =   285
                  TabIndex        =   46
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   11
               Left            =   2445
               Top             =   4815
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   11
                  Left            =   1860
                  TabIndex        =   48
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   11
                  Left            =   60
                  TabIndex        =   47
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   26
               Left            =   1050
               Top             =   5250
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   12
                  Left            =   285
                  TabIndex        =   50
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   12
               Left            =   2445
               Top             =   5250
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   12
                  Left            =   1860
                  TabIndex        =   52
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   12
                  Left            =   60
                  TabIndex        =   51
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   27
               Left            =   1050
               Top             =   5685
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   13
                  Left            =   285
                  TabIndex        =   54
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   13
               Left            =   2445
               Top             =   5685
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   13
                  Left            =   1860
                  TabIndex        =   56
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   13
                  Left            =   60
                  TabIndex        =   55
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   28
               Left            =   1050
               Top             =   6120
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   14
                  Left            =   285
                  TabIndex        =   58
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   14
               Left            =   2445
               Top             =   6120
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   14
                  Left            =   1860
                  TabIndex        =   60
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   14
                  Left            =   60
                  TabIndex        =   59
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   29
               Left            =   1050
               Top             =   6555
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   15
                  Left            =   285
                  TabIndex        =   62
                  Top             =   30
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   15
               Left            =   2445
               Top             =   6555
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   767
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
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   15
                  Left            =   1860
                  TabIndex        =   64
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "00"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   15
                  Left            =   60
                  TabIndex        =   63
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
                  cSize           =   1
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
            End
         End
         Begin vsFlexLib.vsFlexArray VsGrid1 
            Height          =   7050
            Left            =   90
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   540
            Width           =   7860
            _Version        =   196608
            _ExtentX        =   13864
            _ExtentY        =   12435
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
            FormatString    =   "№|品　番                   |背番号    |収容数 |箱 数  |受注数  |分割納入数|完納予定日|予定便|"
            Rows            =   16
            BackColor       =   12648447
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   0
            RowHeightMin    =   440
            GridLines       =   3
            SelectionMode   =   1
            BackColorAlternate=   12648447
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   435
            Left            =   7965
            Top             =   45
            Width           =   4800
            _ExtentX        =   8467
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "出荷日 - 便"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   180
            Begin CisText_V60.CisText iB1_SBin 
               Height          =   375
               Left            =   3810
               TabIndex        =   4
               Top             =   45
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   661
               cFormat         =   "00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "00"
               MaxLength       =   2
               cDataType       =   1
               cDataReplace    =   1
               cFaZero         =   0
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   2
            End
            Begin CisYMD_With_Btn.CisYMDwB iB1_SYmd 
               Height          =   330
               Left            =   1995
               TabIndex        =   3
               Top             =   60
               Width           =   1770
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
               Object.Width           =   1770
               cSize           =   1
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
      IMEMode         =   3  'ｵﾌ固定
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   -1320
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   900
      Top             =   -1380
   End
   Begin VB.Timer Timer2 
      Left            =   1170
      Top             =   -1380
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 出荷実績入力 [個別]  】"
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
      cPositionX      =   3000
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14280
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
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
         Height          =   315
         Left            =   11790
         Top             =   30
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
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
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   661
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "XXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cBoderWidth     =   3
         cFont3D         =   2
         cAlingnment     =   7
      End
   End
   Begin VB.PictureBox Back_HInfo 
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
      Height          =   1020
      Left            =   4695
      ScaleHeight     =   960
      ScaleWidth      =   9390
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   1335
      Width           =   9450
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   840
         Left            =   3615
         Top             =   60
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1482
         ForeColor       =   0
         Caption         =   " 納 入 日 - 便"
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
         Begin Cis3D_v60.CIS3D B1lb_NYmd 
            Height          =   360
            Left            =   90
            Top             =   390
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "9999/99/99"
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
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D B1lb_NBin 
            Height          =   360
            Left            =   1470
            Top             =   390
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "99"
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
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   840
         Left            =   60
         Top             =   60
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   1482
         ForeColor       =   0
         Caption         =   "納  入 先"
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
         Begin Cis3D_v60.CIS3D B1lb_Tornm 
            Height          =   360
            Left            =   60
            Top             =   390
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXX:XXXXXXXXX1XXXXXXXXX2"
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   840
         Left            =   5490
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1482
         ForeColor       =   0
         Caption         =   "受注区"
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
         Begin Cis3D_v60.CIS3D B1lb_JKbn 
            Height          =   360
            Left            =   75
            Top             =   390
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   635
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
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   840
         Left            =   6300
         Top             =   60
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   1482
         ForeColor       =   0
         Caption         =   "出荷管理部署"
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
         Begin Cis3D_v60.CIS3D B1lb_SyukaK 
            Height          =   360
            Left            =   60
            Top             =   390
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXX:XXXXXXXXX1XXXXXXXXX2"
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
   End
End
Attribute VB_Name = "CXJ0220"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   出荷実績入力 [個別]
'**       フォームID    :   CXJ0220
'**       処理概要      :
'**
'**       作  成  日    :   2004/03/27  By CIS
'**       変  更  日    :   2004/04/16  By CIS - 1.受注管理行NOの追加対応＆分納時の受注入力ワーク出力修正
'**       　  　  　                             2.自社伝票発行＆自社伝票ファイル出力の追加
'**       変  更  日    :   2004/04/21  By CIS - QRﾀｯﾁﾘｰﾀﾞｰ機能追加
'**       変  更  日    :   2004/06/28  By CIS - 受領書・出荷時発行の対応
'**       変  更  日    :   2004/09/01  By CIS - 収容数がゼロの時、異常終了する対応
'**       変  更  日    :   2006/11/22  By CIS - 分納時に,受領書の納入数が変更されてない点を修正
'**       変  更  日    :   2006/11/28  By CIS - カレンダマスタ読込時 取引先区分＝発注を受注に
'**       変  更  日    :   2008/04/10  By CIS - 背番号8桁
'**       変  更  日    :   2009/12/15  By CIS - 土岐津Ver(出庫ﾃﾞｰﾀ(実績管理ﾃｰﾌﾞﾙ)作成)
'**       変  更  日    :   2010/09/02  By CIS - 分割時の納入番号変更　廃止
'**       変  更  日    :   2010/12/28  By CIS - 桁溢れｴﾗｰ対応(wInt → wLong)
'**       変  更  日    :   2011/10/17  By CIS - 収容数０対応（実績管理テーブル）
'**       変  更  日    :   2011/10/17  By CIS - 出荷数/分納数　6桁に変更
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl

    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替

    Dim CisVsGrid               As New CisVsGrid3

    Const mMinLine              As Integer = 1
    Const mMaxLine              As Integer = 15
    Dim wLine                   As Integer

    Dim mHeadProc               As String
    Dim mSaveHeadProc           As String

    Dim mBunCount               As Long
    Dim wEndChk                 As Boolean

    Dim wInt                    As Integer
    Dim wLong                   As Long      '2010/12/28
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    If Rv_NCall = "" Then
    '   #-------------------#
    '   # 二 重 起 動 防 止  #
    '   #-------------------#
        If Not Dupli_Start Then End
    '   #-------------------#
    '   # 初 期 内 容 取 得  #
    '   #-------------------#
        gConnectCount = 4
        If Not IniGet Then End
    End If

    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    QRDummy.Left = -1000

    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止

    H4_Area1.Move H1_Area1.Left, H1_Area1.Top
    
    Set QT10QSR_Comm = PDQComm2
    Call QT10QSRInfoGet
    
    If QT10QSR_UseFlg Then
       IRN_NKRead = 1
    Else
       IRN_NKRead = 0
    End If
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 440
        .PatanMax = 1
        .SelectCol = 10
        .InitGet = False
        .Init
    End With
    Call GridInit

'   #-------------------#
'   #  項 目 初 期 化   #
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
    If Rv_NCall <> "K" Then
        K_Sykbnm.Tag = "2"
        Call PB_CHG_Click
        ProcHB = "H1"
        Call DispChange(ProcHB)
    End If
    K_Sykbnm.Caption = "【 指示書 】"

'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)

    If Rv_NCall = "" Then
       Me.Show
    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

    If Rv_NCall = "" Then
'*--------------------------------------------------<< 2004/06/28 Insert Start >>
        gCHT_DelKbn53 = 0
'*--------------------------------------------------<< 2004/06/28 Insert End >>
        If InitProc Then        ' 初期処理（データ削除）
            wEndChk = True
        Else
            wEndChk = False
            Call PB_END_Click
        End If
    End If

    FormAct = Not FormAct

    If Rv_NCall = "K" Then
       K_Sykbnm.Tag = "2"
       Call PB_CHG_Click
       
       Back_BunCount.Visible = True
       mBunCount = 1
       Call ContinuChk
    End If
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
           Case vbKeyF6:        If PB_CHG.Visible Then Call PB_CHG_Click    '【切替】
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyPageUp:    If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown:  If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      '【入力】
                                    Call PB_ENT_Click
                                End If
           Case Else:           Exit Function
    End Select

    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    If Rv_NCall = "" Then
'   【 後方処理 】
    '   #### 終了処理 ####
'*---------------------------------------------<< 2004/04/16 Update Start >>
'       If wEndChk Then Call EndProc
        Call EndProc
'*---------------------------------------------<< 2004/04/16 Update End >>
        Call CisDB.DBDISConnect                         ' ﾃﾞｰﾀﾍﾞｰｽ切断
        Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
        End
    Else
        Call QT10QSR_Close
        Timer1.Enabled = False
        Me.Hide
    End If
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
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    Call CisFrm.MousePT(11)

    If ProcHB = "B1" Then
        If Rv_NCall = "K" Then
           mBunCount = mBunCount + 1
           If ContinuChk Then
              GoTo PB_CAN_Ed
           Else
              Call PB_END_Click
           End If
        End If
        ProcHB = mHeadProc
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        GoTo PB_CAN_Ed
    End If

PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    Call DBInput("Back")
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    Call DBInput("Next")
    Call CisFrm.MousePT(1)
End Sub
'+------------------+
'+    切替   (F6)   +
'+------------------+
Private Sub PB_CHG_Click()
    gLong = CisFun.Val2(K_Sykbnm.Tag)
    
    gLong = gLong + 1
    If gLong > 2 Then gLong = 1
    
    mSaveHeadProc = ""
    mSaveHeadProc = "H1"
    If IRN_NKRead = 1 Then
       mSaveHeadProc = mSaveHeadProc & "H4"
    Else
       mSaveHeadProc = mSaveHeadProc & "  "
    End If
    
    mHeadProc = ""
    gInt = 1
    Do Until Trim(mHeadProc) <> ""
        mHeadProc = Mid(mSaveHeadProc, gLong + (gLong - 1), 2)
        If Trim(mHeadProc) <> "" Then
           K_Sykbnm.Tag = gLong
           Exit Do
        End If
        gLong = gLong + 1
        If gLong > 2 Then gLong = 1
    Loop
    
    H1_Area1.Visible = False
    H4_Area1.Visible = False
    Select Case mHeadProc
        Case "H4"
            K_Sykbnm.Caption = "【 QRﾀｯﾁﾘｰﾀﾞ 】"
            H4_Area1.Visible = True
        Case Else
            K_Sykbnm.Caption = "【 指示書 】"
            H1_Area1.Visible = True
    End Select
    ProcHB = mHeadProc
    Call DispChange(ProcHB)
    
    iH1_SJNo = ""
    Call HeadBodyClear("B")
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iB1_SYmd" Then iB1_SYmd.ShowCalender
    If PB_Look.Tag Like "iB1_BYmd*" Then iB1_BYmd(CisFun.Val2(Mid(PB_Look.Tag, 9))).ShowCalender
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*----------------------------------------------*
'*      出 荷 日  ( iB1_SYmd )       *
'*----------------------------------------------*
Private Sub iB1_SYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_SYmd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*----------------------------------------------*
'*       出荷数  ( iB1_NSuryo )                 *
'*----------------------------------------------*
Private Sub iB1_NSuryo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        For wLine = Index + 1 To mMaxLine Step 1
            If iB1_NSuryo(wLine).Visible Then
               iB1_NSuryo(wLine).SetFocus
               Exit For
            End If
        Next wLine
        Exit Sub
    End If
    If KeyCode = vbKeyUp Then
        For wLine = Index - 1 To mMinLine Step -1
            If iB1_NSuryo(wLine).Visible Then
               iB1_NSuryo(wLine).SetFocus
               Exit For
            End If
        Next wLine
        Exit Sub
    End If
End Sub
Private Sub iB1_NSuryo_LostFocus(Index As Integer)
    If CisFun.Val2(iB1_Suryo(Index).Tag) = 0 Then
       iB1_Suryo(Index) = CisFun.RSetFld(CisFun.Val2(iB1_NSuryo(Index).Tag) - CisFun.Val2(iB1_NSuryo(Index)), 0, iB1_Suryo(1).cFormat)
    End If
End Sub
'*----------------------------------------------*
'*       分割納入数  ( iB1_Suryo )              *
'*----------------------------------------------*
Private Sub iB1_Suryo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        For wLine = Index + 1 To mMaxLine Step 1
            If iB1_Suryo(wLine).Visible Then
               iB1_Suryo(wLine).SetFocus
               Exit For
            End If
        Next wLine
        Exit Sub
    End If
    If KeyCode = vbKeyUp Then
        For wLine = Index - 1 To mMinLine Step -1
            If iB1_Suryo(wLine).Visible Then
               iB1_Suryo(wLine).SetFocus
               Exit For
            End If
        Next wLine
        Exit Sub
    End If
End Sub
Private Sub iB1_Suryo_LostFocus(Index As Integer)
    If CisFun.Val2(iB1_Suryo(Index)) <> 0 Then
       iB1_NSuryo(Index) = CisFun.RSetFld(CisFun.Val2(iB1_NSuryo(Index).Tag) - CisFun.Val2(iB1_Suryo(Index)), 0, iB1_NSuryo(1).cFormat)
    End If
End Sub
'*----------------------------------------------*
'*       完納予定日  ( iB1_BYmd )       *
'*----------------------------------------------*
Private Sub iB1_BYmd_GotFocus(Index As Integer)
    PB_Look.Tag = ActiveControl.Name & Index
    PB_Look.Visible = True
End Sub
Private Sub iB1_BYmd_LostFocus(Index As Integer)
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
Private Sub iB1_BYmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        For wLine = Index + 1 To mMaxLine Step 1
            If iB1_BYmd(wLine).Visible Then
               iB1_BYmd(wLine).SetFocus
               Exit For
            End If
        Next wLine
        Exit Sub
    End If
    If KeyCode = vbKeyUp Then
        For wLine = Index - 1 To mMinLine Step -1
            If iB1_BYmd(wLine).Visible Then
               iB1_BYmd(wLine).SetFocus
               Exit For
            End If
        Next wLine
        Exit Sub
    End If
End Sub
'*----------------------------------------------*
'*       完納予定便  ( iB1_BBin )               *
'*----------------------------------------------*
Private Sub iB1_BBin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        For wLine = Index + 1 To mMaxLine Step 1
            If iB1_BBin(wLine).Visible Then
               iB1_BBin(wLine).SetFocus
               Exit For
          End If
        Next wLine
        Exit Sub
    End If
    If KeyCode = vbKeyUp Then
        For wLine = Index - 1 To mMinLine Step -1
            If iB1_BBin(wLine).Visible Then
               iB1_BBin(wLine).SetFocus
               Exit For
          End If
        Next wLine
        Exit Sub
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
    Dim H1Mode          As Boolean
    Dim H1Color         As Long
    Dim H4Mode          As Boolean
    Dim H4Color         As Long
    Dim B1Mode          As Boolean
    Dim B1Color         As Long

    H1Mode = False:     H1Color = gPLostSel
    H4Mode = False:     H4Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel

'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus

    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "H4":      H4Mode = True:  H4Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
    End Select

    H1_Area1.Enabled = H1Mode:  H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode:  B1_Area1.BackColor = B1Color
    H4_Area1.Enabled = H4Mode:  H4_Area1.BackColor = H4Color

    PB_CAN.Visible = B1Mode
    PB_BACK.Visible = H1Mode
    PB_NEXT.Visible = H1Mode

    If Rv_NCall = "" Then
    Else
       PB_END.BtnCaption = "戻る"
    End If
    
    If IRN_NKRead = 0 Then
       PB_CHG.Visible = False
    Else
       PB_CHG.Visible = H1Mode Or H4Mode
    End If
    PB_Look.Visible = False
    
    If ProcHB = "H4" Then
       Call QT10QSR_Open
       QRMsg.BackColor = &HC0FFC0
       QRMsg.Caption = "QRﾀｯﾁﾘｰﾀﾞ読取中"
    Else
       Call QT10QSR_Close
       QRMsg.BackColor = &HFFC0C0
       QRMsg.Caption = "QRﾀｯﾁﾘｰﾀﾞ読取"
    End If
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False

End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String

    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then .text = "": .Enabled = True
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
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
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
    '   <><><> 一括処理時 <><><>
        If Rv_NCall = "K" Then
           mBunCount = mBunCount + 1
           If ContinuChk Then
              GoTo ReturnPress_Ed
           Else
              Call PB_END_Click
           End If
        End If
    '   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Call HeadBodyClear("B")
        ProcHB = "H1"
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

'( 出荷指示書№ )
    If CisFun.Val2(iH1_SJNo) = 0 Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  指示書№を入力して下さい  "
            .MB_Button = Error
            .MBOX
        End With
        iH1_SJNo.SetFocus
        Exit Function
    End If

    If Not DBInput("INV") Then Exit Function

    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    Dim wClnFlg      As Boolean

    If iB1_SYmd = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    出荷日を入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iB1_SYmd.SetFocus
        Exit Function
    End If
    If Not iB1_SYmd.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    出荷日を正しい日付にて入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iB1_SYmd.SetFocus
        Exit Function
    End If
    If iB1_SBin.Visible And Trim(B1lb_JKbn.Tag) = "1" Then
        If CisFun.Val2(iB1_SBin) = 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    出荷便を入力して下さい                 "
                .MB_Button = Error
                .MBOX
            End With
            iB1_SBin.SetFocus
            Exit Function
        End If
        If B1lb_NYmd.Tag & B1lb_NBin > iB1_SYmd & iB1_SBin Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    出荷日・便は納入日・便以降にて入力して下さい         "
                .MB_Button = Error
                .MBOX
            End With
            iB1_SYmd.SetFocus
            Exit Function
        End If
    Else
        If B1lb_NYmd.Tag > iB1_SYmd Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    出荷日は納入日以降にて入力して下さい         "
                .MB_Button = Error
                .MBOX
            End With
            iB1_SYmd.SetFocus
            Exit Function
        End If
    End If
'--------------------------------------------------------------------------2006/11/28 Update Start
    'gInt = Kdou_Check(iB1_SYmd)
    gInt = Kdou_Check(iB1_SYmd, 0)
'--------------------------------------------------------------------------2006/11/28 End
    If gInt = 1 Then
    '休日
        CisFun.MB_Lines = 4
        CisFun.MB_MSG(1) = "     出荷日が非稼働日です。     "
        CisFun.MB_MSG(3) = "     　よろしいですか？      "
        CisFun.MB_Button = OK_CAN
        If Not CisFun.MBOX Then
            iB1_SYmd.SetFocus
            Exit Function
        End If
    End If
    If gInt = 2 Then
    'カレンダ未登録
        CisFun.MB_Lines = 4
        CisFun.MB_MSG(1) = "     " & RTrim(iB1_SYmd.cYear) & "年" & RTrim(iB1_SYmd.cMonth) & "月のカレンダ未登録！     "
        CisFun.MB_MSG(3) = "     　　(　カレンダマスタ　)　"
        CisFun.MB_Button = Error
        CisFun.MBOX
        iB1_SYmd.SetFocus
        Exit Function
    End If

    Set gObj = Nothing

    For gCnt1 = 1 To 15

        If iB1_Suryo(gCnt1).Visible Then
           gLong = CisFun.Val2(vsGrid1.TextMatrix(gCnt1, 5))

           If CisFun.Val2(iB1_NSuryo(gCnt1)) = 0 And _
              CisFun.Val2(iB1_Suryo(gCnt1)) = 0 Then
              gStr = "出荷数又は分割納入数を入力して下さい"
              Set gObj = iB1_NSuryo(gCnt1)
              Exit For
           End If
           If CisFun.Val2(iB1_NSuryo(gCnt1)) > gLong Then
              gStr = "出荷数は納入数以下にて指定して下さい"
              Set gObj = iB1_NSuryo(gCnt1)
              Exit For
           End If
           If CisFun.Val2(iB1_Suryo(gCnt1)) > gLong Then
              gStr = "分割納入数は納入数以下にて指定して下さい"
              Set gObj = iB1_Suryo(gCnt1)
              Exit For
           End If

           If CisFun.Val2(iB1_Suryo(gCnt1)) <> 0 Then
'''              If Trim(iB1_BYmd(gCnt1)) = "" Then
'''                 gStr = "分割納入数入力時、完納予定日を入力して下さい"
'''                 Set gObj = iB1_BYmd(gCnt1)
'''                 Exit For
'''              End If
              
              '完納予定日未入力時､翌稼働日-'01'を自動ｾｯﾄ---土岐津Ver
              If Trim(iB1_BYmd(gCnt1)) = "" Then
                 gStr = Mid(B1lb_NYmd, 1, 4) & Mid(B1lb_NYmd, 6, 2) & Mid(B1lb_NYmd, 9, 2)
                 wClnFlg = False
                 Do Until wClnFlg
                    '納入日 + 1 稼働日ﾁｪｯｸ
                    With CisFun
                         .Date_RsFormat = fYMD
                         .Date_Add gStr, D, 1, ResultDate
                         gStr = .Date_Result
                    End With
                    gInt = Kdou_Check(gStr, 0)
                    If gInt = 0 Then
                       wClnFlg = True
                    End If
                 Loop
                 iB1_BYmd(gCnt1) = gStr
                 iB1_BBin(gCnt1) = "01"
              End If
              
              If Not iB1_BYmd(gCnt1).cChkResult Then
                 gStr = "完納予定日を正しい日付にて入力して下さい"
                 Set gObj = iB1_BYmd(gCnt1)
                 Exit For
              End If
'--------------------------------------------------------------------------------2006/11/28 Update Start
              'gInt = Kdou_Check(iB1_BYmd(gCnt1))
              gInt = Kdou_Check(iB1_BYmd(gCnt1), 0)
'--------------------------------------------------------------------------------2006/11/28 End
              If gInt = 1 Then
                '休日
                 CisFun.MB_Lines = 4
                 CisFun.MB_MSG(1) = "       " & StrConv(gCnt1, vbWide) & "行目 完納予定日が非稼働日です。     "
                 CisFun.MB_MSG(3) = "     　   よろしいですか？          "
                 CisFun.MB_Button = OK_CAN
                 If Not CisFun.MBOX Then
                    iB1_BYmd(gCnt1).SetFocus
                    Exit Function
                 End If
              End If
              If gInt = 2 Then
                'カレンダ未登録
                 CisFun.MB_Lines = 4
                 CisFun.MB_MSG(1) = "     " & RTrim(iB1_BYmd(gCnt1).cYear) & "年" & RTrim(iB1_BYmd(gCnt1).cMonth) & "月のカレンダ未登録！     "
                 CisFun.MB_MSG(3) = "     　　(　カレンダマスタ　)　"
                 CisFun.MB_Button = Error
                 CisFun.MBOX
                 iB1_BYmd(gCnt1).SetFocus
                 Exit Function
              End If
              If iB1_BBin(gCnt1).Visible Then
                 If Trim(B1lb_JKbn.Tag) = "1" And CisFun.Val2(iB1_BBin(gCnt1)) = 0 Then
                    gStr = "分割納入数入力時、完納予定便を入力して下さい"
                    Set gObj = iB1_BBin(gCnt1)
                    Exit For
                 End If
                 If iB1_SYmd & iB1_SBin > iB1_BYmd(gCnt1) & iB1_BBin(gCnt1) Then
                    gStr = "完納予定日・便は出荷日・便以降にて入力して下さい"
                    Set gObj = iB1_BYmd(gCnt1)
                    Exit For
                 End If
              Else
                 If iB1_SYmd > iB1_BYmd(gCnt1) Then
                    gStr = "完納予定日は出荷日以降にて入力して下さい"
                    Set gObj = iB1_BYmd(gCnt1)
                    Exit For
                 End If
              End If
           End If
        End If
    Next gCnt1

    If Not gObj Is Nothing Then
        With CisFun
            .MB_Lines = 3
            .MB_MSG(2) = "          " & StrConv(gCnt1, vbWide) & "行目 " & gStr & "                "
            .MB_Button = Error
            .MBOX
        End With
        gObj.SetFocus
        Exit Function
    End If
    Body1Chk = True
End Function
'+---------------------------------------+
'+  一括からの連携処理
'+---------------------------------------+
Private Function ContinuChk() As Boolean
    ContinuChk = False

    If mBunCount > Rv_STtlSu Then Exit Function
    Do Until Rv_SJNo(mBunCount) <> 0
       If mBunCount >= Rv_STtlSu Then
          Call PB_END_Click
          Exit Function
       End If
       mBunCount = mBunCount + 1
    Loop

    Bun_Count = CisFun.RSetFld(mBunCount, 0, "###") & " / " & _
                CisFun.RSetFld(Rv_STtlSu, 0, "###")
    iH1_SJNo = Rv_SJNo(mBunCount)
    ProcHB = "H1"
    Call PB_ENT_Click

    ContinuChk = True
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False

    Call HeadBodyClear("B")

    On Error GoTo DBInput_Err:

    Call SqlSelectSet(InputKey)
    If Not JYTRead(gSL_Select, 1) Then
        If InputKey = "Next" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    最終レコードです            "
                .MB_Title = "【次頁】"
                .MB_Button = OK
                .MBOX
            End With
            GoTo DBInput_Ed
        End If
        If InputKey = "Back" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    先頭レコードです            "
                .MB_Title = "【前頁】"
                .MB_Button = OK
                .MBOX
            End With
        End If
        If InputKey = "INV" Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "  該当データがありません  "
                .MB_Button = Error
                .MBOX
            End With
        End If
        GoTo DBInput_Ed
    Else
        iH1_SJNo = Format(JYT.出荷指示書NO, String(10, "0"))

        Call BodySet(InputKey)
        
        If JYT.出荷実績区分 <> 0 Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "          出荷実績済みです。            "
                .MB_Button = Error
                .MBOX
            End With
            Exit Function
        End If
        DBInput = True
    End If

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'+---------------------------
'+      SQL文ｾｯﾄ処理
'+---------------------------
Private Sub SqlSelectSet(InputKey As String)

    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  jyt.出荷指示書NO"
    gSL_Select = gSL_Select & ", jyt.出荷指示書行NO"
    gSL_Select = gSL_Select & ", jyt.納入日"
    gSL_Select = gSL_Select & ", jyt.納入便"
    gSL_Select = gSL_Select & ", jyt.納入先"
    gSL_Select = gSL_Select & ", isnull(trm.略称, '') 納入先名"
    gSL_Select = gSL_Select & ", jyt.受入"
    gSL_Select = gSL_Select & ", isnull(sym.値名称, '') 受注区分名"
    gSL_Select = gSL_Select & ", isnull(hnm.表示品番, '') 表示品番"
    gSL_Select = gSL_Select & ", jyt.背番号"
    gSL_Select = gSL_Select & ", jyt.収容数"
    gSL_Select = gSL_Select & ", jyt.枚数"
    gSL_Select = gSL_Select & ", jyt.納入数"
    gSL_Select = gSL_Select & ", jyt.出荷実績区分"
    gSL_Select = gSL_Select & ", isnull(HT.出荷管理部署,'') 出荷管理部署,ISNULL(TR2.略称,'') 出荷管理部署名"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  受注テーブル jyt"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = jyt.納入先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 0"
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sym"
    gSL_Select = gSL_Select & "    on sym.区分名称 = '受注区分'"
    gSL_Select = gSL_Select & "   and sym.値 = jyt.受注区分"
    gSL_Select = gSL_Select & "   and sym.区分タイプ = ''"
    gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm"
    gSL_Select = gSL_Select & "    on hnm.品番 = jyt.品番"
    gSL_Select = gSL_Select & "  Left Outer Join 品番取引先マスタ HT"
    gSL_Select = gSL_Select & "    ON  HT.品番 = jyt.品番"
    gSL_Select = gSL_Select & "    And HT.取引先 = jyt.納入先"
    gSL_Select = gSL_Select & "    And HT.受入 = jyt.受入"
    gSL_Select = gSL_Select & "    And HT.品目 = 0"
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR2"
    gSL_Select = gSL_Select & "    ON TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "    And TR2.取引先CD = HT.出荷管理部署"
    gSL_Select = gSL_Select & "    And TR2.社内区分 = 1"
    Select Case InputKey
        Case "Back"
            gSL_Select = gSL_Select & " where jyt.出荷指示書NO < " & CisFun.Val2(iH1_SJNo)
            gSL_Select = gSL_Select & "   and jyt.出荷実績区分 = 0"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  jyt.出荷指示書NO DESC"
            gSL_Select = gSL_Select & ", jyt.出荷指示書行NO"
        Case "Next"
            gSL_Select = gSL_Select & " where jyt.出荷指示書NO > " & CisFun.Val2(iH1_SJNo)
            gSL_Select = gSL_Select & "   and jyt.出荷実績区分 = 0"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  jyt.出荷指示書NO"
            gSL_Select = gSL_Select & ", jyt.出荷指示書行NO"
        Case Else
            gSL_Select = gSL_Select & " where jyt.出荷指示書NO = " & CisFun.Val2(iH1_SJNo)
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  jyt.出荷指示書行NO"
    End Select

End Sub
'
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Sub BodySet(SetType As String)
'
    With JYT
    ' 出荷日･初期表示
        If SetType = "INV" Then
            iB1_SYmd = .納入日
            If Rv_NCall = "K" And Trim(Rv_NYmd) <> "" Then
               iB1_SYmd = Rv_NYmd
            End If
            If Trim(JYT.納入便) <> "" Then
               iB1_SBin.Visible = True
               iB1_SBin = .納入便
               If Rv_NCall = "K" And Trim(Rv_NBin) <> "" Then
                  iB1_SBin = Rv_NBin
               End If
            Else
               iB1_SBin.Visible = False
            End If
        End If

        If Trim(.納入先名) = "" Then
           gStr = .納入先
        Else
           gStr = RTrim(.納入先) & ":" & .納入先名
        End If
        B1lb_Tornm = gStr

        B1lb_NYmd.Tag = JYT.納入日
        With CisFun
            .DateE_BefVal = JYT.納入日
            .DateEdit
            B1lb_NYmd = CisFun.DateE_AfVal
            If Trim(JYT.納入便) <> "" Then
                B1lb_NBin = JYT.納入便
            End If
        End With
        B1lb_JKbn = .受注区分名
    
        If Trim(.出荷管理部署名) = "" Then
           gStr = .出荷管理部署
        Else
           gStr = RTrim(.出荷管理部署) & ":" & RTrim(.出荷管理部署名)
        End If
        B1lb_SyukaK = gStr
        
    End With

    Call SqlSelectSet("")
    If JYTRead(gSL_Select) Then
        Call GridSet(SetType)
    End If

End Sub
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

'   ( 確認メッセージ )
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     出荷実績処理を行います            "
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Title = "実績更新処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err

    CisDB.ConnectNo = 0
'    Call CisDB.DBTran(TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始              2009/12/15

    gSL_Select = ""
    gSL_Select = gSL_Select & "select *"
    gSL_Select = gSL_Select & "  from 受注テーブル"
    gSL_Select = gSL_Select & " where 出荷指示書NO = " & CisFun.Val2(iH1_SJNo)
    If JYTRead(gSL_Select, , 1) Then
        Do Until Not JYT_RDSTS
'           >> 出荷テーブル出力
            Call ItemsClearSKT
            With SKT
                .出荷指示書NO = JYT.出荷指示書NO
                .出荷指示書行NO = JYT.出荷指示書行NO
                .出荷日 = iB1_SYmd
                .出荷便 = Format(iB1_SBin, "00")
                .出荷数 = CisFun.Val2(iB1_NSuryo(.出荷指示書行NO))
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'                .出荷枚数 = .出荷数 \ JYT.収容数
'                .出荷端数 = .出荷数 Mod JYT.収容数
''*---------------------------------------------<< 2004/04/16 Update Start >>
''               If .出荷端数 <> 0 Then .出荷数 = .出荷数 + 1
'                If .出荷端数 <> 0 Then .出荷枚数 = .出荷枚数 + 1
''*---------------------------------------------<< 2004/04/16 Update Start >>
'
                If JYT.収容数 <> 0 Then
                    .出荷枚数 = .出荷数 \ JYT.収容数
                    .出荷端数 = .出荷数 Mod JYT.収容数
                    If .出荷端数 <> 0 Then .出荷枚数 = .出荷枚数 + 1
                End If
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
                .分納数 = JYT.納入数 - .出荷数
            End With
            Call SKTInsert
'           >> 実績管理テーブル(出庫ﾃﾞｰﾀ)更新               '2009/12/15
            If Not SyukoInsert Then GoTo DBPut_Err
'           >> 受注テーブル更新
            If Not DBUpdate_JYT Then GoTo DBPut_Err
'           >> 分納判定
            If CisFun.Val2(iB1_Suryo(JYT.出荷指示書行NO)) <> 0 Then
'           >> 受注入力ワーク出力
                Call ItemsClearJNW
                Call DBSet_JNW
                Call JNWInsert
            End If
        Call JYTReadNext(1)
        Loop
'*--------------------------------------------------<< 2004/06/28 Insert Start >>
    '   >> 初回実行の判定
        If gCHT_DelKbn53 = 0 Then
            gCHT_DelKbn53 = 1
        '   >> 帳票出力テーブル削除
            If Not DBDelete_CHT Then GoTo DBPut_Err
        End If
    '   >> 帳票出力テーブル出力
        If Not DBInsert_CHT(CisFun.Val2(iH1_SJNo)) Then GoTo DBPut_Err
'*--------------------------------------------------<< 2004/06/28 Insert End >>
    End If
    Call JYTClose(1)

    CisDB.ConnectNo = 0
'    Call CisDB.DBTran(TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了           2009/12/15

    lb_JMai = Format(CisFun.Val2(lb_JMai) + 1, "#,###")

    If Rv_NCall = "K" Then
       Rv_SExec(mBunCount) = True
    End If

    On Error GoTo 0

    DBPut = True
    Exit Function

DBPut_Err:
    Call JYTClose(1)

    CisDB.ConnectNo = 0
    Call CisDB.DBTran(TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ破棄
    CisFun.ErrorBox
    End
End Function
'=================================================================== 2009/12/15 Start
'+-------------------------------------------------
'+      登録処理：実績管理テーブル---出庫ﾃﾞｰﾀ作成
'+-------------------------------------------------
Private Function SyukoInsert() As Boolean
    SyukoInsert = False
    
    '構成品番情報の取得
     With CisDB
         .SQL = "構成問合せ情報抽出"
         .StoadoCount = 4
         .ParaNo = 0: .ParaIO = Return用
         .ParaNo = 1: .ParaIO = Input用
         .ParaNo = 2: .ParaIO = Input用
         .ParaNo = 3: .ParaIO = Input用
         .ParaNo = 4: .ParaIO = Input用
     
         .ParaNo = 1: .ParaValue = RTrim(JYT.品番)
         .ParaNo = 2: .ParaValue = 0
         .ParaNo = 3: .ParaValue = 0
         .ParaNo = 4: .ParaValue = 0
     
         If Not .DBStored Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = " [出庫作成]構成マスタ情報の抽出に失敗しました。 "
                 .MB_Title = "抽出処理"
                 .MB_Button = Error
                 If .MBOX Then Exit Function
             End With
         End If
     End With
            
    '構成情報読み込み
    gSL_Select = "SELECT * FROM 構成問合せワーク "
    gSL_Select = gSL_Select & " WHERE 処理端末 = Host_name() "
'    gSL_Select = gSL_Select & " AND   レベル <> 0 "
    gSL_Select = gSL_Select & " ORDER BY 順位 "
    If KOLRead(gSL_Select, , 2) Then
        Do Until Not KOL_RDSTS
           '品番工順ﾏｽﾀ読み込み･･･棚卸ﾌﾟﾚｼｰﾄ発行[有]のみ対象
           gSL_Select = "SELECT * FROM 品番工順マスタ"
           gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(KOL.子品番) & "'"
           gSL_Select = gSL_Select & " AND 棚卸プレシート発行 = 0"
           gSL_Select = gSL_Select & " ORDER BY 品番,取引先,受入,工順 "
           If HKTRead(gSL_Select, , 3) Then
               Do Until Not HKT_RDSTS
                   Call ItemsClearJKT
                   
                   With CisDB
                       .SQL = "番号獲得"
                       .StoadoCount = 3
                       .ParaNo = 0: .ParaIO = Return用
                       .ParaNo = 1: .ParaIO = Input用
                       .ParaNo = 2: .ParaIO = OutPut用
                       .ParaNo = 3: .ParaIO = Return用
                   '
                       .ParaNo = 1: .ParaValue = 6
                   '
                       If Not .DBStored Then
                           CisFun.MB_Lines = 4
                           CisFun.MB_MSG(2) = "    実績管理NOの獲得に失敗しました            "
                           CisFun.MB_Title = "採番処理"
                           CisFun.MB_Button = Error
                           CisFun.MBOX
                           Exit Function
                       Else
                           .ParaNo = 2: JKT.実績管理NO = .ParaValue
                       End If
                   End With
                   JKT.実績区分 = 4
                   JKT.納入番号 = JYT.納入番号
                   JKT.納品書頁 = JYT.納品書頁
                   JKT.納品書行 = JYT.納品書行
                   JKT.分納区分 = 0
                   JKT.実績先 = HKT.取引先
                   JKT.受入 = HKT.受入
                   JKT.品番 = HKT.品番
                   If Trim(iB1_SYmd) <> "" Then
                       JKT.実績日 = iB1_SYmd
                       JKT.実績便 = iB1_SBin
                   Else
                       JKT.実績日 = JYT.納入日
                       JKT.実績便 = JYT.納入便
                   End If
                   If HKT.収容数 <> 0 Then
                      JKT.収容数 = HKT.収容数
                   Else
                      JKT.収容数 = JYT.収容数
                   End If
                   
                   
'=========================================== 2010/12/28 Start
'                   wInt = SKT.出荷数 * KOL.構成数
'                   JKT.実績枚数 = wInt \ JYT.収容数
'                   JKT.実績端数 = wInt Mod JKT.収容数
'                   JKT.実績数 = wInt

'==================================== 2011/10/17 Start
'                   wLong = SKT.出荷数 * KOL.構成数
'                   JKT.実績枚数 = wLong \ JYT.収容数
'                   JKT.実績端数 = wLong Mod JKT.収容数
'                   JKT.実績数 = wLong
                   If JYT.収容数 <> 0 Then
                      wLong = SKT.出荷数 * KOL.構成数
                      JKT.実績枚数 = wLong \ JYT.収容数
                      JKT.実績端数 = wLong Mod JKT.収容数
                      JKT.実績数 = wLong
                   Else
                      wLong = SKT.出荷数 * KOL.構成数
                      JKT.実績枚数 = 0
                      JKT.実績端数 = 0
                      JKT.実績数 = wLong
                   End If
'==================================== 2011/10/17 End
'=========================================== 2010/12/28 End
                   
                   
'                   JKT.実績枚数 = SKT.出荷数 \ JYT.収容数
'                   JKT.実績端数 = SKT.出荷数 Mod JKT.収容数
'                   JKT.実績数 = SKT.出荷数
                   JKT.有無償区分 = 1
                   JKT.入力種類 = "S"
                   JKT.工順 = HKT.工順
                   JKT.出荷指示書NO = JYT.出荷指示書NO
                   JKT.出荷指示書行NO = JYT.出荷指示書行NO
                   JKT.作成者 = "CXJ0220"
                   
                   If KOL.レベル = 0 Then
                      JKT.親品番 = 1
                   Else
                      JKT.親品番 = 0
                   End If
                           
                   Call JKTInsert
                                               
                   Call HKTReadNext(3)
              Loop
              Call HKTClose(3)
           End If
           
           Call KOLReadNext(2)
        Loop
        Call KOLClose(2)
    End If
    
    SyukoInsert = True
End Function
'=================================================================== 2009/12/15 End
'+------------------------------------+
'+      セット - 受注入力ワーク
'+------------------------------------+
Private Function DBSet_JNW()
    DBSet_JNW = False

    Call ItemsClearJNW
    With JNW
        .端末番号 = "' + substring(host_name(),1,20) + '"
        .受注管理NO = JYT.受注管理NO
'*---------------------------------------------<< 2004/04/16 Insert Start >>
        .受注管理行NO = JYT.受注管理行NO
'*---------------------------------------------<< 2004/04/16 Insert End >>
        .受注区分 = JYT.受注区分
        .受注日 = JYT.受注日
        .受注便 = JYT.受注便
        .納入日 = iB1_BYmd(JYT.出荷指示書行NO)
        .納入便 = Format(iB1_BBin(JYT.出荷指示書行NO), "00")
        .納入先 = JYT.納入先
        .出荷予定日 = .納入日
        .出荷予定便 = .納入便
        .受入 = JYT.受入
        .機番 = ""
'*---------------------------------------------<< 2004/04/16 Update Start >>
'       .納入番号 = JYT.納入番号
'       .納品書頁 = JYT.納品書頁
'       .納品書行 = JYT.納品書行
'       .納品書発行区分 = 0
'================================ 2010/09/02 Start
'        If JYT.納品書発行区分 = 1 Then
'            gInt = CisFun.Len2(JYT.納入番号)
'            gStr = CisFun.Mid2(JYT.納入番号, gInt, 1)
'            gStr = CisFun.Val2(gStr) + 1
'            .納入番号 = CisFun.Mid2(JYT.納入番号, 1, gInt - 1) & gStr
'            .納品書発行区分 = 0
'        Else
'            .納入番号 = JYT.納入番号
'            .納品書発行区分 = 9
'        End If
        .納入番号 = JYT.納入番号
        .納品書頁 = JYT.納品書頁
        .納品書行 = JYT.納品書行
        .納品書発行区分 = 9
'================================ 2010/09/02 End
        
        .納品書頁 = 0
        .納品書行 = 0
'*---------------------------------------------<< 2004/04/16 Update End >>
        .分納区分 = 1
        .分納回数 = JYT.分納回数 + 1
        .品番 = JYT.品番
        .背番号 = JYT.背番号
        .収容数 = JYT.収容数
        .納入数 = SKT.分納数
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'       .枚数 = .納入数 \ .収容数
'       .端数 = .納入数 Mod .収容数
'       If .端数 <> 0 Then .枚数 = .枚数 + 1
'
        If .収容数 <> 0 Then
            .枚数 = .納入数 \ .収容数
            .端数 = .納入数 Mod .収容数
            If .端数 <> 0 Then .枚数 = .枚数 + 1
        End If
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
        .出荷指示書NO = 0
        .出荷指示書行NO = 0
        .抜取枚数 = 0
        .エフ発行枚数 = 0
        .エフ発行端数 = 0
        .作成種別 = "G"
        .累積区分 = 0
'*---------------------------------------------<< 2004/04/16 Update Start >>
'       .納品書管理NO = .納品書管理NO
        .納品書管理NO = 0
'*---------------------------------------------<< 2004/04/16 Update End >>
        .元出荷指示書NO = JYT.出荷指示書NO
        .元出荷指示書行NO = JYT.出荷指示書行NO
    End With

    DBSet_JNW = True
End Function
'+--------------------------------
'+      更新処理：受注テーブル
'+--------------------------------
Private Function DBUpdate_JYT(Optional DBNo As Integer = 0) As Boolean
    DBUpdate_JYT = False
    gSL_Select = ""
    gSL_Select = gSL_Select & "update 受注テーブル"
    gSL_Select = gSL_Select & "   set 出荷実績区分 = 1"
'----------------------------------------------------------------------------2006/11/22 Insert Start
'    gSL_Select = gSL_Select & "       ,納入数 = " & SKT.分納数 & ""
    gSL_Select = gSL_Select & "       ,納入数 = " & SKT.出荷数 & ""
    gSL_Select = gSL_Select & "       ,枚数 = " & SKT.出荷枚数 & ""
    gSL_Select = gSL_Select & "       ,端数 = " & SKT.出荷端数 & ""
'----------------------------------------------------------------------------2006/11/22 Insert End
    gSL_Select = gSL_Select & " where 出荷指示書NO = " & JYT.出荷指示書NO
    gSL_Select = gSL_Select & "   and 出荷指示書行NO = " & JYT.出荷指示書行NO
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBUpdate_JYT = True
End Function
'*--------------------------------------------------<< 2004/06/28 Insert Start >>
'+--------------------------------
'+      削除処理：帳票出力テーブル
'+--------------------------------
Private Function DBDelete_CHT(Optional DBNo As Integer = 0) As Boolean
    DBDelete_CHT = False
    gSL_Select = ""
    gSL_Select = gSL_Select & "delete"
    gSL_Select = gSL_Select & "  from 帳票出力テーブル"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and 帳票区分 in (53)"
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBDelete_CHT = True
End Function
'+--------------------------------
'+      登録処理：帳票出力テーブル
'+--------------------------------
Private Function DBInsert_CHT(SJNo As Long, Optional DBNo As Integer = 0) As Boolean
    DBInsert_CHT = False
    gSL_Select = ""
    gSL_Select = gSL_Select & "insert into 帳票出力テーブル"
    gSL_Select = gSL_Select & " (端末番号, 帳票区分,管理文字,検索情報,管理NO,管理行NO,補助番号,出力先,作成日) "
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  substring(host_name(),1,20)"
    gSL_Select = gSL_Select & ", 53"
    gSL_Select = gSL_Select & ", jyt.納入番号"
    gSL_Select = gSL_Select & ", '6'"
    gSL_Select = gSL_Select & ", jyt.納品書管理NO"
    gSL_Select = gSL_Select & ", 0"
    gSL_Select = gSL_Select & ", jyt.納品書頁"
    gSL_Select = gSL_Select & ", 0"
    gSL_Select = gSL_Select & ", getdate()"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  受注テーブル jyt"
    gSL_Select = gSL_Select & ", 品番取引先マスタ htm"
    gSL_Select = gSL_Select & ", 出荷テーブル skt"
    gSL_Select = gSL_Select & " where jyt.出荷指示書NO = " & SJNo
    gSL_Select = gSL_Select & "   and htm.品番 = jyt.品番"
    gSL_Select = gSL_Select & "   and htm.取引先 = jyt.納入先"
    gSL_Select = gSL_Select & "   and htm.受入 = jyt.受入"
    gSL_Select = gSL_Select & "   and htm.品目 = 0"
    gSL_Select = gSL_Select & "   and htm.受領書発行 = 1"
    gSL_Select = gSL_Select & "   and skt.出荷指示書NO = jyt.出荷指示書NO"
    gSL_Select = gSL_Select & "   and skt.出荷指示書行NO = jyt.出荷指示書行NO"
    gSL_Select = gSL_Select & "   and not exists"
    gSL_Select = gSL_Select & "("
    gSL_Select = gSL_Select & "select 'X'"
    gSL_Select = gSL_Select & "  from 帳票出力テーブル cht"
    gSL_Select = gSL_Select & " where cht.端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and cht.帳票区分 = 53"
    gSL_Select = gSL_Select & "   and cht.管理文字 = jyt.納入番号"
    gSL_Select = gSL_Select & "   and cht.検索情報 = '6'"
    gSL_Select = gSL_Select & "   and cht.管理NO = jyt.納品書管理NO"
    gSL_Select = gSL_Select & "   and cht.補助番号 = jyt.納品書頁"
    gSL_Select = gSL_Select & ")"
    gSL_Select = gSL_Select & " group by"
    gSL_Select = gSL_Select & "  jyt.納品書管理NO"
    gSL_Select = gSL_Select & ", jyt.納入番号"
    gSL_Select = gSL_Select & ", jyt.納品書頁"
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBInsert_CHT = True
End Function
'*--------------------------------------------------<< 2004/06/28 Insert End >>

'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 0, 0, "№")
        Call .FixedSet(中中, 左中, あり, 0, 1, "品　番")
        Call .FixedSet(中中, 左中, あり, 0, 2, "背番号")
        Call .FixedSet(中中, 右中, あり, 0, 3, "収容数")
        Call .FixedSet(中中, 右中, あり, 0, 4, "箱 数")
        Call .FixedSet(中中, 右中, あり, 0, 5, "納入数")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
'        Call .FixedSet(中中, 左中, なし, 0,9,"Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5")
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet(SetType As String)
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not JYT_RDSTS
        With vsGrid1
            ID = JYT.出荷指示書行NO
        '   >> No
            .TextMatrix(ID, 0) = JYT.出荷指示書行NO
        '   >> 品番
            .TextMatrix(ID, 1) = JYT.表示品番
        '   >> 背番号
            .TextMatrix(ID, 2) = " " & JYT.背番号
        '   >> 収容数
            .TextMatrix(ID, 3) = Format(JYT.収容数, "#,###")
        '   >> 枚数
            .TextMatrix(ID, 4) = Format(JYT.枚数, "#,###")
        '   >> 納入数
            .TextMatrix(ID, 5) = Format(JYT.納入数, "#,###")
        '   >> 出荷数：入力
            iB1_NSuryo(ID) = CisFun.RSetFld(JYT.納入数, 0, iB1_NSuryo(1).cFormat)
            iB1_NSuryo(ID).Tag = JYT.納入数
        End With
        Call JYTReadNext
    Loop
    Call JYTClose
    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
'   >> 入力制御？
    For Each gObj In iB1_Suryo
        gStr = Trim(vsGrid1.TextMatrix(gObj.Index, 0))
        If gStr = "" Then
           gObj.Visible = False
           iB1_NSuryo(gObj.Index).Visible = False
           iB1_BYmd(gObj.Index).Visible = False
           iB1_BBin(gObj.Index).Visible = False
        Else
           gObj.Visible = True
           iB1_NSuryo(gObj.Index).Visible = True
           iB1_BYmd(gObj.Index).Visible = True
           If iB1_SBin.Visible Then
              iB1_BBin(gObj.Index).Visible = True
           Else
              iB1_BBin(gObj.Index).Visible = False
           End If
        End If
    Next gObj

End Sub

'******************************************************************
'*      初期処理
'******************************************************************
Private Function InitProc() As Boolean
    InitProc = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select *"
    gSL_Select = gSL_Select & "  from 受注入力ワーク"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and 累積区分 = 0"
    If JNWRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 6
            .MB_MSG(1) = "    当端末処理分の受注未累積データが存在します。            "
            .MB_MSG(4) = "      『出荷実績入力（個別）』処理を終了しますか？"
            .MB_MSG(5) = "      いいえの場合は、未累積データを削除し続行します"
            .MB_Title = "未累積データの確認"
            .MB_Button = Yes_No
            If .MBOX Then Exit Function
        End With
    End If

    With CisDB
        On Error GoTo InitProc_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
'       << 受注入力テーブル >>
        .SQL = ""
        .SQL = .SQL & "delete"
        .SQL = .SQL & "  from 受注入力ワーク"
        .SQL = .SQL & " where 端末番号 = substring(host_name(),1,20)"
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        On Error GoTo 0
    End With

    InitProc = True
    Exit Function
InitProc_Err:
    CisFun.ErrorBox
    End
End Function
'==========================================================='
'   終了処理（入力データがあれば後方処理を行う）
'==========================================================='
Private Sub EndProc()

'   << 入力件数確認 >>
    gSL_Select = ""
    gSL_Select = gSL_Select & "select count(*) 件数"
    gSL_Select = gSL_Select & "  from 受注入力ワーク"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and"
    gSL_Select = gSL_Select & " (     出荷指示書NO = 0"
    gSL_Select = gSL_Select & "    or 納入番号 = ''"
    gSL_Select = gSL_Select & "    or 累積区分 = 0"
    gSL_Select = gSL_Select & " )"
    If JNWRead(gSL_Select, 1) Then
        If JNW.件数 > 0 Then
        '   出荷指示NO採番
            CXJ0110.Show vbModal
            Unload CXJ0110
            Set CXJ0110 = Nothing
            DoEvents
        '   出荷指示書発行
            Call CXJ0120LPrint(0)
'*---------------------------------------------<< 2004/04/16 Insert Start >>
        '   自社伝票発行
            Call CXJ0140LPrint(0)
        '   自社伝票ﾌｧｲﾙ出力
            Call CXJ0150TFile(0)
'*---------------------------------------------<< 2004/04/16 Insert End >>
'*---------------------------------------------<< 2004/06/28 Insert Start >>
        Else
            If gCHT_DelKbn53 = 1 Then
            '   自社伝票発行
                Call CXJ0140LPrint(0)
            End If
'*---------------------------------------------<< 2004/06/28 Insert End >>
        End If
    End If

End Sub
'+------------------------------+
'+  ﾊﾞｰｺｰﾄﾞ読取処理             *
'+------------------------------+
Private Sub PDQComm2_OnComm()
    
    gStr = QT10QSR_Event
    If gStr = "ERR" Then Exit Sub
    If gStr <> "RCV" Then Exit Sub
    
    If Not QT10QSR_RCVProc Then Exit Sub    ' CR が存在しない時抜ける

'( データ確認 )
    If Trim(QT10QSR_ReadData) = "" Then Exit Sub
    
    iH1_SJNo = Format(CisFun.Val2(QT10QSR_ReadData), iH1_SJNo.cFormat)
    
    QT10QSR_ReadData = ""
    
    If Not DBInput("INV") Then Exit Sub
    
    ProcHB = "B1"
    Call DispChange(ProcHB)
    
End Sub


