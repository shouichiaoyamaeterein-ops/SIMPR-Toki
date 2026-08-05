VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form CXH0520 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "出庫入力"
   ClientHeight    =   11235
   ClientLeft      =   -60
   ClientTop       =   585
   ClientWidth     =   15360
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
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
      Height          =   1005
      Left            =   375
      ScaleHeight     =   945
      ScaleWidth      =   7950
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   495
      Width           =   8010
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   855
         Left            =   60
         Top             =   60
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   1508
         ForeColor       =   16711680
         Caption         =   "出庫年月"
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
         Begin CisYM_V60.CisYM iH1_Ym 
            Height          =   375
            Left            =   30
            TabIndex        =   0
            Top             =   390
            Width           =   1350
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
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   855
         Left            =   1470
         Top             =   60
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   1508
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   120
            TabIndex        =   1
            Top             =   390
            Width           =   3210
            _ExtentX        =   5662
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D MSG_SKbn 
         Height          =   855
         Left            =   4815
         Top             =   60
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   1508
         ForeColor       =   16711680
         Caption         =   "出 庫 区 分"
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
         Begin VB.OptionButton OpH1_SKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "全部"
            Height          =   330
            Index           =   2
            Left            =   2145
            TabIndex        =   17
            Top             =   405
            Width           =   885
         End
         Begin VB.OptionButton OpH1_SKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "構成"
            Height          =   330
            Index           =   1
            Left            =   1110
            TabIndex        =   16
            Top             =   405
            Width           =   885
         End
         Begin VB.OptionButton OpH1_SKbn 
            BackColor       =   &H00C0FFC0&
            Caption         =   "単品"
            Height          =   330
            Index           =   0
            Left            =   135
            TabIndex        =   15
            Top             =   405
            Value           =   -1  'True
            Width           =   885
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   855
         Left            =   4815
         Top             =   60
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   1508
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
      Height          =   1290
      Left            =   10485
      ScaleHeight     =   1230
      ScaleWidth      =   4215
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   420
      Width           =   4275
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1695
         Top             =   630
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
         Left            =   2535
         Top             =   645
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
         Left            =   3360
         Top             =   630
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
         Left            =   1695
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
         Left            =   870
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
         Left            =   60
         Top             =   630
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
         Left            =   2535
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
         Left            =   3360
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
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   3
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
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 出 庫 入 力 】"
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
         Left            =   10110
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
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
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   345
         Left            =   0
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         BackColor       =   14737632
         ForeColor       =   16711935
         Caption         =   "XXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14220
         Top             =   60
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
         Left            =   7560
         Top             =   30
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
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   8790
      Left            =   360
      ScaleHeight     =   8730
      ScaleWidth      =   14355
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1755
      Width           =   14415
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   8565
         Left            =   60
         Top             =   60
         Width           =   14235
         _ExtentX        =   25109
         _ExtentY        =   15108
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
         cBoderStyle     =   1
         Begin VB.PictureBox B3_Area2 
            Height          =   8400
            Left            =   7845
            ScaleHeight     =   8340
            ScaleWidth      =   6210
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   90
            Width           =   6270
            Begin Cis3D_v60.CIS3D CIS3D10 
               Height          =   8310
               Left            =   15
               Top             =   15
               Width           =   6195
               _ExtentX        =   10927
               _ExtentY        =   14658
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cPositionY      =   120
               Begin Cis3D_v60.CIS3D CIS3D27 
                  Height          =   360
                  Left            =   0
                  Top             =   15
                  Width           =   6180
                  _ExtentX        =   10901
                  _ExtentY        =   635
                  BackColor       =   0
                  ForeColor       =   65535
                  Caption         =   "【　構 成 情 報　】"
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
                  cFont3D         =   1
                  cAlingnment     =   7
               End
               Begin vsFlexLib.vsFlexArray vsGrid2 
                  Height          =   7815
                  Left            =   105
                  TabIndex        =   18
                  TabStop         =   0   'False
                  Top             =   405
                  Width           =   6000
                  _Version        =   196608
                  _ExtentX        =   10583
                  _ExtentY        =   13785
                  _StockProps     =   228
                  Appearance      =   1
                  ConvInfo        =   1418783674
                  FormatString    =   "№|品　番　　　　　　　　　　|使用数    |展開区|||||"
                  Rows            =   23
                  Cols            =   14
                  BackColor       =   12648447
                  BackColorFixed  =   12632256
                  BackColorBkg    =   12632256
                  TextStyle       =   4
                  TextStyleFixed  =   3
                  FocusRect       =   2
                  HighLight       =   0
                  RowHeightMin    =   330
                  FillStyle       =   1
                  GridLines       =   2
                  GridColor       =   14737632
                  SelectionMode   =   2
               End
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00E0E0E0&
            Height          =   1335
            Index           =   0
            Left            =   3015
            ScaleHeight     =   1275
            ScaleWidth      =   4665
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   7065
            Width           =   4725
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   315
               Left            =   1380
               Top             =   300
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
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
               Begin Cis3D_v60.CIS3D B1lb_TSuryo 
                  Height          =   255
                  Index           =   1
                  Left            =   165
                  Top             =   30
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   450
                  BackColor       =   12648384
                  Caption         =   "##,###,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D Grid_RowHead 
               Height          =   315
               Index           =   1
               Left            =   30
               Tag             =   "計画数"
               Top             =   300
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               BackColor       =   12648384
               ForeColor       =   0
               Caption         =   "入 荷 数"
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
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D BLB_D 
               Height          =   300
               Index           =   18
               Left            =   1365
               Top             =   0
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   529
               BackColor       =   0
               ForeColor       =   65535
               Caption         =   "当月合計"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   1
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D11 
               Height          =   315
               Left            =   1380
               Top             =   615
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
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
               Begin Cis3D_v60.CIS3D B1lb_TSuryo 
                  Height          =   270
                  Index           =   2
                  Left            =   165
                  Top             =   30
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   476
                  BackColor       =   12640511
                  Caption         =   "##,###,###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D Grid_RowHead 
               Height          =   315
               Index           =   0
               Left            =   30
               Tag             =   "入荷数"
               Top             =   930
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               BackColor       =   16761024
               ForeColor       =   0
               Caption         =   "合計出庫数"
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
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D Grid_RowHead 
               Height          =   315
               Index           =   5
               Left            =   30
               Tag             =   "実績数"
               Top             =   615
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               BackColor       =   12640511
               ForeColor       =   0
               Caption         =   "入力出庫数"
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
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D17 
               Height          =   315
               Index           =   1
               Left            =   1380
               Top             =   930
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
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
               Begin Cis3D_v60.CIS3D B1lb_TSuryo 
                  Height          =   270
                  Index           =   3
                  Left            =   165
                  Top             =   30
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   476
                  BackColor       =   16761024
                  Caption         =   "##,###,###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D2 
               Height          =   315
               Left            =   3030
               Top             =   300
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   255
                  Index           =   1
                  Left            =   165
                  Top             =   30
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   450
                  BackColor       =   12648384
                  Caption         =   "##,###,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D SelectDay 
               Height          =   300
               Left            =   3015
               Top             =   0
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   529
               BackColor       =   0
               ForeColor       =   65535
               Caption         =   "日迄合計"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   1
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   315
               Left            =   3030
               Top             =   615
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   270
                  Index           =   2
                  Left            =   165
                  Top             =   30
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   476
                  BackColor       =   12640511
                  Caption         =   "##,###,###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D17 
               Height          =   315
               Index           =   0
               Left            =   3030
               Top             =   930
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   270
                  Index           =   3
                  Left            =   165
                  Top             =   30
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   476
                  BackColor       =   16761024
                  Caption         =   "##,###,###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D Grid_RowHead 
               Height          =   285
               Index           =   2
               Left            =   30
               Tag             =   "計画数"
               Top             =   15
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               ForeColor       =   0
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
         End
         Begin Cis3D_v60.CIS3D BLB_Sur 
            Height          =   8505
            Index           =   1
            Left            =   30
            Top             =   30
            Width           =   14160
            _ExtentX        =   24977
            _ExtentY        =   15002
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
            cAlingnment     =   1
            Begin Cis3D_v60.CIS3D DayColor 
               Height          =   390
               Left            =   90
               Top             =   75
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   688
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D UC_Week 
               Height          =   390
               Index           =   2
               Left            =   3000
               Top             =   75
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   688
               Caption         =   "火"
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
            Begin Cis3D_v60.CIS3D UC_Week 
               Height          =   390
               Index           =   3
               Left            =   3930
               Top             =   75
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   688
               Caption         =   "水"
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
            Begin Cis3D_v60.CIS3D UC_Week 
               Height          =   390
               Index           =   4
               Left            =   4860
               Top             =   75
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   688
               Caption         =   "木"
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
            Begin Cis3D_v60.CIS3D UC_Week 
               Height          =   390
               Index           =   5
               Left            =   5790
               Top             =   75
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   688
               Caption         =   "金"
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
            Begin Cis3D_v60.CIS3D UC_Week 
               Height          =   390
               Index           =   6
               Left            =   6720
               Top             =   75
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   688
               BackColor       =   16744576
               Caption         =   "土"
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
            Begin Cis3D_v60.CIS3D UC_Week 
               Height          =   390
               Index           =   0
               Left            =   1140
               Top             =   75
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   688
               BackColor       =   255
               Caption         =   "日"
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
            Begin Cis3D_v60.CIS3D UC_Week 
               Height          =   390
               Index           =   1
               Left            =   2070
               Top             =   75
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   688
               Caption         =   "月"
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
            Begin vsFlexLib.vsFlexArray vsGrid1 
               Height          =   8355
               Left            =   60
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   105
               Width           =   7665
               _Version        =   196608
               _ExtentX        =   13520
               _ExtentY        =   14737
               _StockProps     =   228
               Appearance      =   1
               ConvInfo        =   1418783674
               FormatString    =   "日付    |XXXXXXX|XXXXXXX|XXXXXXX|XXXXXXX|XXXXXXX|XXXXXXX|XXXXXXX|"
               Rows            =   25
               Cols            =   11
               BackColor       =   12632256
               BackColorFixed  =   12632256
               BackColorBkg    =   12632256
               TextStyle       =   4
               TextStyleFixed  =   3
               FocusRect       =   2
               HighLight       =   0
               ScrollBars      =   0
               RowHeightMin    =   330
               FillStyle       =   1
               GridColor       =   14737632
               SelectionMode   =   2
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   1
                  Left            =   1050
                  TabIndex        =   21
                  Top             =   975
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   2
                  Left            =   1980
                  TabIndex        =   22
                  Top             =   975
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   3
                  Left            =   2910
                  TabIndex        =   23
                  Top             =   975
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   4
                  Left            =   3840
                  TabIndex        =   24
                  Top             =   975
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   5
                  Left            =   4770
                  TabIndex        =   25
                  Top             =   975
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   6
                  Left            =   5700
                  TabIndex        =   26
                  Top             =   975
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   7
                  Left            =   6630
                  TabIndex        =   27
                  Top             =   975
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   8
                  Left            =   1050
                  TabIndex        =   28
                  Top             =   2295
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   9
                  Left            =   1980
                  TabIndex        =   29
                  Top             =   2295
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   10
                  Left            =   2910
                  TabIndex        =   30
                  Top             =   2295
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   11
                  Left            =   3840
                  TabIndex        =   31
                  Top             =   2295
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   12
                  Left            =   4770
                  TabIndex        =   32
                  Top             =   2295
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   13
                  Left            =   5700
                  TabIndex        =   33
                  Top             =   2295
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   14
                  Left            =   6630
                  TabIndex        =   34
                  Top             =   2295
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   15
                  Left            =   1065
                  TabIndex        =   35
                  Top             =   3615
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   16
                  Left            =   1995
                  TabIndex        =   36
                  Top             =   3615
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   17
                  Left            =   2925
                  TabIndex        =   37
                  Top             =   3615
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   18
                  Left            =   3855
                  TabIndex        =   38
                  Top             =   3615
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   19
                  Left            =   4785
                  TabIndex        =   39
                  Top             =   3615
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   20
                  Left            =   5715
                  TabIndex        =   40
                  Top             =   3615
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   21
                  Left            =   6645
                  TabIndex        =   41
                  Top             =   3615
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   22
                  Left            =   1050
                  TabIndex        =   42
                  Top             =   4935
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   23
                  Left            =   1980
                  TabIndex        =   43
                  Top             =   4935
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   24
                  Left            =   2910
                  TabIndex        =   44
                  Top             =   4935
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   25
                  Left            =   3840
                  TabIndex        =   45
                  Top             =   4935
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   26
                  Left            =   4770
                  TabIndex        =   46
                  Top             =   4935
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   27
                  Left            =   5700
                  TabIndex        =   47
                  Top             =   4935
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   28
                  Left            =   6630
                  TabIndex        =   48
                  Top             =   4935
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   29
                  Left            =   1065
                  TabIndex        =   49
                  Top             =   6255
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   30
                  Left            =   1995
                  TabIndex        =   50
                  Top             =   6255
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   31
                  Left            =   2925
                  TabIndex        =   51
                  Top             =   6255
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   32
                  Left            =   3855
                  TabIndex        =   52
                  Top             =   6255
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   33
                  Left            =   4785
                  TabIndex        =   53
                  Top             =   6255
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   34
                  Left            =   5715
                  TabIndex        =   54
                  Top             =   6255
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   35
                  Left            =   6645
                  TabIndex        =   55
                  Top             =   6255
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   36
                  Left            =   1065
                  TabIndex        =   56
                  Top             =   7575
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   37
                  Left            =   1995
                  TabIndex        =   57
                  Top             =   7575
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   38
                  Left            =   2925
                  TabIndex        =   58
                  Top             =   7605
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   39
                  Left            =   3855
                  TabIndex        =   59
                  Top             =   7605
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   40
                  Left            =   4785
                  TabIndex        =   60
                  Top             =   7605
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   41
                  Left            =   5715
                  TabIndex        =   61
                  Top             =   7605
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
               Begin CisText_V60.CisText iB1_SKSu 
                  Height          =   315
                  Index           =   42
                  Left            =   6645
                  TabIndex        =   62
                  Top             =   7605
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  cFormat         =   "#######"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#######"
                  MaxLength       =   7
                  cDataType       =   4
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   6
               End
            End
         End
      End
   End
   Begin VB.Label ColorKei 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFC0&
      Caption         =   "出庫計"
      Height          =   330
      Left            =   9075
      TabIndex        =   19
      Top             =   1485
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label ColorInput 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0E0FF&
      Caption         =   "入力行"
      Height          =   315
      Left            =   6795
      TabIndex        =   14
      Top             =   1485
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label ColorNyuka 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0FFC0&
      Caption         =   "入荷行"
      Height          =   300
      Left            =   5580
      TabIndex        =   13
      Top             =   1485
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label ColorDay 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00E0E0E0&
      Caption         =   "対象外色"
      Height          =   300
      Left            =   600
      TabIndex        =   12
      Top             =   1500
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label ColorSunday 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C0FF&
      Caption         =   "日曜色"
      Height          =   315
      Left            =   1815
      TabIndex        =   11
      Top             =   1500
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label ColorSaturday 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FF8080&
      Caption         =   "土曜色"
      Height          =   330
      Left            =   2970
      TabIndex        =   10
      Top             =   1500
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label ColorToday 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFC0C0&
      Caption         =   "当日色"
      Height          =   330
      Left            =   4140
      TabIndex        =   9
      Top             =   1500
      Visible         =   0   'False
      Width           =   1050
   End
End
Attribute VB_Name = "CXH0520"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   出庫入力
'**       フォームID    :   CXH0520
'**       処理概要      :
'**
'**       作  成  日    :   2004/12/24  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid1      As New CisVsGrid3
    Dim CisVsGrid2      As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    
    Dim mWeekStart      As Integer
    
    Dim mGotColor       As Long
    Dim mLostColor      As Long

'    Dim mSaturdayColor  As Long
'    Dim mSundayColor    As Long
'    Dim mWeeklyColor    As Long
    Dim mSaveColor      As Long
    Dim mDayRow(31)     As Integer
    Dim mDayCol(31)     As Integer
    Dim mDay(31)        As Integer
    
    Dim Ttl_Kensu1      As Long
    Dim Ttl_Kensu2      As Long
    Dim wSuryo          As Double
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
    If RTrim(Rv_Call2) = "" Then If Not IniGet Then End
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
    
    ' 曜日開始を求める
    gSL_Select = "select 数字1 From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '曜日開始管理'"
    gSL_Select = gSL_Select & "   And 値 = '1'"
    If Not SYMRead(gSL_Select, 1) Then
       SYM.数字1 = 0
    End If
    
    mWeekStart = SYM.数字1
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = No
        .FixedHeight = 310
        .PatanMax = 1
        .SelectCol = 12
        .InitGet = False
        .Init
    End With
    Call GridInit1
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid2
        Set .GridObj = vsGrid2
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 310
        .PatanMax = 1
        .SelectCol = 5
        .InitGet = False
        .Init
    End With
    Call GridInit2
    
    mGotColor = OpH1_SKbn(0).BackColor
    mLostColor = OpH1_SKbn(1).BackColor
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = "ADD"
    Call SyoriName(SyoriKB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    If RTrim(Rv_Call2) = "" Then Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'初期表示
    iH1_Ym = Format(Now(), "yyyymm")
    SelectDay.Tag = Format(Now(), "dd")
    ColorToday.Tag = Format(Now(), "dd")
    Call DaySet
    Call OPH1_SKbn_Click(0)
    
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
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    '【新規】
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
                                 'Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                 'Dummy.Enabled = False
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
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
   Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   Call CisFrm.MousePT(11)
   If ProcHB = "H1" Then
        Call HeadBodyClear("H")
        GoTo PB_CAN_Ed
   End If
   If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        Call SyoriName(SyoriKB)
        iH1_Hinbn.SetFocus
        GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    If SyoriKB <> "ADD" Then
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
        MSG_SKbn.Visible = True
    End If
    iH1_Hinbn.SetFocus
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
        OpH1_SKbn(0).SetFocus
        MSG_SKbn.Visible = False
    End If
    iH1_Hinbn.SetFocus
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        SyoriKB = "DEL"
        Call OPH1_SKbn_Click(0)
        MSG_SKbn.Visible = False
    Else
        SyoriKB = ""
    End If
    Call SyoriName(SyoriKB)
    iH1_Hinbn.SetFocus
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = "Select SUBSTRING(実績日,1,6) 実績日,品番 From 実績管理テーブル"
    gSL_Select = gSL_Select & " WHERE 実績区分 = 2 "
    gSL_Select = gSL_Select & "   AND 納入番号 = '' "
    gSL_Select = gSL_Select & "   AND SUBSTRING(実績日,1,6) + 品番 < '" & RTrim(iH1_Ym) & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " GROUP BY SUBSTRING(実績日,1,6),品番,SUBSTRING(実績日,1,6) + 品番  "
    gSL_Select = gSL_Select & " ORDER BY SUBSTRING(実績日,1,6) + 品番 Desc "
    If JKTRead(gSL_Select, 1) Then
        iH1_Ym = Left(JKT.実績日, 6)
        iH1_Hinbn = RTrim(JKT.品番)
    Else
        With CisFun
            .MB_Lines = 4
            .MB_Title = ""
            .MB_Button = Error
            .MB_MSG(2) = "    先頭レコードです！           "
            .MBOX
            iH1_Hinbn.SetFocus
            Exit Sub
        End With
    
    End If
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Ym.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = "Select SUBSTRING(実績日,1,6) 実績日,品番 From 実績管理テーブル"
    gSL_Select = gSL_Select & " WHERE 実績区分 = 2 "
    gSL_Select = gSL_Select & "   AND 納入番号 = '' "
    gSL_Select = gSL_Select & "   AND SUBSTRING(実績日,1,6) + 品番 > '" & RTrim(iH1_Ym) & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " GROUP BY SUBSTRING(実績日,1,6),品番,SUBSTRING(実績日,1,6)+品番  "
    gSL_Select = gSL_Select & " ORDER BY SUBSTRING(実績日,1,6) + 品番  "
    If JKTRead(gSL_Select, 1) Then
        iH1_Ym = Left(JKT.実績日, 6)
        iH1_Hinbn = RTrim(JKT.品番)
    Else
        With CisFun
            .MB_Lines = 4
            .MB_Title = ""
            .MB_Button = Error
            .MB_MSG(2) = "    最終レコードです！           "
            .MBOX
            iH1_Hinbn.SetFocus
            Exit Sub
        End With
    End If
    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Ym.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
End Sub
'+----------------------------+
'+       品番検索            +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Left = 0: RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 1
    RVI_SK = 9
    RVI_Tehai = "2"
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       iH1_Hinbn = RTrim(RV_Hinbn)
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
'+  OPH1_SKbn
'+-----------------------------+
'Private Sub OPH1_SKbn_GotFocus(Index As Integer)
'    OpH1_SKbn(Index).BackColor = mGotColor
'    OpH1_SKbn(0).Tag = Format(Index, "0")
'End Sub
Private Sub OPH1_SKbn_Click(Index As Integer)
    For gInt = 0 To 2
        OpH1_SKbn(gInt).BackColor = mLostColor
    Next
    OpH1_SKbn(Index).BackColor = mGotColor
    OpH1_SKbn(0).Tag = Format(Index, "0")
End Sub
'+-----------------------------+
'+  iB1_SKSu
'+-----------------------------+
Private Sub iB1_SKSu_LostFocus(Index As Integer)
    If ProcHB = "B1" Then
        B1lb_TSuryo(2) = "": B1lb_SSuryo(2) = ""
        For gInt = 1 To 31
            B1lb_TSuryo(2) = Format(CisFun.Val2(B1lb_TSuryo(2)) + CisFun.Val2(iB1_SKSu(mDay(gInt))), "#,###")
            If gInt <= CisFun.Val2(SelectDay.Tag) Then
                B1lb_SSuryo(2) = Format(CisFun.Val2(B1lb_SSuryo(2)) + CisFun.Val2(iB1_SKSu(mDay(gInt))), "#,###")
            End If
        Next gInt
    End If
End Sub
Private Sub iB1_SKSu_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown Then Exit Sub
    If KeyCode = vbKeyUp Then
        If Index - 7 > 0 Then
            If iB1_SKSu(Index - 7).Visible And iB1_SKSu(Index - 7).Enabled Then
                iB1_SKSu(Index - 7).SetFocus
            End If
        End If
    End If
    If KeyCode = vbKeyDown Then
        If Index + 7 <= 37 Then
            If iB1_SKSu(Index + 7).Visible And iB1_SKSu(Index + 7).Enabled Then
                iB1_SKSu(Index + 7).SetFocus
            End If
        End If
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
   Select Case SYKB
     Case "ADD"
          SyoriNM = "追加"
          K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
          PB_ADD.Visible = False
     Case "MNT"
          SyoriNM = "修正"
          K_Sykbnm.ForeColor = &HFF0000  '(青色)
     Case "DEL"
          SyoriNM = "削除"
          K_Sykbnm.ForeColor = &HFF&     '(赤色)
          w_Color = &HFF&
     Case "LOK"
          SyoriNM = "照会"
          K_Sykbnm.ForeColor = &HFF00FF   '(ピンク)
          PB_ADD.Visible = False
          PB_MNT.Visible = False
          PB_DEL.Visible = False
     Case Else
          SyoriNM = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
       If gObj.Name Like "*vsGrid*" Then gObj.ForeColor = w_Color
'      * PF-Key 表示設定 *
       If SyoriKB <> "LOC" Then
          If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
             gObj.Name = "PB_DEL" Then
             gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
          End If
       End If
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
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    
    PB_ADD.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_DEL.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    
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
    If HeadBody = "B1" Or HeadBody = "B" Then
        With vsGrid1
            For gInt = 1 To .Rows - 1
                If gInt Mod 4 <> 1 Then
                    For gCnt1 = 1 To .Cols - 1
                        vsGrid1.TextMatrix(gInt, gCnt1) = ""
                    Next gCnt1
                End If
            Next gInt
        End With
        For gInt = 1 To 37
            iB1_SKSu(gInt).Enabled = True
            iB1_SKSu(gInt).Tag = ""
        Next gInt
        Call CisVsGrid2.Clear
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk("") Then GoTo ReturnPress_Ed
' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        Else
            ProcHB = "B1"
            Call DispChange(ProcHB)
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        If RTrim(Rv_Call2) <> "" Then
            Rv_CallRtn = True
            Call PB_END_Click
            GoTo ReturnPress_Ed
        End If
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
Private Function Head1Chk(ChkKB As String) As Boolean
    Head1Chk = False
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'出庫年月
    If RTrim(iH1_Ym) = "" Then
        CisFun.MB_MSG(1) = "     　【　出庫年月 必須！ 】     "
        CisFun.MB_MSG(3) = "     出庫年月 を入力して下さい。      "
        CisFun.MBOX
        iH1_Ym.SetFocus
        Exit Function
    End If
    If Not iH1_Ym.cChkResult Then
        CisFun.MB_MSG(1) = "     　【　出庫年月 異常！ 】     "
        CisFun.MB_MSG(3) = "     出庫年月 を再入力して下さい。      "
        CisFun.MBOX
        iH1_Ym.SetFocus
        Exit Function
    End If
    gInt = Kdou_Check(iH1_Ym & "01")
    If gInt = 2 Then
    'カレンダ未登録
        CisFun.MB_MSG(1) = "     " & RTrim(iH1_Ym.cYear) & "年" & RTrim(iH1_Ym.cMonth) & "月のカレンダ未登録！     "
        CisFun.MB_MSG(3) = "     　　(　カレンダマスタ　)　"
        CisFun.MBOX
        iH1_Ym.SetFocus
        Exit Function
    End If
'品番
    If Trim(iH1_Hinbn) = "" Then
        CisFun.MB_MSG(2) = "    品番を入力して下さい            "
        CisFun.MBOX
        iH1_Hinbn.SetFocus
        Exit Function
    End If
    HNM.品番 = RTrim(iH1_Hinbn)
    If Not HNMRead("", 1) Then
        CisFun.MB_MSG(2) = "    品番　未登録！（品番マスタ）            "
        CisFun.MBOX
        iH1_Hinbn.SetFocus
        Exit Function
    End If
    
' マスタ入力 SELECT文 ｾｯﾄ
    If Not DBInput("INV") Then
        If SyoriKB <> "ADD" Then
            CisFun.MB_MSG(2) = "    出庫データ　未登録！           "
            CisFun.MBOX
            iH1_Hinbn.SetFocus
            Exit Function
        End If
    Else
'        If SyoriKB = "ADD" Then
'            CisFun.MB_MSG(2) = "    出庫データ　登録済！           "
'            CisFun.MBOX
'            iH1_Hinbn.SetFocus
'            Exit Function
'        End If
    End If
    If SyoriKB = "ADD" Then
        gCnt1 = 0
        For gInt = 1 To 31
            If mDay(gInt) <> 0 Then
                If CisFun.Val2(iB1_SKSu(mDay(gInt))) <> 0 Then
                    iB1_SKSu(mDay(gInt)).Enabled = False
                Else
                    gCnt1 = gCnt1 + 1
                End If
            End If
        Next gInt
        If gCnt1 = 0 Then
            CisFun.MB_MSG(2) = "    出庫データ　登録済！           "
            CisFun.MBOX
            iH1_Hinbn.SetFocus
            Exit Function
        End If
        If OpH1_SKbn(0).Tag = "1" And Ttl_Kensu2 = 0 Then
            CisFun.MB_MSG(2) = "    構成が存在しません！           "
            CisFun.MBOX
            iH1_Hinbn.SetFocus
            Exit Function
        End If
    End If
    
Head1Chk_Ed:
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
'ﾒｯｾｰｼﾞﾎﾞｯｸｽ設定
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'
'
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
'カレンダー再表示
    If RTrim(iH1_Ym) <> RTrim(iH1_Ym.Tag) Then
        Call DaySet
    End If
'出庫表示
    Call Grid1Set
'構成情報表示
    Call Grid2Set

    If Ttl_Kensu1 = 0 Then Exit Function
DBInput_Ed:
    DBInput = True
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+      削 除 処 理      +
'+----------------------+
Private Function DBDelete()
    DBDelete = False
'  ( 確認メッセージ )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 7
        .MB_MSG(1) = "     出庫年月 = " & iH1_Ym.cYear & "年" & iH1_Ym.cMonth & "月"
        .MB_MSG(3) = "     品　番 　　= " & RTrim(iH1_Hinbn)
        .MB_MSG(5) = "     の出庫データを削除します。よろしいですか？    　　　  "
        .MB_Title = "削除処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
    ' 実績管理テーブル
        .SQL = "DELETE FROM 実績管理テーブル "
        .SQL = .SQL & " WHERE 実績区分 = 2"
        .SQL = .SQL & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
        .SQL = .SQL & "   AND 実績日 LIKE '" & RTrim(iH1_Ym) & "%'"
        .SQL = .SQL & "   AND 納入番号 = '' "
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
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Dim AddFlg  As Boolean
    Dim KanriNo  As Long

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "     　更新処理を行います。　　      "
        .MB_Title = "更新処理"
        .MB_MSG(3) = "     　よろしいですか？      　　"
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
    If OpH1_SKbn(0).Tag = "1" Then GoTo DBPut_Kosei
'-------------------------------------'
'   親品番作成（単品、全部）
'-------------------------------------'
    For gInt = 1 To 31
        If mDay(gInt) <> 0 Then
            If CisFun.Val2(iB1_SKSu(mDay(gInt))) <> CisFun.Val2(iB1_SKSu(mDay(gInt)).Tag) Then
                gSL_Select = "SELECT 実績管理NO FROM 実績管理テーブル"
                gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
                gSL_Select = gSL_Select & "   AND 実績日 = '" & RTrim(iH1_Ym) & Format(gInt, "00") & "'"
                gSL_Select = gSL_Select & "   AND 実績区分 = 2"
                gSL_Select = gSL_Select & "   AND 納入番号 = ''"
                If JKTRead(gSL_Select, 1, 0) Then
                    KanriNo = JKT.実績管理NO
                    AddFlg = False
                Else
                    KanriNo = 0
                    AddFlg = True
                End If
                If CisFun.Val2(iB1_SKSu(mDay(gInt))) = 0 Then
                    With CisDB
                ' ( 削除処理 )
                        .SQL = "DELETE FROM 実績管理テーブル "
                        .SQL = .SQL & " WHERE 実績管理NO = " & KanriNo & ""
                        .DBExec
                
                        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + CisDB.RecordCount, "###")
                    End With
                Else
                    If KanriNo = 0 Then
                        With CisDB
                            .SQL = "番号獲得"
                            .StoadoCount = 3
                            .ParaNo = 0: .ParaIO = Return用
                            .ParaNo = 1: .ParaIO = Input用
                            .ParaNo = 2: .ParaIO = OutPut用
                            .ParaNo = 3: .ParaIO = Return用
                            .ParaNo = 1: .ParaValue = 6
                            If Not .DBStored Then
                                With CisFun
                                    .MB_Lines = 4
                                    .MB_MSG(2) = "     実績管理NOの採番に失敗しました    "
                                    .MB_Title = "採番処理"
                                    .MB_Button = Error
                                    If Not .MBOX Then GoTo DBPut_Err
                                End With
                            Else
                                .ParaNo = 2: KanriNo = .ParaValue
                            End If
                        End With
                    
                    End If
                    Call ItemsClearJKT
                    With JKT
                        .実績管理NO = KanriNo
                        .納品書行 = 1
                        .実績区分 = 2
                        .品番 = RTrim(iH1_Hinbn)
                        .実績数 = CisFun.Val2(iB1_SKSu(mDay(gInt)))
                        .実績日 = RTrim(iH1_Ym) & Format(gInt, "00")
                        .作成者 = gTanto
                    End With
        '
                ' ( 追加処理 )
                    If AddFlg Then
                        Call JKTInsert
                        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
                    Else
                ' ( 更新処理 )
                        gSL_Select = "UPDATE 実績管理テーブル SET"
                        gSL_Select = gSL_Select & " 実績数 = " & JKT.実績数 & ","
                        gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
                        gSL_Select = gSL_Select & " 更新端末 = SubString(Host_Name(),1,20),"
                        gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "'"
                        gSL_Select = gSL_Select & " WHERE 実績管理NO = " & JKT.実績管理NO & ""
                        With CisDB
                            .SQL = gSL_Select
                            .DBExec
                        End With
                        
                        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
                    End If
                End If
            End If
        End If
    Next gInt
    If OpH1_SKbn(0).Tag = "0" Then GoTo DBPut_Ed
DBPut_Kosei:
    If SyoriKB <> "ADD" Then GoTo DBPut_Ed
'-------------------------------------'
'   構成品番出庫作成（構成、全部）
'-------------------------------------'
    For gCnt1 = 1 To Ttl_Kensu2
        If RTrim(vsGrid2.TextMatrix(gCnt1, CisVsGrid2.SelectCol)) <> "" Then
            For gInt = 1 To 31
                If mDay(gInt) <> 0 Then
                    If CisFun.Val2(iB1_SKSu(mDay(gInt))) <> CisFun.Val2(iB1_SKSu(mDay(gInt)).Tag) Then
                        gSL_Select = "SELECT 実績管理NO FROM 実績管理テーブル"
                        gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(vsGrid2.TextMatrix(gCnt1, 4)) & "'"
                        gSL_Select = gSL_Select & "   AND 実績日 = '" & RTrim(iH1_Ym) & Format(gInt, "00") & "'"
                        gSL_Select = gSL_Select & "   AND 実績区分 = 2"
                        gSL_Select = gSL_Select & "   AND 納入番号 = ''"
                        If JKTRead(gSL_Select, 1, 0) Then
                            KanriNo = JKT.実績管理NO
                            AddFlg = False
                        Else
                            KanriNo = 0
                            AddFlg = True
                        End If
                        If CisFun.Val2(iB1_SKSu(mDay(gInt))) = 0 Then
                            With CisDB
                        ' ( 削除処理 )
                                .SQL = "DELETE FROM 実績管理テーブル "
                                .SQL = .SQL & " WHERE 実績管理NO = " & KanriNo & ""
                                .DBExec
                        
                                iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + CisDB.RecordCount, "###")
                            End With
                        Else
                            If KanriNo = 0 Then
                                With CisDB
                                    .SQL = "番号獲得"
                                    .StoadoCount = 3
                                    .ParaNo = 0: .ParaIO = Return用
                                    .ParaNo = 1: .ParaIO = Input用
                                    .ParaNo = 2: .ParaIO = OutPut用
                                    .ParaNo = 3: .ParaIO = Return用
                                    .ParaNo = 1: .ParaValue = 6
                                    If Not .DBStored Then
                                        With CisFun
                                            .MB_Lines = 4
                                            .MB_MSG(2) = "     実績管理NOの採番に失敗しました    "
                                            .MB_Title = "採番処理"
                                            .MB_Button = Error
                                            If Not .MBOX Then GoTo DBPut_Err
                                        End With
                                    Else
                                        .ParaNo = 2: KanriNo = .ParaValue
                                    End If
                                End With
                            
                            End If
                            Call ItemsClearJKT
                            With JKT
                                .実績管理NO = KanriNo
                                .納品書行 = 1
                                .実績区分 = 2
                                .品番 = RTrim(vsGrid2.TextMatrix(gCnt1, 4))
                                wSuryo = CisFun.Val2(iB1_SKSu(mDay(gInt))) * CisFun.Val2(vsGrid2.TextMatrix(gCnt1, 2))
                                .実績数 = CisFun.HasuProc(wSuryo, 切上げ)
                                .実績日 = RTrim(iH1_Ym) & Format(gInt, "00")
                                .作成者 = gTanto
                            End With
                '
                        ' ( 追加処理 )
                            If AddFlg Then
                                Call JKTInsert
                                iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
                            Else
                        ' ( 更新処理 )
                                gSL_Select = "UPDATE 実績管理テーブル SET"
                                gSL_Select = gSL_Select & " 実績数 = " & JKT.実績数 & ","
                                gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
                                gSL_Select = gSL_Select & " 更新端末 = SubString(Host_Name(),1,20),"
                                gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "'"
                                gSL_Select = gSL_Select & " WHERE 実績管理NO = " & JKT.実績管理NO & ""
                                With CisDB
                                    .SQL = gSL_Select
                                    .DBExec
                                End With
                                
                                iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
                            End If
                        End If
                    End If
                End If
            Next gInt
        End If
    Next gCnt1
DBPut_Ed:
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 中中, あり, 8, 0, "日付")
        Call .FixedSet(中中, 右中, あり, 7, 1, "1")
        Call .FixedSet(中中, 右中, あり, 7, 2, "2")
        Call .FixedSet(中中, 右中, あり, 7, 3, "3")
        Call .FixedSet(中中, 右中, あり, 7, 4, "4")
        Call .FixedSet(中中, 右中, あり, 7, 5, "5")
        Call .FixedSet(中中, 右中, あり, 7, 6, "6")
        Call .FixedSet(中中, 右中, あり, 7, 7, "7")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7")
        .InitDisp
    End With
    
    Dim WColor  As Long
    For gInt = 0 To 6
        WColor = DayColor.BackColor
        Select Case gInt
            Case 0: gStr = "日"
                    WColor = ColorSunday.BackColor
            Case 1: gStr = "月"
            Case 2: gStr = "火"
            Case 3: gStr = "水"
            Case 4: gStr = "木"
            Case 5: gStr = "金"
            Case 6: gStr = "土"
                    WColor = ColorSaturday.BackColor
        End Select
        If gInt - mWeekStart < 0 Then
            UC_Week(6) = gStr
            UC_Week(6).BackColor = WColor
        Else
            UC_Week(gInt - mWeekStart) = gStr
            UC_Week(gInt - mWeekStart).BackColor = WColor
        End If
    Next
    
End Sub
'+==========================================================+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit2()
    With CisVsGrid2
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 中中, あり, 3, 0, "№")
        Call .FixedSet(中中, 左中, あり, 25, 1, "品　番")
        Call .FixedSet(中中, 右中, あり, 10, 2, "構成数")
        Call .FixedSet(中中, 右中, あり, 6, 3, "出庫数")
        
        Call .FixedSet(中中, 左中, あり, 0, 4, "X品番")
        Call .FixedSet(中中, 左中, あり, 0, 5, "SELECT")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3")
        .InitDisp
    End With
    
End Sub
Private Sub vsGrid1_Click()
    If ProcHB <> "B1" Then Exit Sub
    Select Case vsGrid1.MouseRow
        Case 1 To 4: gInt = 1
        Case 5 To 8: gInt = 5
        Case 9 To 12: gInt = 9
        Case 13 To 16: gInt = 13
        Case 17 To 20: gInt = 17
        Case 21 To 23: gInt = 21
    End Select
    If RTrim(vsGrid1.TextMatrix(gInt, vsGrid1.MouseCol)) = "" Then Exit Sub
    SelectDay.Tag = Format(CisFun.Val2(vsGrid1.TextMatrix(gInt, vsGrid1.MouseCol)), "00")
    Call SelectDay_Gokei
    
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub Grid1Set()
    Dim JDay    As Integer
        
    vsGrid1.Redraw = False
    Ttl_Kensu1 = 0
'出庫数セット
    gSL_Select = " SELECT JK.実績日,JK.実績数,JK.納入番号 "
    gSL_Select = gSL_Select & "  FROM 実績管理テーブル JK"
    gSL_Select = gSL_Select & " WHERE JK.品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND JK.実績区分 = 2"
    gSL_Select = gSL_Select & "   AND JK.実績日 LIKE '" & RTrim(iH1_Ym) & "%'"
    If JKTRead(gSL_Select, 0, 1) Then
        With vsGrid1
           .Redraw = False
            Do Until Not JKT_RDSTS
                JDay = CisFun.Val2(Mid(JKT.実績日, 7, 2))
                If RTrim(JKT.納入番号) = "" Then
                    If OpH1_SKbn(0).Tag <> "1" Then
                        iB1_SKSu(mDay(JDay)) = Format(JKT.実績数, "####")
                        iB1_SKSu(mDay(JDay)).Tag = Format(JKT.実績数, "####")
                        Ttl_Kensu1 = Ttl_Kensu1 + 1
                        B1lb_TSuryo(2) = Format(CisFun.Val2(B1lb_TSuryo(2)) + JKT.実績数, "#,###")
                    End If
                End If
                B1lb_TSuryo(3) = Format(CisFun.Val2(B1lb_TSuryo(3)) + JKT.実績数, "#,###")
                .TextMatrix(mDayRow(JDay) + 3, mDayCol(JDay)) = Format(CisFun.Val2(.TextMatrix(mDayRow(JDay) + 3, mDayCol(JDay))) + JKT.実績数, "#,###")
                
                Call JKTReadNext(1)
            Loop
        End With
    End If
    Call JKTClose(1)
'入荷数セット
    gSL_Select = " SELECT NY.入荷日,NY.入荷数 "
    gSL_Select = gSL_Select & "  FROM 入荷テーブル NY"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 発注テーブル HC"
    gSL_Select = gSL_Select & "   ON HC.発注管理NO = NY.発注管理NO"
    gSL_Select = gSL_Select & " WHERE HC.品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND NY.入荷日 LIKE '" & RTrim(iH1_Ym) & "%'"
    If NYTRead(gSL_Select, 0, 1) Then
        With vsGrid1
           .Redraw = False
            Do Until Not NYT_RDSTS
                JDay = CisFun.Val2(Mid(NYT.入荷日, 7, 2))
                .TextMatrix(mDayRow(JDay) + 1, mDayCol(JDay)) = Format(CisFun.Val2(.TextMatrix(mDayRow(JDay) + 1, mDayCol(JDay))) + NYT.入荷数, "####")
                B1lb_TSuryo(1) = Format(CisFun.Val2(B1lb_TSuryo(1)) + NYT.入荷数, "#,###")
                
                Call NYTReadNext(1)
            Loop
        End With
    End If
    Call NYTClose(1)
    
    vsGrid1.Redraw = True
    Call SelectDay_Gokei

End Sub
Private Sub SelectDay_Gokei()

    For gInt = 1 To 3
        B1lb_SSuryo(gInt) = ""
    Next
'選択日付の色変更
    With vsGrid1
        .Redraw = False
        .Row = mDayRow(CisFun.Val2(ColorToday.Tag))
        .Col = mDayCol(CisFun.Val2(ColorToday.Tag))
        .CellBackColor = mSaveColor
        .CellFontBold = False

        SelectDay = SelectDay.Tag & "日迄合計"
        .Row = mDayRow(CisFun.Val2(SelectDay.Tag))
        .Col = mDayCol(CisFun.Val2(SelectDay.Tag))
        mSaveColor = .CellBackColor
        .CellBackColor = ColorToday.BackColor
        .CellFontBold = True

        ColorToday.Tag = SelectDay.Tag
        .Redraw = True
    End With
'指示数セット
    B1lb_TSuryo(2) = ""
    For gInt = 1 To 31
        B1lb_TSuryo(2) = Format(CisFun.Val2(B1lb_TSuryo(2)) + CisFun.Val2(iB1_SKSu(mDay(gInt))), "#,###")
        If gInt <= CisFun.Val2(SelectDay.Tag) Then
            B1lb_SSuryo(1) = Format(CisFun.Val2(B1lb_SSuryo(1)) + CisFun.Val2(vsGrid1.TextMatrix(mDayRow(gInt) + 1, mDayCol(gInt))), "#,###")
            B1lb_SSuryo(2) = Format(CisFun.Val2(B1lb_SSuryo(2)) + CisFun.Val2(iB1_SKSu(mDay(gInt))), "#,###")
            B1lb_SSuryo(3) = Format(CisFun.Val2(B1lb_SSuryo(3)) + CisFun.Val2(vsGrid1.TextMatrix(mDayRow(gInt) + 3, mDayCol(gInt))), "#,###")
        End If
    Next gInt

End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub Grid2Set()
   Dim ID   As Integer

    Ttl_Kensu2 = 0
    gSL_Select = " SELECT KS.子品番,KS.構成数,KS.展開区分,ISNULL(HN.表示品番,'') 変更前子品番,ISNULL(JK.実績数,0) 件数 "
    gSL_Select = gSL_Select & "  FROM 構成マスタ KS"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "   ON HN.品番 = KS.子品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN (SELECT 品番,SUM(実績数)実績数 FROM 実績管理テーブル"
    gSL_Select = gSL_Select & "                     WHERE 実績区分 = 2"
    gSL_Select = gSL_Select & "                       AND SUBSTRING(実績日,1,6) = '" & RTrim(iH1_Ym) & "'"
    gSL_Select = gSL_Select & "                     GROUP BY 品番) JK "
    gSL_Select = gSL_Select & "   ON JK.品番 = KS.子品番 "
    gSL_Select = gSL_Select & " WHERE 親品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " ORDER BY 子品番 "
    If KOMRead(gSL_Select, 0, 1) Then
' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu2 = CisDB.RecordCount       ' 抽出件数(合計)退避
'  ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid2
       .Redraw = False
        If Ttl_Kensu2 + 1 < CisVsGrid2.DispMax Then
            .Rows = CisVsGrid2.DispMax
        Else
            .Rows = Ttl_Kensu2 + 1
        End If
        ID = 0
'出庫入力済数セット
        Do Until Not KOM_RDSTS
            ID = ID + 1
            .TextMatrix(ID, 0) = Format(ID, "##")
            .TextMatrix(ID, 1) = RTrim(KOM.変更前子品番)
            .TextMatrix(ID, 2) = Format(KOM.構成数, "#,##0.0000")
            .TextMatrix(ID, 3) = Format(KOM.件数, "#,###")
'            Select Case KOM.展開区分
'                Case 1: .TextMatrix(ID, 3) = "対象"
'                Case 2: .TextMatrix(ID, 3) = "対象外"
'                Case 3: .TextMatrix(ID, 3) = "対象外"
'                Case Else: .TextMatrix(ID, 3) = ""
'            End Select
            .TextMatrix(ID, 4) = RTrim(KOM.子品番)
            
            If OpH1_SKbn(0).Tag <> "0" Then
                .Row = ID
                Call CisVsGrid2.vsColor
            End If
            Call KOMReadNext(1)
        Loop
        Call KOMClose(1)
        vsGrid2.Redraw = True
    End With
    End If
End Sub
Private Sub vsGrid2_Click()
    If ProcHB <> "B1" Then Exit Sub
    If OpH1_SKbn(0).Tag = "0" Then Exit Sub
    
    Call CisVsGrid2.vsColor
End Sub
Private Sub DaySet()
    Dim wDayCol         As Integer
    Dim SetDay          As Integer
    Dim EndDay          As Integer
    Dim wKado           As Boolean
    
'*カレンダ部日付 セット  *
    gSL_Select = "Select * From カレンダマスタ"
    gSL_Select = gSL_Select & " Where 取引先区分 = 1"
    gSL_Select = gSL_Select & " And   取引先 = ''"
    gSL_Select = gSL_Select & " And   年 = '" & Format(iH1_Ym.cYear, "0000") & "'"
    gSL_Select = gSL_Select & " And   月 = '" & Format(iH1_Ym.cMonth, "00") & "'"
    gSL_Select = gSL_Select & " And   受入 = ''"
    If Not CLMRead(gSL_Select, 1) Then
        Call ItemsClearCLM
        CLM.稼動区分 = String(31, "0")
        CLM.夜勤区分 = String(31, "0")
        With CisFun
            .Date_RsFormat = fDD
            .Date_End iH1_Ym & "01"
            CLM.末日 = .Val2(.Date_Result)
            CLM.稼動日数 = CLM.末日
        End With
    End If
    
    EndDay = CLM.末日
    gDate = DateValue(iH1_Ym.cYear & "/" & iH1_Ym.cMonth & "/01")
    gInt = Weekday(gDate, mWeekStart + 1)       '開始位置
    If RTrim(iH1_Ym) <> Format(Now(), "yyyymm") Then
        SelectDay.Tag = Format(EndDay, "00")
        ColorToday.Tag = Format(EndDay, "00")
    Else
        SelectDay.Tag = Format(Now(), "dd")
        ColorToday.Tag = Format(Now(), "dd")
    End If

    Erase mDayRow: Erase mDayCol: Erase mDay
    wDayCol = 1: SetDay = 1
    Call CisVsGrid1.Clear
    gCnt3 = 0
    
    With vsGrid1
        .Redraw = False
        For gCnt1 = 1 To .Rows - 1 Step 4
            gCnt3 = gCnt3 + 1
            .Row = gCnt1: .Col = 0
            .TextMatrix(.Row, .Col) = "日 付"
            .Row = gCnt1 + 1: .Col = 0
            .TextMatrix(.Row, .Col) = "入荷数"
            .CellBackColor = ColorNyuka.BackColor
            .Row = gCnt1 + 2: .Col = 0
            .TextMatrix(.Row, .Col) = "出庫数"
            .CellBackColor = ColorInput.BackColor
            .Row = gCnt1 + 3: .Col = 0
            .TextMatrix(.Row, .Col) = "出庫合計"
            .CellBackColor = ColorKei.BackColor

            For gCnt2 = 1 To 7
                .Row = gCnt1
                .Col = gCnt2
                If gCnt2 >= gInt Then
                    If SetDay <= EndDay Then
                        gDate = DateValue(iH1_Ym.cYear & "/" & iH1_Ym.cMonth & "/" & SetDay)
'                        gCnt3 = Weekday(gDate, mWeekStart + 1)
                        .CellBackColor = vsGrid1.BackColor
                        .CellAlignment = flexAlignCenterCenter
                        .TextMatrix(gCnt1, gCnt2) = SetDay
'
                        wKado = False
                        If Mid(CLM.稼動区分, SetDay, 1) = "0" Or _
                           Mid(CLM.夜勤区分, SetDay, 1) = "0" Then
                           wKado = True
                        End If
                        If Not wKado Then
                           .CellBackColor = ColorSunday.BackColor
                        End If
                        If CisFun.Val2(SelectDay.Tag) = SetDay Then
                            mSaveColor = .CellBackColor
                        End If
'
                        mDayRow(SetDay) = gCnt1
                        mDayCol(SetDay) = gCnt2
                        mDay(SetDay) = (gCnt3 - 1) * 7 + gCnt2
'
                        .Row = .Row + 1
                        .CellBackColor = ColorNyuka.BackColor
                        .Row = .Row + 1
                        .CellBackColor = ColorInput.BackColor
                        .Row = .Row + 1
                        .CellBackColor = ColorKei.BackColor
                        iB1_SKSu((gCnt3 - 1) * 7 + gCnt2).Visible = True
                        SetDay = SetDay + 1
                        gInt = 0
                    Else
                        .CellBackColor = ColorDay.BackColor
                        .Row = .Row + 1
                        .CellBackColor = ColorDay.BackColor
                        .Row = .Row + 1
                        .CellBackColor = ColorDay.BackColor
                        .Row = .Row + 1
                        .CellBackColor = ColorDay.BackColor
                        iB1_SKSu((gCnt3 - 1) * 7 + gCnt2).Visible = False
                    End If
                Else
                    .CellBackColor = ColorDay.BackColor
                    .Row = .Row + 1
                    .CellBackColor = ColorDay.BackColor
                    .Row = .Row + 1
                    .CellBackColor = ColorDay.BackColor
                    .Row = .Row + 1
                    .CellBackColor = ColorDay.BackColor
                    iB1_SKSu(gCnt2).Visible = False
                End If
            Next gCnt2
        Next gCnt1
        If mDayRow(SetDay - 1) < 21 Then
            .Row = 21: .Col = 0
            .TextMatrix(.Row, .Col) = ""
            .CellBackColor = ColorDay.BackColor
            .Row = .Row + 1: .Col = 0
            .TextMatrix(.Row, .Col) = ""
            .CellBackColor = ColorDay.BackColor
            .Row = .Row + 1: .Col = 0
            .TextMatrix(.Row, .Col) = ""
            .CellBackColor = ColorDay.BackColor
            .Row = .Row + 1: .Col = 0
            .TextMatrix(.Row, .Col) = ""
            .CellBackColor = ColorDay.BackColor
        End If
    
        .Redraw = True
    End With

    iH1_Ym.Tag = RTrim(iH1_Ym)
End Sub
