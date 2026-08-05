VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form COS0310 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "買掛金明細リスト　発行"
   ClientHeight    =   10980
   ClientLeft      =   -90
   ClientTop       =   765
   ClientWidth     =   15315
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
   ScaleHeight     =   10980
   ScaleWidth      =   15315
   WindowState     =   2  '最大化
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1020
      Top             =   6225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   870
      Left            =   2280
      ScaleHeight     =   810
      ScaleWidth      =   11130
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9390
      Width           =   11190
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   750
         Left            =   15
         Top             =   45
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   1323
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
         Begin VB.PictureBox Csv_Bar 
            Height          =   420
            Left            =   4875
            ScaleHeight     =   360
            ScaleWidth      =   5685
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   180
            Visible         =   0   'False
            Width           =   5745
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "　出　力　先"
            Height          =   600
            Left            =   90
            TabIndex        =   11
            Top             =   45
            Width           =   4755
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "ﾌﾟﾘﾝﾀ"
               Height          =   240
               Index           =   0
               Left            =   225
               TabIndex        =   3
               Top             =   270
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "画面(ﾌﾟﾚﾋﾞｭｰ)"
               Height          =   240
               Index           =   1
               Left            =   1665
               TabIndex        =   4
               Top             =   270
               Width           =   1875
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "CSV"
               Height          =   240
               Index           =   2
               Left            =   3705
               TabIndex        =   5
               Top             =   270
               Width           =   975
            End
         End
         Begin Cis3D_v60.CIS3D lb_PrtDev 
            Height          =   420
            Left            =   4875
            Top             =   180
            Width           =   5745
            _ExtentX        =   10134
            _ExtentY        =   741
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
            Height          =   420
            Left            =   10635
            Top             =   180
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   741
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
            cAlingnment     =   7
            cPositionX      =   30
            cPositionY      =   -90
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      Height          =   1110
      Left            =   2310
      ScaleHeight     =   1050
      ScaleWidth      =   5625
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1005
      Width           =   5685
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   990
         Left            =   1905
         Top             =   30
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   1746
         ForeColor       =   16711680
         Caption         =   "契 約 先"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         cPositionX      =   30
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1110
            Top             =   405
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
            cPositionX      =   40
         End
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   150
            TabIndex        =   1
            Top             =   405
            Width           =   975
            _ExtentX        =   1720
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
            Text            =   "XXXXXXX"
            MaxLength       =   7
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   990
         Left            =   60
         Top             =   30
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1746
         ForeColor       =   16711680
         Caption         =   "印刷年月"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
         cPositionX      =   30
         cPositionY      =   30
         Begin CisYM_V60.CisYM iH1_Ym 
            Height          =   360
            Left            =   195
            TabIndex        =   0
            Top             =   390
            Width           =   1470
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
   End
   Begin Cis3D_v60.CIS3D CIS3D1 
      Height          =   375
      Left            =   9585
      Top             =   0
      Width           =   3825
      _ExtentX        =   6747
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
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   345
         Left            =   1920
         Top             =   15
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_Kensu 
            Height          =   270
            Left            =   975
            Top             =   30
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   476
            BackColor       =   16777152
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
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   345
         Left            =   60
         Top             =   15
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
         Caption         =   "選択件数"
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
         Begin Cis3D_v60.CIS3D B1lb_SKensu 
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
      Left            =   9120
      ScaleHeight     =   660
      ScaleWidth      =   4260
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1020
      Width           =   4320
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1725
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
         Left            =   2565
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
         Left            =   3405
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
      Begin CisBtn_60.CisBtn PB_Prt 
         Height          =   585
         Left            =   885
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
   Begin VB.PictureBox B1_Area1 
      Height          =   7065
      Left            =   2295
      ScaleHeight     =   7005
      ScaleWidth      =   11115
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2280
      Width           =   11175
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   6885
         Left            =   60
         Top             =   60
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   12144
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
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   6690
            Left            =   90
            TabIndex        =   2
            Top             =   90
            Width           =   10800
            _Version        =   196608
            _ExtentX        =   19050
            _ExtentY        =   11800
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
            FormatString    =   "№  |得意先                      |FCD   |識別子|名  称                |件数    |印刷種|||"
            Rows            =   21
            Cols            =   11
            BackColor       =   12648447
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
            BackColorAlternate=   12648447
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
      Left            =   960
      TabIndex        =   6
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
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 買掛金明細リスト　発行  】"
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
      cPositionX      =   1600
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
         Left            =   7065
         Top             =   45
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
   End
End
Attribute VB_Name = "COS0310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   買掛金明細リスト　発行
'**       フォームID    :   COS0310
'**       処理概要      :
'**
'**       作  成  日    :   2004/11/10  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl

    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim CisVsGrid1              As New CisVsGrid3
    Dim Sl_Where        As String       ' ﾌﾟﾘﾝﾀ出力先
'   出力区分の背景色
    Dim mColor_On               As Long
    Dim mColor_Off              As Long
    Dim mPrtDev         As String       ' ﾌﾟﾘﾝﾀ出力先
    Dim mCsvDir         As String
    Dim mCsvKey         As String
    Dim mCsvName        As String
    Dim mSavePrint      As String
    Dim mCsvSelect      As String

    Dim SaveRow         As Integer
    Dim SaveTorcd       As String
    Dim SaveFCD         As String
    Dim SaveSikib       As String
    Dim SavePrtSyu      As Byte
    Dim SaveColName     As String
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
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
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

'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 320
        .PatanMax = 1
        .SelectCol = 9
        Set .SelectDispObj = B1lb_SKensu
        .InitGet = False
        .Init
    End With
    Call GridInit

'   背景色の設定（出力区分）
    mColor_On = Op_Out(0).BackColor
    mColor_Off = Op_Out(1).BackColor

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
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
' 対象年月初期表示値を求める
    gSL_Select = "Select Max(対象年月) 対象年月 From TEXT買掛金テーブル "
    If TKKRead(gSL_Select, 1) Then
        iH1_Ym = TKK.対象年月
    End If

'CSVﾌｧｲﾙ出力先
    mCsvName = "買掛金明細リスト"
    mCsvKey = "COS0310"
    mCsvDir = GetSetting(appname:=gRegKey, section:="CSV", _
                       Key:=mCsvKey, Default:=Mid(CisFun.WinRoot, 1, 3) & mCsvName & ".csv")
''ﾌﾟﾘﾝﾀ出力先
    Op_Out(0).Value = True
    mPrtDev = Printer.DeviceName
    lb_PrtDev = mPrtDev
    mSavePrint = mPrtDev
    
    
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
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click      '【検索】
           Case vbKeyF8:        If PB_Prt.Visible Then Call PB_Prt_Click        '【印刷】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click        '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click        '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                          '【入力】
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
    Call CisDB.DBDISConnect
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
       ProcHB = "H1"
       Call DispChange(ProcHB)
       Call HeadBodyClear("B")
       GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    印　刷  キ ー(F8)       +
'+----------------------------+
Private Sub PB_Prt_Click()
    If ProcHB = "B1" Then
        If CisFun.Val2(B1lb_SKensu) = 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    選択されていません｡            "
                .MB_Title = "抽出処理"
                .MB_Button = Error
                .MBOX
            End With
            Exit Sub
        End If
        
        If Not PrintProc Then Exit Sub
        
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
End Sub
'+----------------------------+
'+      取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RVI_Kubun = "照合対象契約先"
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
       iH1_Torcd = RV_Code          ' 値
       H1lb_Tornm = RV_Kubun        ' 値名称
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
    If MeisyoGet("照合対象契約先", iH1_Torcd) Then
       H1lb_Tornm = SYM.値名称
    End If
End Sub
'+----------------------------------------+
'+  Op_Out   出力先
'+----------------------------------------+
Private Sub Op_Out_GotFocus(Index As Integer)
    Op_Out(0).Tag = Index
    For Each gObj In Op_Out
        If gObj.Index = Index Then
            gObj.BackColor = mColor_On
        Else
            gObj.BackColor = mColor_Off
        End If
    Next gObj
    Select Case Index
        Case 0: lb_PrtDev = mPrtDev
        Case 1: lb_PrtDev = mPrtDev
        Case 2: lb_PrtDev = mCsvDir
        Case Else
            lb_PrtDev = ""
    End Select
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

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1": H1Mode = True: H1Color = gPGotSel
       Case "B1": B1Mode = True: B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode: H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode: B1_Area1.BackColor = B1Color

    PB_Look.Visible = False
    PB_Prt.Visible = Not H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
'    PB_ENT.Visible = H1Mode

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
            If .Name Like "i" & wbc_Enm & "_*" Then .text = "": .Enabled = True
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Cmb_*" Then .ListIndex = 0
        End With
    Next gObj
    If HeadBody = "B" Then
        CisVsGrid1.SelectCount = 0
        Call CisVsGrid1.Clear
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
' 印刷年月
    If Trim(iH1_Ym) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　印刷年月を入力して下さい    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ym.SetFocus
        Exit Function
    End If
' 得意先
    If Trim(iH1_Torcd) <> "" Then
        If Not MeisyoGet("照合対象契約先", iH1_Torcd) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "   　契約先　未登録    "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
    End If
' 対象データ存在確認
    If Not DBInput Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    該当データが存在しません｡            "
            .MB_MSG(3) = "    再度抽出条件を入力してください       "
            .MB_Title = "抽出処理"
            .MB_Button = Error
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

    If RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 1)) = "" Then Exit Function

    Call Shikibetsu_Name(vsGrid1.Row)
    
    Body1Chk = True
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput() As Boolean
    Dim DateStr As String
    Dim DateSt  As String
    Dim DateEd  As String
    Dim BinSt   As String
    Dim BinEd   As String
    
    DBInput = False
    On Error GoTo DBInput_Err:

    With CisFun
         .Where_CreateKB = NewDocument
         ' 区分名称
         .Where_Create AlphaMe, "SY1.区分名称", "TEXT買掛金"
         .Where_Create AlphaMe, "SY1.区分タイプ", "", , Create
         ' 納入先
         .Where_Create AlphaMe, "SY1.英数字1", iH1_Torcd
    End With
    
    gSL_Select = "SELECT SY1.英数字1 契約先,ISNULL(SY2.値名称,'') 契約先名,SY1.値 データ,"
    gSL_Select = gSL_Select & "    SY1.英数字2 FCD,SY1.英数字3 識別子,SY1.値名称 識別子名,ISNULL(件数,0) 件数,"
    gSL_Select = gSL_Select & "    CASE WHEN SY1.数字3 = 0 THEN '複数' ELSE '単一' END 帳票種類"
    gSL_Select = gSL_Select & " From 名称マスタ SY1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "   ON SY2.区分名称 = '照合対象契約先'"
    gSL_Select = gSL_Select & "   AND SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY2.値 = SY1.英数字1"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN (SELECT FCD,識別子,契約先,COUNT(品番)件数 FROM TEXT買掛金テーブル"
    gSL_Select = gSL_Select & "             WHERE 対象年月 = '" & RTrim(iH1_Ym) & "'"
    gSL_Select = gSL_Select & "             GROUP BY FCD,識別子,契約先 ) TK"
    gSL_Select = gSL_Select & "   ON TK.FCD = SY1.英数字2"
    gSL_Select = gSL_Select & "   AND TK.識別子 = SY1.英数字3"
    gSL_Select = gSL_Select & "   AND TK.契約先 = SY1.英数字1"
    
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    
    gSL_Select = gSL_Select & " ORDER BY SY1.英数字1,SY1.値"
    
    
    If Not TKKRead(gSL_Select) Then
       Call TKKClose
       Exit Function
    End If
    
    Call GridSet
    
    DBInput = True

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'+-------------------------------+
'+     プリンタ・CSV出力先 変更    +
'+-------------------------------+
Private Sub PB_Print_Click()
    If Op_Out(2).Value Then
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
    Else
    'If Op_Out(0).Value Or Op_Out(1).Value Then
        ' プリンタ種類変更
         On Error Resume Next
         CommonDialog1.Flags = 64 + cdlPDHidePrintToFile
         CommonDialog1.Max = 0
         CommonDialog1.Min = 0
         CommonDialog1.copies = 1
         CommonDialog1.ToPage = 1
         CommonDialog1.FromPage = 1
        ' 以下のｿｰｽｺｰﾄﾞは、<ｷｬﾝｾﾙ>ﾎﾞﾀﾝがｸﾘｯｸされたときに実行されます。
         CommonDialog1.CancelError = True
         CommonDialog1.ShowPrinter
         If Err = 32755 Then
        '<ｷｬﾝｾﾙ>ﾎﾞﾀﾝが押されたときにｴﾗｰｺｰﾄﾞを返すようにする設定です。
        '必要なｿｰｽｺｰﾄﾞを入力してください。
         Else
             mPrtDev = Printer.DeviceName
             lb_PrtDev.Caption = mPrtDev
         End If
         On Error GoTo 0
    End If
End Sub
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
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 左中, あり, 27, 1, "契 約 先")
        Call .FixedSet(中中, 中中, あり, 6, 2, "コード")
        Call .FixedSet(中中, 中中, あり, 6, 3, "識別子")
        Call .FixedSet(中中, 左中, あり, 26, 4, "名　称")
        Call .FixedSet(中中, 右中, あり, 8, 5, "件数")
        Call .FixedSet(中中, 中中, あり, 6, 6, "帳票種")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 7, "値")
        Call .FixedSet(中中, 左中, なし, 0, 8, "S契約先")
        Call .FixedSet(中中, 左中, なし, 0, 9, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'***********************************************'
'*****      グリッド内容セット - 流動日付
'***********************************************'
Private Function GridSet() As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

    GridSet = False

'       抽出処理 初期ｾｯﾄ
    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
'       ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
            .Rows = CisVsGrid1.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
        ID = 0
        Do Until Not TKK_RDSTS
            ID = ID + 1
                
            .TextMatrix(ID, 1) = TKK.契約先 & ":" & TKK.契約先名
            .TextMatrix(ID, 2) = TKK.FCD
            .TextMatrix(ID, 3) = TKK.識別子
            .TextMatrix(ID, 4) = TKK.識別子名
            .TextMatrix(ID, 5) = Format(TKK.件数, "##,###")
            .TextMatrix(ID, 6) = TKK.帳票種類
            .TextMatrix(ID, 7) = TKK.データ
            .TextMatrix(ID, 8) = TKK.契約先

            Call TKKReadNext
        Loop
        .Row = 1: .Col = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With

    GridSet = True
    
    CisVsGrid1.ReNumber
    
    Call TKKClose
End Function
'******************************************'
'*****      行選択（クリック）
'******************************************'
Private Sub VSGrid1_Click()
    If ProcHB = "B1" Then
        Call CisVsGrid1.vsColor
    End If
End Sub
'******************************************'
'*****      行選択（ＤblClick）
'******************************************'
Private Sub vsGrid1_DblClick()
    If vsGrid1.MouseCol <> CisVsGrid1.FixedGet(4) Then Exit Sub
    If vsGrid1.TextMatrix(vsGrid1.MouseRow, CisVsGrid1.FixedGet(1)) = "" Then Exit Sub
'
    Call Shikibetsu_Name(vsGrid1.MouseRow)
End Sub
'******************************************'
'*****      行選択（スペース）
'******************************************'
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If ProcHB = "B1" Then
        If KeyCode = vbKeySpace Then
            Call CisVsGrid1.vsColor
        End If
    End If
End Sub
'******************************************'
'*****      識別子　名称登録
'******************************************'
Private Sub Shikibetsu_Name(SaveRow As Integer)
'
    With vsGrid1
        RV_Torcd = .TextMatrix(SaveRow, CisVsGrid1.FixedGet(1))
        RV_Code = .TextMatrix(SaveRow, CisVsGrid1.FixedGet(2))
        RV_KOKBN = .TextMatrix(SaveRow, CisVsGrid1.FixedGet(3))
        RV_Name = .TextMatrix(SaveRow, CisVsGrid1.FixedGet(4))
        If RTrim(.TextMatrix(SaveRow, CisVsGrid1.FixedGet(6))) = "複数" Then
            RV_Su1 = 0
        Else
            RV_Su1 = 1
        End If
        RV_Kubun = .TextMatrix(SaveRow, CisVsGrid1.FixedGet(7))
    End With
    
    COS0315.Show vbModal
    Unload COS0315
    Set COS0315 = Nothing

    If RV_Rtn Then
        vsGrid1.TextMatrix(SaveRow, CisVsGrid1.FixedGet(4)) = RV_Name
        vsGrid1.TextMatrix(SaveRow, CisVsGrid1.FixedGet(6)) = RV_NameR
    End If
End Sub
'+---------------------------------------------+
'+      印刷処理                               +
'+---------------------------------------------+
Private Function PrintProc() As Boolean
    PrintProc = False
    Dim KeyTorcd    As String
    Dim TCnt        As Integer
'
    ReDim PrtSelF(1): ReDim PrtSelT(1): ReDim PrtTorcd(1)
    TCnt = 0
    
    S310_Ym = RTrim(iH1_Ym)
    S310_TotalCnt = 0
'-----------------------------'
'   対象SELECT分作成
'-----------------------------'
    With vsGrid1
        For gInt = 1 To .Rows - 1
            If .TextMatrix(gInt, CisVsGrid1.SelectCol) <> "" Then
                S310_Torcd = .TextMatrix(gInt, CisVsGrid1.FixedGet(8))
                S310_FCD = .TextMatrix(gInt, CisVsGrid1.FixedGet(2))
                S310_SCD = .TextMatrix(gInt, CisVsGrid1.FixedGet(3))
                If RTrim(.TextMatrix(gInt, CisVsGrid1.FixedGet(6))) = "複数" Then
                    S310_PrtSyu = 0
                Else
                    S310_PrtSyu = 1
                End If
                
                If Not Op_Out(2).Value Then
                    '--(( 印刷用SELECT分作成 ))--
                    If RTrim(S310_Torcd) <> KeyTorcd Then
                        If RTrim(KeyTorcd) <> "" Then
                            If RTrim(PrtSelF(0)) <> "" Then PrtSelF(TCnt) = Sl_Where & PrtSelF(0) & ")"
                            If RTrim(PrtSelT(0)) <> "" Then PrtSelT(TCnt) = Sl_Where & PrtSelT(0) & ")"
                        End If
                        TCnt = TCnt + 1
                        If TCnt > UBound(PrtSelF) Then
                            ReDim Preserve PrtSelT(TCnt)
                            ReDim Preserve PrtSelF(TCnt)
                            ReDim Preserve PrtTorcd(TCnt)
                        End If
                        PrtSelT(TCnt) = "": PrtSelF(TCnt) = "": PrtTorcd(TCnt) = .TextMatrix(gInt, CisVsGrid1.FixedGet(1))
                        PrtSelF(0) = "": PrtSelT(0) = ""
                        Sl_Where = " WHERE 対象年月 = '" & RTrim(iH1_Ym) & "'"
                        Sl_Where = Sl_Where & " AND 契約先 = '" & RTrim(S310_Torcd) & "'"
                    End If
                    
                    gStr = " FCD + 識別子 = '" & Trim(S310_FCD) & RTrim(S310_SCD) & "'"
                    If S310_PrtSyu = 0 Then
                        If PrtSelF(0) = "" Then
                            PrtSelF(0) = " AND (" & gStr
                        Else
                            PrtSelF(0) = PrtSelF(0) & " OR " & gStr
                        End If
                    Else
                        If PrtSelT(0) = "" Then
                            PrtSelT(0) = " AND (" & gStr
                        Else
                            PrtSelT(0) = PrtSelT(0) & " OR " & gStr
                        End If
                    End If
                    S310_TotalCnt = S310_TotalCnt + CisFun.Val2(.TextMatrix(gInt, CisVsGrid1.FixedGet(5)))
                    
                    KeyTorcd = S310_Torcd
                Else
                    '--(( CSV用SELECT分作成 ))--
                    If RTrim(mCsvSelect) = "" Then
                        mCsvSelect = " WHERE 対象年月 = '" & RTrim(iH1_Ym) & "'"
                        mCsvSelect = mCsvSelect & " AND ( "
                    Else
                        mCsvSelect = mCsvSelect & " OR "
                    End If
                    mCsvSelect = mCsvSelect & " (契約先 = '" & RTrim(S310_Torcd) & "'"
                    mCsvSelect = mCsvSelect & " AND FCD + 識別子 = '" & Trim(S310_FCD) & RTrim(S310_SCD) & "')"
                        
                    S310_TotalCnt = S310_TotalCnt + CisFun.Val2(.TextMatrix(gInt, CisVsGrid1.FixedGet(5)))
                End If
            End If
        Next
    End With
    If Not Op_Out(2).Value Then
        If RTrim(KeyTorcd) <> "" Then
            If RTrim(PrtSelF(0)) <> "" Then PrtSelF(TCnt) = Sl_Where & PrtSelF(0) & ")"
            If RTrim(PrtSelT(0)) <> "" Then PrtSelT(TCnt) = Sl_Where & PrtSelT(0) & ")"
        End If
    Else
        If RTrim(mCsvSelect) <> "" Then
            mCsvSelect = mCsvSelect & " )"
        End If
    End If
    If S310_TotalCnt = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    印刷対象データが存在しません｡            "
            .MB_Title = "抽出処理"
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
'-----------------------------'
'   出力処理
'-----------------------------'
    If Not Op_Out(2).Value Then
    '【 印刷 】
        For TCnt = 1 To UBound(PrtSelF)
            S310_GKbn = "": S310_SelCnt = TCnt
            If RTrim(PrtSelF(TCnt)) <> "" Then
                If RTrim(PrtSelT(TCnt)) = "" Then               '契約先計印刷
                    S310_GKbn = "G"
                End If
                If DataReadF Then
                    With CisPrt
                        Set .iPrtForm = COS0310L
                        .cDataMaxCnt = CisDB.RecordCount
                        .dpMSG = PrtTorcd(TCnt) & " 印刷中"
                        If Op_Out(0).Value Then
                           .iPreview = pv_Print
                        Else
                           .iPreview = pv_Preview
                        End If
                        .PrintStart
                    End With
                End If
            End If
            If RTrim(PrtSelT(TCnt)) <> "" Then
                S310_GKbn = "G"                                 '契約先計印刷
                If DataReadT Then
                    With CisPrt
                        Set .iPrtForm = COS0311L
                        .cDataMaxCnt = CisDB.RecordCount
                        .dpMSG = PrtTorcd(TCnt) & " 印刷中"
                        If Op_Out(0).Value Then
                           .iPreview = pv_Print
                        Else
                           .iPreview = pv_Preview
                        End If
                        .PrintStart
                    End With
                End If
            End If
        Next
    Else
    '【 CSV出力 】
        Call CsvPut
    End If
    
    PrintProc = True
End Function
'****************************
'*       CSV出力処理         *
'****************************
Private Function CsvPut() As Boolean
    CsvPut = False

    Dim OutStr      As String
    Dim OutFNo      As Long
    Dim OutCnt      As Long
    
    gSL_Select = "SELECT TW.*,"
    gSL_Select = gSL_Select & " ISNULL(SY1.値名称,'') 契約先名,ISNULL(SY2.値名称,'') 識別子名"
    gSL_Select = gSL_Select & " FROM TEXT買掛金テーブル TW"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "   ON  SY1.区分名称 = '照合対象契約先'"
    gSL_Select = gSL_Select & "   AND SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY1.値 = 契約先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "   ON  SY2.区分名称 = 'TEXT買掛金'"
    gSL_Select = gSL_Select & "   AND SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY2.値 = 契約先 + FCD + 識別子"
    gSL_Select = gSL_Select & mCsvSelect
    gSL_Select = gSL_Select & " ORDER BY 対象年月,契約先,FCD,識別子,検収日,品番,納入番号,受入"
    
    If Not TKKRead(gSL_Select, 0) Then
        Call TKKClose
        Exit Function
    End If
    
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
'
'    lb_msg.Caption = "【 買掛金明細リスト 出力中 】"
    Csv_Bar.Visible = True
    
    DoEvents
        
    OutFNo = FreeFile
    Open mCsvDir For Output As #OutFNo

    Call CisFrm.MousePT(11)
    
    With CisFrm
        Set .MeterPic = Csv_Bar
        .MeterMax = CisDB.RecordCount
    End With
        
    OutStr = "№,対象年月,契約先,契約先名,FCD,識別子,品番,検収日,納入番号,数量,単価,金額,"
    OutStr = OutStr & "管理番号,受入,カード,データ,購買,単位,重量,品名,依頼部署,貸借,工区,"
    OutStr = OutStr & "支給先,持込先,出庫者,支給区分,科目コード,繰越,補助品番,自動遡り区分,"
    OutStr = OutStr & "購入区分,仮号試,訂正理由,決済番号,部品管理区分,仮単価,マイナス,データ元,"
    OutStr = OutStr & "型台帳№,支払開始年月,設定月数,経過月数,型投資額,照合年月,照合区分"
    Print #OutFNo, OutStr
    OutCnt = 0
    Do Until Not TKK_RDSTS
        With TKK
            OutCnt = OutCnt + 1
            OutStr = OutCnt & ","
            OutStr = OutStr & Mid(.対象年月, 1, 4) & "/" & Mid(.対象年月, 5, 2) & ","
            OutStr = OutStr & RTrim(.契約先) & "," & RTrim(.契約先名) & ","
            OutStr = OutStr & RTrim(.FCD) & ","
            OutStr = OutStr & RTrim(.識別子) & ":" & RTrim(.識別子名) & ",_"
            OutStr = OutStr & RTrim(.品番) & ","
            If RTrim(.検収日) <> "" Then
                OutStr = OutStr & Left(.検収日, 4) & "/" & Mid(.検収日, 5, 2) & "/" & Mid(.検収日, 7, 2) & ","
            Else
                OutStr = OutStr & ","
            End If
            OutStr = OutStr & RTrim(.納入番号) & ","
            OutStr = OutStr & .数量 & ","
            OutStr = OutStr & .単価 & ","
            OutStr = OutStr & .金額 & ","
            OutStr = OutStr & RTrim(.管理番号) & ","
            OutStr = OutStr & RTrim(.受入) & ","
            OutStr = OutStr & RTrim(.カード) & ","
            OutStr = OutStr & RTrim(.データ) & ","
            OutStr = OutStr & RTrim(.購買) & ","
            OutStr = OutStr & RTrim(.単位) & ","
            OutStr = OutStr & .重量 & ","
            OutStr = OutStr & RTrim(.品名) & ","
            OutStr = OutStr & RTrim(.依頼部署) & ","
            OutStr = OutStr & RTrim(.貸借) & ","
            OutStr = OutStr & RTrim(.工区) & ","
            OutStr = OutStr & RTrim(.支給先) & ","
            OutStr = OutStr & RTrim(.持込先) & ","
            OutStr = OutStr & RTrim(.出庫者) & ","
            OutStr = OutStr & RTrim(.支給区分) & ","
            OutStr = OutStr & RTrim(.科目コード) & ","
            OutStr = OutStr & RTrim(.繰越) & ","
            OutStr = OutStr & RTrim(.補助品番) & ","
            OutStr = OutStr & RTrim(.自動遡り区分) & ","
            OutStr = OutStr & RTrim(.購入区分) & ","
            OutStr = OutStr & RTrim(.仮号試) & ","
            OutStr = OutStr & RTrim(.訂正理由) & ","
            OutStr = OutStr & RTrim(.決済番号) & ","
            OutStr = OutStr & RTrim(.部品管理区分) & ","
            OutStr = OutStr & RTrim(.仮単価) & ","
            OutStr = OutStr & RTrim(.マイナス) & ","
            OutStr = OutStr & RTrim(.データ元) & ","
            OutStr = OutStr & RTrim(.型台帳NO) & ","
            OutStr = OutStr & RTrim(.支払開始年月) & ","
            OutStr = OutStr & .設定月数 & ","
            OutStr = OutStr & .経過月数 & ","
            OutStr = OutStr & .型投資額 & ","
            OutStr = OutStr & RTrim(.照合年月) & ","
            OutStr = OutStr & RTrim(.照合区分) & ","

            Print #OutFNo, OutStr
    
            CisFrm.MeterUpdate OutCnt

            Call TKKReadNext
        End With
    Loop

CsvPut_Ed:
    Call TKKClose

    Close #OutFNo
    SaveSetting appname:=gRegKey, section:="CSV", _
                    Key:=mCsvKey, setting:=RTrim(lb_PrtDev)
    
    Call CisFrm.MousePT(1)
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "   出力処理が終了しました。    "
        .MB_MSG(4) = "   　 " & "出力先    = " & mCsvDir
        .MB_MSG(5) = "   　 " & "出力件数 = " & StrConv(Format(CisDB.RecordCount, "##,##0"), vbWide) & " 件 "
        .MB_Title = "出力確認"
        .MB_Button = OK
        .MBOX
    End With
    
    Csv_Bar.Visible = False
    CsvPut = True
    
End Function

