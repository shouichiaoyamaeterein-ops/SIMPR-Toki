VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form COK0320 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "支払金額連絡書　発行"
   ClientHeight    =   10980
   ClientLeft      =   -150
   ClientTop       =   1005
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10980
   ScaleWidth      =   15315
   WindowState     =   2  '最大化
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   45
      Top             =   9270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   870
      Left            =   465
      ScaleHeight     =   810
      ScaleWidth      =   14100
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9405
      Width           =   14160
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   750
         Left            =   15
         Top             =   45
         Width           =   14040
         _ExtentX        =   24765
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
            Left            =   7050
            ScaleHeight     =   360
            ScaleWidth      =   6030
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   180
            Visible         =   0   'False
            Width           =   6090
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "　出　力　先"
            Height          =   600
            Left            =   90
            TabIndex        =   10
            Top             =   45
            Width           =   6735
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "ﾌﾟﾘﾝﾀ"
               Height          =   240
               Index           =   0
               Left            =   225
               TabIndex        =   2
               Top             =   270
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "画面(ﾌﾟﾚﾋﾞｭｰ)"
               Height          =   240
               Index           =   1
               Left            =   2310
               TabIndex        =   3
               Top             =   270
               Width           =   1875
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "CSV"
               Height          =   240
               Index           =   2
               Left            =   4935
               TabIndex        =   4
               Top             =   270
               Width           =   975
            End
         End
         Begin Cis3D_v60.CIS3D lb_PrtDev 
            Height          =   420
            Left            =   7050
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
            Left            =   12810
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
      Left            =   495
      ScaleHeight     =   1050
      ScaleWidth      =   8670
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1005
      Width           =   8730
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   990
         Left            =   1905
         Top             =   30
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   1746
         ForeColor       =   16711680
         Caption         =   "契　約　先"
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
         cPositionY      =   60
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
         Caption         =   "印 刷 年 月"
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
         cPositionY      =   60
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
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   990
         Left            =   5625
         Top             =   30
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1746
         ForeColor       =   16711680
         Caption         =   "明 細 発 行"
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
         cPositionY      =   60
         Begin VB.CheckBox B1Chk_PrtFlag 
            BackColor       =   &H00C0C0C0&
            Caption         =   "支給明細"
            Height          =   300
            Index           =   2
            Left            =   1545
            TabIndex        =   14
            Top             =   450
            Width           =   1275
         End
         Begin VB.CheckBox B1Chk_PrtFlag 
            BackColor       =   &H00C0C0C0&
            Caption         =   "仕入明細"
            Height          =   300
            Index           =   1
            Left            =   225
            TabIndex        =   13
            Top             =   450
            Width           =   1275
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
      Left            =   10275
      ScaleHeight     =   660
      ScaleWidth      =   4260
      TabIndex        =   7
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
      Height          =   6930
      Left            =   480
      ScaleHeight     =   6870
      ScaleWidth      =   14085
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2460
      Width           =   14145
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   6765
         Left            =   60
         Top             =   60
         Width           =   13980
         _ExtentX        =   24659
         _ExtentY        =   11933
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
         Begin vsFlexLib.vsFlexArray VsGrid1 
            Height          =   6585
            Left            =   60
            TabIndex        =   12
            Top             =   90
            Width           =   13860
            _Version        =   196608
            _ExtentX        =   24447
            _ExtentY        =   11615
            _StockProps     =   228
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "No|契 約 先              |請求書№ |前月繰越  |仕入金額  |消費税  |相殺金額  |消費税  |請求金額  |支払金額  |繰越金額  |||"
            Rows            =   19
            Cols            =   22
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   330
            SelectionMode   =   1
            AllowUserResizing=   1
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
      TabIndex        =   5
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
      Caption         =   "【 支払金額連絡書　発行  】"
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
Attribute VB_Name = "COK0320"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   支払金額連絡書　発行
'**       フォームID    :   COK0320
'**       処理概要      :
'**
'**       作  成  日    :   2004/12/09  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替

    Dim CisVsGrid1      As New CisVsGrid3

'   出力区分の背景色
    Dim mColor_On       As Long
    Dim mColor_Off      As Long

    Dim mPrtDev         As String       ' ﾌﾟﾘﾝﾀ出力先
    Dim mCsvDir         As String
    Dim mCsvKey         As String
    Dim mCsvName        As String
    Dim mSavePrint      As String
    Dim wPrintOut       As Integer
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
    gConnectCount = 3
    If Not IniGet Then End
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Set CisFrm = New CISFormContorl
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
'
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
'
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid1
        Set .GridObj = VsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 320
        .PatanMax = 1
        .SelectCol = 12
        Set .SelectDispObj = B1lb_SKensu
        .InitGet = False
        .Init
    End With
    Call GridInit
'
'   背景色の設定（出力区分）
    mColor_On = Op_Out(0).BackColor
    mColor_Off = Op_Out(1).BackColor
'
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
'
'   >> 印刷年月の初期表示値を求める
    gSL_Select = ""
    gSL_Select = gSL_Select & "select max(実績年月) 実績年月"
    gSL_Select = gSL_Select & "  from 買掛支払テーブル "
    If KSHRead(gSL_Select, 1) Then
        iH1_Ym = KSH.実績年月
    End If
'
'   >> CSVﾌｧｲﾙ出力先
    mCsvName = "支払金額連絡表"
    mCsvKey = "COK0320"
    mCsvDir = GetSetting(appname:=gRegKey, section:="CSV", _
                       Key:=mCsvKey, Default:=Mid(CisFun.WinRoot, 1, 3) & mCsvName & ".csv")
'   >> ﾌﾟﾘﾝﾀ出力先
    Op_Out(0).Value = True
    mPrtDev = Printer.DeviceName
    lb_PrtDev = mPrtDev
    mSavePrint = mPrtDev
'
'   >> 帳票可変項目テーブル追加・処理
    Call DBInsert_CKT
'   >> 名称マスタ（帳票管理）追加・処理
    Call DBInsert_SYM
'
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
'
    Select Case Key_Code
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click      '【検索】
           Case vbKeyF8:        If PB_Prt.Visible Then Call PB_Prt_Click        '【印刷】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_Can_Click        '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click        '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                          '【入力】
                                '   Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                '   Dummy.Enabled = False
                                End If
           Case Else:           Exit Function
    End Select
'
    Key_Acc = True
End Function
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
Private Sub PB_Can_Click()
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
        If Body1Chk Then
            If PrintProc Then
                Call HeadBodyClear("B")
                ProcHB = "H1"
                Call DispChange(ProcHB)
            End If
        End If
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    Select Case PB_Look.Tag
        Case "iH1_Torcd":  Call Look_Torcd
    End Select
End Sub
'+----------------------------+
'+       契約先検索
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 1
    RV_TorKb1 = 1
    RV_TorKb1 = 0
'
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing
'
    If RV_Rtn Then
        iH1_Torcd = RTrim$(RV_TorcdK)
        Call iH1_Torcd_LostFocus
    End If
End Sub
'+-----------------------------+
'+      iH1_Torcd
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
    If RTrim$(iH1_Torcd) <> "" Then
        If TorNmGet(iH1_Torcd, 5) Then
            H1lb_Tornm = TRM.略称
        End If
    End If
End Sub
'+----------------------------------------+
'+      Op_Out   出力先
'+----------------------------------------+
Private Sub Op_Out_GotFocus(Index As Integer)
    wPrintOut = Index
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
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
        Case "H1": H1Mode = True: H1Color = gPGotSel: PB_ENT.BtnCaption = "入力"
        Case "B1": B1Mode = True: B1Color = gPGotSel: PB_ENT.BtnCaption = "選択"
    End Select
    H1_Area1.Enabled = H1Mode: H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode: B1_Area1.BackColor = B1Color

    PB_Look.Visible = False
    PB_Prt.Visible = Not H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode

'   ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
'
    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = "": .Enabled = True
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
        Call VSGrid1_Click
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
'
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'   >> 印刷年月
    If Trim$(iH1_Ym) = "" Then
        CisFun.MB_MSG(2) = "    印刷年月を入力して下さい            "
        CisFun.MBOX
        iH1_Ym.SetFocus
        Exit Function
    End If
'   >> 契約先
    If RTrim$(iH1_Torcd) <> "" Then
        If Not TorNmGet(iH1_Torcd, 5) Then
            CisFun.MB_MSG(2) = "    契約先　未登録！（取引先マスタ）            "
            CisFun.MBOX
            iH1_Torcd.SetFocus
            Exit Function
        End If
    End If
'
'   >> 対象データ存在確認
    Call SqlSelectSet
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
'
'-------------------'
'   印刷対象のｾｯﾄ
'-------------------'
    Dim TCnt        As Integer

    ReDim PrtTorcd(1)
    ReDim PrtDidcd(1)
    TCnt = 0
    With VsGrid1
        For gInt = 1 To .Rows - 1
            If .TextMatrix(gInt, CisVsGrid1.SelectCol) <> "" Then
                TCnt = TCnt + 1
                PrtTorcd(TCnt) = .TextMatrix(gInt, CisVsGrid1.FixedGet(11))
                PrtDidcd(TCnt) = ""
                ReDim Preserve PrtTorcd(TCnt + 1)
                ReDim Preserve PrtDidcd(TCnt + 1)
            End If
        Next
    End With
'
    If TCnt = 0 Then
        With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    印刷データを選択して下さい            "
                .MB_Title = "抽出処理"
                .MB_Button = Error
                .MBOX
        End With
        Exit Function
    End If
'
    PrtYM = Trim$(iH1_Ym)
    PrtFlagHyosi = True
    If B1Chk_PrtFlag(1).Value = 1 Then PrtFlagSire = True Else PrtFlagSire = False
    If B1Chk_PrtFlag(2).Value = 1 Then PrtFlagSkyu = True Else PrtFlagSkyu = False
'
    Body1Chk = True
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput() As Boolean
    DBInput = False
'
    On Error GoTo DBInput_Err:
'
    If Not Grid1Set Then Exit Function
'
    DBInput = True
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'****************************
'*      SQL文ｾｯﾄ処理
'****************************
Private Sub SqlSelectSet()
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  ksh.*"
    gSL_Select = gSL_Select & ", isnull(trm.略称, '') 契約先名"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  買掛支払テーブル ksh"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = ksh.契約先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    gSL_Select = gSL_Select & " where ksh.実績年月 = '" & Trim$(iH1_Ym) & "'"
    If RTrim$(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & "   and ksh.契約先 = '" & RTrim$(iH1_Torcd) & "'"
    End If
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  ksh.契約先"
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    請求書NO 採番     +
'+----------------------+
Private Function DBPut(Optional DBNo As Integer = 0) As Boolean
    DBPut = False

    On Error GoTo DBPut_Err
    Call CisDB.DBTran(TransBegin, DBNo)     ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 更新処理 )
    Dim wIndex      As Integer

    For wIndex = 1 To UBound(PrtTorcd)
        gSL_Select = ""
        gSL_Select = gSL_Select & "update 買掛支払テーブル"
        gSL_Select = gSL_Select & "   set 請求書NO ="
    '   ↓ 最終の請求書NOの算出開始
        gSL_Select = gSL_Select & " substring(実績年月, 3, 4) + right('000' + cast(cast("
        gSL_Select = gSL_Select & "("
        gSL_Select = gSL_Select & "select max(cast(case when 請求書NO = '' then '000' else substring(請求書NO, 5, 3) end as numeric))"
        gSL_Select = gSL_Select & "  from 買掛支払テーブル"
        gSL_Select = gSL_Select & " where 実績年月 = '" & Trim$(PrtYM) & "'"
        gSL_Select = gSL_Select & ") + 1"
        gSL_Select = gSL_Select & " as decimal (3)) as varchar (3)), 3)"
    '   ↑ 最終の請求書NOの算出終了
        gSL_Select = gSL_Select & " where 実績年月 = '" & Trim$(PrtYM) & "'"
        gSL_Select = gSL_Select & "   and 請求書NO = ''"
        gSL_Select = gSL_Select & "   and 契約先 = '" & RTrim$(PrtTorcd(wIndex)) & "'"
        With CisDB
            .ConnectNo = DBNo
            .SQL = gSL_Select
            If Not .DBExec Then GoTo DBPut_Err
        End With
    Next wIndex

    Call CisDB.DBTran(TransCommit, DBNo)    ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    On Error GoTo 0

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
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
        CommonDialog1.Copies = 1
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
'+---------------------------------------------+
'+      印刷処理                               +
'+---------------------------------------------+
Private Function PrintProc() As Boolean
    PrintProc = False
'
    If wPrintOut <> 2 Then
        Call DBPut
        pTitle = "支払金額連絡書"
        Call COK0320LPrint(wPrintOut)
    Else
    '   >> 出力ﾌｧｲﾙの存在ﾁｪｯｸ
        pCsvKey = mCsvKey
        pCsvDir = mCsvDir
        If Not COK0320CsvCheck Then Exit Function

        Call DBPut
        Set CisFrmCOK0320.MeterPic = Csv_Bar
        Csv_Bar.Visible = True
        DoEvents
        Call COK0320LPrint(wPrintOut)
        Csv_Bar.Visible = False
    End If
'
    PrintProc = True
End Function



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
        Call .FixedSet(中中, 右中, あり, 2, 0, "No")
        Call .FixedSet(中中, 左中, あり, 28, 1, "契約先")
        Call .FixedSet(中中, 中中, あり, 8, 2, "請求書№")
        Call .FixedSet(中中, 右中, あり, 11, 3, "前月繰越")
        Call .FixedSet(中中, 右中, あり, 11, 4, "仕入金額")
        Call .FixedSet(左中, 右中, あり, 9, 5, "　消費税　仕入")
        Call .FixedSet(中中, 右中, あり, 11, 6, "相殺金額")
        Call .FixedSet(左中, 右中, あり, 9, 7, "　消費税　相殺")
        Call .FixedSet(中中, 右中, あり, 11, 8, "請求金額")
        Call .FixedSet(中中, 右中, あり, 11, 9, "支払金額")
        Call .FixedSet(中中, 右中, あり, 11, 10, "繰越金額")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 11, "X契約先")
        Call .FixedSet(中中, 左中, なし, 0, 12, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'       Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'***********************************************'
'*****      グリッド内容セット
'***********************************************'
Private Function Grid1Set(Optional DBNo As Integer = 0) As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
'
    Grid1Set = False
'
    If KSHRead(gSL_Select, , DBNo) Then
'
'       抽出処理 初期ｾｯﾄ
        B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
        Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
'       ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
        With VsGrid1
           .Redraw = False
            If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
                .Rows = CisVsGrid1.DispMax
            Else
                .Rows = Ttl_Kensu + 1
            End If
'
            ID = 0
            Do Until Not KSH_RDSTS
                ID = ID + 1
'
                Call DBInfoToGrid1(ID)
                .Row = ID
                Call CisVsGrid1.vsColor
'
                Call KSHReadNext(DBNo)
            Loop
            CisVsGrid1.ReNumber
'
            .Row = 1: .Col = 1: .ColSel = .Cols - 1
            .Redraw = True
        End With
'
        Grid1Set = True
    End If
    Call KSHClose(DBNo)
End Function
'+----------------------------------------------+
'+      テーブルよりグリッドに転送
'+----------------------------------------------+
Private Sub DBInfoToGrid1(SetRow As Long)
    With VsGrid1
    '>> 契約先名
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(1)) = KSH.契約先 & Space(7 - CisFun.Len2(KSH.契約先)) & ":" & RTrim$(KSH.契約先名)
    '>> 請求書NO
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = RTrim$(KSH.請求書NO)
    '>> 前月繰越
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(3)) = Format(KSH.前月繰越, "#,###")
    '>> 仕入金額
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(4)) = Format(KSH.仕入金額, "#,###")
    '>> 仕入消費税
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(5)) = Format(KSH.仕入消費税, "#,###")
    '>> 相殺金額
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(6)) = Format(KSH.相殺金額, "#,###")
    '>> 相殺消費税
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(7)) = Format(KSH.相殺消費税, "#,###")
    '>> 請求金額
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(8)) = Format(KSH.請求金額, "#,###")
    '>> 支払金額
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(9)) = Format(KSH.支払金額, "#,###")
    '>> 繰越金額
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(10)) = Format(KSH.繰越金額, "#,###")
    '>> 退避
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(11)) = RTrim$(KSH.契約先)
    End With
End Sub
'******************************************'
'*****      行選択（クリック）
'******************************************'
Private Sub VSGrid1_Click()
    If ProcHB = "B1" Then
        Call CisVsGrid1.vsColor
    End If
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
