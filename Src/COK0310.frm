VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form COK0310 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "買掛実績明細リスト　発行"
   ClientHeight    =   10980
   ClientLeft      =   90
   ClientTop       =   1515
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
      Left            =   1020
      Top             =   6225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   870
      Left            =   2850
      ScaleHeight     =   810
      ScaleWidth      =   9075
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9390
      Width           =   9135
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   750
         Left            =   15
         Top             =   45
         Width           =   9000
         _ExtentX        =   15875
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
            Left            =   4365
            ScaleHeight     =   360
            ScaleWidth      =   4455
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   180
            Visible         =   0   'False
            Width           =   4515
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "　出　力　先"
            Height          =   600
            Left            =   90
            TabIndex        =   11
            Top             =   45
            Width           =   4230
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "ﾌﾟﾘﾝﾀ"
               Height          =   240
               Index           =   0
               Left            =   180
               TabIndex        =   3
               Top             =   270
               Value           =   -1  'True
               Width           =   1035
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "画面(ﾌﾟﾚﾋﾞｭｰ)"
               Height          =   240
               Index           =   1
               Left            =   1335
               TabIndex        =   4
               Top             =   270
               Width           =   1875
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "CSV"
               Height          =   240
               Index           =   2
               Left            =   3315
               TabIndex        =   5
               Top             =   270
               Width           =   750
            End
         End
         Begin Cis3D_v60.CIS3D lb_PrtDev 
            Height          =   420
            Left            =   4365
            Top             =   180
            Width           =   4185
            _ExtentX        =   7382
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
            Left            =   8550
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
      Left            =   2895
      ScaleHeight     =   1050
      ScaleWidth      =   5655
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1005
      Width           =   5715
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
      Left            =   2880
      ScaleHeight     =   7005
      ScaleWidth      =   9045
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2280
      Width           =   9105
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   6885
         Left            =   60
         Top             =   60
         Width           =   8925
         _ExtentX        =   15743
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
            Width           =   8715
            _Version        =   196608
            _ExtentX        =   15372
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
            FormatString    =   "№  |支払先                      |識別子|名  称                |件数    ||||"
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
      Caption         =   "【 買掛実績明細リスト　発行  】"
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
Attribute VB_Name = "COK0310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   買掛実績明細リスト　発行
'**       フォームID    :   COK0310
'**       処理概要      :
'**
'**       作  成  日    :   2004/12/13  By CIS
'**       変  更  日    :   2008/01/24  By CIS - 受給追加
'**       変  更  日    :   2008/06/15  By CIS - 材料実績 追加
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
'
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替

    Dim CisVsGrid1              As New CisVsGrid3

'   出力区分の背景色
    Dim mColor_On               As Long
    Dim mColor_Off              As Long

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
        .SelectCol = 7
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
    If MeisyoGet("処理内容", "POK0110") Then
        iH1_Ym = SYM.英数字1
    End If

'   >> CSVﾌｧｲﾙ出力先
    mCsvName = "買掛実績明細リスト"
    mCsvKey = "COK0310"
    mCsvDir = GetSetting(appname:=gRegKey, section:="CSV", _
                       Key:=mCsvKey, Default:=Mid(CisFun.WinRoot, 1, 3) & mCsvName & ".csv")
'   >> ﾌﾟﾘﾝﾀ出力先
    Op_Out(0).Value = True
    mPrtDev = Printer.DeviceName
    lb_PrtDev = mPrtDev
    mSavePrint = mPrtDev
'
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

    Select Case Key_Code
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click      '【検索】
           Case vbKeyF8:        If PB_Prt.Visible Then Call PB_Prt_Click        '【印刷】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click        '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click        '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                          '【入力】
                                '   Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                '   Dummy.Enabled = False
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
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
End Sub
'+----------------------------+
'+      取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 1
    RV_TorKb1 = 1
    
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RTrim(RV_TorcdK)
        Call iH1_Torcd_LostFocus
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
    If TorNmGet(iH1_Torcd, 5) Then
        H1lb_Tornm = TRM.略称
    End If
End Sub
'+----------------------------------------+
'+  Op_Out   出力先
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
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
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

' 印刷年月
    If Trim(iH1_Ym) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    印刷年月を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ym.SetFocus
        Exit Function
    End If
' 契約先
    If Trim(iH1_Torcd) <> "" Then
        If Not TorNmGet(iH1_Torcd, 5) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    契約先　未登録！（取引先マスタ）            "
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
'
'-------------------'
'   印刷対象のｾｯﾄ
'-------------------'
    Dim TCnt        As Integer

    ReDim PrtTorcd(1)
    ReDim PrtDidcd(1)
    TCnt = 0
    With vsGrid1
        For gInt = 1 To .Rows - 1
            If .TextMatrix(gInt, CisVsGrid1.SelectCol) <> "" Then
                If PrtTorcd(TCnt) = .TextMatrix(gInt, CisVsGrid1.FixedGet(6)) Then
                    PrtDidcd(TCnt) = PrtDidcd(TCnt) & "," & .TextMatrix(gInt, CisVsGrid1.FixedGet(5))
                Else
                    TCnt = TCnt + 1
                    PrtTorcd(TCnt) = .TextMatrix(gInt, CisVsGrid1.FixedGet(6))
                    PrtDidcd(TCnt) = .TextMatrix(gInt, CisVsGrid1.FixedGet(5))
                    ReDim Preserve PrtTorcd(TCnt + 1)
                    ReDim Preserve PrtDidcd(TCnt + 1)
                End If
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
    PrtFlagHyosi = False
    PrtFlagSire = True
    PrtFlagSkyu = True
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
    On Error GoTo DBInput_Err:

    With CisFun
         .Where_CreateKB = NewDocument
         ' 実績年月
         .Where_Create AlphaMe, "KJ.実績年月", iH1_Ym
         ' 契約先
         .Where_Create AlphaMe, "KJ.契約先", iH1_Torcd
    End With
    
    gSL_Select = "SELECT KJ.実績年月,KJ.契約先,ISNULL(TR.略称,'') 契約先名,"
    gSL_Select = gSL_Select & "     KJ.識別子,ISNULL(SY.値名称,'') 識別子名,COUNT(*) 件数"
    gSL_Select = gSL_Select & " FROM 購入実績テーブル KJ"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR.取引先CD = KJ.契約先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "   ON SY.区分名称 = '買掛識別子'"
    gSL_Select = gSL_Select & "   AND SY.値 = KJ.識別子"
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    gSL_Select = gSL_Select & " GROUP BY KJ.実績年月,KJ.契約先,KJ.識別子,TR.略称,SY.値名称"
    gSL_Select = gSL_Select & " Union"
    gSL_Select = gSL_Select & " SELECT KJ.実績年月,KJ.契約先,ISNULL(TR.略称,'') 契約先名,"
    gSL_Select = gSL_Select & "     KJ.識別子,ISNULL(SY.値名称,'') 識別子名,COUNT(*) 件数"
    gSL_Select = gSL_Select & " FROM 支給実績テーブル KJ"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR.取引先CD = KJ.契約先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "   ON SY.区分名称 = '買掛識別子'"
    gSL_Select = gSL_Select & "   AND SY.値 = KJ.識別子"
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    gSL_Select = gSL_Select & " GROUP BY KJ.実績年月,KJ.契約先,KJ.識別子,TR.略称,SY.値名称"
    
    '================================================ *** 2008/06/16 Start
    gSL_Select = gSL_Select & " Union"
    gSL_Select = gSL_Select & " SELECT KJ.実績年月,KJ.契約先,ISNULL(TR.略称,'') 契約先名,"
    gSL_Select = gSL_Select & "     KJ.識別子,ISNULL(SY.値名称,'') 識別子名,COUNT(*) 件数"
    gSL_Select = gSL_Select & " FROM 材料購入実績テーブル KJ"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR.取引先CD = KJ.契約先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "   ON SY.区分名称 = '買掛識別子'"
    gSL_Select = gSL_Select & "   AND SY.値 = KJ.識別子"
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    gSL_Select = gSL_Select & " GROUP BY KJ.実績年月,KJ.契約先,KJ.識別子,TR.略称,SY.値名称"
    gSL_Select = gSL_Select & " Union"
    gSL_Select = gSL_Select & " SELECT KJ.実績年月,KJ.契約先,ISNULL(TR.略称,'') 契約先名,"
    gSL_Select = gSL_Select & "     KJ.識別子,ISNULL(SY.値名称,'') 識別子名,COUNT(*) 件数"
    gSL_Select = gSL_Select & " FROM 材料支給実績テーブル KJ"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR.取引先CD = KJ.契約先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "   ON SY.区分名称 = '買掛識別子'"
    gSL_Select = gSL_Select & "   AND SY.値 = KJ.識別子"
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    gSL_Select = gSL_Select & " GROUP BY KJ.実績年月,KJ.契約先,KJ.識別子,TR.略称,SY.値名称"
    '================================================ *** 2008/06/16 End
    
    gSL_Select = gSL_Select & " ORDER BY KJ.実績年月,KJ.契約先,KJ.識別子"
    
    If Not KJSRead(gSL_Select) Then
       Call KJSClose
       Exit Function
    End If

    Call GridSet

    DBInput = True
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
        pTitle = "買掛実績明細リスト"
        Call COK0320LPrint(wPrintOut)
    Else
    '   >> 出力ﾌｧｲﾙの存在ﾁｪｯｸ
        pCsvKey = mCsvKey
        pCsvDir = mCsvDir
        If Not COK0320CsvCheck Then Exit Function

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
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 左中, あり, 28, 1, "契　約　先")
        Call .FixedSet(中中, 中中, あり, 6, 2, "識別子")
        Call .FixedSet(中中, 左中, あり, 22, 3, "名　　称")
        Call .FixedSet(中中, 右中, あり, 7, 4, "件数")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 5, "S識別子")
        Call .FixedSet(中中, 左中, なし, 0, 6, "S契約先")
        Call .FixedSet(中中, 左中, なし, 0, 7, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'       Call .EditSet(iH_Text1, 1)
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
        Do Until Not KJS_RDSTS
            ID = ID + 1
                
            .TextMatrix(ID, 1) = KJS.契約先 & Space(7 - CisFun.Len2(KJS.契約先)) & ":" & KJS.契約先名
            .TextMatrix(ID, 2) = KJS.識別子
            .TextMatrix(ID, 3) = KJS.識別子名
            .TextMatrix(ID, 4) = Format(KJS.件数, "##,###")
            .TextMatrix(ID, 5) = KJS.識別子
            .TextMatrix(ID, 6) = KJS.契約先

            Call KJSReadNext
            .Row = ID
            Call CisVsGrid1.vsColor
        Loop
        
        .Row = 1: .Col = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With

    GridSet = True
    
    CisVsGrid1.ReNumber
    
    Call KJSClose
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
'*****      行選択（スペース）
'******************************************'
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If ProcHB = "B1" Then
        If KeyCode = vbKeySpace Then
            Call CisVsGrid1.vsColor
        End If
    End If
End Sub
