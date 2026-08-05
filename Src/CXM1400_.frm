VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CXM1400 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "工程かんばん発行内容パターン登録"
   ClientHeight    =   11025
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   15195
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   14.25
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
   ScaleHeight     =   11025
   ScaleWidth      =   15195
   WindowState     =   2  '最大化
   Begin VB.PictureBox H2_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   14205
      ScaleHeight     =   915
      ScaleWidth      =   1005
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   1065
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   795
         Left            =   60
         Top             =   75
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "値桁数"
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
         Begin CisText_V60.CisText iH2_Keta 
            Height          =   405
            Left            =   210
            TabIndex        =   1
            Top             =   315
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   714
            cFormat         =   "##"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "##"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
            cAutoEject      =   0   'False
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4800
      Left            =   3960
      ScaleHeight     =   4740
      ScaleWidth      =   10050
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3120
      Width           =   10110
      Begin Cis3D_v60.CIS3D Back_B1Area 
         Height          =   4620
         Left            =   60
         Top             =   60
         Width           =   9930
         _ExtentX        =   17515
         _ExtentY        =   8149
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
         Begin vsFlexLib.vsFlexArray VsGrid 
            Height          =   4620
            Left            =   0
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   9930
            _Version        =   196608
            _ExtentX        =   17515
            _ExtentY        =   8149
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
            FormatString    =   "登録№|部署   |１                            |返 |１                            |Select||"
            Rows            =   14
            Cols            =   8
            BackColor       =   12648384
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   330
            SelectionMode   =   1
            BackColorAlternate=   12648384
            Begin Cis3D_v60.CIS3D CIS3D2 
               Height          =   300
               Left            =   855
               Top             =   1935
               Width           =   4200
               _ExtentX        =   7408
               _ExtentY        =   529
               BackColor       =   -2147483633
               Caption         =   "部署１"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483633
               cBoderColor2    =   -2147483633
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cBoderStyle     =   2
            End
            Begin CisText_V60.CisText iB1_Kikai 
               Height          =   300
               Left            =   810
               TabIndex        =   10
               Top             =   330
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   529
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
            Begin Cis3D_v60.CIS3D B1lb_KikaiNM 
               Height          =   300
               Left            =   1740
               Top             =   330
               Width           =   3660
               _ExtentX        =   6456
               _ExtentY        =   529
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
            Begin CisText_V60.CisText iB1_KikaiNo 
               Height          =   300
               Left            =   5430
               TabIndex        =   11
               Top             =   330
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXX"
               MaxLength       =   3
            End
            Begin Cis3D_v60.CIS3D B1lb_KikaiNoNM 
               Height          =   300
               Left            =   5895
               Top             =   330
               Width           =   3660
               _ExtentX        =   6456
               _ExtentY        =   529
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   300
               Left            =   5475
               Top             =   1920
               Width           =   3000
               _ExtentX        =   5292
               _ExtentY        =   529
               BackColor       =   -2147483633
               Caption         =   "返却場所１"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483633
               cBoderColor2    =   -2147483633
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cFont3D         =   2
               cBoderStyle     =   2
            End
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   915
      ScaleHeight     =   1170
      ScaleWidth      =   5460
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1500
      Width           =   5520
      Begin Cis3D_v60.CIS3D UC_3D 
         Height          =   555
         Left            =   30
         Top             =   30
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   979
         ForeColor       =   16711680
         Caption         =   "部　　署"
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
         cPositionX      =   120
         Begin CisText_V60.CisText iH1_Kubun 
            Height          =   375
            Left            =   1170
            TabIndex        =   0
            Top             =   90
            Width           =   960
            _ExtentX        =   1693
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
            IMEMode         =   4
            cAutoEject      =   0   'False
         End
         Begin Cis3D_v60.CIS3D B1lb_KTorNm 
            Height          =   360
            Left            =   2130
            Top             =   90
            Width           =   2505
            _ExtentX        =   4419
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
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   555
         Left            =   30
         Top             =   585
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   979
         ForeColor       =   16711680
         Caption         =   "返却場所"
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
         cPositionX      =   120
         Begin CisText_V60.CisText CisText1 
            Height          =   375
            Left            =   1170
            TabIndex        =   12
            Top             =   90
            Width           =   450
            _ExtentX        =   794
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
            Text            =   "XXX"
            MaxLength       =   3
            IMEMode         =   4
            cAutoEject      =   0   'False
         End
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   360
            Left            =   1620
            Top             =   90
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
      Height          =   1350
      Left            =   9630
      ScaleHeight     =   1290
      ScaleWidth      =   4335
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1500
      Width           =   4395
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1755
         Top             =   675
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
         Left            =   2610
         Top             =   675
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
         Left            =   3465
         Top             =   675
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
      Begin CisBtn_60.CisBtn PB_BACK 
         Height          =   585
         Left            =   2610
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   3465
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1755
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
         Left            =   30
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
         Left            =   900
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
         Left            =   900
         Top             =   675
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
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   14190
      ScaleHeight     =   360
      ScaleWidth      =   1830
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1875
      Visible         =   0   'False
      Width           =   1890
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   345
         Left            =   0
         Top             =   0
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   609
         ForeColor       =   16711680
         Caption         =   "表示件数"
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
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B2lb_Kensu 
            Height          =   315
            Left            =   1035
            Top             =   15
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            Caption         =   "Z,ZZ9"
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
   End
   Begin Cis3D_v60.CIS3D K_Sykbnm 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      BackColor       =   14737632
      Caption         =   "追加"
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
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 工程かんばん発行内容パターン登録  】"
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
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   11940
         Top             =   60
         Width           =   2520
         _ExtentX        =   4445
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14520
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
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
   End
   Begin VB.Timer Timer3 
      Left            =   1440
      Top             =   75
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
      IMEMode         =   3  'ｵﾌ固定
      Left            =   330
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   630
      Top             =   60
   End
   Begin VB.Timer Timer2 
      Left            =   1050
      Top             =   60
   End
   Begin vsFlexLib.vsFlexArray VsGrid_ 
      Height          =   1560
      Left            =   7050
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7920
      Width           =   8940
      _Version        =   196608
      _ExtentX        =   15769
      _ExtentY        =   2752
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
      FormatString    =   "№ |値|値名称|英数字1 |英数字2|英数字3|数字1|数字2|数字3|X|作成日|更新日|"
      Rows            =   18
      Cols            =   13
      BackColor       =   12648447
      BackColorBkg    =   12632256
      FocusRect       =   0
      HighLight       =   2
      RowHeightMin    =   300
      SelectionMode   =   1
      AllowUserResizing=   1
      BackColorAlternate=   12648447
      Begin CisText_V60.CisText CisText2 
         Height          =   360
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   570
         _ExtentX        =   1005
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
         Text            =   "XXX"
         MaxLength       =   3
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   300
         Left            =   1050
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         BackColor       =   16777152
         Caption         =   "300t"
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
         cAlingnment     =   1
         cBoderStyle     =   1
         cPositionX      =   30
      End
      Begin CisText_V60.CisText CisText3 
         Height          =   360
         Left            =   2340
         TabIndex        =   14
         Top             =   360
         Width           =   810
         _ExtentX        =   1429
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
         Text            =   "XXX"
         MaxLength       =   3
      End
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   300
         Left            =   3150
         Top             =   360
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         BackColor       =   16777152
         Caption         =   "P56"
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
         cAlingnment     =   1
         cBoderStyle     =   1
         cPositionX      =   30
      End
   End
End
Attribute VB_Name = "CXM1400"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   名称マスタメンテナンス
'**       フォームID    :   CXM0010
'**       処理概要      :
'**
'**       作  成  日    :   2003/11/18  By CIS
'**       変  更  日    :
'**
'**********************************************************'

    Dim CisFrm          As CISFormContorl
    Dim mVSObj          As New CisVsGrid3     ' VSGridｸﾗｽ定義


    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    Dim SyoriKBB        As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNMB        As String       ' 処理名称(新規･修正･削除 等)
    
    Dim mSeigyo         As String       ' 名称マスタ管理区分
    Const c_CharWidth   As Long = 130
    
    Dim mLock           As Boolean      ' 照会フラグ
    
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    Select Case Trim(App.Title)
        Case "PXM0010"
                mSeigyo = "X"
                mLock = False
                PGName = "【 制御マスタメンテナンス  】"
        Case "PXM0013"
                mSeigyo = "X"
                mLock = True
                PGName = "【 制御マスタ 照会  】"
        Case "PXM0015"
                mSeigyo = ""
                mLock = False
                PGName = "【 名称マスタメンテナンス  】"
        Case "PXM0018"
                mSeigyo = ""
                mLock = True
                PGName = "【 名称マスタ 照会  】"
    End Select

    If Not mLock Then
    '   #-------------------#
    '   # 二 重 起 動 防 止  #
    '   #-------------------#
        If Not Dupli_Start Then End
    '   #-------------------#
    '   # 初 期 内 容 取 得  #
    '   #-------------------#
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
'
    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
    Timer3.Interval = 500    ' 処理区分点滅間隔
    Timer3.Enabled = False   ' 処理区分点滅禁止
'   #------------------#
'   #  ｸﾞﾘｯﾄﾞ初期処理   #
'   #------------------#
    With mVSObj
        Set .GridObj = VsGrid               ' ｸﾞﾘｯﾄﾞｵﾌﾞｼﾞｪｸﾄ名 ｾｯﾄ
        .FixedHeight = 350                  ' 固定行(見出し)高さ設定(省略時は300)
        .PatanMax = 1                       ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ最大数
        .PatanInit = 0                      ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ初期値
        .SelectCol = 5                      ' セル選択処理ﾜｰｸ列
'        Set .SelectDispObj = B1lb_SKensu    ' セル選択件数表示ｵﾌﾞｼﾞｪｸﾄｾｯﾄ
' ﾌﾟﾛｸﾞﾗﾑ作成時、以下のﾌﾗｸﾞをｾｯﾄし、ｸﾘｯﾌﾟﾎﾞｰﾄﾞよりｺｰﾄﾞの貼り付けを行う。(GridInit)
        .InitGet = False                    ' ｸﾞﾘｯﾄﾞ内容初期設定 ｸﾘｯﾌﾟﾎﾞｰﾄﾞ作成ﾌﾗｸﾞ

        .Init                               ' ｸﾞﾘｯﾄﾞ初期処理 実行
    End With
    
    
    
         Set .GridObj = VsGrid
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 19
        .InitGet = False
        .Init
   
    
    
''''    Call GridInit(2)                        ' ｸﾞﾘｯﾄﾞ初期表示
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

    If Not mLock Then
        Me.Show
    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
' 名称マスタ検索　確認
    Call ItemsClearSYM
    With SYM
        .区分名称 = "名称検索"
        .区分桁数 = 2
        .値 = "01"
        .値名称 = "名称検索"
        .英数字1 = "区   分   検   索"
        .英数字2 = "値"
        .英数字3 = "区分名"
        .制御区分 = "X"
        .作成者 = gTanto
    End With
    Call SYMReadChk
    
    FormAct = Not FormAct
    
'    If mLock Then
'        SyoriKB = "MNT"
'        iH1_Kubun = RV_Kubun
'        Call ReturnPress
'        On Error Resume Next
'        VsGrid.SetFocus
'        On Error GoTo 0
'    End If
'    If mLock Then
'        iH1_Kubun = RV_Kubun
'        SyoriKB = "MNT"
'        Call Head1Chk
'    End If
    
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
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyF5
                If Shift = 1 Then
                   Call ChgProc
                End If
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
    Cancel = 1
    Call PB_END_Click
End Sub
Private Sub ChgProc()
    If ProcHB <> "H1" Then Exit Sub
    
    If mSeigyo = "X" Then
       mSeigyo = ""
       mLock = False
       PGName = "【 名称マスタメンテナンス  】"
    Else
       mSeigyo = "X"
       mLock = False
       PGName = "【 制御マスタメンテナンス  】"
    End If
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    If Not mLock Then
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
    If mLock And ProcHB = "B1" Then
       Me.Hide
       GoTo PB_CAN_Ed
    End If
    If mLock And ProcHB = "H1" Then
       Me.Hide
       GoTo PB_CAN_Ed
    End If
   
   If ProcHB = "H2" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      SyoriKBB = ""
      Call SyoriNameB(SyoriKBB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      GoTo PB_CAN_Ed
   End If
   If ProcHB = "B1" And SyoriKBB = "MNT" Then
      ProcHB = "B2"
      Call DispChange(ProcHB)
      SyoriKBB = "MNT"
      Call SyoriNameB(SyoriKBB)
      Call mVSObj.vsColor(消去全部)
      VsGrid.SetFocus
      GoTo PB_CAN_Ed
   End If
   If (ProcHB = "B1" Or ProcHB = "B2") And SyoriKB = "ADD" Then
      ProcHB = "H2"
      Call DispChange(ProcHB)
      SyoriKBB = ""
      Call SyoriNameB(SyoriKBB)
      Call HeadBodyClear("B")
      GoTo PB_CAN_Ed
   End If
   If ProcHB = "B1" Or ProcHB = "B2" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      SyoriKBB = ""
      Call SyoriNameB(SyoriKBB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    If ProcHB = "H1" Then
        If SyoriKB <> "ADD" Then
            SyoriKB = "ADD"
            Call SyoriName(SyoriKB)
        End If
    Else
        If SyoriKBB <> "ADD" Then
            ProcHB = "B1"
            SyoriKBB = "ADD"
            iB1_Atai = ""
            Call DispChange(ProcHB)
            Call SyoriNameB(SyoriKBB)
            Call mVSObj.vsColor(消去全部)
        End If
    End If
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If ProcHB = "H1" Then
        If SyoriKB <> "MNT" Then
            SyoriKB = "MNT"
            Call SyoriName(SyoriKB)
        End If
    Else
        If SyoriKBB <> "MNT" Then
            ProcHB = "B2"
            SyoriKBB = "MNT"
            Call DispChange(ProcHB)
            Call SyoriNameB(SyoriKBB)
            VsGrid.SetFocus
            Call BodySet2
        End If
    End If
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If ProcHB = "H1" Then
        If SyoriKB <> "DEL" Then
            SyoriKB = "DEL"
            Call SyoriName(SyoriKB)
        End If
    Else
        If SyoriKBB <> "DEL" Then
            ProcHB = "B2"
            SyoriKBB = "DEL"
            Call DispChange(ProcHB)
            Call SyoriNameB(SyoriKBB)
            Call mVSObj.vsColor(消去全部)
            VsGrid.SetFocus
            Call BodySet2
        End If
    End If
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = "SELECT * FROM 名称マスタ "
    gSL_Select = gSL_Select & " WHERE  区分名称 < '" & RTrim(iH1_Kubun) & "'"
    gSL_Select = gSL_Select & "   AND  区分タイプ = 'K'"
    gSL_Select = gSL_Select & "   AND  制御区分 = '" & mSeigyo & "'"
    gSL_Select = gSL_Select & " ORDER BY 区分名称 DESC "
    Call DBInput("Back")
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = "SELECT * FROM 名称マスタ "
    gSL_Select = gSL_Select & " WHERE  区分名称 > '" & RTrim(iH1_Kubun) & "'"
    gSL_Select = gSL_Select & "   AND  区分タイプ = 'K'"
    gSL_Select = gSL_Select & "   AND  制御区分 = '" & mSeigyo & "'"
    gSL_Select = gSL_Select & " ORDER BY 区分名称 ASC "
    Call DBInput("Next")
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Kubun" Then Call Look_Kubun: Exit Sub
End Sub
'+----------------------------+
'+        区分名検索           +
'+----------------------------+
Private Sub Look_Kubun()

    RV_Left = 0
    RV_Top = 0
    RVI_Kubun = ""
    RVI_Seigyo = mSeigyo
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
       iH1_Kubun = RV_Kubun
       iH2_Keta = CisFun.RSetFld(RV_Code, 0, "##")
    End If

End Sub
'+----------------------------+
'+      区分名
'+----------------------------+
Private Sub iH1_Kubun_GotFocus()
    If Not mLock Then
        PB_Look.Visible = True
        PB_Look.Tag = ActiveControl.Name
    End If
End Sub
Private Sub iH1_Kubun_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
       PB_Look.Visible = False
       PB_Look.Tag = ""
    End If
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示(HEAD)  *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim W_Color     As Long
   
   If mLock Then Exit Sub
   
   W_Color = &H80000008 '初期色ｾｯﾄ(黒)
   Select Case SyoriKB
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
          W_Color = &HFF&
     Case "LOC"
          SyoriNM = "参照"
          K_Sykbnm.ForeColor = &HFF00FF  '
          PB_ADD.Visible = False
          PB_MNT.Visible = False
          PB_DEL.Visible = False
     Case Else
          SyoriNM = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
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
'**************************************
'*                                    *
'*      処理区分により区分名表示(BODY)  *
'*                                    *
'**************************************
Private Sub SyoriNameB(SYKB As String)
   Dim W_Color     As Long
   
   If mLock Then Exit Sub
   
   W_Color = &H80000008 '初期色ｾｯﾄ(黒)
   Select Case SyoriKBB
     Case "ADD"
          SyoriNMB = "追加"
          K_SykbnmB.ForeColor = &HFFFF&   '(黄色)
          PB_ADD.Visible = False
     Case "MNT"
          SyoriNMB = "修正"
          K_SykbnmB.ForeColor = &HFF0000  '(青色)
     Case "DEL"
          SyoriNMB = "削除"
          K_SykbnmB.ForeColor = &HFF&     '(赤色)
          W_Color = &HFF&
     Case Else
          SyoriNMB = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "iB*" Then gObj.ForeColor = W_Color
'      * PF-Key 表示設定 *
       If SyoriKBB <> "LOC" Then
          If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
             gObj.Name = "PB_DEL" Then
             If SyoriKBB = "" Then
                gObj.Visible = False
             Else
                gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKBB)
             End If
          End If
       End If
   Next gObj
   Timer3.Enabled = False
   If SyoriNMB = "" Then
      K_SykbnmB.Visible = False
      K_SykbnmB.Caption = ""
      Exit Sub
   End If
   K_SykbnmB.Caption = SyoriNMB
   K_SykbnmB.Visible = True
   Timer3.Enabled = True
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
'*******************************
'*      処理区分名点滅制御      *
'*******************************
Private Sub Timer3_Timer()
    If mLock Then Exit Sub
   If K_SykbnmB.Caption <> "" Then
      K_SykbnmB.Caption = ""
   Else
      K_SykbnmB.Caption = SyoriNMB
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
    Dim H2Mode      As Boolean
    Dim H2Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    Dim B2Mode      As Boolean
    Dim B2Color     As Long
    H1Mode = False
    H1Color = gPLostSel
    H2Mode = False
    H2Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
    B2Mode = False
    B2Color = gPLostSel
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
       Case "H2"
          H2Mode = True
          H2Color = gPGotSel
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
       Case "B2"
          B2Mode = True
          B2Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    H2_Area1.Enabled = H2Mode
    H2_Area1.BackColor = H2Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    PB_ADD.Visible = H1Mode Or B1Mode Or B2Mode
    PB_MNT.Visible = H1Mode Or B1Mode Or B2Mode
    PB_DEL.Visible = H1Mode Or B1Mode Or B2Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    
    VsGrid.TabStop = B2Mode
    
''''    Atai_Back.Enabled = True
    If Pro = "B2" Then
       B1_Area1.BackColor = B1Color
    End If
    If Pro = "B1" And SyoriKBB = "MNT" Then
''''       Atai_Back.Enabled = False
    End If
' 照会
    Back_B1Area.Enabled = True
    If mLock Then
        Back_B1Area.Enabled = False
        PB_ADD.Visible = False
        PB_MNT.Visible = False
        PB_DEL.Visible = False
        PB_Look.Visible = False
        PB_ENT.Visible = False
        PB_END.Visible = True
        PB_CAN.Visible = True
        B1_Area1.Enabled = True
        K_Sykbnm.Visible = False
''''        K_SykbnmB.Visible = False
    End If
'    PB_END.Visible = H1Mode
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
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
    If HeadBody = "H2" Or HeadBody = "B" Then
        Call mVSObj.Clear
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
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            GoTo ReturnPress_Ed
        End If
        If SyoriKB = "ADD" Then
           ProcHB = "H2"
           Call DispChange(ProcHB)
        Else
           ProcHB = "B1"
           Call DispChange(ProcHB)
           SyoriKBB = "ADD"
           Call SyoriNameB(SyoriKBB)
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "H2" Then
        If Not Head2Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        SyoriKBB = "ADD"
        Call SyoriNameB(SyoriKBB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        Call GridSet
        If SyoriKBB = "ADD" Then
        Else
           ProcHB = "B2"
           Call DispChange(ProcHB)
           SyoriKBB = "MNT"
           Call SyoriNameB(SyoriKBB)
           VsGrid.SetFocus
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B2" Then
       Call FromGrid
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    If iH1_Kubun = "" Then
        With CisFun
             .MB_Lines = 5
             .MB_MSG(2) = "     区分名を入力して下さい           "
             .MB_Title = "【必須エラー】"
             .MB_Button = Error
             .MBOX
        End With
       Exit Function
    End If
' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = "SELECT * FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = '" & RTrim(iH1_Kubun) & "'"
    gSL_Select = gSL_Select & "   AND 区分タイプ   = 'K'"
    If DBInput("INV") Then
       If SYM.制御区分 <> mSeigyo Then
          If mSeigyo = "" Then
             With CisFun
                  .MB_Lines = 5
                  .MB_MSG(2) = "    制御マスタ登録済(管理者用)            "
                  .MB_Title = ""
                  .MB_Button = Error
                  .MBOX
              End With
             Exit Function
          End If
          If mSeigyo <> "" Then
             With CisFun
                  .MB_Lines = 5
                  .MB_MSG(2) = "     名称マスタ登録済(一般用)            "
                  .MB_Title = ""
                  .MB_Button = Error
                  .MBOX
              End With
             Exit Function
          End If
       End If
       If SyoriKB = "ADD" Then
          With CisFun
              .MB_Lines = 5
              .MB_MSG(2) = "     名称マスタ登録済            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          Exit Function
       End If
    Else
       If SyoriKB <> "ADD" Then
          With CisFun
              .MB_Lines = 5
              .MB_MSG(2) = "     名称マスタ未登録            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          Exit Function
       End If
    End If
    Head1Chk = True
End Function
'****************************
'*      HEAD2ﾁｪｯｸ処理
'****************************
Private Function Head2Chk() As Boolean
    Head2Chk = False
    If CisFun.Val2(iH2_Keta) = 0 Then
       With CisFun
           .MB_Lines = 5
           .MB_MSG(2) = "     桁数を入力して下さい            "
           .MB_Title = "【必須エラー】"
           .MB_Button = Error
           .MBOX
       End With
       Exit Function
    End If
    
    iB1_Atai.MaxLength = CisFun.Val2(iH2_Keta)
    iB1_Atai.Width = c_CharWidth * CisFun.Val2(iH2_Keta) + 90
    iB1_Atai = ""
    
    Call GridInit(CisFun.Val2(iH2_Keta))                        ' ｸﾞﾘｯﾄﾞ初期表示
    
    Head2Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    If Trim(iB1_AtaiNm) = "" And Trim(iB1_Atai) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "     値名称を入力して下さい            "
            .MB_Title = "【必須エラー】"
            .MB_Button = Error
            .MBOX
        End With
       iB1_AtaiNm.SetFocus
       Exit Function
    End If
'   値存在確認
    If SyoriKBB = "ADD" Then
        gSL_Select = "SELECT * FROM 名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = '" & RTrim(iH1_Kubun) & "'"
        gSL_Select = gSL_Select & "   AND 区分タイプ   = ''"
        gSL_Select = gSL_Select & "   AND 値   = '" & RTrim(iB1_Atai) & "'"
        gSL_Select = gSL_Select & "   AND  制御区分 = '" & mSeigyo & "'"
        If SYMRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "     値 登 録 済 み            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB1_Atai.SetFocus
           Exit Function
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
    If Not SYMRead(gSL_Select, 1) Then
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
        Call HeadBodyClear("H2")
    Else
        Call HeadBodyClear("B")
        Call BodySet("")
    DBInput = True
    End If
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Sub BodySet(SetType As String)
    iH1_Kubun = SYM.区分名称
'
    iH2_Keta = CisFun.RSetFld(SYM.区分桁数, 0, "##")
    Call GridSet
    
    iB1_Atai.MaxLength = CisFun.Val2(iH2_Keta)
    iB1_Atai.Width = c_CharWidth * CisFun.Val2(iH2_Keta) + 90
    iB1_Atai = ""
End Sub
'+---------------------------+
'+      ｸﾞﾘｯﾄﾞより内容セット  +
'+---------------------------+
Private Sub BodySet2()
    With mVSObj
        If .GridObj.Row < 1 Then Exit Sub
        iB1_Atai = .GridObj.TextMatrix(.GridObj.Row, .FixedGet(1))
        iB1_AtaiNm = .GridObj.TextMatrix(.GridObj.Row, .FixedGet(2))
        iB1_ESu1 = .GridObj.TextMatrix(.GridObj.Row, .FixedGet(3))
        iB1_ESu2 = .GridObj.TextMatrix(.GridObj.Row, .FixedGet(4))
        iB1_ESu3 = .GridObj.TextMatrix(.GridObj.Row, .FixedGet(5))
        iB1_Su1 = CisFun.RSetFld(.GridObj.TextMatrix(.GridObj.Row, .FixedGet(6)), 0, "#########0.0000")
        iB1_Su2 = CisFun.RSetFld(.GridObj.TextMatrix(.GridObj.Row, .FixedGet(7)), 0, "#########0.0000")
        iB1_Su3 = CisFun.RSetFld(.GridObj.TextMatrix(.GridObj.Row, .FixedGet(8)), 0, "#########0.0000")
    End With
'
End Sub
'+-------------------------------------+
'+      ｸﾞﾘｯﾄﾞより内容セット後入力処理   +
'+-------------------------------------+
Private Sub FromGrid()
    If SyoriKBB <> "MNT" And SyoriKBB <> "DEL" And Not mLock Then Exit Sub
    Call BodySet2
    If Trim(VsGrid.TextMatrix(VsGrid.Row, mVSObj.FixedGet(2))) = "" And _
       Trim(VsGrid.TextMatrix(VsGrid.Row, mVSObj.FixedGet(1))) = "" Then
       VsGrid.SetFocus
       Exit Sub
    End If
    If SyoriKBB = "MNT" Then
       ProcHB = "B1"
       Call DispChange(ProcHB)
       Call mVSObj.vsColor(消去個別)
       Call mVSObj.vsColor
       Call SyoriNameB(SyoriKBB)
       iB1_AtaiNm.SetFocus
       Exit Sub
    End If
    If SyoriKBB = "DEL" Then
       Call mVSObj.vsColor
       Call DBDeleteB
       Call mVSObj.vsColor(消去個別)
               
       VsGrid.SetFocus
    End If
End Sub
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
         .MB_Lines = 5
         .MB_MSG(2) = "   区分削除を行います。           "
         .MB_MSG(4) = "     よろしいですか？           "
         .MB_Title = "削除確認"
         .MB_Button = OK_CAN
         If Not .MBOX Then Exit Function
    End With
' ( 削除処理 )
    With CisDB
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
        .SQL = "DELETE FROM 名称マスタ"
        .SQL = .SQL & " WHERE 区分名称 = '" & RTrim(iH1_Kubun) & "'"
        .SQL = .SQL & "   AND  制御区分 = '" & mSeigyo & "'"
        On Error GoTo DBDelete_Err
    
        .DBExec
        
'        .SQL = "DELETE FROM 名称マスタ"
'        .SQL = .SQL & " WHERE 区分名称 = '名称検索'"
'        .SQL = .SQL & "   AND 値名称 = '" & RTrim(iH1_Kubun) & "'"
'        .SQL = .SQL & "   AND 制御区分 = 'X'"
'        .DBExec
        On Error GoTo 0
        Call .DBTran(TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    End With
    
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+------------------------------+
'+      削 除 処 理(値削除)      +
'+------------------------------+
Private Function DBDeleteB()
    DBDeleteB = False
'  ( 確認メッセージ )
    K_SykbnmB.Caption = SyoriNMB
    K_SykbnmB.Visible = True
    With CisFun
         .MB_Lines = 5
         .MB_MSG(2) = "    値削除を行います。           "
         .MB_MSG(4) = "     よろしいですか？           "
         .MB_Title = "削除確認"
         .MB_Button = OK_CAN
         If Not .MBOX Then
            VsGrid.SetFocus
            Exit Function
         End If
    End With
    
' ( 削除処理 )
    With CisDB
        .DBTran (TransBegin)
        .SQL = "DELETE FROM 名称マスタ"
        .SQL = .SQL & " WHERE 区分名称 = '" & RTrim(iH1_Kubun) & "'"
        .SQL = .SQL & "   AND 区分タイプ   = ''"
        .SQL = .SQL & "   AND  制御区分 = '" & mSeigyo & "'"
        .SQL = .SQL & "   AND 値 = '" & RTrim(iB1_Atai) & "'"
        On Error GoTo DBDeleteB_Err
        .DBExec
        On Error GoTo 0
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    End With
    Call HeadBodyClear("B")
    Call GridSet
    ProcHB = "B2"
    Call DispChange(ProcHB)
    SyoriKBB = "MNT"
    Call SyoriNameB(SyoriKBB)
    VsGrid.SetFocus
    
    DBDeleteB = True
    Exit Function
DBDeleteB_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    Dim UpdFlg  As Boolean
    DBPut = False
    UpdFlg = False
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    K_SykbnmB.Caption = SyoriNMB
    K_SykbnmB.Visible = True
    With CisFun
         .MB_Lines = 5
         If SyoriKBB = "ADD" Then
            .MB_MSG(2) = "    追加処理を行います。           "
            .MB_Title = "追加確認"
         Else
            .MB_MSG(2) = "    修正処理を行います。           "
            .MB_Title = "修正確認"
         End If
         .MB_MSG(4) = "     よろしいですか？           "
         .MB_Button = OK_CAN
         If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err
    
    With CisDB
    
         .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    '   マスタ存在確認(索引レコード）
        gSL_Select = "SELECT * FROM 名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = '" & RTrim(iH1_Kubun) & "'"
        gSL_Select = gSL_Select & "   AND 区分タイプ   = 'K'"
        gSL_Select = gSL_Select & "   AND  制御区分 = '" & mSeigyo & "'"
        If Not SYMRead(gSL_Select, 1) Then
           Call ItemsClearSYM
           With SYM
                .区分名称 = RTrim(iH1_Kubun)
                .区分タイプ = "K"
                .区分桁数 = CisFun.Val2(iH2_Keta)
                If mSeigyo <> "A" Then
                   .制御区分 = mSeigyo
                End If
           End With
           Call SYMInsert
        End If
        
        gSL_Select = "SELECT * FROM 名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = '" & RTrim(iH1_Kubun) & "'"
        gSL_Select = gSL_Select & "   AND 区分タイプ   = ''"
        gSL_Select = gSL_Select & "   AND 値   = '" & RTrim(iB1_Atai) & "'"
        gSL_Select = gSL_Select & "   AND  制御区分 = '" & mSeigyo & "'"
        If Not SYMRead(gSL_Select, 1) Then
           Call ItemsClearSYM
           UpdFlg = True
           SYM.区分名称 = RTrim(iH1_Kubun)
           SYM.値 = RTrim(iB1_Atai)
           SYM.区分桁数 = CisFun.Val2(iH2_Keta)
           SYM.制御区分 = mSeigyo
           SYM.作成者 = gTanto
        End If
    '   マスタ項目移送
        With SYM
            .値名称 = RTrim(iB1_AtaiNm)
            .英数字1 = RTrim(iB1_ESu1)
            .英数字2 = RTrim(iB1_ESu2)
            .英数字3 = RTrim(iB1_ESu3)
            .数字1 = CisFun.Val2(iB1_Su1)
            .数字2 = CisFun.Val2(iB1_Su2)
            .数字3 = CisFun.Val2(iB1_Su3)
            .更新者 = gTanto
            .制御区分 = mSeigyo
        End With
    ' ( 追加処理 )
        If UpdFlg Then
            Call SYMInsert
        Else
    ' ( 更新処理 )
            Call SYMUpdate
        End If
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    End With
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit(Keta As Integer)
    With mVSObj
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容)
         Call .FixedSet(中中, 中中, なし, 6, 0, "登録№")
         Call .FixedSet(中中, 左中, なし, 7, 1, "部 ")       ' 部署コード１
         Call .FixedSet(中中, 左中, なし, 30, 2, "１")       ' 部署名１
         Call .FixedSet(中中, 左中, なし, 3, 3, "返 ")       ' 返却場所コード１
         Call .FixedSet(中中, 左中, なし, 30, 4, "１")       ' 返却場所名１
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
         Call .FixedSet(中中, 左中, なし, 0, 5, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4")
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
    
    Call GridInit(CisFun.Val2(iH2_Keta))                        ' ｸﾞﾘｯﾄﾞ初期表示

' 抽出処理 初期ｾｯﾄ
    gSL_Select = "SELECT * FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = '" & RTrim(iH1_Kubun) & "'"
    gSL_Select = gSL_Select & "   AND 区分タイプ   = ''"
    gSL_Select = gSL_Select & "   AND  制御区分 = '" & mSeigyo & "'"
    gSL_Select = gSL_Select & " ORDER BY 値"
    If Not SYMRead(gSL_Select) Then Exit Sub
    
    Call CisFrm.MousePT(11)  ' ﾏｳｽﾎﾟｲﾝﾀ (砂時計)
'
    PNL_PF.Enabled = False      ' ﾌｧﾝｸｼｮﾝ使用禁止
    

    B2lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount      ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid
        .Redraw = True
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = mVSObj.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < mVSObj.DispMax Then
            .Rows = mVSObj.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 1
    Do Until Not SYM_RDSTS
        With VsGrid
            .TextMatrix(ID, GridCol(0)) = Format(ID, "###")
            .TextMatrix(ID, GridCol(1)) = SYM.値
            .TextMatrix(ID, GridCol(2)) = SYM.値名称
            .TextMatrix(ID, GridCol(3)) = SYM.英数字1
            .TextMatrix(ID, GridCol(4)) = SYM.英数字2
            .TextMatrix(ID, GridCol(5)) = SYM.英数字3
            .TextMatrix(ID, GridCol(6)) = Format(SYM.数字1, "#########0.0000")
            .TextMatrix(ID, GridCol(7)) = Format(SYM.数字2, "#########0.0000")
            .TextMatrix(ID, GridCol(8)) = Format(SYM.数字3, "#########0.0000")
            .TextMatrix(ID, GridCol(9)) = Format(SYM.作成日, "YYYY/MM/DD")
            If SYM.更新日 <> 0 Then
                .TextMatrix(ID, GridCol(10)) = Format(SYM.更新日, "YYYY/MM/DD")
            End If
            ID = ID + 1
            Call SYMReadNext
        End With
    Loop
    With VsGrid
        .Row = 1: .Col = 1
        .Redraw = True
    End With

    PNL_PF.Enabled = True      ' ﾌｧﾝｸｼｮﾝ使用禁止 解除
    Call CisFrm.MousePT(11)        ' ﾏｳｽﾎﾟｲﾝﾀ (砂時計)解除
'
End Sub

'******************************************'
'*****        行選択（クリック）       *****
'******************************************'
Private Sub VSGrid_Click()
    If mLock Then
       Call FromGrid
       Exit Sub
    End If
    If ProcHB <> "B2" Or SyoriKBB = "ADD" Or SyoriKBB = "" Then
       SendKeys "{TAB}"
       Exit Sub
    End If
    Call FromGrid
End Sub
'******************************************'
'*****        行選択（スペースｷｰ)      *****
'******************************************'
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If mLock Then
       Call FromGrid
       Exit Sub
    End If
    
    If ProcHB <> "B2" Or SyoriKBB = "ADD" Or SyoriKBB = "" Then
       SendKeys "{TAB}"
       Exit Sub
    End If
    If KeyCode <> vbKeySpace Then Exit Sub
    Call FromGrid
End Sub
'+-------------------------------+
'+      行変更時　内容表示        +
'+-------------------------------+
Private Sub VsGrid_RowColChange()
    If ProcHB <> "B2" Or SyoriKBB = "ADD" Or SyoriKBB = "" Then
       Exit Sub
    End If
    Call BodySet2
End Sub
'******************************************'
'*****       移動選択項目選択          *****
'******************************************'
Private Sub VSGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Button : 1 = 左ﾎﾞﾀﾝ ,2 = 右ﾎﾞﾀﾝ
    With VsGrid
        .Tag = ""
        If Button = 2 Then
            .Row = .MouseRow
             If .MouseRow = 0 Then
                Exit Sub
            End If
            If .MouseRow = 0 Then
                .Col = 1: .ColSel = .Cols - 1
                Exit Sub
            End If
            .Col = 1: .ColSel = .Cols - 1
            Exit Sub
        End If
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
Private Sub VSGrid_DragDrop(Source As Control, X As Single, Y As Single)
    With VsGrid
    ' 固定列は行わない
        If .MouseCol = 0 Then Exit Sub
        If .Tag = "" Then Exit Sub
' ｸﾞﾘｯﾄﾞ表示 禁止
        .Redraw = False
' 列移動
        .ColPosition(CisFun.Val2(.Tag)) = .MouseCol
        .Col = 1: .ColSel = .Cols - 1
' ｸﾞﾘｯﾄﾞ表示 解除
        .Redraw = True
'
    End With
End Sub

