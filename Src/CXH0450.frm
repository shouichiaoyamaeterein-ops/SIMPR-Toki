VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{8F4FF474-3E18-11D4-AEC2-00000EA57F9E}#1.0#0"; "CISMSG~1.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CXH0450 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "発注打切問合せ"
   ClientHeight    =   10980
   ClientLeft      =   75
   ClientTop       =   1470
   ClientWidth     =   15315
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10980
   ScaleWidth      =   15315
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D B2_Area1 
      Height          =   5730
      Left            =   6060
      Top             =   4440
      Visible         =   0   'False
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   10107
      BackColor       =   12632319
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      Begin CisMsg_v60.CISMSG CISMSG3 
         Height          =   5610
         Left            =   60
         TabIndex        =   25
         Top             =   60
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   9895
         BackColor       =   -2147483633
         ForeColor       =   65280
         Caption         =   "※　発　注　打　切　明　細　※"
         BackStyle       =   0
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
         cFont3DColor1   =   16777152
         cFont3DColor2   =   0
         cAlingnment     =   6
         cBoderStyle     =   1
         cPositionX      =   30
         cPositionY      =   940
         gGC2_Green      =   160
         gGradation      =   -1  'True
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   13
         Left            =   435
         Top             =   720
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "打切種別"
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
         cAlingnment     =   7
         cBoderStyle     =   1
         cPositionX      =   30
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   0
         Left            =   435
         Top             =   60
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "品 番"
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
         cPositionX      =   30
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   13
         Left            =   1605
         Top             =   720
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XXXXXXXXXX"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cBoderStyle     =   1
         cPositionX      =   75
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   1
         Left            =   4575
         Top             =   60
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "仕入先"
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
         cPositionX      =   30
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   1
         Left            =   4575
         Top             =   390
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XXXXXXX:XXXXXXXXX1XXXXXXXXX2"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cBoderStyle     =   1
         cPositionX      =   45
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   0
         Left            =   435
         Top             =   390
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cBoderStyle     =   1
         cPositionX      =   45
         cPositionY      =   15
      End
      Begin vsFlexLib.vsFlexArray vsGrid1 
         Height          =   2790
         Left            =   435
         TabIndex        =   26
         Top             =   2880
         Width           =   8055
         _Version        =   196608
         _ExtentX        =   14208
         _ExtentY        =   4921
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
         FormatString    =   "№ |品　番　　　　　　　     |仕入先    　　　|受入|背番号  |構成|||"
         Rows            =   9
         Cols            =   9
         BackColor       =   12648384
         BackColorBkg    =   14737632
         FocusRect       =   0
         HighLight       =   0
         RowHeightMin    =   300
         GridLines       =   3
         SelectionMode   =   1
      End
      Begin Cis3D_v60.CIS3D UC_3D12 
         Height          =   480
         Index           =   0
         Left            =   435
         Top             =   2400
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   847
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "※　対象品番 ※"
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
         cAlingnment     =   7
         cPositionX      =   30
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   2
         Left            =   5475
         Top             =   1050
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "起算日－便"
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
         cAlingnment     =   7
         cBoderStyle     =   1
         cPositionX      =   30
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   2
         Left            =   6825
         Top             =   1050
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "YYYY/MM/DD-00"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   7
         cBoderStyle     =   1
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   3
         Left            =   435
         Top             =   1380
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "設定数"
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
         cAlingnment     =   7
         cBoderStyle     =   1
         cPositionX      =   30
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   3
         Left            =   1605
         Top             =   1380
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "ZZZ,ZZZ"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   4
         cBoderStyle     =   1
         cPositionX      =   -45
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   4
         Left            =   435
         Top             =   1710
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "警告数"
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
         cAlingnment     =   7
         cBoderStyle     =   1
         cPositionX      =   30
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   4
         Left            =   1605
         Top             =   1710
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "ZZZ,ZZZ"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   4
         cBoderStyle     =   1
         cPositionX      =   -45
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   5
         Left            =   435
         Top             =   1050
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "発注数"
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
         cAlingnment     =   7
         cBoderStyle     =   1
         cPositionX      =   30
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   5
         Left            =   1605
         Top             =   1050
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "ZZZ,ZZZ"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   4
         cBoderStyle     =   1
         cPositionX      =   -45
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   6
         Left            =   2955
         Top             =   1050
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "入荷数"
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
         cAlingnment     =   7
         cBoderStyle     =   1
         cPositionX      =   30
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   6
         Left            =   4125
         Top             =   1050
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "ZZZ,ZZZ"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   4
         cBoderStyle     =   1
         cPositionX      =   -45
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   7
         Left            =   5475
         Top             =   1380
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "設定残数"
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
         cAlingnment     =   7
         cBoderStyle     =   1
         cPositionX      =   30
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   8
         Left            =   5475
         Top             =   1710
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "完了日－便"
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
         cAlingnment     =   7
         cBoderStyle     =   1
         cPositionX      =   30
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   7
         Left            =   6825
         Top             =   1380
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "ZZZ,ZZZ"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   4
         cBoderStyle     =   1
         cPositionX      =   -45
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   8
         Left            =   6825
         Top             =   1710
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "YYYY/MM/DD-00"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cBoderStyle     =   1
         cPositionX      =   45
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   9
         Left            =   2955
         Top             =   1380
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "打切日"
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
         cAlingnment     =   7
         cBoderStyle     =   1
         cPositionX      =   30
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   9
         Left            =   4125
         Top             =   1380
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "YYYY/MM/DD"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cBoderStyle     =   1
         cPositionX      =   45
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   10
         Left            =   2955
         Top             =   1710
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "警告日"
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
         cAlingnment     =   7
         cBoderStyle     =   1
         cPositionX      =   30
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   10
         Left            =   4125
         Top             =   1710
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "YYYY/MM/DD"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cBoderStyle     =   1
         cPositionX      =   45
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   11
         Left            =   435
         Top             =   2040
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "切替仕入先"
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
         cAlingnment     =   7
         cBoderStyle     =   1
         cPositionX      =   30
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   11
         Left            =   2235
         Top             =   2040
         Width           =   4590
         _ExtentX        =   8096
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XXXXXXX:XXXXXXXXX1XXXXXXXXX2"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cBoderStyle     =   1
         cPositionX      =   60
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   12
         Left            =   6825
         Top             =   2040
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "切替受入"
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
         cAlingnment     =   7
         cBoderStyle     =   1
         cPositionX      =   30
         cPositionY      =   15
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   12
         Left            =   7995
         Top             =   2040
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "00"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   7
         cBoderStyle     =   1
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   330
         Index           =   14
         Left            =   2955
         Top             =   720
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         ForeColor       =   16711680
         Caption         =   "品　名"
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
         cAlingnment     =   7
         cBoderStyle     =   1
         cPositionX      =   30
      End
      Begin Cis3D_v60.CIS3D B2lb_Mei 
         Height          =   330
         Index           =   14
         Left            =   4125
         Top             =   720
         Width           =   4350
         _ExtentX        =   7673
         _ExtentY        =   582
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cBoderStyle     =   1
         cPositionX      =   75
         cPositionY      =   15
      End
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   7920
      Left            =   270
      ScaleHeight     =   7860
      ScaleWidth      =   14655
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2685
      Width           =   14715
      Begin Cis3D_v60.CIS3D PNL_Wait 
         Height          =   930
         Left            =   5160
         Top             =   1200
         Visible         =   0   'False
         Width           =   5160
         _ExtentX        =   9102
         _ExtentY        =   1640
         BackColor       =   16761024
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
         cBoderWidth     =   5
         Begin Cis3D_v60.CIS3D Message 
            Height          =   525
            Left            =   240
            Top             =   210
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   926
            BackColor       =   16761024
            ForeColor       =   0
            Caption         =   "【抽出中】しばらくお待ち下さい"
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
            cBoderWidth     =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CXH0450.frx":0000
         Height          =   7710
         Left            =   60
         OleObjectBlob   =   "CXH0450.frx":0015
         TabIndex        =   20
         Top             =   60
         Width           =   14535
      End
   End
   Begin VB.PictureBox H1_Area1 
      Height          =   1095
      Left            =   270
      ScaleHeight     =   1035
      ScaleWidth      =   14655
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1260
      Width           =   14715
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   930
         Index           =   4
         Left            =   8970
         Top             =   60
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1640
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
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin VB.OptionButton H1Op_DspKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "完了"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1110
            TabIndex        =   6
            Top             =   330
            Width           =   855
         End
         Begin VB.OptionButton H1Op_DspKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "未完了"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   1065
         End
         Begin VB.OptionButton H1Op_DspKbn 
            BackColor       =   &H00C0FFC0&
            Caption         =   "全て"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   330
            Width           =   855
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   930
         Index           =   2
         Left            =   1860
         Top             =   60
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   1640
         ForeColor       =   16711680
         Caption         =   "品　　　番"
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
         cPositionX      =   50
         cPositionY      =   75
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   360
            Left            =   120
            TabIndex        =   2
            Top             =   420
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   635
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   930
         Index           =   3
         Left            =   5250
         Top             =   60
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   1640
         ForeColor       =   16711680
         Caption         =   "仕　入　先"
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
         cPositionX      =   50
         cPositionY      =   75
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   120
            TabIndex        =   3
            Top             =   420
            Width           =   990
            _ExtentX        =   1746
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   345
            Left            =   1110
            Top             =   420
            Width           =   2460
            _ExtentX        =   4339
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
      Begin Cis3D_v60.CIS3D H1_DateTitle 
         Height          =   930
         Left            =   11010
         Top             =   60
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   1640
         ForeColor       =   16711680
         Caption         =   "起算日"
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cPositionX      =   75
         cPositionY      =   75
         Begin Cis3D_v60.CIS3D PB_SDate 
            Height          =   345
            Left            =   810
            Top             =   105
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   609
            Caption         =   "開始"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_EDate 
            Height          =   345
            Left            =   810
            Top             =   465
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   609
            Caption         =   "終了"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin CisText_V60.CisText iH1_SBin 
            Height          =   360
            Left            =   3120
            TabIndex        =   8
            Top             =   105
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            cFormat         =   "00"
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
            Text            =   "00"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin CisText_V60.CisText iH1_EBin 
            Height          =   360
            Left            =   3120
            TabIndex        =   10
            Top             =   465
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            cFormat         =   "00"
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
            Text            =   "00"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_SDate 
            Height          =   345
            Left            =   1290
            TabIndex        =   7
            Top             =   105
            Width           =   1830
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
            Object.Height          =   345
            Object.Width           =   1830
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_EDate 
            Height          =   345
            Left            =   1290
            TabIndex        =   9
            Top             =   465
            Width           =   1830
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
            Object.Height          =   345
            Object.Width           =   1830
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   930
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   1640
         ForeColor       =   16711680
         Caption         =   "打切種別"
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
         cPositionY      =   75
         Begin VB.ComboBox H1Cmb_UKbn 
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
            ItemData        =   "CXH0450.frx":1722D
            Left            =   120
            List            =   "CXH0450.frx":1722F
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   1
            Top             =   420
            Width           =   1560
         End
      End
   End
   Begin VB.PictureBox Back_Dsp 
      BackColor       =   &H00404040&
      Height          =   720
      Left            =   270
      ScaleHeight     =   660
      ScaleWidth      =   5295
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   450
      Width           =   5355
      Begin Cis3D_v60.CIS3D UC_3D5 
         Height          =   645
         Index           =   0
         Left            =   15
         Top             =   0
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   1138
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D UC_3D5 
            Height          =   570
            Index           =   1
            Left            =   60
            Top             =   30
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   1005
            BackColor       =   16761024
            Caption         =   "表示切替"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
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
         End
         Begin Cis3D_v60.CIS3D UC_3D5 
            Height          =   570
            Index           =   2
            Left            =   735
            Top             =   30
            Width           =   3510
            _ExtentX        =   6191
            _ExtentY        =   1005
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderStyle     =   1
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "⑤"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   4
               Left            =   2730
               TabIndex        =   17
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "④"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   3
               Left            =   2070
               TabIndex        =   16
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "③"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   2
               Left            =   1395
               TabIndex        =   15
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "②"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   1
               Left            =   720
               TabIndex        =   14
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H008080FF&
               Caption         =   "①"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   0
               Left            =   60
               TabIndex        =   13
               Top             =   150
               Value           =   -1  'True
               Width           =   615
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D5 
            Height          =   570
            Index           =   3
            Left            =   4260
            Top             =   30
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   1005
            ForeColor       =   16711935
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
            Begin VB.OptionButton Op_Sort 
               BackColor       =   &H00C0C0C0&
               Caption         =   "昇順"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   0
               Left            =   30
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   60
               Value           =   -1  'True
               Width           =   825
            End
            Begin VB.OptionButton Op_Sort 
               BackColor       =   &H00C0C0C0&
               Caption         =   "降順"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   1
               Left            =   30
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   315
               Width           =   825
            End
         End
      End
   End
   Begin VB.TextBox Dummy 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   345
      Left            =   240
      TabIndex        =   0
      Text            =   "D"
      Top             =   210
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   510
      Top             =   210
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   1230
      Top             =   210
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   705
      Left            =   8130
      ScaleHeight     =   645
      ScaleWidth      =   6795
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   450
      Width           =   6855
      Begin CisBtn_60.CisBtn PB_SET 
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   15
         BtnCaption      =   "設定"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   0
         BtnShadowColor  =   16777215
         PFCaption       =   "( F11 )"
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
         PFBackColor     =   0
      End
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   5925
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
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
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   5085
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   4245
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
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
      Begin CisBtn_60.CisBtn PB_Csv 
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
            Weight          =   700
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
      Begin CisBtn_60.CisBtn PB_Prt 
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
            Weight          =   700
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   885
         Top             =   30
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
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
      Begin CisBtn_60.CisBtn PB_MNT 
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
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
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
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Cis3D_v60.CIS3D UC_3D31 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "　　　【 発注打切問合せ 】"
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
      Begin VB.PictureBox Back_Over 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   4665
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   45
         Width           =   1995
         Begin VB.OptionButton Op_Overlap 
            BackColor       =   &H00C0C0C0&
            Caption         =   "重複"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   780
         End
         Begin VB.OptionButton Op_Overlap 
            BackColor       =   &H00C0C0C0&
            Caption         =   "単一"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   1020
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Value           =   -1  'True
            Width           =   780
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D32 
         Height          =   255
         Left            =   7500
         Top             =   45
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14415
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
         Left            =   11925
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
   End
End
Attribute VB_Name = "CXH0450"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   発注打切問合せ
'**       フォームID    :   CXH0450
'**       処理概要      :
'**
'**       作  成  日    :   2004/02/07  By CIS
'**       変  更  日    :   2008/04/11  By CIS 背番号８桁対応
'**       変  更  日    :   2014/02/04  By CIS WIN7対応
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean          ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2       ' 入力域切替
    Dim CisDBGrid1              As New CisDBGrid    ' TDBGridPro60ｸﾗｽ定義
    Dim CisVsGrid1              As New CisVsGrid3
'   背景色
    Dim mGotColor               As Long
    Dim mLostColor              As Long
'   退避
    Dim SaveRow                 As Long
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
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止

    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If

'+---------------------+
'+ TrueGrid60 初期設定  +
'+---------------------+
    With CisDBGrid1
        Set .DBGridObj = DBGrid1
        .FrmCaption = Me.Caption
        .MaxSort = 7                        ' ﾚｲｱｳﾄｿｰﾄ最大数
        .LayOut = True                      ' ﾚｲｱｳﾄ変更処理 True:有り,False:無し
        .LayOutMax = 5                      ' ﾚｲｱｳﾄ情報最大数
        .LayOutColMax = 30                  ' ﾚｲｱｳﾄ表示可能項目最大数
        .LayOutPath = gIniExe               ' ﾚｲｱｳﾄ情報格納先、ﾌｧｲﾙ場所
        .LayOutName = CisFun.WSID & "H0450.Lay"       ' ﾚｲｱｳﾄ情報格納先、ﾌｧｲﾙ名
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        .CsvName = Trim(Me.Caption) & ".csv"
        .CsvRegKey = gRegKey
        .CsvValue = "CXH0450"
        .GridInitGet = False                 ' ｸﾞﾘｯﾄﾞ内容初期設定 ｸﾘｯﾌﾟﾎﾞｰﾄﾞ作成ﾌﾗｸﾞ
        .GridProcGet = False
        .Init
    End With
    Call GridInitSet1

    '+---------------------+
    '+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
    '+---------------------+
'   打切種別
    With H1Cmb_UKbn
        .AddItem "全て"
        .AddItem "残数"
        .AddItem "品番"
        .AddItem "品番仕入先"
        .AddItem "仕入先"
    End With

    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 300
        .PatanMax = 1
        .SelectCol = 6
        .InitGet = False
        .Init
    End With
    Call GridInit1

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    mGotColor = H1Op_DspKbn(0).BackColor
    mLostColor = H1Op_DspKbn(1).BackColor

    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

' Head部初期化
    B1Op_Patn(0).Value = True
    Op_Sort(0).Value = True:    Op_Sort(0).BackColor = mGotColor
    Op_Overlap(0).Value = True: Op_Overlap(0).BackColor = mGotColor

    Call HeadBodyClear("H")

' ADOｺﾝﾄﾛｰﾙ　接続情報　セット
    Adodc1.ConnectionString = CisDB.ConnectString
    '-----2014/02/04 EDIT START-----
    gStr = CisDB.ConnectString
    gConnect = "server=" & CisDB.Server & ";database=" & CisDB.DBName & ";User ID=" & CisDB.User & ";password=" & CisDB.PassWord(GetSetting(gRegKey, "Environment", "DrowsSap", ""))
    
    gStr = Replace(gStr, "Use Procedure", "Extended Properties=" & gConnect & ";Use Procedure")
    
    Adodc1.ConnectionString = gStr
    '-----2014/02/04 EDIT END-------

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
           Case vbKeyF2:        If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF8:        If PB_PRT.Visible Then Call PB_PRT_Click    '【印刷】
           Case vbKeyF10:       If PB_Csv.Visible Then Call PB_Csv_Click    '【出力】
           Case vbKeyF11:       If PB_SET.Visible Then Call PB_SET_Click    '【設定】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      '【入力】
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
'+---------------------+
'+   ﾃﾞｰﾀﾍﾞｰｽ切断
'+---------------------+
'   #### ﾃﾞｰﾀﾍﾞｰｽ切断 ####
    Call CisDB.DBDISConnect
'   #### ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去 ####
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
    If ProcHB = "H1" Then
        Call HeadBodyClear("H")
        H1Cmb_UKbn.SetFocus
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call DBGridClr1
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B2" Then
        ProcHB = "B1"
        Call DispChange(ProcHB)
        DBGrid1.SetFocus
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()

    RV_Call = "CHX0450"
    With HUT
        .打切種別 = CisFun.Val2(DBGrid1.Columns(13))
        .品番 = RTrim(DBGrid1.Columns(15))
        .仕入先 = RTrim(DBGrid1.Columns(16))
    End With

    CXH0050.Show vbModal, Me
    Unload CXH0050
    Set CXH0050 = Nothing

    SaveRow = DBGrid1.Row
    Call DBInput
    DBGrid1.Row = SaveRow

End Sub
'+----------------------------+
'+    CSV出力           　     +
'+----------------------------+
Private Sub PB_Csv_Click()
    With CisDBGrid1
        Set .Recordset = Adodc1.Recordset
        .PutCsv
    End With
End Sub
'+----------------------------+
'+    印刷             　     +
'+----------------------------+
Private Sub PB_PRT_Click()
    With CisDBGrid1
        Set .Recordset = Adodc1.Recordset
        .PutPrt
    End With
End Sub
'+----------------------------+
'+    設定              　     +
'+----------------------------+
Private Sub PB_SET_Click()
    With CisDBGrid1
        .LayOutPatan
        If .LayOutSetFlg Then
           B1Op_Patn(.LayOutSetNo).Value = True
           If ProcHB = "B1" Then
              Call DBInput("R")
           End If
        End If
    End With
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_SDate" Then iH1_SDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_EDate" Then iH1_EDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
End Sub
'+-----------------------------+
'+       取引先検索             +
'+-----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 8

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+-----------------------------+
'+       品番検索              +
'+-----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0

    RVI_Hinmoku = 1         ' 発注
    RVI_SK = 9
    RV_Call = ""

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Hinbn = RV_Hinbn
    End If
End Sub
'+-----------------------------+
'+  H1Cmb_UKbn
'+-----------------------------+
Private Sub H1Cmb_UKbn_GotFocus()
    H1Cmb_UKbn.BackColor = iH1_SDate.cGotColor
End Sub
Private Sub H1Cmb_UKbn_LostFocus()
    H1Cmb_UKbn.BackColor = iH1_SDate.cLostColor
End Sub
Private Sub H1Cmb_UKbn_Click()
    Select Case H1Cmb_UKbn.ListIndex
        Case 0
            If H1Op_DspKbn(2).Value Then
                H1_DateTitle.Caption = "完了日" & vbCr & "　・" & vbCr & "打切日"
            Else
                H1_DateTitle.Caption = "起算日" & vbCr & "　・" & vbCr & "打切日"
            End If
        Case 1
            If H1Op_DspKbn(2).Value Then
                H1_DateTitle.Caption = vbCr & "完了日"
            Else
                H1_DateTitle.Caption = vbCr & "起算日"
            End If
        Case 2 To 4
            H1_DateTitle.Caption = vbCr & "打切日"
    End Select
    If H1_DateTitle.Caption = vbCr & "打切日" Then
        iH1_SBin.Enabled = False
        iH1_EBin.Enabled = False
        iH1_SBin = ""
        iH1_EBin = ""
    Else
        iH1_SBin.Enabled = True
        iH1_EBin.Enabled = True
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
    If TorNmGet(iH1_Torcd, 8) Then
        H1lb_Tornm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iH1_SDate
'+-----------------------------+
Private Sub iH1_SDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_SDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_EDate
'+-----------------------------+
Private Sub iH1_EDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_EDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+----------------------------------+
'+      H1Op_DspKbn
'+----------------------------------+
Private Sub H1Op_DspKbn_GotFocus(Index As Integer)
    For Each gObj In H1Op_DspKbn
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
Private Sub H1Op_DspKbn_Click(Index As Integer)
    Select Case H1Cmb_UKbn.ListIndex
        Case 0
            Select Case Index
                Case 2:     H1_DateTitle.Caption = "完了日" & vbCr & "　・" & vbCr & "打切日"
                Case Else:  H1_DateTitle.Caption = "起算日" & vbCr & "　・" & vbCr & "打切日"
            End Select
        Case 1
            Select Case Index
                Case 2:     H1_DateTitle.Caption = vbCr & "完了日"
                Case Else:  H1_DateTitle.Caption = vbCr & "起算日"
            End Select
    End Select
    If H1_DateTitle.Caption = vbCr & "打切日" Then
        iH1_SBin.Enabled = False
        iH1_EBin.Enabled = False
        iH1_SBin = ""
        iH1_EBin = ""
    Else
        iH1_SBin.Enabled = True
        iH1_EBin.Enabled = True
    End If
End Sub
'+----------------------------+
'+    表示切替ｵﾌﾟｼｮﾝﾎﾞﾀﾝ　     +
'+----------------------------+
Private Sub B1Op_Patn_Click(Index As Integer)
    Call CisFrm.MousePT(11)

    For Each gObj In B1Op_Patn
        If Index = gObj.Index Then
           gObj.BackColor = &H8080FF
        Else
           gObj.BackColor = &HC0C0C0
        End If
    Next gObj

    CisDBGrid1.LayOutNow = Index + 1
    CisDBGrid1.LayOutDisp

    If ProcHB = "B1" Then
        Call DBInput("R")
    End If

    Call CisFrm.MousePT(1)
End Sub
'+---------------------------+
'+    重複表示　ﾁｪｯｸﾎﾞｯｸｽ     +
'+---------------------------+
Private Sub Op_Overlap_GotFocus(Index As Integer)
    For Each gObj In Op_Overlap
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
Private Sub Op_Overlap_Click(Index As Integer)
    If Not FormAct Then Exit Sub

    Call CisFrm.MousePT(11)

    Call CisDBGrid1.MargeDisp(Op_Overlap(1).Value)

    If ProcHB = "H1" Then
        H1Cmb_UKbn.SetFocus
    Else
        DBGrid1.SetFocus
    End If

    Call CisFrm.MousePT(1)
End Sub
'+---------------------------+
'+    分類表示　ﾁｪｯｸﾎﾞｯｸｽ     +
'+---------------------------+
Private Sub Op_Sort_GotFocus(Index As Integer)
    For Each gObj In Op_Sort
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
Private Sub Op_Sort_Click(Index As Integer)
    If ProcHB = "B1" Or _
       ProcHB = "B2" Then
        Call DBInput
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
    Dim B2Mode      As Boolean
    Dim B2Color     As Long

    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    B2Mode = False:     B2Color = gPLostSel

'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus

    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel: PB_ENT.BtnCaption = "入力"
        Case "B1":      B1Mode = True:  B1Color = gPGotSel: PB_ENT.BtnCaption = "詳細"
        Case "B2":      B2Mode = True:  B2Color = gPGotSel
    End Select

    H1_Area1.Enabled = H1Mode:  H1_Area1.BackColor = H1Color
   'B1_Area1.Enabled = B1Mode:  B1_Area1.BackColor = B1Color
   'B2_Area1.Enabled = B2Mode:  B2_Area1.BackColor = B2Color
    Select Case Pro
        Case "B2":  B1_Area1.Enabled = B2Mode:  B1_Area1.BackColor = B2Color
        Case Else:  B1_Area1.Enabled = B1Mode:  B1_Area1.BackColor = B1Color
    End Select

    PB_Look.Visible = H1Mode
    PB_END.Visible = H1Mode
    B2_Area1.Visible = B2Mode
    PB_ENT.Visible = H1Mode Or B1Mode
    PB_SET.Visible = H1Mode
    PB_PRT.Visible = B1Mode
    PB_Csv.Visible = B1Mode
    PB_MNT.Visible = B1Mode

'   ﾀﾞﾐｰ項目(入力禁止)
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
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = "": .Enabled = True
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False: .BackColor = mLostColor
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Cmb_*" Then .ListIndex = 0
        End With
    Next gObj
    If HeadBody = "H" Then
        H1Op_DspKbn(0).Value = True
        H1Op_DspKbn(0).BackColor = mGotColor
        H1_DateTitle.Caption = "起算日" & vbCr & "　・" & vbCr & "打切日"
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
        DBGrid1.SetFocus
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        Call BodySet
        ProcHB = "B2"
        Call DispChange(ProcHB)
        DBGrid1.SetFocus
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

'   日付
    If Trim(iH1_SDate) <> "" Then
        If CisFun.Val2(iH1_SDate.cDay) = 0 Then
            iH1_EDate = ""
            iH1_EBin = ""
        End If
    End If
    If Trim(iH1_SDate) <> "" And Trim(iH1_EDate) <> "" Then
        If iH1_SDate & Format(iH1_SBin, "00") > iH1_EDate & Format(iH1_EBin, "00") Then
            With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    " & H1_DateTitle.Caption & "を正しく入力して下さい（開始＞終了）            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
            End With
            iH1_SDate.SetFocus
            Exit Function
        End If
    End If

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
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示     *
'*                                    *
'**************************************
Private Function DBInput(Optional InType As String = "") As Boolean
    Dim DateStr     As String
    Dim DateSt      As String
    Dim DateEd      As String
    Dim BinSt       As String
    Dim BinEd       As String

    DBInput = True

    Call CisFrm.MousePT(11)

    On Error GoTo DBInput_Err:

' 抽出中メッセージ表示
    PNL_Wait.Visible = True
    PNL_PF.Enabled = False
    DoEvents

' 初期表示時、条件文作成(Where)
    If ProcHB = "H1" And InType = "" Then
'       表示条件 セット
        With CisDBGrid1
            .Select_Where = ""

        '   打切種別
            .Select_Where = .Select_Where & " where"
            If H1Cmb_UKbn.ListIndex = 0 Then
                .Select_Where = .Select_Where & " ("
            End If
            '>> 残数設定
            If H1Cmb_UKbn.ListIndex = 0 Or _
               H1Cmb_UKbn.ListIndex = 1 Then
                .Select_Where = .Select_Where & " ("
                .Select_Where = .Select_Where & " 打切種別 = 0"
                '> 未完了
                If H1Op_DspKbn(1).Value Then
                    .Select_Where = .Select_Where & " and 完了日 = ''"
                End If
                '> 完了
                If H1Op_DspKbn(2).Value Then
                    .Select_Where = .Select_Where & " and 完了日 <> ''"
                End If
            '   日付
                If H1Op_DspKbn(2).Value Then
                    DateStr = "完了"
                Else
                    DateStr = "起算"
                End If
                '> 日付　月指定
                If Trim(iH1_SDate) <> "" And CisFun.Val2(iH1_SDate.cDay) = 0 Then
                    .Select_Where = .Select_Where & " and substring(" & DateStr & "日, 1, 6) = '" & iH1_SDate & "'"
                   'If CisFun.Val2(iH1_SBin) <> 0 Then
                   '    .Select_Where = .Select_Where & " and " & DateStr & "便 = '" & iH1_SBin & "'"
                   'End If
                Else
                    DateSt = "": BinSt = ""
                    DateEd = "": BinEd = ""
                    '> 日付＆便　編集
                    If Trim(iH1_SDate) <> "" Then DateSt = iH1_SDate
                    If Trim(iH1_EDate) <> "" Then DateEd = iH1_EDate
                    If CisFun.Val2(iH1_SBin) <> 0 Then BinSt = Format(iH1_SBin, "00")
                    If CisFun.Val2(iH1_EBin) <> 0 Then BinEd = Format(iH1_EBin, "00")
                    '> 日付開始のみ
                    If Trim(DateSt) <> "" And Trim(DateEd) = "" Then
                        If Trim(BinSt) = "" Then
                            .Select_Where = .Select_Where & " and " & DateStr & "日 = '" & DateSt & "'"
                        Else
                            .Select_Where = .Select_Where & " and " & DateStr & "日+" & DateStr & "便 = '" & DateSt & BinSt & "'"
                        End If
                    End If
                    '> 日付終了のみ
                    If Trim(DateSt) = "" And Trim(DateEd) <> "" Then
                        If Trim(BinEd) = "" Then
                            .Select_Where = .Select_Where & " and " & DateStr & "日 <= '" & DateEd & "'"
                        Else
                            .Select_Where = .Select_Where & " and " & DateStr & "日+" & DateStr & "便 <= '" & DateSt & BinSt & "'"
                        End If
                    End If
                    '> 日付開始・終了
                    If Trim(DateSt) <> "" And Trim(DateEd) <> "" Then
                        If Trim(BinSt) = "" And Trim(BinEd) = "" Then
                            .Select_Where = .Select_Where & " and " & DateStr & "日 between '" & DateSt & "' and '" & DateEd & "'"
                        Else
                            If Trim(BinSt) = "" Then BinSt = "00"
                            If Trim(BinEd) = "" Then BinEd = "99"
                            .Select_Where = .Select_Where & " and " & DateStr & "日+" & DateStr & "便" _
                                                                    & " between '" & DateSt & BinSt & "'" _
                                                                        & " and '" & DateEd & BinEd & "'"
                        End If
                    End If
                End If
                .Select_Where = .Select_Where & " )"
            End If
            If H1Cmb_UKbn.ListIndex = 0 Then
                .Select_Where = .Select_Where & "   or"
            End If
            '>> 品番＆品番仕入先＆仕入先＆
            If H1Cmb_UKbn.ListIndex = 0 Or _
               H1Cmb_UKbn.ListIndex = 2 Or _
               H1Cmb_UKbn.ListIndex = 3 Or _
               H1Cmb_UKbn.ListIndex = 4 Then
                .Select_Where = .Select_Where & " ("
                If H1Cmb_UKbn.ListIndex = 0 Then
                    .Select_Where = .Select_Where & " 打切種別 in (1, 2, 3)"
                Else
                    .Select_Where = .Select_Where & " 打切種別 = " & H1Cmb_UKbn.ListIndex - 1
                End If
                '> 未完了
                If H1Op_DspKbn(1).Value Then
                    .Select_Where = .Select_Where & " and 打切日 > '" & Format(Now, "yyyymmdd") & "'"
                End If
                '> 完了
                If H1Op_DspKbn(2).Value Then
                    .Select_Where = .Select_Where & " and 打切日 <= '" & Format(Now, "yyyymmdd") & "'"
                End If
                '> 日付　月指定
                If Trim(iH1_SDate) <> "" And CisFun.Val2(iH1_SDate.cDay) = 0 Then
                    .Select_Where = .Select_Where & " and substring(打切日, 1, 6) = '" & iH1_SDate & "'"
                Else
                    DateSt = ""
                    DateEd = ""
                    '> 日付＆便　編集
                    If Trim(iH1_SDate) <> "" Then DateSt = iH1_SDate
                    If Trim(iH1_EDate) <> "" Then DateEd = iH1_EDate
                    '> 日付開始のみ
                    If Trim(DateSt) <> "" And Trim(DateEd) = "" Then
                        .Select_Where = .Select_Where & " and 打切日 = '" & DateSt & "'"
                    End If
                    '> 日付終了のみ
                    If Trim(DateSt) = "" And Trim(DateEd) <> "" Then
                        .Select_Where = .Select_Where & " and 打切日 <= '" & DateEd & "'"
                    End If
                    '> 日付開始・終了
                    If Trim(DateSt) <> "" And Trim(DateEd) <> "" Then
                        .Select_Where = .Select_Where & " and 打切日 between '" & DateSt & "' and '" & DateEd & "'"
                    End If
                End If
                .Select_Where = .Select_Where & " )"
            End If
            If H1Cmb_UKbn.ListIndex = 0 Then
                .Select_Where = .Select_Where & " )"
            End If

        '   品番
            If RTrim(iH1_Hinbn) <> "" Then
                If InStr(1, iH1_Hinbn, "*") = 0 Then
                    .Select_Where = .Select_Where & "   and 品番 = '" & RTrim(iH1_Hinbn) & "'"
                Else
                    .Select_Where = .Select_Where & "   and 品番 like '" & RTrim(Replace(iH1_Hinbn, "*", "%")) & "'"
                End If
            End If
        '   仕入先
            If RTrim(iH1_Torcd) <> "" Then
                .Select_Where = .Select_Where & "   and 仕入先 = '" & RTrim(iH1_Torcd) & "'"
            End If
        End With
    End If

' ﾚｲｱｳﾄ表示内容よりﾃﾞｰﾀ連結名を求める
    CisDBGrid1.LayOutGet

    gSL_Select = "select " & CisDBGrid1.LayOutfields
    gSL_Select = gSL_Select & ", 打切種別"
    gSL_Select = gSL_Select & ", 品名"
    gSL_Select = gSL_Select & ", 品番"
    gSL_Select = gSL_Select & ", 仕入先"
    gSL_Select = gSL_Select & " from ( "
    gSL_Select = gSL_Select & "select hut.*"
    gSL_Select = gSL_Select & ", case when hut.起算日 = '' then ''"
    gSL_Select = gSL_Select & "       else substring(hut.起算日,1,4)+'/'+substring(hut.起算日,5,2)+'/'+substring(hut.起算日,7,2)+'-'+hut.起算便"
    gSL_Select = gSL_Select & "   end 編集起算日"
    gSL_Select = gSL_Select & ", case when hut.完了日 = '' then ''"
    gSL_Select = gSL_Select & "       else substring(hut.完了日,1,4)+'/'+substring(hut.完了日,5,2)+'/'+substring(hut.完了日,7,2)+'-'+hut.完了便"
    gSL_Select = gSL_Select & "   end 編集完了日"
    gSL_Select = gSL_Select & ", case when hut.打切日 = '' then ''"
    gSL_Select = gSL_Select & "       else substring(hut.打切日,1,4)+'/'+substring(hut.打切日,5,2)+'/'+substring(hut.打切日,7,2)"
    gSL_Select = gSL_Select & "   end 編集打切日"
    gSL_Select = gSL_Select & ", case when hut.警告日 = '' then ''"
    gSL_Select = gSL_Select & "       else substring(hut.警告日,1,4)+'/'+substring(hut.警告日,5,2)+'/'+substring(hut.警告日,7,2)"
    gSL_Select = gSL_Select & "   end 編集警告日"
    gSL_Select = gSL_Select & ", case when isnull(hnm.表示品番,'') = '' then hut.品番 else hnm.表示品番"
    gSL_Select = gSL_Select & "   end 表示品番"
    gSL_Select = gSL_Select & ", case when hut.仕入先 = '' then ''"
    gSL_Select = gSL_Select & "       else hut.仕入先+space(7-len(hut.仕入先))+':'+isnull(tr1.略称,'')"
    gSL_Select = gSL_Select & "   end 編集仕入先"
    gSL_Select = gSL_Select & ", case when hut.切替仕入先 = '' then ''"
    gSL_Select = gSL_Select & "       else hut.切替仕入先+space(7-len(hut.切替仕入先))+':'+isnull(tr2.略称,'')"
    gSL_Select = gSL_Select & "   end 編集切替仕入先"
    gSL_Select = gSL_Select & ", isnull(hnm.品名,'') 品名"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  発注打切テーブル hut"
    gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm"
    gSL_Select = gSL_Select & "    on hnm.品番 = hut.品番"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ tr1"
    gSL_Select = gSL_Select & "    on tr1.取引先CD = hut.仕入先"
    gSL_Select = gSL_Select & "   and tr1.取引先区分 = 1"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ tr2"
    gSL_Select = gSL_Select & "    on tr2.取引先CD = hut.切替仕入先"
    gSL_Select = gSL_Select & "   and tr2.取引先区分 = 1"
    gSL_Select = gSL_Select & " ) xxx "
' 分類内容(Order By句) ｾｯﾄ
    With CisDBGrid1
        .OrderSet (Op_Sort(0).Value)
        If .Select_OrderBy = "" Then
            DBInput = False
            GoTo DBInput_Ed
        End If
    End With
' 抽出文　作成
    gSL_Select = gSL_Select & CisDBGrid1.Select_Where & CisDBGrid1.Select_OrderBy
' グリッド内表示
    If Not DBGridDisp Then
        DBInput = False
        GoTo DBInput_Ed
    End If
    On Error Resume Next
' 表示データなし
    On Error GoTo 0
    If Adodc1.Recordset.EOF Then
        DBInput = False
        GoTo DBInput_Ed
    End If

' 抽出件数表示
    lb_Kensu.Visible = True
    Call KensuDisp

    On Error GoTo 0
DBInput_Ed:
    Call CisFrm.MousePT(0)

    PNL_Wait.Visible = False
    PNL_PF.Enabled = True
    SendKeys "{END}"
    SendKeys "{RIGHT}"
    SendKeys "{LEFT}"
    SendKeys "{HOME}"
    Exit Function
DBInput_Err:
    Call CisFrm.MousePT(0)
    CisFun.ErrorBox
    PNL_Wait.Visible = False
    PNL_PF.Enabled = True
End Function



'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'
'+==========================================================+
'+      グリッド処理(TrueDBGrid)
'+==========================================================+
'+----------------------------------------------------------+
'+   グリッド内容初期設定
'+----------------------------------------------------------+
Private Sub GridInitSet1()
' Call .GridInitSet(ﾏｰｼﾞ処理,分類処理,ﾃﾞｰﾀ連結名,印刷見出し,印刷桁数,合計印刷)
    With CisDBGrid1
         Call .GridInitSet(True, , "表示品番", "品     番", 30)
         Call .GridInitSet(True, , "編集仕入先", "仕入先", 25)
         Call .GridInitSet(True, , "編集起算日", "起算日 - 便", 15)
         Call .GridInitSet(, , "設定数", "設定数", 10, False)
         Call .GridInitSet(, , "警告数", "警告数", 10, False)
         Call .GridInitSet(, , "発注数", "発注数", 10, True)
         Call .GridInitSet(, , "入荷数", "入荷数", 10, True)
         Call .GridInitSet(, , "設定残数", "設定残数", 10, False)
         Call .GridInitSet(True, , "編集完了日", "完了日 - 便", 15)
         Call .GridInitSet(True, , "編集打切日", "打切日", 10)
         Call .GridInitSet(True, , "編集警告日", "警告日", 10)
         Call .GridInitSet(True, , "編集切替仕入先", "切替仕入先", 25)
         Call .GridInitSet(True, , "切替受入", "切受", 5)
    End With
End Sub
'+----------------------------------------------------------+
'+      ADODC1(件数表示)                                     +
'+----------------------------------------------------------+
Private Sub KensuDisp()
    Static NowCnt       As Long
    Static MaxCnt       As Long
    NowCnt = Adodc1.Recordset.AbsolutePosition
    MaxCnt = Adodc1.Recordset.RecordCount
    If NowCnt < 0 Then Exit Sub
    lb_Kensu.Caption = Format(NowCnt, "###,##0") & " / " & _
                       Format(MaxCnt, "###,##0")

End Sub
'+---------------------+
'+  グリッド内容表示    +
'+---------------------+
Private Function DBGridDisp()
    DBGridDisp = False

    On Error GoTo DBGridDisp_ER:

    On Error Resume Next

    Adodc1.RecordSource = gSL_Select
    DBGrid1.ReBind
    Adodc1.Refresh
    DBGrid1.Refresh

' 重複表示
    Call CisDBGrid1.MargeDisp(Op_Overlap(1).Value)
    Call KensuDisp
    DBGrid1.SetFocus
    SendKeys "{HOME}"

    On Error GoTo 0

    DBGridDisp = True
    Exit Function
DBGridDisp_ER:
    If Err <> 40002 Then        ' RDO タイムアウト
       CisFun.ErrorBox
    End If
End Function
'+-----------------------+
'+  グリッド内容初期化    +
'+-----------------------+
Private Sub DBGridClr1()

    Call CisDBGrid1.SplitsClr("", 0)   ' 分割ﾍﾟｲﾝ解除

    Adodc1.Recordset.Close
    DBGrid1.Close True

    lb_Kensu.Caption = ""

End Sub
'+----------------------------------------------------------+
'+ キー入力にて行移動 《 KeyUp イベント   》                  +
'+----------------------------------------------------------+
Private Sub DBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or _
       KeyCode = vbKeyDown Or _
       KeyCode = vbKeyPageUp Or _
       KeyCode = vbKeyPageDown Then
        Call KensuDisp
        If B2_Area1.Visible Then
            Call BodySet
        End If
    End If
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 ColMove イベント   》            +
'+----------------------------------------------------------+
Private Sub DBGrid1_ColMove(ByVal Position As Integer, Cancel As Integer)
    CisDBGrid1.ColMoveFlg = True
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 Paint イベント   》              +
'+----------------------------------------------------------+
Private Sub DBGrid1_Paint()
    If Not FormAct Then Exit Sub
    If CisDBGrid1.ColMoveFlg Then  ' 再表示(分類)
        CisDBGrid1.ColMoveFlg = False
        Call DBInput
    End If
End Sub
'+----------------------------------------------------------+
'+ 明細画面表示(B2_AREA1)
'+----------------------------------------------------------+
Private Sub DBGrid1_DblClick()
    If Not B2_Area1.Visible Then
        Call ReturnPress
    End If
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseDown  イベント   》               +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If DBGrid1.RowContaining(Y) = -1 And _
        DBGrid1.ColContaining(x) = -1 Then
        Call CisDBGrid1.SplitsClr("C", 0)           ' 分割ﾍﾟｲﾝのみ解除
        Call DBInput
        Exit Sub
    End If
    If DBGrid1.RowContaining(Y) < 0 Then            ' 分割ﾍﾟｲﾝ表示
        If Button = 2 Then
            Call CisDBGrid1.SplitsAdd(DBGrid1.ColContaining(x))
            Call DBInput
        End If
        Exit Sub
    End If
    If DBGrid1.RowContaining(Y) < 0 Then Exit Sub
    If Button = 2 Then
        DBGrid1.Row = DBGrid1.RowContaining(Y)
        If B2_Area1.Visible Then
            Call BodySet
        End If
    End If
    If DBGrid1.RowContaining(x) < 0 Then Exit Sub

    If Button = 2 Then
        DBGrid1.Row = DBGrid1.RowContaining(Y)
        If B2_Area1.Visible Then
            Call BodySet
        End If
    Else
        DBGrid1.Row = DBGrid1.RowContaining(Y)
        If B2_Area1.Visible Then
            Call BodySet
        End If
    End If
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseUp  イベント   》                +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call KensuDisp
End Sub
'**********************************************************
'*  明細表示
'**********************************************************
Private Sub BodySet()
    Dim ICnt        As Integer
    Dim wUKbn       As Byte
    Dim wHinbn      As String
    Dim wTorcd      As String

    Call CisFrm.MousePT(11)

    For ICnt = 0 To 14
        If ICnt = 13 Then
            Select Case CisFun.Val2(DBGrid1.Columns(13))
                Case 0: B2lb_Mei(13) = "残数"
                Case 1: B2lb_Mei(13) = "品番"
                Case 2: B2lb_Mei(13) = "品番仕入先"
                Case 3: B2lb_Mei(13) = "仕入先"
            End Select
        Else
            B2lb_Mei(ICnt) = RTrim(DBGrid1.Columns(ICnt))
        End If
    Next ICnt

    wUKbn = CisFun.Val2(DBGrid1.Columns(13))
    wHinbn = RTrim(DBGrid1.Columns(15))
    wTorcd = RTrim(DBGrid1.Columns(16))

    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  htm.取引先"
    gSL_Select = gSL_Select & ", htm.受入"
    gSL_Select = gSL_Select & ", htm.背番号"
    gSL_Select = gSL_Select & ", isnull(trm.略称, '') 取引先名"
    gSL_Select = gSL_Select & ", isnull(hnm.表示品番, htm.品番) 表示品番"
    gSL_Select = gSL_Select & ", ("
    gSL_Select = gSL_Select & "select count(*)"
    gSL_Select = gSL_Select & "  from 構成マスタ kom"
    gSL_Select = gSL_Select & " where kom.親品番 = htm.品番"
    gSL_Select = gSL_Select & "    or kom.子品番 = htm.品番"
    gSL_Select = gSL_Select & "  ) 件数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  品番取引先マスタ htm"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = htm.取引先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm"
    gSL_Select = gSL_Select & "    on hnm.品番 = htm.品番 "
    Select Case wUKbn
        Case 0: gSL_Select = gSL_Select & " where htm.品番 = '" & wHinbn & "'"
        Case 1: gSL_Select = gSL_Select & " where htm.品番 = '" & wHinbn & "'"
        Case 3: gSL_Select = gSL_Select & " where htm.取引先 = '" & wTorcd & "'"
        Case Else
            gSL_Select = gSL_Select & " where htm.品番 = '" & wHinbn & "'"
            gSL_Select = gSL_Select & "   and htm.取引先 = '" & wTorcd & "'"
    End Select
    gSL_Select = gSL_Select & "   and htm.品目 = 1"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  htm.品番"
    gSL_Select = gSL_Select & ", htm.取引先"
    gSL_Select = gSL_Select & ", htm.受入"
    CisVsGrid1.Clear
    Call GridSet1

    Call CisFrm.MousePT(0)
End Sub

'+==========================================================+
'+      グリッド処理
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 左中, あり, 25, 1, "品　番")
        Call .FixedSet(中中, 左中, あり, 16, 2, "仕入先")
        Call .FixedSet(中中, 中中, あり, 4, 3, "受入")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
        'Call .FixedSet(中中, 左中, あり, 6, 4, "背番号")
        Call .FixedSet(中中, 左中, あり, 8, 4, "背番号")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
        Call .FixedSet(中中, 中中, あり, 4, 5, "構成")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 6, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5")
        .InitDisp
    End With
End Sub
'***********************************************'
'*****      グリッド内容セット
'***********************************************'
Private Function GridSet1() As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

    GridSet1 = False

    If HTMRead(gSL_Select, , 1) Then

'       抽出処理 初期ｾｯﾄ
        Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
'       ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
        With vsGrid1
           .Redraw = False
            If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
                .Rows = CisVsGrid1.DispMax
            Else
                .Rows = Ttl_Kensu + 1
            End If
        End With
        ID = 0
        Do Until Not HTM_RDSTS
            ID = ID + 1

            With vsGrid1
                .TextMatrix(ID, CisVsGrid1.FixedGet(0)) = ID
                .TextMatrix(ID, CisVsGrid1.FixedGet(1)) = RTrim(HTM.表示品番)
                .TextMatrix(ID, CisVsGrid1.FixedGet(2)) = HTM.取引先 & Space(7 - CisFun.Len2(HTM.取引先)) & ":" & RTrim(HTM.取引先名)
                .TextMatrix(ID, CisVsGrid1.FixedGet(3)) = RTrim(HTM.受入)
                .TextMatrix(ID, CisVsGrid1.FixedGet(4)) = RTrim(HTM.背番号)
                If HTM.件数 <> 0 Then
                    .TextMatrix(ID, CisVsGrid1.FixedGet(5)) = "○"
                End If
            End With

            Call HTMReadNext(1)
        Loop
        With vsGrid1
             .Row = 1: .Col = 1: .ColSel = .Cols - 1
             .Redraw = True
        End With

        GridSet1 = True
    End If
    Call HTMClose
End Function
