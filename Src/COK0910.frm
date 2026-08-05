VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form COK0910 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "購入追加・訂正入力"
   ClientHeight    =   11025
   ClientLeft      =   2730
   ClientTop       =   2475
   ClientWidth     =   15480
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11025
   ScaleWidth      =   15480
   Begin VB.PictureBox H3_Area1 
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
      Height          =   930
      Left            =   5445
      ScaleHeight     =   870
      ScaleWidth      =   4965
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1680
      Width           =   5025
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   795
         Left            =   30
         Top             =   45
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "納入番号"
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
         Begin CisText_V60.CisText iH3_Nouno 
            Height          =   360
            Left            =   150
            TabIndex        =   5
            Top             =   375
            Width           =   1230
            _ExtentX        =   2170
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
            Text            =   "XXXXXXXXX"
            MaxLength       =   9
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   795
         Left            =   1560
         Top             =   45
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "品　番"
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
         cPositionX      =   90
         cPositionY      =   30
         Begin CisText_V60.CisText iH3_Hinbn 
            Height          =   360
            Left            =   105
            TabIndex        =   6
            Top             =   360
            Width           =   3210
            _ExtentX        =   5662
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
            IMEMode         =   2
         End
      End
   End
   Begin VB.PictureBox H2_Area1 
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
      Height          =   930
      Left            =   435
      ScaleHeight     =   870
      ScaleWidth      =   4830
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1680
      Width           =   4890
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   795
         Left            =   60
         Top             =   45
         Width           =   4740
         _ExtentX        =   8361
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
         Begin VB.OptionButton OpH2_INKbn 
            BackColor       =   &H00C0FFC0&
            Caption         =   "追加"
            Height          =   300
            Index           =   2
            Left            =   270
            TabIndex        =   2
            Top             =   390
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton OpH2_INKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "入力訂正"
            Height          =   300
            Index           =   1
            Left            =   2970
            TabIndex        =   4
            Top             =   390
            Width           =   1380
         End
         Begin VB.OptionButton OpH2_INKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "伝票訂正"
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   0
            Left            =   1455
            TabIndex        =   3
            Top             =   390
            Width           =   1380
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   405
      Index           =   0
      Left            =   11115
      ScaleHeight     =   345
      ScaleWidth      =   2325
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2205
      Width           =   2385
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   345
         Index           =   1
         Left            =   15
         Top             =   0
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   609
         Caption         =   " 登録済件数"
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
         Begin Cis3D_v60.CIS3D Dsp_Kensu 
            Height          =   270
            Left            =   1290
            Top             =   30
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   476
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
   Begin VB.Timer Timer4 
      Left            =   210
      Top             =   810
   End
   Begin VB.PictureBox B2_Area1 
      Height          =   6180
      Left            =   420
      ScaleHeight     =   6120
      ScaleWidth      =   14385
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4725
      Width           =   14445
      Begin vsFlexLib.vsFlexArray VsGrid1 
         Height          =   5940
         Left            =   90
         TabIndex        =   21
         Top             =   90
         Width           =   14205
         _Version        =   196608
         _ExtentX        =   25056
         _ExtentY        =   10477
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "No |検 収 日  |納入番号 |品　番|赤黒|数量|仮区|単価|金額|背番号|車型|納入先|受入|検収管理NO|検収日|赤黒|納入先||||"
         Rows            =   17
         Cols            =   25
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   330
         SelectionMode   =   1
         AllowUserResizing=   1
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
      Height          =   930
      Left            =   450
      ScaleHeight     =   870
      ScaleWidth      =   5610
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   660
      Width           =   5670
      Begin Cis3D_v60.CIS3D HLB_UYm 
         Height          =   795
         Left            =   60
         Top             =   45
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "購入年月"
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
         Begin CisYM_V60.CisYM iH1_KYm 
            Height          =   390
            Left            =   135
            TabIndex        =   0
            Top             =   345
            Width           =   1485
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
      Begin Cis3D_v60.CIS3D HLB_Hinnm 
         Height          =   795
         Left            =   1815
         Top             =   45
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "契 約 先"
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1125
            Top             =   360
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   635
            BackColor       =   16777152
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   135
            TabIndex        =   1
            Top             =   360
            Width           =   1005
            _ExtentX        =   1773
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
   End
   Begin VB.Timer Timer3 
      Left            =   1980
      Top             =   810
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
      Left            =   7050
      ScaleHeight     =   660
      ScaleWidth      =   7785
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   570
      Width           =   7845
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   5175
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
         Left            =   6045
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
         Left            =   6900
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
         Left            =   2610
         Top             =   30
         Visible         =   0   'False
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
         Left            =   4320
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1740
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
         Left            =   885
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
      Left            =   990
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   540
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   930
      Top             =   660
   End
   Begin VB.Timer Timer2 
      Left            =   1440
      Top             =   720
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
      Caption         =   "【　購入追加・訂正入力  】"
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
         Left            =   10620
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   30
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14400
         Top             =   90
         Width           =   885
         _ExtentX        =   1561
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
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
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
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   8070
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
   Begin VB.PictureBox B1_Area1 
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
      Height          =   1905
      Left            =   435
      ScaleHeight     =   1845
      ScaleWidth      =   13020
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2790
      Width           =   13080
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   885
         Left            =   75
         Top             =   45
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "検 収 日"
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
         Begin CisYMD_With_Btn.CisYMDwB iB1_KYmd 
            Height          =   345
            Left            =   120
            TabIndex        =   7
            Top             =   390
            Width           =   1905
            _ExtentX        =   2831
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
            Object.Width           =   1905
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Cis3D_v60.CIS3D B2Msg_Nouno 
         Height          =   885
         Left            =   2220
         Top             =   45
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "納入番号"
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
         Begin CisText_V60.CisText iB1_NouNo 
            Height          =   360
            Left            =   150
            TabIndex        =   8
            Top             =   375
            Width           =   1230
            _ExtentX        =   2170
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
            Text            =   "XXXXXXXXX"
            MaxLength       =   9
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   885
         Left            =   4680
         Top             =   930
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "数 量"
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
         Begin CisText_V60.CisText iB1_Suryo 
            Height          =   360
            Left            =   165
            TabIndex        =   17
            Top             =   360
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   635
            cFormat         =   "#,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "#,###,###"
            MaxLength       =   9
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "#######"
            cILength        =   7
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D19 
         Height          =   885
         Left            =   7260
         Top             =   930
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "単 価"
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
         Begin CisText_V60.CisText iB1_Tanka 
            Height          =   360
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   635
            cFormat         =   "#,###,##0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "#,###,##0.00"
            MaxLength       =   12
            cDataType       =   2
            cDataReplace    =   1
            cFcDecimal      =   2
            cFdAutoFormat   =   1
            cGFormat        =   "######0.00"
            cILength        =   7
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   885
         Left            =   9090
         Top             =   930
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "金　額"
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
         Begin CisText_V60.CisText iB1_Kingk 
            Height          =   360
            Left            =   120
            TabIndex        =   20
            Top             =   345
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   635
            cFormat         =   "###,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###,###,###"
            MaxLength       =   11
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "#########"
            cILength        =   9
         End
      End
      Begin Cis3D_v60.CIS3D B_AKAMsg 
         Height          =   885
         Left            =   3465
         Top             =   930
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "取消"
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
         Begin VB.CheckBox B1Chk_Aka 
            BackColor       =   &H00C0C0C0&
            Caption         =   "訂正"
            Height          =   345
            Left            =   210
            TabIndex        =   16
            Top             =   360
            Width           =   825
         End
      End
      Begin Cis3D_v60.CIS3D B2Msg_Hinbn 
         Height          =   885
         Left            =   3750
         Top             =   45
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "品　番"
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
         cPositionX      =   90
         cPositionY      =   30
         Begin CisText_V60.CisText iB1_Hinbn 
            Height          =   360
            Left            =   105
            TabIndex        =   9
            Top             =   360
            Width           =   3210
            _ExtentX        =   5662
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D B2Msg_Seban 
         Height          =   885
         Left            =   7125
         Top             =   45
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "背番号"
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
         Begin CisText_V60.CisText iB1_Seban 
            Height          =   360
            Left            =   165
            TabIndex        =   10
            Top             =   360
            Width           =   1125
            _ExtentX        =   1984
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
            Text            =   "XXXXXXXX"
            MaxLength       =   8
            IMEMode         =   2
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D B2Msg_Ukeir 
         Height          =   885
         Left            =   12105
         Top             =   45
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "受入"
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
         Begin CisText_V60.CisText iB1_Ukeir 
            Height          =   360
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   390
            _ExtentX        =   688
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
            Text            =   "XX"
            MaxLength       =   2
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D B2Msg_Syakei 
         Height          =   885
         Left            =   75
         Top             =   930
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   1561
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
         Begin CisText_V60.CisText iB1_Syakei 
            Height          =   360
            Left            =   135
            TabIndex        =   13
            Top             =   360
            Width           =   750
            _ExtentX        =   1323
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
            Text            =   "XXXXX"
            MaxLength       =   5
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D B2Msg_Siire 
         Height          =   885
         Left            =   8505
         Top             =   45
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "仕 入 先"
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
         Begin Cis3D_v60.CIS3D B1lb_Siire 
            Height          =   360
            Left            =   1020
            Top             =   360
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   635
            BackColor       =   16777152
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin CisText_V60.CisText iB1_Siire 
            Height          =   360
            Left            =   30
            TabIndex        =   11
            Top             =   360
            Width           =   1005
            _ExtentX        =   1773
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
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   885
         Left            =   6210
         Top             =   930
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "仮単区"
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
         Begin VB.CheckBox B1Chk_Kari 
            BackColor       =   &H00C0C0C0&
            Caption         =   "正式"
            Height          =   345
            Left            =   90
            TabIndex        =   18
            Top             =   345
            Width           =   810
         End
      End
      Begin Cis3D_v60.CIS3D Back_Riyu 
         Height          =   885
         Left            =   1065
         Top             =   930
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "訂正理由"
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
         Begin Cis3D_v60.CIS3D B1lb_TKbn 
            Height          =   360
            Left            =   420
            Top             =   360
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   635
            BackColor       =   16777152
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
         End
         Begin CisText_V60.CisText iB1_TKbn 
            Height          =   360
            Left            =   150
            TabIndex        =   14
            Top             =   360
            Width           =   255
            _ExtentX        =   450
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
            Text            =   "X"
            MaxLength       =   1
         End
      End
      Begin Cis3D_v60.CIS3D B_JYmMsg 
         Height          =   885
         Left            =   10830
         Top             =   930
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "実績年月"
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
         Begin Cis3D_v60.CIS3D B1lb_JYm 
            Height          =   360
            Left            =   300
            Top             =   360
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "9999/99"
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
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D Back_Shiki 
         Height          =   885
         Left            =   1065
         Top             =   930
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "識別子"
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
         Begin Cis3D_v60.CIS3D B1lb_Shiki 
            Height          =   360
            Left            =   480
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            BackColor       =   16777152
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   50
         End
         Begin CisText_V60.CisText iB1_Shiki 
            Height          =   360
            Left            =   105
            TabIndex        =   15
            Top             =   360
            Width           =   375
            _ExtentX        =   661
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
            Text            =   "XX"
            MaxLength       =   2
         End
      End
   End
End
Attribute VB_Name = "COK0910"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************'
'**
'**       フォーム名    :   購入追加・訂正入力
'**       フォームID    :   COK0910
'**       処理概要      :
'**
'**       作  成  日    :   2004/12/06  By CIS SIMPR-A
'**       変  更  日    :   2005/05/19  追加機能を追加
'**                                     識別子　伝票訂正 "10" AND "11"
'**                                             入力訂正 "10" OR "11"
'**                                             追加 "04","05","06","07","08"
'**       変  更  日    :   2005/07/15  金額計算部修正( データ型:Doubleによる障害 )
'**                                      < ex. >　数量×単価＝金額
'**                                               200 × 4.1 = 819 (?)
'**       変  更  日    :   2008/04/11  背番号8桁,車型5桁
'**
'********************************************************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    Dim CisVsGrid1      As New CisVsGrid3

    Dim SaveKanriNo    As Long
    Dim SaveKYm         As String       ' 買掛処理年月
    Dim SaveRow         As Integer
    Dim SaveKeisan      As Byte
    Dim SaveKHoho       As Byte
    Dim SaveShime       As Byte
    Dim RowColFlg       As Boolean

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
'---------------------------------------< 2005.07.15 MNT START >
'    Dim KDouble         As Double
    Dim KCurren         As Currency
'---------------------------------------< 2005.07.15 MNT  END  >

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
'    Timer3.Interval = 500    ' 処理区分点滅間隔(取引先)
'    Timer3.Enabled = False   ' 処理区分点滅禁止(取引先)
'
    mGotColor = OpH2_INKbn(2).BackColor
    mLostColor = OpH2_INKbn(0).BackColor
    '+--------------------------------+
    '+  グリッド初期設定(材料⇒品番)
    '+--------------------------------+
    With CisVsGrid1
        Set .GridObj = VsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 23
        .InitGet = False
        .Init
    End With
    Call GridInit1
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = ""
    Call SyoriName(SyoriKB)
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
    If MeisyoGet("処理内容", "POK0110") Then
        SaveKYm = RTrim(SYM.英数字1)
    End If
'
    Call OpH2_INKbn_Click(2)
'   < 名称マスタ初期設定 >
    Call MeisyoInit
    
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
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
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
   
    If ProcHB = "H1" Then
        Call HeadBodyClear("H")
        Call HeadBodyClear("B")
        SyoriKB = ""
        Call SyoriName(SyoriKB)
        ProcHB = "H1"
        Call DispChange(ProcHB)
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "H2" Then
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "H3" Then
        Call HeadBodyClear("B1")
        ProcHB = "H2"
        Call DispChange(ProcHB)
        Call HeadBodyClear("H3")
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B1" Then
        If SyoriKB = "MNT" Then
            Call CisVsGrid1.vsColor(消去個別)
            ProcHB = "B2"
        Else
            If OpH2_INKbn(0).Value Then
                ProcHB = "H3"
            Else
                ProcHB = "H2"
            End If
            Call HeadBodyClear("B1")
        End If
        Call SyoriName(SyoriKB)
        Call DispChange(ProcHB)
        GoTo PB_CAN_Ed
    End If
            
    If ProcHB = "B2" Then
        SyoriKB = ""
        Call SyoriName(SyoriKB)
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        GoTo PB_CAN_Ed
    End If

PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新　規  キ ー(F1)        +
'+----------------------------+
Private Sub PB_ADD_Click()
    
    If SyoriKB <> "ADD" Then
        Call HeadBodyClear("B1")
        ProcHB = "H2"
        Call DispChange(ProcHB)
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
        OpH2_INKbn(2).SetFocus
        
        RowColFlg = True
        Call CisVsGrid1.vsColor(消去個別)
        RowColFlg = False
    End If
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)        +
'+----------------------------+
Private Sub PB_MNT_Click()
    
    If SyoriKB <> "MNT" Then
        Call HeadBodyClear("H3")
        ProcHB = "B2"
        Call DispChange(ProcHB)
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
        RowColFlg = True
        Call CisVsGrid1.vsColor(消去個別)
        RowColFlg = False

            
        VsGrid1.SetFocus
        Call VSGrid1_Click
    End If
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)        +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        Call HeadBodyClear("H3")
        ProcHB = "B2"
        Call DispChange(ProcHB)
        SyoriKB = "DEL"
        Call SyoriName(SyoriKB)
        RowColFlg = True
        Call CisVsGrid1.vsColor(消去個別)
        RowColFlg = False
        
        VsGrid1.SetFocus
        Call VSGrid1_Click
    End If

End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = " SELECT 実績年月,契約先 "
    gSL_Select = gSL_Select & " FROM 購入実績テーブル "
    gSL_Select = gSL_Select & " WHERE 入力区分 = ''"
    gSL_Select = gSL_Select & "   AND 実績年月 + 契約先 < '" & iH1_KYm + iH1_Torcd & "'"
    gSL_Select = gSL_Select & " GROUP BY 実績年月,契約先"
    gSL_Select = gSL_Select & " ORDER BY 実績年月 DESC,契約先 DESC"
    If KJSRead(gSL_Select, 1) Then
        iH1_KYm = KJS.実績年月
        iH1_Torcd = KJS.契約先
        Call iH1_Torcd_LostFocus
    End If

    Dummy.Enabled = True: Dummy.SetFocus
    iH1_KYm.SetFocus
    Dummy.Enabled = False
    
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = " SELECT 実績年月,契約先 "
    gSL_Select = gSL_Select & " FROM 購入実績テーブル "
    gSL_Select = gSL_Select & " WHERE 入力区分 = ''"
    gSL_Select = gSL_Select & "   AND 実績年月 + 契約先 > '" & iH1_KYm + iH1_Torcd & "'"
    gSL_Select = gSL_Select & " GROUP BY 実績年月,契約先"
    gSL_Select = gSL_Select & " ORDER BY 実績年月 ,契約先 "
    If KJSRead(gSL_Select, 1) Then
        iH1_KYm = KJS.実績年月
        iH1_Torcd = KJS.契約先
        Call iH1_Torcd_LostFocus
    End If

    Dummy.Enabled = True: Dummy.SetFocus
    iH1_KYm.SetFocus
    Dummy.Enabled = False
    
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
'
    If PB_Look.Tag = "iB1_KYmd" Then
        iB1_KYmd.ShowCalender
        Exit Sub
    End If
    If PB_Look.Tag = "iH3_Nouno" Then Call Look_Jisseki(0): Exit Sub
    If PB_Look.Tag = "iH3_Hinbn" Then Call Look_Jisseki(0): Exit Sub
    If PB_Look.Tag = "iB1_NouNo" Then Call Look_Jisseki(1): Exit Sub
    If PB_Look.Tag = "iB1_Siire" Then Call Look_Siire: Exit Sub
    If PB_Look.Tag = "iB1_Ukeir" Then Call Look_Ukeir: Exit Sub
    If PB_Look.Tag = "iB1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iB1_TKbn" Then Call Look_Kubun(0): Exit Sub
    If PB_Look.Tag = "iB1_Shiki" Then Call Look_Kubun(1): Exit Sub
    If PB_Look.Tag = "iB1_Tanka" Then Call Look_Tanka: Exit Sub

End Sub
'+----------------------------+
'+       契約先検索            +
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
'+----------------------------+
'+       品番検索              +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 1
    RVI_SK = 9
    RVI_Torcd = RTrim(iB1_Siire)
    RVI_Ukeir = RTrim(iB1_Ukeir)
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iB1_Hinbn = RTrim(RV_Hinbn)
        iB1_Seban = RTrim(RV_Seban)
    End If
End Sub
'+----------------------------+
'+       仕入先検索            +
'+----------------------------+
Private Sub Look_Siire()

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 1
    RV_TorKb1 = 0
    
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
        iB1_Siire = RTrim(RV_TorcdK)
        Call iB1_Siire_LostFocus
    End If
End Sub
'+----------------------------+
'+     仕入先受入検索          +
'+----------------------------+
Private Sub Look_Ukeir()

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 8
    RV_Torcd = RTrim(iB1_Siire)
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iB1_Siire = RTrim(RV_TorcdK)
        Call iB1_Siire_LostFocus
        iB1_Ukeir = RTrim(RV_Ukeir)
    End If
End Sub
'+----------------------------+
'+       名称マスタ検索        +
'+----------------------------+
Private Sub Look_Kubun(CKb As Integer)

    RV_Left = 0
    RV_Top = 0

    Rv_Select = ""
    Select Case CKb
        Case 0: RVI_Kubun = "買掛訂正理由"
        Case 1: RVI_Kubun = "買掛識別子"
                Rv_Select = " AND 英数字1 = 'K' AND 数字1 = 1"
    End Select
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        Select Case CKb
            Case 0: iB1_TKbn = RV_Code
                    Call iB1_TKbn_LostFocus
            Case 1: iB1_Shiki = RV_Code
                    Call iB1_Shiki_LostFocus
        End Select
    End If

End Sub
'+----------------------------+
'+       購入単価マスタ検索              +
'+----------------------------+
Private Sub Look_Tanka()

    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Torcd = RTrim(iB1_Siire)
    RV_Hinbn = RTrim(iB1_Hinbn)
    RV_SYmd = RTrim(iB1_KYmd)
    
    CKK0180.Show vbModal
    Unload CKK0180
    Set CKK0180 = Nothing

    If RV_Rtn Then
        iB1_Siire = RTrim(RV_Torcd)
        Call iB1_Siire_LostFocus
        iB1_Hinbn = RTrim(RV_Hinbn)
        iB1_Tanka = RTrim(RV_Tanka)
        Call iB1_Tanka_LostFocus
    End If
End Sub
'+----------------------------+
'+     購入実績検索          +
'+----------------------------+
Private Sub Look_Jisseki(JKB As Byte)

    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RV_YM = RTrim(iH1_KYm)
    RV_Torcd = RTrim(iH1_Torcd)
    RV_Hinbn = RTrim(iH3_Hinbn)
    RV_Code = RTrim(iH3_Nouno)
    
    COK0915.Show vbModal
    Unload COK0915
    Set COK0915 = Nothing

    If RV_Rtn Then
        If JKB = 0 Then
            iH3_Nouno = RTrim(RV_Code)
            iH3_Hinbn = RTrim(RV_Hinbn)
            SaveKanriNo = Rv_KanriNo
        Else
            iB1_NouNo = RTrim(RV_Code)
            iB1_Hinbn = RTrim(RV_Hinbn)
            SaveKanriNo = Rv_KanriNo
        End If
        Call DBInput2
    End If
End Sub
'+---( H1 )----------------------------------------------------------------------------
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
'+---( H2 )----------------------------------------------------------------------------
'+-----------------------------+
'+  OpH2_INKbn
'+-----------------------------+
Private Sub OpH2_INKbn_GotFocus(Index As Integer)
    OpH2_INKbn(0).BackColor = mLostColor
    OpH2_INKbn(1).BackColor = mLostColor
    OpH2_INKbn(2).BackColor = mLostColor
    OpH2_INKbn(Index).BackColor = mGotColor
End Sub
Private Sub OpH2_INKbn_LostFocus(Index As Integer)
'''    OpH2_INKbn(Index).BackColor = mLostColor
End Sub
Private Sub OpH2_INKbn_Click(Index As Integer)
    If Index = 0 Then
        B_AKAMsg = "取消"
        B1Chk_Aka.Caption = "訂正"
        B_JYmMsg = "実績年月"
        B1lb_JYm.Visible = True
    Else
        B_AKAMsg = "赤黒"
        B1Chk_Aka.Caption = "黒"
        B_JYmMsg = ""
        B1lb_JYm.Visible = False
    End If
    If Index = 2 Then
        Back_Shiki.Visible = True
        Back_Riyu.Visible = False
    Else
        Back_Shiki.Visible = False
        Back_Riyu.Visible = True
    End If
End Sub
'+---( H3 )----------------------------------------------------------------------------
'+-----------------------------+
'+  iH3_Nouno
'+-----------------------------+
Private Sub iH3_Nouno_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH3_Nouno_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH3_Hinbn
'+-----------------------------+
Private Sub iH3_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH3_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+---( B1 )----------------------------------------------------------------------------
'+-----------------------------+
'+  iB1_Nouno
'+-----------------------------+
Private Sub iB1_Nouno_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Nouno_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Hinbn
'+-----------------------------+
Private Sub iB1_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_KYmd
'+-----------------------------+
Private Sub iB1_KYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_KYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Siire
'+-----------------------------+
Private Sub iB1_Siire_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Siire_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_Siire = ""
    If TorNmGet(iB1_Siire, 8) Then
        B1lb_Siire = TRM.略称
    End If
End Sub
'+-----------------------------+
'+  iB1_Ukeir
'+-----------------------------+
Private Sub iB1_Ukeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_TKbn
'+-----------------------------+
Private Sub iB1_TKbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_TKbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_TKbn = ""
    If MeisyoGet("買掛訂正理由", iB1_TKbn) Then
        B1lb_TKbn = RTrim(SYM.値名称)
    End If
End Sub
'+-----------------------------+
'+  iB1_Shiki
'+-----------------------------+
Private Sub iB1_Shiki_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Shiki_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_Shiki = ""
    If MeisyoGet("買掛識別子", iB1_Shiki) Then
        If SYM.数字1 = 1 And SYM.英数字1 = "K" Then
            B1lb_Shiki = RTrim(SYM.値名称)
        End If
    End If
End Sub
'+-----------------------------+
'+  B1Chk_Aka
'+-----------------------------+
Private Sub B1Chk_Aka_Click()
    If OpH2_INKbn(0).Value Then
        If B1Chk_Aka.Value = 0 Then
            B1Chk_Aka.Caption = "訂正"
        Else
            B1Chk_Aka.Caption = "取消"
        End If
    Else
        If B1Chk_Aka.Value = 0 Then
            B1Chk_Aka.Caption = "黒"
        Else
            B1Chk_Aka.Caption = "赤"
        End If
    End If
End Sub
'+-----------------------------+
'+  iB1_Suryo
'+-----------------------------+
Private Sub iB1_Suryo_LostFocus()
    If CisFun.Val2(iB1_Suryo) <> 0 And CisFun.Val2(iB1_Tanka) <> 0 Then
        
'----------------------------------------------------------------------------< 2005.07.15 MNT START >
'
'        KDouble = CisFun.Val2(iB1_Tanka) * CisFun.Val2(iB1_Suryo)
'        Select Case SaveKeisan
'            Case 1: KDouble = CisFun.HasuProc(KDouble, 切上げ, 0)
'            Case 2: KDouble = CisFun.HasuProc(KDouble, 四捨五入, 0)
'            Case Else: KDouble = CisFun.HasuProc(KDouble, 切捨て, 0)
'        End Select
'        iB1_Kingk = KDouble
        
        KCurren = CisFun.Val2(iB1_Tanka) * CisFun.Val2(iB1_Suryo)
        Select Case SaveKeisan
            Case 1: KCurren = CisFun.HasuProc(CCur(KCurren), 切上げ, 0)
            Case 2: KCurren = CisFun.HasuProc(CCur(KCurren), 四捨五入, 0)
            Case Else: KCurren = CisFun.HasuProc(CCur(KCurren), 切捨て, 0)
        End Select
        iB1_Kingk = KCurren
'
'----------------------------------------------------------------------------< 2005.07.15 MNT START >

    End If
End Sub
'+-----------------------------+
'+  B1Chk_Kari
'+-----------------------------+
Private Sub B1Chk_Kari_Click()
    If B1Chk_Kari.Value = 0 Then
        B1Chk_Kari.Caption = "正式"
    Else
        B1Chk_Kari.Caption = "仮"
    End If
End Sub
'+-----------------------------+
'+  iB1_Tanka
'+-----------------------------+
Private Sub iB1_Tanka_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Tanka_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    If CisFun.Val2(iB1_Suryo) <> 0 And CisFun.Val2(iB1_Tanka) <> 0 Then

'----------------------------------------------------------------------------< 2005.07.15 MNT START >
'
'        KDouble = CisFun.Val2(iB1_Tanka) * CisFun.Val2(iB1_Suryo)
'        Select Case SaveKeisan
'            Case 1: KDouble = CisFun.HasuProc(KDouble, 切上げ, 0)
'            Case 2: KDouble = CisFun.HasuProc(KDouble, 四捨五入, 0)
'            Case Else: KDouble = CisFun.HasuProc(KDouble, 切捨て, 0)
'        End Select
'        iB1_Kingk = KDouble
        
        KCurren = CisFun.Val2(iB1_Tanka) * CisFun.Val2(iB1_Suryo)
        Select Case SaveKeisan
            Case 1: KCurren = CisFun.HasuProc(CCur(KCurren), 切上げ, 0)
            Case 2: KCurren = CisFun.HasuProc(CCur(KCurren), 四捨五入, 0)
            Case Else: KCurren = CisFun.HasuProc(CCur(KCurren), 切捨て, 0)
        End Select
        iB1_Kingk = KCurren
'
'----------------------------------------------------------------------------< 2005.07.15 MNT START >
    
    End If
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
'+------------------------------------------------------------------+
'+      基本情報
'+------------------------------------------------------------------+
Private Sub SyoriName(SYKB As String)
    Dim w_Color     As Long
    w_Color = &H80000008 '初期色ｾｯﾄ(黒)
   
'   H2LB_Size_B.Enabled = False
   
    Select Case SyoriKB
      Case "ADD"
           SyoriNM = "追加"
           K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
               
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
'      * PF-Key 表示設定 *
        If SyoriKB <> "LOK" Then
           If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
              gObj.Name = "PB_DEL" Then
              gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
           End If
        End If
    Next gObj
    If SyoriKB = "ADD" Then
'        Call OpH2_INKbn_LostFocus(0)
'        OpH2_INKbn(2).Value = True
'        Call OpH2_INKbn_Click(2)
    Else
        If SyoriKB <> "" Then
            Call OpH2_INKbn_LostFocus(0)
            OpH2_INKbn(2).Value = True
            Call OpH2_INKbn_Click(2)
        End If
    End If
   
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
    Dim H2Mode      As Boolean
    Dim H2Color     As Long
    Dim H3Mode      As Boolean
    Dim H3Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    Dim B2Mode      As Boolean
    Dim B2Color     As Long
    
    H1Mode = False
    H1Color = gPLostSel
    H2Mode = False
    H2Color = gPLostSel
    H3Mode = False
    H3Color = gPLostSel
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
       Case "H3"
          H3Mode = True
          H3Color = gPGotSel
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
    H3_Area1.Enabled = H3Mode
    H3_Area1.BackColor = H3Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    If Pro = "H1" And SyoriKB = "LOK" Then
        B2_Area1.Enabled = H1Mode
        B2_Area1.BackColor = H1Color
    Else
        B2_Area1.Enabled = B2Mode
        B2_Area1.BackColor = B2Color
    End If

    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_END.Visible = H1Mode Or H2Mode Or H3Mode
    PB_ADD.Visible = Not H1Mode
    PB_MNT.Visible = Not H1Mode
    PB_DEL.Visible = Not H1Mode
    
    If Pro = "B1" Then
        If OpH2_INKbn(0).Value Then
            iB1_NouNo.cLostColor = &HFFFFC0
            iB1_Hinbn.cLostColor = &HFFFFC0
            iB1_Siire.cLostColor = &HFFFFC0
            iB1_Seban.cLostColor = &HFFFFC0
            iB1_Ukeir.cLostColor = &HFFFFC0
            iB1_Syakei.cLostColor = &HFFFFC0
            B2Msg_Nouno.Enabled = False
            B2Msg_Hinbn.Enabled = False
            B2Msg_Siire.Enabled = False
            B2Msg_Seban.Enabled = False
            B2Msg_Ukeir.Enabled = False
            B2Msg_Syakei.Enabled = False
        End If
    Else
        iB1_NouNo.cLostColor = &HC0E0FF
        iB1_Hinbn.cLostColor = &HC0E0FF
        iB1_Siire.cLostColor = &HC0E0FF
        iB1_Seban.cLostColor = &HC0E0FF
        iB1_Ukeir.cLostColor = &HC0E0FF
        iB1_Syakei.cLostColor = &HC0E0FF
        B2Msg_Nouno.Enabled = True
        B2Msg_Hinbn.Enabled = True
        B2Msg_Siire.Enabled = True
        B2Msg_Seban.Enabled = True
        B2Msg_Ukeir.Enabled = True
        B2Msg_Syakei.Enabled = True
    End If
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
    
    On Error GoTo HeadBodyClear_ER
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
              '  .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
                .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    
    
    If HeadBody = "B1" Or HeadBody = "B" Or HeadBody = "ALL" Then
        B1Chk_Aka.Value = 0
        If OpH2_INKbn(0).Value Then
            B1Chk_Aka.Caption = "訂正"
        Else
            B1Chk_Aka.Caption = "黒"
        End If
        B1Chk_Kari.Value = 0
        B1Chk_Kari.Caption = "正式"
    End If
    If HeadBody = "B2" Or HeadBody = "B" Or HeadBody = "ALL" Then
        CisVsGrid1.Clear
        Call CisVsGrid1.Clear
        Dsp_Kensu = ""
    End If
    Exit Sub
HeadBodyClear_ER:
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
'( HEAD1 )
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        If SyoriKB <> "LOK" Then
            If ProcHB = "B2" Then
                Call VSGrid1_Click
            Else
                OpH2_INKbn(2).SetFocus
            End If
        End If
        GoTo ReturnPress_Ed
    End If
'( HEAD2 )
    If ProcHB = "H2" Then
        If OpH2_INKbn(0).Value Then
            ProcHB = "H3"
        Else
            If SyoriKB = "ADD" Then
                ProcHB = "B1"
            Else
                ProcHB = "B2"
            End If
        End If
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
'( HEAD3 )
    If ProcHB = "H3" Then
        If Not Head3Chk Then GoTo ReturnPress_Ed
        
        ProcHB = "B1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        GoTo ReturnPress_Ed
    End If
'( BODY1 )
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut(0) Then
            GoTo ReturnPress_Ed
        End If
        If SyoriKB = "ADD" Then
            ProcHB = "H2"
            Call DispChange(ProcHB)
        Else
            ProcHB = "B2"
            Call DispChange(ProcHB)
            Call VSGrid1_Click
            VsGrid1.SetFocus
        End If
        GoTo ReturnPress_Ed
    End If
    
'( BODY2 vsGrid1 )
    If ProcHB = "B2" Then
        If Not Body2Chk Then GoTo ReturnPress_Ed
' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then
                RowColFlg = True
                Call CisVsGrid1.vsColor(消去個別)
                RowColFlg = False
                GoTo ReturnPress_Ed
            End If
            Call HeadBodyClear("B1")
            If CisFun.Val2(Dsp_Kensu) <> 0 Then
                ProcHB = "B2"
                Call DispChange(ProcHB)
                Call SyoriName(SyoriKB)
                Call VSGrid1_Click
                Call VsGrid1.SetFocus
            Else
                ProcHB = "H2"
                Call DispChange(ProcHB)
                SyoriKB = "ADD"
                Call SyoriName(SyoriKB)
                OpH2_INKbn(0).SetFocus
            End If
        Else
            ProcHB = "B1"
            Call DispChange(ProcHB)
        End If
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
' 購入年月
    If Trim(iH1_KYm) = "" Then
        CisFun.MB_MSG(2) = "    購入年月を入力して下さい            "
        CisFun.MBOX
       iH1_KYm.SetFocus
       Exit Function
    End If
    If Not iH1_KYm.cChkResult Then
        CisFun.MB_MSG(2) = "    正しい購入年月を入力して下さい            "
        CisFun.MBOX
       iH1_KYm.SetFocus
       Exit Function
    End If
    SyoriKB = ""
    If Trim(iH1_KYm) < RTrim(SaveKYm) Then
        SyoriKB = "LOK"
        CisFun.MB_MSG(2) = "    買掛更新年月より前の年月の入力は不可！            "
        CisFun.MB_MSG(4) = "    　　　【参照】処理とします            "
        CisFun.MBOX
       iH1_KYm.SetFocus
'       Exit Function
    End If
' 契約先
    If Trim(iH1_Torcd) = "" Then
        CisFun.MB_MSG(2) = "    契約先を入力して下さい            "
        CisFun.MBOX
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If Not TorNmGet(iH1_Torcd, 5) Then
        CisFun.MB_MSG(2) = "     契約先 未登録！            "
        CisFun.MB_MSG(4) = "    （取引先マスタ）            "
        CisFun.MBOX
       iH1_Torcd.SetFocus
       Exit Function
    End If
    SaveKeisan = TRM.計算単位
    SaveKHoho = TRM.計算方法
    SaveShime = TRM.締日
'購入実績テーブル
    If SyoriKB = "LOK" Then
    '(照会)
        If Not DBInput("INV") Then
            CisFun.MB_MSG(2) = "    　対象データが存在しません            "
            CisFun.MB_MSG(4) = ""
            CisFun.MBOX
            iH1_KYm.SetFocus
            Exit Function
        End If
        ProcHB = "H1"
        
    Else
        If Not DBInput("INV") Then
            SyoriKB = "ADD"
            ProcHB = "H2"
        Else
            SyoriKB = "MNT"
            ProcHB = "B2"
        End If
    End If
    Head1Chk = True
End Function
'****************************
'*      HEAD3ﾁｪｯｸ処理
'****************************
Private Function Head3Chk() As Boolean
    Head3Chk = False
'
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
' 納入番号
'    If Trim(iH3_Nouno) = "" Then
'        CisFun.MB_MSG(2) = "    納入番号を入力して下さい            "
'        CisFun.MBOX
'        iH3_Nouno.SetFocus
'        Exit Function
'    End If
' 品番
    If Trim(iH3_Hinbn) = "" Then
        CisFun.MB_MSG(2) = "    品番を入力して下さい            "
        CisFun.MBOX
        iH3_Hinbn.SetFocus
        Exit Function
    End If
'
    SaveKanriNo = 0
    gSL_Select = "SELECT COUNT(*) 件数 FROM 購入実績テーブル"
    gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 納入番号 = '" & RTrim(iH3_Nouno) & "'"
    gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH3_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND 実績年月 < '" & RTrim(iH1_KYm) & "'"
    gSL_Select = gSL_Select & "   AND 識別子 <> '11'"                       '赤訂正以外
    If KJSRead(gSL_Select, 1) Then
        If KJS.件数 = 0 Then
            CisFun.MB_MSG(2) = "    　対象データが存在しません            "
            CisFun.MBOX
            iH3_Nouno.SetFocus
            Exit Function
        End If
        If KJS.件数 > 1 Then
            With CisFun
                .MB_Lines = 4
                .MB_Title = ""
                .MB_Button = OK
                .MB_MSG(2) = "    　この納入番号・品番の購入実績が複数存在します。            "
                .MB_MSG(4) = "    　　　　　検索画面にて選択して下さい。            "
                .MBOX
            End With
        
            RV_Left = 0: RV_Top = 0
            RV_Call = "C"
            RV_YM = RTrim(iH1_KYm)
            RV_Torcd = RTrim(iH1_Torcd)
            RV_Hinbn = RTrim(iH3_Hinbn)
            RV_Code = RTrim(iH3_Nouno)
            
            COK0915.Show vbModal
            Unload COK0915
            Set COK0915 = Nothing
        
            If RV_Rtn Then
                iH3_Nouno = RTrim(RV_Code)
                iH3_Hinbn = RTrim(RV_Hinbn)
                SaveKanriNo = Rv_KanriNo
            Else
                Exit Function
            End If
        End If
    Else
        Exit Function
    End If
'購入実績テーブル参照
    If Not DBInput2 Then
        CisFun.MB_MSG(2) = "    　対象データが存在しません            "
        CisFun.MBOX
        iH3_Nouno.SetFocus
        Exit Function
    End If
    
    Head3Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    
    Body1Chk = False
        
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
    ' 検収日
    If Trim(iB1_KYmd) = "" Then
        CisFun.MB_MSG(2) = "    検収日を入力して下さい            "
        CisFun.MBOX
        iB1_KYmd.SetFocus
        Exit Function
    End If
    If Not iB1_KYmd.cChkResult Then
        CisFun.MB_MSG(2) = "    正しい検収日を入力して下さい            "
        CisFun.MBOX
        iB1_KYmd.SetFocus
        Exit Function
    End If
    ' 納入番号
'    If Trim(iB1_NouNo) = "" Then
'        CisFun.MB_MSG(2) = "    納入番号を入力して下さい            "
'        CisFun.MBOX
'        iB1_NouNo.SetFocus
'        Exit Function
'    End If
    ' 品番
    If Trim(iB1_Hinbn) = "" Then
        CisFun.MB_MSG(2) = "    品番を入力して下さい            "
        CisFun.MBOX
        iB1_Hinbn.SetFocus
        Exit Function
    End If
    ' 背番号半角チェック
    If RTrim(RTrim(iB1_Seban)) <> "" And Not HalfSizeChk(RTrim(iB1_Seban)) Then
        With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "    背番号は半角にて入力して下さい            "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
        End With
        iB1_Seban.SetFocus
        Exit Function
    End If
    ' 仕入先
    If RTrim(iB1_Siire) <> "" Then
        If Not TorNmGet(iB1_Siire, 8) Then
            CisFun.MB_MSG(2) = "     仕入先 未登録！            "
            CisFun.MB_MSG(4) = "    （取引先マスタ）            "
            CisFun.MBOX
           iB1_Siire.SetFocus
           Exit Function
        End If
    End If
    
    If Not OpH2_INKbn(2).Value Then
        ' 訂正理由
        If RTrim(iB1_TKbn) <> "" And Not MeisyoGet("買掛訂正理由", iB1_TKbn) Then
            CisFun.MB_MSG(2) = "     訂正理由 未登録！            "
            CisFun.MBOX
            iB1_TKbn.SetFocus
            Exit Function
        End If
    Else
        ' 識別子
        If RTrim(iB1_Shiki) = "" Then
            CisFun.MB_MSG(2) = "     識別子 必須！            "
            CisFun.MB_MSG(4) = "    　　　　　　　　　　　     "
            CisFun.MBOX
           iB1_Shiki.SetFocus
           Exit Function
        End If
        If Not MeisyoGet("買掛識別子", iB1_Shiki) Then
            CisFun.MB_MSG(2) = "     識別子 未登録！            "
            CisFun.MBOX
            iB1_Shiki.SetFocus
            Exit Function
        Else
            If SYM.数字1 <> 1 Or SYM.英数字1 <> "K" Then
                CisFun.MB_MSG(2) = "     識別子 対象外！            "
                CisFun.MBOX
                iB1_Shiki.SetFocus
                Exit Function
            End If
        End If
    End If
    ' 金額
    If CisFun.Val2(iB1_Suryo) <> 0 And CisFun.Val2(iB1_Tanka) <> 0 Then

'----------------------------------------------------------------------------< 2005.07.15 MNT START >
'
'        KDouble = CisFun.Val2(iB1_Tanka) * CisFun.Val2(iB1_Suryo)
'        Select Case SaveKeisan
'            Case 1: KDouble = CisFun.HasuProc(KDouble, 切上げ, 0)
'            Case 2: KDouble = CisFun.HasuProc(KDouble, 四捨五入, 0)
'            Case Else: KDouble = CisFun.HasuProc(KDouble, 切捨て, 0)
'        End Select
'        iB1_Kingk = KDouble
'
'        If CisFun.Val2(iB1_Kingk) <> KDouble Then
'            CisFun.MB_MSG(2) = "    金額 <> 数量 × 単価            "
'            CisFun.MBOX
'            iB1_Kingk.SetFocus
'            Exit Function
'        End If
        
        KCurren = CisFun.Val2(iB1_Tanka) * CisFun.Val2(iB1_Suryo)
        Select Case SaveKeisan
            Case 1: KCurren = CisFun.HasuProc(CCur(KCurren), 切上げ, 0)
            Case 2: KCurren = CisFun.HasuProc(CCur(KCurren), 四捨五入, 0)
            Case Else: KCurren = CisFun.HasuProc(CCur(KCurren), 切捨て, 0)
        End Select

        If CisFun.Val2(iB1_Kingk) <> KCurren Then
            CisFun.MB_MSG(2) = "    金額 <> 数量 × 単価            "
            CisFun.MBOX
            iB1_Kingk.SetFocus
            Exit Function
        End If
'
'----------------------------------------------------------------------------< 2005.07.15 MNT START >
    
    End If
    ' 金額
    If CisFun.Val2(iB1_Kingk) = 0 Then
        CisFun.MB_MSG(2) = "    金額を入力して下さい            "
        CisFun.MBOX
        iB1_Kingk.SetFocus
        Exit Function
    End If
    
    
    Body1Chk = True
End Function
'****************************
'*      Body2ﾁｪｯｸ処理
'****************************
Private Function Body2Chk() As Boolean
    
    Body2Chk = False
        
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
    ' 選択
    If CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(15))) = 0 Then Exit Function
    
    If CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(21))) <> 0 And _
        CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(20))) = "11" And _
        SyoriKB = "MNT" Then
        CisFun.MB_MSG(2) = "    　伝票訂正の元データの修正不可！            "
        CisFun.MBOX
        VsGrid1.SetFocus
        Exit Function
    End If
    
    RowColFlg = True
    If SyoriKB = "MNT" Then
        Call CisVsGrid1.vsColor(個別, &HC0FFC0)
    Else
        Call CisVsGrid1.vsColor(個別, &HFFC0FF)
    End If
    RowColFlg = False
    
    
    Body2Chk = True
End Function
'***********************************************************************************
'*
'*       データベース入力 ＆ 表示
'*
'***********************************************************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("H3")
    Call HeadBodyClear("B")
    
    gSL_Select = "SELECT KJ.*,ISNULL(TR.略称,'')仕入先名,ISNULL(SY.値名称,'')訂正理由名, "
    gSL_Select = gSL_Select & " ISNULL(HN.表示品番,'')表示品番,ISNULL(SY2.値名称,'')識別子名"
    gSL_Select = gSL_Select & " FROM 購入実績テーブル KJ"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN "
    gSL_Select = gSL_Select & "   ON HN.品番 = KJ.品番 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR "
    gSL_Select = gSL_Select & "   ON TR.取引先区分 = 1 "
    gSL_Select = gSL_Select & "   AND TR.取引先CD = KJ.仕入先 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY "
    gSL_Select = gSL_Select & "   ON SY.区分名称 = '買掛訂正理由' "
    gSL_Select = gSL_Select & "   AND SY.区分タイプ = '' "
    gSL_Select = gSL_Select & "   AND SY.値 = KJ.訂正理由 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2 "
    gSL_Select = gSL_Select & "   ON SY2.区分名称 = '買掛識別子' "
    gSL_Select = gSL_Select & "   AND SY2.区分タイプ = '' "
    gSL_Select = gSL_Select & "   AND SY2.値 = KJ.識別子 "
    gSL_Select = gSL_Select & " WHERE KJ.実績年月 = '" & RTrim(iH1_KYm) & "'"
    gSL_Select = gSL_Select & "   AND KJ.契約先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND KJ.入力区分 = ''"
    gSL_Select = gSL_Select & " ORDER BY KJ.検収日,KJ.品番,KJ.納入番号,KJ.検収管理NO,KJ.識別子"
    
    If Not KJSRead(gSL_Select, 0) Then
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
    Else
        Call Grid1Set
        DBInput = True
    End If

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
Private Function DBInput2() As Boolean
    DBInput2 = False
    On Error GoTo DBInput2_Err:
    Call HeadBodyClear("B1")
    
    gSL_Select = "SELECT * FROM 購入実績テーブル"
    If SaveKanriNo = 0 Then
        gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 納入番号 = '" & RTrim(iH3_Nouno) & "'"
        gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH3_Hinbn) & "'"
        gSL_Select = gSL_Select & "   AND 実績年月 < '" & RTrim(iH1_KYm) & "'"
        gSL_Select = gSL_Select & "   AND 識別子 <> '11'"                       '赤訂正以外
        gSL_Select = gSL_Select & " ORDER BY 実績年月 DESC,作成日 DESC"
    Else
        gSL_Select = gSL_Select & " WHERE 検収管理NO = " & SaveKanriNo & ""
    End If
    If KJSRead(gSL_Select, 1) Then
        With KJS
            iB1_KYmd = RTrim(.検収日)
            iB1_NouNo = RTrim(.納入番号)
            iB1_Hinbn = RTrim(.品番)
            iB1_Seban = RTrim(.背番号)
            iB1_Syakei = RTrim(.車型)
            iB1_Siire = RTrim(.仕入先)
            Call iB1_Siire_LostFocus
            iB1_Ukeir = RTrim(.受入)
            If .赤黒区分 = 0 Then
                B1Chk_Aka.Value = 0
            Else
                B1Chk_Aka.Value = 1
            End If
            If .部品区分 = "K" Then
                iB1_TKbn = "4"
                Call iB1_TKbn_LostFocus
            End If
            iB1_Suryo = CisFun.Val2(.数量)
            
            If RTrim(.仮単価区分) = "" Then
                B1Chk_Kari.Value = 0
            Else
                B1Chk_Kari.Value = 1
            End If
            iB1_Tanka = CisFun.RSetFld(.単価, 12, "#,##0.00")
            iB1_Kingk = CisFun.RSetFld(.金額, 11, "#,##0")
            B1lb_JYm = Left(.実績年月, 4) & "/" & Mid(.実績年月, 5, 2)
                    
        End With
        DBInput2 = True
    End If

DBInput2_Ed:
    Exit Function
DBInput2_Err:
     CisFun.ErrorBox
End Function
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------------+
'+      削 除 処 理            +
'+----------------------------+
Private Function DBDelete()
    DBDelete = False
'  ( 確認メッセージ )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     削除処理を行います      "
        .MB_Title = "削除処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
        .SQL = "DELETE FROM 購入実績テーブル "
        .SQL = .SQL & " WHERE 検収管理NO = " & CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(15))) & ""
        .DBExec

        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")
        
        On Error GoTo 0
    End With
    
'Grid再表示
    Call vsGrid1_ReSetDisp
    
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut(KB As Byte)
    DBPut = False
    Dim AddFlg  As Boolean

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        If SyoriKB = "ADD" Then
            .MB_MSG(2) = "     追加処理を行います      "
            .MB_Title = "追加処理"
        Else
            .MB_MSG(2) = "     修正処理を行います      "
            .MB_Title = "修正処理"
        End If
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err

'*-----------------------------------------------'
'*  伝票訂正の時、赤伝出力
'*-----------------------------------------------'
    If OpH2_INKbn(0).Value Then
        With CisDB
            .SQL = "番号獲得"
            .StoadoCount = 3
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = OutPut用
            .ParaNo = 3: .ParaIO = Return用
            .ParaNo = 1: .ParaValue = 16
            If Not .DBStored Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "     検収管理NOの採番に失敗しました    "
                    .MB_Title = "採番処理"
                    .MB_Button = Error
                    If Not .MBOX Then GoTo DBPut_Err
                End With
            Else
                .ParaNo = 2: SaveKanriNo = .ParaValue
            End If
        End With
        With KJS
            .訂正管理NO = .検収管理NO
            .検収管理NO = SaveKanriNo
            .実績年月 = RTrim(iH1_KYm)
            .入力区分 = ""
            .識別子 = "11"
            .契約先 = RTrim(iH1_Torcd)
            .実績区分 = 1
            .実績処理日 = Format(Now(), "yyyymmdd")
            .赤黒区分 = 1
            .数量 = .数量 * -1
            .金額 = .金額 * -1
            .訂正理由 = RTrim(iB1_TKbn)
        End With
        Call KJSInsert
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
    Else
        Call ItemsClearKJS
    End If
    
    If OpH2_INKbn(0).Value And B1Chk_Aka.Value = 1 Then GoTo DBPut_Ed
    
'*-----------------------------------------------'
'*  伝票訂正の黒伝、入力訂正分の出力
'*-----------------------------------------------'
    If SyoriKB = "ADD" Then
        With CisDB
            .SQL = "番号獲得"
            .StoadoCount = 3
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = OutPut用
            .ParaNo = 3: .ParaIO = Return用

            .ParaNo = 1: .ParaValue = 16

            If Not .DBStored Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "     検収管理NOの採番に失敗しました    "
                    .MB_Title = "採番処理"
                    .MB_Button = Error
                    If Not .MBOX Then GoTo DBPut_Err
                End With
            Else
                .ParaNo = 2: SaveKanriNo = .ParaValue
            End If
        End With
    Else
        SaveKanriNo = CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(15)))
    End If
    
    With KJS
        .検収管理NO = SaveKanriNo
        .実績年月 = RTrim(iH1_KYm)
        .入力区分 = ""
        If OpH2_INKbn(0).Value Then
            .識別子 = "10"
        Else
            If OpH2_INKbn(1).Value Then
                If B1Chk_Aka.Value = 0 Then
                    .識別子 = "10"
                Else
                    .識別子 = "11"
                End If
            Else
                .識別子 = RTrim(iB1_Shiki)
            End If
        End If
        .契約先 = RTrim(iH1_Torcd)
        .品番 = RTrim(iB1_Hinbn)
        .背番号 = RTrim(iB1_Seban)
        .仕入先 = RTrim(iB1_Siire)
        .受入 = RTrim(iB1_Ukeir)
        .車型 = RTrim(iB1_Syakei)
        .納入番号 = RTrim(iB1_NouNo)
        .検収日 = RTrim(iB1_KYmd)
        .実績区分 = 1
        .実績処理日 = Format(Now(), "yyyymmdd")
        If OpH2_INKbn(0).Value Then
            .赤黒区分 = 0
        Else
            .赤黒区分 = B1Chk_Aka.Value
        End If
        .単価 = CisFun.Val2(iB1_Tanka)
        If B1Chk_Kari.Value = 0 Then
            .仮単価区分 = ""
        Else
            .仮単価区分 = "1"
        End If
        If B1Chk_Aka.Value = 0 Then
            .数量 = CisFun.Val2(iB1_Suryo)
            .金額 = CisFun.Val2(iB1_Kingk)
        Else
            .数量 = CisFun.Val2(iB1_Suryo) * -1
            .金額 = CisFun.Val2(iB1_Kingk) * -1
        End If
        .訂正理由 = RTrim(iB1_TKbn)
        .計算方法 = SaveKHoho
        .計算単位 = SaveKeisan
        .締日 = SaveShime
    End With

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 追加処理 )
    If SyoriKB = "ADD" Then
'        KJS.作成者 = gTanto
        Call KJSInsert
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
    Else
' ( 更新処理 )
'        KJS.更新者 = gTanto
        Call KJSUpdate_Proc
        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End If
'

DBPut_Ed:
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

'Grid再表示
    Call vsGrid1_ReSetDisp
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
Public Function KJSUpdate_Proc() As Boolean
    KJSUpdate_Proc = False
'購入実績テーブル更新
    With KJS
        gSL_Select = "Update 購入実績テーブル  Set "
        gSL_Select = gSL_Select & "識別子 = '" & RTrim(.識別子) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "車型 = '" & RTrim(.車型) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "検収日 = '" & RTrim(.検収日) & "',"
        gSL_Select = gSL_Select & "実績区分 = 1,"
        gSL_Select = gSL_Select & "実績処理日 = '" & RTrim(.実績処理日) & "',"
        gSL_Select = gSL_Select & "赤黒区分 = " & .赤黒区分 & ","
        gSL_Select = gSL_Select & "訂正理由 = '" & .訂正理由 & "',"
        gSL_Select = gSL_Select & "単価 = " & .単価 & ","
        gSL_Select = gSL_Select & "仮単価区分 = '" & RTrim(.仮単価区分) & "',"
        gSL_Select = gSL_Select & "数量 = " & .数量 & ","
        gSL_Select = gSL_Select & "金額 = " & .金額 & ""
        gSL_Select = gSL_Select & " Where 検収管理NO = " & .検収管理NO & ""
    End With
    
    With CisDB
        .ConnectNo = 0
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    
    KJSUpdate_Proc = True
End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        ReDim SColZ(21)

        Call .FixedSet(中中, 右中, あり, 3, 0, "No")
        Call .FixedSet(中中, 中中, あり, 10, 1, "検 収 日")
        Call .FixedSet(中中, 中中, あり, 9, 2, "納入番号")
        Call .FixedSet(中中, 左中, あり, 25, 3, "品　　番")
        Call .FixedSet(中中, 中中, あり, 4, 4, "赤黒")
        Call .FixedSet(中中, 右中, あり, 7, 5, "数量")
        Call .FixedSet(中中, 中中, あり, 4, 6, "仮区")
        Call .FixedSet(中中, 右中, あり, 12, 7, "単 価")
        Call .FixedSet(中中, 右中, あり, 11, 8, "金 額")
        Call .FixedSet(中中, 左中, あり, 10, 9, "訂正/識別")
        Call .FixedSet(中中, 左中, あり, 28, 10, "仕入先")
        '*-*-*-*-* 2008/04/11 Update Start
        'Call .FixedSet(中中, 中中, あり, 6, 11, "背番号")
        Call .FixedSet(中中, 中中, あり, 8, 11, "背番号")
        '*-*-*-*-* 2008/04/11 Update ENd
        Call .FixedSet(中中, 中中, あり, 4, 12, "受入")
        '*-*-*-*-* 2008/04/11 Update Start
        'Call .FixedSet(中中, 中中, あり, 4, 13, "車型")
        Call .FixedSet(中中, 中中, あり, 5, 13, "車型")
        '*-*-*-*-* 2008/04/11 Update End
        Call .FixedSet(中中, 中中, あり, 6, 14, "入力")
        
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 中中, あり, 0, 15, "検収管理NO")
        Call .FixedSet(中中, 中中, あり, 0, 16, "X検収日")
        Call .FixedSet(中中, 中中, あり, 0, 17, "X仕入先")
        Call .FixedSet(中中, 中中, あり, 0, 18, "X仮単価")
        Call .FixedSet(中中, 左中, あり, 0, 19, "X訂正理由")
        Call .FixedSet(中中, 左中, あり, 0, 20, "X識別子")
        Call .FixedSet(中中, 左中, あり, 0, 21, "X訂正管理NO")
        Call .FixedSet(中中, 左中, あり, 0, 22, "X品番")
        Call .FixedSet(中中, 左中, あり, 0, 23, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp

    End With

End Sub
'****************************************************************************************'
'*      グリッド内容セット
'****************************************************************************************'
Private Sub Grid1Set()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
    Dsp_Kensu = Format(Ttl_Kensu, "#,###")
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid1.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
            .Rows = CisVsGrid1.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    
    ID = 0
    
    Do Until Not KJS_RDSTS
        With VsGrid1
            ID = ID + 1
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            ' 検収日
            If Trim(KJS.検収日) <> "" Then
               With CisFun
                    .DateE_BefVal = KJS.検収日
                    .DateEdit
                    gStr = .DateE_AfVal
               End With
               .TextMatrix(ID, GridCol(1)) = gStr
            End If
            '納入番号
            .TextMatrix(ID, GridCol(2)) = RTrim(KJS.納入番号)
            '品番
            If RTrim(KJS.表示品番) <> "" Then
                .TextMatrix(ID, GridCol(3)) = RTrim(KJS.表示品番)
            Else
                .TextMatrix(ID, GridCol(3)) = RTrim(KJS.品番)
            End If
            '赤黒
            If KJS.赤黒区分 <> 0 Then
                .TextMatrix(ID, GridCol(4)) = "赤"
            Else
                .TextMatrix(ID, GridCol(4)) = "黒"
            End If
            ' 数量
            If KJS.数量 <> 0 Then
                If KJS.数量 < 0 Then
                    .TextMatrix(ID, GridCol(5)) = Format(KJS.数量 * -1, "#,###")
                Else
                    .TextMatrix(ID, GridCol(5)) = Format(KJS.数量, "#,###")
                End If
            End If
            '仮区分
            If RTrim(KJS.仮単価区分) = "" Then
                .TextMatrix(ID, GridCol(6)) = "正式"
            Else
                .TextMatrix(ID, GridCol(6)) = "仮"
            End If
            ' 単価
            If KJS.単価 <> 0 Then
                .TextMatrix(ID, GridCol(7)) = Format(KJS.単価, "#,##0.00")
            End If
            ' 金額
            If KJS.金額 <> 0 Then
                If KJS.金額 < 0 Then
                    .TextMatrix(ID, GridCol(8)) = Format(KJS.金額 * -1, "#,###")
                Else
                    .TextMatrix(ID, GridCol(8)) = Format(KJS.金額, "#,###")
                End If
            End If
            If KJS.識別子 = "10" Or KJS.識別子 = "11" Then
                '訂正理由
                .TextMatrix(ID, GridCol(9)) = RTrim(KJS.訂正理由名)
            Else
                '識別子
                .TextMatrix(ID, GridCol(9)) = RTrim(KJS.識別子名)
            End If
            '背番号
            .TextMatrix(ID, GridCol(11)) = RTrim(KJS.背番号)
            '仕入先・名
            If RTrim(KJS.仕入先) <> "" Then
                .TextMatrix(ID, GridCol(10)) = RTrim(KJS.仕入先) & ":" & RTrim(KJS.仕入先名)
            End If
            '受入
            .TextMatrix(ID, GridCol(12)) = RTrim(KJS.受入)
            '車型
            .TextMatrix(ID, GridCol(13)) = RTrim(KJS.車型)
            '入力種別
            If KJS.訂正管理NO <> 0 Then
                .TextMatrix(ID, GridCol(14)) = "伝票"
            Else
                .TextMatrix(ID, GridCol(14)) = "入力"
            End If
            '検収管理NO
            .TextMatrix(ID, GridCol(15)) = KJS.検収管理NO
            '検収日
            .TextMatrix(ID, GridCol(16)) = RTrim(KJS.検収日)
            '仕入先
            .TextMatrix(ID, GridCol(17)) = RTrim(KJS.仕入先)
            '仮単価区分
            .TextMatrix(ID, GridCol(18)) = RTrim(KJS.仮単価区分)
            '訂正理由
            .TextMatrix(ID, GridCol(19)) = RTrim(KJS.訂正理由)
            '識別子
            .TextMatrix(ID, GridCol(20)) = RTrim(KJS.識別子)
            '訂正管理NO
            .TextMatrix(ID, GridCol(21)) = RTrim(KJS.訂正管理NO)
            '品番
            .TextMatrix(ID, GridCol(22)) = RTrim(KJS.品番)
            
        End With
        
        Call KJSReadNext
    Loop
    
    Call KJSClose
    
    With VsGrid1
        .Row = 1: .Col = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  グリッドより入力域へセット
'+---------------------------------------------------------------------------------------+
Private Sub GridToText()
    
    With VsGrid1
        If Trim(.TextMatrix(.Row, CisVsGrid1.FixedGet(15))) = "" Then Exit Sub
        Select Case RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(20)))
            Case "10", "11": OpH2_INKbn(1).Value = True: Call OpH2_INKbn_GotFocus(1)
            Case Else: OpH2_INKbn(2).Value = True: Call OpH2_INKbn_GotFocus(2)
        End Select
        
        
        iB1_KYmd = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(16)))
        iB1_NouNo = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(2)))
        iB1_Hinbn = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(22)))
        iB1_Seban = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(11)))
        iB1_Syakei = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(13)))
        iB1_Siire = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(17)))
        Call iB1_Siire_LostFocus
        iB1_Ukeir = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(12)))
        iB1_TKbn = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(19)))
        Call iB1_TKbn_LostFocus
        iB1_Shiki = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(20)))
        Call iB1_Shiki_LostFocus
        If RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(4))) = "黒" Then
            B1Chk_Aka.Value = 0
        Else
            B1Chk_Aka.Value = 1
        End If
        iB1_Suryo = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(5)))
        
        If RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(18))) = "" Then
            B1Chk_Kari.Value = 0
        Else
            B1Chk_Kari.Value = 1
        End If
        iB1_Tanka = CisFun.RSetFld(.TextMatrix(.Row, CisVsGrid1.FixedGet(7)), 12, "#,##0.00")
        iB1_Kingk = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(8)))
    End With

End Sub
'+---------------------------------------+
'+  グリッド (Click) サイズ情報
'+---------------------------------------+
Private Sub VSGrid1_Click()
    With VsGrid1
        If ProcHB = "B2" Then
            Call HeadBodyClear("B1")

            Call GridToText
        End If
    End With
End Sub
'+---------------------------------------+
'+  グリッド (DblClick)
'+---------------------------------------+
Private Sub vsGrid1_DblClick()
    
    With VsGrid1
        If ProcHB = "B2" Then
            If Trim(.TextMatrix(.Row, CisVsGrid1.FixedGet(15))) <> "" Then
               Call ReturnPress
            End If
        End If
    End With
End Sub
'+---------------------------------------+
'+  グリッド (GotFocus)
'+---------------------------------------+
Private Sub vsGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace And Trim(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(15))) <> "" Then
       Call ReturnPress
    End If
End Sub
'+---------------------------------------+
'+  グリッド (RowColChange)
'+---------------------------------------+
Private Sub vsGrid1_RowColChange()
    If Not FormAct Then Exit Sub
    If ProcHB <> "B2" Then Exit Sub
    If RowColFlg Then Exit Sub
    Call VSGrid1_Click
End Sub
'+---------------------------------------+
'+  グリッド 再表示
'+---------------------------------------+
Private Sub vsGrid1_ReSetDisp()
    
    SaveRow = VsGrid1.Row
    RowColFlg = True
    Call DBInput("INV")
    RowColFlg = False
    If SaveRow > VsGrid1.Rows - 1 Then
        SaveRow = SaveRow - 1
    End If
    If SyoriKB <> "ADD" Then
        VsGrid1.Row = SaveRow
    End If
    If VsGrid1.RowIsVisible(SaveRow) = False Then
        VsGrid1.TopRow = SaveRow
    End If

End Sub
'+---------------------------------------+
'+      名称マスタ初期設定
'+---------------------------------------+
Private Sub MeisyoInit()

' 訂正理由
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "買掛訂正理由"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "単価訂正"
        .制御区分 = "X"
        .作成者 = gTanto
        .検索名 = "買掛訂正理由検索"
        .コード名 = "CD"
        .名称名 = "買掛訂正理由"
    End With
    Call SYMInitCreate(True)
    With SYMCreate
        .値 = "2"
        .値名称 = "金額訂正"
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "3"
        .値名称 = "数量訂正"
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "4"
        .値名称 = "型費訂正"
    End With
    Call SYMInitCreate(False)

' 識別子
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "買掛識別子"
        .区分桁数 = 2
        .値 = "01"
        .値名称 = "部品費"
        .英数字1 = "K"
        .数字1 = 0
        .制御区分 = "X"
        .作成者 = gTanto
        
        .検索名 = "識別子(買掛)　検索"
        .コード名 = "識別子"
        .名称名 = "識別子名称"
    End With
    Call SYMInitCreate(True)
    With SYMCreate
        .値 = "02"
        .値名称 = "型費"
        .英数字1 = "K"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "04"
        .値名称 = "その他部品費"
        .英数字1 = "K"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "05"
        .値名称 = "その他仕入"
        .英数字1 = "K"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "06"
        .値名称 = "金型"
        .英数字1 = "K"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "07"
        .値名称 = "金型修繕費"
        .英数字1 = "K"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "08"
        .値名称 = "治工具"
        .英数字1 = "K"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "10"
        .値名称 = "訂正(黒)"
        .英数字1 = "K"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "11"
        .値名称 = "訂正(赤)"
        .英数字1 = "K"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "12"
        .値名称 = "遡り訂正"
        .英数字1 = "K"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)

    With SYMCreate
        .値 = "21"
        .値名称 = "支給部品費"
        .英数字1 = "S"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "24"
        .値名称 = "その他部品費"
        .英数字1 = "S"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "25"
        .値名称 = "その他支給"
        .英数字1 = "S"
        .数字1 = 1
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "30"
        .値名称 = "支給訂正(黒)"
        .英数字1 = "S"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "31"
        .値名称 = "支給訂正(赤)"
        .英数字1 = "S"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)
    With SYMCreate
        .値 = "32"
        .値名称 = "支給遡り訂正"
        .英数字1 = "S"
        .数字1 = 0
    End With
    Call SYMInitCreate(False)


End Sub




