VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form COK0410 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "買掛実績問合せ"
   ClientHeight    =   11025
   ClientLeft      =   90
   ClientTop       =   1485
   ClientWidth     =   15195
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11025
   ScaleWidth      =   15195
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   5040
      Top             =   3765
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
   Begin VB.PictureBox B1_Area1 
      Height          =   8985
      Left            =   30
      ScaleHeight     =   8925
      ScaleWidth      =   15060
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1845
      Width           =   15120
      Begin vsFlexLib.vsFlexArray VSGrid1 
         DragIcon        =   "COK0410.frx":0000
         Height          =   8670
         Left            =   45
         TabIndex        =   4
         Top             =   165
         Width           =   15000
         _Version        =   196608
         _ExtentX        =   26458
         _ExtentY        =   15293
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
         FormatString    =   "NO   |実績年月|契約先|車型|品番|納入番号|数量|金額|"
         Rows            =   26
         Cols            =   9
         BackColor       =   12648447
         BackColorBkg    =   12632256
         HighLight       =   0
         RowHeightMin    =   330
         SelectionMode   =   1
         BackColorAlternate=   12648447
      End
   End
   Begin VB.PictureBox H1_Area1 
      Height          =   1290
      Left            =   30
      ScaleHeight     =   1230
      ScaleWidth      =   10770
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   465
      Width           =   10830
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   1140
         Left            =   7350
         Top             =   45
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   2011
         ForeColor       =   16711680
         Caption         =   "品番"
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
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   135
            TabIndex        =   3
            Top             =   510
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   661
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   1140
         Left            =   60
         Top             =   45
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   2011
         ForeColor       =   16711680
         Caption         =   ""
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
         cPositionY      =   60
         Begin VB.OptionButton OpH1_KKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "支給実績"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   105
            TabIndex        =   13
            Top             =   600
            Width           =   1365
         End
         Begin VB.OptionButton OpH1_KKbn 
            BackColor       =   &H00C0FFC0&
            Caption         =   "購入実績"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   105
            TabIndex        =   12
            Top             =   210
            Value           =   -1  'True
            Width           =   1365
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   1140
         Left            =   3720
         Top             =   45
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   2011
         ForeColor       =   16711680
         Caption         =   "契約先"
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
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   45
            TabIndex        =   2
            Top             =   510
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
         Begin Cis3D_v60.CIS3D H1lb_Torcd 
            Height          =   360
            Left            =   1035
            Top             =   510
            Width           =   2460
            _ExtentX        =   4339
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
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   1140
         Index           =   0
         Left            =   1620
         Top             =   45
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   2011
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
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisYM_V60.CisYM iH1_SYM 
            Height          =   375
            Left            =   570
            TabIndex        =   0
            Top             =   360
            Width           =   1455
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
         Begin Cis3D_v60.CIS3D PB_SDate 
            Height          =   345
            Left            =   90
            Top             =   360
            Width           =   465
            _ExtentX        =   820
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
            Left            =   90
            Top             =   720
            Width           =   465
            _ExtentX        =   820
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
         Begin CisYM_V60.CisYM iH1_EYM 
            Height          =   375
            Left            =   570
            TabIndex        =   1
            Top             =   705
            Width           =   1455
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
   Begin Cis3D_v60.CIS3D UC_3D31 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "　　　【 買掛実績問合せ 】"
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
         Height          =   315
         Left            =   10245
         ScaleHeight     =   255
         ScaleWidth      =   1590
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   45
         Width           =   1650
         Begin VB.OptionButton Op_Overlap 
            BackColor       =   &H00C0C0C0&
            Caption         =   "明細"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   45
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   705
         End
         Begin VB.OptionButton Op_Overlap 
            BackColor       =   &H00C0C0C0&
            Caption         =   "集計"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   780
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Value           =   -1  'True
            Width           =   765
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D32 
         Height          =   255
         Left            =   7260
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
   Begin VB.TextBox Dummy 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   3  'ｵﾌ固定
      Left            =   240
      TabIndex        =   5
      Text            =   "D"
      Top             =   210
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   510
      Top             =   210
   End
   Begin VB.Timer Timer2 
      Left            =   1230
      Top             =   210
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   705
      Left            =   11400
      ScaleHeight     =   645
      ScaleWidth      =   3480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   450
      Width           =   3540
      Begin CisBtn_60.CisBtn PB_END 
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
         Left            =   1770
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
         Left            =   930
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   45
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
   End
End
Attribute VB_Name = "COK0410"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  買掛実績問合せ
'**       フォームID    :  COK0410
'**       処理概要      :
'**
'**       作  成  日    :  2004/12/13  By CIS
'**       変  更  日    :  2005/03/04  By CIS - 重複起動ﾁｪｯｸで起動済の時、『実行時エラー'91'』を表示する対応（Timerﾌﾟﾛﾊﾟﾃｨの Interval=0 に修正）
'**       変  更  日    :  2008/01/24  By CIS - 受給追加
'**                        2008/04/11  By CIS   車型5桁
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim wObj            As Object       ' ｵﾌﾞｼﾞｪｸﾄﾜｰｸｴﾘｱ
    Dim Sl_Where        As String       ' SQL条件
    Dim Ttl_Kensu       As Integer      ' 表示件数
    Dim mDispMax        As Integer      ' ｸﾞﾘｯﾄﾞ全行数退避
    
    Dim CisVsGrid       As New CisVsGrid3
    Dim mGotColor       As Long
    Dim mLostColor      As Long
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
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = VSGrid1
        .NoSpaceCol = 1
        .NumberSet = No
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 10
        .InitGet = False
        .Init
    End With
    
    Call GridInit
    
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Call CisFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
    
    mDispMax = CisVsGrid.DispMax
    lb_Kensu = ""

    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
'
    mGotColor = OpH1_KKbn(0).BackColor
    mLostColor = OpH1_KKbn(1).BackColor
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    Op_Overlap(0).Value = True
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
' 買掛実績年月を求める
    If MeisyoGet("処理内容", "POK0110") Then
       iH1_SYM = RTrim(SYM.英数字1)
    End If

    OpH1_KKbn(0).SetFocus
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False
    If Not FormAct Then Exit Function
    If UCase(IVENT) <> "DOWN" Then Exit Function
    
    Select Case Key_Code
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
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
    If ProcHB = "H1" Then
        Call HeadClear
        iH1_SYM.SetFocus
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B1" Then
        Call DispTotal(False)
        lb_Kensu = ""
        CisVsGrid.DispMax = mDispMax
        VSGrid1.Rows = mDispMax
        CisVsGrid.Clear
        Call GridInit
        ProcHB = "H1"
        Call DispChange(ProcHB)
        iH1_SYM.SetFocus
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検索   (F7)             +
'+----------------------------+
Private Sub PB_Look_Click()
    If ProcHB = "H1" Then
       If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
       If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    End If
End Sub
'+----------------------------+
'+      検         索         +
'+----------------------------+
Private Sub Look_Torcd()

    RV_TorKb = 5    ' 契約先
    RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    iH1_Torcd.SetFocus

    If RV_Rtn Then
       iH1_Torcd = RV_TorcdK
       Call iH1_Torcd_LostFocus
    End If

End Sub
'+----------------------------+
'+       品番検索            +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Left = 0
    RV_Top = 0
    RVI_Hinmoku = 1
    
    CKK0030.Show vbModal
    Unload CKK0030
    Set CKK0030 = Nothing

    If RV_Rtn Then
       iH1_Hinbn = RV_Hinbn
    End If
End Sub
'+----------------------------------------+
'+  iH1_Torcd ( GotFocus )   契約先
'+----------------------------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
'+----------------------------------------+
'+  iH1_Torcd ( LostFocus )  契約先
'+----------------------------------------+
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
       PB_Look.Visible = False
       PB_Look.Tag = ""
    End If
    If TorNmGet(iH1_Torcd, 5) Then
        H1lb_Torcd = TRM_RName
        If Trim(H1lb_Torcd) = "" Then
           H1lb_Torcd = TRM_Name
        End If
    Else
        H1lb_Torcd = ""
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
'+  OpH1_KKbn
'+-----------------------------+
Private Sub OpH1_KKbn_GotFocus(Index As Integer)
    OpH1_KKbn(Index).BackColor = mGotColor
End Sub
Private Sub OpH1_KKbn_LostFocus(Index As Integer)
    OpH1_KKbn(Index).BackColor = mLostColor
End Sub
'+----------------------------------------+
'+      小計 表示･非表示 切替              +
'+----------------------------------------+
Private Sub Op_Overlap_Click(Index As Integer)
    Call CisFrm.MousePT(11)

    If ProcHB = "B1" Then
        CisVsGrid.DispMax = mDispMax
        VSGrid1.Rows = mDispMax
        CisVsGrid.Clear
        Call GridInit
        Call BodySet
    Else
        If Op_Overlap(0).Value Then
            Call DispTotal(True)
        Else
            Call DispTotal(False)
        End If
    End If

    On Error Resume Next
    If VSGrid1.Enabled Then
        VSGrid1.SetFocus
    Else
        iH1_SYM.SetFocus
    End If
    On Error GoTo 0
    
    Call CisFrm.MousePT(1)
    
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

'    PB_CAN.Visible = Not H1Mode
    PB_ENT.Visible = H1Mode
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
   Call CisFrm.MousePT(11)
   If ProcHB = "H1" Then
      ProcHB = "B1"
      If Not Head1Chk Then
          ProcHB = "H1"
          GoTo ReturnPress_Ed
      End If
      Call DispChange(ProcHB)
      VSGrid1.SetFocus
      GoTo ReturnPress_Ed
   End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1初期化
'****************************
Private Sub HeadClear()
    For Each wObj In Me.Controls
        If wObj.Name Like "iH*" Or _
           wObj.Name Like "H1lb_*" Then wObj = ""
        If wObj.Name Like "Ch_*" Then wObj.Value = 0
    Next wObj
    Op_Overlap(0).Value = True
End Sub

'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
 
' 印刷年月
    If Trim(iH1_SYM) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "     開始年月を入力して下さい      "
            .MB_MSG(3) = "   "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SYM.SetFocus
        Exit Function
    End If
    If Not IsDate(iH1_SYM.cYear & "/" & iH1_SYM.cMonth & "/01") Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "        開始年月             "
            .MB_MSG(4) = "   正しい年月にて入力して下さい        "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SYM.SetFocus
        Exit Function
    End If
' 年月【終了】
    If Trim(iH1_EYM) <> "" Then
        If Not IsDate(iH1_EYM.cYear & "/" & iH1_EYM.cMonth & "/01") Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "        終了年月              "
                .MB_MSG(4) = "   正しい年月にて入力して下さい        "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_EYM.SetFocus
            Exit Function
        End If
    End If
    If Trim(iH1_SYM) <> "" And Trim(iH1_EYM) <> "" Then
        If iH1_SYM > iH1_EYM Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "     年月【開始＞終了】        "
                .MB_MSG(4) = ""
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_SYM.SetFocus
            Exit Function
        End If
    End If

    If Not BodySet Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    該当データが存在しません｡            "
            .MB_MSG(3) = "    再度抽出条件を入力してください        "
            .MB_Button = Error
            .MB_Title = "抽出処理"
            Call .MBOX
       End With
       PNL_Wait.Visible = False
       Exit Function
    End If
    Head1Chk = True
End Function
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear ("NO")
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容)
        Call .FixedSet(中中, 右中, なし, 7, 0, "№")
        Call .FixedSet(中中, 中中, あり, 8, 1, "実績年月")
        '==============================================***2008/06/23 Start
        'Call .FixedSet(中中, 左中, あり, 27, 2, "契約先")
        Call .FixedSet(中中, 左中, あり, 20, 2, "契約先")
        '==============================================***2008/06/23 End
        '*-*-*-*-* 2008/04/11 Update Start
        'Call .FixedSet(中中, 中中, あり, 5, 3, "車型")
        '==============================================***2008/06/23 Start
        'Call .FixedSet(中中, 中中, あり, 7, 3, "車型")
        Call .FixedSet(中中, 中中, あり, 6, 3, "車型")
        '==============================================***2008/06/23 End
        '*-*-*-*-* 2008/04/11 Update End
        '==============================================***2008/06/23 Start
        'Call .FixedSet(中中, 左中, あり, 27, 4, "品番")
        Call .FixedSet(中中, 左中, あり, 40, 4, "品番 / 材料")
        '==============================================***2008/06/23 End
        Call .FixedSet(中中, 左中, なし, 8, 5, "納入番号")
        Call .FixedSet(中中, 右中, なし, 10, 6, "数量")
        Call .FixedSet(中中, 右中, なし, 16, 7, "金額")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 8, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7")
        .InitDisp
    End With
End Sub
'****************************
'*      ボディ内容セット
'****************************
Private Function BodySet() As Boolean
    Dim ID          As Long
    Dim ID2         As Long
    Dim wChk        As Boolean
    Dim GridCol()   As Long

    BodySet = False

    On Error GoTo BodySet_Err:
    
    Sl_Where = "": gSL_Select = ""
' 表示条件 セット
    With CisFun
        .Where_CreateKB = NewDocument
        If Trim(iH1_EYM) <> "" Then
           If iH1_SYM = iH1_EYM Then
               .Where_Create AlphaMe, "KJ.実績年月", iH1_SYM
           Else
               .Where_Create AlphaMe, "KJ.実績年月", iH1_SYM, Between, , iH1_EYM
           End If
        Else
          .Where_Create AlphaMe, "KJ.実績年月", iH1_SYM
        End If
         ' 契約先
         .Where_Create AlphaMe, "KJ.契約先", iH1_Torcd
         ' 品番
         .Where_Create AlphaMe, "KJ.品番", iH1_Hinbn
    End With

'=================================================================== ***2008/06/23 Start
    CisDB.SQL = "DELETE FROM 買掛実績問合ワーク"
    CisDB.SQL = CisDB.SQL & "  WHERE 処理端末 = SUBSTRING(HOST_NAME(), 1, 20)"
    Call CisDB.DBExec
    
    gSL_Select = " INSERT INTO 買掛実績問合ワーク "
'=================================================================== ***2008/06/23 End
    
    gSL_Select = gSL_Select & " SELECT SUBSTRING(HOST_NAME(), 1, 20),SUBSTRING(KJ.実績年月,1,4) + '/' + SUBSTRING(KJ.実績年月,5,2) 実績年月," & vbCrLf
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(TR.略称,'') <> '' THEN " & vbCrLf
    gSL_Select = gSL_Select & "      KJ.契約先 + ':' + TR.略称 " & vbCrLf
    gSL_Select = gSL_Select & " ELSE KJ.契約先 END 契約先名," & vbCrLf
    gSL_Select = gSL_Select & " KJ.契約先," & vbCrLf
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(HN.表示品番,'') = '' THEN KJ.品番 Else HN.表示品番 End 品番," & vbCrLf
    gSL_Select = gSL_Select & " 納入番号,KJ.数量, KJ.金額," & vbCrLf
    gSL_Select = gSL_Select & " CASE WHEN KJ.入力区分 = '' THEN HN.車型 ELSE KJ.車型 END 車型" & vbCrLf
    If OpH1_KKbn(0).Value Then
        gSL_Select = gSL_Select & " From 購入実績テーブル KJ" & vbCrLf
    Else
        gSL_Select = gSL_Select & " From 支給実績テーブル KJ" & vbCrLf
    End If
    gSL_Select = gSL_Select & "  Left Outer Join 品番マスタ HN" & vbCrLf
    gSL_Select = gSL_Select & "   ON HN.品番 = KJ.品番" & vbCrLf
    gSL_Select = gSL_Select & "    LEFT Outer Join 取引先マスタ TR" & vbCrLf
    gSL_Select = gSL_Select & "    ON TR.取引先CD = KJ.契約先" & vbCrLf
    gSL_Select = gSL_Select & "    And TR.取引先区分 = 1" & vbCrLf
'=================================================================== ***2008/06/23 Start
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    
    gSL_Select = gSL_Select & " UNION " & vbCrLf
    gSL_Select = gSL_Select & " SELECT SUBSTRING(HOST_NAME(), 1, 20),SUBSTRING(KJ.実績年月,1,4) + '/' + SUBSTRING(KJ.実績年月,5,2) 実績年月," & vbCrLf
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(TR.略称,'') <> '' THEN " & vbCrLf
    gSL_Select = gSL_Select & "      KJ.契約先 + ':' + TR.略称 " & vbCrLf
    gSL_Select = gSL_Select & " ELSE KJ.契約先 END 契約先名," & vbCrLf
    gSL_Select = gSL_Select & " KJ.契約先," & vbCrLf
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(ZN.表示寸法,'') = '' THEN RTRIM(ZN.材質) + ' ' + CONVERT(VARCHAR(7),ZN.板厚) +  CONVERT(VARCHAR(7),ZN.幅) + CONVERT(VARCHAR(7),ZN.長さ)" & vbCrLf
    gSL_Select = gSL_Select & " Else RTRIM(ZN.材質) + ' ' + ZN.表示寸法 End 品番," & vbCrLf
    gSL_Select = gSL_Select & " 納入番号,KJ.数量, KJ.金額," & vbCrLf
    gSL_Select = gSL_Select & " '' 車型" & vbCrLf
    If OpH1_KKbn(0).Value Then
        gSL_Select = gSL_Select & " From 材料購入実績テーブル KJ" & vbCrLf
    Else
        gSL_Select = gSL_Select & " From 材料支給実績テーブル KJ" & vbCrLf
    End If
    gSL_Select = gSL_Select & "  Left Outer Join 材料マスタ ZN" & vbCrLf
    gSL_Select = gSL_Select & "   ON ZN.材料管理番号 = KJ.材料管理番号" & vbCrLf
    gSL_Select = gSL_Select & "    LEFT Outer Join 取引先マスタ TR" & vbCrLf
    gSL_Select = gSL_Select & "    ON TR.取引先CD = KJ.契約先" & vbCrLf
    gSL_Select = gSL_Select & "    And TR.取引先区分 = 1" & vbCrLf

'=================================================================== ***2008/06/23 End

' 抽出文　作成
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    
'=================================================================== ***2008/06/23 Start
    CisDB.SQL = ""
    CisDB.SQL = gSL_Select
    Call CisDB.DBExec

    gSL_Select = "SELECT * FROM 買掛実績問合ワーク KJ"
    gSL_Select = gSL_Select & "  WHERE 処理端末 = SUBSTRING(HOST_NAME(), 1, 20)"
'=================================================================== ***2008/06/23 End
    
    gSL_Select = gSL_Select & " ORDER BY KJ.実績年月,KJ.契約先,KJ.車型,KJ.品番,KJ.納入番号 "
    
    If Not KJSRead(gSL_Select) Then GoTo BodySet_ED1

    Ttl_Kensu = CisDB.RecordCount
    lb_Kensu.Caption = Format(Ttl_Kensu, "#,###,##0")
    Back_Over.Enabled = False    ' 分類 入力禁止

    PNL_Wait.Visible = True      ' 抽出中ﾒｯｾｰｼﾞ表示
    DoEvents

' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VSGrid1
        .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(ID)
        Next ID
    End With
    ID = 1
    VSGrid1.Rows = CisDB.RecordCount + 1

    Do Until Not KJS_RDSTS
        With VSGrid1
            If Op_Overlap(1).Value Then
                .TextMatrix(ID, GridCol(0)) = CisFun.RSetFld(ID, 3, "###")
            End If
            If Trim(KJS.実績年月) = "" Then
                .TextMatrix(ID, GridCol(1)) = "    "
            Else
                .TextMatrix(ID, GridCol(1)) = KJS.実績年月
            End If
            If Trim(KJS.契約先名) = "" Then
                .TextMatrix(ID, GridCol(2)) = "    "
            Else
                .TextMatrix(ID, GridCol(2)) = KJS.契約先名
            End If
            If Trim(KJS.車型) = "" Then
                .TextMatrix(ID, GridCol(3)) = "    "
            Else
                .TextMatrix(ID, GridCol(3)) = KJS.車型
            End If
            If Trim(KJS.品番) = "" Then
                .TextMatrix(ID, GridCol(4)) = "    "
            Else
                .TextMatrix(ID, GridCol(4)) = KJS.品番
            End If
            If Trim(KJS.納入番号) = "" Then
                .TextMatrix(ID, GridCol(5)) = "    "
            Else
                .TextMatrix(ID, GridCol(5)) = KJS.納入番号
            End If
            .TextMatrix(ID, GridCol(6)) = CisFun.RSetFld(KJS.数量, 7, "###,###")
            .TextMatrix(ID, GridCol(7)) = CisFun.RSetFld(KJS.金額, 10, "##,###,###")
            ID = ID + 1
        End With
        Call KJSReadNext
    Loop
    Call KJSClose

    CisVsGrid.DispMax = VSGrid1.Rows

BodySet_ED:

    Call GridInit
    If Op_Overlap(0).Value Then
        Call DispTotal(True)
    Else
        Call DispTotal(False)
    End If
    PNL_Wait.Visible = False
    With VSGrid1
         .Col = 1: .ColSel = .Cols - 1
         .Row = 1: .Redraw = True
    End With
    BodySet = True

BodySet_ED1:
    Back_Over.Enabled = True    ' 分類 入力禁止 解除
    PNL_Wait.Visible = False   ' 抽出中ﾒｯｾｰｼﾞ非表示
    Exit Function
BodySet_Err:
     CisFun.ErrorBox ("BodySet")
End Function
'+----------------------------------------+
'+      数量・金額 小計設定                 +
'+----------------------------------------+
Private Sub DispTotal(Disp As Boolean)
    Static FlexSt   As Integer
    Static FixedCol As Long
    
    With VSGrid1
        Call CisVsGrid.VSSort(True)
        
        .SubtotalPosition = flexSTAbove
        FlexSt = flexSTClear
        If Not Disp Or ProcHB = "H1" Then
            .OutlineBar = flexOutlineBarNone
            .Outline -1
            .Subtotal flexSTClear
        Else
            .Subtotal flexSTClear
            FlexSt = flexSTSum
        End If
        
        ' 数量・型費・印額　合計
        For gLong = 6 To 7
            FixedCol = CisVsGrid.FixedGet(gLong)
            .Subtotal FlexSt, 5, FixedCol, ",", &HFFFFC0, , False, " %s", 1
            .Subtotal FlexSt, 4, FixedCol, ",", &HC0FFC0, , False, " %s", 1
            .Subtotal FlexSt, 3, FixedCol, ",", &HC0C0FF, , False, " %s", 1
            .Subtotal FlexSt, 2, FixedCol, ",", &HC0FFC0, , False, " %s", 1
            .Subtotal FlexSt, 1, FixedCol, ",", &HC0C0FF, , False, " %s", 1
            .Subtotal FlexSt, -1, FixedCol, ",", &H80000008, &HFFFFFF, False, "総計"
        Next gLong

        If Not Disp Or ProcHB = "H1" Then
            .Outline -1
        Else
            .OutlineBar = flexOutlineBarComplete
            .Outline 2
        End If
    End With

'    If Not Disp Then
'        CisVsGrid.DispMax = mDispMax
'        VSGrid1.Rows = mDispMax
'    End If
End Sub
'******************************************'
'*****       移動選択項目選択          *****
'******************************************'
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Button : 1 = 左ﾎﾞﾀﾝ ,2 = 右ﾎﾞﾀﾝ
    With VSGrid1
        .Tag = ""
        If Button <> 1 Then Exit Sub
    ' 固定行以外は行わない
        If .MouseRow <> 0 Then Exit Sub
    ' 固定列は行わない
        If .MouseCol = 0 Then Exit Sub
    ' 数量・金額は行わない
        If .MouseCol >= 6 Then Exit Sub
    ' 移動元列をタグへ退避
        .Tag = Str(.MouseCol)
        .Drag 1
    End With
End Sub
'******************************************'
'*****        表 示 項 目 移 動        *****
'******************************************'
Private Sub VSGrid1_DragDrop(Source As Control, X As Single, Y As Single)

    With VSGrid1
    ' 固定列は行わない
        If .MouseCol = 0 Then Exit Sub
        If .Tag = "" Then Exit Sub
    ' 数量・金額より前のみ移動可能
        If .MouseCol >= 6 Then Exit Sub
' ｸﾞﾘｯﾄﾞ表示 禁止
        Call CisFrm.MousePT(11)

        .Redraw = False
' 列移動
        .ColPosition(CisFun.Val2(.Tag)) = .MouseCol
' 列分類
        .Col = 1
        .ColSel = .Cols - 1
        '.Sort = GridSortAsc
' 合計表示・非表示
        If Op_Overlap(0).Value Then
            Call DispTotal(True)
            If .MouseCol > 0 And .MouseCol < 6 Then
                VSGrid1.Outline .MouseCol
            End If
        Else
            Call DispTotal(False)
        End If
' ｸﾞﾘｯﾄﾞ表示 解除
        .Redraw = True

       Call CisFrm.MousePT(1)

    End With
End Sub


