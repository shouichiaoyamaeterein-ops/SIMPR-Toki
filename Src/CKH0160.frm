VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Begin VB.Form CKH0160 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "かんばん増減・平準化開始日設定"
   ClientHeight    =   11235
   ClientLeft      =   2370
   ClientTop       =   2010
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D CIS3D3 
      Height          =   825
      Left            =   3360
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1455
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
      Begin Cis3D_v60.CIS3D PB_Base 
         Height          =   675
         Left            =   2655
         Top             =   90
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1191
         Caption         =   "変更"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   2
         cAlingnment     =   7
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   705
         Left            =   60
         Top             =   60
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1244
         BackColor       =   16777152
         Caption         =   "基本開始日"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         cAlingnment     =   6
         cBoderStyle     =   1
         cPositionY      =   45
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   330
            Left            =   465
            Top             =   375
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "(１日)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   16777152
            cBoderColor2    =   16777152
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D H1lb_BaseDay 
         Height          =   705
         Left            =   1680
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1244
         BackColor       =   16777152
         Caption         =   " 1日"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         cAlingnment     =   7
         cBoderStyle     =   1
         Begin CisText_V60.CisText iH1_BaseStart 
            Height          =   360
            Left            =   120
            TabIndex        =   2
            Top             =   195
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   635
            cFormat         =   "###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###"
            MaxLength       =   3
            cDataType       =   4
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cMinusColor     =   0
            cGFormat        =   "###"
            cILength        =   3
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   7830
      Left            =   3353
      ScaleHeight     =   7770
      ScaleWidth      =   9270
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2940
      Width           =   9330
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   7635
         Index           =   1
         Left            =   75
         Top             =   75
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   13467
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   510
            Index           =   3
            Left            =   5010
            Top             =   60
            Width           =   3690
            _ExtentX        =   6509
            _ExtentY        =   900
            ForeColor       =   16711680
            Caption         =   "品     番"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   510
            Index           =   2
            Left            =   4455
            Top             =   60
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   900
            ForeColor       =   16711680
            Caption         =   "受"
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
            cAlingnment     =   7
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   510
            Index           =   1
            Left            =   780
            Top             =   60
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   900
            ForeColor       =   16711680
            Caption         =   "仕  入  先"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   510
            Index           =   0
            Left            =   75
            Top             =   60
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   900
            ForeColor       =   16711680
            Caption         =   "№"
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
            cAlingnment     =   7
         End
         Begin vsFlexLib.vsFlexArray vsGrid 
            Height          =   7470
            Left            =   60
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   105
            Width           =   8940
            _Version        =   196608
            _ExtentX        =   15769
            _ExtentY        =   13176
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
            FormatString    =   "№   |自社取引|取引先名             |   |契約先    |支給元②  |代行      |   |                     |||"
            Rows            =   17
            Cols            =   12
            BackColor       =   12648384
            BackColorBkg    =   12632256
            FocusRect       =   2
            HighLight       =   2
            RowHeightMin    =   430
            SelectionMode   =   1
            Begin Cis3D_v60.CIS3D Input_Row 
               Height          =   435
               Left            =   675
               Top             =   435
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   767
               BackColor       =   0
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
               Begin Cis3D_v60.CIS3D Back_JTorcd 
                  Height          =   420
                  Left            =   -30
                  Top             =   0
                  Width           =   4275
                  _ExtentX        =   7541
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   16711680
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
                  cBoderColor1    =   -2147483643
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
                  Begin Cis3D_v60.CIS3D B1lb_Tornm 
                     Height          =   345
                     Left            =   1065
                     Top             =   30
                     Width           =   2655
                     _ExtentX        =   4683
                     _ExtentY        =   609
                     BackColor       =   16777152
                     Caption         =   "XXXXXXXXX1XXXXXXXXXX"
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
                     cAlingnment     =   1
                     cBoderStyle     =   1
                     cPositionX      =   50
                  End
                  Begin CisText_V60.CisText iB1_Torcd 
                     Height          =   360
                     Left            =   75
                     TabIndex        =   3
                     Top             =   45
                     Width           =   960
                     _ExtentX        =   1693
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
                     cDataReplace    =   1
                     cFaZero         =   0
                  End
                  Begin CisText_V60.CisText iB1_Uke 
                     Height          =   360
                     Left            =   3765
                     TabIndex        =   4
                     Top             =   30
                     Width           =   465
                     _ExtentX        =   820
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
                     cFaZero         =   0
                  End
               End
               Begin Cis3D_v60.CIS3D Back_KTorcd 
                  Height          =   420
                  Left            =   4245
                  Top             =   0
                  Width           =   3720
                  _ExtentX        =   6562
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   16711680
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
                  cBoderColor1    =   -2147483643
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
                  Begin CisText_V60.CisText iB1_Hinbn 
                     Height          =   375
                     Left            =   75
                     TabIndex        =   5
                     Top             =   30
                     Width           =   3645
                     _ExtentX        =   6429
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
                     cDataReplace    =   1
                  End
               End
            End
         End
         Begin CisText_V60.CisText B_TabDummy 
            Height          =   165
            Left            =   5325
            TabIndex        =   10
            Top             =   180
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   291
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
      Height          =   945
      Left            =   3353
      ScaleHeight     =   885
      ScaleWidth      =   6945
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1740
      Width           =   7005
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Left            =   60
         Top             =   60
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "開始日"
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
         Begin CisText_V60.CisText iH1_StartDay 
            Height          =   360
            Left            =   165
            TabIndex        =   0
            Top             =   330
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   635
            cFormat         =   "###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###"
            MaxLength       =   3
            cDataType       =   4
            cDataReplace    =   1
            cFbComma        =   0
            cFdAutoFormat   =   1
            cMinusColor     =   0
            cGFormat        =   "###"
            cILength        =   3
            cAutoEject      =   0   'False
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   780
         Left            =   960
         Top             =   60
         Width           =   5910
         _ExtentX        =   10425
         _ExtentY        =   1376
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
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   600
            Left            =   60
            TabIndex        =   1
            Top             =   90
            Width           =   5775
            _Version        =   196608
            _ExtentX        =   10186
            _ExtentY        =   1058
            _StockProps     =   228
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "    |    |    |    |    |    |    |    |    |    "
            Rows            =   2
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            ScrollBars      =   0
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
      Height          =   750
      Left            =   7628
      ScaleHeight     =   690
      ScaleWidth      =   4320
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   4380
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1755
         Top             =   60
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
         Left            =   2595
         Top             =   60
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
         Left            =   3435
         Top             =   60
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
         Left            =   75
         Top             =   60
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   915
         Top             =   60
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
      TabIndex        =   11
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
      Caption         =   "【 かんばん増減・平準化開始日設定  】"
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
   End
End
Attribute VB_Name = "CKH0160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   かんばん増減・平準化開始日　メンテ
'**       フォームID    :   CKH0160
'**       処理概要      :
'**
'**       作  成  日    :   2006/02/20  By CIS SIMPR-A
'**       変  更  日    :   2007/02/23  By CIS 基本開始日(±設定/表示)
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    Dim CisVsGrid1      As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim ActObj          As Object
    Dim ErrObj          As Object
    Dim ActRow          As Integer
    
    Dim RD_Count        As Long
    Dim mBaseStart_Day  As Integer
    Dim mBaseHani(20)   As String

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
    
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGrid
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 430
        .PatanMax = 1
        .SelectCol = 19
        .InitGet = False
        .Init
    End With
    Call GridInit
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = ""
    Call SyoriName(SyoriKB)
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
'
'   >> 名称マスタより基本開始日を獲得
'基本開始日
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "かんばん増減基本開始日"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "1:基本開始日 2:開始日範囲"
        .英数字1 = "01"
        .英数字2 = "-05,-04,-03,-02,-01"
        .作成者 = gTanto
    End With
    Call SYMInitCreate(True)
    Call MeisyoGet("かんばん増減基本開始日", "1")
    
    mBaseStart_Day = CisFun.Val2(SYM.英数字1)
    iH1_BaseStart.Visible = False
    H1lb_BaseDay = Format(mBaseStart_Day, "+###;-###;±0") & "日"
    iH1_BaseStart = Format(mBaseStart_Day, "###")
    
    gCnt2 = 0: gCnt3 = 0
    For gCnt1 = 1 To Len(RTrim(SYM.英数字2))
        If Mid(SYM.英数字2, gCnt1, 1) = "," Then
           mBaseHani(gCnt3) = Mid(SYM.英数字2, gCnt1 - gCnt2, gCnt2)
           gCnt2 = 0
           gCnt3 = gCnt3 + 1
           If gCnt3 = 20 Then Exit For
        Else
           gCnt2 = gCnt2 + 1
        End If
    Next gCnt1
    gCnt1 = Len(RTrim(SYM.英数字2))
    mBaseHani(gCnt3) = Mid(SYM.英数字2, (gCnt1 + 1) - gCnt2, gCnt2)
'
    For gCnt1 = 0 To gCnt3
        If gCnt1 <= 10 Then
           vsGrid1.TextMatrix(0, gCnt1) = mBaseHani(gCnt1)
        Else
           vsGrid1.TextMatrix(1, gCnt1) = mBaseHani(gCnt1)
        End If
    Next gCnt1
    
    FormAct = Not FormAct
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
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
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
   Call CisFrm.MousePT(11)
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   
   If ProcHB = "H1" Then
      SyoriKB = ""
      Call SyoriName(SyoriKB)
      PB_CAN.Visible = False
      iH1_StartDay.SetFocus
      GoTo PB_CAN_Ed
   End If
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      iH1_StartDay.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)    +
'+----------------------------+
Private Sub PB_DEL_Click()
   
    If SyoriKB = "" Then
        SyoriKB = "DEL"
        PB_CAN.Visible = True
    Else
        SyoriKB = ""
        PB_CAN.Visible = False
    End If
    Call SyoriName(SyoriKB)
End Sub
Private Sub PB_Base_Click()
    If PB_Base.Caption = "変更" Then
       PB_Base.Caption = "更新"
       iH1_BaseStart.Visible = True
       iH1_BaseStart.SetFocus
    Else
       PB_Base.Caption = "変更"
       H1lb_BaseDay.Caption = Format(CisFun.Val2(iH1_BaseStart), "+###;-###;±0") & "日"
       iH1_BaseStart.Visible = False
    
    '基本開始日
       gSL_Select = "UPDATE 名称マスタ SET "
       gSL_Select = gSL_Select & " 英数字1 = '" & Format(CisFun.Val2(iH1_BaseStart), "00") & "'"
       gSL_Select = gSL_Select & " WHERE 区分名称 = 'かんばん増減基本開始日'"
       gSL_Select = gSL_Select & " AND 値 = '1'"
       gSL_Select = gSL_Select & " AND 区分タイプ = ''"
    
       With CisDB
            .ConnectNo = 0
            .SQL = gSL_Select
            .DBExec
       End With
    
    End If
    
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iB1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iB1_Uke" Then Call Look_UTorcd: Exit Sub
    If PB_Look.Tag = "iB1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_StartDay" Then Call Look_StartDay: Exit Sub
End Sub
'+----------------------------+
'+       開始日検索            +
'+----------------------------+
Private Sub Look_StartDay()

    RV_Left = 0:    RV_Top = 0
    RV_Call = "C"
    
    CKK0160.Show vbModal
    Unload CKK0160
    Set CKK0160 = Nothing

    If RV_Rtn Then
       iH1_StartDay = Format(CisFun.Val2(RV_SYmd), "###")
    End If

End Sub
'+----------------------------+
'+       取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()

    RV_Left = 0:    RV_Top = 0
    RV_TorKb = 9
    RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
       iB1_Torcd = RTrim(RV_TorcdK)
       B1lb_Tornm = RV_TorRName
    End If

End Sub
'+-----------------------------+
'+       取引先受入検索         +
'+-----------------------------+
Private Sub Look_UTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 9
    RV_Torcd = iB1_Torcd
    RV_Kojyo = ""

    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iB1_Torcd = RV_TorcdK
        iB1_Uke = RV_Ukeir
        B1lb_Tornm = RV_TorRName
    End If
End Sub
'+----------------------------+
'+      品番検索              +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RVI_Torcd = RTrim(iB1_Torcd)
    RVI_Ukeir = RTrim(iB1_Uke)
    RVI_Hinmoku = 1

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       iB1_Hinbn = RV_Hinbn
       iB1_Torcd = RV_TorcdK
       iB1_Uke = RV_Ukeir
       B1lb_Tornm = ""
       If TorNmGet(iB1_Torcd, 9) Then
           B1lb_Tornm = TRM_RName
       End If
    End If
End Sub
'+-----------------------------+
'+  iB1_Torcd
'+-----------------------------+
Private Sub iB1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB1_Torcd
'    Back_JTorcd.Tag = "S"
End Sub
Private Sub iB1_Torcd_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 1) = iB1_Torcd
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 2) = ""
    B1lb_Tornm = ""
    If RTrim(iB1_Torcd) <> "" Then
        If TorNmGet(iB1_Torcd, 9) Then
            B1lb_Tornm = RTrim(TRM.略称)
            vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 2) = RTrim(TRM.略称)
        End If
    End If
End Sub
'+-----------------------------+
'+  iB1_Uke
'+-----------------------------+
Private Sub iB1_Uke_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB1_Uke
End Sub
Private Sub iB1_Uke_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Uke_LostFocus()
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
    Set ActObj = iB1_Hinbn
End Sub
Private Sub iB1_Hinbn_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_StartDay
'+-----------------------------+
Private Sub iH1_STartDay_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iH1_StartDay
End Sub
Private Sub iH1_StartDay_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If Back_JTorcd.Tag = "S" Then
        Call TextToVsGrid(Val(Input_Row.Tag))
        Call Grid_RowChange(vbKeyUp)
    End If
    If Back_JTorcd.Tag = "E" Then
        Call TextToVsGrid(Val(Input_Row.Tag))
        Call Grid_RowChange(vbKeyDown)
    End If
    SendKeys "{Tab}"
End Sub

'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim w_Color     As Long
   w_Color = &H80000008 '初期色ｾｯﾄ(黒)
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
     Case Else
          SyoriNM = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
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
    PB_Look.Visible = H1Mode
    PB_DEL.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    
    Call GridTextEnabled(B1Mode)
    
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
    If HeadBody = "B" Then
        CisVsGrid.Clear
        CisVsGrid.NumberSet = All
        CisVsGrid.ReNumber
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            SyoriKB = ""
            Call SyoriName(SyoriKB)
            GoTo ReturnPress_Ed
        End If
'
        ProcHB = "B1"
        Call DispChange(ProcHB)
        vsGrid.Row = 1
        Call VsGridToText(vsGrid.Row)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        SyoriKB = ""
        Call SyoriName(SyoriKB)
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
'開始日
    If Trim(iH1_StartDay) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    開始日を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_StartDay.SetFocus
       Exit Function
    End If
    
    gInt = 0
    For gCnt1 = 0 To 9
        If CisFun.Val2(vsGrid1.TextMatrix(0, gCnt1)) = CisFun.Val2(iH1_StartDay) Then
           gInt = 1
           Exit For
        End If
    Next gCnt1
    For gCnt1 = 0 To 9
        If CisFun.Val2(vsGrid1.TextMatrix(1, gCnt1)) = CisFun.Val2(iH1_StartDay) Then
           gInt = 1
           Exit For
        End If
    Next gCnt1
    If gInt = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    開始日が範囲外です       "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_StartDay.SetFocus
       Exit Function
    End If
' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = " SELECT KS.*,TR.略称 取引先名,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(HN.表示品番,'') <> '' THEN HN.表示品番 ELSE KS.品番 END 表示品番"
    gSL_Select = gSL_Select & " FROM かんばん増減開始日マスタ KS"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR.取引先CD = KS.取引先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "   ON  HN.品番 = KS.品番 "
    gSL_Select = gSL_Select & " WHERE KS.開始日 = '" & Format(CisFun.Val2(iH1_StartDay), "00") & "'"
    gSL_Select = gSL_Select & " ORDER BY KS.取引先,KS.受入,KS.品番"
    If DBInput("INV") Then
       SyoriKB = "MNT"
       Call SyoriName(SyoriKB)
    Else
        If SyoriKB = "DEL" Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    該当データが存在しません！        "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
            iH1_StartDay.SetFocus
            Exit Function
        End If
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
    End If
       
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False

    Call TextToVsGrid(Val(Input_Row.Tag))
    Call GridTextEnabled(False)

'ﾒｯｾｰｼﾞﾎﾞｯｸｽ設定
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'<< 行を詰める >>
    With vsGrid
        For gCnt1 = 1 To .Rows - 1
            If Not BodyInput(gCnt1) Then
                For gCnt2 = gCnt1 + 1 To .Rows - 1
                    If BodyInput(gCnt2) Then
                        For gCnt3 = 1 To 6
                            .TextMatrix(gCnt1, gCnt3) = .TextMatrix(gCnt2, gCnt3)
                            .TextMatrix(gCnt2, gCnt3) = ""
                        Next gCnt3
                        Exit For
                    End If
                Next gCnt2
            End If
        Next gCnt1
'
        For gCnt1 = 1 To .Rows - 1
            If Not BodyInput(gCnt1) Then Exit For
        '<< 取引先入力必須 >>
            If RTrim(.TextMatrix(gCnt1, 1)) = "" And _
               (RTrim(.TextMatrix(gCnt1, 3)) <> "" Or RTrim(.TextMatrix(gCnt1, 4)) <> "") Then
                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　取引先　未入力！】            "
                CisFun.MBOX
                Set ErrObj = iB1_Torcd
                GoTo Body1Chk_Err
            End If
            If RTrim(.TextMatrix(gCnt1, 1)) <> "" Then
            '<< 取引先マスタ存在チェック >>
                If RTrim(.TextMatrix(gCnt1, 1)) <> "" Then
                   If Not TorNmGet(.TextMatrix(gCnt1, 1), 9) Then
                       CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　取引先マスタ　未登録！】            "
                       CisFun.MBOX
                       Set ErrObj = iB1_Torcd
                       GoTo Body1Chk_Err
                   End If
                End If
            End If
        '<< 取引先受入入力 >>
            If RTrim(.TextMatrix(gCnt1, 3)) = "" And _
               RTrim(.TextMatrix(gCnt1, 4)) <> "" Then
                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　受入　未入力！】            "
                CisFun.MBOX
                Set ErrObj = iB1_Uke
                GoTo Body1Chk_Err
            End If
            If RTrim(.TextMatrix(gCnt1, 3)) <> "" And _
               RTrim(.TextMatrix(gCnt1, 1)) = "" Then
                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　取引先　未入力！】            "
                CisFun.MBOX
                Set ErrObj = iB1_Torcd
                GoTo Body1Chk_Err
            End If
            If RTrim(.TextMatrix(gCnt1, 3)) <> "" Then
        '<< 取引先受入マスタ存在チェック >>
                TUM.取引先CD = RTrim$(.TextMatrix(gCnt1, 1))
                TUM.取引先区分 = 1
                TUM.受入 = RTrim$(.TextMatrix(gCnt1, 3))
                If Not TUMRead("", 1) Then
                    CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　取引先受入マスタ　未登録！】            "
                    CisFun.MBOX
                    Set ErrObj = iB1_Uke
                    GoTo Body1Chk_Err
                End If
            End If
        '<< 品番入力 >>
            If RTrim(.TextMatrix(gCnt1, 4)) <> "" And _
               RTrim(.TextMatrix(gCnt1, 1)) = "" Then
                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　取引先　未入力！】            "
                CisFun.MBOX
                Set ErrObj = iB1_Torcd
                GoTo Body1Chk_Err
            End If
            If RTrim(.TextMatrix(gCnt1, 4)) <> "" And _
               RTrim(.TextMatrix(gCnt1, 3)) = "" Then
                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　受入　未入力！】            "
                CisFun.MBOX
                Set ErrObj = iB1_Uke
                GoTo Body1Chk_Err
            End If
            If RTrim(.TextMatrix(gCnt1, 4)) <> "" Then
        '<< 品番取引先マスタ存在チェック >>
                HTM.取引先 = RTrim$(.TextMatrix(gCnt1, 1))
                HTM.受入 = RTrim$(.TextMatrix(gCnt1, 3))
                HTM.品番 = RTrim$(.TextMatrix(gCnt1, 5))
                HTM.品目 = 1
                If Not HTMRead("", 1) Then
                    CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　品番取引先マスタ　未登録！】            "
                    CisFun.MBOX
                    Set ErrObj = iB1_Hinbn
                    GoTo Body1Chk_Err
                End If
            End If
        '<< 重複チェック >>
            For gCnt2 = gCnt1 + 1 To .Rows - 1
                If RTrim(.TextMatrix(gCnt1, 1)) = "" Then Exit For
                If RTrim(.TextMatrix(gCnt1, 1)) = RTrim(.TextMatrix(gCnt2, 1)) Then
                    CisFun.MB_MSG(1) = "    　取引先が重複しています。            "
                    CisFun.MB_MSG(3) = "    " & gCnt1 & "行目と " & gCnt2 & "行目            "
                    CisFun.MBOX
                    Set ErrObj = iB1_Torcd
                    gCnt1 = gCnt2
                    GoTo Body1Chk_Err
                End If
                If RTrim(.TextMatrix(gCnt1, 1)) = RTrim(.TextMatrix(gCnt2, 1)) And _
                   RTrim(.TextMatrix(gCnt1, 3)) = RTrim(.TextMatrix(gCnt2, 3)) Then
                    CisFun.MB_MSG(1) = "    　取引先・受入が重複しています。            "
                    CisFun.MB_MSG(3) = "    " & gCnt1 & "行目と " & gCnt2 & "行目            "
                    CisFun.MBOX
                    Set ErrObj = iB1_Torcd
                    gCnt1 = gCnt2
                    GoTo Body1Chk_Err
                End If
                If RTrim(.TextMatrix(gCnt1, 1)) = RTrim(.TextMatrix(gCnt2, 1)) And _
                   RTrim(.TextMatrix(gCnt1, 3)) = RTrim(.TextMatrix(gCnt2, 3)) And _
                   RTrim(.TextMatrix(gCnt1, 5)) = RTrim(.TextMatrix(gCnt2, 5)) Then
                    CisFun.MB_MSG(1) = "    　取引先・受入・品番が重複しています。            "
                    CisFun.MB_MSG(3) = "    " & gCnt1 & "行目と " & gCnt2 & "行目            "
                    CisFun.MBOX
                    Set ErrObj = iB1_Torcd
                    gCnt1 = gCnt2
                    GoTo Body1Chk_Err
                End If
            Next gCnt2
            
        Next gCnt1
    End With
    
    
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    vsGrid.Row = gCnt1
    Call VsGridToText(vsGrid.Row)
    Call GridTextEnabled(True)
    ErrObj.SetFocus
End Function
'+----------------------+
'+      入 力 判 定
'+----------------------+
Private Function BodyInput(RowNo As Long)
    BodyInput = False

    With vsGrid
        If RTrim(.TextMatrix(RowNo, 1)) = "" And _
           RTrim(.TextMatrix(RowNo, 3)) = "" And _
           RTrim(.TextMatrix(RowNo, 4)) = "" Then Exit Function
    End With

    BodyInput = True
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
    
    RD_Count = 0
    
    If Not KZSRead(gSL_Select, 0, 1) Then
        Call KZSClose(1)
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
        RD_Count = CisDB.RecordCount
        Call BodySet
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
Private Sub BodySet()

'Grid表示
    Call GridSet
    Call KZSClose(1)
        
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
        ' 取引先変換マスタ
        .SQL = "DELETE FROM かんばん増減開始日マスタ "
        .SQL = .SQL & " WHERE 開始日 = '" & Format(CisFun.Val2(iH1_StartDay), "00") & "'"
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

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
    Dim wYDJun As Byte

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     更新処理を行います      "
        .MB_Title = "更新処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call VsGridToText(Val(Input_Row.Tag))
            Call GridTextEnabled(True)
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
    If SyoriKB <> "ADD" Then
        With CisDB
            On Error GoTo DBPut_Err
    ' ( 削除処理 )
            ' 発注サイクルマスタ
            .SQL = "DELETE FROM かんばん増減開始日マスタ "
            .SQL = .SQL & " WHERE 開始日 = '" & Format(CisFun.Val2(iH1_StartDay), "00") & "'"
            .DBExec
    
            On Error GoTo 0
        End With
    End If
'   マスタ項目移送
    For gInt = 1 To vsGrid.Rows - 1
        If RTrim(vsGrid.TextMatrix(gInt, 1)) <> "" Then
            With KZS
                .取引先 = RTrim(vsGrid.TextMatrix(gInt, 1))
                .受入 = RTrim(vsGrid.TextMatrix(gInt, 3))
                .品番 = RTrim(vsGrid.TextMatrix(gInt, 5))
                .開始日 = Format(CisFun.Val2(iH1_StartDay), "00")
                .作成者 = gTanto
            End With
        ' ( 追加処理 )
            Call KZSInsert
        End If
    Next gInt
    
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
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 中中, あり, 5, 0, "№")
        Call .FixedSet(中中, 中中, あり, 8, 1, "取引先")
        Call .FixedSet(中中, 左中, あり, 21, 2, "取引先名")
        Call .FixedSet(中中, 中中, あり, 4, 3, "受入")
        Call .FixedSet(中中, 中中, あり, 30, 4, "表示品番")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, あり, 0, 5, "X品番")
        Call .FixedSet(中中, 左中, なし, 0, 6, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4")
        .InitDisp
    End With
'*--------------------------------------------------<< 2004/09/16 >> Update Start
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu = RD_Count

' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(CisFun.Val2(ID))
        Next ID
        If Ttl_Kensu + 1 > .Rows Then
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not KZS_RDSTS
        'Grid Row,Col
        ID = ID + 1
        With vsGrid
            .TextMatrix(ID, GridCol(0)) = ID
            .TextMatrix(ID, GridCol(1)) = KZS.取引先 + Space(7 - CisFun.Len2(KZS.取引先))
            .TextMatrix(ID, GridCol(2)) = RTrim$(KZS.取引先名)
            .TextMatrix(ID, GridCol(3)) = KZS.受入 + Space(2 - CisFun.Len2(KZS.受入))
            .TextMatrix(ID, GridCol(4)) = KZS.表示品番 + Space(30 - CisFun.Len2(KZS.表示品番))
            .TextMatrix(ID, GridCol(5)) = KZS.品番 + Space(25 - CisFun.Len2(KZS.品番))
        End With
        Call KZSReadNext(1)
    Loop

    With vsGrid
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    Input_Row.Visible = EnabledFlg
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'テキストからグリッドに退避
    Call TextToVsGrid(Val(Input_Row.Tag))
    gInt = vsGrid.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > vsGrid.Rows - 1 Then
            If RTrim$(vsGrid.TextMatrix(gInt, 1)) <> "" Then
                vsGrid.Rows = vsGrid.Rows + 1
                vsGrid.TextMatrix(gInt + 1, 0) = gInt + 1
            Else
                GoTo Grid_RowChange_End
            End If
        End If
        gInt = gInt + 1
    End If
    'PageUp
    If KeyCode = vbKeyPageUp Then
        If gInt - CisVsGrid.DispMax < 1 Then
            gInt = 1
        Else
            gInt = gInt - CisVsGrid.DispMax
        End If
    End If
    'PageDown
    If KeyCode = vbKeyPageDown Then
        If gInt + CisVsGrid.DispMax > vsGrid.Rows - 1 Then
            gInt = vsGrid.Rows - 1
        Else
            gInt = gInt + CisVsGrid.DispMax
        End If
    End If
    vsGrid.Row = gInt
'グリッドからテキストに表示
    Call VsGridToText(vsGrid.Row)
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With vsGrid
'       >> 自社取引先
        .TextMatrix(tRow, 1) = iB1_Torcd + Space(7 - CisFun.Len2(iB1_Torcd))
        .TextMatrix(tRow, 2) = ""
        If RTrim(iB1_Torcd) <> "" Then
            If TorNmGet(iB1_Torcd, 9) Then
                .TextMatrix(tRow, 2) = RTrim(TRM.略称)
            End If
        End If
        .TextMatrix(tRow, 3) = iB1_Uke + Space(2 - CisFun.Len2(iB1_Uke))
        HNM.品番 = RTrim(iB1_Hinbn)
        If HNMRead("", 1) Then
           .TextMatrix(tRow, 4) = HNM.表示品番 + Space(30 - CisFun.Len2(HNM.表示品番))
        Else
           .TextMatrix(tRow, 4) = RTrim(iB1_Hinbn)
        End If
        
'       >>
        .TextMatrix(tRow, 5) = RTrim(iB1_Hinbn)
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With vsGrid
        Input_Row.Tag = tRow
        Input_Row.Move Input_Row.Left, vsGrid.CellTop, Input_Row.Width, vsGrid.CellHeight
        '>> 取引先
        iB1_Torcd = RTrim(.TextMatrix(tRow, 1))
        '>> 自社取引先名
        B1lb_Tornm = RTrim(.TextMatrix(tRow, 2))
        '>> 受入
        iB1_Uke = RTrim$(.TextMatrix(tRow, 3))
        '>> 品番
        iB1_Hinbn = RTrim$(.TextMatrix(tRow, 5))
    End With
End Sub
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    Call TextToVsGrid(Val(Input_Row.Tag))
    Call VsGridToText(vsGrid.Row)
End Sub
Private Sub VSGrid_Click()
    Call TextToVsGrid(Val(Input_Row.Tag))
    If RTrim$(vsGrid.TextMatrix(vsGrid.Row, 0)) = "" Then Exit Sub
    Call GridTextEnabled(True)
    Call VsGridToText(vsGrid.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
Private Sub vsGrid_Scroll()
    'スクロール時はテキストボックスを非表示にします。
'    Call GridTextEnabled(False)
End Sub
Private Sub vsGrid1_Click()
    
    If RTrim(vsGrid1.TextMatrix(vsGrid1.Row, vsGrid1.Col)) <> "" Then
       iH1_StartDay = RTrim(vsGrid1.TextMatrix(vsGrid1.Row, vsGrid1.Col))
    End If

End Sub
