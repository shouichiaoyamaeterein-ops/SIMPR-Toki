VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "Cis3D_6.0.ocx"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CisBtn_6.0.ocx"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CisText_6.0.ocx"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Begin VB.Form CXM2030 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "品番マスタメンテ[一括]"
   ClientHeight    =   10980
   ClientLeft      =   5805
   ClientTop       =   3900
   ClientWidth     =   15330
   ControlBox      =   0   'False
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
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10980
   ScaleWidth      =   15330
   WindowState     =   2  '最大化
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   405
      Index           =   1
      Left            =   10995
      ScaleHeight     =   405
      ScaleWidth      =   3525
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1695
      Width           =   3525
      Begin Cis3D_v60.CIS3D UC_3D6 
         Height          =   345
         Left            =   30
         Top             =   30
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   609
         BackColor       =   14737632
         ForeColor       =   0
         Caption         =   "更新件数"
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
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_Select 
            Height          =   285
            Left            =   975
            Top             =   30
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
            BackColor       =   16777152
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
         Left            =   1755
         Top             =   30
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   609
         BackColor       =   14737632
         ForeColor       =   0
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   20
         Begin Cis3D_v60.CIS3D B1lb_Kensu 
            Height          =   285
            Left            =   960
            Top             =   30
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   503
            BackColor       =   12648384
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
      Height          =   1290
      Left            =   825
      ScaleHeight     =   1230
      ScaleWidth      =   5520
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   780
      Width           =   5580
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   0
         Left            =   3390
         Top             =   60
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "在庫発注(設定値)"
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
         cPositionX      =   30
         cPositionY      =   60
         Begin VB.ComboBox H1Comb_Zaiko 
            BackColor       =   &H00C0E0FF&
            Height          =   345
            Left            =   420
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   3
            Top             =   510
            Width           =   1290
         End
         Begin CisText_V60.CisText iH1_Zaiko 
            Height          =   360
            Left            =   210
            TabIndex        =   2
            Top             =   510
            Visible         =   0   'False
            Width           =   210
            _ExtentX        =   370
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
            cDataReplace    =   1
            cFbComma        =   0
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   4
         Left            =   60
         Top             =   60
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "品　　番"
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
         cPositionX      =   30
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   360
            Left            =   90
            TabIndex        =   1
            Top             =   510
            Width           =   3195
            _ExtentX        =   5636
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
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   2
         Left            =   60
         Top             =   60
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1984
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   90
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   510
            Visible         =   0   'False
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   635
            Enabled         =   0   'False
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
            Height          =   360
            Left            =   1080
            Top             =   510
            Width           =   2460
            _ExtentX        =   4339
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
      Height          =   735
      Left            =   11025
      ScaleHeight     =   675
      ScaleWidth      =   3420
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   570
      Width           =   3480
      Begin CisBtn_60.CisBtn PB_CAN 
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
   End
   Begin VB.Timer Timer2 
      Left            =   0
      Top             =   495
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   8070
      Left            =   795
      ScaleHeight     =   8010
      ScaleWidth      =   13665
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2175
      Width           =   13725
      Begin Cis3D_v60.CIS3D BLB_Back 
         Height          =   7860
         Left            =   75
         Top             =   90
         Width           =   13530
         _ExtentX        =   23865
         _ExtentY        =   13864
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
         cAlingnment     =   6
         Begin vsFlexLib.vsFlexArray vsGrid 
            Height          =   7665
            Left            =   90
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   105
            Width           =   13335
            _Version        =   196608
            _ExtentX        =   23521
            _ExtentY        =   13520
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
            FormatString    =   $"CXM2030.frx":0000
            Rows            =   21
            Cols            =   20
            BackColor       =   12648384
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   360
            SelectionMode   =   1
            Begin Cis3D_v60.CIS3D Input_Area 
               Height          =   375
               Left            =   10635
               Top             =   360
               Width           =   2610
               _ExtentX        =   4604
               _ExtentY        =   661
               BackColor       =   0
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
               Begin CisText_V60.CisText iB1_Keisu 
                  Height          =   360
                  Left            =   1320
                  TabIndex        =   6
                  Top             =   0
                  Width           =   735
                  _ExtentX        =   1296
                  _ExtentY        =   635
                  cFormat         =   "#0.0"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#0.0"
                  MaxLength       =   4
                  cDataType       =   2
                  cFcDecimal      =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#0.0"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iB1_ZaikoDay 
                  Height          =   360
                  Left            =   1305
                  TabIndex        =   13
                  Top             =   15
                  Visible         =   0   'False
                  Width           =   600
                  _ExtentX        =   1058
                  _ExtentY        =   635
                  cFormat         =   "#0.0"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#0.0"
                  MaxLength       =   4
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#0.0"
                  cILength        =   2
               End
               Begin Cis3D_v60.CIS3D B1lb_KMsg 
                  Height          =   330
                  Left            =   2055
                  Top             =   15
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   582
                  Caption         =   "%"
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
                  cAlingnment     =   7
               End
               Begin VB.ComboBox B1Comb_Zaiko 
                  BackColor       =   &H00C0E0FF&
                  Height          =   345
                  Left            =   210
                  Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
                  TabIndex        =   5
                  Top             =   0
                  Width           =   1095
               End
               Begin CisText_V60.CisText iB1_ZaikoL 
                  Height          =   375
                  Left            =   1275
                  TabIndex        =   14
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   661
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
               Begin CisText_V60.CisText iB1_Zaiko 
                  Height          =   360
                  Left            =   0
                  TabIndex        =   4
                  Top             =   0
                  Width           =   210
                  _ExtentX        =   370
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
                  cDataReplace    =   1
                  cFbComma        =   0
               End
            End
         End
         Begin VB.TextBox B_TabDummy 
            Height          =   330
            Left            =   150
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   180
            Width           =   240
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
      Left            =   480
      TabIndex        =   7
      Text            =   "X"
      Top             =   450
      Width           =   345
   End
   Begin VB.Timer Timer1 
      Left            =   900
      Top             =   450
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
      Caption         =   "【 品番マスタメンテ[一括]  】"
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
         Left            =   11685
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
   End
End
Attribute VB_Name = "CXM2030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   品番マスタメンテ[一括]
'**       フォームID    :   CXJ2030
'**       処理概要      :
'**
'**       作  成  日    :   2005/06/20 By CIS SIMPR-A
'**       変  更  日    :   2019/08/08 By CIS 仕入先入力削除
'**                                           コンボボックスから係数､日数を削除
'**       変  更  日    :   2019/10/24 By CIS コンボボックスから削除した日数を復活
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    Dim CisComboZaikoH  As CisCombo
    Dim CisComboZaikoB  As CisCombo
    Dim SL_Where        As String
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKb         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim ActObj          As Object
    Dim ErrObj          As Object
    Dim ActRow          As Integer
    Dim BinMax          As Byte         '設定最大便（８便）
    Dim mCycle          As String       'サイクル退避

    Dim SaveRow         As Integer
    
    Dim CmbIdx()        As String       ' 在庫区分（値）        ' 2019/08/09 ADD
    Dim CmbStr()        As String       ' 在庫区分（名称）      ' 2019/08/09 ADD
    Dim CmbCnt          As Integer      ' 在庫区分数            ' 2019/08/09 ADD
    Dim CmbSel()        As Boolean      ' 有効値                ' 2019/10/24 ADD
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
        .NumberSet = All
        .FixedHeight = 360
        .PatanMax = 1
        .SelectCol = 11
        .InitGet = False
        .Init
    End With
    Call GridInit
    '+---------------------+
    '+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
    '+---------------------+
    '===============================================================================' 2019/08/09 ADD START
    Dim iCnt As Integer: iCnt = 0
    gSL_Select = "SELECT 値, 値名称, COUNT(*) OVER() 件数 FROM 名称マスタ" + vbCrLf
    gSL_Select = gSL_Select + " WHERE 区分名称   = '在庫区分'" + vbCrLf
    gSL_Select = gSL_Select + "   AND 区分タイプ = ''        " + vbCrLf
    Call Str_Monitor(gSL_Select)
    If SYMRead(gSL_Select) Then
        With SYM
            CmbCnt = .件数
            ReDim CmbIdx(CmbCnt - 1)
            ReDim CmbStr(CmbCnt - 1)
            ReDim CmbSel(CmbCnt - 1)        ' 2019/10/24 ADD
            For iCnt = 0 To CmbCnt - 1
                CmbIdx(iCnt) = .値
                CmbStr(iCnt) = .値名称
                CmbSel(iCnt) = False        ' 2019/10/24 ADD
                Call SYMReadNext
            Next
        End With
    End If
    ' 在庫区分
    Set CisComboZaikoH = New CisCombo
    With CisComboZaikoH
       'Set .Connect = CisDB
        Set .CombObj = H1Comb_Zaiko
        .Control = ""
        .ComboName = ""
        .Ess = False
       '.IniPath = gIniExe
        .SetSpace = 50
        
        .CombowkSet (4)                                                 ' 2019/10/24 UPD：(3)→(4)
        For iCnt = 0 To CmbCnt - 1
           'If CmbStr(iCnt) <> "係数" And CmbStr(iCnt) <> "日数" Then   ' 2019/10/24 DEL
            If CmbStr(iCnt) <> "係数" Then                              ' 2019/10/24 ADD
                .Code = CmbIdx(iCnt): .CombText = CmbStr(iCnt)
                .CombowkSet
            End If
        Next

       '.ComboInit
       .ComboSet
    End With
    '===============================================================================' 2019/08/09 ADD END

    '===============================================================================' 2019/08/09 DEL START
'    ' 在庫区分
'    Set CisComboZaikoH = New CisCombo
'    With CisComboZaikoH
'        Set .Connect = CisDB
'        Set .CombObj = H1Comb_Zaiko
'        .Control = "名称マスタ"
'        .ComboName = "在庫区分"
'        .Ess = False
'        .IniPath = gIniExe
'        .SetSpace = 50
'        .ComboInit
'        .ComboSet
'    End With
    '===============================================================================' 2019/08/09 DEL END
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKb = ""
    Call SyoriName(SyoriKb)
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
'初期表示
    
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
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
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
   Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   Call CisFrm.MousePT(11)
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKb)
     'iH1_Torcd.SetFocus            ' 2019/08/09 DEL
      iH1_Hinbn.SetFocus            ' 2019/08/09 ADD
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
End Sub
'+----------------------------+
'+       取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()

    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 9
    RV_Call = "C"

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
       iH1_Torcd = RTrim(RV_TorcdK)
       Call iH1_Torcd_LostFocus
    End If

End Sub
'+----------------------------+
'+       品番検索            +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Left = 0: RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 1
    RVI_SK = 9
    RVI_Torcd = RTrim(iH1_Torcd)
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       iH1_Hinbn = RTrim(RV_Hinbn)
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
    If RTrim(iH1_Torcd) <> "" Then
        If TorNmGet(iH1_Torcd, 9) Then
            H1lb_Tornm = RTrim(TRM.略称)
        End If
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
'+  iB1_Zaiko
'+-----------------------------+
Private Sub iB1_Zaiko_GotFocus()
    Set ActObj = iB1_Zaiko
    B_TabDummy.Tag = "S"
End Sub
Private Sub iB1_Zaiko_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Zaiko_LostFocus()
    '===========================================================' 2019/10/24 ADD START
    Dim CmbNum As Integer                   ' 在庫発注の入力値
    CmbNum = CisFun.Val2(iB1_Zaiko.text)
    If CmbNum = 0 Then                      ' 数値0なら文字"0"にする（数値以外も0になるため）
        iB1_Zaiko.text = "0"
    ElseIf CmbNum >= CmbCnt Then            ' 登録数以上は無効
        iB1_Zaiko.text = "0"
    ElseIf CmbSel(CmbNum) = False Then      ' 許可していない値は無効
        iB1_Zaiko.text = "0"
    End If
    '===========================================================' 2019/10/24 ADD END

    vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 8) = iB1_Zaiko
    CisComboZaikoB.Code = iB1_Zaiko
    CisComboZaikoB.TextGet
End Sub
'+-----------------------------+
'+  B1Comb_Zaiko
'+-----------------------------+
Private Sub B1Comb_Zaiko_Click()
    CisComboZaikoB.CodeGet
    iB1_Zaiko = CisComboZaikoB.Code
    Select Case iB1_Zaiko
        '係数
        Case "2": iB1_Keisu.cDataType = ｢9｣整数
                    iB1_Keisu.cFcDecimal = 0
                    iB1_Keisu.cFormat = "###"
                    iB1_Keisu.cFbComma = 0
                    iB1_Keisu.MaxLength = 3
                    iB1_Keisu.Width = 735
                    iB1_Keisu = ""
                    iB1_Keisu = Format(CisFun.Val2(vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6)), "###")
                    B1lb_KMsg = "%"
        Case "3": iB1_Keisu.cDataType = ｢9｣整数
                    iB1_Keisu.cFcDecimal = 0
                    iB1_Keisu.cFormat = "#,###"
                    iB1_Keisu.cFbComma = 1
                    iB1_Keisu.MaxLength = 7
                    iB1_Keisu.Width = 1005
                    iB1_Keisu = ""
                    iB1_Keisu = Format(CisFun.Val2(vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6)), "#,###")
        Case "4":
                    iB1_Keisu.MaxLength = 5
                    iB1_Keisu.cDataType = ｢V9｣小数
                    iB1_Keisu.cFcDecimal = 1
                    iB1_Keisu.cFormat = "#0.0"
                    iB1_Keisu.cFbComma = 0
                    iB1_Keisu.Width = 735
                    iB1_Keisu = ""
                    iB1_Keisu = Format(CisFun.Val2(vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6)), "#0.0")
                    B1lb_KMsg = "日"
        Case Else: iB1_Keisu.cDataType = ｢9｣整数
                    iB1_Keisu.cFcDecimal = 0
                    iB1_Keisu.cFormat = "###"
                    iB1_Keisu.MaxLength = 1
                    iB1_Keisu.Width = 735
                    iB1_Keisu = ""
                    B1lb_KMsg = ""
    End Select
End Sub
'+-----------------------------+
'+  iB1_Keisu
'+-----------------------------+
Private Sub iB1_Keisu_GotFocus()
    Set ActObj = iB1_Keisu
    B_TabDummy.Tag = "E"
End Sub
Private Sub iB1_Keisu_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Keisu_LostFocus()
    If CisFun.Val2(iB1_Keisu) <> 0 Then
        vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = iB1_Keisu
        Select Case iB1_Zaiko
            '係数
            Case "2":  vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = Format(CisFun.Val2(iB1_Keisu), "###")
            Case "3":  vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = Format(CisFun.Val2(iB1_Keisu), "#,###")
            Case "4":  vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = Format(CisFun.Val2(iB1_Keisu), "#0.0")
            Case Else: vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = ""
        End Select
    Else
        vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = ""
    End If
End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If B_TabDummy.Tag = "S" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyUp)
    End If
    If B_TabDummy.Tag = "E" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
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
   Select Case SyoriKb
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
       If gObj.Name Like "*vsGrid*" Then gObj.ForeColor = w_Color
'      * PF-Key 表示設定 *
       If SyoriKb <> "LOC" Then
          If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
             gObj.Name = "PB_DEL" Then
             gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKb)
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
    PB_Look.Visible = H1Mode
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
                .text = ""
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
    End If
    If HeadBody = "B1" Then
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
        Call SyoriName(SyoriKb)
       'iH1_Torcd.SetFocus          ' 2019/10/24 DEL
        iH1_Hinbn.SetFocus          ' 2019/10/24 ADD
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
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'===============================================================================' 2019/10/24 DEL START
''仕入先
'    If Trim(iH1_Torcd) <> "" And Not TorNmGet(iH1_Torcd, 9) Then
'        CisFun.MB_MSG(2) = "    仕入先　未登録！（取引先マスタ）            "
'        CisFun.MBOX
'        iH1_Torcd.SetFocus
'        Exit Function
'    End If
'===============================================================================' 2019/10/24 DEL END

' マスタ入力 SELECT文 ｾｯﾄ
    SL_Where = ""
    
    If RTrim(iH1_Torcd) = "" Then
        gSL_Select = " SELECT 品番,表示品番,車型,社内背番,在庫管理区分,在庫係数,在庫日数,在庫発注点,品名,ISNULL(SY.値名称,'')在庫管理区分名 "
        gSL_Select = gSL_Select & "  FROM 品番マスタ HN"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY"
        gSL_Select = gSL_Select & "     ON SY.区分名称 = '在庫区分'"
        gSL_Select = gSL_Select & "     AND SY.区分タイプ = ''"
        gSL_Select = gSL_Select & "     AND SY.値 = HN.在庫管理区分"
        With CisFun
            .Where_CreateKB = NewDocument
            .Where_Create AlphaMe, "HN.品番", RTrim(iH1_Hinbn)
            CisComboZaikoH.CodeGet
            .Where_Create AlphaMe, "在庫管理区分", RTrim(CisComboZaikoH.Code)
            SL_Where = .Where_Phrase
        End With
        gSL_Select = gSL_Select & SL_Where
    Else
        gSL_Select = " SELECT HN.品番,表示品番,車型,社内背番,在庫管理区分,在庫係数,在庫日数,在庫発注点,品名,ISNULL(SY.値名称,'')在庫管理区分名"
        gSL_Select = gSL_Select & "  FROM 品番取引先マスタ HT"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
        gSL_Select = gSL_Select & "     ON HN.品番 = HT.品番"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY"
        gSL_Select = gSL_Select & "     ON SY.区分名称 = '在庫区分'"
        gSL_Select = gSL_Select & "     AND SY.区分タイプ = ''"
        gSL_Select = gSL_Select & "     AND SY.値 = HN.在庫管理区分"
        With CisFun
            .Where_CreateKB = NewDocument
            .Where_Create Numeric, "HT.品目", 1
            .Where_Create AlphaMe, "HT.品番", RTrim(iH1_Hinbn)
            .Where_Create AlphaMe, "HT.取引先", RTrim(iH1_Torcd)
            CisComboZaikoH.CodeGet
            .Where_Create AlphaMe, "在庫管理区分", RTrim(CisComboZaikoH.Code)
            SL_Where = .Where_Phrase
        End With
        gSL_Select = gSL_Select & SL_Where
        gSL_Select = gSL_Select & " GROUP BY HN.品番,表示品番,車型,社内背番,在庫管理区分,在庫係数,在庫日数,在庫発注点,品名,SY.値名称"
    End If
    gSL_Select = gSL_Select & " ORDER BY HN.品番"
    If Not DBInput("INV") Then
        CisFun.MB_MSG(2) = "    対象品番　未登録！           "
        CisFun.MBOX
       'iH1_Torcd.SetFocus      ' 2019/10/24 DEL
        iH1_Hinbn.SetFocus      ' 2019/10/24 ADD
        Exit Function
    End If
       
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call GridTextEnabled(False)
    
    With vsGrid
        For gCnt1 = 1 To .Rows - 1
            If RTrim(.TextMatrix(gCnt1, 1)) = "" Then
                Exit For
            End If
            ' 在庫管理(係数)
            If RTrim(.TextMatrix(gCnt1, 8)) = "2" Then
               If CisFun.Val2(.TextMatrix(gCnt1, 6)) = 0 Then
                   With CisFun
                       .MB_Lines = 4
                       .MB_MSG(2) = "    在庫係数を入力して下さい            "
                       .MB_Title = ""
                       .MB_Button = Error
                       .MBOX
                   End With
                    Set ErrObj = iB1_Keisu
                    GoTo Body1Chk_Err
                End If
            End If
            ' 在庫管理(発注点)
            If RTrim(.TextMatrix(gCnt1, 8)) = "3" Then
               If CisFun.Val2(.TextMatrix(gCnt1, 6)) = 0 Then
                   With CisFun
                       .MB_Lines = 4
                       .MB_MSG(2) = "    在庫発注点を入力して下さい            "
                       .MB_Title = ""
                       .MB_Button = Error
                       .MBOX
                   End With
                    Set ErrObj = iB1_Keisu
                    GoTo Body1Chk_Err
                End If
            End If
            ' 在庫管理(日数)
            If RTrim(.TextMatrix(gCnt1, 8)) = "4" Then
               If CisFun.Val2(.TextMatrix(gCnt1, 6)) = 0 Then
                   With CisFun
                       .MB_Lines = 4
                       .MB_MSG(2) = "    在庫日数を入力して下さい            "
                       .MB_Title = ""
                       .MB_Button = Error
                       .MBOX
                   End With
                    Set ErrObj = iB1_Keisu
                    GoTo Body1Chk_Err
                End If
            End If
            
            
        Next gCnt1
    End With
'
'
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    vsGrid.Row = gCnt1
    Call VsGridToText(vsGrid.Row)
    Call GridTextEnabled(True)
    ErrObj.SetFocus
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("B1")
    
    If Not HNMRead(gSL_Select, 0, 1) Then
        Call HNMClose
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
        Call GridSet
        DBInput = True
    End If
    
DBInput_Ed:
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
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Dim AddFlg  As Boolean

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     更新処理を行います      "
        .MB_Title = "更新処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call VsGridToText(Val(Input_Area.Tag))
            Call GridTextEnabled(True)
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
'   マスタ項目移送
    For gIndex = 1 To vsGrid.Rows - 1
        If RTrim(vsGrid.TextMatrix(gIndex, 1)) <> "" And RTrim(vsGrid.TextMatrix(gIndex, CisVsGrid.SelectCol)) <> "" Then
            Call ItemsClearHNM
            With HNM
                .品番 = RTrim(vsGrid.TextMatrix(gIndex, 7))
                .在庫管理区分 = RTrim(vsGrid.TextMatrix(gIndex, 8))
                Select Case vsGrid.TextMatrix(gIndex, 8)
                    Case "2": .在庫係数 = CisFun.Val2(vsGrid.TextMatrix(gIndex, 6))
                    Case "3": .在庫発注点 = CisFun.Val2(vsGrid.TextMatrix(gIndex, 6))
                    Case "4": .在庫日数 = CisFun.Val2(vsGrid.TextMatrix(gIndex, 6))
                End Select
                .更新者 = gTanto
            End With
        ' ( 更新処理 )
            gSL_Select = " UPDATE 品番マスタ SET"
            gSL_Select = gSL_Select & " 在庫管理区分 = " & HNM.在庫管理区分 & ","
            gSL_Select = gSL_Select & " 在庫係数 = " & HNM.在庫係数 & ","
            gSL_Select = gSL_Select & " 在庫発注点 = " & HNM.在庫発注点 & ","
            gSL_Select = gSL_Select & " 在庫日数 = " & HNM.在庫日数 & ","
            gSL_Select = gSL_Select & " 更新者 = '" & HNM.更新者 & "',"
            gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
            gSL_Select = gSL_Select & " 更新端末 = HOST_NAME()"
            gSL_Select = gSL_Select & " WHERE 品番 = '" & HNM.品番 & "'"
            CisDB.SQL = gSL_Select
            On Error GoTo DBPut_Err
            CisDB.DBExec
            
           ' iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
        End If
    Next gIndex
'
' ( 追加処理 )
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
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 左中, あり, 30, 1, "品　番")
        Call .FixedSet(中中, 中中, あり, 4, 2, "車型")
        Call .FixedSet(中中, 中中, あり, 8, 3, "社内背番")
        Call .FixedSet(中中, 左中, あり, 40, 4, "品　　名")
        Call .FixedSet(中中, 中中, あり, 10, 5, "在庫発注")
        Call .FixedSet(中中, 右中, あり, 8, 6, "係数")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 7, "X品番")
        Call .FixedSet(中中, 左中, なし, 0, 8, "X在庫発注")
        Call .FixedSet(中中, 左中, なし, 0, 9, "X初期在庫区分")
        Call .FixedSet(中中, 左中, なし, 0, 10, "X初期係数")
        Call .FixedSet(中中, 左中, なし, 0, 11, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6")
        .InitDisp
    End With
    
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
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
    Do Until Not HNM_RDSTS
        'Grid Row,Col
        ID = ID + 1
        With vsGrid
            .TextMatrix(ID, GridCol(0)) = ID
            .TextMatrix(ID, GridCol(1)) = RTrim(HNM.表示品番)
            .TextMatrix(ID, GridCol(2)) = RTrim(HNM.車型)
            .TextMatrix(ID, GridCol(3)) = RTrim(HNM.社内背番)
            .TextMatrix(ID, GridCol(4)) = RTrim(HNM.品名)
            .TextMatrix(ID, GridCol(5)) = RTrim(HNM.在庫管理区分名)
            Select Case HNM.在庫管理区分
                Case 2: .TextMatrix(ID, GridCol(6)) = Format(HNM.在庫係数, "###")
                Case 3: .TextMatrix(ID, GridCol(6)) = Format(HNM.在庫発注点, "#,###")
                Case 4: .TextMatrix(ID, GridCol(6)) = Format(HNM.在庫日数, "#0.0")
            End Select
            .TextMatrix(ID, GridCol(7)) = RTrim(HNM.品番)
            .TextMatrix(ID, GridCol(8)) = RTrim(HNM.在庫管理区分)
            .TextMatrix(ID, GridCol(9)) = RTrim(HNM.在庫管理区分)
            Select Case HNM.在庫管理区分
                Case 2: .TextMatrix(ID, GridCol(10)) = Format(HNM.在庫係数, "###")
                Case 3: .TextMatrix(ID, GridCol(10)) = Format(HNM.在庫発注点, "#,###")
                Case 4: .TextMatrix(ID, GridCol(10)) = Format(HNM.在庫日数, "#0.0")
            End Select
        End With
        Call HNMReadNext(1)
    Loop

    With vsGrid
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    Call HNMClose(1)
    
    '===============================================================================' 2019/08/09 ADD START
    ' 在庫区分
    Dim iCnt As Integer
    Set CisComboZaikoB = New CisCombo
    With CisComboZaikoB
        Set .CombObj = B1Comb_Zaiko
        .Control = ""
        .ComboName = ""
        .Ess = True
        .SetSpace = 50
        .CombowkSet (4)                                                 ' 2019/10/24 UPD：(3)→(4)
        For iCnt = 0 To CmbCnt - 1
           'If CmbStr(iCnt) <> "係数" And CmbStr(iCnt) <> "日数" Then   ' 2019/10/24 DEL
            If CmbStr(iCnt) <> "係数" Then                              ' 2019/10/24 ADD
                CmbSel(iCnt) = True                                     ' 2019/10/24 ADD
                .Code = CmbIdx(iCnt): .CombText = CmbStr(iCnt)
                .CombowkSet
            End If
        Next
       .ComboSet
    End With
    '===============================================================================' 2019/08/09 ADD END

    '===============================================================================' 2019/08/09 DEL START
'    ' 在庫区分
'    Set CisComboZaikoB = New CisCombo
'    With CisComboZaikoB
'        Set .Connect = CisDB
'        Set .CombObj = B1Comb_Zaiko
'        .Control = "名称マスタ"
'        .ComboName = "在庫区分"
'        .Ess = False
'        .IniPath = gIniExe
'        .SetSpace = 50
'        .ComboInit
'        .ComboSet
'    End With
    '===============================================================================' 2019/08/09 DEL END
End Sub
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    Input_Area.Visible = EnabledFlg
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'テキストからグリッドに退避
    Call TextToVsGrid(Val(Input_Area.Tag))
    gInt = vsGrid.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > vsGrid.Rows - 1 Then
                GoTo Grid_RowChange_End
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
    B1Comb_Zaiko.Refresh
    iB1_Zaiko.Refresh
    iB1_Keisu.Refresh
    B1lb_KMsg.Refresh
    Input_Area.Refresh
Grid_RowChange_End:
    KeyCode = 0
'    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    SaveRow = vsGrid.Row
    With vsGrid
    '在庫区分
        .TextMatrix(tRow, 8) = RTrim(iB1_Zaiko)
        CisComboZaikoB.NameGet
        .TextMatrix(tRow, 5) = CisComboZaikoB.CombText
    '係数
        If CisFun.Val2(iB1_Keisu) <> 0 Then
            Select Case iB1_Zaiko
                '係数
                Case "2":  .TextMatrix(tRow, 6) = Format(CisFun.Val2(iB1_Keisu), "###")
                Case "3":  .TextMatrix(tRow, 6) = Format(CisFun.Val2(iB1_Keisu), "#,###")
                Case "4":  .TextMatrix(tRow, 6) = Format(CisFun.Val2(iB1_Keisu), "#0.0")
                Case Else: .TextMatrix(tRow, 6) = ""
            End Select
        Else
            .TextMatrix(tRow, 6) = ""
        End If
    '値の変更
        If RTrim(.TextMatrix(tRow, 8)) <> RTrim(.TextMatrix(tRow, 9)) Or _
           CisFun.Val2(.TextMatrix(tRow, 6)) <> CisFun.Val2(.TextMatrix(tRow, 10)) Then
            If RTrim(.TextMatrix(tRow, CisVsGrid.SelectCol)) = "" Then
                vsGrid.Row = tRow
                Call CisVsGrid.vsColor(個別, &HC0E0FF)
                B1lb_Select = Format(CisFun.Val2(B1lb_Select) + 1, "#,###")
                vsGrid.Row = SaveRow
            End If
        Else
            If RTrim(.TextMatrix(tRow, CisVsGrid.SelectCol)) <> "" Then
                vsGrid.Row = tRow
                Call CisVsGrid.vsColor(消去個別)
                B1lb_Select = Format(CisFun.Val2(B1lb_Select) - 1, "#,###")
                vsGrid.Row = SaveRow
            End If
        End If
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With vsGrid
        Input_Area.Tag = tRow
        Input_Area.Move Input_Area.Left, vsGrid.CellTop, Input_Area.Width, vsGrid.CellHeight
        '在庫区分
        iB1_Zaiko = RTrim(.TextMatrix(tRow, 8))
        CisComboZaikoB.Code = RTrim(.TextMatrix(tRow, 8))
        CisComboZaikoB.TextGet
            
        '係数
        Select Case iB1_Zaiko
            Case "2": iB1_Keisu.cDataType = ｢9｣整数
                        iB1_Keisu.cFcDecimal = 0
                        iB1_Keisu.cFormat = "###"
                        iB1_Keisu.cFbComma = 0
                        iB1_Keisu.MaxLength = 3
                        iB1_Keisu.Width = 735
                        iB1_Keisu = ""
                        iB1_Keisu = Format(CisFun.Val2(vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6)), "###")
                        B1lb_KMsg = "%"
            Case "3": iB1_Keisu.cDataType = ｢9｣整数
                        iB1_Keisu.cFcDecimal = 0
                        iB1_Keisu.cFormat = "#,###"
                        iB1_Keisu.cFbComma = 1
                        iB1_Keisu.MaxLength = 7
                        iB1_Keisu.Width = 1005
                        iB1_Keisu = ""
                        iB1_Keisu = Format(CisFun.Val2(vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6)), "#,###")
            Case "4":
                        iB1_Keisu.MaxLength = 5
                        iB1_Keisu.cDataType = ｢V9｣小数
                        iB1_Keisu.cFcDecimal = 1
                        iB1_Keisu.cFormat = "#0.0"
                        iB1_Keisu.cFbComma = 0
                        iB1_Keisu.Width = 735
                        iB1_Keisu = ""
                        iB1_Keisu = Format(CisFun.Val2(vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6)), "#0.0")
                        B1lb_KMsg = "日"
            Case Else: iB1_Keisu.cDataType = ｢9｣整数
                        iB1_Keisu.cFcDecimal = 0
                        iB1_Keisu.cFormat = "###"
                        iB1_Keisu.MaxLength = 1
                        iB1_Keisu.Width = 735
                        iB1_Keisu = ""
                        B1lb_KMsg = ""
        End Select
        DoEvents
    End With
End Sub
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call VsGridToText(vsGrid.Row)
End Sub
Private Sub VSGrid_Click()
    Call TextToVsGrid(Val(Input_Area.Tag))
    If vsGrid.TextMatrix(vsGrid.Row, 0) = "" Then Exit Sub
    Call GridTextEnabled(True)
    Call VsGridToText(vsGrid.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
Private Sub VSGrid_DblClick()
    If RTrim(vsGrid.TextMatrix(vsGrid.Row, 1)) = "" Then
        Exit Sub
    End If
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     " & vsGrid.Row & " 行目をクリアしますか？      "
        .MB_Title = ""
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Sub
    End With
    iB1_Zaiko = "0": iB1_Keisu = ""
    CisComboZaikoB.Code = "0": CisComboZaikoB.TextGet
    
    Call TextToVsGrid(Val(Input_Area.Tag))

End Sub
Private Sub vsGrid_Scroll()
    'スクロール時はテキストボックスを非表示にします。
'    Call GridTextEnabled(False)
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++    '++ デバッグ用です +++
'+                                                 +
'+  文字列をイミディエイトウィンドウに表示する     +
'+                                                 +
'+++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub Str_Monitor(ByVal wStr As String)
    Debug.Print vbCrLf
    Debug.Print Format(Now, "yyyy/MM/dd HH:mm:ss") + _
                " .................................................." + _
                vbCrLf + wStr
End Sub
