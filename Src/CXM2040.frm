VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CXM2040 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "工順置場メンテ[一括]"
   ClientHeight    =   10980
   ClientLeft      =   105
   ClientTop       =   3795
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
      TabIndex        =   8
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
      ScaleWidth      =   7140
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   780
      Width           =   7200
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   2
         Left            =   60
         Top             =   60
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "取  引  先"
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
            Top             =   510
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   4
         Left            =   3720
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
      TabIndex        =   5
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
      TabIndex        =   3
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
            Left            =   75
            TabIndex        =   4
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
            FormatString    =   $"CXM2040.frx":0000
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
               Left            =   8550
               Top             =   360
               Width           =   1395
               _ExtentX        =   2461
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
               Begin CisText_V60.CisText iB1_Okiba1 
                  Height          =   375
                  Left            =   15
                  TabIndex        =   9
                  Top             =   0
                  Width           =   330
                  _ExtentX        =   582
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XX"
                  MaxLength       =   2
               End
               Begin CisText_V60.CisText iB1_Okiba2 
                  Height          =   375
                  Left            =   345
                  TabIndex        =   10
                  Top             =   0
                  Width           =   330
                  _ExtentX        =   582
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XX"
                  MaxLength       =   2
               End
               Begin CisText_V60.CisText iB1_Okiba3 
                  Height          =   375
                  Left            =   675
                  TabIndex        =   11
                  Top             =   0
                  Width           =   330
                  _ExtentX        =   582
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XX"
                  MaxLength       =   2
               End
               Begin CisText_V60.CisText iB1_Okiba4 
                  Height          =   375
                  Left            =   1005
                  TabIndex        =   12
                  Top             =   0
                  Width           =   330
                  _ExtentX        =   582
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
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
         Begin VB.TextBox B_TabDummy 
            Height          =   330
            Left            =   150
            TabIndex        =   6
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
      TabIndex        =   2
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
      Caption         =   "【 工順置場メンテメンテ[一括]  】"
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
Attribute VB_Name = "CXM2040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   工順置場メンテ[一括]
'**       フォームID    :   CXM2040
'**       処理概要      :
'**
'**       作  成  日    :   2010/01/20 By CIS SIMPR-A
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
'    Dim CisComboZaikoH  As CisCombo
'    Dim CisComboZaikoB  As CisCombo
    Dim SL_Where        As String
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKb         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim ActObj          As Object
    Dim ErrObj          As Object
    Dim ActRow          As Integer

    Dim SaveRow         As Integer
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
        .SelectCol = 13
        .InitGet = False
        .Init
    End With
    Call GridInit
'    '+---------------------+
'    '+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
'    '+---------------------+
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
      iH1_Torcd.SetFocus
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
    If PB_Look.Tag = "iB1_Okiba1" Then Call Look_Kubun(13): Exit Sub
    If PB_Look.Tag = "iB1_Okiba2" Then Call Look_Kubun(14): Exit Sub
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
'+----------------------------+
'+       名称マスタ検索        +
'+----------------------------+
Private Sub Look_Kubun(KB As Integer)

    RV_Left = 0
    RV_Top = 0

    Select Case KB
        Case 13: RVI_Kubun = "管理場所"
        Case 14: RVI_Kubun = "管理場所ユーザー"
    End Select
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        Select Case KB
            Case 13
                iB1_Okiba1 = RV_Code
                Call iB1_Okiba1_LostFocus
            Case 14
                iB1_Okiba2 = RV_Code
                Call iB1_Okiba2_LostFocus
        End Select
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
'+  iB1_Okiba1
'+-----------------------------+
Private Sub iB1_Okiba1_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB1_Okiba1
    B_TabDummy.Tag = "S"
End Sub
Private Sub iB1_Okiba1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Okiba1_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Okiba2
'+-----------------------------+
Private Sub iB1_Okiba2_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB1_Okiba2
    B_TabDummy.Tag = "S"
End Sub
Private Sub iB1_Okiba2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Okiba2_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Okiba3
'+-----------------------------+
Private Sub iB1_Okiba3_GotFocus()
    Set ActObj = iB1_Okiba3
    B_TabDummy.Tag = "S"
End Sub
Private Sub iB1_Okiba3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'+-----------------------------+
'+  iB1_Okiba4
'+-----------------------------+
Private Sub iB1_Okiba4_GotFocus()
    Set ActObj = iB1_Okiba4
    B_TabDummy.Tag = "E"
End Sub
Private Sub iB1_Okiba4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'Private Sub iB1_Okiba1_LostFocus()
'    vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 8) = iB1_Okiba1
'    CisComboZaikoB.Code = iB1_Zaiko
'    CisComboZaikoB.TextGet
'End Sub
''+-----------------------------+
''+  B1Comb_Zaiko
''+-----------------------------+
'Private Sub B1Comb_Zaiko_Click()
'    CisComboZaikoB.CodeGet
'    iB1_Zaiko = CisComboZaikoB.Code
'    Select Case iB1_Zaiko
'        '係数
'        Case "2": iB1_Keisu.cDataType = ｢9｣整数
'                    iB1_Keisu.cFcDecimal = 0
'                    iB1_Keisu.cFormat = "###"
'                    iB1_Keisu.cFbComma = 0
'                    iB1_Keisu.MaxLength = 3
'                    iB1_Keisu.Width = 735
'                    iB1_Keisu = ""
'                    iB1_Keisu = Format(CisFun.Val2(vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6)), "###")
'                    B1lb_KMsg = "%"
'        Case "3": iB1_Keisu.cDataType = ｢9｣整数
'                    iB1_Keisu.cFcDecimal = 0
'                    iB1_Keisu.cFormat = "#,###"
'                    iB1_Keisu.cFbComma = 1
'                    iB1_Keisu.MaxLength = 7
'                    iB1_Keisu.Width = 1005
'                    iB1_Keisu = ""
'                    iB1_Keisu = Format(CisFun.Val2(vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6)), "#,###")
'        Case "4":
'                    iB1_Keisu.MaxLength = 5
'                    iB1_Keisu.cDataType = ｢V9｣小数
'                    iB1_Keisu.cFcDecimal = 1
'                    iB1_Keisu.cFormat = "#0.0"
'                    iB1_Keisu.cFbComma = 0
'                    iB1_Keisu.Width = 735
'                    iB1_Keisu = ""
'                    iB1_Keisu = Format(CisFun.Val2(vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6)), "#0.0")
'                    B1lb_KMsg = "日"
'        Case Else: iB1_Keisu.cDataType = ｢9｣整数
'                    iB1_Keisu.cFcDecimal = 0
'                    iB1_Keisu.cFormat = "###"
'                    iB1_Keisu.MaxLength = 1
'                    iB1_Keisu.Width = 735
'                    iB1_Keisu = ""
'                    B1lb_KMsg = ""
'    End Select
'End Sub
''+-----------------------------+
''+  iB1_Keisu
''+-----------------------------+
'Private Sub iB1_Keisu_GotFocus()
'    Set ActObj = iB1_Keisu
'    B_TabDummy.Tag = "E"
'End Sub
'Private Sub iB1_Keisu_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call Grid_RowChange(KeyCode)
'End Sub
'Private Sub iB1_Keisu_LostFocus()
'    If CisFun.Val2(iB1_Keisu) <> 0 Then
'        vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = iB1_Keisu
'        Select Case iB1_Zaiko
'            '係数
'            Case "2":  vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = Format(CisFun.Val2(iB1_Keisu), "###")
'            Case "3":  vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = Format(CisFun.Val2(iB1_Keisu), "#,###")
'            Case "4":  vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = Format(CisFun.Val2(iB1_Keisu), "#0.0")
'            Case Else: vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = ""
'        End Select
'    Else
'        vsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = ""
'    End If
'End Sub
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
        iH1_Torcd.SetFocus
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
'仕入先
    If Trim(iH1_Torcd) <> "" And Not TorNmGet(iH1_Torcd, 9) Then
        CisFun.MB_MSG(2) = "    取引先　未登録！（取引先マスタ）            "
        CisFun.MBOX
        iH1_Torcd.SetFocus
        Exit Function
    End If

' マスタ入力 SELECT文 ｾｯﾄ
    SL_Where = ""
    
    gSL_Select = " SELECT HK.取引先,TR.略称 取引先名,HK.受入,HK.品番,HN.表示品番,HK.工順,HN.車型,HT.背番号,HK.置場,HN.品名"
    gSL_Select = gSL_Select & "  FROM 品番工順マスタ HK"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "     ON HN.品番 = HK.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "     ON  TR.取引先CD = HK.取引先"
    gSL_Select = gSL_Select & "     AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT"
    gSL_Select = gSL_Select & "     ON  HT.品番 = HK.品番"
    gSL_Select = gSL_Select & "     AND HT.取引先 = HK.取引先"
    gSL_Select = gSL_Select & "     AND HT.受入 = HK.受入"
    With CisFun
        .Where_CreateKB = NewDocument
        .Where_Create Numeric, "HT.品目", 1
        .Where_Create AlphaMe, "HT.品番", RTrim(iH1_Hinbn)
        .Where_Create AlphaMe, "HT.取引先", RTrim(iH1_Torcd)
'            CisComboZaikoH.CodeGet
'            .Where_Create AlphaMe, "在庫管理区分", RTrim(CisComboZaikoH.Code)
        SL_Where = .Where_Phrase
    End With
    gSL_Select = gSL_Select & SL_Where
'        gSL_Select = gSL_Select & " GROUP BY HN.品番,表示品番,車型,社内背番,在庫管理区分,在庫係数,在庫日数,在庫発注点,品名,SY.値名称"
    gSL_Select = gSL_Select & " ORDER BY HK.取引先,HK.受入,HK.品番,HK.工順"
    If Not DBInput("INV") Then
        CisFun.MB_MSG(2) = "    対象品番　未登録！           "
        CisFun.MBOX
        iH1_Torcd.SetFocus
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
    
'    With vsGrid
'        For gCnt1 = 1 To .Rows - 1
'            If RTrim(.TextMatrix(gCnt1, 1)) = "" Then
'                Exit For
'            End If
'            ' 在庫管理(係数)
'            If RTrim(.TextMatrix(gCnt1, 8)) = "2" Then
'               If CisFun.Val2(.TextMatrix(gCnt1, 6)) = 0 Then
'                   With CisFun
'                       .MB_Lines = 4
'                       .MB_MSG(2) = "    在庫係数を入力して下さい            "
'                       .MB_Title = ""
'                       .MB_Button = Error
'                       .MBOX
'                   End With
'                    Set ErrObj = iB1_Keisu
'                    GoTo Body1Chk_Err
'                End If
'            End If
'            ' 在庫管理(発注点)
'            If RTrim(.TextMatrix(gCnt1, 8)) = "3" Then
'               If CisFun.Val2(.TextMatrix(gCnt1, 6)) = 0 Then
'                   With CisFun
'                       .MB_Lines = 4
'                       .MB_MSG(2) = "    在庫発注点を入力して下さい            "
'                       .MB_Title = ""
'                       .MB_Button = Error
'                       .MBOX
'                   End With
'                    Set ErrObj = iB1_Keisu
'                    GoTo Body1Chk_Err
'                End If
'            End If
'            ' 在庫管理(日数)
'            If RTrim(.TextMatrix(gCnt1, 8)) = "4" Then
'               If CisFun.Val2(.TextMatrix(gCnt1, 6)) = 0 Then
'                   With CisFun
'                       .MB_Lines = 4
'                       .MB_MSG(2) = "    在庫日数を入力して下さい            "
'                       .MB_Title = ""
'                       .MB_Button = Error
'                       .MBOX
'                   End With
'                    Set ErrObj = iB1_Keisu
'                    GoTo Body1Chk_Err
'                End If
'            End If
'
'
'        Next gCnt1
'    End With
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
    
    If Not HKTRead(gSL_Select, 0, 1) Then
        Call HKTClose
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
            Call ItemsClearHKT
            With HKT
                .品番 = RTrim(vsGrid.TextMatrix(gIndex, 9))
                .取引先 = RTrim(vsGrid.TextMatrix(gIndex, 10))
                .受入 = RTrim(vsGrid.TextMatrix(gIndex, 11))
                .工順 = RTrim(vsGrid.TextMatrix(gIndex, 4))
                gStr = Replace(vsGrid.TextMatrix(gIndex, 7), "-", "")
                .置場 = gStr
                .更新者 = gTanto
            End With
        ' ( 更新処理 )
            gSL_Select = " UPDATE 品番工順マスタ SET"
            gSL_Select = gSL_Select & " 置場 = '" & HKT.置場 & "',"
            gSL_Select = gSL_Select & " 更新者 = '" & HKT.更新者 & "',"
            gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
            gSL_Select = gSL_Select & " 更新端末 = HOST_NAME()"
            gSL_Select = gSL_Select & " WHERE 品番 = '" & HKT.品番 & "'"
            gSL_Select = gSL_Select & "   AND 取引先 = '" & HKT.取引先 & "'"
            gSL_Select = gSL_Select & "   AND 受入 = '" & HKT.受入 & "'"
            gSL_Select = gSL_Select & "   AND 工順 = '" & HKT.工順 & "'"
            CisDB.SQL = gSL_Select
            On Error GoTo DBPut_Err
            CisDB.DBExec
            
           ' iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
        End If
    Next gIndex
'
' ( 更新処理 )
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
        Call .FixedSet(中中, 左中, あり, 12, 1, "取引先")
        Call .FixedSet(中中, 中中, あり, 4, 2, "受入")
        Call .FixedSet(中中, 左中, あり, 30, 3, "品  番")
        Call .FixedSet(中中, 中中, あり, 4, 4, "工順")
        Call .FixedSet(中中, 中中, あり, 8, 5, "背番号")
        Call .FixedSet(中中, 中中, あり, 5, 6, "車型")
        Call .FixedSet(中中, 中中, あり, 11, 7, "置場")
        Call .FixedSet(中中, 左中, あり, 24, 8, "品　　名")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 9, "X品番")
        Call .FixedSet(中中, 左中, なし, 0, 10, "X取引先")
        Call .FixedSet(中中, 左中, なし, 0, 11, "X受入")
        Call .FixedSet(中中, 左中, なし, 0, 12, "X置場")
        Call .FixedSet(中中, 左中, なし, 0, 13, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8")
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
    Do Until Not HKT_RDSTS
        'Grid Row,Col
        ID = ID + 1
        With vsGrid
            .TextMatrix(ID, GridCol(0)) = ID
            If RTrim(HKT.取引先名) = "" Then
               .TextMatrix(ID, GridCol(1)) = RTrim(HKT.取引先)
            Else
               .TextMatrix(ID, GridCol(1)) = RTrim(HKT.取引先) & ":" & RTrim(HKT.取引先名)
            End If
            .TextMatrix(ID, GridCol(2)) = RTrim(HKT.受入)
            .TextMatrix(ID, GridCol(3)) = RTrim(HKT.表示品番)
            .TextMatrix(ID, GridCol(4)) = Format(HKT.工順, "##")
            .TextMatrix(ID, GridCol(5)) = RTrim(HKT.背番号)
            .TextMatrix(ID, GridCol(6)) = RTrim(HKT.車型)
            If RTrim(HKT.置場) <> "" Then
               .TextMatrix(ID, GridCol(7)) = Mid(HKT.置場, 1, 2) & "-" & Mid(HKT.置場, 3, 2) & "-" & Mid(HKT.置場, 5, 2) & "-" & Mid(HKT.置場, 7, 2)
            Else
               .TextMatrix(ID, GridCol(7)) = ""
            End If
'            Select Case Len(RTrim(HKT.置場))
'               WHEN 8: .TextMatrix(ID, GridCol(7)) = Mid(RTrim(HKT.置場), 1, 2) & "-" & Mid(RTrim(HKT.置場), 3, 2) & "-" & Mid(RTrim(HKT.置場), 5, 2) & "-" & Mid(RTrim(HKT.置場), 7, 2)
'               WHEN 7: .TextMatrix(ID, GridCol(7)) = Mid(RTrim(HKT.置場), 1, 2) & "-" & Mid(RTrim(HKT.置場), 3, 2) & "-" & Mid(RTrim(HKT.置場), 5, 2) & "-" & Mid(RTrim(HKT.置場), 7, 2)
'               WHEN 6: .TextMatrix(ID, GridCol(7)) = Mid(RTrim(HKT.置場), 1, 2) & "-" & Mid(RTrim(HKT.置場), 3, 2) & "-" & Mid(RTrim(HKT.置場), 5, 4)
'               WHEN 5: .TextMatrix(ID, GridCol(7)) = Mid(RTrim(HKT.置場), 1, 2) & "-" & Mid(RTrim(HKT.置場), 3, 2) & "-" & Mid(RTrim(HKT.置場), 5, 4)
'               WHEN 4: .TextMatrix(ID, GridCol(7)) = Mid(RTrim(HKT.置場), 1, 2) & "-" & Mid(RTrim(HKT.置場), 3, 6)
'               WHEN 3: .TextMatrix(ID, GridCol(7)) = Mid(RTrim(HKT.置場), 1, 2) & "-" & Mid(RTrim(HKT.置場), 3, 6)
'               WHEN 2: .TextMatrix(ID, GridCol(7)) = RTrim(HKT.置場)
'               WHEN 1: .TextMatrix(ID, GridCol(7)) = RTrim(HKT.置場)
'               WHEN 0: .TextMatrix(ID, GridCol(7)) = RTrim(HKT.置場)
'               Else: .TextMatrix(ID, GridCol(7)) = RTrim(HKT.置場)
'            End Select
            .TextMatrix(ID, GridCol(8)) = RTrim(HKT.品名)
            
            .TextMatrix(ID, GridCol(9)) = RTrim(HKT.品番)
            .TextMatrix(ID, GridCol(10)) = RTrim(HKT.取引先)
            .TextMatrix(ID, GridCol(11)) = RTrim(HKT.受入)
            .TextMatrix(ID, GridCol(12)) = RTrim(HKT.置場)
        End With
        Call HKTReadNext(1)
    Loop

    With vsGrid
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    Call HKTClose(1)
    
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
    iB1_Okiba1.Refresh
    iB1_Okiba2.Refresh
    iB1_Okiba3.Refresh
    iB1_Okiba4.Refresh
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
        If RTrim(iB1_Okiba1) & RTrim(iB1_Okiba2) & RTrim(iB1_Okiba3) & RTrim(iB1_Okiba4) <> "" Then
           vsGrid.TextMatrix(tRow, 7) = RTrim(iB1_Okiba1) & Space(2 - Len(iB1_Okiba1)) & "-" & _
                                  RTrim(iB1_Okiba2) & Space(2 - Len(iB1_Okiba2)) & "-" & _
                                  RTrim(iB1_Okiba3) & Space(2 - Len(iB1_Okiba3)) & "-" & _
                                  RTrim(iB1_Okiba4) & Space(2 - Len(iB1_Okiba4))
        Else
           vsGrid.TextMatrix(tRow, 7) = ""
        End If
    '値の変更
        If Replace(RTrim(vsGrid.TextMatrix(tRow, 7)), "-", "") <> RTrim(vsGrid.TextMatrix(tRow, 12)) Then
            If RTrim(vsGrid.TextMatrix(tRow, CisVsGrid.SelectCol)) = "" Then
                vsGrid.Row = tRow
                Call CisVsGrid.vsColor(個別, &HC0E0FF)
                B1lb_Select = Format(CisFun.Val2(B1lb_Select) + 1, "#,###")
                vsGrid.Row = SaveRow
            End If
        Else
            If RTrim(vsGrid.TextMatrix(tRow, CisVsGrid.SelectCol)) <> "" Then
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
        
        gStr = Replace(RTrim(.TextMatrix(tRow, 7)), "-", "")
        iB1_Okiba1 = Mid(gStr, 1, 2)
        iB1_Okiba2 = Mid(gStr, 3, 2)
        iB1_Okiba3 = Mid(gStr, 5, 2)
        iB1_Okiba4 = Mid(gStr, 7, 2)
            
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
    iB1_Okiba1 = "": iB1_Okiba2 = "": iB1_Okiba3 = "": iB1_Okiba4 = "":
    
    Call TextToVsGrid(Val(Input_Area.Tag))

End Sub
Private Sub vsGrid_Scroll()
    'スクロール時はテキストボックスを非表示にします。
'    Call GridTextEnabled(False)
End Sub


