VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CXM0035 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "品番マスタメンテナンス"
   ClientHeight    =   7125
   ClientLeft      =   2550
   ClientTop       =   1245
   ClientWidth     =   16155
   ControlBox      =   0   'False
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   7125
   ScaleWidth      =   16155
   Begin Cis3D_v60.CIS3D Back_Grid4 
      Height          =   2085
      Left            =   3825
      Top             =   4680
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   3678
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   315
         Left            =   15
         Top             =   30
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
         BackColor       =   16711680
         ForeColor       =   16777215
         Caption         =   "※ セット品番入力 ※"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cFont3D         =   2
         cAlingnment     =   7
         Begin Cis3D_v60.CIS3D PB_SetClose 
            Height          =   300
            Left            =   4320
            Top             =   0
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   529
            BackColor       =   8421631
            ForeColor       =   16777215
            Caption         =   "×"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   18
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
      End
      Begin vsFlexLib.vsFlexArray vsGrid4 
         Height          =   1740
         Left            =   30
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   315
         Width           =   4740
         _Version        =   196608
         _ExtentX        =   8361
         _ExtentY        =   3069
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "  |品番                        |取数"
         Rows            =   5
         Cols            =   4
         BackColor       =   12632319
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   0
         RowHeightMin    =   330
         SelectionMode   =   1
         Begin CisText_V60.CisText iB3_SetHinbn 
            Height          =   315
            Left            =   315
            TabIndex        =   44
            Top             =   330
            Width           =   3510
            _ExtentX        =   6191
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
         End
         Begin CisText_V60.CisText iB3_TSu 
            Height          =   330
            Left            =   3765
            TabIndex        =   45
            Top             =   315
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   582
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
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "###"
            cILength        =   3
         End
      End
   End
   Begin Cis3D_v60.CIS3D Back_Kotei 
      Height          =   6840
      Left            =   330
      Top             =   180
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   12065
      BackColor       =   16711680
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      Begin VB.PictureBox PNL_PFKotei 
         BackColor       =   &H00808080&
         Height          =   630
         Left            =   9120
         ScaleHeight     =   570
         ScaleWidth      =   6120
         TabIndex        =   42
         Top             =   30
         Width           =   6180
         Begin CisBtn_60.CisBtn PB_KEND 
            Height          =   585
            Left            =   5250
            Top             =   -15
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
         Begin CisBtn_60.CisBtn PB_KENT 
            Height          =   585
            Left            =   4365
            Top             =   -15
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
         Begin CisBtn_60.CisBtn PB_KDEL 
            Height          =   585
            Left            =   1755
            Top             =   -15
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
         Begin CisBtn_60.CisBtn PB_KLook 
            Height          =   585
            Left            =   2620
            Top             =   -15
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
         Begin CisBtn_60.CisBtn PB_KCAN 
            Height          =   585
            Left            =   3495
            Top             =   -15
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
         Begin CisBtn_60.CisBtn PB_KADD 
            Height          =   585
            Left            =   30
            Top             =   -15
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
            BtnType         =   7
            BtnCaption      =   "新規"
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
         Begin CisBtn_60.CisBtn PB_KMNT 
            Height          =   585
            Left            =   900
            Top             =   -15
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
      Begin Cis3D_v60.CIS3D Back_KoteiS 
         Height          =   1485
         Left            =   90
         Top             =   675
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   2619
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
         cBoderWidth     =   2
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D CIS3D21 
            Height          =   1410
            Left            =   45
            Top             =   45
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   2487
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            Begin Cis3D_v60.CIS3D CIS3D38 
               Height          =   705
               Left            =   14385
               Top             =   0
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "工程数"
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
               Begin CisText_V60.CisText iB3_KoteiSu 
                  Height          =   360
                  Left            =   75
                  TabIndex        =   11
                  Top             =   300
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
            End
            Begin Cis3D_v60.CIS3D CIS3D55 
               Height          =   705
               Left            =   0
               Top             =   705
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "棚卸ﾌﾟﾚｼｰﾄ"
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
               cAlingnment     =   6
               Begin VB.CheckBox B3Chk_Pure 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "発行有"
                  Height          =   240
                  Left            =   150
                  TabIndex        =   12
                  Top             =   345
                  Width           =   1095
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D57 
               Height          =   705
               Left            =   1320
               Top             =   705
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "置　　場"
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
               cAlingnment     =   6
               Begin CisText_V60.CisText iB3_Okiba1 
                  Height          =   375
                  Left            =   45
                  TabIndex        =   13
                  Top             =   285
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
               Begin CisText_V60.CisText iB3_Okiba2 
                  Height          =   375
                  Left            =   375
                  TabIndex        =   14
                  Top             =   285
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
               Begin CisText_V60.CisText iB3_Okiba3 
                  Height          =   375
                  Left            =   705
                  TabIndex        =   15
                  Top             =   285
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
               Begin CisText_V60.CisText iB3_Okiba4 
                  Height          =   375
                  Left            =   1035
                  TabIndex        =   16
                  Top             =   285
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
            Begin Cis3D_v60.CIS3D CIS3D52 
               Height          =   705
               Left            =   7395
               Top             =   0
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "ﾛﾎﾞｯﾄ/単発"
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
               cAlingnment     =   6
               Begin VB.CheckBox B3Chk_Robot 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "ﾛﾎﾞｯﾄ"
                  Height          =   285
                  Left            =   75
                  TabIndex        =   5
                  Top             =   315
                  Width           =   915
               End
               Begin Cis3D_v60.CIS3D B3lb_Daisu 
                  Height          =   270
                  Left            =   975
                  Top             =   315
                  Width           =   420
                  _ExtentX        =   741
                  _ExtentY        =   476
                  BackColor       =   16777152
                  Caption         =   "3台"
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
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D45 
               Height          =   705
               Left            =   2745
               Top             =   705
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "セット品番"
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
               cAlingnment     =   6
               Begin VB.CommandButton PB_SetHin 
                  Caption         =   "SET品番"
                  BeginProperty Font 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Left            =   615
                  TabIndex        =   38
                  Top             =   270
                  Width           =   1260
               End
               Begin VB.CheckBox B3Chk_Set 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "有"
                  Height          =   285
                  Left            =   45
                  TabIndex        =   18
                  Top             =   330
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D36 
               Height          =   705
               Left            =   8835
               Top             =   0
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "工　　程"
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
               cAlingnment     =   6
               cPositionY      =   30
               Begin CisText_V60.CisText iB3_Kotei 
                  Height          =   375
                  Left            =   15
                  TabIndex        =   6
                  Top             =   300
                  Width           =   465
                  _ExtentX        =   820
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
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFbComma        =   0
               End
               Begin Cis3D_v60.CIS3D B3lb_Koteinm 
                  Height          =   360
                  Left            =   465
                  Top             =   300
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "XXXXXXXXX1XXXX"
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
            Begin Cis3D_v60.CIS3D CIS3D29 
               Height          =   705
               Left            =   13455
               Top             =   705
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "日報"
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
               cAlingnment     =   6
               cPositionY      =   30
               Begin VB.CheckBox B3Chk_Nipo 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "○"
                  Height          =   285
                  Left            =   45
                  TabIndex        =   35
                  Top             =   330
                  Width           =   525
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D22 
               Height          =   705
               Left            =   15
               Top             =   0
               Width           =   510
               _ExtentX        =   900
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "工順"
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
               cAlingnment     =   6
               cPositionY      =   30
               Begin CisText_V60.CisText iB3_Kojyun 
                  Height          =   360
                  Left            =   75
                  TabIndex        =   0
                  Top             =   300
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   635
                  cFormat         =   "00"
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
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D23 
               Height          =   705
               Index           =   0
               Left            =   525
               Top             =   0
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "機械（作業）"
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
               cAlingnment     =   6
               cPositionY      =   30
               Begin CisText_V60.CisText iB3_Kikai 
                  Height          =   375
                  Left            =   15
                  TabIndex        =   1
                  Top             =   300
                  Width           =   465
                  _ExtentX        =   820
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
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFbComma        =   0
               End
               Begin Cis3D_v60.CIS3D B3lb_Kikainm 
                  Height          =   360
                  Left            =   465
                  Top             =   300
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "XXXXXXXXX1XXXX"
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
            Begin Cis3D_v60.CIS3D CIS3D25 
               Height          =   705
               Index           =   3
               Left            =   6975
               Top             =   705
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "係数2"
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
               Begin CisText_V60.CisText iB3_KKeisu2 
                  Height          =   360
                  Left            =   30
                  TabIndex        =   24
                  Top             =   300
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   635
                  cFormat         =   "0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "0.00"
                  MaxLength       =   4
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "0.00"
                  cILength        =   1
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D25 
               Height          =   705
               Index           =   2
               Left            =   6360
               Top             =   705
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "係数1"
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
               Begin CisText_V60.CisText iB3_KKeisu1 
                  Height          =   360
                  Left            =   30
                  TabIndex        =   23
                  Top             =   300
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   635
                  cFormat         =   "0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "0.00"
                  MaxLength       =   4
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "0.00"
                  cILength        =   1
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D25 
               Height          =   705
               Index           =   4
               Left            =   7590
               Top             =   705
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "LT"
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
               Begin CisText_V60.CisText iB3_KLT 
                  Height          =   360
                  Left            =   30
                  TabIndex        =   25
                  Top             =   300
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   635
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
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D25 
               Height          =   705
               Index           =   0
               Left            =   4635
               Top             =   705
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "ｻｲｸﾙ"
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
               cAlingnment     =   6
               cPositionY      =   20
               Begin CisText_V60.CisText iB3_KCycle1 
                  Height          =   330
                  Left            =   15
                  TabIndex        =   19
                  Top             =   300
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   582
                  cFormat         =   "#"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#"
                  MaxLength       =   1
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#"
                  cILength        =   1
               End
               Begin CisText_V60.CisText iB3_KCycle2 
                  Height          =   345
                  Left            =   255
                  TabIndex        =   20
                  Top             =   300
                  Width           =   330
                  _ExtentX        =   582
                  _ExtentY        =   609
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
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iB3_KCycle3 
                  Height          =   345
                  Left            =   570
                  TabIndex        =   21
                  Top             =   300
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   609
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
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   705
               Index           =   5
               Left            =   10590
               Top             =   0
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "基準時間①"
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
               Begin CisText_V60.CisText iB3_Kosunm1 
                  Height          =   360
                  Left            =   15
                  TabIndex        =   7
                  Top             =   300
                  Width           =   1050
                  _ExtentX        =   1852
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
               End
               Begin CisText_V60.CisText iB3_Kosu1 
                  Height          =   360
                  Left            =   1050
                  TabIndex        =   8
                  Top             =   300
                  Width           =   825
                  _ExtentX        =   1455
                  _ExtentY        =   635
                  cFormat         =   "#0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#0.000"
                  MaxLength       =   6
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "#0.000"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D25 
               Height          =   705
               Index           =   1
               Left            =   5580
               Top             =   705
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "実工数"
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
               Begin CisText_V60.CisText iB3_JKosu 
                  Height          =   360
                  Left            =   15
                  TabIndex        =   22
                  Top             =   300
                  Width           =   750
                  _ExtentX        =   1323
                  _ExtentY        =   635
                  cFormat         =   "#0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#0.00"
                  MaxLength       =   5
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "#0.00"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D26 
               Height          =   705
               Left            =   7995
               Top             =   705
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "在庫管理"
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
               Begin VB.ComboBox B3Comb_KZaiko 
                  BackColor       =   &H00C0E0FF&
                  Height          =   360
                  Left            =   15
                  Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
                  TabIndex        =   26
                  Top             =   300
                  Width           =   1110
               End
               Begin CisText_V60.CisText iB3_KZaikoDay 
                  Height          =   360
                  Left            =   1155
                  TabIndex        =   29
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   585
                  _ExtentX        =   1032
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
               Begin CisText_V60.CisText iB3_KZaiko 
                  Height          =   360
                  Left            =   1140
                  TabIndex        =   27
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   480
                  _ExtentX        =   847
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
                  cDataType       =   1
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin CisText_V60.CisText iB3_KZaikoL 
                  Height          =   375
                  Left            =   1125
                  TabIndex        =   28
                  Top             =   285
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
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "日"
                  Height          =   240
                  Left            =   1830
                  TabIndex        =   39
                  Top             =   405
                  Visible         =   0   'False
                  Width           =   240
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D27 
               Height          =   705
               Left            =   10140
               Top             =   705
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "在庫管理(最大)"
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
               Begin VB.ComboBox B3Comb_KMZaiko 
                  BackColor       =   &H00C0E0FF&
                  Height          =   360
                  Left            =   30
                  Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
                  TabIndex        =   30
                  Top             =   300
                  Width           =   1170
               End
               Begin CisText_V60.CisText iB3_KMZaikoDay 
                  Height          =   360
                  Left            =   1215
                  TabIndex        =   33
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   585
                  _ExtentX        =   1032
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
               Begin CisText_V60.CisText iB3_KMZaiko 
                  Height          =   360
                  Left            =   1200
                  TabIndex        =   31
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   525
                  _ExtentX        =   926
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
                  cDataType       =   1
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin CisText_V60.CisText iB3_KMZaikoL 
                  Height          =   375
                  Left            =   1200
                  TabIndex        =   32
                  Top             =   300
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
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "日"
                  Height          =   240
                  Left            =   1965
                  TabIndex        =   40
                  Top             =   405
                  Visible         =   0   'False
                  Width           =   240
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D28 
               Height          =   705
               Left            =   12375
               Top             =   705
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "収容数"
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
               Begin CisText_V60.CisText iB3_KSyuyo 
                  Height          =   360
                  Left            =   45
                  TabIndex        =   34
                  Top             =   300
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   635
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
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   705
               Index           =   6
               Left            =   12480
               Top             =   0
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "基準時間②"
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
               Begin CisText_V60.CisText iB3_Kosunm2 
                  Height          =   360
                  Left            =   15
                  TabIndex        =   9
                  Top             =   300
                  Width           =   1050
                  _ExtentX        =   1852
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
               End
               Begin CisText_V60.CisText iB3_Kosu2 
                  Height          =   360
                  Left            =   1050
                  TabIndex        =   10
                  Top             =   300
                  Width           =   825
                  _ExtentX        =   1455
                  _ExtentY        =   635
                  cFormat         =   "#0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#0.000"
                  MaxLength       =   6
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "#0.000"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D53 
               Height          =   705
               Index           =   1
               Left            =   2745
               Top             =   705
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "出荷管理"
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
               cAlingnment     =   6
               Begin VB.CheckBox B3Chk_Syuka 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "○"
                  Height          =   285
                  Left            =   240
                  TabIndex        =   17
                  TabStop         =   0   'False
                  Top             =   330
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   720
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D23 
               Height          =   705
               Index           =   2
               Left            =   2280
               Top             =   0
               Width           =   3360
               _ExtentX        =   5927
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "外注加工"
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
               cAlingnment     =   6
               cPositionY      =   30
               Begin VB.CheckBox B3Chk_NaiGai 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "内"
                  Height          =   285
                  Left            =   60
                  TabIndex        =   2
                  Top             =   345
                  Width           =   570
               End
               Begin Cis3D_v60.CIS3D B3lb_GTornm 
                  Height          =   360
                  Left            =   1605
                  Top             =   300
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "XXXXXXXXX1XXXX"
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
               Begin CisText_V60.CisText iB3_GTorcd 
                  Height          =   375
                  Left            =   630
                  TabIndex        =   3
                  Top             =   300
                  Width           =   990
                  _ExtentX        =   1746
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
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D23 
               Height          =   705
               Index           =   3
               Left            =   5640
               Top             =   0
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   1244
               ForeColor       =   16711680
               Caption         =   "機械№"
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
               cAlingnment     =   6
               cPositionY      =   30
               Begin CisText_V60.CisText iB3_KikaiNo 
                  Height          =   375
                  Left            =   15
                  TabIndex        =   4
                  Top             =   300
                  Width           =   465
                  _ExtentX        =   820
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
                  cDataReplace    =   1
                  cFaZero         =   0
                  cFbComma        =   0
               End
               Begin Cis3D_v60.CIS3D B3lb_KikaiNonm 
                  Height          =   360
                  Left            =   465
                  Top             =   300
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "XXXXXXXXX1XXXX"
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
      End
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   375
         Left            =   -240
         Top             =   -6990
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   661
         BackColor       =   12632319
         ForeColor       =   0
         Caption         =   "【 工    順    管    理  】"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Back_Grid3 
         Height          =   4590
         Left            =   90
         Top             =   2175
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   8096
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   2
         cAlingnment     =   1
         cBoderStyle     =   1
         cPositionX      =   60
         Begin vsFlexLib.vsFlexArray vsGrid3 
            Height          =   4530
            Left            =   45
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   30
            Width           =   15240
            _Version        =   196608
            _ExtentX        =   26882
            _ExtentY        =   7990
            _StockProps     =   228
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   $"CXM0035.frx":0000
            Rows            =   14
            Cols            =   25
            BackColor       =   12648384
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   330
            SelectionMode   =   1
            BackColorAlternate=   12648384
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D32 
         Height          =   540
         Left            =   90
         Top             =   90
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   953
         BackColor       =   12632319
         ForeColor       =   0
         Caption         =   "【 工    順    管    理  】"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   18
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
         Begin Cis3D_v60.CIS3D K_Sykbnm_K 
            Height          =   360
            Left            =   90
            Top             =   90
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   635
            BackColor       =   16761087
            Caption         =   "XXXX"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   15.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
      End
   End
   Begin VB.Timer Timer31 
      Left            =   1980
      Top             =   810
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
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   540
      Width           =   585
   End
   Begin VB.Timer Timer11 
      Left            =   930
      Top             =   660
   End
   Begin VB.Timer Timer21 
      Left            =   1440
      Top             =   720
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 品番マスタメンテナンス  】"
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
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   30
         Width           =   3675
         Begin Cis3D_v60.CIS3D CIS3D15 
            Height          =   345
            Left            =   2460
            Top             =   0
            Visible         =   0   'False
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
            Visible         =   0   'False
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
            Visible         =   0   'False
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
         Visible         =   0   'False
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
         Visible         =   0   'False
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
End
Attribute VB_Name = "CXM0035"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   品番マスタメンテナンス　工順管理画面
'**       フォームID    :   CXM0035
'**       処理概要      :
'**
'**       作  成  日    :   2009/10/16  By CIS SIMPR-A
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    Dim SyoriKB2        As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM2        As String       ' 処理名称(新規･修正･削除 等)
    Dim mCHK            As Boolean
    Dim RowColFlg       As Boolean
    Dim GridBack        As String
    
    Dim DayString()     As String
    
    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
    Dim CisVsGrid       As New CisVsGrid3
    'Dim CisVsGridS      As New CisVsGrid3          2009/07/22
    '+------------------------------------------- insert 2004/12/13
    'Dim CisVsGridK      As New CisVsGrid3          2009/07/22
    '+------------------------------------------- insert 2004/12/13 End
    
    '============================================*** 2009/07/22 Start
    Dim SyoriKB_K       As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM_K       As String       ' 処理名称(新規･修正･削除 等)
    Dim K_Select_Data   As Boolean      ' カーソル位置
    Dim CisVsGrid2      As New CisVsGrid3
    Dim CisVsGrid3      As New CisVsGrid3
    Dim CisVsGrid4      As New CisVsGrid3
    Dim CisComboKZaiko  As CisCombo
    Dim CisComboKMZaiko As CisCombo
    Dim CisComboTest As CisCombo
    '============================================*** 2009/07/22 End
    
    Dim CisComboTehai   As CisCombo
    Dim CisComboZaiko   As CisCombo
'''2007/12/11DEL    Dim CisComboHNou    As CisCombo
'''2007/12/11DEL    Dim CisComboHJyu    As CisCombo
    Dim CisComboHMei    As CisCombo
'''2007/12/11DEL    Dim CisComboJNou    As CisCombo
'''2007/12/11DEL    Dim CisComboJJyu    As CisCombo
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
    Dim CisComboJsya    As CisCombo
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
        
    Dim mIx             As Long
    Dim mIx2            As Long
    Dim mIx3            As Long
    Dim mStr            As String
    
    Dim ActObj          As Object
    Dim ActRow          As Integer
    
    Dim KCol            As Long         ' 非表示項目開始ｶﾗﾑ
    Dim SCol()          As Long         ' 表示項目ｶﾗﾑ

    Dim mTorFlg         As Boolean
    
    Dim mBodyErrFld     As Object
    
    Dim mImage          As String
    Dim maxPNo          As Byte         ' 最大ポケット№ ( 2006/02/16 ADD )

'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 2006/12/19 Update Satart
    Dim BD1Chk_Flg      As Byte
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 2006/12/19 Update End
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 2007/12/11 Update Satart
    Dim SaveKeitai      As String
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 2007/12/11 Update End

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
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
    'Dummy2.Left = -1000                                        2009/07/22
    'Dummy3.Left = -1000                                        2009/07/22
    '+----------------------------------- 2004/12/13 Insert
    'DummyKS.Left = -1000                                       2009/07/22
    'DummyKE.Left = -1000                                       2009/07/22
    '+----------------------------------- 2004/12/13 Insert End
    
'-----------------------------------------------------------' 2009.10.16 del start
''''    Timer1.Interval = 1000   ' 時刻表示間隔
''''    Timer1.Enabled = True    ' 時刻表示開始
''''    Timer2.Interval = 500    ' 処理区分点滅間隔(基本)
''''    Timer2.Enabled = False   ' 処理区分点滅禁止(基本)
''''    Timer3.Interval = 500    ' 処理区分点滅間隔(取引先)
''''    Timer3.Enabled = False   ' 処理区分点滅禁止(取引先)
'-----------------------------------------------------------' 2009.10.16 del end
    
    '+--------------------------------+
    '+  グリッド初期設定(工程詳細情報)
    '+--------------------------------+
    With CisVsGrid3
        Set .GridObj = vsGrid3
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 23
        .InitGet = False
        .Init
    End With
    Call GridInit3
    '+--------------------------------+
    '+  グリッド初期設定(ｾｯﾄ品番情報)
    '+--------------------------------+
    With CisVsGrid4
        Set .GridObj = vsGrid4
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 4
        .InitGet = False
        .Init
    End With
    Call GridInit4

    ' 在庫区分（工程）
    Set CisComboKZaiko = New CisCombo
    With CisComboKZaiko
        Set .Connect = CisDB
        Set .CombObj = B3Comb_KZaiko
        .Control = "名称マスタ"
        .ComboName = "在庫区分"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .Ess = True
        .ComboInit
    End With
    ' 在庫区分最大（工程）
    Set CisComboKMZaiko = New CisCombo
    With CisComboKMZaiko
        Set .Connect = CisDB
        Set .CombObj = B3Comb_KMZaiko
        .Control = "名称マスタ"
        .ComboName = "在庫区分"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .Ess = True
        .ComboInit
    End With
    
    Call CisComboKZaiko.ComboSet        ' 2009.10.19 add
    Call CisComboKMZaiko.ComboSet       ' 2009.10.19 add
    '
    Back_Grid4.Visible = False
' =============================================*** 2009/07/22 End

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
'-----------------------------------------------------------' 2009.10.16 add start
    RV_Top = CXM0030.Top + 450: RV_Left = CXM0030.Left - 300
    If RV_Left <> 0 Or RV_Top <> 0 Then
        Me.Move RV_Left, RV_Top
    Else
        Call CisFrm.Frm_Center(False, x1024y768)
    End If
    Back_Kotei.Visible = True
    Call PB_KADD_Click
'-----------------------------------------------------------' 2009.10.16 add end

'-----------------------------------------------------------' 2009.10.16 del start
''''    If Rv_Call3 = "" Then
''''        Me.Show
''''    End If
'-----------------------------------------------------------' 2009.10.16 del end
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
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
           Case vbKeyF1:
                If PB_KADD.Visible Then
                    Call PB_KADD_Click    '【新規】
                End If
           Case vbKeyF2:
                If PB_KMNT.Visible Then
                    Call PB_KMNT_Click    '【修正】
                End If
           Case vbKeyF3
                               If PB_KDEL.Visible Then Call PB_KDEL_Click       '【削除】
           Case vbKeyF7
                               If PB_KLook.Visible Then Call PB_KLook_Click     '【検索】
           Case vbKeyEscape
                               If PB_KCAN.Visible Then Call PB_KCAN_Click       '【取消】
           Case vbKeyF12
                               If PB_KEND.Visible Then Call PB_KEND_Click       '【終了】
           Case vbKeyReturn
                               If PB_KENT.Visible Then                          '【入力】
                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_KENT_Click
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
' *CHG*  2004/04/05
'    Cancel = 1
    Call PB_KEND_Click
End Sub

'=============================================== *** 2009/07/22 Start
'+--------------------------------------+
'+    工順情報ﾃﾞｰﾀ登録/削除確認         +
'+--------------------------------------+
Private Sub KJ_Delete_Chk(KB As Byte)
    If SyoriKB = "ADD" Or SyoriKB = "MNT" Then
       If KB = 0 Then
           '品番単位
           HNM.品番 = RTrim(CXM0030.iH1_Hinbn)
           If Not HNMRead("", 1) Then
              With CisDB
                  On Error GoTo KJ_DBDelete_Err1
                  .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
          ' ( 削除処理 )
                  ' 品番工順マスタ
                  .SQL = "DELETE FROM 品番工順マスタ "
                  .SQL = .SQL & " WHERE 品番 = '" & RTrim(CXM0030.iH1_Hinbn) & "'"
                  On Error GoTo KJ_DBDelete_Err1
                  .DBExec
                  ' 品番工順セット品番マスタ
                  .SQL = "DELETE FROM 品番工順セット品番マスタ "
                  .SQL = .SQL & " WHERE 品番 = '" & RTrim(CXM0030.iH1_Hinbn) & "'"
                  On Error GoTo KJ_DBDelete_Err1
                  .DBExec
                    
                  .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
                   
                  On Error GoTo 0
              End With
              Exit Sub
KJ_DBDelete_Err1:
              CisFun.ErrorBox
              End
           End If
       End If
       
       If KB = 1 Then
           '品番取引先単位
           gSL_Select = "Select * from 品番取引先マスタ "
           gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(CXM0030.iH1_Hinbn) & "'"
           gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(CXM0030.iB2_Torcd) & "'"
           gSL_Select = gSL_Select & " And   受入 = '" & RTrim(CXM0030.iB2_Ukeir) & "'"
           If CXM0030.B2Op_Hinmoku(0) Then
              gSL_Select = gSL_Select & " And   品目 = 0"
           Else
              gSL_Select = gSL_Select & " And   品目 = 1"
           End If
           If Not HTMRead(gSL_Select, 1) Then
              With CisDB
                  On Error GoTo KJ_DBDelete_Err2
                  .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
          ' ( 削除処理 )
                  ' 品番工順マスタ
                  .SQL = "DELETE FROM 品番工順マスタ "
                  .SQL = .SQL & " Where 品番 = '" & RTrim(CXM0030.iH1_Hinbn) & "'"
                  .SQL = .SQL & " And   取引先 = '" & RTrim(CXM0030.iB2_Torcd) & "'"
                  .SQL = .SQL & " And   受入 = '" & RTrim(CXM0030.iB2_Ukeir) & "'"
                  If CXM0030.B2Op_Hinmoku(0) Then
                     .SQL = .SQL & " And   品目 = 0"
                  Else
                     .SQL = .SQL & " And   品目 = 1"
                  End If
                  On Error GoTo KJ_DBDelete_Err2
                  .DBExec
                  ' 品番工順セット品番マスタ
                  .SQL = "DELETE FROM 品番工順セット品番マスタ "
                  .SQL = .SQL & " Where 品番 = '" & RTrim(CXM0030.iH1_Hinbn) & "'"
                  .SQL = .SQL & " And   取引先 = '" & RTrim(CXM0030.iB2_Torcd) & "'"
                  .SQL = .SQL & " And   受入 = '" & RTrim(CXM0030.iB2_Ukeir) & "'"
                  If CXM0030.B2Op_Hinmoku(0) Then
                     .SQL = .SQL & " And   品目 = 0"
                  Else
                     .SQL = .SQL & " And   品目 = 1"
                  End If
                  On Error GoTo KJ_DBDelete_Err2
                  .DBExec
                  
                  .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
            
              End With
              Exit Sub
KJ_DBDelete_Err2:
              CisFun.ErrorBox
              End
           End If
       End If
       
    End If
End Sub

'+--------------------------------------+
'+    検  索  キ ー：工程情報(F7)       +
'+--------------------------------------+
Private Sub PB_KLook_Click()
    If PB_KLook.Tag = "iB3_Kikai" Then Call Look_Kubun(10): Exit Sub
    If PB_KLook.Tag = "iB3_KikaiNo" Then Call Look_Kubun(11): Exit Sub
    If PB_KLook.Tag = "iB3_Kotei" Then Call Look_Kubun(12): Exit Sub
    If PB_KLook.Tag = "iB3_GTorcd" Then Call Look_GTorcd: Exit Sub
    If PB_KLook.Tag = "iB3_SetHinbn" Then Call Look_SetHinbn: Exit Sub
End Sub
'+----------------------------+
'+    工程画面表示  キ ー     +
'+----------------------------+
Private Sub PB_Kotei_Click()
    Back_Kotei.Visible = True
    
    SyoriKB_K = "ADD"
    SyoriNM_K = "新規"
    Call B3Comb_KZaiko_Click
    Call B3Comb_KMZaiko_Click
    iB3_Kojyun.SetFocus
    Call CisComboKZaiko.ComboSet
    Call CisComboKMZaiko.ComboSet
    Call PB_KADD_Click
       
End Sub
'+----------------------------+
'+    工程画面表示終了キ ー   +
'+----------------------------+
Private Sub PB_KEND_Click()
    Unload CXM0035          ' 2009.10.16 add
End Sub
'+----------------------------+
'+    工程画面　　新規キ ー   +
'+----------------------------+
Private Sub PB_KADD_Click()
    Call Kojyun_SQL_Set
    If HKTRead(gSL_Select, , 1) Then
       Call GridSet3
    End If
        
    PB_KEND.Visible = True
    PB_KCAN.Visible = False
    PB_KADD.Visible = False
    SyoriKB_K = "ADD"
    SyoriNM_K = "新規"
    K_Sykbnm_K = SyoriNM_K
    K_Sykbnm_K.ForeColor = &HFFFF&   '(黄色)
    K_Select_Data = False
End Sub
'+----------------------------+
'+    工程画面　　修正キ ー   +
'+----------------------------+
Private Sub PB_KMNT_Click()
    Back_KoteiS.Enabled = False
    vsGrid3.Enabled = True
    vsGrid3.Row = 1
    vsGrid3.SetFocus
    PB_KEND.Visible = False
    PB_KCAN.Visible = True
    PB_KADD.Visible = False
    PB_KMNT.Visible = False
    PB_KDEL.Visible = False
    PB_KLook.Visible = False
    SyoriKB_K = "MNT"
    SyoriNM_K = "修正"
    K_Sykbnm_K = SyoriNM_K
    K_Sykbnm_K.ForeColor = &HFF0000  '(青色)
    K_Select_Data = True
End Sub
'+----------------------------+
'+    工程画面　　削除キ ー   +
'+----------------------------+
Private Sub PB_KDEL_Click()
    Back_KoteiS.Enabled = False
    vsGrid3.Enabled = True
    vsGrid3.Row = 1
    vsGrid3.SetFocus
    PB_KEND.Visible = False
    PB_KCAN.Visible = True
    PB_KADD.Visible = False
    PB_KMNT.Visible = False
    PB_KDEL.Visible = False
    PB_KLook.Visible = False
    SyoriKB_K = "DEL"
    SyoriNM_K = "削除"
    K_Sykbnm_K = SyoriNM_K
    K_Sykbnm_K.ForeColor = &HFF&     '(赤色)
    K_Select_Data = True
End Sub
'+----------------------------+
'+    工程画面　　取消キ ー   +
'+----------------------------+
Private Sub PB_KCAN_Click()
    If K_Select_Data Then
        Back_KoteiS.Enabled = True
        PB_KEND.Visible = True
        PB_KCAN.Visible = False
        PB_KMNT.Visible = True
        PB_KDEL.Visible = True
        iB3_Kojyun.Enabled = True
        iB3_Kojyun.SetFocus
        SyoriKB_K = "ADD"
        SyoriNM_K = "新規"
        K_Sykbnm_K = SyoriNM_K
        K_Sykbnm_K.ForeColor = &HFFFF&   '(黄色)
        K_Select_Data = False
    Else
        Call Kotei_Area_Clear
        Call PB_KMNT_Click
    End If
End Sub
'+------------------------------+
'+    工程画面　　ｾｯﾄ品番入力   +
'+------------------------------+
Private Sub PB_SetHin_Click()
'    PB_KEND.Visible = False
'    PB_KCAN.Visible = False
'    PB_KENT.Visible = False
'    PB_KADD.Visible = False
'    PB_KMNT.Visible = False
'    PB_KDEL.Visible = False
    PNL_PFKotei.Visible = False
    
    Back_KoteiS.Enabled = False
    Back_Grid3.Enabled = False
    Back_Grid4.Visible = True
    vsGrid4.Row = 1
    
    iB3_SetHinbn.SetFocus
End Sub
Private Sub PB_SetClose_Click()
    PNL_PFKotei.Visible = True
    
    Call VSGrid4_Click
    Back_KoteiS.Enabled = True
    Back_Grid3.Enabled = True
    Back_Grid4.Visible = False
    vsGrid4.Row = 1

    If SyoriKB_K = "ADD" Then
       iB3_Kojyun.SetFocus
    Else
       iB3_Kikai.SetFocus
    End If
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_KENT_Click()
   Call ReturnPress_K
End Sub
'+----------------------------+
'+       外注取引先検索       +
'+----------------------------+
Private Sub Look_GTorcd()

    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 1
    
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
       iB3_GTorcd = RV_TorcdK
       B3lb_GTornm = RV_TorRName
    End If
End Sub
'+----------------------------+
'+       品番検索              +
'+----------------------------+
Private Sub Look_SetHinbn()

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RVI_Hinmoku = 9
    
    CKK0030.Show vbModal
    Unload CKK0030
    Set CKK0030 = Nothing

    If RV_Rtn Then
       iB3_SetHinbn = RTrim(RV_Hinbn)
    End If
End Sub
'+-----------------------------+
'+  B3Comb_KZaiko
'+-----------------------------+
Private Sub B3Comb_KZaiko_GotFocus()
    B3Comb_KZaiko.BackColor = gIGotSel
End Sub
Private Sub B3Comb_KZaiko_LostFocus()
    B3Comb_KZaiko.BackColor = gILostSel
End Sub
Private Sub B3Comb_KZaiko_Click()
    CisComboKZaiko.CodeGet
    gStr = CisComboKZaiko.Code
    Select Case gStr
        Case "2"    ' 係数
            iB3_KZaiko.Visible = True
            iB3_KZaikoDay.Visible = False
            iB3_KZaikoL.Visible = False
            CXM0030.Zaiko_Label.Left = iB3_KZaiko.Left + iB3_KZaiko.Width + 60
            CXM0030.Zaiko_Label = "％"
            CXM0030.Zaiko_Label.Visible = True
        Case "3"    ' 発注点
            iB3_KZaiko.Visible = False
            iB3_KZaikoDay.Visible = False
            iB3_KZaikoL.Visible = True
        Case "4"    ' 日数
            iB3_KZaiko.Visible = False
            iB3_KZaikoDay.Visible = True
            iB3_KZaikoL.Visible = False
            CXM0030.Zaiko_Label.Left = iB3_KZaikoDay.Left + iB3_KZaikoDay.Width + 60
            CXM0030.Zaiko_Label = "日"
            CXM0030.Zaiko_Label.Visible = True
        Case Else
            iB3_KZaiko.Visible = False
            iB3_KZaikoDay.Visible = False
            iB3_KZaikoL.Visible = False
            CXM0030.Zaiko_Label.Visible = False
    End Select
End Sub
'+-----------------------------+
'+  B3Comb_KZaiko
'+-----------------------------+
Private Sub B3Comb_KMZaiko_GotFocus()
    B3Comb_KMZaiko.BackColor = gIGotSel
End Sub
Private Sub B3Comb_KMZaiko_LostFocus()
    B3Comb_KMZaiko.BackColor = gILostSel
End Sub
Private Sub B3Comb_KMZaiko_Click()
    CisComboKMZaiko.CodeGet
    gStr = CisComboKMZaiko.Code
    Select Case gStr
        Case "2"    ' 係数
            iB3_KMZaiko.Visible = True
            iB3_KMZaikoDay.Visible = False
            iB3_KMZaikoL.Visible = False
            CXM0030.Zaiko_Label.Left = iB3_KMZaiko.Left + iB3_KMZaiko.Width + 60
            CXM0030.Zaiko_Label = "％"
            CXM0030.Zaiko_Label.Visible = True
        Case "3"    ' 発注点
            iB3_KMZaiko.Visible = False
            iB3_KMZaikoDay.Visible = False
            iB3_KMZaikoL.Visible = True
        Case "4"    ' 日数
            iB3_KMZaiko.Visible = False
            iB3_KMZaikoDay.Visible = True
            iB3_KMZaikoL.Visible = False
            CXM0030.Zaiko_Label.Left = iB3_KMZaikoDay.Left + iB3_KMZaikoDay.Width + 60
            CXM0030.Zaiko_Label = "日"
            CXM0030.Zaiko_Label.Visible = True
        Case Else
            iB3_KMZaiko.Visible = False
            iB3_KMZaikoDay.Visible = False
            iB3_KMZaikoL.Visible = False
            CXM0030.Zaiko_Label.Visible = False
    End Select
End Sub
'+-----------------------------+
'+  B3Chk_NaiGai　内外区分
'+-----------------------------+
Private Sub B3Chk_NaiGai_Click()
    If B3Chk_NaiGai.Value = 0 Then
       B3Chk_NaiGai.Caption = "内"
    Else
       B3Chk_NaiGai.Caption = "外"
    End If
    If B3Chk_NaiGai.Value = 1 Then
       iB3_GTorcd.Enabled = True
    Else
       iB3_GTorcd.Enabled = False
       iB3_GTorcd = "": B3lb_GTornm = ""
    End If
End Sub
'+-----------------------------+
'+  B3Chk_Robot　ﾛﾎﾞｯﾄ/単発区分
'+-----------------------------+
Private Sub B3Chk_Robot_Click()
    If B3Chk_Robot.Value = 0 Then
       B3Chk_Robot.Caption = "ﾛﾎﾞｯﾄ"
       Call Robot_Dsp
'       gSL_Select = "Select 台数 From 取引先工程マスタ "
'       gSL_Select = gSL_Select & " Where 取引先CD = '" & RTrim(CXM0030.iB2_Torcd) & "'"
'       'gSL_Select = gSL_Select & " And   機械 = '" & RTrim(CXM0030.iB3_Kijyun) & "'"
'       'gSL_Select = gSL_Select & " And   機械NO = '" & RTrim(CXM0030.iB3_Kijyun) & "'"
'       gSL_Select = gSL_Select & " And   機械 = '" & RTrim(iB3_Kikai) & "'"
'       gSL_Select = gSL_Select & " And   機械NO = '" & RTrim(iB3_KikaiNo) & "'"
'       If TKTRead(gSL_Select, 1) Then
'          If TKT.台数 <> 0 Then
'             B3lb_Daisu = Format(TKT.台数, "#0") & "台"
'          Else
'             B3lb_Daisu = ""
'          End If
'       Else
'             B3lb_Daisu = ""
'       End If
    Else
       B3Chk_Robot.Caption = "単発"
       B3lb_Daisu = ""
    End If
End Sub
'+-----------------------------+
'+  ﾛﾎﾞｯﾄ台数表示
'+-----------------------------+
Private Sub Robot_Dsp()

    gSL_Select = "Select 台数 From 取引先工程マスタ "
    gSL_Select = gSL_Select & " Where 取引先CD = '" & RTrim(CXM0030.iB2_Torcd) & "'"
    'gSL_Select = gSL_Select & " And   機械 = '" & RTrim(CXM0030.iB3_Kijyun) & "'"
    'gSL_Select = gSL_Select & " And   機械NO = '" & RTrim(CXM0030.iB3_Kijyun) & "'"
    gSL_Select = gSL_Select & " And   機械 = '" & RTrim(iB3_Kikai) & "'"
    gSL_Select = gSL_Select & " And   機械NO = '" & RTrim(iB3_KikaiNo) & "'"
    If TKTRead(gSL_Select, 1) Then
       If TKT.台数 <> 0 Then
          B3lb_Daisu = Format(TKT.台数, "#0") & "台"
       Else
          B3lb_Daisu = ""
       End If
    Else
          B3lb_Daisu = ""
    End If

End Sub
'+-----------------------------+
'+  B3Chk_Pure　棚卸ﾌﾟﾚｼｰﾄ印刷区分
'+-----------------------------+
Private Sub B3Chk_Pure_Click()
    If B3Chk_Pure.Value = 0 Then
       B3Chk_Pure.Caption = "発行有"
    Else
       B3Chk_Pure.Caption = "発行無"
    End If
End Sub
'+-----------------------------+
'+  B3Chk_Syuka　出荷管理
'+-----------------------------+
Private Sub B3Chk_Syuka_Click()
    If B3Chk_Syuka.Value = 0 Then
       B3Chk_Syuka.Caption = "○"
    Else
       B3Chk_Syuka.Caption = "×"
    End If
End Sub
'+-----------------------------+
'+  B3Chk_Set　ｾｯﾄ品番
'+-----------------------------+
Private Sub B3Chk_Set_Click()
    If B3Chk_Set.Value = 0 Then
       B3Chk_Set.Caption = "無"
    Else
       B3Chk_Set.Caption = "有"
    End If
    If B3Chk_Set.Value = 1 Then
       PB_SetHin.Enabled = True
    Else
       PB_SetHin.Enabled = False
    End If
End Sub
'+-----------------------------+
'+  B3Chk_Nipo　日報管理
'+-----------------------------+
Private Sub B3Chk_Nipo_Click()
    If B3Chk_Nipo.Value = 0 Then
       B3Chk_Nipo.Caption = "○"
    Else
       B3Chk_Nipo.Caption = "×"
    End If
End Sub
'+---------------------------------------------+
'+  iB3_Kikai    機械
'+---------------------------------------------+
Private Sub iB3_Kikai_GotFocus()
    PB_KLook.Tag = ActiveControl.Name
    PB_KLook.Visible = True
End Sub
Private Sub iB3_Kikai_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_KLook.Tag = ""
        PB_KLook.Visible = False
    End If
    
    B3lb_Kikainm = ""
    If RTrim(iB3_Kikai) <> "" Then
       Call MeisyoGet("機械", RTrim(iB3_Kikai))
       B3lb_Kikainm = RTrim(SYM_Meisyo)
       
       Call Robot_Dsp
    End If
End Sub
'+-----------------------------+
'+  iB3_GTorcd
'+-----------------------------+
Private Sub iB3_GTorcd_GotFocus()
    PB_KLook.Tag = ActiveControl.Name
    PB_KLook.Visible = True
End Sub
Private Sub iB3_GTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_KLook.Tag = ""
        PB_KLook.Visible = False
    End If
    B3lb_GTornm = ""
    If TorNmGet(iB3_GTorcd, 8) Then
        B3lb_GTornm = TRM_RName
    End If
End Sub
'+---------------------------------------------+
'+  iB3_KikaiNo    機械№
'+---------------------------------------------+
Private Sub iB3_KikaiNo_GotFocus()
    PB_KLook.Tag = ActiveControl.Name
    PB_KLook.Visible = True
End Sub
Private Sub iB3_KikaiNo_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_KLook.Tag = ""
        PB_KLook.Visible = False
    End If
    
    B3lb_KikaiNonm = ""
    If RTrim(iB3_KikaiNo) <> "" Then
       Call MeisyoGet("機械NO", RTrim(iB3_KikaiNo))
       B3lb_KikaiNonm = RTrim(SYM_Meisyo)
       
       Call Robot_Dsp
    End If
End Sub
'+-----------------------------+
'+  iB3_Kotei
'+-----------------------------+
Private Sub iB3_Kotei_GotFocus()
    PB_KLook.Tag = ActiveControl.Name
    PB_KLook.Visible = True
End Sub
Private Sub iB3_Kotei_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_KLook.Tag = ""
        PB_KLook.Visible = False
    End If
    B3lb_Koteinm = ""
    If MeisyoGet("工程管理", iB3_Kotei) Then
        B3lb_Koteinm = SYM_Meisyo
    End If
End Sub
'**********************************
'*      入力処理_工順(Enter)      *
'**********************************
Private Sub ReturnPress_K()
    Call CisFrm.MousePT(11)
        
    If K_Select_Data Then
       'カーソルがGrid内
        If SyoriKB_K = "DEL" Then
           If CisFun.Val2(vsGrid3.TextMatrix(vsGrid3.Row, 0)) <> 0 Then
              If Not DBDelete_K(CisFun.Val2(vsGrid3.TextMatrix(vsGrid3.Row, 0))) Then GoTo ReturnPress_K_Ed
'              Call PB_KDEL_Click
           End If
        Else
           If CisFun.Val2(vsGrid3.TextMatrix(vsGrid3.Row, 0)) <> 0 Then
              Call GridToText3(CisFun.Val2(vsGrid3.TextMatrix(vsGrid3.Row, 0)))
           
               Back_KoteiS.Enabled = True
               vsGrid3.Enabled = False
               PB_KEND.Visible = False
               PB_KCAN.Visible = True
               PB_KADD.Visible = False
               PB_KMNT.Visible = False
               PB_KDEL.Visible = False
               K_Select_Data = False
               
               iB3_Kikai.SetFocus
           End If
        End If
    Else
       'カーソルがGrid外
        If Not BodyKChk Then
           GoTo ReturnPress_K_Ed
        End If
        
        If Not DBPut_K Then
            GoTo ReturnPress_K_Ed
        End If
        Call Kotei_Area_Clear
        
        Call Kojyun_SQL_Set
        If HKTRead(gSL_Select, , 1) Then
           Call GridSet3
        End If
        
        If SyoriKB_K = "MNT" Then
           Call PB_KMNT_Click
        End If
        
        GoTo ReturnPress_K_Ed
    End If

ReturnPress_K_Ed:
   Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+      削 除 処 理 (工順)    +
'+----------------------------+
Private Function DBDelete_K(wKojyun As Byte)
    DBDelete_K = False
'  ( 確認メッセージ )
    CXM0030.K_Sykbnm2.Caption = SyoriNM2
    CXM0030.K_Sykbnm2.Visible = True
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     工順：" & Format(wKojyun, "#0") & " の削除処理を行います      "
        .MB_Title = "削除処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    With CisDB
        On Error GoTo DBDelete_K_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
        ' 品番工順マスタ
        .SQL = "DELETE FROM 品番工順マスタ "
        .SQL = .SQL & " Where 品番 = '" & RTrim(CXM0030.iH1_Hinbn) & "'"
        .SQL = .SQL & " And   取引先 = '" & RTrim(CXM0030.iB2_Torcd) & "'"
        .SQL = .SQL & " And   受入 = '" & RTrim(CXM0030.iB2_Ukeir) & "'"
        If CXM0030.B2Op_Hinmoku(0) Then
           .SQL = .SQL & " And   品目 = 0"
        Else
           .SQL = .SQL & " And   品目 = 1"
        End If
        .SQL = .SQL & " And   工順 = " & wKojyun
        On Error GoTo DBDelete_K_Err
        .DBExec
        ' 品番工順セット品番マスタ
        .SQL = "DELETE FROM 品番工順セット品番マスタ "
        .SQL = .SQL & " Where 品番 = '" & RTrim(CXM0030.iH1_Hinbn) & "'"
        .SQL = .SQL & " And   取引先 = '" & RTrim(CXM0030.iB2_Torcd) & "'"
        .SQL = .SQL & " And   受入 = '" & RTrim(CXM0030.iB2_Ukeir) & "'"
        If CXM0030.B2Op_Hinmoku(0) Then
           .SQL = .SQL & " And   品目 = 0"
        Else
           .SQL = .SQL & " And   品目 = 1"
        End If
        .SQL = .SQL & " And   工順 = " & wKojyun
        On Error GoTo DBDelete_K_Err
        .DBExec
        
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
        Call Kojyun_SQL_Set
        If HKTRead(gSL_Select, , 1) Then
           Call GridSet3
           Call PB_KDEL_Click
        Else
           CisVsGrid3.Clear
           Call PB_KCAN_Click
'           Call PB_KADD_Click
        End If
        
'        gLong = CisFun.Val2(iB2_Torcd.Tag)
'        With vsGrid1
'             .RemoveItem gLong
'             If .Rows < CisVsGrid.DispMax Then
'                .Rows = CisVsGrid.DispMax
'             End If
'        End With
'        CisVsGrid.ReNumber
        
        On Error GoTo 0
    End With
    DBDelete_K = True
    Exit Function
DBDelete_K_Err:
    CisFun.ErrorBox
    End
End Function

'+----------------------------+
'+    工順 追加･更新 処理     +
'+----------------------------+
Private Function DBPut_K()
    DBPut_K = False
    Dim AddFlg  As Boolean

    With CisFun
        .MB_Lines = 4
        If SyoriKB_K = "ADD" Then
            .MB_MSG(2) = "     追加処理を行います      "
            .MB_Title = "追加処理"
        Else
            .MB_MSG(2) = "     修正処理を行います      "
            .MB_Title = "修正処理"
        End If
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
'---------------------------------------' 2009.10.19 add start
           If SyoriKB_K = "ADD" Then
              iB3_Kojyun.SetFocus
           Else
              iB3_Kikai.SetFocus
           End If
'---------------------------------------' 2009.10.19 add end
          'iB3_Kikai.SetFocus           ' 2009.10.19 del
           Exit Function
        End If
    End With
    
    On Error GoTo DBPut_K_Err
'   マスタ項目移送
    gSL_Select = "SELECT * From 品番工順マスタ"
    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(CXM0030.iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(CXM0030.iB2_Torcd) & "'"
    gSL_Select = gSL_Select & " And   受入 = '" & RTrim(CXM0030.iB2_Ukeir) & "'"
    If CXM0030.B2Op_Hinmoku(0) Then
       gSL_Select = gSL_Select & " And   品目 = 0"
    Else
       gSL_Select = gSL_Select & " And   品目 = 1"
    End If
    gSL_Select = gSL_Select & " And   工順 = " & CisFun.Val2(iB3_Kojyun)
    
    If HKTRead(gSL_Select, 1) Then
       AddFlg = False
    Else
       Call ItemsClearHKT
       AddFlg = True
       HKT.品番 = RTrim(CXM0030.iH1_Hinbn)
    End If
    
    With HKT
        .品番 = RTrim(CXM0030.iH1_Hinbn)
        .取引先 = RTrim(CXM0030.iB2_Torcd)
        .受入 = RTrim(CXM0030.iB2_Ukeir)
        If CXM0030.B2Op_Hinmoku(0) Then
           .品目 = 0
        Else
           .品目 = 1
        End If
        .工順 = CisFun.Val2(iB3_Kojyun)
        
        .機械 = RTrim(iB3_Kikai)
        .外注加工 = B3Chk_NaiGai
        If .外注加工 = 1 Then
           .外注加工先CD = RTrim(iB3_GTorcd)
        Else
           .外注加工先CD = ""
        End If
        .機械NO = RTrim(iB3_KikaiNo)
        .ロボット区分 = B3Chk_Robot
        .工程 = RTrim(iB3_Kotei)
        .基準工数名1 = RTrim(iB3_Kosunm1)
        .基準工数1 = CisFun.Val2(iB3_Kosu1)
        .基準工数名2 = RTrim(iB3_Kosunm2)
        .基準工数2 = CisFun.Val2(iB3_Kosu2)
        .工程数 = CisFun.Val2(iB3_KoteiSu)
        .棚卸プレシート発行 = B3Chk_Pure.Value
        .置場 = RTrim(iB3_Okiba1) & Space(2 - Len(iB3_Okiba1)) & _
                RTrim(iB3_Okiba2) & Space(2 - Len(iB3_Okiba2)) & _
                RTrim(iB3_Okiba3) & Space(2 - Len(iB3_Okiba3)) & _
                RTrim(iB3_Okiba4) & Space(2 - Len(iB3_Okiba4))
         .出荷管理 = B3Chk_Syuka.Value
         .セット品番 = B3Chk_Set.Value
         If CisFun.Val2(iB3_KCycle1) <> 0 Or CisFun.Val2(iB3_KCycle2) <> 0 Or _
            CisFun.Val2(iB3_KCycle3) <> 0 Then
            .サイクル = Format(iB3_KCycle1, "0") & _
                          Format(iB3_KCycle2, "00") & _
                          Format(iB3_KCycle3, "00")
         Else
            .サイクル = ""
         End If
         .実工数 = CisFun.Val2(iB3_JKosu)
         .係数1 = CisFun.Val2(iB3_KKeisu1)
         .係数2 = CisFun.Val2(iB3_KKeisu2)
         .LT = CisFun.Val2(iB3_KLT)
         CisComboKZaiko.CodeGet
         .在庫管理区分 = CisFun.Val2(CisComboKZaiko.Code)
         .在庫係数 = CisFun.Val2(iB3_KZaiko)
         .在庫発注点 = CisFun.Val2(iB3_KZaikoL)
         .在庫日数 = CisFun.Val2(iB3_KZaikoDay)
         Select Case .在庫管理区分
              Case 2: .在庫管理数値 = .在庫係数
              Case 3: .在庫管理数値 = .在庫発注点
              Case 4: .在庫管理数値 = .在庫日数
         End Select
         CisComboKMZaiko.CodeGet
         .最大在庫管理区分 = CisFun.Val2(CisComboKMZaiko.Code)
         .最大在庫係数 = CisFun.Val2(iB3_KMZaiko)
         .最大在庫発注点 = CisFun.Val2(iB3_KMZaikoL)
         .最大在庫日数 = CisFun.Val2(iB3_KMZaikoDay)
         Select Case .最大在庫管理区分
              Case 2: .最大在庫管理数値 = .最大在庫係数
              Case 3: .最大在庫管理数値 = .最大在庫発注点
              Case 4: .最大在庫管理数値 = .最大在庫日数
         End Select
         .収容数 = CisFun.Val2(iB3_KSyuyo)
         .日報出力 = B3Chk_Nipo.Value
    End With

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 追加処理 )
    If AddFlg Then
        HKT.作成者 = gTanto
        Call HKTInsert
    Else
' ( 更新処理 )
        HKT.更新者 = gTanto
        Call HKTUpdate
    End If

'品番工順セット品番マスタ
    With CisDB
        .SQL = "DELETE FROM 品番工順セット品番マスタ "
        .SQL = .SQL & " Where 品番 = '" & RTrim(CXM0030.iH1_Hinbn) & "'"
        .SQL = .SQL & " And   取引先 = '" & RTrim(CXM0030.iB2_Torcd) & "'"
        .SQL = .SQL & " And   受入 = '" & RTrim(CXM0030.iB2_Ukeir) & "'"
        If CXM0030.B2Op_Hinmoku(0) Then
           .SQL = .SQL & " And   品目 = 0"
        Else
           .SQL = .SQL & " And   品目 = 1"
        End If
        .SQL = .SQL & " And   工順 = " & CisFun.Val2(iB3_Kojyun)
        On Error GoTo DBPut_K_Err
        .DBExec
    End With

    HSE.品番 = RTrim(CXM0030.iH1_Hinbn)
    HSE.取引先 = RTrim(CXM0030.iB2_Torcd)
    HSE.受入 = RTrim(CXM0030.iB2_Ukeir)
    If CXM0030.B2Op_Hinmoku(0) Then
       HSE.品目 = 0
    Else
       HSE.品目 = 1
    End If
    HSE.工順 = CisFun.Val2(iB3_Kojyun)
    With vsGrid4
         For mIx = 1 To .Rows - 1
             If Trim(.TextMatrix(mIx, 1)) <> "" Then
                HSE.セット品番 = Trim(.TextMatrix(mIx, 1))
                HSE.取数 = CisFun.Val2(Trim(.TextMatrix(mIx, 2)))
                HSE.作成者 = gTanto
                
                Call HSEInsert
             End If
         Next mIx
    End With
        
'
    Call CisDB.DBTran(TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    DBPut_K = True
    Exit Function
DBPut_K_Err:
    CisFun.ErrorBox
    End
End Function
Private Sub Kotei_Area_Clear()
    iB3_Kojyun = ""
    iB3_Kikai = ""
    B3lb_Kikainm = ""
    B3Chk_NaiGai.Value = 0
    Call B3Chk_NaiGai_Click
    iB3_GTorcd = ""
    B3lb_GTornm = ""
    iB3_KikaiNo = ""
    B3lb_KikaiNonm = ""
    B3Chk_Robot.Value = 0
    Call B3Chk_Robot_Click
    B3lb_Daisu = ""
    iB3_Kotei = ""
    B3lb_Koteinm = ""
    iB3_Kosunm1 = ""
    iB3_Kosu1 = ""
    iB3_Kosunm2 = ""
    iB3_Kosu2 = ""
    iB3_KoteiSu = ""
    B3Chk_Pure.Value = 0
    Call B3Chk_Pure_Click
    iB3_Okiba1 = ""
    iB3_Okiba2 = ""
    iB3_Okiba3 = ""
    iB3_Okiba4 = ""
    B3Chk_Syuka.Value = 0
    Call B3Chk_Syuka_Click
    B3Chk_Set.Value = 0
    Call B3Chk_Set_Click
    iB3_KCycle1 = ""
    iB3_KCycle2 = ""
    iB3_KCycle3 = ""
    iB3_JKosu = ""
    iB3_KKeisu1 = ""
    iB3_KKeisu2 = ""
    iB3_KLT = ""
 If B3Comb_KZaiko.ListCount > 0 Then
    B3Comb_KZaiko.ListIndex = 0
    Call B3Comb_KZaiko_Click
End If
If B3Comb_KZaiko.ListCount > 0 Then
    B3Comb_KMZaiko.ListIndex = 0
    Call B3Comb_KMZaiko_Click
End If
    iB3_KSyuyo = ""
    B3Chk_Nipo.Value = 0
    Call B3Chk_Nipo_Click
End Sub
'=============================================== *** 2009/07/22 End

'+----------------------------+
'+       名称マスタ検索        +
'+    KB : 0     工程         +
'+       : 1     次工程       +
'+       : 3     作業　　      +
'+       : 4     収容器(基本) +
'+       : 5     収容器1      +
'+       : 6     収容器2      +
'+       : 7     伝票種類発注 +
'+       : 8     伝票種類受注 +
'+       : 9     品番区分     +
'+       : 0以外 担当         +
'+----------------------------+
Private Sub Look_Kubun(KB As Integer)

    RV_Left = 0
    RV_Top = 0

    Select Case KB
        Case 0: RVI_Kubun = "工程管理"
        Case 1: RVI_Kubun = "工程管理"
        Case 3: RVI_Kubun = "作業管理"
        Case 4, 5, 6: RVI_Kubun = "収容器管理"
        Case 7: RVI_Kubun = "伝票種類発注"
        Case 8: RVI_Kubun = "伝票種類受注"
        Case 9: RVI_Kubun = "品番区分"
        Case 10: RVI_Kubun = "機械"                     '2009/07/22
        Case 11: RVI_Kubun = "機械NO"                   '2009/07/22
        Case 12: RVI_Kubun = "工程管理"                 '2009/07/22
        Case Else: RVI_Kubun = "担当管理"
    End Select
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        Select Case KB
            '''Case 0
            '''    iB1_Kotei = RV_Code
            '''    Call iB1_Kotei_LostFocus
            '''Case 1
            '''    iB1_NKotei = RV_Code
            '''    Call iB1_NKotei_LostFocus
            '=================================== *** 2009/07/22 Start
            Case 10
                iB3_Kikai = RV_Code
                Call iB3_Kikai_LostFocus
            Case 11
                iB3_KikaiNo = RV_Code
                Call iB3_KikaiNo_LostFocus
            Case 12
                iB3_Kotei = RV_Code
                Call iB3_Kotei_LostFocus
            '=================================== *** 2009/07/22 End
        End Select
    End If

End Sub
'=========================================== *** 2009/07/22 Start
''+---------------------------------------------------------- 2004/12/13 Insert
''+----------------------------+
''+       機械検索             +
''+----------------------------+
'Private Sub Look_Kikai()
'
'    RV_Left = 0
'    RV_Top = 0
'
'    CKP0610.Show vbModal
'    Unload CKP0610
'    Set CKP0610 = Nothing
'
'    If RV_Rtn Then
'       iB3_Kikai = RV_Code
'       Call iB3_Kikai_LostFocus
'    End If
'
'End Sub
''+---------------------------------------------------------- 2004/12/13 Insert End
'=========================================== *** 2009/07/22 End

'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
'+-----------------------+
'+      基本情報
'+-----------------------+
Private Sub SyoriName(SYKB As String)
'
End Sub
'+-----------------------+
'+      取引先情報
'+-----------------------+
Private Sub SyoriName2(SYKB As String)
'
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
'*******************************
'*      処理区分名点滅制御      *
'*******************************
Private Sub Timer3_Timer()
   If CXM0030.K_Sykbnm2.Caption <> "" Then
      CXM0030.K_Sykbnm2.Caption = ""
   Else
      CXM0030.K_Sykbnm2.Caption = SyoriNM2
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
    Dim B2Mode      As Boolean
    Dim B2Color     As Long
    Dim B3Mode      As Boolean
    Dim B3Color     As Long
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
    B2Mode = False
    B2Color = gPLostSel
    B3Mode = False
    B3Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
    
    B3Color = gPGotSel          ' 2009.10.16 add
    
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
    RowColFlg = True
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
    If HeadBody = "B1" Or HeadBody = "B" Then
        CisVsGrid.Clear
        CisVsGrid2.Clear
        CisVsGrid3.Clear
        CisVsGrid4.Clear
        'Call CisVsGridS.Clear                                      2009/07/22
        '+----------------------------------- 2004/12/13 Insert
        'Call CisVsGridK.Clear                                      2009/07/22
        '+----------------------------------- 2004/12/13 Insert End
    End If
    If HeadBody = "B3" Then
        'Call CisVsGridS.EditEnable("C")                            2009/07/22
        'Call CisVsGridS.Clear                                      2009/07/22
        '+----------------------------------- 2004/12/13 Insert
        'Call CisVsGridK.EditEnable("C")                            2009/07/22
        'Call CisVsGridK.Clear                                      2009/07/22
        '+----------------------------------- 2004/12/13 Insert End
        CisVsGrid2.Clear
        CisVsGrid3.Clear
        CisVsGrid4.Clear
''''        Call B3Comb_Tehai_Click                 '2007/12/14 ADD     ' 2009.10.19 del
''''        B3Op_HinKu(0).Value = True              '2009/07/22
    End If
    If HeadBody = "B4" Then
        iB3_SetHinbn = "": iB3_TSu = ""
        CisVsGrid4.Clear
    End If
    RowColFlg = False
    Exit Sub
HeadBodyClear_ER:
    RowColFlg = False
End Sub

'================================================================ *** 2009/07/22 Start
'**************************************
'*      Back_Koteiﾁｪｯｸ処理(工順情報)
'**************************************
Private Function BodyKChk() As Boolean
    Dim ReadFlg     As Boolean
    Dim SetHinFlg   As Boolean
    BodyKChk = False

    If SyoriKB_K = "ADD" Then
        ' 工順チェック
        If RTrim(iB3_Kojyun) = "" Then
            With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    工順を入力して下さい            "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
            End With
            iB3_Kojyun.SetFocus
            Exit Function
        End If
        '  工順重複ﾁｪｯｸ
        With vsGrid3
            For mIx = 1 To .Rows - 1
                If Trim(.TextMatrix(mIx, 0)) <> "" Then
                   If CisFun.Val2(Trim(.TextMatrix(mIx, 0))) = CisFun.Val2(iB3_Kojyun) Then
                       With CisFun
                            .MB_Lines = 4
                            .MB_MSG(2) = "    工順が重複しています         "
                            .MB_Title = ""
                            .MB_Button = Error
                            .MBOX
                       End With
                       iB3_Kojyun.SetFocus
                       Exit Function
                   End If
                End If
            Next mIx
        End With
     End If
 
 ' 機械(作業)/機械№(仕様) どちらか必須入力
    If RTrim(iB3_Kikai) = "" And RTrim(iB3_KikaiNo) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    機械/機械№どちらかを入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB3_Kikai.SetFocus
        Exit Function
    End If
    If RTrim(iB3_Kikai) <> "" Then
       If Not MeisyoGet("機械", RTrim(iB3_Kikai)) Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    機械(作業)が未登録です            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB3_Kikai.SetFocus
           Exit Function
       Else
           B3lb_Kikainm = RTrim(SYM_Meisyo)
       End If
    End If
    If RTrim(iB3_KikaiNo) <> "" Then
       If Not MeisyoGet("機械NO", RTrim(iB3_KikaiNo)) Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    機械№(仕様)が未登録です            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB3_KikaiNo.SetFocus
           Exit Function
       Else
           B3lb_KikaiNonm = RTrim(SYM_Meisyo)
       End If
    End If
    
    '外注加工先
    If B3Chk_NaiGai.Value = 1 Then
       If RTrim(iB3_GTorcd) = "" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    外注加工先を入力して下さい            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB3_GTorcd.SetFocus
           Exit Function
       End If
       B3lb_GTornm = ""
       If Not TorNmGet(iB3_GTorcd, 8) Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    外注加工先が未登録です            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB3_GTorcd.SetFocus
           Exit Function
       Else
           B3lb_GTornm = TRM_RName
       End If
    End If
    
    '工程
    If Trim(iB3_Kotei) <> "" Then
        B3lb_Koteinm = ""
        If MeisyoGet("工程管理", iB3_Kotei) Then
            B3lb_Koteinm = SYM_Meisyo
        Else
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    工程が未登録です            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iB3_Kotei.SetFocus
           Exit Function
        End If
    End If
    
    ' ｾｯﾄ品番
    If B3Chk_Set.Value = 1 Then
       SetHinFlg = False
       With vsGrid4
            gCnt1 = 0: gCnt2 = 0
            For mIx = 1 To .Rows - 1
                If Trim(.TextMatrix(mIx, 1)) <> "" Then
                   SetHinFlg = True
                   Exit For
                End If
            Next mIx
            If Not SetHinFlg Then
                Call PB_SetHin_Click
                With CisFun
                     .MB_Lines = 4
                     .MB_MSG(2) = "    ｾｯﾄ品番を入力して下さい           "
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                  End With
                  .Row = 1
'                  .Row = mIx
                  GoTo SetHin_Err
            End If
            
            For mIx = 1 To .Rows - 1
                If Trim(.TextMatrix(mIx, 1)) = "" Then
                   .TextMatrix(mIx, 2) = ""
                Else
                   If Trim(.TextMatrix(mIx, 2)) = "" Then
                      Call PB_SetHin_Click
                      With CisFun
                           .MB_Lines = 4
                           .MB_MSG(2) = "    取数を入力して下さい           "
                           .MB_Title = ""
                           .MB_Button = Error
                           .MBOX
                        End With
                        .Row = mIx
                        GoTo SetHin_Err
                   End If
                   HNM.品番 = Trim(.TextMatrix(mIx, 1))
                   If Not HNMRead("") Then
                       Call PB_SetHin_Click
                       With CisFun
                          .MB_Lines = 4
                          .MB_MSG(2) = "    品番マスタ未登録            "
                          .MB_Title = ""
                          .MB_Button = Error
                          .MBOX
                       End With
                       .Row = mIx
                       GoTo SetHin_Err
                   End If
                End If
            Next mIx
    
    '        ' 品番重複チェック
            For mIx = 1 To .Rows - 1
                If Trim(.TextMatrix(mIx, 1)) <> "" Then
                    For mIx2 = mIx + 1 To .Rows - 1
                        If Trim(.TextMatrix(mIx2, 1)) <> "" Then
                           If Trim(.TextMatrix(mIx, 1)) = Trim(.TextMatrix(mIx2, 1)) Then
                              Call PB_SetHin_Click
                              With CisFun
                                  .MB_Lines = 4
                                  .MB_MSG(2) = "    品番が重複してます。            "
                                  .MB_Title = ""
                                  .MB_Button = Error
                                  .MBOX
                               End With
                               .Row = mIx2
                               GoTo SetHin_Err
                           End If
                        End If
                    Next mIx2
                End If
            Next mIx
       End With
    End If
    
    ' 在庫管理
    With CisComboKZaiko
        .CodeGet
        gStr = .Code
    End With
    If gStr = "2" Then
       If CisFun.Val2(iB3_KZaiko) = 0 Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    在庫係数を入力して下さい            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB3_KZaiko.SetFocus
           Exit Function
        End If
    End If
    If gStr = "3" Then
       If CisFun.Val2(iB3_KZaikoL) = 0 Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    在庫発注点を入力して下さい            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB3_KZaikoL.SetFocus
           Exit Function
        End If
    End If
    If gStr = "4" Then
       If CisFun.Val2(iB3_KZaikoDay) = 0 Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    在庫日数を入力して下さい            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB3_KZaikoDay.SetFocus
           Exit Function
        End If
    End If
    
    BodyKChk = True
    Exit Function
Chk_Err:
    Exit Function
SetHin_Err:
    Call CisVsGrid4.EditEnable
    Call CisVsGrid4.EditVisible
    iB3_SetHinbn.SetFocus
    Exit Function
End Function
'================================================================ *** 2009/07/22 End

'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
'========================================================== *** 2009/07/22 Start
Private Sub Kojyun_SQL_Set()
    ' 工順情報
    gSL_Select = "SELECT HK.*,ISNULL(SY1.値名称,'') 機械名,ISNULL(SY2.値名称,'') 機械NO名,"
    gSL_Select = gSL_Select & " ISNULL(SY3.値名称,'') 工程名,ISNULL(TR.略称,'') 外注加工名 "
    gSL_Select = gSL_Select & " From 品番工順マスタ HK"
    gSL_Select = gSL_Select & " Left Outer join 名称マスタ SY1"
    gSL_Select = gSL_Select & " On  SY1.区分名称 = '機械' "
    gSL_Select = gSL_Select & " And SY1.値 = HK.機械"
    gSL_Select = gSL_Select & " And SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & " Left Outer join 名称マスタ SY2"
    gSL_Select = gSL_Select & " On  SY2.区分名称 = '機械NO' "
    gSL_Select = gSL_Select & " And SY2.値 = HK.機械NO"
    gSL_Select = gSL_Select & " And SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & " Left Outer join 名称マスタ SY3"
    gSL_Select = gSL_Select & " On  SY3.区分名称 = '工程管理' "
    gSL_Select = gSL_Select & " And SY3.値 = HK.工程"
    gSL_Select = gSL_Select & " And SY3.区分タイプ = ''"
    gSL_Select = gSL_Select & " Left Outer join 取引先マスタ TR"
    gSL_Select = gSL_Select & " On  TR.取引先CD = HK.外注加工先CD "
    gSL_Select = gSL_Select & " And TR.取引先区分 = 1"
        
    gSL_Select = gSL_Select & " WHERE HK.品番 = '" & RTrim(CXM0030.iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   HK.取引先 = '" & RTrim(CXM0030.iB2_Torcd) & "'"
    gSL_Select = gSL_Select & " And   HK.受入 = '" & RTrim(CXM0030.iB2_Ukeir) & "'"
    If CXM0030.B2Op_Hinmoku(0) Then
       gSL_Select = gSL_Select & " And   HK.品目 = 0"
    Else
       gSL_Select = gSL_Select & " And   HK.品目 = 1"
    End If
    gSL_Select = gSL_Select & " ORDER BY HK.工順"
End Sub
Private Sub Kojyun_SQL_Set2()
    ' 工順情報(vsGrid2用)
    gSL_Select = "SELECT HK.*,ISNULL(SY1.値名称,'') 機械名,ISNULL(SY2.値名称,'') 機械NO名,"
    gSL_Select = gSL_Select & " ISNULL(SY3.値名称,'') 工程名,ISNULL(TR.略称,'') 外注加工名 "
    gSL_Select = gSL_Select & " From 品番工順マスタ HK"
    gSL_Select = gSL_Select & " Left Outer join 名称マスタ SY1"
    gSL_Select = gSL_Select & " On  SY1.区分名称 = '機械' "
    gSL_Select = gSL_Select & " And SY1.値 = HK.機械"
    gSL_Select = gSL_Select & " And SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & " Left Outer join 名称マスタ SY2"
    gSL_Select = gSL_Select & " On  SY2.区分名称 = '機械NO' "
    gSL_Select = gSL_Select & " And SY2.値 = HK.機械NO"
    gSL_Select = gSL_Select & " And SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & " Left Outer join 名称マスタ SY3"
    gSL_Select = gSL_Select & " On  SY3.区分名称 = '工程管理' "
    gSL_Select = gSL_Select & " And SY3.値 = HK.工程"
    gSL_Select = gSL_Select & " And SY3.区分タイプ = ''"
    gSL_Select = gSL_Select & " Left Outer join 取引先マスタ TR"
    gSL_Select = gSL_Select & " On  TR.取引先CD = HK.外注加工先CD "
    gSL_Select = gSL_Select & " And TR.取引先区分 = 1"
        
    gSL_Select = gSL_Select & " WHERE HK.品番 = '" & RTrim(CXM0030.iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   HK.取引先 = '" & RTrim(CXM0030.iB2_Torcd) & "'"
    gSL_Select = gSL_Select & " And   HK.受入 = '" & RTrim(CXM0030.iB2_Ukeir) & "'"
    If CXM0030.B2Op_Hinmoku(0) Then
       gSL_Select = gSL_Select & " And   HK.品目 = 0"
    Else
       gSL_Select = gSL_Select & " And   HK.品目 = 1"
    End If
    gSL_Select = gSL_Select & " ORDER BY HK.工順"
End Sub
'========================================================== *** 2009/07/22 End
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+--------------------------------------+
'+      作業情報　更新
'+--------------------------------------+
Public Function DBPutSagyo() As Boolean
    Dim AddFlg  As Boolean
    DBPutSagyo = False
'====================================================== *** 2009/07/22 Start
' 取引先作業情報更新
'    With vsGridS
'        gSL_Select = "Update 品番取引作業マスタ Set 削除区分 = 'D'"
'        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(iH1_Hinbn) & "'"
'        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(iB2_Torcd) & "'"
'        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(iB2_Ukeir) & "'"
'        CisDB.SQL = gSL_Select
'        Call CisDB.DBExec
'
'        For mIx = 1 To .Rows - 1
'            If Trim(.TextMatrix(mIx, 0)) <> "" Or Trim(.TextMatrix(mIx, 1)) <> "" Then
'                gSL_Select = "Select * from 品番取引作業マスタ "
'                gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(iH1_Hinbn) & "'"
'                gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(iB2_Torcd) & "'"
'                gSL_Select = gSL_Select & " And   受入 = '" & RTrim(iB2_Ukeir) & "'"
'                gSL_Select = gSL_Select & " And   作業順 = " & CisFun.Val2(.TextMatrix(mIx, 0))
'                If HSMRead(gSL_Select, 1) Then
'                   AddFlg = False
'                Else
'                   AddFlg = True
'                   HSM.品番 = RTrim(iH1_Hinbn)
'                   HSM.取引先 = RTrim(iB2_Torcd)
'                   HSM.受入 = RTrim(iB2_Ukeir)
'                   HSM.作業順 = CisFun.Val2(.TextMatrix(mIx, 0))
'                End If
'
'                HSM.作業 = Trim(.TextMatrix(mIx, 1))
'                HSM.備考 = RTrim(.TextMatrix(mIx, 3))
'                HSM.削除区分 = ""
'
'                If AddFlg Then
'                   Call HSMInsert
'                Else
'                   Call HSMUpdate
'                End If
'            End If
'        Next mIx
'        With CisDB
'            .SQL = "delete 品番取引作業マスタ "
'            .SQL = .SQL & " Where 削除区分 = 'D'"
'            .DBExec
'        End With
'    End With
'====================================================== *** 2009/07/22 End
    DBPutSagyo = True
End Function
'+---------------------------------------------- 2004/12/13 Insert
'+--------------------------------------+
'+      機械情報　更新
'+--------------------------------------+
Public Function DBPutKikai() As Boolean
    Dim AddFlg  As Boolean
    DBPutKikai = False

'====================================================== *** 2009/07/22 Start
' 取引先機械情報更新
'    With VsGridK
'        gSL_Select = "Update 品番取引機械マスタ Set 削除区分 = 'D'"
'        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(iH1_Hinbn) & "'"
'        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(iB2_Torcd) & "'"
'        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(iB2_Ukeir) & "'"
'        CisDB.SQL = gSL_Select
'        Call CisDB.DBExec
'
'        For mIx = 1 To .Rows - 1
'            If Trim(.TextMatrix(mIx, 1)) <> "" Then
'                gSL_Select = "Select * from 品番取引機械マスタ "
'                gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(iH1_Hinbn) & "'"
'                gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(iB2_Torcd) & "'"
'                gSL_Select = gSL_Select & " And   受入 = '" & RTrim(iB2_Ukeir) & "'"
'                gSL_Select = gSL_Select & " And   機械 = '" & Trim(.TextMatrix(mIx, 1)) & "'"
'                If HKMRead(gSL_Select, 1) Then
'                   AddFlg = False
'                Else
'                   AddFlg = True
'                   HKM.品番 = RTrim(iH1_Hinbn)
'                   HKM.取引先 = RTrim(iB2_Torcd)
'                   HKM.受入 = RTrim(iB2_Ukeir)
'                   HKM.機械 = Trim(.TextMatrix(mIx, 1))
'                End If
'
'                If Trim(.TextMatrix(mIx, 3)) <> "" Then
'                    HKM.基本 = "1"
'                    HKM.機械備考 = RTrim(iB3_SJBiko)
'                Else
'                    HKM.基本 = ""
'                    HKM.機械備考 = ""
'                End If
'
'                HKM.削除区分 = ""
'
'                If AddFlg Then
'                   Call HKMInsert
'                Else
'                   Call HKMUpdate
'                End If
'            End If
'        Next mIx
'        With CisDB
'            .SQL = "delete 品番取引機械マスタ "
'            .SQL = .SQL & " Where 品番 = '" & RTrim(iH1_Hinbn) & "'"
'            .SQL = .SQL & "   And 取引先 = '" & RTrim(iB2_Torcd) & "'"
'            .SQL = .SQL & "   And 受入 = '" & RTrim(iB2_Ukeir) & "'"
'            .SQL = .SQL & "   And 削除区分 = 'D'"
'            .DBExec
'        End With
'    End With
'====================================================== *** 2009/07/22 End
    DBPutKikai = True
End Function
'+---------------------------------------------- 2004/12/13 Insert End

'+------------------------------+
'+  収容器　更新
'+------------------------------+
Private Function DBPutYouki(Youki As String) As Boolean
    
    If Trim(Youki) = "" Then Exit Function
        
    gSL_Select = "Select * from 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '収容器管理'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値 = '" & RTrim(Youki) & "'"
    If SYMRead(gSL_Select, 1, 1) Then Exit Function
    
    Call ItemsClearSYM
    With SYM
        .区分名称 = "収容器管理"
        .値 = RTrim(Youki)
        .区分桁数 = 10
        .作成者 = gTanto
    End With
    Call SYMInsert

End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'=============================================== *** 2009/07/22 Start
''+==========================================================+
''+      グリッド処理(作業情報)                               +
''+==========================================================+
''+----------------------------------------------------------+
''+   見出しｾｯﾄ & 初期状態ｾｯﾄ
''+----------------------------------------------------------+
'Private Sub GridInitS()
'    With CisVsGridS
''　ｸﾞﾘｯﾄﾞ初期化処理
''        .Clear
''       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
''                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
'        Call .FixedSet(中中, 右中, あり, 2, 0, "No")
'        Call .FixedSet(中中, 左中, あり, 4, 1, "作業")
'        Call .FixedSet(中中, 左中, あり, 9, 2, "作 業 名")
'        Call .FixedSet(中中, 左中, あり, 14, 3, "備考")
''　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
'        Call .FixedSet(中中, 左中, なし, 0, 4, "Select")
''  表示ﾊﾟﾀｰﾝｾｯﾄ
''       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
'        Call .PatanSet(0, "0,1,2,3")
''  編集内容セット
''       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iB3_No, 0, 0, iB3_No.cFormat)
'        Call .EditSet(iB3_Sagyo, 1)
'        Call .EditSet(iB3_SBiko, 3)
'        .InitDisp
'    End With
'End Sub
'
'
''+----------------------+
''+  VSGridS_Click
''+----------------------+
'Private Sub VSGridS_Click()
'    Call CisVsGridS.EditEnable
'    Call CisVsGridS.EditVisible
'End Sub
''+---------------------------------------------------------+
''+          グリッド編集エリア
''+---------------------------------------------------------+
''+----------------------+
''+  iB3_No
''+----------------------+
'Private Sub iB3_No_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call CisVsGridS.KeyEdit(iB3_No, KeyCode, Shift)
'End Sub
''+----------------------+
''+  iB3_Sagyo
''+----------------------+
'Private Sub iB3_Sagyo_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call CisVsGridS.KeyEdit(iB3_Sagyo, KeyCode, Shift)
'End Sub
''+-----------------------------+
''+  iB1_NKotei
''+-----------------------------+
'Private Sub iB3_Sagyo_GotFocus()
'    PB_Look.Tag = ActiveControl.Name
'    PB_Look.Visible = True
'End Sub
'Private Sub iB3_Sagyo_LostFocus()
'    If ActiveControl.Name <> "PB_Look" Then
'        PB_Look.Tag = ""
'        PB_Look.Visible = False
'    End If
'    With vsGridS
'        .TextMatrix(.Row, 2) = ""
'        If MeisyoGet("作業管理", iB3_Sagyo) Then
'            .TextMatrix(.Row, 2) = SYM_Meisyo
'        End If
'    End With
'End Sub
''+----------------------+
''+  iB3_SBiko
''+----------------------+
'Private Sub iB3_SBiko_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call CisVsGridS.KeyEdit(iB3_SBiko, KeyCode, Shift)
'End Sub
''+----------------------+
''+  Dummy2
''+----------------------+
'Private Sub iB3_Biko_LostFocus()
'    Set ActObj = iB3_Biko
'End Sub
''+----------------------+
''+  Dummy2
''+----------------------+
'Private Sub Dummy2_GotFocus()
'    Call CisVsGridS.KeyEdit(iB3_No, vbKeyDown, 0)
'End Sub
''+----------------------+
''+  Dummy3
''+----------------------+
'Private Sub Dummy3_GotFocus()
'    If Not ActObj Is Nothing Then
'       Set ActObj = Nothing
'       iB3_No.SetFocus
'       Exit Sub
'    End If
'    If vsGridS.Row = 1 Then
'        '+------------------------------- 2004/12/13 Delete
''       If iB3_Biko.Enabled Then
''          iB3_Biko.SetFocus
''       End If
'        '+------------------------------- 2004/12/13 Delete End
'        '+------------------------------- 2004/12/13 Insert
'       If iB3_SJBiko.Enabled Then
'          iB3_SJBiko.SetFocus
'       End If
'        '+------------------------------- 2004/12/13 Insert End
'       Exit Sub
'    End If
'    Call CisVsGridS.KeyEdit(iB3_SBiko, vbKeyUp, 0)
'End Sub
''+----------------------+
''+  iB3_SJBiko
''+----------------------+
'Private Sub iB3_SJBiko_LostFocus()
'    Set ActObj = iB3_SJBiko
'End Sub
''+----------------------+
''+  VSGrid1_Scroll
''+----------------------+
'Private Sub vsGrid1_Scroll()
'    'スクロール時はテキストボックスを非表示にします。
'    Call CisVsGrid.EditEnable
'End Sub
''+------------------------------------------------------------ 2004/12/13 Insert
''+==========================================================+
''+      グリッド処理(機械情報)                               +
''+==========================================================+
''+==========================================================+
''+      グリッド処理                                         +
''+==========================================================+
''+----------------------------------------------------------+
''+   見出しｾｯﾄ & 初期状態ｾｯﾄ
''+----------------------------------------------------------+
'Private Sub GridInitK()
'    With CisVsGridK
''　ｸﾞﾘｯﾄﾞ初期化処理
''        .Clear
''       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
''                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
'        Call .FixedSet(中中, 右中, あり, 0, 0, "No")
'        Call .FixedSet(中中, 左中, あり, 0, 1, "機械№")
'        Call .FixedSet(中中, 左中, あり, 17, 2, "機械名")
'        Call .FixedSet(中中, 中中, あり, 0, 3, "基本")
''　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
'        Call .FixedSet(中中, 左中, なし, 0, 4, "Select")
''  表示ﾊﾟﾀｰﾝｾｯﾄ
''       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
'        Call .PatanSet(0, "0,1,2,3")
''  編集内容セット
''       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iB3_Kikai, 1)
'        .InitDisp
'    End With
'End Sub
''+----------------------+
''+  VSGridK_Click
''+----------------------+
'Private Sub VSGridK_Click()
'    Call CisVsGridK.EditEnable
'    Call CisVsGridK.EditVisible
'    Call iB3_Kikai_LostFocus
'End Sub
'Private Sub VsGridK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    With VsGridK
'        Call iB3_Kikai_LostFocus
'
'        If .MouseCol <> 3 Then Exit Sub
'        If Trim(.TextMatrix(.MouseRow, 1)) <> "" Then
'            For gLong = 1 To .Rows - 1
'                If gLong = .MouseRow And .MouseCol = 3 And Trim(.TextMatrix(.MouseRow, 1)) <> "" Then
'                   If Trim(.TextMatrix(.MouseRow, 3)) = "" Then
'                      .TextMatrix(.MouseRow, 3) = "○"
'                   Else
'                      .TextMatrix(.MouseRow, 3) = ""
'                   End If
'                Else
'                  .TextMatrix(gLong, 3) = ""
'                End If
'            Next gLong
'        End If
'    End With
'End Sub
''+---------------------------------------------------------+
''+          グリッド編集エリア
''+---------------------------------------------------------+
''+----------------------+
''+  iB3_Kikai
''+----------------------+
'Private Sub iB3_Kikai_KeyDown(KeyCode As Integer, Shift As Integer)
''    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
''       With VsGridK
''           .TextMatrix(.Row, 2) = ""
''           If KMMMeisyoGet(iB3_Kikai) Then
''               .TextMatrix(.Row, 2) = KMM_Meisyo
''           End If
''       End With
''    End If
'
'    Call CisVsGridK.KeyEdit(iB3_Kikai, KeyCode, Shift)
'End Sub
'Private Sub iB3_Kikai_GotFocus()
'    PB_Look.Tag = ActiveControl.Name
'    PB_Look.Visible = True
'End Sub
'Private Sub iB3_Kikai_LostFocus()
'    Set ActObj = iB3_Kikai
'    If ActiveControl.Name <> "PB_Look" Then
'        PB_Look.Tag = ""
'        PB_Look.Visible = False
'    End If
'    With VsGridK
'        .TextMatrix(.Row, 2) = ""
'        If KMMMeisyoGet(iB3_Kikai) Then
'            .TextMatrix(.Row, 2) = KMM_Meisyo
'        End If
'    End With
'End Sub
'''+----------------------+
'''+  DummyKS
'''+----------------------+
'Private Sub DummyKS_GotFocus()
'    If ActObj Is iB3_Biko Then
'       Set ActObj = Nothing
'       iB3_Kikai.SetFocus
'       Exit Sub
'    End If
'    If VsGridK.Row = 1 Then
'       If iB3_Biko.Enabled Then
'          iB3_Biko.SetFocus
'       End If
'       Exit Sub
'    End If
'    Call CisVsGridK.KeyEdit(iB3_Kikai, vbKeyUp, 0)
'End Sub
''+----------------------+
''+  DummyKE
''+----------------------+
'Private Sub DummyKE_GotFocus()
'    If ActObj Is iB3_Kikai Then
'       If Trim(iB3_Kikai) = "" Then
'          On Error Resume Next
'          If iB3_SJBiko.Enabled Then
'             iB3_SJBiko.SetFocus
'          End If
'          On Error GoTo 0
'          Exit Sub
'       End If
'       Call CisVsGridK.KeyEdit(iB3_Kikai, vbKeyDown, 0)
'    End If
'    If Not ActObj Is Nothing Then
'       Set ActObj = Nothing
'       iB3_Kikai.SetFocus
'       Exit Sub
'    End If
'End Sub
'''2007/12/11DEL+----------------------+
'''2007/12/11DEL+  VSGrid1_Scroll
'''2007/12/11DEL+----------------------+
''Private Sub vsGrid1_Scroll()
''    'スクロール時はテキストボックスを非表示にします。
''    Call CisVsGrid.EditEnable
''End Sub
'+------------------------------------------------------------ 2004/12/13 Insert End
'=============================================== *** 2009/07/22 End

'+==========================================================+
'+      グリッド処理(工程詳細情報)                          +
'+==========================================================+
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit3()
    With CisVsGrid3
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
'       Call .FixedSet(中中, 右中, あり, 0, 0, "No")
'|出荷管理|ｾｯﾄ品番|サイクル|実工数|係数1|係数2|LT|在庫管理         |在庫管理(最大)  |収容数  |日報
        Call .FixedSet(中中, 左中, あり, 4, 0, "工順")
        Call .FixedSet(中中, 左中, あり, 12, 1, "機械(作業)")
        Call .FixedSet(中中, 中中, あり, 4, 2, "内外")
        Call .FixedSet(中中, 左中, あり, 12, 3, "外注加工")
        Call .FixedSet(中中, 左中, あり, 12, 4, "機械№(仕様)")
        Call .FixedSet(中中, 中中, あり, 10, 5, "ﾛﾎﾞｯﾄ/単発")
        Call .FixedSet(中中, 左中, あり, 12, 6, "工程")
        Call .FixedSet(中中, 中中, あり, 15, 7, "基準工数1")
        Call .FixedSet(中中, 中中, あり, 15, 8, "基準工数2")
        Call .FixedSet(中中, 右中, あり, 9, 9, "工程数")
        Call .FixedSet(中中, 中中, あり, 6, 10, "ﾌﾟﾚｼｰﾄ")
        Call .FixedSet(中中, 左中, あり, 11, 11, "置場")
        Call .FixedSet(中中, 中中, あり, 4, 12, "出荷")
        Call .FixedSet(中中, 中中, あり, 7, 13, "ｾｯﾄ品番")
        Call .FixedSet(中中, 中中, あり, 5, 14, "ｻｲｸﾙ")
        Call .FixedSet(中中, 右中, あり, 6, 15, "実工数")
        Call .FixedSet(中中, 右中, あり, 5, 16, "係数1")
        Call .FixedSet(中中, 右中, あり, 5, 17, "係数2")
        Call .FixedSet(中中, 右中, あり, 2, 18, "LT")
        Call .FixedSet(中中, 中中, あり, 10, 19, "在庫管理")
        Call .FixedSet(中中, 中中, あり, 10, 20, "最大在庫")
        Call .FixedSet(中中, 右中, あり, 7, 21, "収容数")
        Call .FixedSet(中中, 中中, あり, 4, 22, "日報")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 23, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22")
        .InitDisp
    End With
End Sub
'***********************************************'
'*****       グリッド内容セット(工順情報)  *****
'***********************************************'
Private Sub GridSet3()

    Dim ID          As Long
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
     
    CisVsGrid3.Clear

' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid3
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid3.DispMax Then
            .Rows = CisVsGrid3.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not HKT_RDSTS
        ID = ID + 1
            
        Call HKTInfoToGrid(ID)
        
        Call HKTReadNext(1)
    Loop
    With vsGrid3
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'+-------------------------------------+
'+  品番取引先情報よりグリッドへセット    +
'+-------------------------------------+
Private Sub HKTInfoToGrid(SetRow As Long)
    With vsGrid3
        For mIx = 0 To .Cols - 1
            .TextMatrix(SetRow, mIx) = ""
        Next mIx

        ReDim SCol(12)                              ' 2009.10.19 add
        For mIx = 1 To 12: SCol(mIx) = mIx: Next    ' 2009.10.19 add
        
        ' 工順
        .TextMatrix(SetRow, 0) = Format(HKT.工順, "##0")
        ' 機械(作業)
        If RTrim(HKT.機械) <> "" Then
           .TextMatrix(SetRow, SCol(1)) = RTrim(HKT.機械) & ":" & RTrim(HKT.機械名)
        Else
           .TextMatrix(SetRow, SCol(1)) = ""
        End If
        ' 内外
        If HKT.外注加工 = 0 Then
           .TextMatrix(SetRow, SCol(2)) = "内"
        Else
           .TextMatrix(SetRow, SCol(2)) = "外"
        End If
        ' 外注加工先
        If HKT.外注加工 = 1 Then
           .TextMatrix(SetRow, SCol(3)) = RTrim(HKT.外注加工先CD) & "：" & RTrim(HKT.外注加工名)
        Else
           .TextMatrix(SetRow, SCol(3)) = ""
        End If
        ' 機械NO(仕様)
        If RTrim(HKT.機械NO) <> "" Then
           .TextMatrix(SetRow, SCol(4)) = RTrim(HKT.機械NO) & ":" & RTrim(HKT.機械NO名)
        Else
           .TextMatrix(SetRow, SCol(4)) = ""
        End If
        ' ﾛﾎﾞｯﾄ/単発
        If HKT.ロボット区分 = 0 Then
           .TextMatrix(SetRow, SCol(5)) = "ロボット"
        Else
           .TextMatrix(SetRow, SCol(5)) = "単発"
        End If
        ' 工程
        If RTrim(HKT.工程) <> "" Then
           .TextMatrix(SetRow, SCol(6)) = RTrim(HKT.工程) & ":" & RTrim(HKT.工程名)
        Else
           .TextMatrix(SetRow, SCol(6)) = ""
        End If
        
        ' 基準時間1
        If RTrim(HKT.基準工数名1) <> "" Then
           .TextMatrix(SetRow, SCol(7)) = RTrim(HKT.基準工数名1) & ":" & Format(HKT.基準工数1, "#0.000")
        Else
           .TextMatrix(SetRow, SCol(7)) = Format(HKT.基準工数1, "#0.000")
        End If
        ' 基準時間2
        If RTrim(HKT.基準工数名2) <> "" Then
           .TextMatrix(SetRow, SCol(8)) = RTrim(HKT.基準工数名2) & ":" & Format(HKT.基準工数2, "#0.000")
        Else
           .TextMatrix(SetRow, SCol(8)) = Format(HKT.基準工数2, "#0.000")
        End If
        ' 工程数
        .TextMatrix(SetRow, SCol(9)) = Format(HKT.工程数, "#0.0")
        ' 棚卸ﾌﾟﾚｼｰﾄ
        If HKT.棚卸プレシート発行 = 0 Then
           .TextMatrix(SetRow, SCol(10)) = "発行有"
        Else
           .TextMatrix(SetRow, SCol(10)) = "発行無"
        End If
        .Redraw = True
        
        ' 置場
        If RTrim(HKT.置場) <> "" Then
           .TextMatrix(SetRow, SCol(11)) = Mid(HKT.置場, 1, 2) & "-" & Mid(HKT.置場, 3, 2) & "-" & Mid(HKT.置場, 5, 2) & "-" & Mid(HKT.置場, 7, 2)
        Else
           .TextMatrix(SetRow, SCol(11)) = ""
        End If
        ' 出荷管理
        If HKT.出荷管理 = 0 Then
           .TextMatrix(SetRow, SCol(12)) = "○"
        Else
           .TextMatrix(SetRow, SCol(12)) = "×"
        End If
        ' ｾｯﾄ品番
        If HKT.セット品番 = 0 Then
           .TextMatrix(SetRow, 13) = "無"
        Else
           .TextMatrix(SetRow, 13) = "有"
        End If
        ' サイクル
        If Trim(HKT.サイクル) <> "" Then
            .TextMatrix(SetRow, 14) = Mid(HKT.サイクル, 1, 1) & "-" & _
                                     Mid(HKT.サイクル, 2, 2) & "-" & _
                                     Mid(HKT.サイクル, 4, 2)
        End If
        ' 実工数
        If HKT.実工数 <> 0 Then
           .TextMatrix(SetRow, 15) = Format(HKT.実工数, "#0.00")
        Else
           .TextMatrix(SetRow, 15) = ""
        End If
        ' 係数1
        If HKT.係数1 <> 0 Then
           .TextMatrix(SetRow, 16) = Format(HKT.係数1, "0.00")
        Else
           .TextMatrix(SetRow, 16) = ""
        End If
        ' 係数2
        If HKT.係数2 <> 0 Then
           .TextMatrix(SetRow, 17) = Format(HKT.係数2, "0.00")
        Else
           .TextMatrix(SetRow, 17) = ""
        End If
        ' LT
        If HKT.LT <> 0 Then
           .TextMatrix(SetRow, 18) = Format(HKT.LT, "##")
        Else
           .TextMatrix(SetRow, 18) = ""
        End If
        ' 在庫管理
        Select Case HKT.在庫管理区分
               Case 0: .TextMatrix(SetRow, 19) = "なし"
               Case 1: .TextMatrix(SetRow, 19) = "あり"
               Case 2: .TextMatrix(SetRow, 19) = "係数"
               Case 3: .TextMatrix(SetRow, 19) = "発注点"
               Case 4: .TextMatrix(SetRow, 19) = "日数"
        End Select
        ' 最大在庫管理
        Select Case HKT.最大在庫管理区分
               Case 0: .TextMatrix(SetRow, 20) = "なし"
               Case 1: .TextMatrix(SetRow, 20) = "あり"
               Case 2: .TextMatrix(SetRow, 20) = "係数"
               Case 3: .TextMatrix(SetRow, 20) = "発注点"
               Case 4: .TextMatrix(SetRow, 20) = "日数"
        End Select
        ' 収容数
        If HKT.収容数 <> 0 Then
           .TextMatrix(SetRow, 21) = Format(HKT.収容数, "##,###")
        Else
           .TextMatrix(SetRow, 21) = ""
        End If
        ' 日報管理
        If HKT.日報出力 = 0 Then
           .TextMatrix(SetRow, 22) = "○"
        Else
           .TextMatrix(SetRow, 22) = "×"
        End If
    End With
End Sub
'*************************************************
'*      工順情報グリッドより入力域へセット        *
'*************************************************
Private Sub GridToText3(wKojyun As Byte)
    
    With HKT
        ' 工順情報
        gSL_Select = "SELECT HK.*,ISNULL(SY1.値名称,'') 機械名,ISNULL(SY2.値名称,'') 機械NO名,"
        gSL_Select = gSL_Select & " ISNULL(SY3.値名称,'') 工程名,ISNULL(TR.略称,'') 外注加工名 "
        gSL_Select = gSL_Select & " From 品番工順マスタ HK"
        gSL_Select = gSL_Select & " Left Outer join 名称マスタ SY1"
        gSL_Select = gSL_Select & " On  SY1.区分名称 = '機械' "
        gSL_Select = gSL_Select & " And SY1.値 = HK.機械"
        gSL_Select = gSL_Select & " And SY1.区分タイプ = ''"
        gSL_Select = gSL_Select & " Left Outer join 名称マスタ SY2"
        gSL_Select = gSL_Select & " On  SY2.区分名称 = '機械NO' "
        gSL_Select = gSL_Select & " And SY2.値 = HK.機械NO"
        gSL_Select = gSL_Select & " And SY2.区分タイプ = ''"
        gSL_Select = gSL_Select & " Left Outer join 名称マスタ SY3"
        gSL_Select = gSL_Select & " On  SY3.区分名称 = '工程管理' "
        gSL_Select = gSL_Select & " And SY3.値 = HK.工程"
        gSL_Select = gSL_Select & " And SY3.区分タイプ = ''"
        gSL_Select = gSL_Select & " Left Outer join 取引先マスタ TR"
        gSL_Select = gSL_Select & " On  TR.取引先CD = HK.外注加工先CD "
        gSL_Select = gSL_Select & " And TR.取引先区分 = 1"
            
        gSL_Select = gSL_Select & " WHERE HK.品番 = '" & RTrim(CXM0030.iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " And   HK.取引先 = '" & RTrim(CXM0030.iB2_Torcd) & "'"
        gSL_Select = gSL_Select & " And   HK.受入 = '" & RTrim(CXM0030.iB2_Ukeir) & "'"
        If CXM0030.B2Op_Hinmoku(0) Then
           gSL_Select = gSL_Select & " And   HK.品目 = 0"
        Else
           gSL_Select = gSL_Select & " And   HK.品目 = 1"
        End If
        gSL_Select = gSL_Select & " And HK.工順 = " & wKojyun
        
        If Not HKTRead(gSL_Select, 1) Then
           Exit Sub
        End If
        
        ' 工順
        iB3_Kojyun = .工順
        If SyoriKB_K = "MNT" Then
           iB3_Kojyun.Enabled = False
        Else
           iB3_Kojyun.Enabled = True
        End If
        ' 機械
        iB3_Kikai = .機械
        Call iB3_Kikai_LostFocus
        ' 内外
        B3Chk_NaiGai.Value = .外注加工
        Call B3Chk_NaiGai_Click
        If B3Chk_NaiGai.Value = 1 Then
           iB3_GTorcd = .外注加工先CD
           Call iB3_GTorcd_LostFocus
        End If
        ' 機械NO
        iB3_KikaiNo = .機械NO
        Call iB3_KikaiNo_LostFocus
        ' ﾛﾎﾞｯﾄ/単発
        B3Chk_Robot.Value = .ロボット区分
        Call B3Chk_Robot_Click
        gSL_Select = "Select 台数 From 取引先工程マスタ "
        gSL_Select = gSL_Select & " Where 取引先CD = '" & RTrim(CXM0030.iB2_Torcd) & "'"
        gSL_Select = gSL_Select & " And   機械 = '" & RTrim(iB3_Kikai) & "'"
        gSL_Select = gSL_Select & " And   機械NO = '" & RTrim(iB3_KikaiNo) & "'"
        If TKTRead(gSL_Select, 1) Then
           If TKT.台数 <> 0 Then
              B3lb_Daisu = Format(TKT.台数, "#0") & "台"
           Else
              B3lb_Daisu = ""
           End If
        Else
              B3lb_Daisu = ""
        End If
        
        ' 工程
        iB3_Kotei = .工程
        Call iB3_Kotei_LostFocus
        ' 基準時間1
        iB3_Kosunm1 = RTrim(.基準工数名1)
        iB3_Kosu1 = Format(.基準工数1, "#0.000")
        ' 基準時間2
        iB3_Kosunm2 = RTrim(.基準工数名2)
        iB3_Kosu2 = Format(.基準工数2, "#0.000")
        ' 工程数
        iB3_KoteiSu = Format(.工程数, "#0.0")
        ' 棚卸ﾌﾟﾚｼｰﾄ
        B3Chk_Pure.Value = .棚卸プレシート発行
        Call B3Chk_Pure_Click
        ' 置場
        iB3_Okiba1 = Mid(.置場, 1, 2)
        iB3_Okiba2 = Mid(.置場, 3, 2)
        iB3_Okiba3 = Mid(.置場, 5, 2)
        iB3_Okiba4 = Mid(.置場, 7, 2)
        ' 出荷管理
        B3Chk_Syuka.Value = .出荷管理
        Call B3Chk_Syuka_Click
        ' ｾｯﾄ品番
        B3Chk_Set.Value = .セット品番
        Call B3Chk_Set_Click
        ' サイクル
        iB3_KCycle1 = Mid(.サイクル, 1, 1)
        iB3_KCycle2 = Mid(.サイクル, 2, 2)
        iB3_KCycle3 = Mid(.サイクル, 4, 2)
        ' 実工数
        iB3_JKosu = Format(.実工数, "#0.00")
        ' 係数1
        iB3_KKeisu1 = Format(.係数1, "0.00")
        ' 係数2
        iB3_KKeisu2 = Format(.係数2, "0.00")
        ' LT
        iB3_KLT = Format(.LT, "##")
        ' 在庫管理
        CisComboKZaiko.Code = .在庫管理区分
        CisComboKZaiko.TextGet
        Call B3Comb_KZaiko_Click
        ' 在庫係数
        iB3_KZaiko = CisFun.RSetFld(.在庫係数, 0, iB3_KZaiko.cFormat)
        ' 在庫日数
        iB3_KZaikoDay = CisFun.RSetFld(.在庫日数, 0, iB3_KZaikoDay.cFormat)
        ' 在庫発注点
        iB3_KZaikoL = CisFun.RSetFld(.在庫発注点, 0, iB3_KZaikoL.cFormat)
        ' 最大在庫管理
        CisComboKMZaiko.Code = .最大在庫管理区分
        CisComboKMZaiko.TextGet
        Call B3Comb_KMZaiko_Click
        ' 最大在庫係数
        iB3_KMZaiko = CisFun.RSetFld(.最大在庫係数, 0, iB3_KMZaiko.cFormat)
        ' 最大在庫日数
        iB3_KMZaikoDay = CisFun.RSetFld(.最大在庫日数, 0, iB3_KMZaikoDay.cFormat)
        ' 最大在庫発注点
        iB3_KMZaikoL = CisFun.RSetFld(.最大在庫発注点, 0, iB3_KMZaikoL.cFormat)
        ' 収容数
        iB3_KSyuyo = Format(.収容数, "##,###")
        ' 日報
        B3Chk_Nipo.Value = .日報出力
        Call B3Chk_Nipo_Click
        
        
        'セット品番
        Call HeadBodyClear("B4")
        
        gSL_Select = "SELECT * FROM 品番工順セット品番マスタ "
        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(CXM0030.iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(CXM0030.iB2_Torcd) & "'"
        gSL_Select = gSL_Select & " And   受入 = '" & RTrim(CXM0030.iB2_Ukeir) & "'"
        If CXM0030.B2Op_Hinmoku(0) Then
           gSL_Select = gSL_Select & " And   品目 = 0"
        Else
           gSL_Select = gSL_Select & " And   品目 = 1"
        End If
        gSL_Select = gSL_Select & " And   工順 = " & CisFun.Val2(iB3_Kojyun)
        
        If HSERead(gSL_Select) Then
           mIx = 1
           Do Until Not HSE_RDSTS
              vsGrid4.TextMatrix(mIx, 1) = RTrim(HSE.セット品番)
              vsGrid4.TextMatrix(mIx, 2) = Format(HSE.取数, "###")
              
              If mIx = 1 Then
                 iB3_SetHinbn = RTrim(HSE.セット品番)
                 iB3_TSu = Format(HSE.取数, "###")
                 Call VSGrid4_Click
              End If
              
              mIx = mIx + 1
              Call HSEReadNext
           Loop
        
           vsGrid4.Row = 1
           
        End If
        Call HSEClose
        
    End With

End Sub
'+---------------------------------------+
'+  グリッド (DblClick) 工順情報
'+---------------------------------------+
Private Sub vsGrid3_DblClick()
    If SyoriKB_K = "MNT" Or SyoriKB_K = "DEL" Then
       If K_Select_Data Then
          Call PB_KENT_Click
       End If
    End If
End Sub
'+==========================================================+
'+      グリッド処理(ｾｯﾄ品番情報)                              +
'+==========================================================+
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit4()
    With CisVsGrid4
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 0, 0, "No")
        Call .FixedSet(中中, 左中, あり, 28, 1, "品番")
        Call .FixedSet(中中, 左中, あり, 4, 2, "取数")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 3, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
        Call .EditSet(iB3_SetHinbn, 1)
        Call .EditSet(iB3_TSu, 2)
        .InitDisp
    End With
End Sub
'+----------------------+
'+  VSGrid4_Click
'+----------------------+
Private Sub VSGrid4_Click()
    Call CisVsGrid4.EditEnable
    Call CisVsGrid4.EditVisible
End Sub
'+---------------------------------------------------------+
'+          グリッド編集エリア
'+---------------------------------------------------------+
'+----------------------+
'+  iB3_SetHinbn
'+----------------------+
Private Sub iB3_SetHinbn_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CisVsGrid4.KeyEdit(iB3_SetHinbn, KeyCode, Shift)
End Sub
'+----------------------+
'+  iB3_TSu
'+----------------------+
Private Sub iB3_TSu_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CisVsGrid4.KeyEdit(iB3_TSu, KeyCode, Shift)
End Sub
'+-----------------------------+
'+  iB3_SetHinbn
'+-----------------------------+
Private Sub iB3_SetHinbn_GotFocus()
    PB_KLook.Tag = ActiveControl.Name
    PB_KLook.Visible = True
End Sub
Private Sub iB3_SetHinbn_LostFocus()
    If ActiveControl.Name <> "PB_KLook" Then
        PB_KLook.Tag = ""
        PB_KLook.Visible = False
    End If
End Sub
'======================================================== ***  2009/07/22 End

'+---------------------------------------+
'+      名称マスタ初期設定
'+---------------------------------------+
Private Sub MeisyoInit()
' 品番編集
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "品番編集"
        .区分桁数 = 1
        .値 = "1"
        .英数字1 = "552"
        .値名称 = "品番編集初期表示"
        .作成者 = gTanto
    End With
    If Not SYMInitCreate(True) Then
    End If
' 工程
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "工程管理"
        .区分桁数 = 3
        .値 = "999"
        .値名称 = "その他"
        .作成者 = gTanto
        .検索名 = "工 　程　 検　 索"
        .コード名 = "工程"
        .名称名 = "工　程　名"
    End With
    If Not SYMInitCreate(True) Then
    End If
' 担当
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "担当管理"
        .区分桁数 = 3
        .値 = "999"
        .値名称 = "その他"
        .作成者 = gTanto
        .検索名 = "担 　当　 検　 索"
        .コード名 = "担当"
        .名称名 = "担　当　名"
    End With
    If Not SYMInitCreate(True) Then
    End If
' 手配区分
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "手配区分"
        .区分桁数 = 1
        .値 = "0"
        .値名称 = "購入"
        .英数字1 = "NJ "
        .制御区分 = "X"
        .作成者 = gTanto
        .検索名 = "手 配 区 分 検 索"
        .コード名 = "手配区"
        .名称名 = "手 配 区 分 名"
        If Not SYMInitCreate(True) Then
            .値 = "1"
            .値名称 = "直送"
            .英数字1 = "NJS"
            Call SYMInitCreate
            .値 = "2"
            .値名称 = "社内"
            .英数字1 = "   "
            Call SYMInitCreate
            .値 = "3"
            .値名称 = "受給"
            .英数字1 = "NJS "
            Call SYMInitCreate
            '-----* 2008/06/09 Start
            .値 = "4"
            .値名称 = "支給"
            .英数字1 = "N"
            Call SYMInitCreate
            '-----* 2008/06/09 End
        End If
    End With
' 在庫区分
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "在庫区分"
        .区分桁数 = 1
        .値 = "0"
        .制御区分 = "X"
        .値名称 = "なし"
        .作成者 = gTanto
        .検索名 = "在 庫 区 分 検 索"
        .コード名 = "在庫区"
        .名称名 = "在 庫 区 分 名"
        If Not SYMInitCreate(True) Then
            .値 = "1"
            .値名称 = "あり"
            Call SYMInitCreate
            .値 = "2"
            .値名称 = "係数"
            Call SYMInitCreate
            .値 = "3"
            .値名称 = "発注点"
            Call SYMInitCreate
            .値 = "4"
            .値名称 = "日数"
            Call SYMInitCreate
        End If
    End With
' 作業
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "作業管理"
        .区分桁数 = 4
        .値 = "9999"
        .値名称 = "その他"
        .作成者 = gTanto
        .検索名 = "作　業　検　索"
        .コード名 = "作業"
        .名称名 = "作　　業　　名"
        If Not SYMInitCreate(True) Then
        End If
    End With
' 収容器
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "収容器管理"
        .区分桁数 = 10
        .値 = "XXXXXXXXXX"
        .値名称 = "その他"
        .作成者 = gTanto
        .検索名 = "収 容 器 検 索"
        .コード名 = "収容器"
        .名称名 = "収  容  器  名"
        If Not SYMInitCreate(True) Then
        End If
    End With
' 品番区分
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "品番区分"
        .区分桁数 = 1
        .値 = "1"
        .制御区分 = ""
        .値名称 = "製品"
        .作成者 = gTanto
        .検索名 = "品 番 区 分 検 索"
        .コード名 = "品番区"
        .名称名 = "品 番 区 分 名"
        If Not SYMInitCreate(True) Then
            .値 = "2"
            .値名称 = "部品"
            Call SYMInitCreate
        End If
    End With
End Sub
