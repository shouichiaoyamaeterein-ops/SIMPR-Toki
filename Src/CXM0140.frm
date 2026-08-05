VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Begin VB.Form CXM0140 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "取引先変換マスタ　メンテ"
   ClientHeight    =   11235
   ClientLeft      =   2505
   ClientTop       =   4560
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   6135
      Left            =   1110
      ScaleHeight     =   6075
      ScaleWidth      =   13080
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3330
      Width           =   13140
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   5925
         Index           =   1
         Left            =   75
         Top             =   75
         Width           =   12915
         _ExtentX        =   22781
         _ExtentY        =   10451
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
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   510
            Left            =   12165
            Top             =   60
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   900
            ForeColor       =   16711680
            Caption         =   "品"
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
            Index           =   4
            Left            =   7455
            Top             =   60
            Width           =   4710
            _ExtentX        =   8308
            _ExtentY        =   900
            ForeColor       =   16711680
            Caption         =   "代行ｺｰﾄﾞ/受入/ 代 行 名 称"
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
            cPositionX      =   150
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   510
            Index           =   3
            Left            =   5700
            Top             =   60
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   900
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            Begin Cis3D_v60.CIS3D CIS3D8 
               Height          =   315
               Left            =   585
               Top             =   105
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               ForeColor       =   16711680
               Caption         =   "工区/受入"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D7 
               Height          =   480
               Left            =   120
               Top             =   15
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   847
               ForeColor       =   16711680
               Caption         =   "直直"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   510
            Index           =   2
            Left            =   4410
            Top             =   60
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   900
            ForeColor       =   16711680
            Caption         =   "契約先ｺｰﾄﾞ"
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
            Left            =   420
            Top             =   60
            Width           =   3990
            _ExtentX        =   7038
            _ExtentY        =   900
            ForeColor       =   16711680
            Caption         =   "自 社 取 引 先"
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
            Width           =   345
            _ExtentX        =   609
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
            Height          =   5730
            Left            =   60
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   105
            Width           =   12750
            _Version        =   196608
            _ExtentX        =   22490
            _ExtentY        =   10107
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
            FormatString    =   "№|自社取引|取引先名             |   |契約先    |支給元②  |代行      | 品 |                     |||"
            Rows            =   13
            Cols            =   19
            BackColor       =   12648384
            BackColorBkg    =   12632256
            FocusRect       =   2
            HighLight       =   2
            RowHeightMin    =   430
            SelectionMode   =   1
            Begin Cis3D_v60.CIS3D Input_Row 
               Height          =   435
               Left            =   345
               Top             =   435
               Width           =   12330
               _ExtentX        =   21749
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
               Begin Cis3D_v60.CIS3D Back_DTorcd 
                  Height          =   420
                  Left            =   7020
                  Top             =   0
                  Width           =   4710
                  _ExtentX        =   8308
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
                  Begin CisText_V60.CisText iB1_DTorcd 
                     Height          =   360
                     Left            =   75
                     TabIndex        =   16
                     Top             =   45
                     Width           =   720
                     _ExtentX        =   1270
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
                     cFaZero         =   0
                  End
                  Begin CisText_V60.CisText iB1_DTorKj 
                     Height          =   360
                     Left            =   795
                     TabIndex        =   17
                     Top             =   45
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
                     cDataReplace    =   1
                     cFaZero         =   0
                  End
                  Begin CisText_V60.CisText iB1_DTorNm 
                     Height          =   360
                     Left            =   1635
                     TabIndex        =   22
                     Top             =   45
                     Width           =   3090
                     _ExtentX        =   5450
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
                     Text            =   "XXXXXXXXX1XXXXXXXXX2XXXX"
                     MaxLength       =   24
                     IMEMode         =   4
                     cDataReplace    =   1
                     cFaZero         =   0
                  End
                  Begin CisText_V60.CisText iB1_DUke 
                     Height          =   360
                     Left            =   1245
                     TabIndex        =   18
                     Top             =   45
                     Width           =   360
                     _ExtentX        =   635
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
               Begin Cis3D_v60.CIS3D Back_STorcd 
                  Height          =   420
                  Left            =   5265
                  Top             =   0
                  Width           =   1755
                  _ExtentX        =   3096
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
                  Begin VB.CheckBox B1chk_Tyoku 
                     Caption         =   "非対象"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   9.75
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Left            =   45
                     TabIndex        =   12
                     Top             =   30
                     Width           =   930
                  End
                  Begin CisText_V60.CisText iB1_STorUke2 
                     Height          =   360
                     Left            =   1350
                     TabIndex        =   15
                     Top             =   45
                     Width           =   360
                     _ExtentX        =   635
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
                  Begin CisText_V60.CisText iB1_STorcd2 
                     Height          =   360
                     Left            =   120
                     TabIndex        =   13
                     TabStop         =   0   'False
                     Top             =   45
                     Width           =   720
                     _ExtentX        =   1270
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
                     cFaZero         =   0
                  End
                  Begin CisText_V60.CisText iB1_STorKj2 
                     Height          =   360
                     Left            =   990
                     TabIndex        =   14
                     Top             =   45
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
                     cDataReplace    =   1
                     cFaZero         =   0
                  End
               End
               Begin Cis3D_v60.CIS3D Back_JTorcd 
                  Height          =   420
                  Left            =   -30
                  Top             =   0
                  Width           =   4005
                  _ExtentX        =   7064
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
                  Begin Cis3D_v60.CIS3D B1lb_JTornm 
                     Height          =   345
                     Left            =   1035
                     Top             =   30
                     Width           =   2505
                     _ExtentX        =   4419
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
                  Begin CisText_V60.CisText iB1_JTorcd 
                     Height          =   360
                     Left            =   75
                     TabIndex        =   8
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
                  Begin CisText_V60.CisText iB1_JUke 
                     Height          =   360
                     Left            =   3555
                     TabIndex        =   9
                     Top             =   30
                     Width           =   330
                     _ExtentX        =   582
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
                  Left            =   3975
                  Top             =   0
                  Width           =   1290
                  _ExtentX        =   2275
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
                  Begin CisText_V60.CisText iB1_KTorcd 
                     Height          =   360
                     Left            =   120
                     TabIndex        =   10
                     Top             =   45
                     Width           =   720
                     _ExtentX        =   1270
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
                     cFaZero         =   0
                  End
                  Begin CisText_V60.CisText iB1_KTorKj 
                     Height          =   360
                     Left            =   840
                     TabIndex        =   11
                     Top             =   45
                     Width           =   330
                     _ExtentX        =   582
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
               Begin Cis3D_v60.CIS3D Back_Hin 
                  Height          =   420
                  Left            =   11730
                  Top             =   0
                  Width           =   345
                  _ExtentX        =   609
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
                  Begin Cis3D_v60.CIS3D B1lb_HinCnt 
                     Height          =   375
                     Left            =   15
                     Top             =   15
                     Width           =   300
                     _ExtentX        =   529
                     _ExtentY        =   661
                     BackColor       =   16777152
                     Caption         =   "ZZ"
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
                  End
               End
            End
         End
         Begin CisText_V60.CisText B_TabDummy 
            Height          =   165
            Left            =   5325
            TabIndex        =   23
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
      Left            =   1110
      ScaleHeight     =   885
      ScaleWidth      =   8295
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1230
      Width           =   8355
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Left            =   60
         Top             =   60
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1376
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   90
            TabIndex        =   0
            Top             =   330
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
            cAutoEject      =   0   'False
         End
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1080
            Top             =   330
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
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   780
         Left            =   3720
         Top             =   60
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "代行 契約先"
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
         Begin CisText_V60.CisText iH1_DaikoK 
            Height          =   360
            Left            =   90
            TabIndex        =   1
            Top             =   330
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
            cAutoEject      =   0   'False
         End
         Begin Cis3D_v60.CIS3D H1lb_DaikoK 
            Height          =   360
            Left            =   1080
            Top             =   330
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
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   780
         Left            =   7380
         Top             =   60
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "未設定表示"
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
         cAlingnment     =   6
         Begin VB.CheckBox ChkH1_Disp 
            BackColor       =   &H00C0C0C0&
            Caption         =   "無"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   600
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
      Height          =   1320
      Left            =   9870
      ScaleHeight     =   1260
      ScaleWidth      =   4290
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   885
      Width           =   4350
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2580
         Top             =   630
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
         Left            =   3420
         Top             =   630
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
         Top             =   630
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
      Begin CisBtn_60.CisBtn PB_NEXT 
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
         Left            =   885
         Top             =   630
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
      Begin CisBtn_60.CisBtn PB_Hinbn 
         Height          =   585
         Left            =   1725
         Top             =   630
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
         BtnCaption      =   "品番"
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
      TabIndex        =   24
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
      Caption         =   "【 取引先変換マスタ　メンテ  】"
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
   Begin VB.PictureBox B1_Area2 
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
      Height          =   675
      Left            =   1110
      ScaleHeight     =   615
      ScaleWidth      =   6435
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2580
      Width           =   6495
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   510
         Left            =   3945
         Top             =   60
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "自社コード"
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
         cPositionX      =   50
         Begin CisText_V60.CisText iB1_HiJisya 
            Height          =   360
            Left            =   2385
            TabIndex        =   6
            Top             =   90
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
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
            cAutoEject      =   0   'False
         End
         Begin Cis3D_v60.CIS3D B1lb_HiJisya 
            Height          =   360
            Left            =   1305
            Top             =   75
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXX-XX"
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
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   510
         Left            =   60
         Top             =   60
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "支給元"
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
         cPositionX      =   60
         Begin CisText_V60.CisText iB1_STorcd 
            Height          =   360
            Left            =   840
            TabIndex        =   4
            Top             =   75
            Width           =   735
            _ExtentX        =   1296
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
            cAutoEject      =   0   'False
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   510
         Left            =   1725
         Top             =   60
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "直直支給元"
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
         cPositionX      =   60
         Begin CisText_V60.CisText iB1_TyokTorcd 
            Height          =   360
            Left            =   1395
            TabIndex        =   5
            Top             =   75
            Width           =   735
            _ExtentX        =   1296
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
            cAutoEject      =   0   'False
         End
      End
   End
End
Attribute VB_Name = "CXM0140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   取引先変換マスタ　メンテ（直直支給品番対応Ver）(代行対応）
'**       フォームID    :   CXM1140
'**       処理概要      :
'**
'**       作  成  日    :   2006/02/09  By CIS SIMPR-A
'**       変  更  日    :   2006/03/13  By CIS 被自社を表示に変更(契約先の被自社を表示)
'**                         2006/03/15  By CIS 直直支給対応、「直直支給元」追加・・・DB項目は「被自社」に登録
'**                                            直直チェック追加、チェック有の場合、支給元２は支給元の内容をセット、無の場合は空白
'**                         2006/09/05  By CIS 品番設定有の削除処理追加
'**                         2007/11/17  By CIS 代行対応（代行契約先追加）
'**                                             未設定分の表示
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim ActObj          As Object
    Dim ErrObj          As Object
    Dim ActRow          As Integer
    
    Dim RD_Count        As Long
    Dim Dsp_Flg         As Byte

    Dim mSTorcd       As String
    Dim mHiJisya      As String
    Dim mTyokTorcd    As String
    Dim mTorcd()      As String
    Dim mUkeir()      As String
    Dim mKTorcd()     As String
    Dim mKTorkj()     As String
    Dim mSTorcd2()    As String
    Dim mSTorkj2()    As String
    Dim mSTorUke2()   As String
    Dim mDTorcd()     As String
    Dim mDTorkj()     As String
    Dim mDUke()       As String
    Dim mDTorNm()     As String
    Dim mJTorcd()     As String
    
    Dim mJisyaCd      As String
    Dim mJisyaKj      As String
    Dim mJisya        As String
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

    Dsp_Flg = 0
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
'   >> 名称マスタ確認
    Call MeisyoInit
'
'--------------------------------------------- 2006/03/15 Insert
'--( 自社コード )
    mJisyaCd = Left(gCompany, 4)
'--( JAMA自社工場・JAMA機番 ) [システム情報]
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 8
    Call IRNRead("", 1)
    mJisyaKj = Mid(IRN.oyian, 1, 1)
    
    mJisya = mJisyaCd
    If Trim(mJisyaKj) <> "" Then
       mJisya = mJisya & "-" & RTrim(mJisyaKj)
    End If
    
'--------------------------------------------- 2006/03/15 Insert End

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
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click  '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyPageUp:
                            If ProcHB = "H1" Then
                                If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
                            Else
                                Exit Function
                            End If
           Case vbKeyPageDown:
           
                            If ProcHB = "H1" Then
                                If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
                            Else
                                Exit Function
                            End If
           Case vbKeyEscape:  If PB_CAN.Visible Then Call PB_Can_Click     '【取消】
                                
           Case vbKeyF11:      If PB_Hinbn.Visible Then Call PB_Hinbn_Click    '【品番設定】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_Ent.Visible Then                      '【入力】
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
Private Sub PB_Can_Click()
   Call CisFrm.MousePT(11)
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   
   If ProcHB = "H1" Then
      SyoriKB = ""
      Call SyoriName(SyoriKB)
      PB_CAN.Visible = False
      iH1_Torcd.SetFocus
      GoTo PB_CAN_Ed
   End If
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      iH1_Torcd.SetFocus
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
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = " SELECT *"
    gSL_Select = gSL_Select & " FROM 取引先変換マスタ "
    gSL_Select = gSL_Select & " WHERE 契約先 + 代行契約先 < '" & RTrim(iH1_Torcd) & RTrim(iH1_DaikoK) & "'"             ' 2007/12/17 MNT
    gSL_Select = gSL_Select & " ORDER BY 契約先 Desc"
    If Not THMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     先頭レコードです            "
            .MB_Title = "【後退】"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Back_Ed
    End If
    gSL_Select = " SELECT '1' 区分,TH.契約先,TH.代行契約先,TH.取引先,TH.契約取引先,TH.契約工場,TH.契約取引先CD,TH.受入,"
    gSL_Select = gSL_Select & " TH.支給元,TH.被自社,TH.支給元2,TH.支給元工場2,TH.支給元2受入,TH.代行,TH.代行工場,TH.代行受入,TH.代行工場,TH.代行名称"
    gSL_Select = gSL_Select & " ,TR.略称,COUNT(*) 品番件数 "
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " ,TR2.被自社 契約先被自社,TR2.被自社工場 契約先被工場"
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " FROM 取引先変換マスタ TH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR.取引先CD = TH.取引先"
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2"
    gSL_Select = gSL_Select & "   ON  TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR2.取引先CD = TH.契約先"
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " WHERE TH.契約先 = '" & RTrim(THM.契約先) & "'"
    gSL_Select = gSL_Select & "   AND TH.代行契約先 = '" & RTrim(THM.代行契約先) & "'"                  ' 2007/12/17 ADD
    gSL_Select = gSL_Select & "   AND TH.品番 <> '' "
    gSL_Select = gSL_Select & " GROUP BY "
    gSL_Select = gSL_Select & " TH.契約先,TH.代行契約先,TH.取引先,TH.契約取引先,TH.契約工場,TH.契約取引先CD,TH.受入,"
    gSL_Select = gSL_Select & " TH.支給元,TH.被自社,TH.支給元2,TH.支給元工場2,TH.支給元2受入,TH.代行,TH.代行工場,TH.代行受入,TH.代行工場,TR.略称,TH.代行名称"
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " ,TR2.被自社 ,TR2.被自社工場 "
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " UNION "
    gSL_Select = gSL_Select & " SELECT '0' 区分,TH.契約先,TH.代行契約先,TH.取引先,TH.契約取引先,TH.契約工場,TH.契約取引先CD,TH.受入,"
    gSL_Select = gSL_Select & " TH.支給元,TH.被自社,TH.支給元2,TH.支給元工場2,TH.支給元2受入,TH.代行,TH.代行工場,TH.代行受入,TH.代行工場,TH.代行名称"
    gSL_Select = gSL_Select & " ,TR.略称,0 品番件数 "
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " ,TR2.被自社 契約先被自社,TR2.被自社工場 契約先被工場"
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " FROM 取引先変換マスタ TH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR.取引先CD = TH.取引先"
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2"
    gSL_Select = gSL_Select & "   ON  TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR2.取引先CD = TH.契約先"
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " WHERE TH.契約先 = '" & RTrim(THM.契約先) & "'"
    gSL_Select = gSL_Select & "   AND TH.代行契約先 = '" & RTrim(THM.代行契約先) & "'"                  ' 2007/12/17 ADD
    gSL_Select = gSL_Select & "   AND TH.品番 = '' "
    gSL_Select = gSL_Select & " GROUP BY "
    gSL_Select = gSL_Select & " TH.契約先,TH.代行契約先,TH.取引先,TH.契約取引先,TH.契約工場,TH.契約取引先CD,TH.受入,"
    gSL_Select = gSL_Select & " TH.支給元,TH.被自社,TH.支給元2,TH.支給元工場2,TH.支給元2受入,TH.代行,TH.代行工場,TH.代行受入,TH.代行工場,TR.略称,TH.代行名称"
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " ,TR2.被自社 ,TR2.被自社工場 "
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " ORDER BY TH.取引先,TH.受入,区分"
    
    Call DBInput("Back")
PB_Back_Ed:
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Torcd.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = " SELECT *"
    gSL_Select = gSL_Select & " FROM 取引先変換マスタ "
    gSL_Select = gSL_Select & " WHERE 契約先 + 代行契約先 > '" & RTrim(iH1_Torcd) & RTrim(iH1_DaikoK) & "'"     ' 2007/12/17 MNT
    gSL_Select = gSL_Select & " ORDER BY 契約先 "
    If Not THMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     最終レコードです            "
            .MB_Title = "【前進】"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Next_Ed
    End If
    gSL_Select = " SELECT '1' 区分,TH.契約先,TH.代行契約先,TH.取引先,TH.契約取引先,TH.契約工場,TH.契約取引先CD,TH.受入,"
    gSL_Select = gSL_Select & " TH.支給元,TH.被自社,TH.支給元2,TH.支給元工場2,TH.支給元2受入,TH.代行,TH.代行工場,TH.代行受入,TH.代行工場,TH.代行名称"
    gSL_Select = gSL_Select & " ,TR.略称,COUNT(*) 品番件数 "
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " ,TR2.被自社 契約先被自社,TR2.被自社工場 契約先被工場"
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " FROM 取引先変換マスタ TH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR.取引先CD = TH.取引先"
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2"
    gSL_Select = gSL_Select & "   ON  TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR2.取引先CD = TH.契約先"
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " WHERE TH.契約先 = '" & RTrim(THM.契約先) & "'"
    gSL_Select = gSL_Select & "   AND TH.代行契約先 = '" & RTrim(THM.代行契約先) & "'"              ' 2007/12/17 ADD
    gSL_Select = gSL_Select & "   AND TH.品番 <> '' "
    gSL_Select = gSL_Select & " GROUP BY "
    gSL_Select = gSL_Select & " TH.契約先,TH.代行契約先,TH.取引先,TH.契約取引先,TH.契約工場,TH.契約取引先CD,TH.受入,"
    gSL_Select = gSL_Select & " TH.支給元,TH.被自社,TH.支給元2,TH.支給元工場2,TH.支給元2受入,TH.代行,TH.代行工場,TH.代行受入,TH.代行工場,TR.略称,TH.代行名称"
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " ,TR2.被自社 ,TR2.被自社工場 "
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " UNION "
    gSL_Select = gSL_Select & " SELECT '0' 区分,TH.契約先,TH.代行契約先,TH.取引先,TH.契約取引先,TH.契約工場,TH.契約取引先CD,TH.受入,"
    gSL_Select = gSL_Select & " TH.支給元,TH.被自社,TH.支給元2,TH.支給元工場2,TH.支給元2受入,TH.代行,TH.代行工場,TH.代行受入,TH.代行工場,TH.代行名称"
    gSL_Select = gSL_Select & " ,TR.略称,0 品番件数 "
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " ,TR2.被自社 契約先被自社,TR2.被自社工場 契約先被工場"
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " FROM 取引先変換マスタ TH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR.取引先CD = TH.取引先"
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2"
    gSL_Select = gSL_Select & "   ON  TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR2.取引先CD = TH.契約先"
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " WHERE TH.契約先 = '" & RTrim(THM.契約先) & "'"
    gSL_Select = gSL_Select & "   AND TH.代行契約先 = '" & RTrim(THM.代行契約先) & "'"                  ' 2007/12/17 ADD
    gSL_Select = gSL_Select & "   AND TH.品番 = '' "
    gSL_Select = gSL_Select & " GROUP BY "
    gSL_Select = gSL_Select & " TH.契約先,TH.代行契約先,TH.取引先,TH.契約取引先,TH.契約工場,TH.契約取引先CD,TH.受入,"
    gSL_Select = gSL_Select & " TH.支給元,TH.被自社,TH.支給元2,TH.支給元工場2,TH.支給元2受入,TH.代行,TH.代行工場,TH.代行受入,TH.代行工場,TR.略称,TH.代行名称"
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " ,TR2.被自社 ,TR2.被自社工場 "
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " ORDER BY TH.取引先,TH.受入,区分"
    
    Call DBInput("Next")
PB_Next_Ed:
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Torcd.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd(0): Exit Sub
    If PB_Look.Tag = "iH1_DaikoK" Then Call Look_Torcd(1): Exit Sub
    If PB_Look.Tag = "iB1_JTorcd" Then Call Look_TorcdK: Exit Sub
    If PB_Look.Tag = "iB1_JUke" Then Call Look_UTorcd: Exit Sub
End Sub
'+----------------------------+
'+       取引先検索            +
'+----------------------------+
Private Sub Look_Torcd(tTKb As Byte)

    RV_Left = 0:    RV_Top = 0
    If tTKb = 0 Then
        RV_TorKb = 5
    Else
        RV_TorKb = 8
    End If
    RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        If tTKb = 0 Then
            iH1_Torcd = RTrim(RV_TorcdK)
            Call iH1_Torcd_LostFocus
        Else
            iH1_DaikoK = RTrim(RV_TorcdK)
            Call iH1_DaikoK_LostFocus
        End If
    End If

End Sub
'+-----------------------------------------------+
'+       取引先検索(品番取引先ベース)            +
'+-----------------------------------------------+
Private Sub Look_TorcdK()

    RV_Left = 0:    RV_Top = 0
    RV_TorKb = 1
    RV_JTorcd = RTrim(iH1_Torcd)
    RV_Call = "C"
    
    CKK0029.Show vbModal
    Unload CKK0029
    Set CKK0029 = Nothing

    If RV_Rtn Then
       iB1_JTorcd = RTrim(RV_TorcdK)
       Call iB1_JTorcd_LostFocus
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
    RV_Torcd = iB1_JTorcd
    RV_Kojyo = ""

    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iB1_JTorcd = RV_TorcdK
        iB1_JUke = RV_Ukeir
        B1lb_JTornm = RV_TorRName
    End If
End Sub
'+-----------------------------+
'+   品番設定画面表示
'+-----------------------------+
Private Sub PB_Hinbn_Click()
    Dim wRow        As Integer
    Dim wTRow       As Integer
    
    RV_Left = 0:    RV_Top = 0
    
    RV_TorcdK = RTrim(iH1_Torcd)                            '契約先
    RV_DaikoK = RTrim(iH1_DaikoK)                           '代行契約先     2007/12/17ADD
    RV_Torcd = RTrim(iB1_JTorcd)                            '自社取引先
    RV_Ukeir = RTrim(iB1_JUke)                              '自社受入
    RV_STorcd = RTrim(iB1_STorcd)                           '支給元
    RV_KTorcd = RTrim(iB1_KTorcd)                           '契約先ｺｰﾄﾞ
    RV_KTorkj = RTrim(iB1_KTorKj)                           '     "    工場
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
    RV_TyokTorcd = RTrim(iB1_TyokTorcd)                       '直直支給元
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
    RV_STorcd2 = RTrim(iB1_STorcd2)                         '支給元2
    RV_STorkj2 = RTrim(iB1_STorKj2)                         '     "    工場
    RV_STorUke2 = RTrim(iB1_STorUke2)                       '     "    受入
    RV_DTorcd = RTrim(iB1_DTorcd)                           '代行ｺｰﾄﾞ
    RV_DTorkj = RTrim(iB1_DTorKj)                           '     "    工場
    RV_DUke = RTrim(iB1_DUke)                               '代行受入
    RV_DTorNm = RTrim(iB1_DTorNm)                           '代行名称
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'    RV_HiJisya = RTrim(iB1_HiJisya)                         '被自社
    RV_HiJisya = RTrim(B1lb_HiJisya)                         '被自社
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
    wRow = vsGrid.Row
    wTRow = vsGrid.TopRow
    Call TextToVsGrid(wRow)
    
'(退避データチェック)
    If Not SaveDateCheck Then
        iB1_JTorcd.SetFocus
        Exit Sub
    End If
    
    CXM0145.Show vbModal
    Unload CXM0145
    Set CXM0145 = Nothing

    Dsp_Flg = 1
        
    If RV_Rtn Then
       Input_Row.Visible = False
       Call HeadBodyClear("B")
       Call Head1Chk
       Input_Row.Visible = True
       Call VsGridToText(vsGrid.Row)
       
        If wRow < vsGrid.Rows - 1 Then
            vsGrid.TopRow = wTRow
            vsGrid.Row = wRow
        Else
            vsGrid.Row = 1
        End If
        Call VsGridToText(vsGrid.Row)
       DoEvents
       iB1_JTorcd.SetFocus

'        Input_Row.Visible = False
'        If RV_SelCnt <> 0 Or RV_CrtRep = 1 Then
'            vsGrid.TextMatrix(wRow, 9) = RV_SelCnt
'            Call VsGridToText(wRow)
'        Else
'            With vsGrid
'                For gInt = 1 To vsGrid.Cols - 1
'                    .TextMatrix(wRow, gInt) = ""
'                Next gInt
'            End With
'        End If
'        Input_Row.Visible = True
'        Call VsGridToText(vsGrid.Row)
'        DoEvents
'        iB1_JTorcd.SetFocus
    End If
    
    Dsp_Flg = 0
    
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
        If TorNmGet(iH1_Torcd, 5, 0) Then
            H1lb_Tornm = RTrim(TRM.略称)
        End If
    End If
End Sub
'+-----------------------------+
'+  iH1_DaikoK                              2007/12/17 ADD
'+-----------------------------+
Private Sub iH1_DaikoK_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_DaikoK_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_DaikoK = ""
    If RTrim(iH1_DaikoK) <> "" Then
        If TorNmGet(iH1_DaikoK, 8, 0) Then
            H1lb_DaikoK = RTrim(TRM.略称)
        End If
    End If
End Sub
'+-----------------------------+
'+  ChkH1_Disp                              2007/12/17 ADD
'+-----------------------------+
Private Sub ChkH1_Disp_Click()
    If ChkH1_Disp.Value = 1 Then
       ChkH1_Disp.Caption = "有"
    Else
       ChkH1_Disp.Caption = "無"
    End If
End Sub
'+-----------------------------+
'+  iB1_JTorcd
'+-----------------------------+
Private Sub iB1_JTorcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB1_JTorcd
'    Back_JTorcd.Tag = "S"
End Sub
Private Sub iB1_JTorcd_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_JTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 1) = iB1_JTorcd
    vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 2) = ""
    B1lb_JTornm = ""
    If RTrim(iB1_JTorcd) <> "" Then
        If TorNmGet(iB1_JTorcd, 8) Then
            B1lb_JTornm = RTrim(TRM.略称)
            vsGrid.TextMatrix(CisFun.Val2(Input_Row.Tag), 2) = RTrim(TRM.略称)
        End If
    End If
End Sub
'+-----------------------------+
'+  iB1_JUke
'+-----------------------------+
Private Sub iB1_JUke_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB1_JUke
End Sub
Private Sub iB1_JUke_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_JUke_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_KTorcd
'+-----------------------------+
Private Sub iB1_KTorcd_GotFocus()
    Set ActObj = iB1_KTorcd
End Sub
Private Sub iB1_KTorcd_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'+-----------------------------+
'+  iB1_KTorKj
'+-----------------------------+
Private Sub iB1_KTorKj_GotFocus()
    Set ActObj = iB1_KTorKj
'    Back_JTorcd.Tag = "E"
End Sub
Private Sub iB1_KTorKj_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'+-----------------------------+
'+  iB1_STorcd2
'+-----------------------------+
Private Sub iB1_STorcd2_GotFocus()
    Set ActObj = iB1_STorcd2
End Sub
Private Sub iB1_STorcd2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'+-----------------------------+
'+  iB1_STorKj2
'+-----------------------------+STorUke2
Private Sub iB1_STorKj2_GotFocus()
    Set ActObj = iB1_STorKj2
End Sub
Private Sub iB1_STorKj2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'+-----------------------------+
'+  iB1_STorUke2
'+-----------------------------+
Private Sub iB1_STorUke2_GotFocus()
    Set ActObj = iB1_STorUke2
End Sub
Private Sub iB1_STorUke2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'+-----------------------------+
'+  iB1_DTorcd
'+-----------------------------+
Private Sub iB1_DTorcd_GotFocus()
    Set ActObj = iB1_DTorcd
End Sub
Private Sub iB1_DTorcd_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'+-----------------------------+
'+  iB1_DTorKj
'+-----------------------------+
Private Sub iB1_DTorKj_GotFocus()
    Set ActObj = iB1_DTorKj
End Sub
Private Sub iB1_DTorKj_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'+-----------------------------+
'+  iB1_DUke
'+-----------------------------+
Private Sub iB1_DUke_GotFocus()
    Set ActObj = iB1_DUke
End Sub
Private Sub iB1_DUke_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'+-----------------------------+
'+  iB1_DTorNm
'+-----------------------------+
Private Sub iB1_DTorNm_GotFocus()
    Set ActObj = iB1_DTorNm
'    Back_JTorcd.Tag = "E"
End Sub
Private Sub iB1_DTorNm_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'+-----------------------------+
'+  B1chk_Tyoku
'+-----------------------------+
Private Sub B1chk_Tyoku_GotFocus()
    Set ActObj = B1chk_Tyoku
End Sub
Private Sub B1chk_Tyoku_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub B1chk_Tyoku_Click()
    If B1chk_Tyoku.Value = 1 Then
       B1chk_Tyoku.Caption = "対象"
       iB1_STorcd2 = iB1_STorcd
    Else
       B1chk_Tyoku.Caption = "非対象"
       iB1_STorcd2 = ""
    End If
End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
'    If Back_JTorcd.Tag = "S" Then
'        Call TextToVsGrid(Val(Input_Row.Tag))
'        Call Grid_RowChange(vbKeyUp)
'    End If
'    If Back_JTorcd.Tag = "E" Then
'        Call TextToVsGrid(Val(Input_Row.Tag))
'        Call Grid_RowChange(vbKeyDown)
'    End If
'    SendKeys "{Tab}"
End Sub
Private Sub B_TabDummy_LostFocus()
    On Error Resume Next
'    If iB1_STorcd.Enabled Then iB1_STorcd.SetFocus
    On Error GoTo 0
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim W_Color     As Long
   W_Color = &H80000008 '初期色ｾｯﾄ(黒)
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
          W_Color = &HFF&
     Case "LOK"
          SyoriNM = "照会"
          K_Sykbnm.ForeColor = &HFF00FF   '(ピンク)
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
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID
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
    B1_Area2.Enabled = B1Mode: B1_Area2.BackColor = B1Color
    PB_Look.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_DEL.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    PB_Hinbn.Visible = Not H1Mode
    
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
        If RTrim$(iB1_STorcd) = "" Then iB1_STorcd.SetFocus Else iB1_JTorcd.SetFocus
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
'契約先
    If Trim(iH1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    契約先を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
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
'代行契約先
    If Trim(iH1_DaikoK) <> "" Then
        If Not TorNmGet(iH1_DaikoK, 8) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    代行契約先　未登録！（取引先マスタ）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_DaikoK.SetFocus
           Exit Function
        End If
    End If

' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = " SELECT '1' 区分,TH.契約先,TH.代行契約先,TH.取引先,TH.契約取引先,TH.契約工場,TH.契約取引先CD,TH.受入,"
    gSL_Select = gSL_Select & " TH.支給元,TH.被自社,TH.支給元2,TH.支給元工場2,TH.支給元2受入,TH.代行,TH.代行工場,TH.代行受入,TH.代行工場,TH.代行名称"
    gSL_Select = gSL_Select & " ,TR.略称,COUNT(*) 品番件数"
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " ,TR2.被自社 契約先被自社,TR2.被自社工場 契約先被工場"
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " FROM 取引先変換マスタ TH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR.取引先CD = TH.取引先"
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2"
    gSL_Select = gSL_Select & "   ON  TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR2.取引先CD = TH.契約先"
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " WHERE TH.契約先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND TH.代行契約先 = '" & RTrim(iH1_DaikoK) & "'"              ' 2007/12/17 ADD
    gSL_Select = gSL_Select & "   AND TH.品番 <> '' "
    gSL_Select = gSL_Select & " GROUP BY "
    gSL_Select = gSL_Select & " TH.契約先,TH.代行契約先,TH.取引先,TH.契約取引先,TH.契約工場,TH.契約取引先CD,TH.受入,"
    gSL_Select = gSL_Select & " TH.支給元,TH.被自社,TH.支給元2,TH.支給元工場2,TH.支給元2受入,TH.代行,TH.代行工場,TH.代行受入,TH.代行工場,TR.略称,TH.代行名称"
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " ,TR2.被自社 ,TR2.被自社工場 "
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " UNION "
    gSL_Select = gSL_Select & " SELECT '0' 区分,TH.契約先,TH.代行契約先,TH.取引先,TH.契約取引先,TH.契約工場,TH.契約取引先CD,TH.受入,"
    gSL_Select = gSL_Select & " TH.支給元,TH.被自社,TH.支給元2,TH.支給元工場2,TH.支給元2受入,TH.代行,TH.代行工場,TH.代行受入,TH.代行工場,TH.代行名称"
    gSL_Select = gSL_Select & " ,TR.略称,0 品番件数 "
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " ,TR2.被自社 契約先被自社,TR2.被自社工場 契約先被工場"
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " FROM 取引先変換マスタ TH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR.取引先CD = TH.取引先"
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2"
    gSL_Select = gSL_Select & "   ON  TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR2.取引先CD = TH.契約先"
'------------------------------------------------------------------ 2006/03/13 Insert End
    gSL_Select = gSL_Select & " WHERE TH.契約先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND TH.代行契約先 = '" & RTrim(iH1_DaikoK) & "'"              ' 2007/12/17 ADD
    gSL_Select = gSL_Select & "   AND TH.品番 = '' "
    gSL_Select = gSL_Select & " GROUP BY "
    gSL_Select = gSL_Select & " TH.契約先,TH.代行契約先,TH.取引先,TH.契約取引先,TH.契約工場,TH.契約取引先CD,TH.受入,"
    gSL_Select = gSL_Select & " TH.支給元,TH.被自社,TH.支給元2,TH.支給元工場2,TH.支給元2受入,TH.代行,TH.代行工場,TH.代行受入,TH.代行工場,TR.略称,TH.代行名称"
'------------------------------------------------------------------ 2006/03/13 Insert
    gSL_Select = gSL_Select & " ,TR2.被自社 ,TR2.被自社工場 "
'------------------------------------------------------------------ 2006/03/13 Insert End
'------------------------------------------------------------------ 2007/12/17 Insert
'   未設定分を品番取引先マスタより表示
    If ChkH1_Disp.Value = 1 Then
        gSL_Select = gSL_Select & " Union"
        gSL_Select = gSL_Select & " SELECT '2' 区分,HT.契約先,HT.代行契約先,HT.取引先,'' 契約取引先,'' 契約工場,'' 契約取引先CD,HT.受入,"
        gSL_Select = gSL_Select & "  '' 支給元,'' 被自社,'' 支給元2,'' 支給元工場2,'' 支給元2受入,'' 代行,'' 代行工場,'' 代行受入,"
        gSL_Select = gSL_Select & "  '' 代行工場,'' 代行名称 ,TR.略称,0 品番件数  ,TR2.被自社 契約先被自社,TR2.被自社工場 契約先被工場"
        gSL_Select = gSL_Select & " FROM 品番取引先マスタ HT"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN (SELECT 契約先,代行契約先,取引先,受入 FROM 取引先変換マスタ"
        gSL_Select = gSL_Select & "            WHERE 契約先 = '" & RTrim(iH1_Torcd) & "' AND 代行契約先 = '" & RTrim(iH1_DaikoK) & "'"
        gSL_Select = gSL_Select & "            GROUP BY 契約先,代行契約先,取引先,受入)TH"
        gSL_Select = gSL_Select & "   ON  TH.契約先 = HT.契約先"
        gSL_Select = gSL_Select & "   AND TH.代行契約先 = HT.代行契約先"
        gSL_Select = gSL_Select & "   AND TH.取引先 = HT.取引先"
        gSL_Select = gSL_Select & "   AND TH.受入 = HT.受入"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
        gSL_Select = gSL_Select & "   ON  TR.取引先区分 = 1"
        gSL_Select = gSL_Select & "   AND TR.取引先CD = HT.取引先"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2"
        gSL_Select = gSL_Select & "   ON  TR2.取引先区分 = 1"
        gSL_Select = gSL_Select & "   AND TR2.取引先CD = HT.契約先"
        gSL_Select = gSL_Select & " Where HT.品目 = 1"
        gSL_Select = gSL_Select & "   AND HT.契約先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND HT.代行契約先 = '" & RTrim(iH1_DaikoK) & "'"
        gSL_Select = gSL_Select & "   AND HT.手配区分 = '3'"                                '受給
        gSL_Select = gSL_Select & "   AND TH.契約先 IS NULL"
        gSL_Select = gSL_Select & " GROUP BY  HT.契約先,HT.代行契約先,HT.取引先,HT.受入,TR.略称,TR2.被自社 ,TR2.被自社工場"
    End If
'------------------------------------------------------------------ 2007/12/17 Insert End
    
    gSL_Select = gSL_Select & " ORDER BY TH.取引先,TH.受入,区分"
    
    If DBInput("INV") Then
        If SyoriKB = "" Then
            SyoriKB = "MNT"
        End If
    Else
        If SyoriKB = "DEL" Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    該当データが存在しません！        "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
        SyoriKB = "ADD"
    End If
       
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False

'   >> 支給元
    If Trim(iB1_STorcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    支給元を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_STorcd.SetFocus
        Exit Function
    End If

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
'-*-*-*-*-*-*-*-*- 2006/09/05 Del Start -*-*-*-*-*-*-*-*-*-*-*-*-
'        For gCnt1 = 1 To .Rows - 1
'            If Not BodyInput(gCnt1) Then
'                For gCnt2 = gCnt1 + 1 To .Rows - 1
'                    If BodyInput(gCnt2) Then
'                        For gCnt3 = 1 To 14
'                            .TextMatrix(gCnt1, gCnt3) = .TextMatrix(gCnt2, gCnt3)
'                            .TextMatrix(gCnt2, gCnt3) = ""
'                        Next gCnt3
'                        Exit For
'                    End If
'                Next gCnt2
'            End If
'        Next gCnt1
'-*-*-*-*-*-*-*-*- 2006/09/05 Del End   -*-*-*-*-*-*-*-*-*-*-*-*-
'
        For gCnt1 = 1 To .Rows - 1
            If Not BodyInput(gCnt1) Then Exit For
        '<< 自社取引先入力必須 >>
            If RTrim(.TextMatrix(gCnt1, 1)) = "" Then
                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　自社取引先　未入力！】            "
                CisFun.MBOX
                Set ErrObj = iB1_JTorcd
                GoTo Body1Chk_Err
            End If
        '<< 取引先マスタ存在チェック >>
            If Not TorNmGet(.TextMatrix(gCnt1, 1), 8) Then
                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　取引先マスタ　未登録！】            "
                CisFun.MBOX
                Set ErrObj = iB1_JTorcd
                GoTo Body1Chk_Err
            End If
        '<< 取引先受入マスタ存在チェック >>
            TUM.取引先CD = RTrim$(.TextMatrix(gCnt1, 1))
            TUM.取引先区分 = 1
            TUM.受入 = RTrim$(.TextMatrix(gCnt1, 3))
            If Not TUMRead("", 1) Then
                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　取引先受入マスタ　未登録！】            "
                CisFun.MBOX
                Set ErrObj = iB1_JUke
                GoTo Body1Chk_Err
            End If
        '<< 契約取引先入力必須 >>
            If RTrim(.TextMatrix(gCnt1, 4)) = "" Then
                CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　契約取引先　未入力！】            "
                CisFun.MBOX
                Set ErrObj = iB1_KTorcd
                GoTo Body1Chk_Err
            End If
        '<< 重複チェック >>
            For gCnt2 = gCnt1 + 1 To .Rows - 1
                If RTrim(.TextMatrix(gCnt1, 1)) = "" Then Exit For
                If RTrim(.TextMatrix(gCnt1, 1)) = RTrim(.TextMatrix(gCnt2, 1)) And _
                   RTrim(.TextMatrix(gCnt1, 3)) = RTrim(.TextMatrix(gCnt2, 3)) And _
                   RTrim(.TextMatrix(gCnt1, 9)) = "" And RTrim(.TextMatrix(gCnt2, 9)) = "" Then
                    CisFun.MB_MSG(1) = "    　自社取引先が重複しています。            "
                    CisFun.MB_MSG(3) = "    " & gCnt1 & "行目と " & gCnt2 & "行目            "
                    CisFun.MBOX
                    Set ErrObj = iB1_JTorcd
                    gCnt1 = gCnt2
                    GoTo Body1Chk_Err
                End If
            
                If RTrim(.TextMatrix(gCnt1, 1)) = RTrim(.TextMatrix(gCnt2, 1)) And _
                   RTrim(.TextMatrix(gCnt1, 3)) = RTrim(.TextMatrix(gCnt2, 3)) And _
                   RTrim(.TextMatrix(gCnt1, 4)) = RTrim(.TextMatrix(gCnt2, 4)) And _
                   RTrim(.TextMatrix(gCnt1, 5)) = RTrim(.TextMatrix(gCnt2, 5)) And _
                   RTrim(.TextMatrix(gCnt1, 6)) = RTrim(.TextMatrix(gCnt2, 6)) And _
                   RTrim(.TextMatrix(gCnt1, 7)) = RTrim(.TextMatrix(gCnt2, 7)) And _
                   RTrim(.TextMatrix(gCnt1, 8)) = RTrim(.TextMatrix(gCnt2, 8)) Then
                    CisFun.MB_MSG(1) = "    　内容が重複しています。            "
                    CisFun.MB_MSG(3) = "    " & gCnt1 & "行目と " & gCnt2 & "行目            "
                    CisFun.MBOX
                    Set ErrObj = iB1_JTorcd
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
           RTrim(.TextMatrix(RowNo, 2)) = "" And _
           RTrim(.TextMatrix(RowNo, 3)) = "" And _
           RTrim(.TextMatrix(RowNo, 4)) = "" And _
           RTrim(.TextMatrix(RowNo, 5)) = "" And _
           RTrim(.TextMatrix(RowNo, 6)) = "" And _
           RTrim(.TextMatrix(RowNo, 7)) = "" And _
           RTrim(.TextMatrix(RowNo, 8)) = "" Then Exit Function
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
        
    If Dsp_Flg = 0 Then Call HeadBodyClear("B")
    
    RD_Count = 0
    
    If Not THMRead(gSL_Select, 0, 1) Then
        Call THMClose(1)
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

    With THM
        iH1_Torcd = RTrim(.契約先)
        Call iH1_Torcd_LostFocus
        iH1_DaikoK = RTrim(.代行契約先)                             ' 2007/12/17 ADD
        Call iH1_DaikoK_LostFocus                                   ' 2007/12/17 ADD
        
        iB1_STorcd = RTrim(.支給元)
'        iB1_HiJisya = RTrim(.被自社)               2006/03/15
    
'------------------------------------------------------------------ 2006/03/13 Insert
        gStr = RTrim(.契約先被自社)
        If Trim(.契約先被工場) <> "" Then
           gStr = gStr & "-" & RTrim(.契約先被工場)
        End If
        B1lb_HiJisya = gStr
'------------------------------------------------------------------ 2006/03/13 Insert End
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
        If Trim(B1lb_HiJisya) = "" Then
           B1lb_HiJisya = mJisya
        End If
        
        iB1_TyokTorcd = RTrim(.被自社)
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
        
        mSTorcd = RTrim(.支給元)
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'        mHiJisya = RTrim(.被自社)              2006/03/15
        mHiJisya = RTrim(B1lb_HiJisya)
        mTyokTorcd = RTrim(.被自社)
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
        
    End With
'Grid表示
    Call GridSet
    Call THMClose(1)
        
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
        .SQL = "DELETE FROM 取引先変換マスタ "
        .SQL = .SQL & " WHERE 契約先 = '" & RTrim(iH1_Torcd) & "'"
        .SQL = .SQL & "   AND 代行契約先 = '" & RTrim(iH1_DaikoK) & "'"             '2007/12/17 ADD
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
Private Function DBPut(Optional pKB As Byte = 0)
    DBPut = False
    Dim AddFlg  As Boolean
    Dim wYDJun As Byte

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    If pKB = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     更新処理を行います      "
            .MB_Title = "更新処理"
            .MB_MSG(3) = "     よろしいですか？      "
            .MB_Button = OK_CAN
            If Not .MBOX Then
                Call VsGridToText(Val(Input_Row.Tag))
                Call GridTextEnabled(True)
                iB1_JTorcd.SetFocus
                Exit Function
            End If
        End With
    End If
    On Error GoTo DBPut_Err
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
'    If SyoriKB <> "ADD" Then
        With CisDB
            On Error GoTo DBPut_Err
    ' ( 削除処理 )
            ' 取引先変換マスタ
            .SQL = "DELETE FROM 取引先変換マスタ "
            .SQL = .SQL & " WHERE 契約先 = '" & RTrim(iH1_Torcd) & "'"
            .SQL = .SQL & "   AND 代行契約先 = '" & RTrim(iH1_DaikoK) & "'"             '2007/12/17 ADD
            .SQL = .SQL & "   AND ISNULL(品番,'') = ''"
                        
            .DBExec
    
            On Error GoTo 0
        End With
'    End If
'   マスタ項目移送
    For gInt = 1 To vsGrid.Rows - 1
        If RTrim(vsGrid.TextMatrix(gInt, 1)) <> "" And RTrim(vsGrid.TextMatrix(gInt, 9)) = "" Then
            Call ItemsClearTHM
            With THM
                .契約先 = RTrim(iH1_Torcd)
                .代行契約先 = RTrim(iH1_DaikoK)                         '2007/12/17 ADD
                .取引先 = RTrim(vsGrid.TextMatrix(gInt, 1))
                .契約取引先 = RTrim(vsGrid.TextMatrix(gInt, 10))
                .契約工場 = RTrim(vsGrid.TextMatrix(gInt, 11))
                .契約取引先CD = RTrim(vsGrid.TextMatrix(gInt, 10)) & RTrim(vsGrid.TextMatrix(gInt, 11))
                .受入 = RTrim(vsGrid.TextMatrix(gInt, 3))
                .支給元 = RTrim(iB1_STorcd)
                .支給元2 = RTrim(vsGrid.TextMatrix(gInt, 12))
                .支給元工場2 = RTrim(vsGrid.TextMatrix(gInt, 13))
                .支給元2受入 = RTrim(vsGrid.TextMatrix(gInt, 17))
                .代行 = RTrim(vsGrid.TextMatrix(gInt, 14))
                .代行工場 = RTrim(vsGrid.TextMatrix(gInt, 15))
                .代行受入 = RTrim(vsGrid.TextMatrix(gInt, 7))
                .代行名称 = RTrim(vsGrid.TextMatrix(gInt, 8))
                .作成者 = gTanto
                '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'                .被自社 = RTrim(iB1_HiJisya)
                .被自社 = RTrim(iB1_TyokTorcd)
                '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
            End With
        ' ( 追加処理 )
            Call THMInsert
        Else
            If RTrim(vsGrid.TextMatrix(gInt, 1)) <> "" And RTrim(vsGrid.TextMatrix(gInt, 9)) <> "" Then
               If RTrim(vsGrid.TextMatrix(gInt, 1)) = mTorcd(gInt) Then                                     ' 2006/09/05
                   With HTM
                        gSL_Select = "UPDATE 取引先変換マスタ SET "
                        gSL_Select = gSL_Select & " 取引先 = '" & RTrim(vsGrid.TextMatrix(gInt, 1)) & "',"
                        gSL_Select = gSL_Select & " 受入 = '" & RTrim(vsGrid.TextMatrix(gInt, 3)) & "',"
                        gSL_Select = gSL_Select & " 契約取引先 = '" & RTrim(vsGrid.TextMatrix(gInt, 10)) & "',"
                        gSL_Select = gSL_Select & " 契約工場 = '" & RTrim(vsGrid.TextMatrix(gInt, 11)) & "',"
                        gSL_Select = gSL_Select & " 契約取引先CD = '" & RTrim(vsGrid.TextMatrix(gInt, 10)) & RTrim(vsGrid.TextMatrix(gInt, 11)) & "',"
                        gSL_Select = gSL_Select & " 支給元 = '" & RTrim(iB1_STorcd) & "',"
                        gSL_Select = gSL_Select & " 支給元2 = '" & RTrim(vsGrid.TextMatrix(gInt, 12)) & "',"
                        gSL_Select = gSL_Select & " 支給元工場2 = '" & RTrim(vsGrid.TextMatrix(gInt, 13)) & "',"
                        gSL_Select = gSL_Select & " 支給元2受入 = '" & RTrim(vsGrid.TextMatrix(gInt, 17)) & "',"
                        gSL_Select = gSL_Select & " 代行 = '" & RTrim(vsGrid.TextMatrix(gInt, 14)) & "',"
                        gSL_Select = gSL_Select & " 代行工場 = '" & RTrim(vsGrid.TextMatrix(gInt, 15)) & "',"
                        gSL_Select = gSL_Select & " 代行受入 = '" & RTrim(vsGrid.TextMatrix(gInt, 7)) & "',"
                        gSL_Select = gSL_Select & " 代行名称 = '" & RTrim(vsGrid.TextMatrix(gInt, 8)) & "',"
                        '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
    '                    gSL_Select = gSL_Select & " 被自社 = '" & RTrim(iB1_HiJisya) & "',"
                        gSL_Select = gSL_Select & " 被自社 = '" & RTrim(iB1_TyokTorcd) & "',"
                        '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
                        gSL_Select = gSL_Select & " 作成者 = '" & gTanto & "',"
                        gSL_Select = gSL_Select & " 作成日 = getdate(), "
                        gSL_Select = gSL_Select & " 作成端末 = SubString(Host_Name(),1,20)"
                        gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(iH1_Torcd) & "'"
                        gSL_Select = gSL_Select & "   AND 代行契約先 = '" & RTrim(iH1_DaikoK) & "'"             '2007/12/17 ADD
                        gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(mTorcd(gInt)) & "'"
                        gSL_Select = gSL_Select & "   AND 受入   = '" & RTrim(mUkeir(gInt)) & "'"
                        gSL_Select = gSL_Select & "   AND 契約取引先 = '" & RTrim(mKTorcd(gInt)) & "'"
                        gSL_Select = gSL_Select & "   AND 契約工場 = '" & RTrim(mKTorkj(gInt)) & "'"
                        gSL_Select = gSL_Select & "   AND 支給元 = '" & RTrim(mSTorcd) & "'"
                        gSL_Select = gSL_Select & "   AND 支給元2 = '" & RTrim(mSTorcd2(gInt)) & "'"
                        gSL_Select = gSL_Select & "   AND 支給元工場2 = '" & RTrim(mSTorkj2(gInt)) & "'"
                        gSL_Select = gSL_Select & "   AND 支給元2受入 = '" & RTrim(mSTorUke2(gInt)) & "'"
                        gSL_Select = gSL_Select & "   AND 代行 = '" & RTrim(mDTorcd(gInt)) & "'"
                        gSL_Select = gSL_Select & "   AND 代行工場 = '" & RTrim(mDTorkj(gInt)) & "'"
                        gSL_Select = gSL_Select & "   AND 代行受入 = '" & RTrim(mDUke(gInt)) & "'"
                        gSL_Select = gSL_Select & "   AND 代行名称 = '" & RTrim(mDTorNm(gInt)) & "'"
                        '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
    '                    gSL_Select = gSL_Select & "   AND 被自社 = '" & RTrim(mHiJisya) & "'"
                        gSL_Select = gSL_Select & "   AND 被自社 = '" & RTrim(mTyokTorcd) & "'"
                        '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
                        gSL_Select = gSL_Select & "   AND 品番 <> '' "
                   End With
                
                   With CisDB
                       .SQL = gSL_Select
                       .DBExec
                   End With
               Else
            '-*-*-*-*-*-*-*-*-*-* データ更新処理 Add Start 2006/09/05 -*-*-*-*-*-*-*-*-*-*-*-*-*
                   gSL_Select = "SELECT TH.*, "
                   gSL_Select = gSL_Select & " CASE WHEN ISNULL(HT.品番,'') <> '' THEN 1 ELSE 0 END 品番件数"
                   gSL_Select = gSL_Select & " FROM 取引先変換マスタ TH"
                   gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT "
                   gSL_Select = gSL_Select & " ON  HT.品番 = TH.品番 "
                   gSL_Select = gSL_Select & " AND HT.取引先 = '" & RTrim(vsGrid.TextMatrix(gInt, 1)) & "'"
                   gSL_Select = gSL_Select & " AND HT.受入 = '" & RTrim(vsGrid.TextMatrix(gInt, 3)) & "'"
                   gSL_Select = gSL_Select & " AND HT.品目 = 1 "
                   gSL_Select = gSL_Select & " WHERE TH.契約先 = '" & RTrim(iH1_Torcd) & "'"
                   gSL_Select = gSL_Select & " AND   TH.代行契約先 = '" & RTrim(iH1_DaikoK) & "'"                   '2007/12/17 ADD
                   gSL_Select = gSL_Select & " AND   TH.取引先 = '" & RTrim(mTorcd(gInt)) & "'"
                   gSL_Select = gSL_Select & " AND   TH.受入   = '" & RTrim(mUkeir(gInt)) & "'"
                   gSL_Select = gSL_Select & " AND   ISNULL(TH.品番,'') <> '' "
                   If THMRead(gSL_Select) Then
                      Do Until THM_RDSTS = False
                         If THM.品番件数 = 0 Then      '品番取引先Ｍが未登録の場合削除
                            gSL_Select = "DELETE FROM 取引先変換マスタ "
                            gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(iH1_Torcd) & "'"
                            gSL_Select = gSL_Select & "   AND 代行契約先 = '" & RTrim(iH1_DaikoK) & "'"             '2007/12/17 ADD
                            gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(mTorcd(gInt)) & "'"
                            gSL_Select = gSL_Select & "   AND 受入   = '" & RTrim(mUkeir(gInt)) & "'"
                            gSL_Select = gSL_Select & "   AND 品番   = '" & RTrim(THM.品番) & "'"
              
                            With CisDB
                                .SQL = gSL_Select
                                .DBExec
                            End With
                         
                         End If
                         Call THMReadNext
                      Loop
                      Call THMClose
                                              
                      gSL_Select = "SELECT * FROM 取引先変換マスタ "
                      gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(iH1_Torcd) & "'"
                      gSL_Select = gSL_Select & " AND   代行契約先 = '" & RTrim(iH1_DaikoK) & "'"                   '2007/12/17 ADD
                      gSL_Select = gSL_Select & " AND   取引先 = '" & RTrim(mTorcd(gInt)) & "'"
                      gSL_Select = gSL_Select & " AND   受入   = '" & RTrim(mUkeir(gInt)) & "'"
                      If THMRead(gSL_Select, 1) Then
                         gSL_Select = "UPDATE 取引先変換マスタ SET "
                         gSL_Select = gSL_Select & " 取引先 = '" & RTrim(vsGrid.TextMatrix(gInt, 1)) & "',"
                         gSL_Select = gSL_Select & " 受入 = '" & RTrim(vsGrid.TextMatrix(gInt, 3)) & "',"
                         gSL_Select = gSL_Select & " 契約取引先 = '" & RTrim(vsGrid.TextMatrix(gInt, 10)) & "',"
                         gSL_Select = gSL_Select & " 契約工場 = '" & RTrim(vsGrid.TextMatrix(gInt, 11)) & "',"
                         gSL_Select = gSL_Select & " 契約取引先CD = '" & RTrim(vsGrid.TextMatrix(gInt, 10)) & RTrim(vsGrid.TextMatrix(gInt, 11)) & "',"
                         gSL_Select = gSL_Select & " 支給元 = '" & RTrim(iB1_STorcd) & "',"
                         gSL_Select = gSL_Select & " 支給元2 = '" & RTrim(vsGrid.TextMatrix(gInt, 12)) & "',"
                         gSL_Select = gSL_Select & " 支給元工場2 = '" & RTrim(vsGrid.TextMatrix(gInt, 13)) & "',"
                         gSL_Select = gSL_Select & " 支給元2受入 = '" & RTrim(vsGrid.TextMatrix(gInt, 17)) & "',"
                         gSL_Select = gSL_Select & " 代行 = '" & RTrim(vsGrid.TextMatrix(gInt, 14)) & "',"
                         gSL_Select = gSL_Select & " 代行工場 = '" & RTrim(vsGrid.TextMatrix(gInt, 15)) & "',"
                         gSL_Select = gSL_Select & " 代行受入 = '" & RTrim(vsGrid.TextMatrix(gInt, 7)) & "',"
                         gSL_Select = gSL_Select & " 代行名称 = '" & RTrim(vsGrid.TextMatrix(gInt, 8)) & "',"
                         gSL_Select = gSL_Select & " 被自社 = '" & RTrim(iB1_TyokTorcd) & "',"
                         gSL_Select = gSL_Select & " 作成者 = '" & gTanto & "',"
                         gSL_Select = gSL_Select & " 作成日 = getdate(), "
                         gSL_Select = gSL_Select & " 作成端末 = SubString(Host_Name(),1,20)"
                         gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(iH1_Torcd) & "'"
                         gSL_Select = gSL_Select & "   AND 代行契約先 = '" & RTrim(iH1_DaikoK) & "'"                '2007/12/17 ADD
                         gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(mTorcd(gInt)) & "'"
                         gSL_Select = gSL_Select & "   AND 受入   = '" & RTrim(mUkeir(gInt)) & "'"
                      Else
                          Call ItemsClearTHM
                          With THM
                              .契約先 = RTrim(iH1_Torcd)
                              .代行契約先 = RTrim(iH1_DaikoK)                                                       '2007/12/17 ADD
                              .取引先 = RTrim(vsGrid.TextMatrix(gInt, 1))
                              .契約取引先 = RTrim(vsGrid.TextMatrix(gInt, 10))
                              .契約工場 = RTrim(vsGrid.TextMatrix(gInt, 11))
                              .契約取引先CD = RTrim(vsGrid.TextMatrix(gInt, 10)) & RTrim(vsGrid.TextMatrix(gInt, 11))
                              .受入 = RTrim(vsGrid.TextMatrix(gInt, 3))
                              .支給元 = RTrim(iB1_STorcd)
                              .支給元2 = RTrim(vsGrid.TextMatrix(gInt, 12))
                              .支給元工場2 = RTrim(vsGrid.TextMatrix(gInt, 13))
                              .支給元2受入 = RTrim(vsGrid.TextMatrix(gInt, 17))
                              .代行 = RTrim(vsGrid.TextMatrix(gInt, 14))
                              .代行工場 = RTrim(vsGrid.TextMatrix(gInt, 15))
                              .代行受入 = RTrim(vsGrid.TextMatrix(gInt, 7))
                              .代行名称 = RTrim(vsGrid.TextMatrix(gInt, 8))
                              .作成者 = gTanto
                              .被自社 = RTrim(iB1_TyokTorcd)
                         End With
                         Call THMInsert
                      End If
                   Else
                      Call THMClose
                      
                      gSL_Select = "UPDATE 取引先変換マスタ SET "
                      gSL_Select = gSL_Select & " 取引先 = '" & RTrim(vsGrid.TextMatrix(gInt, 1)) & "',"
                      gSL_Select = gSL_Select & " 受入 = '" & RTrim(vsGrid.TextMatrix(gInt, 3)) & "',"
                      gSL_Select = gSL_Select & " 契約取引先 = '" & RTrim(vsGrid.TextMatrix(gInt, 10)) & "',"
                      gSL_Select = gSL_Select & " 契約工場 = '" & RTrim(vsGrid.TextMatrix(gInt, 11)) & "',"
                      gSL_Select = gSL_Select & " 契約取引先CD = '" & RTrim(vsGrid.TextMatrix(gInt, 10)) & RTrim(vsGrid.TextMatrix(gInt, 11)) & "',"
                      gSL_Select = gSL_Select & " 支給元 = '" & RTrim(iB1_STorcd) & "',"
                      gSL_Select = gSL_Select & " 支給元2 = '" & RTrim(vsGrid.TextMatrix(gInt, 12)) & "',"
                      gSL_Select = gSL_Select & " 支給元工場2 = '" & RTrim(vsGrid.TextMatrix(gInt, 13)) & "',"
                      gSL_Select = gSL_Select & " 支給元2受入 = '" & RTrim(vsGrid.TextMatrix(gInt, 17)) & "',"
                      gSL_Select = gSL_Select & " 代行 = '" & RTrim(vsGrid.TextMatrix(gInt, 14)) & "',"
                      gSL_Select = gSL_Select & " 代行工場 = '" & RTrim(vsGrid.TextMatrix(gInt, 15)) & "',"
                      gSL_Select = gSL_Select & " 代行受入 = '" & RTrim(vsGrid.TextMatrix(gInt, 7)) & "',"
                      gSL_Select = gSL_Select & " 代行名称 = '" & RTrim(vsGrid.TextMatrix(gInt, 8)) & "',"
                      gSL_Select = gSL_Select & " 被自社 = '" & RTrim(iB1_TyokTorcd) & "',"
                      gSL_Select = gSL_Select & " 作成者 = '" & gTanto & "',"
                      gSL_Select = gSL_Select & " 作成日 = getdate(), "
                      gSL_Select = gSL_Select & " 作成端末 = SubString(Host_Name(),1,20)"
                      gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(iH1_Torcd) & "'"
                      gSL_Select = gSL_Select & "   AND 代行契約先 = '" & RTrim(iH1_DaikoK) & "'"               '2007/12/17 ADD
                      gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(mTorcd(gInt)) & "'"
                      gSL_Select = gSL_Select & "   AND 受入   = '" & RTrim(mUkeir(gInt)) & "'"
                      gSL_Select = gSL_Select & "   AND 契約取引先 = '" & RTrim(mKTorcd(gInt)) & "'"
                      gSL_Select = gSL_Select & "   AND 契約工場 = '" & RTrim(mKTorkj(gInt)) & "'"
                      gSL_Select = gSL_Select & "   AND 支給元 = '" & RTrim(mSTorcd) & "'"
                      gSL_Select = gSL_Select & "   AND 支給元2 = '" & RTrim(mSTorcd2(gInt)) & "'"
                      gSL_Select = gSL_Select & "   AND 支給元工場2 = '" & RTrim(mSTorkj2(gInt)) & "'"
                      gSL_Select = gSL_Select & "   AND 支給元2受入 = '" & RTrim(mSTorUke2(gInt)) & "'"
                      gSL_Select = gSL_Select & "   AND 代行 = '" & RTrim(mDTorcd(gInt)) & "'"
                      gSL_Select = gSL_Select & "   AND 代行工場 = '" & RTrim(mDTorkj(gInt)) & "'"
                      gSL_Select = gSL_Select & "   AND 代行受入 = '" & RTrim(mDUke(gInt)) & "'"
                      gSL_Select = gSL_Select & "   AND 代行名称 = '" & RTrim(mDTorNm(gInt)) & "'"
                      gSL_Select = gSL_Select & "   AND 被自社 = '" & RTrim(mTyokTorcd) & "'"
              
                      With CisDB
                          .SQL = gSL_Select
                          .DBExec
                      End With
                   End If
               End If
            '-*-*-*-*-*-*-*-*-*-* データ更新処理 Add End   2006/09/05 -*-*-*-*-*-*-*-*-*-*-*-*-*
            
            '-*-*-*-*-*-*-*-*-*-* データ削除処理 Add Start 2006/09/05 -*-*-*-*-*-*-*-*-*-*-*-*-*
            Else
               If RTrim(vsGrid.TextMatrix(gInt, 9)) <> "" And RTrim(vsGrid.TextMatrix(gInt, 1)) = "" Then
                  gSL_Select = "DELETE FROM 取引先変換マスタ"
                  gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(iH1_Torcd) & "'"
                  gSL_Select = gSL_Select & "   AND 代行契約先 = '" & RTrim(iH1_DaikoK) & "'"                   '2007/12/17 ADD
                  gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(mTorcd(gInt)) & "'"
                  gSL_Select = gSL_Select & "   AND 受入   = '" & RTrim(mUkeir(gInt)) & "'"
                  gSL_Select = gSL_Select & "   AND 契約取引先 = '" & RTrim(mKTorcd(gInt)) & "'"
                  gSL_Select = gSL_Select & "   AND 契約工場 = '" & RTrim(mKTorkj(gInt)) & "'"
                  gSL_Select = gSL_Select & "   AND 支給元 = '" & RTrim(mSTorcd) & "'"
                  gSL_Select = gSL_Select & "   AND 支給元2 = '" & RTrim(mSTorcd2(gInt)) & "'"
                  gSL_Select = gSL_Select & "   AND 支給元工場2 = '" & RTrim(mSTorkj2(gInt)) & "'"
                  gSL_Select = gSL_Select & "   AND 支給元2受入 = '" & RTrim(mSTorUke2(gInt)) & "'"
                  gSL_Select = gSL_Select & "   AND 代行 = '" & RTrim(mDTorcd(gInt)) & "'"
                  gSL_Select = gSL_Select & "   AND 代行工場 = '" & RTrim(mDTorkj(gInt)) & "'"
                  gSL_Select = gSL_Select & "   AND 代行受入 = '" & RTrim(mDUke(gInt)) & "'"
                  gSL_Select = gSL_Select & "   AND 代行名称 = '" & RTrim(mDTorNm(gInt)) & "'"
                  gSL_Select = gSL_Select & "   AND 被自社 = '" & RTrim(mTyokTorcd) & "'"
                  gSL_Select = gSL_Select & "   AND 品番 <> '' "
                  
                  With CisDB
                      .SQL = gSL_Select
                      .DBExec
                  End With
               
               End If
            '-*-*-*-*-*-*-*-*-*-* データ削除処理 Add End   2006/09/05 -*-*-*-*-*-*-*-*-*-*-*-*-*
            End If
        End If
    Next gInt
    
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'*------------------------------------------------------------------------------*
'*      退避データと入力テキストとのチェック ( 異なっていれば更新をうながす )       *
'*------------------------------------------------------------------------------*
Private Function SaveDateCheck()
    SaveDateCheck = False
    
    Static swTopR           As Integer
    Static swRow            As Integer

    swTopR = vsGrid.TopRow
    swRow = vsGrid.Row
    
    If Not Body1Chk Then Exit Function
    Call DBPut(1)
    Input_Row.Visible = False
    Call HeadBodyClear("B")
    Call Head1Chk
    Input_Row.Visible = True
    vsGrid.TopRow = swTopR
    
    If vsGrid.Rows <= swRow Then swRow = vsGrid.Rows - 1
    vsGrid.Row = swRow
    Call VsGridToText(vsGrid.Row)
    DoEvents
    iB1_JTorcd.SetFocus
    

'    With vsGrid
'        If mTorcd(tRow) = RTrim$(.TextMatrix(tRow, 1)) And _
'            mUkeir(tRow) = RTrim$(.TextMatrix(tRow, 3)) And _
'            mKTorcd(tRow) = RTrim$(.TextMatrix(tRow, 10)) And _
'            mKTorkj(tRow) = RTrim$(.TextMatrix(tRow, 11)) And _
'            mSTorcd2(tRow) = RTrim$(.TextMatrix(tRow, 12)) And _
'            mSTorkj2(tRow) = RTrim$(.TextMatrix(tRow, 13)) And _
'            mSTorUke2(tRow) = RTrim$(.TextMatrix(tRow, 14)) And _
'            mDTorcd(tRow) = RTrim$(.TextMatrix(tRow, 15)) And _
'            mDTorkj(tRow) = RTrim$(.TextMatrix(tRow, 16)) And _
'            mDUke(tRow) = RTrim$(.TextMatrix(tRow, 7)) And _
'            mDTorNm(tRow) = RTrim$(.TextMatrix(tRow, 8)) Then
'                GoTo SaveDateCheck_ED
'        End If
'    End With
'
'    With CisFun
'        .MB_Lines = 5
'        .MB_MSG(1) = "      【" & StrConv(tRow, vbWide) & "行目の内容が変更されています。】        "
'        .MB_MSG(2) = "     品番選択の前に更新を行う必要があります。         "
'        .MB_MSG(4) = "              よろしいですか？      "
'        .MB_Title = "更新処理"
'        .MB_Button = OK_CAN
'        If Not .MBOX Then
'            iB1_JTorcd.SetFocus
'            Exit Function
'        End If
'    End With
    
SaveDateCheck_ED:
    SaveDateCheck = True
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
        Call .FixedSet(中中, 中中, あり, 2, 0, "№")
        Call .FixedSet(中中, 中中, あり, 8, 1, "自社取引")
'        Call .FixedSet(中中, 左中, あり, 21, 2, "自社取引先名")            '2006/03/15
        Call .FixedSet(中中, 左中, あり, 20, 2, "自社取引先名")             '2006/03/15
        Call .FixedSet(中中, 中中, あり, 3, 3, "受入")
        Call .FixedSet(中中, 中中, あり, 10, 4, "契約取引先")
'        Call .FixedSet(中中, 中中, あり, 13, 5, "支給元2")                 '2006/03/15
        Call .FixedSet(中中, 中中, あり, 14, 5, "支給元2")                  '2006/03/15
        Call .FixedSet(中中, 中中, あり, 9, 6, "代行")
        Call .FixedSet(中中, 中中, あり, 3, 7, "代行受入")
        Call .FixedSet(中中, 左中, あり, 25, 8, "代行名称")
        Call .FixedSet(中中, 中中, あり, 2, 9, "品")                '品番件数
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 中中, あり, 0, 10, "x契約取引先")
        Call .FixedSet(中中, 中中, あり, 0, 11, "x契約取引工場")
        Call .FixedSet(中中, 中中, あり, 0, 12, "x支給元2")
        Call .FixedSet(中中, 中中, あり, 0, 13, "x支給元工場2")
        Call .FixedSet(中中, 中中, あり, 0, 14, "x代行")
        Call .FixedSet(中中, 中中, あり, 0, 15, "x代行工場")
        Call .FixedSet(中中, 左中, なし, 0, 16, "Select")
        Call .FixedSet(中中, 中中, あり, 0, 17, "x支給元2受入")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9")
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
    
    ReDim mTorcd(Ttl_Kensu + 1)
    ReDim mUkeir(Ttl_Kensu + 1)
    ReDim mKTorcd(Ttl_Kensu + 1)
    ReDim mKTorkj(Ttl_Kensu + 1)
    ReDim mSTorcd2(Ttl_Kensu + 1)
    ReDim mSTorkj2(Ttl_Kensu + 1)
    ReDim mSTorUke2(Ttl_Kensu + 1)
    ReDim mDTorcd(Ttl_Kensu + 1)
    ReDim mDTorkj(Ttl_Kensu + 1)
    ReDim mDUke(Ttl_Kensu + 1)
    ReDim mDTorNm(Ttl_Kensu + 1)
    ReDim mJTorcd(Ttl_Kensu + 1)
    
    Do Until Not THM_RDSTS
        'Grid Row,Col
        ID = ID + 1
        With vsGrid
            .TextMatrix(ID, GridCol(0)) = ID
            .TextMatrix(ID, GridCol(1)) = THM.取引先 + Space(7 - CisFun.Len2(THM.取引先))
            .TextMatrix(ID, GridCol(2)) = RTrim$(THM.略称)
            .TextMatrix(ID, GridCol(3)) = THM.受入 + Space(2 - CisFun.Len2(THM.受入))
            
            gStr = " " & THM.契約取引先 & Space(5 - CisFun.Len2(THM.契約取引先))
            If RTrim$(THM.契約工場) = "" Then
                gStr = gStr & Space(4)
            Else
                gStr = gStr & "-" & THM.契約工場 + Space(2 - CisFun.Len2(THM.契約工場)) & Space(1)
            End If
            .TextMatrix(ID, GridCol(4)) = gStr
            
            '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'            If RTrim(THM.支給元2) <> "" Then gStr = Format(RTrim(THM.支給元2), "!@@@@@") Else gStr = Space(5)
            If RTrim(THM.支給元2) <> "" Then
               gStr = "対象  "
            Else
               gStr = Space(6)
            End If
            '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
            
            If RTrim(THM.支給元工場2) <> "" Then
            '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'                gStr = gStr & "-" & Format(RTrim(THM.支給元工場2), "!@@")
                gStr = gStr & "  " & Format(RTrim(THM.支給元工場2), "!@@")
            '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
            Else
                gStr = gStr & Space(3)
            End If
            If RTrim$(THM.支給元2受入) <> "" Then
            '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'                gStr = gStr & ":" & Format(RTrim(THM.支給元2受入), "!@@")
                gStr = gStr & "/" & Format(RTrim(THM.支給元2受入), "!@@")
            '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
            Else
                gStr = gStr & Space(3)
            End If
            .TextMatrix(ID, GridCol(5)) = gStr
            
            If RTrim(THM.代行工場) <> "" Then
                .TextMatrix(ID, GridCol(6)) = THM.代行 & Space(5 - CisFun.Len2(THM.代行)) _
                                      & "-" & THM.代行工場 + Space(2 - CisFun.Len2(THM.代行工場))
            Else
                .TextMatrix(ID, GridCol(6)) = THM.代行 & Space(5 - CisFun.Len2(THM.代行)) & Space(3)
            End If
            .TextMatrix(ID, GridCol(7)) = THM.代行受入 + Space(2 - CisFun.Len2(THM.代行受入))
            .TextMatrix(ID, GridCol(8)) = RTrim$(THM.代行名称)
            
            .TextMatrix(ID, GridCol(9)) = Format(THM.品番件数, "##")
            
            .TextMatrix(ID, GridCol(10)) = RTrim$(THM.契約取引先)
            .TextMatrix(ID, GridCol(11)) = RTrim$(THM.契約工場)
            .TextMatrix(ID, GridCol(12)) = RTrim$(THM.支給元2)
            .TextMatrix(ID, GridCol(13)) = RTrim$(THM.支給元工場2)
            .TextMatrix(ID, GridCol(17)) = RTrim$(THM.支給元2受入)
            .TextMatrix(ID, GridCol(14)) = RTrim$(THM.代行)
            .TextMatrix(ID, GridCol(15)) = RTrim$(THM.代行工場)
        End With
        '初期情報退避
        
        mTorcd(ID) = RTrim(THM.取引先)
        mUkeir(ID) = RTrim(THM.受入)
        mKTorcd(ID) = RTrim(THM.契約取引先)
        mKTorkj(ID) = RTrim(THM.契約工場)
        mSTorcd2(ID) = RTrim(THM.支給元2)
        mSTorkj2(ID) = RTrim(THM.支給元工場2)
        mSTorUke2(ID) = RTrim(THM.支給元2受入)
        mDTorcd(ID) = RTrim(THM.代行)
        mDTorkj(ID) = RTrim(THM.代行工場)
        mDUke(ID) = RTrim(THM.代行受入)
        mDTorNm(ID) = RTrim(THM.代行名称)
        
        Call THMReadNext(1)
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
        .TextMatrix(tRow, 1) = iB1_JTorcd + Space(7 - CisFun.Len2(iB1_JTorcd))
        .TextMatrix(tRow, 2) = ""
        If RTrim(iB1_JTorcd) <> "" Then
            If TorNmGet(iB1_JTorcd, 8) Then
                .TextMatrix(tRow, 2) = RTrim(TRM.略称)
            End If
        End If
        .TextMatrix(tRow, 3) = iB1_JUke + Space(2 - CisFun.Len2(iB1_JUke))
'       >> 支払先コード
        If RTrim(iB1_KTorKj) <> "" Then
            .TextMatrix(tRow, 4) = " " & iB1_KTorcd + Space(5 - CisFun.Len2(iB1_KTorcd)) _
                                 & "-" & iB1_KTorKj + Space(2 - CisFun.Len2(iB1_KTorKj)) & Space(1)
        Else
            .TextMatrix(tRow, 4) = " " & iB1_KTorcd + Space(5 - CisFun.Len2(iB1_KTorcd)) & Space(4)
        End If
'       >> 支給元②
'        If RTrim(iB1_STorKj2) <> "" Then
'            .TextMatrix(tRow, 5) = " " & iB1_STorcd2 + Space(5 - CisFun.Len2(iB1_STorcd2)) _
'                                 & "-" & iB1_STorKj2 + Space(2 - CisFun.Len2(iB1_STorKj2)) & Space(1)
'        Else
'            .TextMatrix(tRow, 5) = " " & iB1_STorcd2 + Space(5 - CisFun.Len2(iB1_STorcd2)) & Space(4)
'        End If
'        If RTrim(iB1_STorKj2) <> "" Then
'            .TextMatrix(tRow, 5) = " " & iB1_STorcd2 + iB1_STorKj2 + Space(7 - CisFun.Len2(iB1_STorcd2 + iB1_STorKj2)) _
'                                 & "-" & iB1_STorUke2 + Space(2 - CisFun.Len2(iB1_STorUke2)) & Space(1)
'        Else
'            .TextMatrix(tRow, 5) = " " & iB1_STorcd2 + Space(5 - CisFun.Len2(iB1_STorcd2)) & Space(4)
'        End If

        
        '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
''       >> 支給元②
'        If RTrim(iB1_STorcd2) <> "" Then gStr = Format(RTrim(iB1_STorcd2), "!@@@@@") Else gStr = Space(5)
'        If RTrim(iB1_STorKj2) <> "" Then
'            gStr = gStr & "-" & Format(RTrim(iB1_STorKj2), "!@@")
'        Else
'            gStr = gStr & Space(3)
'        End If
'        If RTrim$(iB1_STorUke2) <> "" Then
'            gStr = gStr & ":" & Format(RTrim(iB1_STorUke2), "!@@")
'        Else
'            gStr = gStr & Space(3)
'        End If
'        .TextMatrix(tRow, 5) = gStr

'       >> 支給元②
        If RTrim(iB1_STorcd2) <> "" Then gStr = "対象  " Else gStr = Space(6)
        If RTrim(iB1_STorKj2) <> "" Then
            gStr = gStr & "  " & Format(RTrim(iB1_STorKj2), "!@@")
        Else
            gStr = gStr & Space(3)
        End If
        If RTrim$(iB1_STorUke2) <> "" Then
            gStr = gStr & ":" & Format(RTrim(iB1_STorUke2), "!@@")
        Else
            gStr = gStr & Space(3)
        End If
        .TextMatrix(tRow, 5) = gStr
        '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End


'       >> 代行コード／受入／受入名称
        If RTrim(iB1_DTorKj) <> "" Then
            .TextMatrix(tRow, 6) = iB1_DTorcd + Space(5 - CisFun.Len2(iB1_DTorcd)) _
                           & "-" & iB1_DTorKj + Space(2 - CisFun.Len2(iB1_DTorKj))
        Else
            .TextMatrix(tRow, 6) = iB1_DTorcd + Space(5 - CisFun.Len2(iB1_DTorcd)) & Space(3)
        End If
        .TextMatrix(tRow, 7) = iB1_DUke + Space(2 - CisFun.Len2(iB1_DUke))
        .TextMatrix(tRow, 8) = RTrim$(iB1_DTorNm)
        
        .TextMatrix(tRow, 9) = Format(CisFun.Val2(B1lb_HinCnt), "##")

'       >>
        .TextMatrix(tRow, 10) = iB1_KTorcd
        .TextMatrix(tRow, 11) = iB1_KTorKj
        .TextMatrix(tRow, 12) = iB1_STorcd2
        .TextMatrix(tRow, 13) = iB1_STorKj2
        .TextMatrix(tRow, 14) = iB1_DTorcd
        .TextMatrix(tRow, 15) = iB1_DTorKj
        .TextMatrix(tRow, 17) = iB1_STorUke2
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With vsGrid
        Input_Row.Tag = tRow
        Input_Row.Move Input_Row.Left, vsGrid.CellTop, Input_Row.Width, vsGrid.CellHeight
        
        '自社取引先
        iB1_JTorcd = RTrim(.TextMatrix(tRow, 1))
        '自社取引先名
        B1lb_JTornm = RTrim(.TextMatrix(tRow, 2))
        '>> 受入
        iB1_JUke = RTrim$(.TextMatrix(tRow, 3))
        '>> 品番件数
        B1lb_HinCnt = Format(CisFun.Val2(.TextMatrix(tRow, 9)), "##")
        '>> 契約先取引先
        iB1_KTorcd = RTrim$(.TextMatrix(tRow, 10))
        iB1_KTorKj = RTrim$(.TextMatrix(tRow, 11))
        '>> 支給元
        iB1_STorcd2 = RTrim$(.TextMatrix(tRow, 12))
        '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
        If RTrim(iB1_STorcd2) <> "" Then
           B1chk_Tyoku.Value = 1
        Else
           B1chk_Tyoku.Value = 0
        End If
        '[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
        
        iB1_STorKj2 = RTrim$(.TextMatrix(tRow, 13))
        iB1_STorUke2 = RTrim$(.TextMatrix(tRow, 17))
        '>> 代行
        iB1_DTorcd = RTrim$(.TextMatrix(tRow, 14))
        iB1_DTorKj = RTrim$(.TextMatrix(tRow, 15))
        iB1_DUke = RTrim$(.TextMatrix(tRow, 7))
        iB1_DTorNm = RTrim$(.TextMatrix(tRow, 8))
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

'+---------------------------------------+
'+      名称マスタ初期設定
'+---------------------------------------+
Private Sub MeisyoInit()
'   >> 後工程符号管理
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "後工程符号管理"
        .区分桁数 = 5
        .値 = "1097"
        .値名称 = "アラコ"
        .英数字1 = "AR"
        .作成者 = gTanto
        If Not SYMInitCreate(True) Then
            .値 = "1613"
            .値名称 = "セントラル"
            .英数字1 = "CN"
            Call SYMInitCreate
            .値 = "1875"
            .値名称 = "ダイハツ"
            .英数字1 = "DH"
            Call SYMInitCreate
            .値 = "2474"
            .値名称 = "岐阜車"
            .英数字1 = "GI"
            Call SYMInitCreate
            .値 = "2625"
            .値名称 = "日野"
            .英数字1 = "HN"
            Call SYMInitCreate
            .値 = "3230"
            .値名称 = "関自"
            .英数字1 = "KN"
            Call SYMInitCreate
            .値 = "7371"
            .値名称 = "織機"
            .英数字1 = "KY"
            Call SYMInitCreate
            .値 = "7388"
            .値名称 = "トヨ車"
            .英数字1 = "TY"
            Call SYMInitCreate
            .値 = "8482"
            .値名称 = "トヨタ九州"
            .英数字1 = "TMK"
            Call SYMInitCreate
            .値 = "7364"
            .値名称 = "トヨタ紡織"
            .英数字1 = "TB"
            Call SYMInitCreate
        End If
    End With
End Sub

