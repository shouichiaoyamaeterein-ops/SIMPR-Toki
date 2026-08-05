VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CKH0150 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "かんばん増減・平準化一括メンテナンス"
   ClientHeight    =   11565
   ClientLeft      =   2130
   ClientTop       =   2805
   ClientWidth     =   15315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11565
   ScaleWidth      =   15315
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   4380
      Top             =   5025
      Visible         =   0   'False
      Width           =   6435
      _ExtentX        =   11351
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
         Width           =   6015
         _ExtentX        =   10610
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
      Height          =   9165
      Left            =   45
      ScaleHeight     =   9105
      ScaleWidth      =   15150
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1815
      Width           =   15210
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   9030
         Left            =   0
         Top             =   45
         Width           =   15150
         _ExtentX        =   26723
         _ExtentY        =   15928
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
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   330
            Index           =   25
            Left            =   14280
            Top             =   360
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   582
            ForeColor       =   16711680
            Caption         =   "設定"
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
            Height          =   330
            Index           =   24
            Left            =   13680
            Top             =   360
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   582
            ForeColor       =   16711680
            Caption         =   "区分"
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
            Height          =   300
            Index           =   21
            Left            =   13680
            Top             =   60
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "平準管理"
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
            Height          =   300
            Index           =   20
            Left            =   11970
            Top             =   60
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "増減管理"
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
            Height          =   330
            Index           =   19
            Left            =   13110
            Top             =   360
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
            ForeColor       =   16711680
            Caption         =   "便当"
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
            Height          =   330
            Index           =   18
            Left            =   12540
            Top             =   360
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
            ForeColor       =   16711680
            Caption         =   "設定"
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
            Height          =   330
            Index           =   17
            Left            =   11970
            Top             =   360
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
            ForeColor       =   16711680
            Caption         =   "区分"
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
            Height          =   330
            Index           =   16
            Left            =   10950
            Top             =   360
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
            ForeColor       =   16711680
            Caption         =   "今回"
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
            Height          =   330
            Index           =   15
            Left            =   10380
            Top             =   360
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
            ForeColor       =   16711680
            Caption         =   "前回"
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
            Height          =   300
            Index           =   14
            Left            =   10380
            Top             =   60
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "回転枚数"
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
            Height          =   630
            Index           =   13
            Left            =   8850
            Top             =   60
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   1111
            ForeColor       =   16711680
            Caption         =   "設定期間"
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
            Height          =   630
            Index           =   12
            Left            =   11520
            Top             =   60
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   1111
            ForeColor       =   16711680
            Caption         =   "有効"
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
            cPositionX      =   20
            cPositionY      =   -120
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   630
            Index           =   11
            Left            =   8040
            Top             =   60
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1111
            ForeColor       =   16711680
            Caption         =   "内示数"
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
            Height          =   630
            Index           =   10
            Left            =   7230
            Top             =   60
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1111
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   630
            Index           =   9
            Left            =   3090
            Top             =   60
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   1111
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   630
            Index           =   8
            Left            =   4140
            Top             =   60
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   1111
            ForeColor       =   16711680
            Caption         =   "品          番"
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
            Height          =   630
            Index           =   7
            Left            =   2760
            Top             =   60
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   1111
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionX      =   20
            cPositionY      =   50
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   630
            Index           =   6
            Left            =   630
            Top             =   60
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   1111
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D HEAD_No 
            Height          =   630
            Left            =   60
            Top             =   60
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   1111
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
            cPositionX      =   30
         End
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   8655
            Left            =   30
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   330
            Width           =   15090
            _Version        =   196608
            _ExtentX        =   26617
            _ExtentY        =   15266
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
            FormatString    =   "№  |仕入先           |受|背番号  |品番                     |収容数|内示数|設定期間    |前回|今回|有 |増減|設定|便当|平準|設定|"
            Rows            =   26
            Cols            =   42
            BackColor       =   12648384
            BackColorFixed  =   14737632
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   330
            FillStyle       =   1
            GridLines       =   2
            SelectionMode   =   1
            Begin Cis3D_v60.CIS3D CIS3D1 
               Height          =   330
               Index           =   22
               Left            =   0
               Top             =   0
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   582
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
               cPositionX      =   30
               Begin Cis3D_v60.CIS3D CIS3D1 
                  Height          =   330
                  Index           =   23
                  Left            =   0
                  Top             =   0
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   582
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
                  cPositionX      =   30
               End
            End
            Begin Cis3D_v60.CIS3D Input_Row 
               Height          =   390
               Left            =   570
               Top             =   330
               Width           =   14220
               _ExtentX        =   25083
               _ExtentY        =   688
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
               Begin Cis3D_v60.CIS3D B1lb_Hinbn 
                  Height          =   360
                  Left            =   3525
                  Top             =   0
                  Width           =   3075
                  _ExtentX        =   5424
                  _ExtentY        =   635
                  BackColor       =   16761024
                  Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
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
                  cPositionX      =   50
               End
               Begin Cis3D_v60.CIS3D B1lb_Seban 
                  Height          =   360
                  Left            =   2475
                  Top             =   0
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   635
                  BackColor       =   16761024
                  Caption         =   "XXXXXXXX"
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
                  cPositionX      =   50
               End
               Begin Cis3D_v60.CIS3D B1lb_Uke 
                  Height          =   360
                  Left            =   2145
                  Top             =   0
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   635
                  BackColor       =   16761024
                  Caption         =   "XX"
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
               Begin CisText_V60.CisText iB1_ZKbn 
                  Height          =   360
                  Left            =   11340
                  TabIndex        =   13
                  Top             =   15
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   635
                  cFormat         =   "#"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
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
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "#"
                  cILength        =   1
               End
               Begin CisText_V60.CisText iB1_ZBinSu 
                  Height          =   375
                  Left            =   12480
                  TabIndex        =   15
                  Top             =   0
                  Width           =   600
                  _ExtentX        =   1058
                  _ExtentY        =   661
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
               Begin CisText_V60.CisText iB1_HSetSu 
                  Height          =   375
                  Left            =   13650
                  TabIndex        =   17
                  Top             =   0
                  Width           =   540
                  _ExtentX        =   953
                  _ExtentY        =   661
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
               Begin CisText_V60.CisText iB1_Seban 
                  Height          =   375
                  Left            =   2460
                  TabIndex        =   11
                  Top             =   0
                  Width           =   1065
                  _ExtentX        =   1879
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
                  Text            =   "XXXXXXXX"
                  MaxLength       =   8
                  IMEMode         =   2
                  cDataReplace    =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_SetYM 
                  Height          =   360
                  Left            =   8220
                  Top             =   0
                  Width           =   1530
                  _ExtentX        =   2699
                  _ExtentY        =   635
                  BackColor       =   16761024
                  ForeColor       =   0
                  Caption         =   "MM/DD～MM/DD"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_YukoMai 
                  Height          =   360
                  Left            =   10890
                  Top             =   0
                  Width           =   450
                  _ExtentX        =   794
                  _ExtentY        =   635
                  BackColor       =   16761024
                  ForeColor       =   0
                  Caption         =   "###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -50
               End
               Begin CisText_V60.CisText iB1_ZSetSu 
                  Height          =   375
                  Left            =   11910
                  TabIndex        =   14
                  Top             =   0
                  Width           =   600
                  _ExtentX        =   1058
                  _ExtentY        =   661
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
                  cILength        =   2
               End
               Begin Cis3D_v60.CIS3D B1lb_ZKbnm 
                  Height          =   360
                  Left            =   11580
                  Top             =   0
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   635
                  BackColor       =   16761024
                  Caption         =   "増"
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
               Begin Cis3D_v60.CIS3D Back_Torcd 
                  Height          =   360
                  Left            =   0
                  Top             =   0
                  Width           =   2145
                  _ExtentX        =   3784
                  _ExtentY        =   635
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
                  Begin Cis3D_v60.CIS3D B1lb_Torcd 
                     Height          =   360
                     Left            =   45
                     Top             =   0
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   635
                     BackColor       =   16761024
                     Caption         =   "XXXXXXX"
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
                     cPositionX      =   50
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Tornm 
                     Height          =   375
                     Left            =   930
                     Top             =   0
                     Width           =   1185
                     _ExtentX        =   2090
                     _ExtentY        =   661
                     BackColor       =   16761024
                     Caption         =   "XXXXXXXXX1XXXXXXXXXX"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   9
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
                     cPositionX      =   50
                  End
                  Begin CisText_V60.CisText iB1_Torcd 
                     Height          =   390
                     Left            =   15
                     TabIndex        =   9
                     Top             =   -15
                     Width           =   960
                     _ExtentX        =   1693
                     _ExtentY        =   688
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
               End
               Begin CisText_V60.CisText iB1_Uke 
                  Height          =   375
                  Left            =   2130
                  TabIndex        =   10
                  Top             =   0
                  Width           =   330
                  _ExtentX        =   582
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
                  Text            =   "XX"
                  MaxLength       =   2
                  cDataReplace    =   1
                  cFaZero         =   0
               End
               Begin CisText_V60.CisText iB1_Hinbn 
                  Height          =   375
                  Left            =   3510
                  TabIndex        =   12
                  Top             =   0
                  Width           =   3165
                  _ExtentX        =   5583
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
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   360
                  Left            =   6615
                  Top             =   0
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   635
                  BackColor       =   16761024
                  ForeColor       =   0
                  Caption         =   "##,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -30
               End
               Begin Cis3D_v60.CIS3D B1lb_Naisu 
                  Height          =   360
                  Left            =   7410
                  Top             =   0
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   635
                  BackColor       =   16761024
                  ForeColor       =   0
                  Caption         =   "##,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -50
               End
               Begin Cis3D_v60.CIS3D B1lb_KKMai 
                  Height          =   360
                  Left            =   10320
                  Top             =   0
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   635
                  BackColor       =   16761024
                  ForeColor       =   0
                  Caption         =   "###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -50
               End
               Begin Cis3D_v60.CIS3D B1lb_ZKMai 
                  Height          =   360
                  Left            =   9750
                  Top             =   0
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   635
                  BackColor       =   16761024
                  ForeColor       =   0
                  Caption         =   "###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -50
               End
               Begin CisText_V60.CisText iB1_HKbn 
                  Height          =   375
                  Left            =   13050
                  TabIndex        =   16
                  Top             =   0
                  Width           =   270
                  _ExtentX        =   476
                  _ExtentY        =   661
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
                  cFaZero         =   0
                  cFbComma        =   0
                  cGFormat        =   "#"
                  cILength        =   1
               End
               Begin Cis3D_v60.CIS3D B1lb_HKbnm 
                  Height          =   360
                  Left            =   13320
                  Top             =   0
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   635
                  BackColor       =   16761024
                  Caption         =   "日"
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
      Left            =   45
      ScaleHeight     =   1230
      ScaleWidth      =   11550
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   480
      Width           =   11610
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   1984
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
         cPositionX      =   30
         cPositionY      =   30
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "平準化"
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   1530
            TabIndex        =   24
            Top             =   300
            Width           =   1965
            Begin VB.CheckBox H1Chk_Heijyun 
               BackColor       =   &H00C0C0C0&
               Caption         =   "月"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   1320
               TabIndex        =   8
               Top             =   330
               Width           =   615
            End
            Begin VB.CheckBox H1Chk_Heijyun 
               BackColor       =   &H00C0C0C0&
               Caption         =   "便"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   60
               TabIndex        =   6
               Top             =   330
               Width           =   615
            End
            Begin VB.CheckBox H1Chk_Heijyun 
               BackColor       =   &H00C0C0C0&
               Caption         =   "日"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   690
               TabIndex        =   7
               Top             =   330
               Width           =   615
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "増減"
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   60
            TabIndex        =   23
            Top             =   300
            Width           =   1425
            Begin VB.CheckBox H1Chk_ZouGen 
               BackColor       =   &H00C0C0C0&
               Caption         =   "減"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   690
               TabIndex        =   5
               Top             =   330
               Width           =   615
            End
            Begin VB.CheckBox H1Chk_ZouGen 
               BackColor       =   &H00C0C0C0&
               Caption         =   "増"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   60
               TabIndex        =   4
               Top             =   330
               Width           =   615
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   315
            Left            =   990
            Top             =   30
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
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
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cBoderStyle     =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   3
         Left            =   6345
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1984
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
         cPositionX      =   30
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   360
            Left            =   60
            TabIndex        =   1
            Top             =   480
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   2
         Left            =   3615
         Top             =   60
         Width           =   2730
         _ExtentX        =   4815
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
            Left            =   30
            TabIndex        =   0
            Top             =   480
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
            Left            =   1035
            Top             =   480
            Width           =   1650
            _ExtentX        =   2910
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
         Left            =   6840
         Top             =   60
         Width           =   3225
         _ExtentX        =   5689
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
            Left            =   30
            TabIndex        =   2
            Top             =   480
            Width           =   3195
            _ExtentX        =   5636
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
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   5
         Left            =   10065
         Top             =   60
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "年 月"
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
         Begin CisYM_V60.CisYM iH1_Ym 
            Height          =   360
            Left            =   30
            TabIndex        =   3
            Top             =   480
            Width           =   1395
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
      Left            =   11700
      ScaleHeight     =   1290
      ScaleWidth      =   3495
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   435
      Width           =   3555
      Begin CisBtn_60.CisBtn PB_SMNT 
         Height          =   585
         Left            =   900
         Top             =   660
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "詳細"
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1755
         Top             =   660
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
         Left            =   2610
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
         Left            =   2610
         Top             =   660
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
         Top             =   660
         Visible         =   0   'False
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
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   30
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
         BtnCaption      =   "全て"
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
         Left            =   890
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
         BtnCaption      =   "登録済"
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1750
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
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【かんばん増減・平準化一括メンテナンス 】"
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
         Left            =   14385
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
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
         Height          =   300
         Left            =   11835
         Top             =   30
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   529
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
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   375
         Left            =   7935
         Top             =   0
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   661
         BackColor       =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
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
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   330
            Left            =   30
            Top             =   30
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   582
            Caption         =   "表示件数"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
            cPositionX      =   50
            Begin Cis3D_v60.CIS3D B1lb_Kensu 
               Height          =   285
               Left            =   990
               Top             =   15
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   503
               BackColor       =   16777152
               Caption         =   "##,###"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -50
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   330
            Left            =   1965
            Top             =   30
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   582
            ForeColor       =   255
            Caption         =   "操作件数"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
            cPositionX      =   50
            Begin Cis3D_v60.CIS3D B1lb_SKensu 
               Height          =   285
               Left            =   990
               Top             =   15
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   503
               BackColor       =   16777152
               ForeColor       =   255
               Caption         =   "##,###"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -50
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
         BackColor       =   16777215
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
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cBoderWidth     =   3
         cFont3D         =   2
         cAlingnment     =   7
      End
   End
   Begin VB.Timer Timer2 
      Left            =   705
      Top             =   -120
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   -150
   End
   Begin VB.TextBox Dummy 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Left            =   -15
      TabIndex        =   21
      Text            =   "Dummy"
      Top             =   -15
      Width           =   345
   End
   Begin VB.Label Color_Del 
      BackColor       =   &H00FFC0FF&
      Caption         =   "削除色"
      Height          =   345
      Left            =   2340
      TabIndex        =   25
      Top             =   11310
      Width           =   855
   End
End
Attribute VB_Name = "CKH0150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  かんばん増減一括メンテナンス（流動管理テーブルVer）
'**       フォームID    :  CKH0150
'**       処理概要      :  【かんばん流動管理ビュー】　使用
'**
'**       作  成  日    :  2005/02/15
'**       変  更  日    :  2006/03/16 By CIS 期間を求める時の終了日が変(開始日設定がある時)
'**       変  更  日    :  2006/03/17 By CIS 前回枚数の表示を前回回転枚数から前月回転枚数へ変更
'**                                          今回枚数の表示がされないのを修正
'**                        2006/03/20 By CIS 【無効】表示関連を修正
'**                        2006/03/31 By CIS 重複チェックのロジックを変更
'**                        2007/07/10 By CIS 基本開始日の日付加算修正
'**********************************************************'
    Dim CisFrm                  As CISFormContorl

    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKB                 As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM                 As String       ' 処理名称(新規･修正･削除 等)
    Dim CisVsGrid1              As New CisVsGrid3
    Dim ActObj                  As Object
    Dim ErrObj                  As Object
    Dim ActRow                  As Integer
    Dim mWork_Tag               As String
    Dim mWhere                  As String
'   増減区分の背景色
    Dim mColor_On               As Long
    Dim mColor_Off              As Long
'   グリッドの背景色
    Dim mColorOrg               As Long
    Dim mColorChg               As Long
'   行
    Dim NowRow                  As Long
    Dim NewRow                  As Long
    
    Dim mBaseStartDay           As Integer         '基本開始日
    Dim mStartDay               As Integer         '開始日
    Dim mStartYMD               As String
    Dim mEndYMD                 As String

    Dim mUpdate_Cnt             As Integer         '変更対象件数
    Dim mUpdate_Now             As Integer         '現在変更件数
    
    Dim mUpdate_Color           As Long
    Dim mNomal_Color            As Long
    Dim mNomal_FixColor         As Long

    Dim mBefore_Row             As Long
    
    
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

'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid1
        Set .GridObj = vsGrid1
'        .NoSpaceCol = 12
        .NoSpaceCol = 1
        .NumberSet = All
'        .EditCheckCol = 12
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 19
        .InitGet = False
        .Init
    End With
    Call GridInit1

'年月初期表示
    Call GetNJProc_YM(3)
    iH1_Ym = gn_HNYM
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
    
    PB_Look.Visible = False

    mUpdate_Cnt = 0
    mBefore_Row = 0
    
    mUpdate_Color = &HFFC0FF
    mNomal_Color = &HC0FFC0
    mNomal_FixColor = &HE0E0E0
    
    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

'   背景色の設定（グリッド）
    mColorOrg = vsGrid1.BackColor
    mColorChg = &HC0E0FF
'内示年月
    gSL_Select = "SELECT MAX(年月)年月 FROM かんばん回転枚数テーブル"
    Call KKTRead(gSL_Select, 1)
    If RTrim(KKT.年月) <> "" Then
        iH1_Ym = RTrim(KKT.年月)
    Else
        If CisFun.Val2(Format(Now, "DD")) >= 15 Then
           iH1_Ym = Format(DateAdd("M", 1, Now()), "YYYYMM")
        Else
           iH1_Ym = Format(Now(), "YYYYMM")
        End If
    End If
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

    Select Case Key_Code
        Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click        '【全て】
        Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click        '【登録済】
        Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click        '【削除】
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click      '【検索】
        Case vbKeyF11:      If PB_SMNT.Visible Then Call PB_SMNT_Click      '【詳細設定】
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_Can_Click        '【取消】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click        '【終了】
        Case vbKeyReturn
                            If PB_Ent.Visible Then                          '【入力】
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
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    If SyoriKB <> "ADD" Then
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
    End If
    iH1_Torcd.SetFocus
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
    End If
    iH1_Torcd.SetFocus
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        SyoriKB = "DEL"
    Else
        SyoriKB = ""
    End If
    Call SyoriName(SyoriKB)
    iH1_Torcd.SetFocus
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
Private Sub PB_Can_Click()
    Call CisFrm.MousePT(11)
    If ProcHB = "B1" Then
       ProcHB = "H1"
       Call DispChange(ProcHB)
       Call SyoriName(SyoriKB)
       Call HeadBodyClear("B")
       iH1_Torcd.SetFocus
       GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_Ukeir: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iB1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iB1_Uke" Then Call Look_Ukeir: Exit Sub
    If PB_Look.Tag = "iB1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iB1_Seban" Then Call Look_Hinbn: Exit Sub
End Sub
'+----------------------------+
'+      取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 9            ' [発注]手配先/契約先/社内検索

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        If mWork_Tag = "iH1_Torcd" Then
           iH1_Torcd = RV_TorcdK
           H1lb_Tornm = RV_TorRName
        Else
           iB1_Torcd = RV_TorcdK
           B1lb_Tornm = RV_TorRName
        End If
    End If
End Sub
'+----------------------------+
'+      取引先受入検索        +
'+----------------------------+
Private Sub Look_Ukeir()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 9            ' [発注]手配先/契約先/社内検索
    If mWork_Tag = "iH1_Ukeir" Then
       RV_Torcd = iH1_Torcd
    Else
       RV_Torcd = iB1_Torcd
    End If
    RV_Kojyo = ""

    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        If mWork_Tag = "iH1_Ukeir" Then
           iH1_Torcd = RV_TorcdK
           H1lb_Tornm = RV_TorRName
           iH1_Ukeir = RV_Ukeir
        Else
           iB1_Torcd = RV_TorcdK
           B1lb_Tornm = RV_TorRName
           iB1_Uke = RV_Ukeir
        End If
    End If
End Sub
'+----------------------------+
'+      品番検索              +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    If mWork_Tag = "iH1_Hinbn" Then
       RVI_Torcd = RTrim(iH1_Torcd)
       RVI_Ukeir = RTrim(iH1_Ukeir)
    Else
       RVI_Torcd = RTrim(iB1_Torcd)
       RVI_Ukeir = RTrim(iB1_Uke)
    End If
    RVI_Hinmoku = 1


    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        If mWork_Tag = "iH1_Hinbn" Then
           iH1_Hinbn = RV_Hinbn
           iH1_Torcd = RV_TorcdK
           iH1_Ukeir = RV_Ukeir
           H1lb_Tornm = ""
           If TorNmGet(iH1_Torcd, 9) Then
              H1lb_Tornm = TRM_RName
           End If
        Else
           iB1_Hinbn = RV_Hinbn
           iB1_Torcd = RV_TorcdK
           iB1_Seban = RV_Seban
           iB1_Uke = RV_Ukeir
           B1lb_Tornm = ""
           If TorNmGet(iB1_Torcd, 9) Then
              B1lb_Tornm = TRM_RName
           End If
        End If
    End If
End Sub
'+----------------------------+
'+      詳細画面表示          +
'+----------------------------+
Private Sub PB_SMNT_Click()
    Dim wRow        As Integer
    Dim wTRow       As Integer
    
    If RTrim(iB1_Hinbn) = "" Then
       Exit Sub
    End If
    
    RV_Left = 0
    RV_Top = 0
    RV_YM = RTrim(iH1_Ym)
    RV_Seban = RTrim(iB1_Seban)
    RV_Torcd = RTrim(iB1_Torcd)
    RV_TorRName = RTrim(B1lb_Tornm)
    RV_Ukeir = RTrim(iB1_Uke)
    RV_Hinbn = RTrim(iB1_Hinbn)
    RV_ZG_DSPHin = vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(4))
    RV_ZG_Syuyo = CisFun.Val2(B1lb_Syuyo)
    RV_ZG_Naisu = CisFun.Val2(B1lb_Naisu)
    RV_SYmd = vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(18))
    RV_EYmd = vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(19))
    RV_ZG_ZenKai = CisFun.Val2(B1lb_ZKMai)
    RV_ZG_KonKai = CisFun.Val2(B1lb_KKMai)
    RV_ZG_Yuko = CisFun.Val2(B1lb_YukoMai)
    RV_ZG_ZGKB = CisFun.Val2(iB1_ZKbn)
    RV_ZG_ZGSet = CisFun.Val2(iB1_ZSetSu)
    RV_ZG_ZGBin = CisFun.Val2(iB1_ZBinSu)
    RV_ZG_HJKB = CisFun.Val2(iB1_HKbn)
    RV_ZG_HJSet = CisFun.Val2(iB1_HSetSu)
            
    Rv_Call2 = "X"
    
    wRow = vsGrid1.Row
    wTRow = vsGrid1.TopRow

''(退避データチェック)
'    If Not SaveDateCheck Then
''        iB1_Torcd.SetFocus
'        PNL_Wait.Visible = False
'        Message = "【抽出中】しばらくお待ち下さい"
'        Exit Sub
'    End If
'    PNL_Wait.Visible = False
'    Message = "【抽出中】しばらくお待ち下さい"
    
    CKH0170.Show vbModal
    Unload CKH0170
    Set CKH0170 = Nothing

'    mUpdate_Cnt = 0
    
    If Rv_CallRtn Then
       
       Input_Row.Visible = False
       Call Grid_ReDsp
       Call VsGridToText(vsGrid1.Row)
       Input_Row.Visible = True
       If wRow <= vsGrid1.Rows - 1 Then
           vsGrid1.TopRow = wTRow
           vsGrid1.Row = wRow
       Else
           vsGrid1.Row = 1
       End If
       Call VsGridToText(vsGrid1.Row)
       DoEvents
       
'       Input_Row.Visible = False
'       Call HeadBodyClear("B")
'       Call Head1Chk
'       Input_Row.Visible = True
'       Call VsGridToText(vsGrid1.Row)
'
'        If wRow <= vsGrid1.Rows - 1 Then
'            vsGrid1.TopRow = wTRow
'            vsGrid1.Row = wRow
'        Else
'            vsGrid1.Row = 1
'        End If
'        Call VsGridToText(vsGrid1.Row)
'       DoEvents

    End If
End Sub

'+-----------------------------+
'+  iH1_Torcd
'+-----------------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    mWork_Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_Tornm = ""
    If TorNmGet(iH1_Torcd, 9) Then
       H1lb_Tornm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iH1_Ukeir
'+-----------------------------+
Private Sub iH1_Ukeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    mWork_Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_Hinbn
'+-----------------------------+
Private Sub iH1_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    mWork_Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Torcd
'+-----------------------------+
Private Sub iB1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    mWork_Tag = ActiveControl.Name
    PB_Look.Visible = True
    Set ActObj = iB1_Torcd
    Back_Torcd.Tag = "S"
End Sub
Private Sub iB1_Torcd_LostFocus()
    If iB1_Torcd.Enabled = True Then
        If ActiveControl.Name <> "PB_Look" Then
            PB_Look.Tag = ""
            PB_Look.Visible = False
        End If
    End If

    If RTrim(iB1_Torcd) <> "" Then
       If TorNmGet(iB1_Torcd, 9) Then
           B1lb_Tornm = RTrim(TRM.略称)
       Else
           B1lb_Tornm = ""
       End If
    Else
       B1lb_Tornm = ""
    End If
    Set ActObj = iB1_Torcd
    Back_Torcd.Tag = "E"

    If vsGrid1.Rows < CisFun.Val2(Input_Row.Tag) Then Exit Sub                           '取消時の対応
    If RTrim(iB1_Torcd) <> "" And vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(34)) = "" Then
       
        If RTrim(iB1_Hinbn) <> "" Then
           '収容数
           Call Syuyo_Set
               
           '開始・終了日
           Call YMD_Set
        
           If HNaiji_Chk(iH1_Ym, iB1_Hinbn, iB1_Torcd, iB1_Uke) Then
              B1lb_Naisu = Format(HNJ.当月内示数, "###,###")
              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(6)) = Format(HNJ.当月内示数, "###,###")
           Else
              B1lb_Naisu = ""
              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(6)) = ""
           End If
           
'           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
'              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
'              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
'            End If
        End If
    End If

    If RTrim(iB1_Torcd) <> RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) Then
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
          vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
          mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
                      
          vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
          vsGrid1.CellBackColor = mUpdate_Color
          vsGrid1.ColSel = 15
          vsGrid1.CellBackColor = mUpdate_Color
        End If
    Else
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
          If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
             RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
             RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
             RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
             CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
             CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
             CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
             CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
             CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                  vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                  mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
          
                  vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                  vsGrid1.CellBackColor = mNomal_FixColor
                  vsGrid1.ColSel = 15
                  vsGrid1.CellBackColor = mNomal_Color
          End If
       End If
    End If

    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If

End Sub
'+-----------------------------+
'+  iB1_Uke
'+-----------------------------+
Private Sub iB1_Uke_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    mWork_Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Uke_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    If vsGrid1.Rows < CisFun.Val2(Input_Row.Tag) Then Exit Sub                           '取消時の対応
    If RTrim(iB1_Uke) <> "" And vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(34)) = "" Then
       
        If RTrim(iB1_Hinbn) <> "" Then
           '収容数
           Call Syuyo_Set
               
           '開始・終了日
           Call YMD_Set
             
           If HNaiji_Chk(iH1_Ym, iB1_Hinbn, iB1_Torcd, iB1_Uke) Then
              B1lb_Naisu = Format(HNJ.当月内示数, "###,###")
              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(6)) = Format(HNJ.当月内示数, "###,###")
           Else
              B1lb_Naisu = ""
              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(6)) = ""
           End If
           
'           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
'              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
'              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
'           End If
        End If
    End If

    If RTrim(iB1_Uke) <> RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) Then
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
          vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
          mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
                      
          vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
          vsGrid1.CellBackColor = mUpdate_Color
          vsGrid1.ColSel = 15
          vsGrid1.CellBackColor = mUpdate_Color
        End If
    Else
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
          If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
             RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
             RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
             RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
             CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
             CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
             CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
             CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
             CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                  vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                  mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
          
                  vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                  vsGrid1.CellBackColor = mNomal_FixColor
                  vsGrid1.ColSel = 15
                  vsGrid1.CellBackColor = mNomal_Color
          End If
       End If
    End If

    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If

End Sub
'+-----------------------------+
'+  iB1_Hinbn
'+-----------------------------+
Private Sub iB1_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    mWork_Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Hinbn_LostFocus()
    Dim wYmT        As String
    Dim wYmY        As String
'
    With CisFun
        wYmT = iH1_Ym
        .Date_RsFormat = fYM
        .Date_Add iH1_Ym, M, 1
        wYmY = .Date_Result
    End With
'
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    
    If vsGrid1.Rows < CisFun.Val2(Input_Row.Tag) Then Exit Sub                           '取消時の対応
    If RTrim(iB1_Hinbn) <> "" And vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(34)) = "" Then
       
       '収容数
       Call Syuyo_Set
       '開始・終了日
       Call YMD_Set
         
       If HNaiji_Chk(iH1_Ym, iB1_Hinbn, iB1_Torcd, iB1_Uke) Then
          B1lb_Naisu = Format(HNJ.当月内示数, "###,###")
          vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(6)) = Format(HNJ.当月内示数, "###,###")
       Else
          B1lb_Naisu = ""
          vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(6)) = ""
       End If
          
'       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
'          vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
'          mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
'        End If
    Else
        If RTrim(iB1_Hinbn) = "" Then
           B1lb_Naisu = "": vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(6)) = ""
           B1lb_Syuyo = "": vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(5)) = ""
           B1lb_SetYM = "": vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(7)) = ""
        End If
    End If

    If RTrim(iB1_Hinbn) <> RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) Then
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
          vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
          mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
                      
          vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
          vsGrid1.CellBackColor = mUpdate_Color
          vsGrid1.ColSel = 15
          vsGrid1.CellBackColor = mUpdate_Color
        End If
    Else
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
          If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
             RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
             RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
             RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
             CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
             CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
             CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
             CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
             CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                  vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                  mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
          
                  vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                  vsGrid1.CellBackColor = mNomal_FixColor
                  vsGrid1.ColSel = 15
                  vsGrid1.CellBackColor = mNomal_Color
          End If
       End If
    End If

    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If

End Sub
'+-----------------------------+
'+  iB1_Seban
'+-----------------------------+
Private Sub iB1_Seban_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    mWork_Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Seban_LostFocus()
    Dim wYmT        As String
    Dim wYmY        As String
'
    With CisFun
        wYmT = iH1_Ym
        .Date_RsFormat = fYM
        .Date_Add iH1_Ym, M, 1
        wYmY = .Date_Result
    End With
'
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    If RTrim(iB1_Seban) = "" Then Exit Sub
    gSL_Select = "SELECT * FROM 品番取引先マスタ "
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iB1_Torcd) & "'"
    gSL_Select = gSL_Select & " AND   受入 = '" & RTrim(iB1_Uke) & "'"
    gSL_Select = gSL_Select & " AND   背番号 = '" & RTrim(iB1_Seban) & "'"
    If HTMRead(gSL_Select, 1) Then
       iB1_Hinbn = RTrim(HTM.品番)
    End If

    If RTrim(iB1_Seban) <> "" And vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(34)) = "" Then
       If vsGrid1.Rows < CisFun.Val2(Input_Row.Tag) Then Exit Sub                           '取消時の対応
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
          vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
          mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
        End If
    
       If RTrim(iB1_Hinbn) <> "" Then
          '収容数
          Call Syuyo_Set
          '開始・終了日
          Call YMD_Set
       
          If HNaiji_Chk(iH1_Ym, iB1_Hinbn, iB1_Torcd, iB1_Uke) Then
             B1lb_Naisu = Format(HNJ.当月内示数, "###,###")
             vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(6)) = Format(HNJ.当月内示数, "###,###")
          Else
             B1lb_Naisu = ""
             vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(6)) = ""
          End If
       End If
    End If

    If RTrim(iB1_Seban) <> RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) Then
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
          vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
          mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
                      
          vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
          vsGrid1.CellBackColor = mUpdate_Color
          vsGrid1.ColSel = 15
          vsGrid1.CellBackColor = mUpdate_Color
        End If
    Else
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
          If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
             RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
             RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
             RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
             CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
             CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
             CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
             CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
             CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                  vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                  mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
          
                  vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                  vsGrid1.CellBackColor = mNomal_FixColor
                  vsGrid1.ColSel = 15
                  vsGrid1.CellBackColor = mNomal_Color
          End If
       End If
    End If

    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If

End Sub
Private Function HNaiji_Chk(SYM As String, sHinbn As String, sTorcd As String, sUke As String)
    Dim wYmT        As String
    Dim wYmY        As String
'
    With CisFun
        wYmT = iH1_Ym
        .Date_RsFormat = fYM
        .Date_Add iH1_Ym, M, 1
        wYmY = .Date_Result
    End With
'
    HNaiji_Chk = False
    
    gSL_Select = "SELECT HN.品番,HN.手配先,HN.受入,"
    gSL_Select = gSL_Select & " SUM(CASE WHEN SUBSTRING(HN.納期, 1, 6) = '" & wYmT & "' THEN HN.数量 ELSE 0 END) 当月内示数"
    gSL_Select = gSL_Select & " FROM 発注内示テーブル HN"
    gSL_Select = gSL_Select & " WHERE HN.内示年月 = '" & SYM & "'"
    gSL_Select = gSL_Select & " AND  HN.品番 = '" & RTrim(sHinbn) & "'"
    gSL_Select = gSL_Select & " AND HN.手配先 = '" & RTrim(sTorcd) & "'"
    gSL_Select = gSL_Select & " AND HN.受入 = '" & RTrim(sUke) & "'"
    gSL_Select = gSL_Select & " GROUP BY HN.品番,HN.手配先,HN.受入"
    If Not HNJRead(gSL_Select, 1) Then
       Exit Function
    End If
    
    HNaiji_Chk = True

End Function
Private Sub YMD_Set()
    
    Call SYMD_Get(RTrim(iB1_Torcd), RTrim(iB1_Uke), RTrim(iB1_Hinbn))
    If Trim(mEndYMD) = "" Then
        B1lb_SetYM = CisFun.Mid2(Format(mStartYMD, "0000/00/00"), 6, 5) & Space(7)
    Else
        B1lb_SetYM = CisFun.Mid2(Format(mStartYMD, "0000/00/00"), 6, 5) _
                     & "～" & CisFun.Mid2(Format(mEndYMD, "0000/00/00"), 6, 5)
    End If
    vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(7)) = B1lb_SetYM
     
    'ワークエリア
'    vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(18)) = KRV.開始日
'    vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(19)) = KRV.終了日
'    vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(35)) = KRV.開始日
'    vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(36)) = KRV.終了日

    vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(18)) = mStartYMD
    vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(19)) = mEndYMD
    vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(35)) = mStartYMD
    vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(36)) = mEndYMD
    
End Sub
Private Sub Syuyo_Set()

    gSL_Select = "SELECT * FROM 品番取引先マスタ "
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iB1_Torcd) & "'"
    gSL_Select = gSL_Select & " AND   受入 = '" & RTrim(iB1_Uke) & "'"
    gSL_Select = gSL_Select & " AND   品番 = '" & RTrim(iB1_Hinbn) & "'"
    If HTMRead(gSL_Select, 1) Then
       iB1_Seban = RTrim(HTM.背番号)
       If RTrim(HTM.変更日) <> "" Then
          If RTrim(HTM.変更日) < RTrim(iH1_Ym) & "01" Then
             B1lb_Syuyo = Format(HTM.収容数2, "##,###")
          Else
             B1lb_Syuyo = Format(HTM.収容数1, "##,###")
          End If
       Else
           B1lb_Syuyo = Format(HTM.収容数1, "##,###")
       End If
       vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(5)) = Format(HTM.収容数1, "##,###")
    End If

End Sub
'+-----------------------------+
'+  iB1_ZKbn
'+-----------------------------+
Private Sub iB1_ZKbn_LostFocus()
'    Select Case iB1_ZKbn
'           Case "1": B1lb_ZKbnm = "減": iB1_ZSetSu.Enabled = True: iB1_ZBinSu.Enabled = True: iB1_ZSetSu.SetFocus
'           Case "2": B1lb_ZKbnm = "増": iB1_ZBinSu = "": iB1_ZSetSu.Enabled = True: iB1_ZBinSu.Enabled = False: iB1_ZSetSu.SetFocus
'           Case Else: B1lb_ZKbnm = "": iB1_ZKbn = "": iB1_ZSetSu = "": iB1_ZBinSu = "": iB1_ZSetSu.Enabled = False: iB1_ZBinSu.Enabled = False
'    End Select
    Select Case iB1_ZKbn
           Case "1": B1lb_ZKbnm = "減"
           Case "2": B1lb_ZKbnm = "増": iB1_ZBinSu = ""
           Case Else: B1lb_ZKbnm = "": iB1_ZKbn = "": iB1_ZSetSu = "": iB1_ZBinSu = ""
    End Select

    If vsGrid1.Rows < CisFun.Val2(Input_Row.Tag) Then Exit Sub                           '取消時の対応
    If CisFun.Val2(iB1_ZKbn) <> CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
          vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
          mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
                      
          vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
          vsGrid1.CellBackColor = mUpdate_Color
          vsGrid1.ColSel = 15
          vsGrid1.CellBackColor = mUpdate_Color
        End If
    Else
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
          If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
             RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
             RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
             RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
             CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
             CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
             CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
             CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
             CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                  vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                  mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
          
                  vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                  vsGrid1.CellBackColor = mNomal_FixColor
                  vsGrid1.ColSel = 15
                  vsGrid1.CellBackColor = mNomal_Color
          End If
       End If
    End If

    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If

End Sub
'+-----------------------------+
'+  iB1_HKbn
'+-----------------------------+
Private Sub iB1_HKbn_LostFocus()
'    Select Case iB1_HKbn
'           Case "1": B1lb_HKbnm = "日": iB1_HSetSu.Enabled = True: iB1_HSetSu.SetFocus
'           Case "2": B1lb_HKbnm = "便": iB1_HSetSu.Enabled = True: iB1_HSetSu.SetFocus
'           Case "3": B1lb_HKbnm = "月": iB1_HSetSu.Enabled = True: iB1_HSetSu.SetFocus
'           Case Else: B1lb_HKbnm = "": iB1_HKbn = "": iB1_HSetSu = "": iB1_HSetSu.Enabled = False
'    End Select
    Select Case iB1_HKbn
           Case "1": B1lb_HKbnm = "日"
           Case "2": B1lb_HKbnm = "便"
           Case "3": B1lb_HKbnm = "月"
           Case Else: B1lb_HKbnm = "": iB1_HKbn = "": iB1_HSetSu = ""
    End Select

    If vsGrid1.Rows < CisFun.Val2(Input_Row.Tag) Then Exit Sub                           '取消時の対応
    If CisFun.Val2(iB1_HKbn) <> CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) Then
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
          vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
          mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
                      
          vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
          vsGrid1.CellBackColor = mUpdate_Color
          vsGrid1.ColSel = 15
          vsGrid1.CellBackColor = mUpdate_Color
        End If
    Else
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
          If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
             RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
             RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
             RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
             CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
             CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
             CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
             CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
             CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                  vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                  mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
          
                  vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                  vsGrid1.CellBackColor = mNomal_FixColor
                  vsGrid1.ColSel = 15
                  vsGrid1.CellBackColor = mNomal_Color
          End If
       End If
    End If

    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If

End Sub
Private Sub iB1_ZSetSu_LostFocus()
    If vsGrid1.Rows < CisFun.Val2(Input_Row.Tag) Then Exit Sub                           '取消時の対応
    If CisFun.Val2(iB1_ZSetSu) <> CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) Then
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
          vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
          mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

          vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
          vsGrid1.CellBackColor = mUpdate_Color
          vsGrid1.ColSel = 15
          vsGrid1.CellBackColor = mUpdate_Color
        End If
    Else
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
          If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
             RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
             RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
             RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
             CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
             CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
             CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
             CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
             CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                  vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                  mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

                  vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                  vsGrid1.CellBackColor = mNomal_FixColor
                  vsGrid1.ColSel = 15
                  vsGrid1.CellBackColor = mNomal_Color
          End If
       End If
    End If

    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If

End Sub
Private Sub iB1_ZBinSu_LostFocus()
    If vsGrid1.Rows < CisFun.Val2(Input_Row.Tag) Then Exit Sub                           '取消時の対応
    If CisFun.Val2(iB1_ZBinSu) <> CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) Then
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
          vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
          mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
                      
          vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
          vsGrid1.CellBackColor = mUpdate_Color
          vsGrid1.ColSel = 15
          vsGrid1.CellBackColor = mUpdate_Color
        End If
    Else
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
          If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
             RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
             RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
             RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
             CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
             CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
             CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
             CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
             CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                  vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                  mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
          
                  vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                  vsGrid1.CellBackColor = mNomal_FixColor
                  vsGrid1.ColSel = 15
                  vsGrid1.CellBackColor = mNomal_Color
          End If
       End If
    End If
    
    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If

End Sub
Private Sub iB1_HSetSu_LostFocus()
    If vsGrid1.Rows < CisFun.Val2(Input_Row.Tag) Then Exit Sub                           '取消時の対応
    If CisFun.Val2(iB1_HSetSu) <> CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) Then
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
          vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
          mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
                      
          vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
          vsGrid1.CellBackColor = mUpdate_Color
          vsGrid1.ColSel = 15
          vsGrid1.CellBackColor = mUpdate_Color
        End If
    Else
       If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
          If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
             RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
             RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
             RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
             CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
             CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
             CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
             CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
             CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                  vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                  mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
          
                  vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                  vsGrid1.CellBackColor = mNomal_FixColor
                  vsGrid1.ColSel = 15
                  vsGrid1.CellBackColor = mNomal_Color
          End If
       End If
    End If

    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If

End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If Back_Torcd.Tag = "S" Then
        Call TextToVsGrid(Val(Input_Row.Tag))
        Call Grid_RowChange(vbKeyUp)
    End If
    If Back_Torcd.Tag = "E" Then
        Call TextToVsGrid(Val(Input_Row.Tag))
        Call Grid_RowChange(vbKeyDown)
    End If
    SendKeys "{Tab}"
End Sub
'+-------------------------------+
'+      iB1_Torcd
'+-------------------------------+
Private Sub iB1_Torcd_KeyDown(KeyCode As Integer, Shift As Integer)
    
'    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
'        If RTrim(iB1_Torcd) <> "" Then
'           If vsGrid1.Rows < CisFun.Val2(Input_Row.Tag) Then Exit Sub                           '取消時の対応
'           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
'              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
'              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
'            End If
'        End If
'    End If

    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If RTrim(iB1_Torcd) <> RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) Then
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

              vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
              vsGrid1.CellBackColor = mUpdate_Color
              vsGrid1.ColSel = 15
              vsGrid1.CellBackColor = mUpdate_Color
           End If
        Else
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
              If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
                 RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
                 RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
                 RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
                 CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
                 CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
                 CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
                 CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
                 CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                      vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                      mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

                      vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                      vsGrid1.CellBackColor = mNomal_FixColor
                      vsGrid1.ColSel = 15
                      vsGrid1.CellBackColor = mNomal_Color
              End If
           End If
        End If
    End If
    
    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If
    
    Call Grid_RowChange(KeyCode)

End Sub
'+-------------------------------+
'+      iB1_Uke
'+-------------------------------+
Private Sub iB1_Uke_KeyDown(KeyCode As Integer, Shift As Integer)
    
'    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
'        If RTrim(iB1_Uke) <> "" Then
'           If vsGrid1.Rows < CisFun.Val2(Input_Row.Tag) Then Exit Sub                           '取消時の対応
'           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
'              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
'              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
'            End If
'        End If
'    End If
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If RTrim(iB1_Uke) <> RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) Then
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

              vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
              vsGrid1.CellBackColor = mUpdate_Color
              vsGrid1.ColSel = 15
              vsGrid1.CellBackColor = mUpdate_Color
           End If
        Else
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
              If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
                 RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
                 RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
                 RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
                 CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
                 CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
                 CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
                 CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
                 CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                      vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                      mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

                      vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                      vsGrid1.CellBackColor = mNomal_FixColor
                      vsGrid1.ColSel = 15
                      vsGrid1.CellBackColor = mNomal_Color
              End If
           End If
        End If
    
    End If

    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If
    
    Call Grid_RowChange(KeyCode)

End Sub
'+-------------------------------+
'+      iB1_Seban
'+-------------------------------+
Private Sub iB1_Seban_KeyDown(KeyCode As Integer, Shift As Integer)
    
'    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
'        If RTrim(iB1_Seban) <> "" Then
'           If vsGrid1.Rows < CisFun.Val2(Input_Row.Tag) Then Exit Sub                           '取消時の対応
'           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
'              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
'              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
'            End If
'        End If
'    End If
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If RTrim(iB1_Seban) <> RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) Then
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

              vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
              vsGrid1.CellBackColor = mUpdate_Color
              vsGrid1.ColSel = 15
              vsGrid1.CellBackColor = mUpdate_Color
           End If
        Else
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
              If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
                 RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
                 RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
                 RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
                 CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
                 CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
                 CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
                 CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
                 CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                      vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                      mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

                      vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                      vsGrid1.CellBackColor = mNomal_FixColor
                      vsGrid1.ColSel = 15
                      vsGrid1.CellBackColor = mNomal_Color
              End If
           End If
        End If
    
    End If
    
    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If
    
    Call Grid_RowChange(KeyCode)

End Sub
'+-------------------------------+
'+      iB1_Hinbn
'+-------------------------------+
Private Sub iB1_Hinbn_KeyDown(KeyCode As Integer, Shift As Integer)
    
'    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
'        If RTrim(iB1_Hinbn) <> "" Then
'           If vsGrid1.Rows < CisFun.Val2(Input_Row.Tag) Then Exit Sub                           '取消時の対応
'           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
'              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
'              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")
'            End If
'        End If
'    End If
'
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If RTrim(iB1_Hinbn) <> RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) Then
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

              vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
              vsGrid1.CellBackColor = mUpdate_Color
              vsGrid1.ColSel = 15
              vsGrid1.CellBackColor = mUpdate_Color
           End If
        Else
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
              If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
                 RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
                 RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
                 RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
                 CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
                 CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
                 CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
                 CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
                 CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                      vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                      mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

                      vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                      vsGrid1.CellBackColor = mNomal_FixColor
                      vsGrid1.ColSel = 15
                      vsGrid1.CellBackColor = mNomal_Color
              End If
           End If
        End If
    End If

    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If
    
    Call Grid_RowChange(KeyCode)

End Sub
'+-------------------------------+
'+      iB1_ZKbn
'+-------------------------------+
Private Sub iB1_ZKbn_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
'        Select Case iB1_ZKbn
'               Case "1": B1lb_ZKbnm = "減": iB1_ZSetSu.Enabled = True: iB1_ZBinSu.Enabled = True
'               Case "2": B1lb_ZKbnm = "増": iB1_ZBinSu = "": iB1_ZSetSu.Enabled = True: iB1_ZBinSu.Enabled = False
'               Case Else: B1lb_ZKbnm = "": iB1_ZKbn = "": iB1_ZSetSu = "": iB1_ZBinSu = "": iB1_ZSetSu.Enabled = False: iB1_ZBinSu.Enabled = False
'        End Select
        Select Case iB1_ZKbn
               Case "1": B1lb_ZKbnm = "減"
               Case "2": B1lb_ZKbnm = "増": iB1_ZBinSu = ""
               Case Else: B1lb_ZKbnm = "": iB1_ZKbn = "": iB1_ZSetSu = "": iB1_ZBinSu = ""
        End Select
        
        If CisFun.Val2(iB1_ZKbn) <> CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

              vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
              vsGrid1.CellBackColor = mUpdate_Color
              vsGrid1.ColSel = 15
              vsGrid1.CellBackColor = mUpdate_Color
           End If
        Else
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
              If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
                 RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
                 RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
                 RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
                 CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
                 CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
                 CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
                 CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
                 CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                      vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                      mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

                      vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                      vsGrid1.CellBackColor = mNomal_FixColor
                      vsGrid1.ColSel = 15
                      vsGrid1.CellBackColor = mNomal_Color
              End If
           End If
        End If
    
    End If
    
    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If
    
    Call Grid_RowChange(KeyCode)

End Sub
'+-------------------------------+
'+      iB1_ZSetSu
'+-------------------------------+
Private Sub iB1_ZSetSu_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If CisFun.Val2(iB1_ZSetSu) <> CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) Then
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

              vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
              vsGrid1.CellBackColor = mUpdate_Color
              vsGrid1.ColSel = 15
              vsGrid1.CellBackColor = mUpdate_Color
           End If
        Else
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
              If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
                 RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
                 RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
                 RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
                 CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
                 CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
                 CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
                 CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
                 CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                      vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                      mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

                      vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                      vsGrid1.CellBackColor = mNomal_FixColor
                      vsGrid1.ColSel = 15
                      vsGrid1.CellBackColor = mNomal_Color
              End If
           End If
        End If
    
    End If

    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If
    
    Call Grid_RowChange(KeyCode)

'    Call iB1_ZSetSu_LostFocus
End Sub
'+-------------------------------+
'+      iB1_ZBinSu
'+-------------------------------+
Private Sub iB1_ZBinSu_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If CisFun.Val2(iB1_ZBinSu) <> CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) Then
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

              vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
              vsGrid1.CellBackColor = mUpdate_Color
              vsGrid1.ColSel = 15
              vsGrid1.CellBackColor = mUpdate_Color
           End If
        Else
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
              If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
                 RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
                 RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
                 RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
                 CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
                 CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
                 CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
                 CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
                 CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                      vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                      mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

                      vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                      vsGrid1.CellBackColor = mNomal_FixColor
                      vsGrid1.ColSel = 15
                      vsGrid1.CellBackColor = mNomal_Color
              End If
           End If
        End If
    
    End If
    
    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If
    
    Call Grid_RowChange(KeyCode)

End Sub
'+-------------------------------+
'+      iB1_HKbn
'+-------------------------------+
Private Sub iB1_HKbn_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        
'        Select Case iB1_HKbn
'               Case "1": B1lb_HKbnm = "日": iB1_HSetSu.Enabled = True
'               Case "2": B1lb_HKbnm = "便": iB1_HSetSu.Enabled = True
'               Case "3": B1lb_HKbnm = "月": iB1_HSetSu.Enabled = True
'               Case Else: B1lb_HKbnm = "": iB1_HKbn = "": iB1_HSetSu = "": iB1_HSetSu.Enabled = False
'        End Select
        Select Case iB1_HKbn
               Case "1": B1lb_HKbnm = "日"
               Case "2": B1lb_HKbnm = "便"
               Case "3": B1lb_HKbnm = "月"
               Case Else: B1lb_HKbnm = "": iB1_HKbn = "": iB1_HSetSu = ""
        End Select

        If CisFun.Val2(iB1_HKbn) <> CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) Then
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

              vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
              vsGrid1.CellBackColor = mUpdate_Color
              vsGrid1.ColSel = 15
              vsGrid1.CellBackColor = mUpdate_Color
           End If
        Else
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
              If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
                 RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
                 RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
                 RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
                 CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
                 CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
                 CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
                 CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
                 CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                      vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                      mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

                      vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                      vsGrid1.CellBackColor = mNomal_FixColor
                      vsGrid1.ColSel = 15
                      vsGrid1.CellBackColor = mNomal_Color
              End If
           End If
        End If
    
    End If

    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If
    
    Call Grid_RowChange(KeyCode)

End Sub
'+-------------------------------+
'+      iB1_HSetSu
'+-------------------------------+
Private Sub iB1_HSetSu_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
        If CisFun.Val2(iB1_HSetSu) <> CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) Then
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "" Then
              vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X"
              mUpdate_Cnt = mUpdate_Cnt + 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

              vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
              vsGrid1.CellBackColor = mUpdate_Color
              vsGrid1.ColSel = 15
              vsGrid1.CellBackColor = mUpdate_Color
           End If
        Else
           If vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = "X" Then
              If RTrim(iB1_Torcd) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(16))) And _
                 RTrim(iB1_Hinbn) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(17))) And _
                 RTrim(iB1_Uke) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(26))) And _
                 RTrim(iB1_Seban) = RTrim(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(27))) And _
                 CisFun.Val2(iB1_HSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(25))) And _
                 CisFun.Val2(iB1_ZBinSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(21))) And _
                 CisFun.Val2(iB1_ZSetSu) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(20))) And _
                 CisFun.Val2(iB1_HKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(24))) And _
                 CisFun.Val2(iB1_ZKbn) = CisFun.Val2(vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(23))) Then
                      vsGrid1.TextMatrix(CisFun.Val2(Input_Row.Tag), CisVsGrid1.FixedGet(40)) = ""
                      mUpdate_Cnt = mUpdate_Cnt - 1: B1lb_SKensu = Format(mUpdate_Cnt, "###,###")

                      vsGrid1.Row = CisFun.Val2(Input_Row.Tag): vsGrid1.Col = 0
                      vsGrid1.CellBackColor = mNomal_FixColor
                      vsGrid1.ColSel = 15
                      vsGrid1.CellBackColor = mNomal_Color
              End If
           End If
        End If
    End If
    
    Call Grid_RowChange(KeyCode)

End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim w_Color     As Long
   w_Color = &H80000008 '初期色ｾｯﾄ(黒)
   K_Sykbnm.BackColor = &HFFFFFF  '(白)
   K_Sykbnm.cFont3D = 0           ' 浮き
   Select Case SyoriKB
     Case "ADD"
          SyoriNM = "全て"
          K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
          K_Sykbnm.BackColor = &H0&      '(黒)
          K_Sykbnm.cFont3D = 2           ' ﾌﾗｯﾄ
          PB_ADD.Visible = False
     Case "MNT"
          SyoriNM = "登録済"
          K_Sykbnm.ForeColor = &HFF0000  '(青色)
     Case "DEL"
          SyoriNM = "削除"
          K_Sykbnm.ForeColor = &HFF&     '(赤色)
'          W_Color = &HFF&
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

    PB_Look.Visible = False
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    PB_SMNT.Visible = Not H1Mode
    If SyoriKB = "DEL" Then
       PB_SMNT.Visible = False
    End If
    PB_Look.Visible = H1Mode
    PB_ADD.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_DEL.Visible = H1Mode
    
    If Pro = "H1" Then
       mUpdate_Cnt = 0
    End If
    
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
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = ""
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
    If HeadBody = "B" Then
        Call CisVsGrid1.Clear
    End If
End Sub
'*********************************
'*          入力処理(Enter)      *
'*********************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
'
        ProcHB = "B1"
        Call DispChange(ProcHB)
        vsGrid1.Row = 1
        Call VsGridToText(vsGrid1.Row)
        Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
        If SyoriKB = "DEL" Then
           vsGrid1.HighLight = flexHighlightWithFocus
           vsGrid1.SetFocus
        Else
           vsGrid1.HighLight = 0
        End If
        
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
        Else
            If Not DBPut Then GoTo ReturnPress_Ed
        End If
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk(Optional pKB As Byte = 0) As Boolean
    Head1Chk = False

'   取引先
    If Trim(iH1_Torcd) <> "" Then
        If Not TorNmGet(iH1_Torcd, 9) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    仕入先　未登録！（取引先マスタ）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Torcd.SetFocus
           Exit Function
        End If
'       受入
        If Trim(iH1_Ukeir) <> "" Then
            gSL_Select = ""
            gSL_Select = gSL_Select & "select *"
            gSL_Select = gSL_Select & "  from 取引先受入マスタ"
            gSL_Select = gSL_Select & " where 取引先CD = '" & RTrim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "   and 取引先区分 = 1"
            gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(iH1_Ukeir) & "'"
            If Not TUMRead(gSL_Select, 1) Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    受入　未登録！（取引先受入マスタ）            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
               iH1_Ukeir.SetFocus
               Exit Function
            End If
        End If
    End If
'   品番
    If Trim(iH1_Hinbn) <> "" Then
    '   >> 半角ﾁｪｯｸ（品番）
        If CisFun.Len2(RTrim(iH1_Hinbn)) <> Len(RTrim(iH1_Hinbn)) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    品番は半角にて入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Hinbn.SetFocus
            Exit Function
        End If
        gSL_Select = ""
        gSL_Select = gSL_Select & "select *"
        gSL_Select = gSL_Select & "  from 品番マスタ"
        gSL_Select = gSL_Select & " where 品番 = '" & RTrim(iH1_Hinbn) & "'"
        If Not HNMRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    品番　未登録！（品番マスタ）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Hinbn.SetFocus
           Exit Function
        End If
    End If
'   年月
    If Trim(iH1_Ym) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    年月を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ym.SetFocus
        Exit Function
    Else
        If Not iH1_Ym.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    年月を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Ym.SetFocus
            Exit Function
        End If
    End If

    If pKB = 0 Then
       PNL_Wait.Visible = True
    End If
    
    DoEvents
    
    Call SqlSelectSet
    
    If Not DBInput("INV") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    該当データが存在しません｡            "
            .MB_MSG(3) = "    再度抽出条件を入力してください       "
            .MB_Title = "抽出処理"
            .MB_Button = Error
            .MBOX
        End With
        
        PNL_Wait.Visible = False
        DoEvents

        Exit Function
    End If

    PNL_Wait.Visible = False
    DoEvents

    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False

    Call TextToVsGrid(Val(Input_Row.Tag))
    Call GridTextEnabled(False)

    Call CisVsGrid1.EditEnable
'    Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ

    Dim ID          As Long
    If SyoriKB <> "DEL" Then
        With vsGrid1
            If mUpdate_Cnt = 0 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    データが変更されていません｡          "
                    .MB_Title = "変更処理"
                    .MB_Button = Error
                    .MBOX
                End With
                Call CisVsGrid1.EditEnable
                Call CisVsGrid1.EditVisible
                Call VsGridToText(vsGrid1.Row)
                Call GridTextEnabled(True)
                Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                Exit Function
            End If
            
            For gCnt1 = 1 To .Rows - 1
                If RTrim(.TextMatrix(gCnt1, 1)) = "" And _
                    RTrim(.TextMatrix(gCnt1, 2)) = "" And _
                    RTrim(.TextMatrix(gCnt1, 4)) = "" And _
                    RTrim(.TextMatrix(gCnt1, 11)) = "" And _
                    RTrim(.TextMatrix(gCnt1, 12)) = "" And _
                    RTrim(.TextMatrix(gCnt1, 13)) = "" And _
                    RTrim(.TextMatrix(gCnt1, 14)) = "" And _
                    RTrim(.TextMatrix(gCnt1, 15)) = "" Then
                    For gCnt2 = gCnt1 + 1 To .Rows - 1
                        If RTrim(.TextMatrix(gCnt2, 1)) <> "" Or _
                            RTrim(.TextMatrix(gCnt2, 2)) <> "" Then
                            For gCnt3 = 1 To 40
                                .TextMatrix(gCnt1, gCnt3) = .TextMatrix(gCnt2, gCnt3)
                            Next gCnt3
                            For gCnt3 = 1 To 40
                                .TextMatrix(gCnt2, gCnt3) = ""
                            Next gCnt3
                            '移動元クリアー
                            .Row = gCnt2: vsGrid1.Col = 0
                            .CellBackColor = mNomal_FixColor
                            .ColSel = 15
                            .CellBackColor = mNomal_Color
                            '移動先色更新
                            .Row = gCnt1: vsGrid1.Col = 0
                            .CellBackColor = mUpdate_Color
                            .ColSel = 15
                            .CellBackColor = mUpdate_Color
                                                        
                            Exit For
                        End If
                    Next gCnt2
                End If
            Next gCnt1
'
            mUpdate_Now = 0
            For ID = 1 To .Rows - 1
                If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(0))) = "" Then Exit For
                If mUpdate_Now = mUpdate_Cnt Then Exit For                                  '変更対象件数まで処理したら終了
                If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(40))) <> "" Then                '変更対象行のみチェック
                '   キー項目入力チェック
                    If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))) <> "" And Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(2))) = "" Then
                       With CisFun
                            .MB_Lines = 4
                            .MB_MSG(2) = "    " & ID & "行目 ： 【受入】未入力       "
                            .MB_Title = ""
                            .MB_Button = Error
                           .MBOX
                       End With
                       .Row = ID
                       Call CisVsGrid1.EditEnable
                       Call CisVsGrid1.EditVisible
                       Call VsGridToText(vsGrid1.Row)
                       Call GridTextEnabled(True)
                       Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                       Exit Function
                    End If
                    If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))) <> "" And Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) = "" Then
                       With CisFun
                            .MB_Lines = 4
                            .MB_MSG(2) = "    " & ID & "行目 ： 【品番】未入力       "
                            .MB_Title = ""
                            .MB_Button = Error
                           .MBOX
                       End With
                       .Row = ID
                       Call CisVsGrid1.EditEnable
                       Call CisVsGrid1.EditVisible
                       Call VsGridToText(vsGrid1.Row)
                       Call GridTextEnabled(True)
                       Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                       Exit Function
                    End If
                
                    If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(2))) <> "" And Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))) = "" Then
                       With CisFun
                            .MB_Lines = 4
                            .MB_MSG(2) = "    " & ID & "行目 ： 【仕入先】未入力       "
                            .MB_Title = ""
                            .MB_Button = Error
                           .MBOX
                       End With
                       .Row = ID
                       Call CisVsGrid1.EditEnable
                       Call CisVsGrid1.EditVisible
                       Call VsGridToText(vsGrid1.Row)
                       Call GridTextEnabled(True)
                       Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                       Exit Function
                    End If
                    If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(2))) <> "" And Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) = "" Then
                       With CisFun
                            .MB_Lines = 4
                            .MB_MSG(2) = "    " & ID & "行目 ： 【品番】未入力       "
                            .MB_Title = ""
                            .MB_Button = Error
                           .MBOX
                       End With
                       .Row = ID
                       Call CisVsGrid1.EditEnable
                       Call CisVsGrid1.EditVisible
                       Call VsGridToText(vsGrid1.Row)
                       Call GridTextEnabled(True)
                       Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                       Exit Function
                    End If
                
                    If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) <> "" And Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))) = "" Then
                       With CisFun
                            .MB_Lines = 4
                            .MB_MSG(2) = "    " & ID & "行目 ： 【仕入先】未入力       "
                            .MB_Title = ""
                            .MB_Button = Error
                           .MBOX
                       End With
                       .Row = ID
                       Call CisVsGrid1.EditEnable
                       Call CisVsGrid1.EditVisible
                       Call VsGridToText(vsGrid1.Row)
                       Call GridTextEnabled(True)
                       Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                       Exit Function
                    End If
                    If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) <> "" And Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(2))) = "" Then
                       With CisFun
                            .MB_Lines = 4
                            .MB_MSG(2) = "    " & ID & "行目 ： 【受入】未入力       "
                            .MB_Title = ""
                            .MB_Button = Error
                           .MBOX
                       End With
                       .Row = ID
                       Call CisVsGrid1.EditEnable
                       Call CisVsGrid1.EditVisible
                       Call VsGridToText(vsGrid1.Row)
                       Call GridTextEnabled(True)
                       Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                       Exit Function
                    End If
                
                    If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(30))) <> 0 Or CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(12))) <> 0 Or _
                       CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(33))) <> 0 Or _
                       CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(31))) <> 0 Or CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(15))) <> 0 Then
                       
                       If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))) = "" Then
                          With CisFun
                               .MB_Lines = 4
                               .MB_MSG(2) = "    " & ID & "行目 ： 【仕入先】未入力       "
                               .MB_Title = ""
                               .MB_Button = Error
                              .MBOX
                          End With
                          .Row = ID
                          Call CisVsGrid1.EditEnable
                          Call CisVsGrid1.EditVisible
                          Call VsGridToText(vsGrid1.Row)
                          Call GridTextEnabled(True)
                          Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                          Exit Function
                       End If
                       If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(2))) = "" Then
                          With CisFun
                               .MB_Lines = 4
                               .MB_MSG(2) = "    " & ID & "行目 ： 【受入】未入力       "
                               .MB_Title = ""
                               .MB_Button = Error
                              .MBOX
                          End With
                          .Row = ID
                          Call CisVsGrid1.EditEnable
                          Call CisVsGrid1.EditVisible
                          Call VsGridToText(vsGrid1.Row)
                          Call GridTextEnabled(True)
                          Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                          Exit Function
                       End If
                       If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) = "" Then
                          With CisFun
                               .MB_Lines = 4
                               .MB_MSG(2) = "    " & ID & "行目 ： 【品番】未入力       "
                               .MB_Title = ""
                               .MB_Button = Error
                              .MBOX
                          End With
                          .Row = ID
                          Call CisVsGrid1.EditEnable
                          Call CisVsGrid1.EditVisible
                          Call VsGridToText(vsGrid1.Row)
                          Call GridTextEnabled(True)
                          Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                          Exit Function
                       End If
                       
                    End If
    
                '   キー項目マスタ入力チェック
                    '取引先ﾁｪｯｸ
                    If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))) <> "" Then
                       If Not TorNmGet(Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))), 9) Then
                          With CisFun
                               .MB_Lines = 4
                               .MB_MSG(2) = "    " & ID & "行目 ： 取引先マスタ未登録       "
                               .MB_Title = ""
                               .MB_Button = Error
                              .MBOX
                          End With
                          .Row = ID
                          Call CisVsGrid1.EditEnable
                          Call CisVsGrid1.EditVisible
                          Call VsGridToText(vsGrid1.Row)
                          Call GridTextEnabled(True)
                          Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                          Exit Function
                       End If
                    End If
                    '取引先受入ﾁｪｯｸ
                    If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(2))) <> "" Then
                       gSL_Select = "SELECT * FROM 取引先受入マスタ "
                       gSL_Select = gSL_Select & " WHERE 取引先CD = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))) & "'"
                       gSL_Select = gSL_Select & " AND 受入 = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(2))) & "'"
                       gSL_Select = gSL_Select & " AND 取引先区分 = 1 "
                       If Not TUMRead(gSL_Select, 1) Then
                          With CisFun
                               .MB_Lines = 4
                               .MB_MSG(2) = "    " & ID & "行目 ： 取引先受入マスタ未登録       "
                               .MB_Title = ""
                               .MB_Button = Error
                              .MBOX
                          End With
                          .Row = ID
                          Call CisVsGrid1.EditEnable
                          Call CisVsGrid1.EditVisible
                          Call VsGridToText(vsGrid1.Row)
                          Call GridTextEnabled(True)
                          Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                          Exit Function
                       End If
                    End If
                    '品番ﾁｪｯｸ
                    If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) <> "" Then
                       gSL_Select = "SELECT * FROM 品番マスタ "
                       gSL_Select = gSL_Select & " WHERE 品番 = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) & "'"
                       If Not HNMRead(gSL_Select, 1) Then
                          With CisFun
                               .MB_Lines = 4
                               .MB_MSG(2) = "    " & ID & "行目 ： 品番マスタ未登録       "
                               .MB_Title = ""
                               .MB_Button = Error
                              .MBOX
                          End With
                          .Row = ID
                          Call CisVsGrid1.EditEnable
                          Call CisVsGrid1.EditVisible
                          Call VsGridToText(vsGrid1.Row)
                          Call GridTextEnabled(True)
                          Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                          Exit Function
                       End If
                    End If
                    '品番取引先ﾁｪｯｸ
                    If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) <> "" Then
                       gSL_Select = "SELECT * FROM 品番取引先マスタ "
                       gSL_Select = gSL_Select & " WHERE 取引先 = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))) & "'"
                       gSL_Select = gSL_Select & " AND 受入 = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(2))) & "'"
                       gSL_Select = gSL_Select & " AND 品番 = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) & "'"
                       gSL_Select = gSL_Select & " AND 品目 = 1 "
                       If Not HTMRead(gSL_Select, 1) Then
                          With CisFun
                               .MB_Lines = 4
                               .MB_MSG(2) = "    " & ID & "行目 ： 品番取引先マスタ未登録       "
                               .MB_Title = ""
                               .MB_Button = Error
                              .MBOX
                          End With
                          .Row = ID
                          Call CisVsGrid1.EditEnable
                          Call CisVsGrid1.EditVisible
                          Call VsGridToText(vsGrid1.Row)
                          Call GridTextEnabled(True)
                          Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                          Exit Function
                       End If
                    End If
                
                '   枚数項目チェック
                    If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(30))) = 2 Then       '2006/03/20  【増】の場合は便枚数０
                       .TextMatrix(ID, CisVsGrid1.FixedGet(13)) = ""
                    End If
                    
                    If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(30))) <> 0 Then
                       If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(12))) = 0 Then
                          With CisFun
                              .MB_Lines = 4
                              .MB_MSG(2) = "    " & ID & "行目 ： 【増減設定枚数】を入力して下さい            "
                              .MB_Title = ""
                              .MB_Button = Error
                              .MBOX
                          End With
                          .Row = ID
                          Call CisVsGrid1.EditEnable
                          Call CisVsGrid1.EditVisible
                          Call VsGridToText(vsGrid1.Row)
                          Call GridTextEnabled(True)
                          Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                          Exit Function
                       End If
                    End If
                    If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(30))) = 0 Then
                       If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(12))) <> 0 Then
                          With CisFun
                              .MB_Lines = 4
                              .MB_MSG(2) = "    " & ID & "行目 ： 【増減区分】を入力して下さい            "
                              .MB_Title = ""
                              .MB_Button = Error
                              .MBOX
                          End With
                          .Row = ID
                          Call CisVsGrid1.EditEnable
                          Call CisVsGrid1.EditVisible
                          Call VsGridToText(vsGrid1.Row)
                          Call GridTextEnabled(True)
                          Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                          Exit Function
                       End If
                    End If
'-------------------------------- 2006/03/20 Add
                    If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(30))) = 1 Then
                       If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(12))) < CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(13))) Then
                          With CisFun
                              .MB_Lines = 4
                              .MB_MSG(2) = "    " & ID & "行目 ： 便枚数が設定枚数をオーバーしています     "
                              .MB_Title = ""
                              .MB_Button = Error
                              .MBOX
                          End With
                          .Row = ID
                          Call CisVsGrid1.EditEnable
                          Call CisVsGrid1.EditVisible
                          Call VsGridToText(vsGrid1.Row)
                          Call GridTextEnabled(True)
                          Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                          Exit Function
                       End If
                    End If
'-------------------------------- 2006/03/20 Add End
                    
                    If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(31))) <> 0 Then
                       If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(15))) = 0 Then
                          With CisFun
                              .MB_Lines = 4
                              .MB_MSG(2) = "    " & ID & "行目 ： 【平準設定枚数】を入力して下さい            "
                              .MB_Title = ""
                              .MB_Button = Error
                              .MBOX
                          End With
                          .Row = ID
                          Call CisVsGrid1.EditEnable
                          Call CisVsGrid1.EditVisible
                          Call VsGridToText(vsGrid1.Row)
                          Call GridTextEnabled(True)
                          Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                          Exit Function
                       End If
                    End If
                    If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(31))) = 0 Then
                       If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(15))) <> 0 Then
                          With CisFun
                              .MB_Lines = 4
                              .MB_MSG(2) = "    " & ID & "行目 ： 【平準区分】を入力して下さい            "
                              .MB_Title = ""
                              .MB_Button = Error
                              .MBOX
                          End With
                          .Row = ID
                          Call CisVsGrid1.EditEnable
                          Call CisVsGrid1.EditVisible
                          Call VsGridToText(vsGrid1.Row)
                          Call GridTextEnabled(True)
                          Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                          Exit Function
                       End If
                    End If
                
'                    If RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(34))) = "" Then           '2006/03/20
                       'データ重複チェック
'                        gSL_Select = "SELECT * FROM かんばん増減テーブル "
'                        gSL_Select = gSL_Select & " WHERE 取引先 = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))) & "'"
'                        gSL_Select = gSL_Select & " AND   受入   = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(2))) & "'"
'                        gSL_Select = gSL_Select & " AND   品番   = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) & "'"
'                        gSL_Select = gSL_Select & " AND   開始日 = '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(35))) & "'"
'                        If KZTRead(gSL_Select, 1) Then
'                           With CisFun
'                               .MB_Lines = 4
'                               .MB_MSG(2) = "    " & ID & "行目 ： 【追加】既に増減データが存在します         "
'                               .MB_Title = ""
'                               .MB_Button = Error
'                               .MBOX
'                           End With
'                           .Row = ID
'                           Call CisVsGrid1.EditEnable
'                           Call CisVsGrid1.EditVisible
'                           Call VsGridToText(vsGrid1.Row)
'                           Call GridTextEnabled(True)
'                           Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
'                           Exit Function
'                        End If
                        If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(12))) <> 0 Then
                            gSL_Select = ""
                            gSL_Select = gSL_Select & "SELECT KZ.開始日"
                            gSL_Select = gSL_Select & "     , KZ.終了日"
                            gSL_Select = gSL_Select & "  FROM かんばん増減テーブル KZ"
                            gSL_Select = gSL_Select & " WHERE KZ.取引先 = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))) & "'"
                            gSL_Select = gSL_Select & "   AND KZ.受入 = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(2))) & "'"
                            gSL_Select = gSL_Select & "   AND KZ.品番 = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) & "'"
                            gSL_Select = gSL_Select & "   AND KZ.開始日 <> '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(35))) & "'"   '2006/03/31
                            gSL_Select = gSL_Select & "   AND KZ.開始日 <= '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(36))) & "'"
                            gSL_Select = gSL_Select & "   AND CASE WHEN KZ.終了日 = ''"
                            gSL_Select = gSL_Select & "            THEN '99991231'"
                            gSL_Select = gSL_Select & "            ELSE KZ.終了日"
                            gSL_Select = gSL_Select & "        END >= '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(35))) & "'"
                            If KZTRead(gSL_Select, 1) Then
                                With CisFun
                                    .MB_Lines = 5
                                    .MB_MSG(1) = "    【増減テーブル】           "
                                    .MB_MSG(2) = "    下記の登録済日付と重複しています！            "
                                    .MB_MSG(4) = "        > 開始日 : " & Format(KZT.開始日, "0000/00/00")
                                    .MB_MSG(5) = "        > 終了日 : " & Format(KZT.終了日, "0000/00/00")
                                    .MB_Title = ""
                                    .MB_Button = Error
                                    .MBOX
                                End With
                               .Row = ID
                               Call CisVsGrid1.EditEnable
                               Call CisVsGrid1.EditVisible
                               Call VsGridToText(vsGrid1.Row)
                               Call GridTextEnabled(True)
                               Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                               Exit Function
                            End If
                         End If
                                            
'                        gSL_Select = "SELECT * FROM かんばん平準化テーブル "
'                        gSL_Select = gSL_Select & " WHERE 取引先 = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))) & "'"
'                        gSL_Select = gSL_Select & " AND   受入   = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(2))) & "'"
'                        gSL_Select = gSL_Select & " AND   品番   = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) & "'"
'                        gSL_Select = gSL_Select & " AND   開始日 = '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(35))) & "'"
'                        If KZTRead(gSL_Select, 1) Then
'                           With CisFun
'                               .MB_Lines = 4
'                               .MB_MSG(2) = "    " & ID & "行目 ： 【追加】既に平準化データが存在します         "
'                               .MB_Title = ""
'                               .MB_Button = Error
'                               .MBOX
'                           End With
'                           .Row = ID
'                           Call CisVsGrid1.EditEnable
'                           Call CisVsGrid1.EditVisible
'                           Call VsGridToText(vsGrid1.Row)
'                           Call GridTextEnabled(True)
'                           Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
'                           Exit Function
'                        End If
                    
                        If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(15))) <> 0 Then
                            gSL_Select = ""
                            gSL_Select = gSL_Select & "select kjt.開始日"
                            gSL_Select = gSL_Select & "     , kjt.終了日"
                            gSL_Select = gSL_Select & "  from かんばん平準化テーブル kjt"
                            gSL_Select = gSL_Select & " where kjt.取引先 = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))) & "'"
                            gSL_Select = gSL_Select & "   and kjt.受入 = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(2))) & "'"
                            gSL_Select = gSL_Select & "   and kjt.品番 = '" & Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) & "'"
'                            gSL_Select = gSL_Select & "   and kjt.設定内容区分 = 0"            '2006/03/20
                            gSL_Select = gSL_Select & "   AND kjt.開始日 <> '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(35))) & "'"   '2006/03/31
                            gSL_Select = gSL_Select & "   and kjt.開始日 <= '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(36))) & "'"
                            gSL_Select = gSL_Select & "   and case when kjt.終了日 = ''"
                            gSL_Select = gSL_Select & "            then '99991231'"
                            gSL_Select = gSL_Select & "            else kjt.終了日"
                            gSL_Select = gSL_Select & "        end >= '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(35))) & "'"
                            If KJTRead(gSL_Select, 1) Then
                                With CisFun
                                    .MB_Lines = 5
                                    .MB_MSG(1) = "    【平準化テーブル】           "
                                    .MB_MSG(2) = "    下記の登録済日付と重複しています！            "
                                    .MB_MSG(4) = "        > 開始日 : " & Format(KJT.開始日, "0000/00/00")
                                    .MB_MSG(5) = "        > 終了日 : " & Format(KJT.終了日, "0000/00/00")
                                    .MB_Title = ""
                                    .MB_Button = Error
                                    .MBOX
                                End With
                               .Row = ID
                               Call CisVsGrid1.EditEnable
                               Call CisVsGrid1.EditVisible
                               Call VsGridToText(vsGrid1.Row)
                               Call GridTextEnabled(True)
                               Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                               Exit Function
                            End If
                        End If
'                    End If                     '2006/03/20
                                
                    mUpdate_Now = mUpdate_Now + 1
                    
                End If
             '<< 重複チェック >>
                For gCnt2 = ID + 1 To .Rows - 1
                    If RTrim(.TextMatrix(ID, 1)) = "" Then Exit For
                    If RTrim(.TextMatrix(gCnt2, 34)) = "" And _
                       RTrim(.TextMatrix(ID, 29)) = RTrim(.TextMatrix(gCnt2, 29)) And _
                       RTrim(.TextMatrix(ID, 28)) = RTrim(.TextMatrix(gCnt2, 28)) And _
                       RTrim(.TextMatrix(ID, 2)) = RTrim(.TextMatrix(gCnt2, 2)) Then
                        With CisFun
                             .MB_Lines = 4
                             .MB_MSG(1) = "    　仕入先・受入・品番が重複しています。            "
                             .MB_MSG(3) = "    " & ID & "行目と " & gCnt2 & "行目            "
                             .MB_Button = Error
                             .MBOX
                        End With
                        .Row = gCnt2
                        Call CisVsGrid1.EditEnable
                        Call CisVsGrid1.EditVisible
                        Call VsGridToText(vsGrid1.Row)
                        Call GridTextEnabled(True)
                        Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
                        Exit Function
                    End If
                Next gCnt2
               
            Next ID
        End With
    Else
        With vsGrid1
            For ID = 1 To .Rows - 1
                If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(33))) = "1" Then
                   Body1Chk = True
                   Exit Function
                End If
            Next ID
            
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(2) = "    削除行を選択して下さい          "
            CisFun.MB_Title = "削除処理"
            CisFun.MB_Button = Error
            CisFun.MBOX
            vsGrid1.SetFocus
            Exit Function
       
        End With
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
    
    If GridSet1 Then
        DBInput = True
    End If
    
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'****************************
'*      SQL文ｾｯﾄ処理
'****************************
Private Sub SqlSelectSet(Optional InputKey As String = "")
    Dim wYmT        As String
    Dim wYmY        As String
'
    With CisFun
        wYmT = iH1_Ym
        .Date_RsFormat = fYM
        .Date_Add iH1_Ym, M, 1
        wYmY = .Date_Result
    End With
'
    '名称マスタ【基本開始日】獲得
    If MeisyoGet("かんばん増減基本開始日", "1") Then
       mBaseStartDay = CisFun.Val2(SYM.英数字1)
    Else
       mBaseStartDay = 1
    End If
        
    '
    With CisDB
         .ConnectNo = 0
         .SQL = "DELETE FROM かんばん増減一括新規ワーク "
         .SQL = .SQL & " WHERE 処理端末 = '" & RTrim(CisFun.WSID) & "'"
         If Not .DBExec Then Exit Sub
    
         .SQL = "DELETE FROM かんばん増減一括更新ワーク "
         .SQL = .SQL & " WHERE 処理端末 = '" & RTrim(CisFun.WSID) & "'"
         If Not .DBExec Then Exit Sub
    End With

'
    '流動管理テーブル
    If SyoriKB = "ADD" Then
       gSL_Select = "INSERT INTO かんばん増減一括新規ワーク "
       gSL_Select = gSL_Select & "SELECT '" & RTrim(CisFun.WSID) & "',"
       gSL_Select = gSL_Select & " KR.取引先,KR.取引先名,KR.受入,KR.品番,"
       gSL_Select = gSL_Select & " CASE WHEN ISNULL(HN.表示品番,'') = '' THEN KR.品番 ELSE HN.表示品番 END 表示品番,"
       gSL_Select = gSL_Select & " HT.背番号,"
       gSL_Select = gSL_Select & " CASE WHEN ISNULL(HT.変更日,'') < '" & RTrim(iH1_Ym) & "01" & "'" & " THEN HT.収容数1 ELSE HT.収容数2 END 収容数,"
'-------------------------------------------- 2006/03/17 Delete
'       gSL_Select = gSL_Select & " ISNULL(KK.前回回転枚数,0) 前回回転枚数 , ISNULL(KK.回転枚数,0) 回転枚数, ISNULL(HNX.当月内示数,0) 当月内示数,KRK.有効枚数,"
'-------------------------------------------- 2006/03/17 Delete End
'-------------------------------------------- 2006/03/17 Insert
       gSL_Select = gSL_Select & " ISNULL(KK.前月回転枚数,0) 前回回転枚数 , ISNULL(KK.回転枚数,0) 回転枚数, ISNULL(HNX.当月内示数,0) 当月内示数,KRK.有効枚数,"
'-------------------------------------------- 2006/03/17 Insert End
       gSL_Select = gSL_Select & " 0 増減区分,0 設定枚数,0 便当枚数,0 累積枚数,'' 開始日, '' 終了日,0 増減設定内容区分,'' 完了日,"
       gSL_Select = gSL_Select & " 0 平準化区分,0 平準設定枚数,0 平準設定内容区分"
    '------------------------------------------------ 2006/03/17 Insert
          gSL_Select = gSL_Select & ",KK.回転枚数"
    '------------------------------------------------ 2006/03/17 Insert End
'       gSL_Select = gSL_Select & " INTO WK9" & Replace(RTrim(CisFun.WSID), "-", "")
       gSL_Select = gSL_Select & " FROM かんばん流動管理ビュー KR"
       gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
       gSL_Select = gSL_Select & " ON  HN.品番 = KR.品番"
       gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT"
       gSL_Select = gSL_Select & " ON  HT.取引先 = KR.取引先"
       gSL_Select = gSL_Select & " AND HT.受入 = KR.受入"
       gSL_Select = gSL_Select & " AND HT.品番 = KR.品番"
       gSL_Select = gSL_Select & " AND HT.品目 = 1"
       gSL_Select = gSL_Select & " LEFT OUTER JOIN かんばん回転枚数テーブル KK"
       gSL_Select = gSL_Select & " ON   KK.年月 = '" & RTrim(iH1_Ym) & "'"
       gSL_Select = gSL_Select & " AND KK.仕入先 = KR.取引先"
       gSL_Select = gSL_Select & " AND KK.受入 = KR.受入"
       gSL_Select = gSL_Select & " AND KK.品番 = KR.品番"
       gSL_Select = gSL_Select & " LEFT OUTER JOIN ( "
       gSL_Select = gSL_Select & "            SELECT HN.品番,HN.手配先,HN.受入,"
       gSL_Select = gSL_Select & "            SUM(CASE WHEN SUBSTRING(HN.納期, 1, 6) = '" & wYmT & "' THEN HN.数量 ELSE 0 END) 当月内示数"
       gSL_Select = gSL_Select & "            FROM 発注内示テーブル HN"
       gSL_Select = gSL_Select & "            WHERE HN.内示年月 = '" & iH1_Ym & "'"
       gSL_Select = gSL_Select & "            GROUP BY HN.品番,HN.手配先,HN.受入 ) HNX"
       gSL_Select = gSL_Select & " ON  HNX.品番 = KR.品番"
       gSL_Select = gSL_Select & " AND HNX.手配先 = KR.取引先"
       gSL_Select = gSL_Select & " AND HNX.受入 = KR.受入"
       gSL_Select = gSL_Select & " LEFT OUTER JOIN (SELECT  品番, 取引先, 受入, COUNT(品番) 有効枚数"
       gSL_Select = gSL_Select & "            FROM かんばん流動管理テーブル"
       gSL_Select = gSL_Select & "            Where 発行区分 = 1"
       gSL_Select = gSL_Select & "            GROUP BY  品番,取引先,受入) KRK"
       gSL_Select = gSL_Select & " ON  KRK.品番 = KR.品番 "
       gSL_Select = gSL_Select & " AND KRK.取引先 = KR.取引先 "
       gSL_Select = gSL_Select & " AND KRK.受入 = KR.受入"
       gSL_Select = gSL_Select & " LEFT OUTER JOIN かんばん増減テーブル KZ "
       gSL_Select = gSL_Select & " ON  KZ.取引先 = KR.取引先 "
       gSL_Select = gSL_Select & " AND KZ.受入 = KR.受入 "
       gSL_Select = gSL_Select & " AND KZ.品番 = KR.品番 "
       gSL_Select = gSL_Select & " AND SUBSTRING(KZ.開始日, 1, 6) <= '" & iH1_Ym & "'"
       gSL_Select = gSL_Select & " AND CASE WHEN KZ.終了日 = '' THEN '999912' ELSE SUBSTRING(KZ.終了日,1,6) END >= '" & iH1_Ym & "'"
       gSL_Select = gSL_Select & " LEFT OUTER JOIN かんばん平準化テーブル HJ"
       gSL_Select = gSL_Select & " ON  HJ.取引先 = KR.取引先 "
       gSL_Select = gSL_Select & " AND HJ.受入 = KR.受入 "
       gSL_Select = gSL_Select & " AND HJ.品番 = KR.品番 "
       gSL_Select = gSL_Select & " AND SUBSTRING(HJ.開始日, 1, 6) <= '" & iH1_Ym & "'"
       gSL_Select = gSL_Select & " AND CASE WHEN HJ.終了日 = '' THEN '999912' ELSE SUBSTRING(HJ.終了日,1,6) END >= '" & iH1_Ym & "'"

'-------------------------------------------- 2006/03/20 Add Start
       gSL_Select = gSL_Select & " LEFT OUTER JOIN ( SELECT 取引先,受入,品番,開始日,終了日,COUNT(品番) 品番件数 FROM かんばん増減テーブル"
       gSL_Select = gSL_Select & "                                                                     WHERE SUBSTRING(開始日, 1, 6) <= '" & iH1_Ym & "'"
       gSL_Select = gSL_Select & "                                                                       AND 設定内容区分 = 1 "
       gSL_Select = gSL_Select & "                                                                       AND  CASE WHEN 終了日 = '' THEN '999912' ELSE SUBSTRING(終了日,1,6) END >= '" & iH1_Ym & "'"
       gSL_Select = gSL_Select & "                                                          GROUP BY 取引先,受入,品番,開始日,終了日 ) KZM "
       gSL_Select = gSL_Select & " ON  KZM.取引先 = KR.取引先 "
       gSL_Select = gSL_Select & " AND KZM.受入 = KR.受入 "
       gSL_Select = gSL_Select & " AND KZM.品番 = KR.品番 "
       gSL_Select = gSL_Select & " AND SUBSTRING(KZM.開始日, 1, 6) <= '" & iH1_Ym & "'"
       gSL_Select = gSL_Select & " AND CASE WHEN KZM.終了日 = '' THEN '999912' ELSE SUBSTRING(KZM.終了日,1,6) END >= '" & iH1_Ym & "'"
       gSL_Select = gSL_Select & " LEFT OUTER JOIN ( SELECT 取引先,受入,品番,開始日,終了日,COUNT(品番) 品番件数 FROM かんばん平準化テーブル"
       gSL_Select = gSL_Select & "                                                                     WHERE SUBSTRING(開始日, 1, 6) <= '" & iH1_Ym & "'"
       gSL_Select = gSL_Select & "                                                                       AND 設定内容区分 = 1 "
       gSL_Select = gSL_Select & "                                                                       AND  CASE WHEN 終了日 = '' THEN '999912' ELSE SUBSTRING(終了日,1,6) END >= '" & iH1_Ym & "'"
       gSL_Select = gSL_Select & "                                                          GROUP BY 取引先,受入,品番,開始日,終了日 ) HJM "
       gSL_Select = gSL_Select & " ON  HJM.取引先 = KR.取引先 "
       gSL_Select = gSL_Select & " AND HJM.受入 = KR.受入 "
       gSL_Select = gSL_Select & " AND HJM.品番 = KR.品番 "
       gSL_Select = gSL_Select & " AND SUBSTRING(HJM.開始日, 1, 6) <= '" & iH1_Ym & "'"
       gSL_Select = gSL_Select & " AND CASE WHEN HJM.終了日 = '' THEN '999912' ELSE SUBSTRING(HJM.終了日,1,6) END >= '" & iH1_Ym & "'"
'-------------------------------------------- 2006/03/20 Add End

'       gSL_Select = gSL_Select & " WHERE ISNULL(KZ.品番,'') = '' "
'       gSL_Select = gSL_Select & " AND   ISNULL(HJ.品番,'') = '' "
       gSL_Select = gSL_Select & " WHERE ( ISNULL(KZ.品番,'') = '' OR ( ISNULL(KZ.品番,'') <> '' AND ISNULL(KZ.設定内容区分,0) <> 0  AND ISNULL(KZM.品番件数,0) <> 0 ) )"        '2006/03/20
       gSL_Select = gSL_Select & " AND   ( ISNULL(HJ.品番,'') = '' OR ( ISNULL(HJ.品番,'') <> '' AND ISNULL(HJ.設定内容区分,0) <> 0  AND ISNULL(HJM.品番件数,0) <> 0 ) )"        '2006/03/20
       gSL_Select = gSL_Select & " AND   ISNULL(HT.品番,'') <> '' "
       gSL_Select = gSL_Select & " AND   ISNULL(HT.手配区分,'') <> '2' "
'   取引先
       If RTrim(iH1_Torcd) <> "" Then
           gSL_Select = gSL_Select & " AND KR.取引先 = '" & RTrim(iH1_Torcd) & "'"
       End If
'   受入
       If RTrim(iH1_Ukeir) <> "" Then
           gSL_Select = gSL_Select & " AND KR.受入 = '" & RTrim(iH1_Ukeir) & "'"
       End If
'   品番
       If RTrim(iH1_Hinbn) <> "" Then
           gSL_Select = gSL_Select & " AND KR.品番 = '" & RTrim(iH1_Hinbn) & "'"
       End If
    End If
    
    'かんばん増減テーブル
    If SyoriKB = "ADD" Then
       gSL_Select = gSL_Select & " UNION  "
       gSL_Select = gSL_Select & " SELECT '" & RTrim(CisFun.WSID) & "',"
    Else
       gSL_Select = " INSERT INTO かんばん増減一括新規ワーク "
       gSL_Select = gSL_Select & "SELECT '" & RTrim(CisFun.WSID) & "',"
    End If
    gSL_Select = gSL_Select & " KZ.取引先,TR.略称 取引先名,KZ.受入,KZ.品番,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(HN.表示品番,'') = '' THEN KZ.品番 ELSE HN.表示品番 END 表示品番,"
    gSL_Select = gSL_Select & " HT.背番号,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(HT.変更日,'') < '" & RTrim(iH1_Ym) & "01" & "'" & " THEN HT.収容数1 ELSE HT.収容数2 END 収容数,"
'-------------------------------------------- 2006/03/17 Delete
'    gSL_Select = gSL_Select & " ISNULL(KK.前回回転枚数,'') 前回回転枚数 , ISNULL(KK.回転枚数,0) 回転枚数, ISNULL(HNX.当月内示数,0) 当月内示数,ISNULL(KRK.有効枚数,0) 有効枚数,"
'-------------------------------------------- 2006/03/17 Delete End
'-------------------------------------------- 2006/03/17 Insert
    gSL_Select = gSL_Select & " ISNULL(KK.前月回転枚数,'') 前回回転枚数 , ISNULL(KK.回転枚数,0) 回転枚数, ISNULL(HNX.当月内示数,0) 当月内示数,ISNULL(KRK.有効枚数,0) 有効枚数,"
'-------------------------------------------- 2006/03/17 Insert End
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(KZ.品番,'') = '' THEN 0 "
    gSL_Select = gSL_Select & " ELSE "
    gSL_Select = gSL_Select & "      CASE WHEN KZ.増減区分 = 0 THEN 1 "
    gSL_Select = gSL_Select & "      ELSE 2 END "
    gSL_Select = gSL_Select & " END 増減区分,"
    gSL_Select = gSL_Select & " KZ.設定枚数,KZ.便当枚数,KZ.累積枚数,KZ.開始日,KZ.終了日,KZ.設定内容区分 増減設定内容区分,KZ.完了日,"
    gSL_Select = gSL_Select & " 0 平準化区分,0 平準設定枚数,0 平準設定内容区分"
    '------------------------------------------------ 2006/03/17 Insert
    gSL_Select = gSL_Select & ",KK.回転枚数"
    '------------------------------------------------ 2006/03/17 Insert End
'    If SyoriKB <> "ADD" Then
'       gSL_Select = gSL_Select & " INTO WK9" & Replace(RTrim(CisFun.WSID), "-", "")
'    End If
    gSL_Select = gSL_Select & " FROM かんばん増減テーブル KZ"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & " ON  TR.取引先CD = KZ.取引先"
    gSL_Select = gSL_Select & " AND TR.取引先区分 = 1 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & " ON  HN.品番 = KZ.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " ON  HT.取引先 = KZ.取引先"
    gSL_Select = gSL_Select & " AND HT.受入 = KZ.受入"
    gSL_Select = gSL_Select & " AND HT.品番 = KZ.品番"
    gSL_Select = gSL_Select & " AND HT.品目 = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN かんばん回転枚数テーブル KK"
    gSL_Select = gSL_Select & " ON   KK.年月 = '" & RTrim(iH1_Ym) & "'"
    gSL_Select = gSL_Select & " AND KK.仕入先 = KZ.取引先"
    gSL_Select = gSL_Select & " AND KK.受入 = KZ.受入"
    gSL_Select = gSL_Select & " AND KK.品番 = KZ.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN ( "
    gSL_Select = gSL_Select & "            SELECT HN.品番,HN.手配先,HN.受入,"
    gSL_Select = gSL_Select & "            SUM(CASE WHEN SUBSTRING(HN.納期, 1, 6) = '" & wYmT & "' THEN HN.数量 ELSE 0 END) 当月内示数"
    gSL_Select = gSL_Select & "            FROM 発注内示テーブル HN"
    gSL_Select = gSL_Select & "            WHERE HN.内示年月 = '" & iH1_Ym & "'"
    gSL_Select = gSL_Select & "            GROUP BY HN.品番,HN.手配先,HN.受入 ) HNX"
    gSL_Select = gSL_Select & " ON  HNX.品番 = KZ.品番"
    gSL_Select = gSL_Select & " AND HNX.手配先 = KZ.取引先"
    gSL_Select = gSL_Select & " AND HNX.受入 = KZ.受入"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN (SELECT  品番, 取引先, 受入, COUNT(品番) 有効枚数"
    gSL_Select = gSL_Select & "            FROM かんばん流動管理テーブル"
    gSL_Select = gSL_Select & "            Where 発行区分 = 1"
    gSL_Select = gSL_Select & "            GROUP BY  品番,取引先,受入) KRK"
    gSL_Select = gSL_Select & " ON  KRK.品番 = KZ.品番 "
    gSL_Select = gSL_Select & " AND KRK.取引先 = KZ.取引先 "
    gSL_Select = gSL_Select & " AND KRK.受入 = KZ.受入"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN かんばん平準化テーブル HJ"
    gSL_Select = gSL_Select & " ON  HJ.取引先 = KZ.取引先 "
    gSL_Select = gSL_Select & " AND HJ.受入 = KZ.受入 "
    gSL_Select = gSL_Select & " AND HJ.品番 = KZ.品番 "
    gSL_Select = gSL_Select & " AND SUBSTRING(HJ.開始日, 1, 6) <= '" & iH1_Ym & "'"
    gSL_Select = gSL_Select & " AND CASE WHEN HJ.終了日 = '' THEN '999912' ELSE SUBSTRING(HJ.終了日,1,6) END >= '" & iH1_Ym & "'"
    gSL_Select = gSL_Select & " WHERE ISNULL(HT.品番,'') <> '' "
    gSL_Select = gSL_Select & " AND   ISNULL(HT.手配区分,'') <> '2' "
    gSL_Select = gSL_Select & " AND   ISNULL(KZ.設定内容区分,0) <> 1"
'   取引先
    If RTrim(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & " AND KZ.取引先 = '" & RTrim(iH1_Torcd) & "'"
    End If
'   受入
    If RTrim(iH1_Ukeir) <> "" Then
        gSL_Select = gSL_Select & " AND KZ.受入 = '" & RTrim(iH1_Ukeir) & "'"
    End If
'   品番
    If RTrim(iH1_Hinbn) <> "" Then
        gSL_Select = gSL_Select & " AND KZ.品番 = '" & RTrim(iH1_Hinbn) & "'"
    End If
    gSL_Select = gSL_Select & " AND SUBSTRING(KZ.開始日, 1, 6) <= '" & iH1_Ym & "'"
    gSL_Select = gSL_Select & " AND CASE WHEN KZ.終了日 = '' THEN '999912' ELSE SUBSTRING(KZ.終了日,1,6) END >= '" & iH1_Ym & "'"
    
    'かんばん平準化テーブル
    gSL_Select = gSL_Select & " UNION  "
    gSL_Select = gSL_Select & " SELECT '" & RTrim(CisFun.WSID) & "',"
    gSL_Select = gSL_Select & " HJ.取引先,TR.略称 取引先名,HJ.受入,HJ.品番,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(HN.表示品番,'') = '' THEN HJ.品番 ELSE HN.表示品番 END 表示品番,"
    gSL_Select = gSL_Select & " HT.背番号,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(HT.変更日,'') < '" & RTrim(iH1_Ym) & "01" & "'" & " THEN HT.収容数1 ELSE HT.収容数2 END 収容数,"
'--------------------------------------- 2006/03/17 Delete
    'gSL_Select = gSL_Select & " ISNULL(KK.前回回転枚数,0) 前回回転枚数 , ISNULL(KK.回転枚数,0) 回転枚数, ISNULL(HNX.当月内示数,0) 当月内示数,ISNULL(KRK.有効枚数,0) 有効枚数,"
'--------------------------------------- 2006/03/17 Delete End
'--------------------------------------- 2006/03/17 Insert
    gSL_Select = gSL_Select & " ISNULL(KK.前月回転枚数,0) 前回回転枚数 , ISNULL(KK.回転枚数,0) 回転枚数, ISNULL(HNX.当月内示数,0) 当月内示数,ISNULL(KRK.有効枚数,0) 有効枚数,"
'--------------------------------------- 2006/03/17 Insert End
    gSL_Select = gSL_Select & " 0 増減区分,0 設定枚数,0 便当枚数,0 累積枚数,HJ.開始日,HJ.終了日,0 増減設定内容区分,'' 完了日,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(HJ.品番,'') = '' THEN 0 ELSE HJ.平準化区分 END 平準化区分,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(HJ.平準化区分,0) = 1 THEN HJ.日上限枚数 "
    gSL_Select = gSL_Select & "      WHEN ISNULl(HJ.平準化区分,0) = 2 THEN HJ.便上限枚数 "
    gSL_Select = gSL_Select & "      WHEN ISNULL(HJ.平準化区分,0) = 3 THEN HJ.月上限枚数 "
    gSL_Select = gSL_Select & " END 平準設定枚数,HJ.設定内容区分 平準設定内容区分"
    '------------------------------------------------ 2006/03/17 Insert
    gSL_Select = gSL_Select & ",KK.回転枚数"
    '------------------------------------------------ 2006/03/17 Insert End
    gSL_Select = gSL_Select & " FROM かんばん平準化テーブル HJ"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & " ON  TR.取引先CD = HJ.取引先"
    gSL_Select = gSL_Select & " AND TR.取引先区分 = 1 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & " ON  HN.品番 = HJ.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " ON  HT.取引先 = HJ.取引先"
    gSL_Select = gSL_Select & " AND HT.受入 = HJ.受入"
    gSL_Select = gSL_Select & " AND HT.品番 = HJ.品番"
    gSL_Select = gSL_Select & " AND HT.品目 = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN かんばん回転枚数テーブル KK"
    gSL_Select = gSL_Select & " ON   KK.年月 = '" & RTrim(iH1_Ym) & "'"
    gSL_Select = gSL_Select & " AND KK.仕入先 = HJ.取引先"
    gSL_Select = gSL_Select & " AND KK.受入 = HJ.受入"
    gSL_Select = gSL_Select & " AND KK.品番 = HJ.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN ( "
    gSL_Select = gSL_Select & "            SELECT HN.品番,HN.手配先,HN.受入,"
    gSL_Select = gSL_Select & "            SUM(CASE WHEN SUBSTRING(HN.納期, 1, 6) = '" & wYmT & "' THEN HN.数量 ELSE 0 END) 当月内示数"
    gSL_Select = gSL_Select & "            FROM 発注内示テーブル HN"
    gSL_Select = gSL_Select & "            WHERE HN.内示年月 = '" & iH1_Ym & "'"
    gSL_Select = gSL_Select & "            GROUP BY HN.品番,HN.手配先,HN.受入 ) HNX"
    gSL_Select = gSL_Select & " ON  HNX.品番 = HJ.品番"
    gSL_Select = gSL_Select & " AND HNX.手配先 = HJ.取引先"
    gSL_Select = gSL_Select & " AND HNX.受入 = HJ.受入"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN (SELECT  品番, 取引先, 受入, COUNT(品番) 有効枚数"
    gSL_Select = gSL_Select & "            FROM かんばん流動管理テーブル"
    gSL_Select = gSL_Select & "            Where 発行区分 = 1"
    gSL_Select = gSL_Select & "            GROUP BY  品番,取引先,受入) KRK"
    gSL_Select = gSL_Select & " ON  KRK.品番 = HJ.品番 "
    gSL_Select = gSL_Select & " AND KRK.取引先 = HJ.取引先 "
    gSL_Select = gSL_Select & " AND KRK.受入 = HJ.受入"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN かんばん増減テーブル KZ "
    gSL_Select = gSL_Select & " ON  KZ.取引先 = HJ.取引先 "
    gSL_Select = gSL_Select & " AND KZ.受入 = HJ.受入 "
    gSL_Select = gSL_Select & " AND KZ.品番 = HJ.品番 "
    gSL_Select = gSL_Select & " AND SUBSTRING(KZ.開始日, 1, 6) <= '" & iH1_Ym & "'"
    gSL_Select = gSL_Select & " AND CASE WHEN KZ.終了日 = '' THEN '999912' ELSE SUBSTRING(KZ.終了日,1,6) END >= '" & iH1_Ym & "'"
    gSL_Select = gSL_Select & " WHERE ISNULL(HT.品番,'') <> '' "
    gSL_Select = gSL_Select & " AND   ISNULL(HT.手配区分,'') <> '2' "
    gSL_Select = gSL_Select & " AND   ISNULL(HJ.設定内容区分,0) <> 1"
'   取引先
    If RTrim(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & " AND HJ.取引先 = '" & RTrim(iH1_Torcd) & "'"
    End If
'   受入
    If RTrim(iH1_Ukeir) <> "" Then
        gSL_Select = gSL_Select & " AND HJ.受入 = '" & RTrim(iH1_Ukeir) & "'"
    End If
'   品番
    If RTrim(iH1_Hinbn) <> "" Then
        gSL_Select = gSL_Select & " AND HJ.品番 = '" & RTrim(iH1_Hinbn) & "'"
    End If
    gSL_Select = gSL_Select & " AND SUBSTRING(HJ.開始日, 1, 6) <= '" & iH1_Ym & "'"
    gSL_Select = gSL_Select & " AND CASE WHEN HJ.終了日 = '' THEN '999912' ELSE SUBSTRING(HJ.終了日,1,6) END >= '" & iH1_Ym & "'"

    With CisDB
        .ConnectNo = 0
        .SQL = gSL_Select
        If Not .DBExec Then Exit Sub
    End With
    
    gSL_Select = " INSERT INTO かんばん増減一括更新ワーク "
    gSL_Select = gSL_Select & " SELECT '" & RTrim(CisFun.WSID) & "',"
    gSL_Select = gSL_Select & " 取引先,取引先名,受入,品番,表示品番,背番号,収容数,前回回転枚数, 回転枚数, 当月内示数,"
    gSL_Select = gSL_Select & " 有効枚数,MAX(増減区分) 増減区分,MAX(設定枚数) 設定枚数,MAX(便当枚数) 便当枚数,MAX(累積枚数) 累積枚数,開始日,終了日,"
    gSL_Select = gSL_Select & " MAX(増減設定内容区分) 増減設定内容区分,MAX(完了日) 完了日,MAX(平準化区分) 平準化区分,MAX(平準設定枚数) 平準設定枚数,MAX(平準設定内容区分) 平準設定内容区分"
    '------------------------------------------------ 2006/03/17 Insert
    gSL_Select = gSL_Select & ",MAX(今回回転枚数) "
    '------------------------------------------------ 2006/03/17 Insert End
'    gSL_Select = gSL_Select & " INTO WK8" & Replace(RTrim(CisFun.WSID), "-", "")
'    gSL_Select = gSL_Select & " FROM WK9" & Replace(RTrim(CisFun.WSID), "-", "")
    gSL_Select = gSL_Select & " FROM かんばん増減一括新規ワーク "
    gSL_Select = gSL_Select & " WHERE ISNULl(品番,'') <> '' "
    gSL_Select = gSL_Select & " AND   ISNULL(処理端末,'') = '" & RTrim(CisFun.WSID) & "'"
    gSL_Select = gSL_Select & " GROUP BY"
    gSL_Select = gSL_Select & " 取引先,取引先名,受入,品番,表示品番,背番号,収容数,前回回転枚数, 回転枚数, 当月内示数,"
    gSL_Select = gSL_Select & " 有効枚数 , 開始日, 終了日"
    gSL_Select = gSL_Select & " ORDER BY 取引先,受入,背番号,品番,ISNULL(開始日,'') "
    With CisDB
        .ConnectNo = 0
        .SQL = gSL_Select
        If Not .DBExec Then Exit Sub
    End With

'    gSL_Select = "SELECT * FROM WK8" & Replace(RTrim(CisFun.WSID), "-", "")
    gSL_Select = "SELECT * FROM かんばん増減一括更新ワーク "
    gSL_Select = gSL_Select & " WHERE 品番 <> '' "
    gSL_Select = gSL_Select & " AND   ISNULL(処理端末,'') = '" & RTrim(CisFun.WSID) & "'"
'   増減区分
    If H1Chk_ZouGen(0).Value = 1 And H1Chk_ZouGen(1).Value = 0 Then gSL_Select = gSL_Select & " AND ISNULL(増減区分,0) = 2"
    If H1Chk_ZouGen(1).Value = 1 And H1Chk_ZouGen(0).Value = 0 Then gSL_Select = gSL_Select & " AND ISNULL(増減区分,0) = 1"
'    If H1Chk_ZouGen(0).Value = 1 And H1Chk_ZouGen(1).Value = 1 Then gSL_Select = gSL_Select & " AND ( ISNULL(増減区分,0) = 0 OR ISNULL(増減区分,0) = 1 ) "
    
'   平準区分
    '便のみ
    If H1Chk_Heijyun(0).Value = 1 And H1Chk_Heijyun(1).Value = 0 And H1Chk_Heijyun(2).Value = 0 Then
       gSL_Select = gSL_Select & " AND ISNULL(平準化区分,0) = 2"
    End If
    '日のみ
    If H1Chk_Heijyun(0).Value = 0 And H1Chk_Heijyun(1).Value = 1 And H1Chk_Heijyun(2).Value = 0 Then
       gSL_Select = gSL_Select & " AND ISNULL(平準化区分,0) = 1"
    End If
    '月のみ
    If H1Chk_Heijyun(0).Value = 0 And H1Chk_Heijyun(1).Value = 0 And H1Chk_Heijyun(2).Value = 1 Then
       gSL_Select = gSL_Select & " AND ISNULL(平準化区分,0) = 3"
    End If
    '便・日のみ
    If H1Chk_Heijyun(0).Value = 1 And H1Chk_Heijyun(1).Value = 1 And H1Chk_Heijyun(2).Value = 0 Then
       gSL_Select = gSL_Select & " AND ( ISNULL(平準化区分,0) = 2 OR ISNULL(平準化区分,0) = 1 )"
    End If
    '便・月のみ
    If H1Chk_Heijyun(0).Value = 1 And H1Chk_Heijyun(1).Value = 0 And H1Chk_Heijyun(2).Value = 1 Then
       gSL_Select = gSL_Select & " AND ( ISNULL(平準化区分,0) = 2 OR ISNULL(平準化区分,0) = 3 )"
    End If
    '日・月のみ
    If H1Chk_Heijyun(0).Value = 0 And H1Chk_Heijyun(1).Value = 1 And H1Chk_Heijyun(2).Value = 1 Then
       gSL_Select = gSL_Select & " AND ( ISNULL(平準化区分,0) = 1 OR ISNULL(平準化区分,0) = 3 )"
    End If
    '全て
    If H1Chk_Heijyun(0).Value = 1 And H1Chk_Heijyun(1).Value = 1 And H1Chk_Heijyun(2).Value = 1 Then
       gSL_Select = gSL_Select & " AND ( ISNULL(平準化区分,0) = 1 OR ISNULL(平準化区分,0) = 2 OR ISNULL(平準化区分,0) = 3 )"
    End If
    
    gSL_Select = gSL_Select & " ORDER BY 取引先,受入,背番号,品番,開始日,終了日"

'
End Sub
'*------------------------------------------------------------------------------*
'*      退避データと入力テキストとのチェック ( 異なっていれば更新をうながす )   *
'*------------------------------------------------------------------------------*
Private Function SaveDateCheck()
    SaveDateCheck = False
    
    Static swTopR           As Integer
    Static swRow            As Integer

    swTopR = vsGrid1.TopRow
    swRow = vsGrid1.Row
    
    Call iB1_Torcd_LostFocus
    Call iB1_Uke_LostFocus
    Call iB1_ZBinSu_LostFocus
    Call iB1_ZSetSu_LostFocus
    Call iB1_ZKbn_LostFocus
    Call iB1_Hinbn_LostFocus
    Call iB1_Seban_LostFocus
    Call iB1_HKbn_LostFocus
    Call iB1_HSetSu_LostFocus
    
    If mUpdate_Cnt <> 0 Then
        If Not Body1Chk Then Exit Function
        Call DBPut(1)
        Input_Row.Visible = False
        vsGrid1.Redraw = False
        Call HeadBodyClear("B")
    '    Call Head1Chk(1)
        Call Head1Chk
        Input_Row.Visible = True
        vsGrid1.TopRow = swTopR
        vsGrid1.Redraw = True
        
        If vsGrid1.Rows <= swRow Then swRow = vsGrid1.Rows - 1
        vsGrid1.Row = swRow
        Call VsGridToText(vsGrid1.Row)
    End If
'    iB1_Torcd.SetFocus
    
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
        If Not .MBOX Then
           vsGrid1.SetFocus
           Exit Function
        End If
    End With
    
    For gInt = 1 To vsGrid1.Rows - 1
        If RTrim(vsGrid1.TextMatrix(gInt, CisVsGrid1.FixedGet(33))) = "1" Then
            With CisDB
                On Error GoTo DBDelete_Err
                .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
        ' ( 増減テーブル削除処理 )
                .SQL = "DELETE FROM かんばん増減テーブル "
                .SQL = .SQL & " WHERE 取引先 = '" & RTrim(vsGrid1.TextMatrix(gInt, CisVsGrid1.FixedGet(29))) & "'"
                .SQL = .SQL & " AND   受入 = '" & RTrim(vsGrid1.TextMatrix(gInt, CisVsGrid1.FixedGet(2))) & "'"
                .SQL = .SQL & " AND   品番 = '" & RTrim(vsGrid1.TextMatrix(gInt, CisVsGrid1.FixedGet(28))) & "'"
                .SQL = .SQL & " AND   開始日 = '" & RTrim(vsGrid1.TextMatrix(gInt, CisVsGrid1.FixedGet(35))) & "'"
                .DBExec
                
        ' ( 平準化テーブル削除処理 )
                .SQL = "DELETE FROM かんばん平準化テーブル "
                .SQL = .SQL & " WHERE 取引先 = '" & RTrim(vsGrid1.TextMatrix(gInt, CisVsGrid1.FixedGet(29))) & "'"
                .SQL = .SQL & " AND   受入 = '" & RTrim(vsGrid1.TextMatrix(gInt, CisVsGrid1.FixedGet(2))) & "'"
                .SQL = .SQL & " AND   品番 = '" & RTrim(vsGrid1.TextMatrix(gInt, CisVsGrid1.FixedGet(28))) & "'"
                .SQL = .SQL & " AND   開始日 = '" & RTrim(vsGrid1.TextMatrix(gInt, CisVsGrid1.FixedGet(35))) & "'"
                .DBExec
                
                .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
                On Error GoTo 0
            End With
        End If
    Next gInt
    
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

        
    If pKB = 0 Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "    更新処理を行います。            "
            .MB_MSG(4) = "     よろしいですか？            "
            .MB_Title = "更新処理"
            .MB_Button = OK_CAN
            If Not .MBOX Then
               Call VsGridToText(Val(Input_Row.Tag))
               Call GridTextEnabled(True)
               Exit Function
            End If
        End With
    End If
    
    On Error GoTo DBPut_Err

    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 更新処理 )
    Dim ID          As Long
    mUpdate_Now = 0
    With vsGrid1
        For ID = 1 To .Rows - 1
            If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(0))) = "" Then Exit For
            If mUpdate_Now = mUpdate_Cnt Then Exit For                                  '変更対象件数まで処理したら終了
            If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(40))) <> "" Then                '変更対象のみ更新
        
        ' ( 増減テーブル削除処理 )
                CisDB.SQL = "DELETE FROM かんばん増減テーブル "
                CisDB.SQL = CisDB.SQL & " WHERE 取引先 = '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))) & "'"
                CisDB.SQL = CisDB.SQL & " AND   受入 = '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(2))) & "'"
                CisDB.SQL = CisDB.SQL & " AND   品番 = '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) & "'"
                CisDB.SQL = CisDB.SQL & " AND   開始日 = '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(35))) & "'"
                CisDB.DBExec

                If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(12))) <> 0 Or _
                   CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(13))) <> 0 Then
                      Call DBINsert_KZT(ID)           'かんばん増減テーブル追加
                End If

        ' ( 平準化テーブル削除処理 )
                CisDB.SQL = "DELETE FROM かんばん平準化テーブル "
                CisDB.SQL = CisDB.SQL & " WHERE 取引先 = '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(29))) & "'"
                CisDB.SQL = CisDB.SQL & " AND   受入 = '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(2))) & "'"
                CisDB.SQL = CisDB.SQL & " AND   品番 = '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(28))) & "'"
                CisDB.SQL = CisDB.SQL & " AND   開始日 = '" & RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(35))) & "'"
                CisDB.DBExec
                
                If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(15))) <> 0 Then
                   Call DBINsert_KJT(ID)           'かんばん平準テーブル追加
                End If
            
                mUpdate_Now = mUpdate_Now + 1
                
            End If
        
        Next ID
    End With

    CisDB.DBTran (TransCommit)      ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+------------------------------------+
'+      かんばん増減テーブル更新
'+------------------------------------+
Public Function DBINsert_KZT(SetRow As Long, Optional DBNo As Integer) As Boolean
    DBINsert_KZT = False
    
    '開始日獲得
    Call SYMD_Get(RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(29))), RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(2))), RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(28))))
    
    Call ItemsClearKZT
    With KZT
        
        .取引先 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(29)))
        .受入 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(2)))
        .品番 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(28)))
        If RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(35))) <> "" Then
           .開始日 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(35)))
        Else
           .開始日 = mStartYMD
        End If
        If RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(36))) <> "" Then
           .終了日 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(36)))
        Else
           .終了日 = mEndYMD
        End If
        .設定内容区分 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(37)))
        If CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(30))) = 1 Then
           .増減区分 = 0
        Else
           .増減区分 = 1
        End If
        .設定枚数 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(12)))
        If .増減区分 = 0 Then
           .便当枚数 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(13)))
        Else
           .便当枚数 = 0
        End If
        
        '区分が変更されたら累積枚数・完了日をクリアー
        If CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(30))) = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(23))) Then
           .累積枚数 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(22)))
           .完了日 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(38)))
           If .設定枚数 > .累積枚数 Then
               .完了日 = ""
           End If
        Else
           .累積枚数 = 0
           .完了日 = ""
        End If
        .作成者 = gTanto

    End With
    
    Call KZTInsert
    
    DBINsert_KZT = True
End Function
'+------------------------------------+
'+      かんばん平準テーブル更新
'+------------------------------------+
Public Function DBINsert_KJT(SetRow As Long, Optional DBNo As Integer) As Boolean
    DBINsert_KJT = False
    
    '開始日獲得
    Call SYMD_Get(RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(29))), RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(2))), RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(28))))
    
    Call ItemsClearKJT
    With KJT
        
        .取引先 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(29)))
        .受入 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(2)))
        .品番 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(28)))
        If RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(35))) <> "" Then
           .開始日 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(35)))
        Else
           .開始日 = mStartYMD
        End If
        If RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(36))) <> "" Then
           .終了日 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(36)))
        Else
           .終了日 = mEndYMD
        End If
        .設定内容区分 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(39)))
        .平準化区分 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(31)))
        If .平準化区分 = 1 Then .日上限枚数 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(15)))
        If .平準化区分 = 2 Then .便上限枚数 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(15)))
        If .平準化区分 = 3 Then .月上限枚数 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(15)))
        .作成者 = gTanto

    End With
    
    Call KJTInsert
    
    DBINsert_KJT = True
End Function
Private Sub SYMD_Get(sTorcd As String, sUke As String, sHinbn As String)

    '開始日の検索、存在しなければ基本開始日をセット
    gSL_Select = "SELECT 1 区分,* FROM かんばん増減開始日マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(sTorcd) & "'"
    gSL_Select = gSL_Select & " AND   受入 = ''"
    gSL_Select = gSL_Select & " AND   品番 = ''"
    gSL_Select = gSL_Select & " UNION "
    gSL_Select = gSL_Select & " SELECT 2 区分,* FROM かんばん増減開始日マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(sTorcd) & "'"
    gSL_Select = gSL_Select & " AND   受入 = '" & RTrim(sUke) & "'"
    gSL_Select = gSL_Select & " AND   品番 = ''"
    gSL_Select = gSL_Select & " UNION "
    gSL_Select = gSL_Select & " SELECT 3 区分,* FROM かんばん増減開始日マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(sTorcd) & "'"
    gSL_Select = gSL_Select & " AND   受入 = '" & RTrim(sUke) & "'"
    gSL_Select = gSL_Select & " AND   品番 = '" & RTrim(sHinbn) & "'"
    gSL_Select = gSL_Select & " ORDER BY 区分 DESC"
    If KZSRead(gSL_Select, 1, 1) Then
       mStartDay = CisFun.Val2(KZS.開始日)
    Else
       mStartDay = mBaseStartDay
    End If
    
'    If mStartDay <> 1 Then                         '2007/07/10修正
    If mStartDay <> 0 Then
       mStartYMD = RTrim(iH1_Ym) & "01"
       CisFun.Date_RsFormat = fYMD
       CisFun.Date_Add mStartYMD, D, mStartDay
       mStartYMD = CisFun.Date_Result
    Else
       mStartYMD = RTrim(iH1_Ym) & "01"
    End If
    
'    If mStartDay <> 1 Then                         '2007/07/10修正
    If mStartDay <> 0 Then
       CisFun.Date_Add mStartYMD, M, 1
       mEndYMD = CisFun.Date_Result
    Else
       CisFun.Date_RsFormat = fYMD
       CisFun.Date_End RTrim(iH1_Ym)
       mEndYMD = CisFun.Date_Result
    End If

End Sub


'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ - 流動日付
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 4, 0, "№")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/10 UPDATE START
        'Call .FixedSet(中中, 左中, あり, 19, 1, "仕 入 先")
        Call .FixedSet(中中, 左中, あり, 17, 1, "仕 入 先")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/10 UPDATE END
        Call .FixedSet(中中, 左中, あり, 2, 2, "受")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/10 UPDATE START
        'Call .FixedSet(中中, 左中, あり, 6, 3, "背番号")
        Call .FixedSet(中中, 左中, あり, 8, 3, "背番号")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/10 UPDATE END
        Call .FixedSet(中中, 左中, あり, 25, 4, "品　 番")
        Call .FixedSet(中中, 右中, あり, 6, 5, "収容数")
        Call .FixedSet(中中, 右中, あり, 6, 6, "内示数")
        Call .FixedSet(中中, 左中, あり, 12, 7, "設定期間")
        Call .FixedSet(中中, 右中, あり, 4, 8, "前回")
        Call .FixedSet(中中, 右中, あり, 4, 9, "今回")
        Call .FixedSet(中中, 右中, あり, 3, 10, "有効")
        Call .FixedSet(中中, 中中, あり, 4, 11, "増減")
        Call .FixedSet(中中, 右中, あり, 4, 12, "設定")
        Call .FixedSet(中中, 右中, あり, 4, 13, "便当")
        Call .FixedSet(中中, 中中, あり, 4, 14, "平準")
        Call .FixedSet(中中, 右中, あり, 4, 15, "平設定")
        Call .FixedSet(中中, 左中, あり, 0, 16, "X元仕入先")
        Call .FixedSet(中中, 左中, あり, 0, 17, "X元品番")
        Call .FixedSet(中中, 左中, あり, 0, 18, "X元開始日")
        Call .FixedSet(中中, 左中, あり, 0, 19, "X元終了日")
        Call .FixedSet(中中, 右中, あり, 0, 20, "X元設定枚数")
        Call .FixedSet(中中, 右中, あり, 0, 21, "X元便当枚数")
        Call .FixedSet(中中, 右中, あり, 0, 22, "X元累積枚数")
        Call .FixedSet(中中, 右中, あり, 0, 23, "X元増減区分")
        Call .FixedSet(中中, 右中, あり, 0, 24, "X元平準化区分")
        Call .FixedSet(中中, 右中, あり, 0, 25, "X元平準設定枚数")
        Call .FixedSet(中中, 左中, あり, 0, 26, "X元受入")
        Call .FixedSet(中中, 左中, あり, 0, 27, "X元背番号")
        
        Call .FixedSet(中中, 左中, あり, 0, 28, "X品番")
        Call .FixedSet(中中, 左中, あり, 0, 29, "X仕入先")
        Call .FixedSet(中中, 左中, あり, 0, 30, "X増減区分")
        Call .FixedSet(中中, 左中, あり, 0, 31, "X平準化区分")
        
        Call .FixedSet(中中, 左中, なし, 0, 32, "NotDisp")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 33, "Select")
        Call .FixedSet(中中, 左中, なし, 0, 34, "Add")
        
        Call .FixedSet(中中, 左中, あり, 0, 35, "X開始日")
        Call .FixedSet(中中, 左中, あり, 0, 36, "X終了日")
        Call .FixedSet(中中, 左中, あり, 0, 37, "Xかんばん設定内容")
        Call .FixedSet(中中, 左中, あり, 0, 38, "X完了日")
        Call .FixedSet(中中, 左中, あり, 0, 39, "X平準設定内容")

        Call .FixedSet(中中, 左中, あり, 0, 40, "Update")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15")
'   編集内容ｾｯﾄ
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO,非表示ｶﾗﾑ,編集ﾌｫｰﾏｯﾄ,日付編集)
'        Call .EditSet(iB1_ZSetSu, 10)
'        Call .EditSet(iB1_ZBinSu, 11, 18)
        .InitDisp
    
    End With
End Sub
'***********************************************'
'*****      グリッド内容セット - 流動日付
'***********************************************'
Private Function GridSet1() As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

    GridSet1 = False

    If KRVRead(gSL_Select) Then

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
        End With
        ID = 0
        Do Until Not KRV_RDSTS
            ID = ID + 1
            
'            If ID = 132 Then
'               MsgBox "AAA"
'            End If

            Call KrvInfoToGrid1(ID)

            Call KRVReadNext
        Loop
        With vsGrid1
             .Row = 1: .Col = 1: .ColSel = .Cols - 1
             .Redraw = True
        End With
        NowRow = vsGrid1.Row

        GridSet1 = True
    End If
    Call KRVClose

    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With

End Function
'+----------------------------------------------+
'+  かんばん増減テーブルよりグリッドに転送
'+----------------------------------------------+
Private Sub KrvInfoToGrid1(SetRow As Long)
    
    '開始日の検索、存在しなければ基本開始日をセット
    gSL_Select = "SELECT 1 区分,* FROM かんばん増減開始日マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(KRV.取引先) & "'"
    gSL_Select = gSL_Select & " AND   受入 = ''"
    gSL_Select = gSL_Select & " AND   品番 = ''"
    gSL_Select = gSL_Select & " UNION "
    gSL_Select = gSL_Select & " SELECT 2 区分,* FROM かんばん増減開始日マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(KRV.取引先) & "'"
    gSL_Select = gSL_Select & " AND   受入 = '" & RTrim(KRV.受入) & "'"
    gSL_Select = gSL_Select & " AND   品番 = ''"
    gSL_Select = gSL_Select & " UNION "
    gSL_Select = gSL_Select & " SELECT 3 区分,* FROM かんばん増減開始日マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(KRV.取引先) & "'"
    gSL_Select = gSL_Select & " AND   受入 = '" & RTrim(KRV.受入) & "'"
    gSL_Select = gSL_Select & " AND   品番 = '" & RTrim(KRV.品番) & "'"
    gSL_Select = gSL_Select & " ORDER BY 区分 DESC"
    If KZSRead(gSL_Select, 1, 1) Then
       mStartDay = CisFun.Val2(KZS.開始日)
    Else
       mStartDay = mBaseStartDay
    End If
    
'    If mStartDay <> 1 Then                         '2007/07/10修正
    If mStartDay <> 0 Then
       mStartYMD = RTrim(iH1_Ym) & "01"
       CisFun.Date_RsFormat = fYMD
       CisFun.Date_Add mStartYMD, D, mStartDay
       mStartYMD = CisFun.Date_Result
    Else
       mStartYMD = RTrim(iH1_Ym) & "01"
    End If
'------------------------------------------------ 2006/03/16 Delete
'    If mStartDay <> 1 Then
'       CisFun.Date_Add mStartYMD, M, 1
'       mEndYMD = CisFun.Date_Result
'    Else
'       CisFun.Date_RsFormat = fYMD
'       CisFun.Date_End RTrim(iH1_Ym)
'       mEndYMD = CisFun.Date_Result
'    End If
'------------------------------------------------ 2006/03/16 Delete End
    
'------------------------------------------------ 2006/03/16 Insert
    With CisFun
        .Date_RsFormat = fYMD
        .Date_End RTrim(iH1_Ym) & "01"
        mEndYMD = .Date_Result
    End With
'    If mStartDay <> 1 Then                         '2007/07/10修正
    If mStartDay <> 0 Then
       With CisFun
           .Date_RsFormat = fYMD
           .Date_Add mEndYMD, D, mStartDay
           mEndYMD = .Date_Result
       End With
    End If
'------------------------------------------------ 2006/03/16 Insert End
    
    With vsGrid1
        'No
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(0)) = SetRow
        '仕入先
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(1)) = RTrim(KRV.取引先) & "：" & RTrim(KRV.取引先名)
        '受入
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = RTrim(KRV.受入)
        '背番号
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(3)) = RTrim(KRV.背番号)
        '品番
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(4)) = RTrim(KRV.表示品番)
        '収容数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(5)) = Format(KRV.収容数, "##,###")
        '当月内示数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(6)) = Format(KRV.当月内示数, "##,###")
        '設定期間
        If RTrim(KRV.開始日) = "" And RTrim(KRV.終了日) = "" Then
           KRV.開始日 = mStartYMD
           KRV.終了日 = mEndYMD
        End If
        If Trim(KRV.終了日) = "" Then
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(7)) = CisFun.Mid2(Format(KRV.開始日, "0000/00/00"), 6, 5) & Space(7)
        Else
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(7)) = CisFun.Mid2(Format(KRV.開始日, "0000/00/00"), 6, 5) _
                                                 & "～" & CisFun.Mid2(Format(KRV.終了日, "0000/00/00"), 6, 5)
        End If
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(35)) = KRV.開始日
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(36)) = KRV.終了日
        
        If Trim(KRV.終了日) = "" Then
            gStr = CisFun.Mid2(Format(KRV.開始日, "0000/00/00"), 6, 5) & Space(7)
        Else
            gStr = CisFun.Mid2(Format(KRV.開始日, "0000/00/00"), 6, 5) _
                   & "～" & CisFun.Mid2(Format(KRV.終了日, "0000/00/00"), 6, 5)
        End If
        '前回回転枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(8)) = Format(KRV.前回回転枚数, "#,###")
        '今回回転枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(9)) = Format(KRV.今回回転枚数, "#,###")
        '有効枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(10)) = Format(KRV.有効枚数, "#,###")
        '増減区分
        Select Case KRV.増減区分
               Case 1: .TextMatrix(SetRow, CisVsGrid1.FixedGet(11)) = "減"
               Case 2: .TextMatrix(SetRow, CisVsGrid1.FixedGet(11)) = "増"
               Case Else: .TextMatrix(SetRow, CisVsGrid1.FixedGet(11)) = ""
        End Select
        '設定枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(12)) = Format(KRV.設定枚数, "###")
        '便当枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(13)) = Format(KRV.便当枚数, "###")
        '平準化区分
        Select Case KRV.平準化区分
               Case 1: .TextMatrix(SetRow, CisVsGrid1.FixedGet(14)) = "日"
               Case 2: .TextMatrix(SetRow, CisVsGrid1.FixedGet(14)) = "便"
               Case 3: .TextMatrix(SetRow, CisVsGrid1.FixedGet(14)) = "月"
               Case Else: .TextMatrix(SetRow, CisVsGrid1.FixedGet(14)) = ""
        End Select
        '平準設定枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(15)) = Format(KRV.平準設定枚数, "###")
        '退避
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(16)) = KRV.取引先
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(17)) = KRV.品番
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(18)) = KRV.開始日
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(19)) = KRV.終了日
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(20)) = KRV.設定枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(21)) = KRV.便当枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(22)) = KRV.累積枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(23)) = KRV.増減区分
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(24)) = KRV.平準化区分
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(25)) = KRV.平準設定枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(26)) = KRV.受入
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(27)) = KRV.背番号
        
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(29)) = KRV.取引先
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(28)) = KRV.品番
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(30)) = KRV.増減区分
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(31)) = KRV.平準化区分
        
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(37)) = KRV.増減設定内容区分
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(38)) = KRV.完了日
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(39)) = KRV.平準設定内容区分
        
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(34)) = "X"
        '表示制御
'        Select Case OpH1_ZGKbn(0)
'            Case True:  .TextMatrix(SetRow, CisVsGrid1.FixedGet(18)) = ""
'            Case False: .TextMatrix(SetRow, CisVsGrid1.FixedGet(18)) = "NotDisp"
'        End Select
    End With
End Sub
Private Sub HEAD_No_Click()
    If ProcHB <> "B1" Then Exit Sub
    If SyoriKB <> "DEL" Then Exit Sub
    vsGrid1.Row = 1: vsGrid1.RowSel = vsGrid1.Rows - 1
    Call vsGrid1_Click
End Sub


'+-------------------------------+
'+      VSGrid1_Click
'+-------------------------------+
Private Sub vsGrid1_Click()
    If ProcHB = "B1" Then

        If Trim(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(0))) = "" Then
            vsGrid1.Row = NowRow
        End If

'        Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
            
        If SyoriKB <> "DEL" Then Call TextToVsGrid(Val(Input_Row.Tag))
        If RTrim$(vsGrid1.TextMatrix(vsGrid1.Row, 0)) = "" Then Exit Sub
        If SyoriKB <> "DEL" Then Call GridTextEnabled(True)
        If SyoriKB <> "DEL" Then Call VsGridToText(vsGrid1.Row)
        
        
        Call CisVsGrid1.EditEnable
        Call CisVsGrid1.EditVisible
        
        If Trim(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(0))) <> "" Then
            NowRow = vsGrid1.Row
        End If
    
        If SyoriKB = "DEL" Then
            With vsGrid1
                .Redraw = False
                gCnt1 = .Row: gCnt2 = .RowSel
                If .RowSel < .Row Then
                   gCnt2 = .Row: gCnt1 = .RowSel
                End If
                For gLong = gCnt1 To gCnt2
                    
                    If Trim(.TextMatrix(gLong, CisVsGrid1.NoSpaceCol)) <> "" Then
                        If .TextMatrix(gLong, CisVsGrid1.FixedGet(33)) = "1" Then
                           .TextMatrix(gLong, CisVsGrid1.FixedGet(33)) = "0"
                           .Row = gLong
                           .Col = 0
                           .CellBackColor = .BackColorFixed
                           .Col = 1: vsGrid1.ColSel = 6
                           .CellBackColor = .BackColor
        '                      mClick_Cnt = mClick_Cnt - 1
        '                      B1lb_SKensu = Format(mClick_Cnt, "###,###")
                        Else
                           .TextMatrix(gLong, CisVsGrid1.FixedGet(33)) = "1"
                           .Row = gLong
                           .Col = 0
                           .CellBackColor = Color_Del.BackColor
                           .Col = 1: .ColSel = 6
                           .CellBackColor = Color_Del.BackColor
        '                      mClick_Cnt = mClick_Cnt + 1
        '                      B1lb_SKensu = Format(mClick_Cnt, "###,###")
                        End If
                    End If
                Next gLong
                
                .Redraw = True
            End With
        End If
    End If
End Sub
'Private Sub VSGrid1_Click()
'    Call TextToVsGrid(Val(Input_Row.Tag))
'    If RTrim$(vsGrid1.TextMatrix(vsGrid1.Row, 0)) = "" Then Exit Sub
'    Call GridTextEnabled(True)
'    Call VsGridToText(vsGrid1.Row)
'    On Error Resume Next
'    If ActObj.Visible Then ActObj.SetFocus
'    On Error GoTo 0
'End Sub

'+---------------------------------------------------------+
'+          グリッド編集エリア
'+---------------------------------------------------------+
'+--------------------------------------+
'+      ボディ入力項目キーチェック
'+--------------------------------------+
Private Function BD_InKeyCheck(TObj As Object, KeyCode As Integer, Shift As Integer)

'< PageUp・PageDown >
    If KeyCode = vbKeyPageDown Or _
       KeyCode = vbKeyPageUp Then
        
        If vsGrid1.Rows <= CisVsGrid1.DispMax Then Exit Function
        
'        Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
        
        If KeyCode = vbKeyPageDown Then
            
            '< PageDown >
            If vsGrid1.TopRow + 25 <= vsGrid1.Rows - 1 Then
                vsGrid1.TopRow = vsGrid1.TopRow + 25
                
                If vsGrid1.TopRow + 24 <= vsGrid1.Rows - 1 Then
                    vsGrid1.Row = vsGrid1.TopRow + 24
                Else
                    vsGrid1.Row = vsGrid1.Rows - 1
                End If
            Else
                vsGrid1.TopRow = vsGrid1.Rows - 1 - 24
            End If
        Else
            
            '< PageUp >
            If vsGrid1.TopRow - 25 > 0 Then
                vsGrid1.TopRow = vsGrid1.TopRow - 25
            Else
                vsGrid1.TopRow = 1
            End If
            vsGrid1.Row = vsGrid1.TopRow
        End If
        
        vsGrid1.SetFocus
        Call CisVsGrid1.EditEnable
        Call CisVsGrid1.EditVisible
        NowRow = vsGrid1.Row
        TObj.SetFocus
        Exit Function
    End If
    
'< ↑・↓ >
    If KeyCode = vbKeyUp Or _
       KeyCode = vbKeyDown Then
'        Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
        vsGrid1.SetFocus
        Call CisVsGrid1.KeyEdit(TObj, KeyCode, Shift)
        NowRow = vsGrid1.Row
    End If

End Function
'+-------------------------------+
'+      背景色の再設定
'+-------------------------------+
Private Sub ChgColorGrid1(SetRow As Long)
    With vsGrid1
        .Redraw = False
        NewRow = .Row
        .Row = SetRow
        If RTrim(iB1_Torcd) <> RTrim(.TextMatrix(SetRow, CisVsGrid1.FixedGet(16))) Or _
           RTrim(iB1_Uke) <> RTrim(.TextMatrix(SetRow, CisVsGrid1.FixedGet(26))) Or _
           RTrim(iB1_Seban) <> RTrim(.TextMatrix(SetRow, CisVsGrid1.FixedGet(27))) Or _
           RTrim(iB1_Hinbn) <> RTrim(.TextMatrix(SetRow, CisVsGrid1.FixedGet(17))) Or _
           CisFun.Val2(iB1_ZKbn) <> CisFun.Val2(.TextMatrix(SetRow, CisVsGrid1.FixedGet(23))) Or _
           CisFun.Val2(iB1_ZSetSu) <> CisFun.Val2(.TextMatrix(SetRow, CisVsGrid1.FixedGet(20))) Or _
           CisFun.Val2(iB1_ZBinSu) <> CisFun.Val2(.TextMatrix(SetRow, CisVsGrid1.FixedGet(21))) Or _
           CisFun.Val2(iB1_HKbn) <> CisFun.Val2(.TextMatrix(SetRow, CisVsGrid1.FixedGet(24))) Or _
           CisFun.Val2(iB1_HSetSu) <> CisFun.Val2(.TextMatrix(SetRow, CisVsGrid1.FixedGet(25))) Then
            For gInt = 1 To 11
                .Col = gInt
                .CellBackColor = mColorChg
            Next gInt
        Else
            For gInt = 1 To 11
                .Col = gInt
                .CellBackColor = mColorOrg
            Next gInt
        End If
        .Row = NewRow
        .Redraw = True
    End With
End Sub
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    If SyoriKB <> "DEL" Then Input_Row.Visible = EnabledFlg
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'テキストからグリッドに退避
    Call TextToVsGrid(Val(Input_Row.Tag))
    gInt = vsGrid1.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > vsGrid1.Rows - 1 Then
            If RTrim$(vsGrid1.TextMatrix(gInt, 1)) <> "" Then
                vsGrid1.Rows = vsGrid1.Rows + 1
                vsGrid1.TextMatrix(gInt + 1, 0) = gInt + 1
            Else
                GoTo Grid_RowChange_End
            End If
        End If
        gInt = gInt + 1
    End If
    'PageUp
    If KeyCode = vbKeyPageUp Then
        If gInt - CisVsGrid1.DispMax < 1 Then
            gInt = 1
        Else
            gInt = gInt - CisVsGrid1.DispMax
        End If
    End If
    'PageDown
    If KeyCode = vbKeyPageDown Then
        If gInt + CisVsGrid1.DispMax > vsGrid1.Rows - 1 Then
            gInt = vsGrid1.Rows - 1
        Else
            gInt = gInt + CisVsGrid1.DispMax
        End If
    End If
    vsGrid1.Row = gInt

    iB1_ZKbn.SetFocus
'グリッドからテキストに表示
    If SyoriKB <> "DEL" Then Call VsGridToText(vsGrid1.Row)
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With vsGrid1
'       >> 取引先
        TRM.略称 = ""
        If RTrim(iB1_Torcd) <> "" Then
            If TorNmGet(iB1_Torcd, 9) Then
                .TextMatrix(tRow, CisVsGrid1.FixedGet(2)) = RTrim(TRM.略称)
            End If
        End If
        If RTrim(TRM.略称) <> "" Then
           .TextMatrix(tRow, CisVsGrid1.FixedGet(1)) = RTrim(iB1_Torcd) & "：" & RTrim(TRM.略称)
        Else
           .TextMatrix(tRow, CisVsGrid1.FixedGet(1)) = RTrim(iB1_Torcd)
        End If
'       >> 受入
        .TextMatrix(tRow, CisVsGrid1.FixedGet(2)) = iB1_Uke + Space(2 - CisFun.Len2(iB1_Uke))
'       >> 背番号
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/10 UPDATE START
        '.TextMatrix(tRow, CisVsGrid1.FixedGet(3)) = iB1_Seban + Space(5 - CisFun.Len2(iB1_Seban))
        .TextMatrix(tRow, CisVsGrid1.FixedGet(3)) = iB1_Seban + Space(8 - CisFun.Len2(iB1_Seban))
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/10 UPDATE END
'       >> 品番
        gSL_Select = "SELECT 表示品番 FROM 品番マスタ "
        gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iB1_Hinbn) & "'"
        If HNMRead(gSL_Select, 1) Then
           .TextMatrix(tRow, CisVsGrid1.FixedGet(4)) = HNM.表示品番 + Space(30 - CisFun.Len2(HNM.表示品番))
        Else
           .TextMatrix(tRow, CisVsGrid1.FixedGet(4)) = iB1_Hinbn + Space(30 - CisFun.Len2(iB1_Hinbn))
        End If
'       >> 増減区分
        Select Case iB1_ZKbn
               Case "1": .TextMatrix(tRow, CisVsGrid1.FixedGet(11)) = "減"
               Case "2": .TextMatrix(tRow, CisVsGrid1.FixedGet(11)) = "増"
               Case Else: .TextMatrix(tRow, CisVsGrid1.FixedGet(11)) = ""
        End Select
'       >> 増減設定枚数
        .TextMatrix(tRow, CisVsGrid1.FixedGet(12)) = Format(CisFun.Val2(iB1_ZSetSu), "##,###")
'       >> 増減便当枚数
        .TextMatrix(tRow, CisVsGrid1.FixedGet(13)) = Format(CisFun.Val2(iB1_ZBinSu), "##,###")
'       >> 平準化区分
        
        Select Case iB1_HKbn
               Case "1": .TextMatrix(tRow, CisVsGrid1.FixedGet(14)) = "日"
               Case "2": .TextMatrix(tRow, CisVsGrid1.FixedGet(14)) = "便"
               Case "3": .TextMatrix(tRow, CisVsGrid1.FixedGet(14)) = "月"
               Case Else: .TextMatrix(tRow, CisVsGrid1.FixedGet(14)) = ""
        End Select
'       >> 平準設定枚数
        .TextMatrix(tRow, CisVsGrid1.FixedGet(15)) = Format(CisFun.Val2(iB1_HSetSu), "##,###")

'       >>
        .TextMatrix(tRow, CisVsGrid1.FixedGet(28)) = RTrim(iB1_Hinbn)
        .TextMatrix(tRow, CisVsGrid1.FixedGet(29)) = RTrim(iB1_Torcd)
        .TextMatrix(tRow, CisVsGrid1.FixedGet(30)) = iB1_ZKbn
        .TextMatrix(tRow, CisVsGrid1.FixedGet(31)) = iB1_HKbn
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With vsGrid1
        Input_Row.Tag = tRow
        Input_Row.Move Input_Row.Left, vsGrid1.CellTop, Input_Row.Width, vsGrid1.CellHeight
        '仕入先
        iB1_Torcd = RTrim(.TextMatrix(tRow, CisVsGrid1.FixedGet(29)))
'        Call iB1_Torcd_LostFocus        '仕入先名表示
        If RTrim(iB1_Torcd) <> "" Then
           If TorNmGet(iB1_Torcd, 9) Then
               B1lb_Tornm = RTrim(TRM.略称)
           Else
               B1lb_Tornm = ""
           End If
        Else
           B1lb_Tornm = ""
        End If
        
        B1lb_Torcd = RTrim(.TextMatrix(tRow, CisVsGrid1.FixedGet(29)))
       '>> 受入
        iB1_Uke = RTrim$(.TextMatrix(tRow, CisVsGrid1.FixedGet(2)))
        B1lb_Uke = RTrim$(.TextMatrix(tRow, CisVsGrid1.FixedGet(2)))
        '>> 背番号
        iB1_Seban = RTrim$(.TextMatrix(tRow, CisVsGrid1.FixedGet(3)))
        B1lb_Seban = RTrim$(.TextMatrix(tRow, CisVsGrid1.FixedGet(3)))
        '>> 品番
        iB1_Hinbn = RTrim$(.TextMatrix(tRow, CisVsGrid1.FixedGet(28)))
        B1lb_Hinbn = RTrim$(.TextMatrix(tRow, CisVsGrid1.FixedGet(4)))      '表示品番
        '>> 収容数
        B1lb_Syuyo = Format(CisFun.Val2(RTrim$(.TextMatrix(tRow, CisVsGrid1.FixedGet(5)))), "##,###")
        '>> 内示数
        B1lb_Naisu = Format(CisFun.Val2(RTrim$(.TextMatrix(tRow, CisVsGrid1.FixedGet(6)))), "##,###")
        '>> 設定期間
        B1lb_SetYM = .TextMatrix(tRow, CisVsGrid1.FixedGet(7))
        '>> 前回回転枚数
        B1lb_ZKMai = Format(CisFun.Val2(RTrim$(.TextMatrix(tRow, CisVsGrid1.FixedGet(8)))), "##,###")
        '>> 今回回転枚数
        B1lb_KKMai = Format(CisFun.Val2(RTrim$(.TextMatrix(tRow, CisVsGrid1.FixedGet(9)))), "##,###")
        '>> 有効枚数
        B1lb_YukoMai = Format(CisFun.Val2(RTrim$(.TextMatrix(tRow, CisVsGrid1.FixedGet(10)))), "##,###")
        '>> 増減区分
        iB1_ZKbn = RTrim$(.TextMatrix(tRow, 30))
        '>> 増減設定枚数
        iB1_ZSetSu = Format(CisFun.Val2(RTrim$(.TextMatrix(tRow, CisVsGrid1.FixedGet(12)))), "##,###")
        '>> 増減便当枚数
        iB1_ZBinSu = Format(CisFun.Val2(RTrim$(.TextMatrix(tRow, CisVsGrid1.FixedGet(13)))), "##,###")
        '>> 平準区分
        iB1_HKbn = RTrim$(.TextMatrix(tRow, 31))
        '>> 平準設定枚数
        iB1_HSetSu = Format(CisFun.Val2(RTrim$(.TextMatrix(tRow, CisVsGrid1.FixedGet(15)))), "##,###")
    
        Call iB1_ZKbn_LostFocus
        Call iB1_HKbn_LostFocus
    End With

    'かんばん流動・増減・平準化・登録済分は、【取引先】【受入】【品番】【背番号】入力不可
    If RTrim$(vsGrid1.TextMatrix(tRow, CisVsGrid1.FixedGet(34))) = "X" Then
       iB1_Torcd.Visible = False: B1lb_Torcd.Visible = True
       iB1_Uke.Visible = False: B1lb_Uke.Visible = True
       iB1_Seban.Visible = False: B1lb_Seban.Visible = True
       iB1_Hinbn.Visible = False: B1lb_Hinbn.Visible = True
    Else
       iB1_Torcd.Visible = True: B1lb_Torcd.Visible = False
       iB1_Uke.Visible = True: B1lb_Uke.Visible = False
       iB1_Seban.Visible = True: B1lb_Seban.Visible = False
       iB1_Hinbn.Visible = True: B1lb_Hinbn.Visible = False
    End If
    
    If RTrim(iB1_Hinbn) <> "" And RTrim(iB1_Seban) <> "" And RTrim(iB1_Torcd) <> "" And RTrim(iB1_Uke) <> "" Then
       PB_SMNT.Visible = True
    Else
       PB_SMNT.Visible = False
    End If
    
End Sub
Private Sub vsGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown And KeyCode <> vbKeySpace Then Exit Sub
    If SyoriKB <> "DEL" Then Call TextToVsGrid(Val(Input_Row.Tag))
    If SyoriKB <> "DEL" Then Call VsGridToText(vsGrid1.Row)
    If SyoriKB = "DEL" And KeyCode = vbKeySpace Then Call vsGrid1_Click
End Sub
'*******************************************************************
'* 詳細画面から戻ってきた時の再表示
'*******************************************************************
Private Sub Grid_ReDsp()
    Dim Ix      As Integer
    Dim Iy      As Integer
    Dim wRow    As Integer
    Dim wChk    As Byte
    
'    vsGrid1.Redraw = False
    
    DoEvents
    
    '詳細ボタンが押された行をクリアー
    wChk = 0
    For Iy = 1 To vsGrid1.Rows - 1
        If vsGrid1.TextMatrix(Iy, CisVsGrid1.FixedGet(16)) = RV_Torcd And _
           vsGrid1.TextMatrix(Iy, CisVsGrid1.FixedGet(17)) = RV_Hinbn And _
           vsGrid1.TextMatrix(Iy, CisVsGrid1.FixedGet(26)) = RV_Ukeir And _
           vsGrid1.TextMatrix(Iy, CisVsGrid1.FixedGet(27)) = RV_Seban Then
               vsGrid1.TextMatrix(Iy, 0) = ""
               vsGrid1.TextMatrix(Iy, 1) = ""
               wChk = wChk + 1
        End If
    Next Iy
    
    '更新データをGrid最終行以降に追加
    vsGrid1.Rows = CisFun.Val2(B1lb_Kensu) + RVS_ZG_Count + 1
       
    With vsGrid1
        For Ix = 0 To RVS_ZG_Count - 1
        
            wRow = CisFun.Val2(B1lb_Kensu) + Ix + 1
            'No
            .TextMatrix(wRow, CisVsGrid1.FixedGet(0)) = CisFun.Val2(B1lb_Kensu) + Ix + 1
            '仕入先
            .TextMatrix(wRow, CisVsGrid1.FixedGet(1)) = RTrim(RV_Torcd) & "：" & RTrim(RV_TorRName)
            '受入
            .TextMatrix(wRow, CisVsGrid1.FixedGet(2)) = RTrim(RV_Ukeir)
            '背番号
            .TextMatrix(wRow, CisVsGrid1.FixedGet(3)) = RTrim(RV_Seban)
            '品番
            .TextMatrix(wRow, CisVsGrid1.FixedGet(4)) = RTrim(RV_ZG_DSPHin)
            '収容数
            .TextMatrix(wRow, CisVsGrid1.FixedGet(5)) = Format(RV_ZG_Syuyo, "##,###")
            '当月内示数
            .TextMatrix(wRow, CisVsGrid1.FixedGet(6)) = Format(RV_ZG_Naisu, "##,###")
            '設定期間
            If RTrim(RVS_ZG_SYMD(Ix)) = "" And RTrim(RVS_ZG_EYMD(Ix)) = "" Then
               RVS_ZG_SYMD(Ix) = mStartYMD
               RVS_ZG_EYMD(Ix) = mEndYMD
            End If
            If Trim(RVS_ZG_EYMD(Ix)) = "" Then
                .TextMatrix(wRow, CisVsGrid1.FixedGet(7)) = CisFun.Mid2(Format(RVS_ZG_SYMD(Ix), "0000/00/00"), 6, 5) & Space(7)
            Else
                .TextMatrix(wRow, CisVsGrid1.FixedGet(7)) = CisFun.Mid2(Format(RVS_ZG_SYMD(Ix), "0000/00/00"), 6, 5) _
                                                     & "～" & CisFun.Mid2(Format(RVS_ZG_EYMD(Ix), "0000/00/00"), 6, 5)
            End If
            .TextMatrix(wRow, CisVsGrid1.FixedGet(35)) = RVS_ZG_SYMD(Ix)
            .TextMatrix(wRow, CisVsGrid1.FixedGet(36)) = RVS_ZG_EYMD(Ix)
            
            If Trim(RVS_ZG_EYMD(Ix)) = "" Then
                gStr = CisFun.Mid2(Format(RVS_ZG_SYMD(Ix), "0000/00/00"), 6, 5) & Space(7)
            Else
                gStr = CisFun.Mid2(Format(RVS_ZG_SYMD(Ix), "0000/00/00"), 6, 5) _
                       & "～" & CisFun.Mid2(Format(RVS_ZG_EYMD(Ix), "0000/00/00"), 6, 5)
            End If
            '前回回転枚数
            .TextMatrix(wRow, CisVsGrid1.FixedGet(8)) = Format(RV_ZG_ZenKai, "#,###")
            '今回回転枚数
            .TextMatrix(wRow, CisVsGrid1.FixedGet(9)) = Format(RV_ZG_KonKai, "#,###")
            '有効枚数
            .TextMatrix(wRow, CisVsGrid1.FixedGet(10)) = Format(RV_ZG_Yuko, "#,###")
        
            '増減区分
            Select Case RVS_ZG_ZGKB(Ix)
                   Case 1: .TextMatrix(wRow, CisVsGrid1.FixedGet(11)) = "減"
                   Case 2: .TextMatrix(wRow, CisVsGrid1.FixedGet(11)) = "増"
                   Case Else: .TextMatrix(wRow, CisVsGrid1.FixedGet(11)) = ""
            End Select
            '設定枚数
            .TextMatrix(wRow, CisVsGrid1.FixedGet(12)) = Format(RVS_ZG_ZGSet(Ix), "###")
            '便当枚数
            .TextMatrix(wRow, CisVsGrid1.FixedGet(13)) = Format(RVS_ZG_ZGBin(Ix), "###")
            '平準化区分
            Select Case RVS_ZG_HJKB(Ix)
                   Case 1: .TextMatrix(wRow, CisVsGrid1.FixedGet(14)) = "日"
                   Case 2: .TextMatrix(wRow, CisVsGrid1.FixedGet(14)) = "便"
                   Case 3: .TextMatrix(wRow, CisVsGrid1.FixedGet(14)) = "月"
                   Case Else: .TextMatrix(wRow, CisVsGrid1.FixedGet(14)) = ""
            End Select
            '平準設定枚数
            .TextMatrix(wRow, CisVsGrid1.FixedGet(15)) = Format(RVS_ZG_HJSet(Ix), "###")
        
            '退避
            .TextMatrix(wRow, CisVsGrid1.FixedGet(16)) = RV_Torcd
            .TextMatrix(wRow, CisVsGrid1.FixedGet(17)) = RV_Hinbn
            .TextMatrix(wRow, CisVsGrid1.FixedGet(18)) = RVS_ZG_SYMD(Ix)
            .TextMatrix(wRow, CisVsGrid1.FixedGet(19)) = RVS_ZG_EYMD(Ix)
            .TextMatrix(wRow, CisVsGrid1.FixedGet(20)) = RVS_ZG_ZGSet(Ix)
            .TextMatrix(wRow, CisVsGrid1.FixedGet(21)) = RVS_ZG_ZGBin(Ix)
            .TextMatrix(wRow, CisVsGrid1.FixedGet(22)) = RVS_ZG_RuiMai(Ix)
            .TextMatrix(wRow, CisVsGrid1.FixedGet(23)) = RVS_ZG_ZGKB(Ix)
            .TextMatrix(wRow, CisVsGrid1.FixedGet(24)) = RVS_ZG_HJKB(Ix)
            .TextMatrix(wRow, CisVsGrid1.FixedGet(25)) = RVS_ZG_HJSet(Ix)
            .TextMatrix(wRow, CisVsGrid1.FixedGet(26)) = RV_Ukeir
            .TextMatrix(wRow, CisVsGrid1.FixedGet(27)) = RV_Seban
            
            .TextMatrix(wRow, CisVsGrid1.FixedGet(29)) = RV_Torcd
            .TextMatrix(wRow, CisVsGrid1.FixedGet(28)) = RV_Hinbn
            .TextMatrix(wRow, CisVsGrid1.FixedGet(30)) = RVS_ZG_ZGKB(Ix)
            .TextMatrix(wRow, CisVsGrid1.FixedGet(31)) = RVS_ZG_HJKB(Ix)
            
            .TextMatrix(wRow, CisVsGrid1.FixedGet(37)) = RVS_ZG_ZGNKB(Ix)
            .TextMatrix(wRow, CisVsGrid1.FixedGet(38)) = RVS_ZG_Kan(Ix)
            .TextMatrix(wRow, CisVsGrid1.FixedGet(39)) = RVS_ZG_HJNKB(Ix)
            
            .TextMatrix(wRow, CisVsGrid1.FixedGet(34)) = "X"
        
        Next Ix
    
    End With
    
    'クリアーした行を削除
    Ix = vsGrid1.Rows - wChk
    For Iy = 1 To Ix - 1
        If vsGrid1.TextMatrix(Iy, CisVsGrid1.FixedGet(0)) = "" Or vsGrid1.TextMatrix(Iy, CisVsGrid1.FixedGet(0)) = "" Then
            vsGrid1.RemoveItem (Iy)
            Iy = Iy - 1
        End If
    Next Iy
    
    '仕入先・受入・背番号・品番にてソート
    vsGrid1.Col = 1: vsGrid1.ColSel = 7
    vsGrid1.Sort = flexSortGenericAscending
    
    DoEvents
    
    If vsGrid1.Rows < 26 Then
       vsGrid1.Rows = 26
    End If
    
    DoEvents
    
    '連番振り直し
    For Iy = 1 To vsGrid1.Rows - 1
        vsGrid1.TextMatrix(Iy, CisVsGrid1.FixedGet(0)) = ""
        vsGrid1.TextMatrix(Iy, CisVsGrid1.FixedGet(0)) = Format(Iy, "##0")
    Next Iy
    
    vsGrid1.Redraw = True
    
    Ix = 0
    For Iy = 1 To vsGrid1.Rows - 1
        If vsGrid1.TextMatrix(Iy, CisVsGrid1.FixedGet(1)) <> "" Then
           Ix = Ix + 1
        End If
    Next Iy
    
    B1lb_Kensu = Format(Ix, "###,###")

End Sub
