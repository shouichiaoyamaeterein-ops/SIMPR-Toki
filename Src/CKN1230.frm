VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CKN1230 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "工程かんばん発行"
   ClientHeight    =   11100
   ClientLeft      =   105
   ClientTop       =   360
   ClientWidth     =   15120
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
   ScaleHeight     =   11100
   ScaleWidth      =   15120
   WindowState     =   2  '最大化
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1665
      Left            =   1410
      ScaleHeight     =   1605
      ScaleWidth      =   10785
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4560
      Width           =   10845
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   1470
         Left            =   60
         Top             =   60
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   2593
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
         Begin Cis3D_v60.CIS3D CIS3D28 
            Height          =   1050
            Left            =   30
            Top             =   390
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   1852
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
            cAlingnment     =   6
            cPositionY      =   60
            Begin Cis3D_v60.CIS3D B1lb_SHYmd 
               Height          =   420
               Left            =   105
               Top             =   300
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   741
               BackColor       =   16777152
               Caption         =   "9999/99/99"
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
            Begin Cis3D_v60.CIS3D B1lb_SHBin 
               Height          =   420
               Left            =   1515
               Top             =   300
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   741
               BackColor       =   16777152
               Caption         =   "99"
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
         Begin Cis3D_v60.CIS3D BLB_Sur 
            Height          =   525
            Index           =   1
            Left            =   2040
            Top             =   390
            Width           =   7365
            _ExtentX        =   12991
            _ExtentY        =   926
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   5
               Left            =   2100
               Top             =   0
               Width           =   1950
               _ExtentX        =   3440
               _ExtentY        =   900
               ForeColor       =   0
               Caption         =   "　"
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
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   330
                  Index           =   0
                  Left            =   375
                  Top             =   75
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   582
                  BackColor       =   16777152
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -100
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   6
               Left            =   255
               Top             =   0
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   900
               ForeColor       =   0
               Caption         =   "　"
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
               Begin Cis3D_v60.CIS3D B1lb_Syuki 
                  Height          =   330
                  Index           =   0
                  Left            =   210
                  Top             =   75
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "XXXXXXXXX1"
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
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   7
               Left            =   4395
               Top             =   0
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   900
               ForeColor       =   16711680
               Caption         =   "　"
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
               Begin CisText_V60.CisText iB1_HSu 
                  Height          =   360
                  Index           =   0
                  Left            =   315
                  TabIndex        =   4
                  Top             =   75
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
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
            End
            Begin Cis3D_v60.CIS3D Back_SSu 
               Height          =   510
               Index           =   0
               Left            =   5655
               Top             =   0
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   900
               ForeColor       =   16711680
               Caption         =   "　"
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
               Begin CisText_V60.CisText iB1_SSSu 
                  Height          =   360
                  Index           =   0
                  Left            =   150
                  TabIndex        =   5
                  Top             =   75
                  Width           =   540
                  _ExtentX        =   953
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
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin CisText_V60.CisText iB1_ESSu 
                  Height          =   360
                  Index           =   0
                  Left            =   1035
                  TabIndex        =   6
                  Top             =   75
                  Width           =   540
                  _ExtentX        =   953
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
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin Cis3D_v60.CIS3D CIS3D17 
                  Height          =   255
                  Left            =   780
                  Top             =   120
                  Width           =   210
                  _ExtentX        =   370
                  _ExtentY        =   450
                  Caption         =   "～"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   9
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   510
               Left            =   15
               Top             =   0
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   900
               BackColor       =   12648384
               Caption         =   "1"
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
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   1395
            Index           =   0
            Left            =   9405
            Top             =   45
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   2461
            ForeColor       =   0
            Caption         =   "最終№"
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
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D B1lb_EndSu 
               Height          =   615
               Left            =   120
               Top             =   510
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   1085
               BackColor       =   16777152
               Caption         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   18
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
               cAlingnment     =   7
               cBoderStyle     =   1
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   345
            Left            =   30
            Top             =   45
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   609
            Caption         =   "収容変更日･便"
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
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   345
            Index           =   2
            Left            =   4140
            Top             =   45
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   609
            ForeColor       =   0
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
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   345
            Index           =   50
            Left            =   2055
            Top             =   45
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   609
            ForeColor       =   0
            Caption         =   "収容器"
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
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   345
            Index           =   1
            Left            =   6435
            Top             =   45
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   609
            ForeColor       =   16711680
            Caption         =   "追加枚数"
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
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   345
            Index           =   4
            Left            =   7695
            Top             =   45
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   609
            ForeColor       =   16711680
            Caption         =   "再発行№"
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
         End
         Begin Cis3D_v60.CIS3D BLB_Sur 
            Height          =   525
            Index           =   0
            Left            =   2040
            Top             =   915
            Width           =   7365
            _ExtentX        =   12991
            _ExtentY        =   926
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   10
               Left            =   2100
               Top             =   15
               Width           =   1950
               _ExtentX        =   3440
               _ExtentY        =   900
               ForeColor       =   0
               Caption         =   "　"
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
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   330
                  Index           =   1
                  Left            =   375
                  Top             =   75
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   582
                  BackColor       =   16777152
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -100
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   510
               Index           =   11
               Left            =   255
               Top             =   15
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   900
               ForeColor       =   0
               Caption         =   "　"
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
               Begin Cis3D_v60.CIS3D B1lb_Syuki 
                  Height          =   330
                  Index           =   1
                  Left            =   225
                  Top             =   75
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "XXXXXXXXX1"
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
            Begin Cis3D_v60.CIS3D Back_HSu 
               Height          =   510
               Index           =   1
               Left            =   4395
               Top             =   15
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   900
               ForeColor       =   16711680
               Caption         =   "　"
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
               Begin CisText_V60.CisText iB1_HSu 
                  Height          =   360
                  Index           =   1
                  Left            =   315
                  TabIndex        =   7
                  Top             =   75
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
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
            End
            Begin Cis3D_v60.CIS3D Back_SSu 
               Height          =   510
               Index           =   1
               Left            =   5655
               Top             =   15
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   900
               ForeColor       =   16711680
               Caption         =   "　"
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
               Begin CisText_V60.CisText iB1_SSSu 
                  Height          =   360
                  Index           =   1
                  Left            =   150
                  TabIndex        =   8
                  Top             =   75
                  Width           =   540
                  _ExtentX        =   953
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
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin CisText_V60.CisText iB1_ESSu 
                  Height          =   360
                  Index           =   1
                  Left            =   1035
                  TabIndex        =   9
                  Top             =   75
                  Width           =   540
                  _ExtentX        =   953
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
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin Cis3D_v60.CIS3D CIS3D26 
                  Height          =   255
                  Left            =   780
                  Top             =   120
                  Width           =   210
                  _ExtentX        =   370
                  _ExtentY        =   450
                  Caption         =   "～"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   9
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
            Begin Cis3D_v60.CIS3D CIS3D27 
               Height          =   510
               Left            =   15
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   900
               BackColor       =   12648384
               Caption         =   "2"
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
      Height          =   930
      Left            =   1440
      ScaleHeight     =   870
      ScaleWidth      =   9090
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1260
      Width           =   9150
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Left            =   60
         Top             =   60
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1376
         ForeColor       =   192
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
         Begin CisText_V60.CisText iH1_Seban 
            Height          =   360
            Left            =   105
            TabIndex        =   0
            Top             =   345
            Width           =   1140
            _ExtentX        =   2011
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
      End
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   780
         Left            =   5685
         Top             =   60
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "品  番"
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
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   360
            Left            =   90
            TabIndex        =   3
            Top             =   330
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Index           =   1
         Left            =   1395
         Top             =   60
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "社内手配先"
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
            Left            =   60
            TabIndex        =   1
            Top             =   330
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
            Height          =   360
            Left            =   1035
            Top             =   315
            Width           =   2475
            _ExtentX        =   4366
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
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   780
         Index           =   1
         Left            =   4980
         Top             =   60
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   1376
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
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   360
            Left            =   150
            TabIndex        =   2
            Top             =   330
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
         End
      End
   End
   Begin VB.PictureBox B1_Area2 
      Height          =   2205
      Left            =   1410
      ScaleHeight     =   2145
      ScaleWidth      =   10785
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2310
      Width           =   10845
      Begin Cis3D_v60.CIS3D CIS3D16 
         Height          =   705
         Left            =   2880
         Top             =   15
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   1244
         Caption         =   "⇒"
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
         cPositionY      =   100
      End
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   705
         Left            =   15
         Top             =   720
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1244
         ForeColor       =   0
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Syakei 
            Height          =   330
            Left            =   135
            Top             =   315
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XXXXX"
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
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   705
         Left            =   1065
         Top             =   720
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   1244
         ForeColor       =   0
         Caption         =   "品    名"
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
         Begin Cis3D_v60.CIS3D B1lb_Hinnm 
            Height          =   330
            Left            =   90
            Top             =   315
            Width           =   4890
            _ExtentX        =   8625
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
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
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   705
         Index           =   0
         Left            =   15
         Top             =   15
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   1244
         ForeColor       =   0
         Caption         =   "工程"
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
         Begin Cis3D_v60.CIS3D B1lb_Kotei 
            Height          =   330
            Left            =   135
            Top             =   315
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   582
            BackColor       =   12648384
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   705
         Left            =   3405
         Top             =   15
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   1244
         ForeColor       =   0
         Caption         =   "後工程"
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
         Begin Cis3D_v60.CIS3D B1lb_AKotei 
            Height          =   330
            Left            =   120
            Top             =   315
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   582
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   2115
         Left            =   6285
         Top             =   15
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3731
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin vsFlexLib.vsFlexArray vsGrid2 
            Height          =   1995
            Left            =   120
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   75
            Width           =   4230
            _Version        =   196608
            _ExtentX        =   7461
            _ExtentY        =   3519
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
            FormatString    =   "№|作　業       |備　考　　　　|"
            Rows            =   6
            Cols            =   5
            BackColor       =   16777152
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   320
            GridLines       =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   705
         Index           =   0
         Left            =   15
         Top             =   1425
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1244
         ForeColor       =   0
         Caption         =   "生産ﾛｯﾄ"
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
         Begin Cis3D_v60.CIS3D B1lb_SLot 
            Height          =   330
            Index           =   0
            Left            =   210
            Top             =   315
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
            BackColor       =   16777152
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D B1lb_SLot 
            Height          =   330
            Index           =   1
            Left            =   1200
            Top             =   330
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
            BackColor       =   16777152
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   360
            Left            =   60
            Top             =   300
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   635
            BackColor       =   12648384
            Caption         =   "1"
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
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   330
            Left            =   1050
            Top             =   330
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   582
            BackColor       =   12648384
            Caption         =   "2"
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
      End
      Begin Cis3D_v60.CIS3D CIS3D19 
         Height          =   705
         Left            =   2100
         Top             =   1425
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   1244
         ForeColor       =   0
         Caption         =   "基準生産日"
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
         Begin Cis3D_v60.CIS3D B1lb_SKijyun 
            Height          =   330
            Left            =   30
            Top             =   315
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XX XX XX XX XX"
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
            cPositionX      =   60
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D24 
         Height          =   705
         Left            =   5475
         Top             =   1425
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1244
         ForeColor       =   0
         Caption         =   "工数"
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
         Begin Cis3D_v60.CIS3D B1lb_KKosu 
            Height          =   330
            Left            =   60
            Top             =   315
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "##0.0"
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
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D34 
         Height          =   705
         Left            =   4005
         Top             =   1425
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1244
         ForeColor       =   0
         Caption         =   "置　場"
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
         Begin Cis3D_v60.CIS3D B1lb_Okiba 
            Height          =   330
            Left            =   60
            Top             =   315
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "XXXXXXXXXX"
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
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   495
      Index           =   1
      Left            =   11220
      ScaleHeight     =   495
      ScaleWidth      =   2370
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6270
      Width           =   2370
      Begin Cis3D_v60.CIS3D UC_3D6 
         Height          =   435
         Left            =   30
         Top             =   30
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   767
         ForeColor       =   0
         Caption         =   " 発行総枚数"
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
         Begin Cis3D_v60.CIS3D iT1_Addsu 
            Height          =   285
            Left            =   1335
            Top             =   75
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            BackColor       =   12648447
            ForeColor       =   0
            Caption         =   "ZZ,ZZ9"
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
      Left            =   7335
      ScaleHeight     =   660
      ScaleWidth      =   7770
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   480
      Width           =   7830
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   5175
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
         Left            =   6030
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
         Left            =   6885
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
         Left            =   3465
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
      Begin CisBtn_60.CisBtn PB_NEXT 
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
         Left            =   1755
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
      Begin CisBtn_60.CisBtn PB_Prt 
         Height          =   585
         Left            =   4320
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
      Begin CisBtn_60.CisBtn PB_Del 
         Height          =   585
         Left            =   900
         Top             =   30
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
      Begin CisBtn_60.CisBtn PB_Add 
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
      Caption         =   "【 工程かんばん発行  】"
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
         Left            =   14295
         Top             =   75
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
         Left            =   11775
         Top             =   30
         Width           =   2520
         _ExtentX        =   4445
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
         ForeColor       =   65535
         Caption         =   "追加"
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
   Begin Cis3D_v60.CIS3D B2_Area1 
      Height          =   3810
      Left            =   1380
      Top             =   6750
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   6720
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
      Begin Cis3D_v60.CIS3D CIS3D37 
         Height          =   3660
         Left            =   75
         Top             =   60
         Width           =   12060
         _ExtentX        =   21273
         _ExtentY        =   6456
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   3570
            Left            =   60
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   60
            Width           =   11925
            _Version        =   196608
            _ExtentX        =   21034
            _ExtentY        =   6297
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
            FormatString    =   "№ |手配先 |手配先名(社内)　　|受入|品　番　　　　　　　　　|背番号|収容器   |収容数  |発行枚数||"
            Rows            =   11
            Cols            =   11
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
         End
      End
   End
End
Attribute VB_Name = "CKN1230"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   工程かんばん発行
'**       フォームID    :   CKN0230
'**       処理概要      :
'**
'**       作  成  日    :   2003/12/24  By CIS SIMPR-A
'**       変  更  日    :   2004/09/10  By CIS ｾﾝﾄﾛﾆｸｽ対応
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim CisVsGrid       As New CisVsGrid3
    Dim CisVsGrid2      As New CisVsGrid3
    
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
    
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 9
        .InitGet = False
        .Init
    End With
    Call GridInit
    
    With CisVsGrid2
        Set .GridObj = vsGrid2
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 3
        .InitGet = False
        .Init
    End With
    Call GridInit2
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
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
    FormAct = Not FormAct
' 自社情報取得
    Kn_Jisya = ""
    Kn_JisyaNM = ""
    Kn_JisyaKJNM = ""
    
    Kn_Jisya = RTrim(gCompany)
' 自社伝票発行用(工場)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 7
        IRN.oyian = "0"
        IRN.oyian2 = "00"
    End If
    ' 工場
    Kn_Jisya = Kn_Jisya & RTrim(IRN.oyian)
    
    gSL_Select = "Select * From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & Kn_Jisya & "'"
    If TRMRead(gSL_Select, 1) Then
       Kn_JisyaNM = TRM.略称
       Kn_JisyaKJNM = TRM.工場名
    End If
'
    Call GridSet
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
           Case vbKeyF1:       If PB_Add.Visible Then Call PB_Add_Click    '【新規】
           Case vbKeyF3:       If PB_Del.Visible Then Call PB_Del_Click    '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF8:       If PB_Prt.Visible Then Call PB_Prt_Click    '【印刷】
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyEscape:   If PB_Can.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_Ent.Visible Then                      '【入力】
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
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      iH1_Seban.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    追　加  キ ー(F1)        +
'+----------------------------+
Private Sub PB_Add_Click()
    If SyoriKB <> "ADD" Then
       K_Sykbnm = "追加"
       K_Sykbnm.ForeColor = PB_Add.BtnForeColor
       SyoriKB = "ADD"
       ProcHB = "H1"
       vsGrid1.HighLight = flexHighlightNever
       Call DispChange(ProcHB)
    End If
End Sub
'+----------------------------+
'+    削　除  キ ー(F3)        +
'+----------------------------+
Private Sub PB_Del_Click()
    If SyoriKB <> "DEL" Then
       K_Sykbnm = "削除"
       K_Sykbnm.ForeColor = PB_Del.BtnForeColor
       SyoriKB = "DEL"
       ProcHB = "B2"
       Call DispChange(ProcHB)
       vsGrid1.HighLight = flexHighlightAlways
       vsGrid1.SetFocus
    End If
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = "Select HT.* From 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " Where HT.背番号 + HT.取引先 + HT.受入 + HT.品番 < '"
    gSL_Select = gSL_Select & RTrim(iH1_Seban) & RTrim(iH1_Torcd) & RTrim(iH1_Ukeir) & RTrim(iH1_Hinbn)
    gSL_Select = gSL_Select & "'"
    gSL_Select = gSL_Select & " And 品目 = 1 and SK区分 = 0 And 手配区分 = '2'"
    gSL_Select = gSL_Select & " ORDER BY HT.背番号 + HT.取引先 + HT.受入 + HT.品番 Desc "
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Seban.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = "Select HT.* From 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " Where HT.背番号 + HT.取引先 + HT.受入 + HT.品番 > '"
    gSL_Select = gSL_Select & RTrim(iH1_Seban) & RTrim(iH1_Torcd) & RTrim(iH1_Ukeir) & RTrim(iH1_Hinbn)
    gSL_Select = gSL_Select & "'"
    gSL_Select = gSL_Select & " And 品目 = 1 and SK区分 = 0 And 手配区分 = '2'"
    gSL_Select = gSL_Select & " ORDER BY HT.背番号 + HT.取引先 + HT.受入 + HT.品番 "
    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Seban.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    印　刷  キ ー(F8)       +
'+----------------------------+
Private Sub PB_Prt_Click()
    kn_Call = "C"
    kn_PrtKB = 1
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "   　かんばん発行を行いますか？       "
        .MB_Title = ""
        .MB_Button = Yes_No
        If Not .MBOX Then Exit Sub
    End With
    
    
    If Not CKN1240LPrint Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　かんばん発行データがありません       "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Sub
    End If
    Call GridSet
    
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Seban" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd: Exit Sub
End Sub
'+----------------------------+
'+       仕入先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 6
    RV_Call = ""
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+-----------------------------+
'+       取引先受入検索         +
'+-----------------------------+
Private Sub Look_UTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 6
    RV_Torcd = iH1_Torcd
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+-----------------------------+
'+       品番検索              +
'+-----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
        
    RVI_Hinmoku = 1 ' 発注
    RVI_SK = 0      ' かんばん
    RV_Call = ""
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
        iH1_Hinbn = RV_Hinbn
        iH1_Seban = RV_Seban
    End If
End Sub
'+-----------------------------+
'+  iH1_Seban
'+-----------------------------+
Private Sub iH1_Seban_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Seban_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    ' 背番号より他情報表示
    If Trim(iH1_Seban) <> "" Then
        gSL_Select = "Select 取引先,受入,品番 From 品番取引先マスタ"
        gSL_Select = gSL_Select & " Where 背番号 = '" & iH1_Seban & "'"
        gSL_Select = gSL_Select & "   And 品目 = 1"
        gSL_Select = gSL_Select & "   And SK区分 = 0"
        gSL_Select = gSL_Select & "   And 手配区分 = '2'"
        If HTMRead(gSL_Select, 0, 1) Then
           If CisDB.RecordCount = 1 Then
              iH1_Torcd = HTM.取引先
              iH1_Ukeir = HTM.受入
              iH1_Hinbn = HTM.品番
              Call iH1_Torcd_LostFocus
           End If
        End If
        Call HTMClose(1)
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
    If TorNmGet(iH1_Torcd, 6, 1) Then
       H1lb_Tornm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iH1_Ukeir
'+-----------------------------+
Private Sub iH1_Ukeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
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
    PB_Look.Visible = True
End Sub
Private Sub iH1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
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
    H1Mode = False
    H1Color = gPLostSel
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
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
       Case "B2"
          B2Mode = True
          B2Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B1_Area2.Enabled = B1Mode
    B2_Area1.Enabled = True
    B2_Area1.BackColor = B2Color
    
    PB_Look.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_Can.Visible = B1Mode
    PB_Del.Visible = H1Mode
    PB_Add.Visible = B2Mode
    
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
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        CisVsGrid2.Clear
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
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B2" Then
        If Not DBDelete Then
           vsGrid1.SetFocus
           GoTo ReturnPress_Ed
        End If
        vsGrid1.SetFocus
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
    If Trim(iH1_Seban) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    背番号を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Seban.SetFocus
       Exit Function
    End If
    
    gSL_Select = "select * From 品番取引先マスタ"
    gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & " And   背番号 = '" & RTrim(iH1_Seban) & "'"
    If Not HTMRead(gSL_Select, 1, 1) Then
        gSL_Select = "Select HT.背番号,HT.取引先,HT.受入,HT.品番,HN.表示品番,HN.社内背番,"
        gSL_Select = gSL_Select & " HT.最終発行連番,Rtrim(TR.略称) 取引先名,HT.手配区分"
        gSL_Select = gSL_Select & " From 品番取引先マスタ HT"
        gSL_Select = gSL_Select & " Left Outer Join 品番マスタ HN"
        gSL_Select = gSL_Select & "   On HN.品番 = HT.品番"
        gSL_Select = gSL_Select & " Left Outer Join 取引先マスタ TR"
        gSL_Select = gSL_Select & "   On TR.取引先CD = HT.取引先"
        gSL_Select = gSL_Select & "   And TR.取引先区分 = 1"
        gSL_Select = gSL_Select & " Where HT.背番号 = '" & iH1_Seban & "'"
        gSL_Select = gSL_Select & "   And HT.品目 = 1"
        gSL_Select = gSL_Select & "   And HT.SK区分 = 0"
        gSL_Select = gSL_Select & "   And HT.手配区分 = '2'"
        gSL_Select = gSL_Select & " Order By HT.品番,HT.取引先,HT.受入"
        If HTMRead(gSL_Select, 0, 1) Then
           If CisDB.RecordCount > 1 Then
              CKN0215.Show vbModal
              Unload CKN0215
              Set CKN0215 = Nothing
              If RV_Rtn Then
                 iH1_Torcd = RV_TorcdK
                 iH1_Ukeir = RV_Ukeir
                 iH1_Hinbn = RV_Hinbn
                Call iH1_Torcd_LostFocus
              Else
                 iH1_Seban.SetFocus
                 Exit Function
              End If
           End If
        End If
        Call HTMClose(1)
    End If
    
    If Trim(iH1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    社内手配先を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If Not TorNmGet(iH1_Torcd, 6) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    社内手配先　取引先マスタ未登録            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If Trim(iH1_Ukeir) <> "" Then
        gSL_Select = "select * From 取引先受入マスタ "
        gSL_Select = gSL_Select & " WHERE 取引先CD = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & "   And 取引先区分 = 1"
        If Not TUMRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    受入未登録            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Ukeir.SetFocus
           Exit Function
        End If
    End If
    If Trim(iH1_Hinbn) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    品番を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Hinbn.SetFocus
       Exit Function
    End If

' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = " Select HT.* From 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " Where HT.品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   HT.取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   HT.受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & " And   HT.背番号 = '" & RTrim(iH1_Seban) & "'"
    If Not DBInput("INV") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    品番取引先マスタ未登録            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Seban.SetFocus
        Exit Function
    Else
        If HTM.品目 <> 1 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     発注品番を指定して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Seban.SetFocus
            Exit Function
        End If
        If HTM.SK区分 <> 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     かんばんを指定して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Seban.SetFocus
            Exit Function
        End If
        If HTM.手配区分 <> "2" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     手配先が'社内'を指定して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Seban.SetFocus
            Exit Function
        End If
    End If
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Dim NoStr       As String
    Body1Chk = False
    
'    ' 追加枚数のみ（１）
    If Not Back_HSu(1).Visible And Not Back_SSu(0).Visible And Not Back_SSu(1).Visible Then
       If CisFun.Val2(iB1_HSu(0)) = 0 Then
          With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     追加枚数を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
          End With
          Exit Function
       End If
    End If
    If Not Back_HSu(1).Visible And Back_SSu(0).Visible And Not Back_SSu(1).Visible Then
       If CisFun.Val2(iB1_HSu(0)) = 0 And _
          CisFun.Val2(iB1_SSSu(0)) = 0 And _
          CisFun.Val2(iB1_ESSu(0)) = 0 Then
          With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     追加枚数又は再発行№を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
          End With
          Exit Function
       End If
    End If
    If Trim(B1lb_SHYmd) <> "" Then
       If CisFun.Val2(iB1_HSu(0)) = 0 And _
          CisFun.Val2(iB1_SSSu(0)) = 0 And _
          CisFun.Val2(iB1_ESSu(0)) = 0 And _
          CisFun.Val2(iB1_HSu(1)) = 0 And _
          CisFun.Val2(iB1_SSSu(1)) = 0 And _
          CisFun.Val2(iB1_ESSu(1)) = 0 Then
          With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     追加枚数又は再発行№を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
          End With
          Exit Function
       End If
     End If
    For gInt = 0 To 1
        If gInt = 0 Then
           NoStr = "①"
        Else
           NoStr = "②"
        End If
        If Trim(B1lb_SHYmd) = "" Then
           NoStr = ""
        End If
        If Back_SSu(gInt).Visible And _
           CisFun.Val2(iB1_HSu(gInt)) = 0 And _
           CisFun.Val2(iB1_SSSu(gInt)) = 0 And _
           CisFun.Val2(iB1_ESSu(gInt)) <> 0 Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "     再発行№" & NoStr & "(開始)を入力して下さい            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iB1_SSSu(gInt).SetFocus
           Exit Function
        End If
        If Back_SSu(gInt).Visible And _
           CisFun.Val2(iB1_HSu(gInt)) = 0 And _
           CisFun.Val2(iB1_ESSu(gInt)) = 0 And _
           CisFun.Val2(iB1_SSSu(gInt)) <> 0 Then
           With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "     再発行№" & NoStr & "(終了)を入力して下さい            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
            End With
            iB1_ESSu(gInt).SetFocus
            Exit Function
         End If
         If CisFun.Val2(iB1_SSSu(gInt)) > CisFun.Val2(iB1_ESSu(gInt)) Then
            With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "     再発行№" & NoStr & "を正しく入力して下さい（開始＞終了）            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
            End With
            iB1_ESSu(gInt).SetFocus
            Exit Function
         End If
         If CisFun.Val2(iB1_ESSu(gInt)) > CisFun.Val2(B1lb_EndSu) Then
            With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "     再発行№" & NoStr & "は最終№以下を指定して下さい            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
            End With
            iB1_ESSu(gInt).SetFocus
            Exit Function
         End If
    Next gInt
    
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
    Call HeadBodyClear("B")
    
    If Not HTMRead(gSL_Select, 1) Then
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
        iH1_Seban = RTrim(HTM.背番号)
        iH1_Torcd = RTrim(HTM.取引先)
        iH1_Ukeir = RTrim(HTM.受入)
        iH1_Hinbn = RTrim(HTM.品番)
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
    If HTM.品目 <> 1 Or HTM.SK区分 <> 0 Then Exit Sub

    With HTM
        B1lb_SLot(0) = Format(.ロット1, "#,###")
        B1lb_SLot(1) = Format(.ロット2, "#,###")
        B1lb_Okiba = .置場
        If .基準工数 <> 0 Then
           B1lb_KKosu = Format(.基準工数, "##0.0")
        End If
        
        B1lb_Syuki(0) = .収容器1
        B1lb_Syuyo(0) = Format(.収容数1, "#,###")
        B1lb_Syuki(1) = .収容器2
        B1lb_Syuyo(1) = Format(.収容数2, "#,###")
        
        If Trim(.変更日) <> "" Then
           B1lb_SHYmd = Mid(.変更日, 1, 4) & "/" & Mid(.変更日, 5, 2) & "/" & _
                        Mid(.変更日, 7, 2)
           B1lb_SHBin = .変更便
           Back_HSu(1).Visible = True
           Back_SSu(1).Visible = True
        Else
           Back_HSu(1).Visible = False
           Back_SSu(1).Visible = False
        End If
        If .最終発行連番 <> 0 Then
           B1lb_EndSu = Format(.最終発行連番, "000")
           Back_SSu(0).Visible = True
        Else
           Back_SSu(0).Visible = False
           Back_SSu(1).Visible = False
        End If
        gStr = ""
        If .Day1 <> 0 Then
            gStr = CisFun.RSetFld(.Day1, 2, "##") & " "
            If .Day2 <> 0 Then
                gStr = gStr & CisFun.RSetFld(.Day2, 2, "##") & " "
            End If
            If .Day3 <> 0 Then
                gStr = gStr & CisFun.RSetFld(.Day3, 2, "##") & " "
            End If
            If .Day4 <> 0 Then
                gStr = gStr & CisFun.RSetFld(.Day4, 2, "##") & " "
            End If
            If .Day5 <> 0 Then
                gStr = gStr & CisFun.RSetFld(.Day5, 2, "##")
            End If
        Else
            If .Week1 <> 0 Then gStr = "月 " Else gStr = Space(3)
            If .Week2 <> 0 Then gStr = gStr & "火 " Else gStr = gStr & Space(3)
            If .Week3 <> 0 Then gStr = gStr & "水 " Else gStr = gStr & Space(3)
            If .Week4 <> 0 Then gStr = gStr & "木 " Else gStr = gStr & Space(3)
            If .Week5 <> 0 Then gStr = gStr & "金 " Else gStr = gStr & Space(3)
        End If
        B1lb_SKijyun = gStr
    End With
    ' 品番情報
    gSL_Select = "select * From 品番マスタ "
    gSL_Select = gSL_Select & " Where 品番 = '" & HTM.品番 & "'"
    If HNMRead(gSL_Select, 1) Then
       With HNM
            B1lb_Syakei = .車型
            B1lb_Hinnm = .品名
            If MeisyoGet("工程管理", .工程) Then
               B1lb_Kotei = SYM_Meisyo
            End If
            If MeisyoGet("工程管理", .次工程) Then
               B1lb_AKotei = SYM_Meisyo
            End If
       End With
    End If
    ' 作業情報
    Call GridSet2
    '
    Call iH1_Torcd_LostFocus
    
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    削除 処理          +
'+----------------------+
Private Function DBDelete() As Boolean
    DBDelete = False
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     選択行を削除します      "
        .MB_Title = "削除処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
'
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    With CisDB
        .SQL = "Delete From かんばん発行テーブル"
        .SQL = .SQL & " Where かんばん種類 = 1"
        .SQL = .SQL & " And   背番号 = '" & RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 5)) & "'"
        .SQL = .SQL & " And   品番 = '" & RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 9)) & "'"
        .SQL = .SQL & " And   取引先 = '" & RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 1)) & "'"
        .SQL = .SQL & " And   受入 = '" & RTrim(vsGrid1.TextMatrix(vsGrid1.Row, 3)) & "'"
        .SQL = .SQL & " And   収容数 = " & CisFun.Val2(vsGrid1.TextMatrix(vsGrid1.Row, 7))
        If Not .DBExec Then GoTo DBDelete_Err
    End With
'
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

'
    Call GridSet
'
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisDB.DBTran (TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加 処理          +
'+----------------------+
Private Function DBPut() As Boolean
    DBPut = False
    Dim AddMai  As Long
    Dim PrtStart    As Long
    Dim PrtEnd      As Long
    Dim wTorNm      As String
    Dim wwTorcd     As String
    Dim wNTorNm     As String
    Dim PrtKb       As Integer
    Dim wPutFlg     As Boolean
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     発行データ作成処理を行います      "
        .MB_Title = "処理確認"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err
'   品番取引先情報
    gSL_Select = "Select *"
    gSL_Select = gSL_Select & " From 品番取引先マスタ "
    gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & " And   背番号 = '" & RTrim(iH1_Seban) & "'"
    gSL_Select = gSL_Select & " And   品目 = 1"
    gSL_Select = gSL_Select & " And   SK区分 = 0"
    If Not HTMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     品番取引先マスタ　異常      "
            .MB_Title = ""
            .MB_Button = Error
            Call .MBOX
        End With
        Exit Function
    End If
    ' 品番情報
    gSL_Select = "Select * From 品番マスタ "
    gSL_Select = gSL_Select & " Where 品番 = '" & HTM.品番 & "'"
    If Not HNMRead(gSL_Select, 1) Then
       Call ItemsClearHNM
    End If
    ' 取引先情報
    gSL_Select = "Select * From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先CD = '" & HTM.取引先 & "'"
    gSL_Select = gSL_Select & "   And 取引先区分 = 1"
    If Not TRMRead(gSL_Select, 1) Then
       Call ItemsClearTRM
    End If
    wTorNm = TRM.略称
    wwTorcd = TRM.取引先
    If Trim(TRM.工場CD) <> "" Then
       wwTorcd = wwTorcd & "-" & TRM.工場CD
    End If
    ' 納入先情報
    wNTorNm = ""
    If Trim(HTM.納入先) <> "" Then
        gSL_Select = "Select * From 取引先マスタ "
        gSL_Select = gSL_Select & " Where 取引先CD = '" & HTM.納入先 & "'"
        gSL_Select = gSL_Select & "   And 取引先区分 = 1"
        If Not TRMRead(gSL_Select, 1) Then
           Call ItemsClearTRM
        End If
        wNTorNm = TRM.略称
    End If
    
    ' かんばん発行情報出力
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' 発行済データ削除
    With CisDB
        .SQL = "Delete かんばん発行テーブル "
        .SQL = .SQL & "    Where かんばん種類 = 1"
        .SQL = .SQL & "      And 発行済 = 1"
        .DBExec
    End With
' ( 追加発行 )
    For gInt = 0 To 1
        PrtStart = 0: PrtEnd = 0
        AddMai = CisFun.Val2(iB1_HSu(gInt))
        If AddMai <> 0 Then
           PrtStart = CisFun.Val2(B1lb_EndSu) + 1
           PrtEnd = PrtStart + AddMai - 1
           PrtKb = 0        ' 追加発行
        End If
        If CisFun.Val2(iB1_SSSu(gInt)) <> 0 Then
           PrtStart = CisFun.Val2(iB1_SSSu(gInt))
           PrtKb = 1        ' 再発行
        End If
        If CisFun.Val2(iB1_ESSu(gInt)) <> 0 Then
           PrtEnd = CisFun.Val2(iB1_ESSu(gInt))
           PrtKb = 1        ' 再発行
        End If
        If PrtEnd = 0 Then
           PrtEnd = PrtStart
        End If
        
        If PrtStart <> 0 Then
            For gLong = PrtStart To PrtEnd
                wPutFlg = False
                gSL_Select = "Select * From かんばん発行テーブル"
                gSL_Select = gSL_Select & " Where かんばん種類 = 1"
                gSL_Select = gSL_Select & " And   背番号 = '" & RTrim(HTM.背番号) & "'"
                gSL_Select = gSL_Select & " And   品番 = '" & RTrim(HTM.品番) & "'"
                gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(HTM.取引先) & "'"
                gSL_Select = gSL_Select & " And   受入 = '" & RTrim(HTM.受入) & "'"
                gSL_Select = gSL_Select & " And   発行NO = " & gLong
                If Not KHTRead(gSL_Select, 1) Then
                    Call ItemsClearKHT
                    With KHT
                        .かんばん種類 = 1
                        .品番 = HTM.品番
                        .取引先 = HTM.取引先
                        .受入 = HTM.受入
                        .発行NO = gLong
                    End With
                    wPutFlg = True
                End If
                With KHT
                    .背番号 = HTM.背番号
                    .発行区分 = PrtKb
                    .発行種類 = 1
                    .表示品番 = HNM.表示品番
                    If Len(HNM.表示品番) < 15 Then
                       .品番14 = HNM.表示品番
                    Else
                       .品番30 = HNM.表示品番
                    End If
                    .品目 = 3       'HTM.品目
                    .品名 = HNM.品名
                    .受入名 = TUM.受入名称
                    .収容器 = B1lb_Syuki(gInt)
                    .収容数 = CisFun.Val2(B1lb_Syuyo(gInt))
                    .収容区分 = gInt + 1
                    .手配区分 = HTM.手配区分
                    .手配区分名 = "社内"
                    .アドレス = HTM.アドレス
                    .IMage = HNM.IMage
                    .取引先CD = wwTorcd
                    .取引先名 = TRM.略称
                    .取引先名1 = CisFun.Mid2(TRM.略称, 1, 10)
                    .取引先名2 = CisFun.Mid2(TRM.略称, 11, 10)
                    .取引先工場名 = TRM.工場名
                    .取引先工場名1 = CisFun.Mid2(TRM.工場名, 1, 10)
                    .取引先工場名2 = CisFun.Mid2(TRM.工場名, 11, 10)
                    If Trim(HTM.サイクル) <> "" Then
                        .サイクル = CisFun.RSetFld(Mid(HTM.サイクル, 1, 1), 1, "#") & "-" & _
                                    CisFun.RSetFld(Mid(HTM.サイクル, 2, 2), 2, "#") & "-" & _
                                    CisFun.RSetFld(Mid(HTM.サイクル, 4, 2), 2, "#")
                    End If
                    .置場 = HTM.置場
                    .車型 = HNM.車型
                    .自社CD = Kn_Jisya
                    .自社名1 = CisFun.Mid2(Kn_JisyaNM, 1, 20)
                    .自社名2 = CisFun.Mid2(Kn_JisyaKJNM, 1, 14)
                    .実数区分 = HNM.実数区分
                    .備考 = HTM.備考
                
                    .工程 = HNM.工程
                    .工程名 = CisFun.Mid2(RTrim(B1lb_Kotei), 1, 16)
                    .次工程 = HNM.次工程
                    .次工程名 = CisFun.Mid2(RTrim(B1lb_AKotei), 1, 16)
                    .基準工数 = HTM.基準工数
                    .基準生産日 = B1lb_SKijyun
                    .生産ロット = CisFun.Val2(B1lb_SLot(gInt))
                    .収容区分 = gInt + 1
                    
                    .作成者 = gTanto
                    
                End With
            
                If wPutFlg Then
                    Call KHTInsert
                Else
                    Call KHTUpdate
                End If
                B1lb_EndSu = Format(gLong, "000")
            Next gLong
        End If
        
        If AddMai <> 0 Then
           With CisDB
                .SQL = "Update 品番取引先マスタ Set  "
                .SQL = .SQL & "最終発行連番 = " & PrtEnd & ","
                .SQL = .SQL & "更新日 = GETDATE(),"
                .SQL = .SQL & "更新者 = '" & gTanto & "',"
                .SQL = .SQL & "更新端末 = SubString(Host_Name(),1,20)"
                .SQL = .SQL & " Where 品番 = '" & RTrim(HTM.品番) & "'"
                .SQL = .SQL & "   And 取引先 = '" & RTrim(HTM.取引先) & "'"
                .SQL = .SQL & "   And 受入 = '" & RTrim(HTM.受入) & "'"
                If Not .DBExec Then GoTo DBPut_Err
           End With
           Back_SSu(gInt).Enabled = True
        End If
    Next gInt

    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    Call GridSet
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisDB.DBTran (TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
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
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 0, 0, "№")
        Call .FixedSet(中中, 左中, あり, 7, 1, "手配先")
        Call .FixedSet(中中, 左中, あり, 20, 2, "手配先名(社内)")
        Call .FixedSet(中中, 中中, あり, 2, 3, "受")
        Call .FixedSet(中中, 左中, あり, 25, 4, "品　番")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
        'Call .FixedSet(中中, 左中, あり, 6, 5, "背番号")
        Call .FixedSet(中中, 左中, あり, 8, 5, "背番号")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
        Call .FixedSet(中中, 左中, あり, 10, 6, "収容器")
        Call .FixedSet(中中, 右中, あり, 7, 7, "収容数")
        Call .FixedSet(中中, 右中, あり, 0, 8, "発行枚数")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 9, "編集前品番")
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
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
    
    CisVsGrid.Clear

' 抽出処理 初期ｾｯﾄ
    gSL_Select = "select KT.取引先,KT.受入,KT.表示品番,KT.背番号,KT.収容器,KT.品番,"
    gSL_Select = gSL_Select & "KT.収容数,Count(KT.発行NO) 件数,"
    gSL_Select = gSL_Select & "    TR.略称 取引先名1"
    gSL_Select = gSL_Select & "       from かんばん発行テーブル KT"
    gSL_Select = gSL_Select & "    Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "    On TR.取引先CD = KT.取引先"
    gSL_Select = gSL_Select & "    And TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "    Where KT.かんばん種類 = 1"
    gSL_Select = gSL_Select & "      And KT.発行済 = 0"
    gSL_Select = gSL_Select & "    Group by KT.取引先,KT.受入,KT.表示品番,KT.背番号,KT.収容器,KT.品番,"
    gSL_Select = gSL_Select & "             KT.収容数,TR.略称"
    gSL_Select = gSL_Select & "    Order by MAX(KT.作成日) Desc"
    If Not KHTRead(gSL_Select) Then GoTo GridSet_Ed
    
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        If CisDB.RecordCount + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = CisDB.RecordCount + 1
        End If
    End With
    ID = 0: Ttl_Kensu = 0
    Do Until Not KHT_RDSTS
        With vsGrid1
            ID = ID + 1
            '№
            .TextMatrix(ID, 0) = Format(ID, "####")
            .TextMatrix(ID, 1) = KHT.取引先
            .TextMatrix(ID, 2) = KHT.取引先名1
            .TextMatrix(ID, 3) = KHT.受入
            .TextMatrix(ID, 4) = KHT.表示品番
            .TextMatrix(ID, 5) = KHT.背番号
            .TextMatrix(ID, 6) = KHT.収容器
            .TextMatrix(ID, 7) = Format(KHT.収容数, "#,###")
            .TextMatrix(ID, 8) = Format(KHT.件数, "###")
            .TextMatrix(ID, 9) = KHT.品番
            Ttl_Kensu = Ttl_Kensu + KHT.件数
       End With
       Call KHTReadNext
    Loop
    With vsGrid1
         iT1_Addsu = Format(Ttl_Kensu, "#,###")
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
GridSet_Ed:
    Call KHTClose
End Sub
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit2()
    With CisVsGrid2
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 0, 0, "№")
        Call .FixedSet(中中, 左中, あり, 14, 1, "作　業")
        Call .FixedSet(中中, 左中, あり, 14, 2, "備　考")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 3, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2")
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet2()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
    CisVsGrid2.Clear
    
    gSL_Select = "Select RTRIM(ISNULL(SY.値名称,'')) 作業,備考"
    gSL_Select = gSL_Select & " FROM 品番取引作業マスタ HS"
    gSL_Select = gSL_Select & "   Left Outer Join 名称マスタ SY"
    gSL_Select = gSL_Select & "    On  SY.区分名称 = '作業管理'"
    gSL_Select = gSL_Select & "    And SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "    And SY.値 = HS.作業"
    gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(HTM.品番) & "'"
    gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(HTM.取引先) & "'"
    gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(HTM.受入) & "'"
    gSL_Select = gSL_Select & " Order BY 作業順"
    If Not HSMRead(gSL_Select) Then
       Call HSMClose
       Exit Sub
    End If
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid2
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid2.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid2.DispMax Then
            .Rows = CisVsGrid2.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not HSM_RDSTS
        With vsGrid2
            ID = ID + 1
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            '項目1
            .TextMatrix(ID, GridCol(1)) = HSM.作業
            .TextMatrix(ID, GridCol(2)) = HSM.備考
        End With
        Call HSMReadNext
    Loop
    Call HSMClose
    With vsGrid2
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
Private Sub vsGrid1_DblClick()
    If ProcHB = "B2" Then
       If Trim(vsGrid1.TextMatrix(vsGrid1.Row, 1)) = "" Then Exit Sub
       Call ReturnPress
    End If
End Sub

