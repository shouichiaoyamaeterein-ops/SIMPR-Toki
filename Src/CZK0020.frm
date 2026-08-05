VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BA5D7819-2B67-11D4-806A-00000EA57F9E}#3.0#0"; "CISYMD~1.OCX"
Begin VB.Form CZK0020 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "棚卸プレシート発行"
   ClientHeight    =   11190
   ClientLeft      =   3720
   ClientTop       =   5430
   ClientWidth     =   15390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11190
   ScaleWidth      =   15390
   Begin VB.PictureBox Csv_Bar 
      Height          =   345
      Left            =   4365
      ScaleHeight     =   285
      ScaleWidth      =   6600
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   6660
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   585
      Left            =   5640
      Top             =   8865
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1032
      Caption         =   "終  了"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   4
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   5625
      Left            =   4365
      ScaleHeight     =   5792.151
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6585
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2760
      Width           =   6645
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   5460
         Index           =   1
         Left            =   55
         Top             =   60
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   9631
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Height          =   540
            Left            =   1965
            TabIndex        =   26
            Top             =   1965
            Width           =   4425
            Begin VB.OptionButton H1Op_Buhin 
               BackColor       =   &H00C0C0C0&
               Caption         =   "完成品"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   945
               TabIndex        =   4
               Top             =   165
               Width           =   1035
            End
            Begin VB.OptionButton H1Op_Buhin 
               BackColor       =   &H00C0C0C0&
               Caption         =   "購入品"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   2160
               TabIndex        =   5
               Top             =   165
               Width           =   1035
            End
            Begin VB.OptionButton H1Op_Buhin 
               BackColor       =   &H00C0C0C0&
               Caption         =   "仕掛品"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   3330
               TabIndex        =   6
               Top             =   165
               Width           =   1035
            End
            Begin VB.OptionButton H1Op_Buhin 
               BackColor       =   &H00C0FFC0&
               Caption         =   "全て"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
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
               TabIndex        =   3
               Top             =   165
               Value           =   -1  'True
               Width           =   765
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0C0&
            Height          =   525
            Left            =   1965
            TabIndex        =   25
            Top             =   2625
            Width           =   3015
            Begin VB.OptionButton H1Op_Gass 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ガス"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   2175
               TabIndex        =   9
               Top             =   150
               Width           =   795
            End
            Begin VB.OptionButton H1Op_Gass 
               BackColor       =   &H00C0C0C0&
               Caption         =   "自動車"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   960
               TabIndex        =   8
               Top             =   150
               Width           =   1020
            End
            Begin VB.OptionButton H1Op_Gass 
               BackColor       =   &H00C0FFC0&
               Caption         =   "全て"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
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
               TabIndex        =   7
               Top             =   150
               Value           =   -1  'True
               Width           =   765
            End
         End
         Begin CisYMD_V60.CisYMD iH1_SYMD 
            Height          =   375
            Left            =   1965
            TabIndex        =   12
            Top             =   3930
            Width           =   1740
            _ExtentX        =   3069
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
            Object.Height          =   375
            cSize           =   -1  'True
            cChkResult      =   0   'False
         End
         Begin CisYM_V60.CisYM iH1_TYM 
            Height          =   375
            Left            =   1965
            TabIndex        =   1
            Top             =   870
            Width           =   1470
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
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "　出　力　先"
            Height          =   600
            Left            =   60
            TabIndex        =   24
            Top             =   4395
            Width           =   6345
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "プリンター"
               Height          =   240
               Index           =   0
               Left            =   1065
               TabIndex        =   14
               Top             =   270
               Value           =   -1  'True
               Width           =   1260
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "画面(ﾌﾟﾚﾋﾞｭｰ)"
               Height          =   240
               Index           =   1
               Left            =   2505
               TabIndex        =   15
               Top             =   270
               Width           =   1875
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "CSV"
               Height          =   240
               Index           =   2
               Left            =   4380
               TabIndex        =   16
               Top             =   240
               Width           =   975
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   375
            Left            =   1965
            Top             =   3345
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   661
            BackColor       =   14737632
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
            Begin VB.OptionButton H1Op_Zaiko 
               BackColor       =   &H00E0E0E0&
               Caption         =   "表示無"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   1260
               TabIndex        =   11
               Top             =   60
               Width           =   1020
            End
            Begin VB.OptionButton H1Op_Zaiko 
               BackColor       =   &H00C0FFC0&
               Caption         =   "表示有"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   60
               TabIndex        =   10
               Top             =   60
               Value           =   -1  'True
               Width           =   1020
            End
         End
         Begin CisText_V60.CisText iH1_ZaikoBasyo 
            Height          =   360
            Left            =   1965
            TabIndex        =   2
            Top             =   1515
            Width           =   435
            _ExtentX        =   767
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
         Begin Cis3D_v60.CIS3D H1lb_Zaikonm 
            Height          =   360
            Left            =   2730
            Top             =   1515
            Width           =   2940
            _ExtentX        =   5186
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D UC_3D28 
            Height          =   615
            Index           =   1
            Left            =   60
            Top             =   60
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   1085
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
            Begin VB.PictureBox B1_Area 
               BackColor       =   &H00FFC0FF&
               Height          =   555
               Left            =   45
               ScaleHeight     =   495
               ScaleWidth      =   6210
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   30
               Width           =   6270
               Begin Cis3D_v60.CIS3D lb_msg 
                  Height          =   405
                  Left            =   60
                  Top             =   60
                  Width           =   6090
                  _ExtentX        =   10742
                  _ExtentY        =   714
                  BackColor       =   16761087
                  Caption         =   "棚卸プレシートの発行を行います"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   14.25
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cBoderColor1    =   16761087
                  cBoderColor2    =   16761087
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   6
                  cBoderStyle     =   2
               End
            End
         End
         Begin Cis3D_v60.CIS3D lb_PrtDev 
            Height          =   375
            Left            =   75
            Top             =   5025
            Width           =   5970
            _ExtentX        =   10530
            _ExtentY        =   661
            BackColor       =   16777152
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
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D PB_Print 
            Height          =   345
            Left            =   6075
            Top             =   5055
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   609
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
            cPositionY      =   -90
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   360
            Index           =   1
            Left            =   405
            Top             =   900
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   635
            ForeColor       =   16711680
            Caption         =   "対象年月"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   300
            Left            =   540
            Top             =   1530
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "在庫場所"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   300
            Left            =   540
            Top             =   3360
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "理論在庫"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   420
            Left            =   540
            Top             =   3915
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   741
            ForeColor       =   16711680
            Caption         =   "最終出庫日"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D PB_ZaiBasyo 
            Height          =   360
            Left            =   2385
            Top             =   1515
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            ForeColor       =   65280
            Caption         =   "▼"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cButton         =   -1  'True
            cPositionX      =   20
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   300
            Index           =   0
            Left            =   540
            Top             =   2130
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "部品区分"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   300
            Left            =   540
            Top             =   2775
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "自動車/ガス"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   360
            Left            =   3810
            Top             =   3930
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
            Caption         =   "～"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
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
         Begin CisYMD_V60.CisYMD iH1_EYMD 
            Height          =   375
            Left            =   4215
            TabIndex        =   13
            Top             =   3930
            Width           =   1875
            _ExtentX        =   3069
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
            Object.Height          =   375
            Object.Width           =   1875
            cSize           =   -1  'True
            cChkResult      =   0   'False
         End
      End
   End
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 棚卸プレシート発行 】"
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
         Left            =   14400
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
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
         Height          =   285
         Left            =   11850
         Top             =   30
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
      Left            =   -15
      TabIndex        =   0
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
   Begin Cis3D_v60.CIS3D PB_Ent 
      Height          =   585
      Left            =   8295
      Top             =   8865
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1032
      Caption         =   "開  始"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   4
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D PB_Ok 
      Height          =   585
      Left            =   6930
      Top             =   8865
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1032
      Caption         =   "確　認"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   4
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4365
      Top             =   5130
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   1410
      Top             =   1965
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label PNL_Ok 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6915
      TabIndex        =   23
      Top             =   9540
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8295
      TabIndex        =   21
      Top             =   9540
      Width           =   1485
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5625
      TabIndex        =   20
      Top             =   9540
      Width           =   1485
   End
   Begin VB.Label BG_Logo 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  '透明
      Caption         =   "SIMPR-A"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   19
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CZK0020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  棚卸プレシート 発行
'**       フォームID    :  CZK0020
'**       処理概要      :
'**
'**       作  成  日    :  2009/12/15
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
'
    Dim mSaveMsg        As String       ' ﾒｯｾｰｼﾞ
    Dim mSavePrint      As String
    Dim mPrtDev         As String       ' ﾌﾟﾘﾝﾀ出力先
    Dim mCsvKey         As String
    Dim mCsvDir         As String
    Dim mCsvName        As String

'   背景色
    Dim mSelColor       As Long
    Dim mUnSelColor     As Long
    Dim mUnZaikoColor   As Long
    
    Dim wStr            As String
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
    mSaveMsg = lb_msg
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
    Timer2.Interval = 1000   ' 時刻表示間隔
    Timer2.Enabled = True    ' 時刻表示開始

    With CisFrm
        Set .MSGObject = lb_msg
        Set .MSGTimer = Timer2
        .MSGStopCount = 10
        .MSGType = Left_To_Right
        .MSGText = mSaveMsg
        .MSGInterval = 100
        .MSGRelate
    End With

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

'   背景色の退避
    mSelColor = Op_Out(0).BackColor
    mUnSelColor = Op_Out(1).BackColor
    mUnZaikoColor = H1Op_Zaiko(1).BackColor

    If gLogoDisp Then
       BG_Logo(0).Caption = gLogo
       CisFrm.BGLDisp BG_Logo, 10
    End If

    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

    H1Op_Zaiko(0).Value = True
    H1Op_Buhin(0).Value = True
    H1Op_Gass(0).Value = True
    Op_Out(0).Value = True
    
'   CSVﾌｧｲﾙ出力先
    mCsvName = Mid(CisFun.WinRoot, 1, 3) & "棚卸プレシート.csv"
    mCsvKey = "CZK0020"
    mCsvDir = GetSetting(appname:=gRegKey, section:="CSV", Key:=mCsvKey, Default:=mCsvName)
'   ﾌﾟﾘﾝﾀ出力先
    Op_Out(0).Value = True
    mPrtDev = Printer.DeviceName
    lb_PrtDev = mPrtDev
    mSavePrint = mPrtDev

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
        Case vbKeyF7:       If Dummy.Tag <> "" Then Call PB_Look_Click      '【検索】
        Case vbKeyF12
                            If PB_END.Visible Then Call PB_END_Click        '【終了】
                            If PB_Ok.Visible Then Call PB_END_Click         '【終了】
        Case vbKeyReturn
                            If PB_Ent.Visible Then                          '【入力】
    '                            Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                Call PB_ENT_Click
    '                            Dummy.Enabled = False
                            End If
        Case Else:          Exit Function
    End Select

    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      フォーム    《Ｐａｉｎｔ イベント》                   +
'+----------------------------------------------------------+
Private Sub Form_Paint()
'ｸﾞﾗﾃﾞｰｼｮﾝ表示
    With CisFrm
        .Gradation_Direction = Landscape
        .Gradation_Shading = Lighten
        .Gradation 100, 130, 200, 4, 2, 3
    End With
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Call CisDB.DBDISConnect
    Call CisFrm.UnLoadDisp

    CisFrm.PrintSet mSavePrint

    End
End Sub
'+----------------------------+
'+    確　認  キ ー(F12)       +
'+----------------------------+
Private Sub PB_OK_Click()
    Call PB_END_Click
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
'+      検         索         +
'+----------------------------+
Private Sub PB_Look_Click()
    If Dummy.Tag = "iH1_ZaikoBasyo" Then Call PB_ZaiBasyo_Click: Exit Sub
End Sub

'+----------------------------+
'+       名称マスタ検索        +
'+----------------------------+
Private Sub PB_ZaiBasyo_Click()

    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "管理場所"
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        iH1_ZaikoBasyo = RV_Code
        Call iH1_ZaikoBasyo_LostFocus
    End If

End Sub
'+----------------------------------------+
'+  H1Op_Zaiko   在庫数表示
'+----------------------------------------+
Private Sub H1Op_Zaiko_GotFocus(Index As Integer)
    H1Op_Zaiko(0).Tag = Index
    For Each gObj In H1Op_Zaiko
        If gObj.Index = Index Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnZaikoColor
        End If
    Next gObj
End Sub
'+----------------------------------------+
'+  H1Op_Buhin   部品区分
'+----------------------------------------+
Private Sub H1Op_Buhin_GotFocus(Index As Integer)
    H1Op_Buhin(0).Tag = Index
    For Each gObj In H1Op_Buhin
        If gObj.Index = Index Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj
End Sub
'+----------------------------------------+
'+  H1Op_Gass   自動車/ガス
'+----------------------------------------+
Private Sub H1Op_Gass_GotFocus(Index As Integer)
    H1Op_Gass(0).Tag = Index
    For Each gObj In H1Op_Gass
        If gObj.Index = Index Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj
End Sub
'+----------------------------------------+
'+  Op_Out   出力先
'+----------------------------------------+
Private Sub Op_Out_GotFocus(Index As Integer)
    Op_Out(0).Tag = Index
    For Each gObj In Op_Out
        If gObj.Index = Index Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj
    Select Case Index
        Case 0: lb_PrtDev = mPrtDev
        Case 1: lb_PrtDev = mPrtDev
        Case 2: lb_PrtDev = mCsvDir
        Case Else
            lb_PrtDev = ""
    End Select
End Sub
'+--------------------------------+
'+    iH1_ZaikoBasyo(GotFocus)    +
'+--------------------------------+
Private Sub iH1_ZaikoBasyo_GotFocus()
    Dummy.Tag = ActiveControl.Name
End Sub
'+---------------------------------+
'+    iH1_ZaikoBasyo(LostFocus)    +
'+---------------------------------+
Private Sub iH1_ZaikoBasyo_LostFocus()
    Dummy.Tag = ""
    H1lb_Zaikonm = ""
    If RTrim(iH1_ZaikoBasyo) <> "" Then
       Call MeisyoGet("管理場所", RTrim(iH1_ZaikoBasyo))
       H1lb_Zaikonm = RTrim(SYM_Meisyo)
    End If
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*      メッセージ表示        *
'*******************************
Private Sub Timer2_Timer()
    CisFrm.MSGRelate
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
    Dim B2Mode      As Boolean

    H1Mode = False
    B2Mode = False
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1":   H1Mode = True
       Case "B2":   B2Mode = True
    End Select
    H1_Area1.Enabled = H1Mode

    PB_Ent.Visible = H1Mode
    PNL_ENT.Visible = H1Mode
    PB_END.Visible = H1Mode
    PNL_END.Visible = H1Mode
    Timer2.Enabled = H1Mode Or B2Mode
    PB_Ok.Visible = B2Mode
    PNL_Ok.Visible = B2Mode

    Csv_Bar.Visible = False

' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'*********************************
'*          入力処理(Enter)      *
'*********************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)

    If Not Head1Chk Then GoTo ReturnPress_Ed

    Call DispChange("B1")

    If Not Create_Pure Then GoTo ReturnPress_Ee

    If Not PrintStart Then GoTo ReturnPress_Ee

    lb_msg.Caption = "【  出  力  終  了  】"
    DoEvents

    Call CisFrm.MousePT(1)
    With CisFrm
        .MSGType = FSize_Changes
        .MSGText = "【  出  力  終  了  】"
    End With
    ProcHB = "B2"
    Call DispChange(ProcHB)
    Exit Sub
ReturnPress_Ee:
    lb_msg = mSaveMsg
    DoEvents
    ProcHB = "H1"
    Call DispChange(ProcHB)
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False

'   対象年月
    If Trim(iH1_TYM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    対象年月を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_TYM.SetFocus
        Exit Function
    End If
    If Not iH1_TYM.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    対象年月を正しく入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_TYM.SetFocus
        Exit Function
    End If
    
    If Not Trim(iH1_SYMD) = "" Then
        If Not iH1_SYMD.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    最終出庫日（開始）を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_SYMD.SetFocus
            Exit Function
        End If
    End If
    
    If Not Trim(iH1_EYMD) = "" Then
        If Not iH1_EYMD.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    最終出庫日（終了）を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_EYMD.SetFocus
            Exit Function
        End If
    End If
    
    DoEvents
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "   　棚卸プレシート発行を行います。            "
        .MB_MSG(3) = "   　よろしいですか？            "
        .MB_Title = "棚卸プレシート発行"
        .MB_Button = Yes_No
        If Not .MBOX Then Exit Function
    End With
    
'    If H1Op_OutPut(0) Then COK0530_OutPut = 0 Else COK0530_OutPut = 1

    Head1Chk = True
End Function

'+-------------------------------+
'+     プリンタ・CSV出力先 変更    +
'+-------------------------------+
Private Sub PB_Print_Click()
    If Op_Out(2).Value Then
        ' CSV出力先変更
        On Error Resume Next
        CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly
        CommonDialog1.FileName = mCsvDir
        CommonDialog1.Filter = "ＣＳＶ ファイル (*.csv)|*.csv"
        ' 以下のｿｰｽｺｰﾄﾞは、<ｷｬﾝｾﾙ>ﾎﾞﾀﾝがｸﾘｯｸされたときに実行されます。
        CommonDialog1.CancelError = True
        CommonDialog1.ShowOpen
        If Err = 32755 Then
            '<ｷｬﾝｾﾙ>ﾎﾞﾀﾝが押されたときにｴﾗｰｺｰﾄﾞを返すようにする設定です。
            '必要なｿｰｽｺｰﾄﾞを入力してください。
        End If
        If Trim(CommonDialog1.FileName) <> "" Then
            mCsvDir = CommonDialog1.FileName
            lb_PrtDev.Caption = mCsvDir
        End If
        On Error GoTo 0
    Else
    'If Op_Out(0).Value Or Op_Out(1).Value Then
        ' プリンタ種類変更
         On Error Resume Next
         CommonDialog1.Flags = 64 + cdlPDHidePrintToFile
         CommonDialog1.Max = 0
         CommonDialog1.Min = 0
         CommonDialog1.Copies = 1
         CommonDialog1.ToPage = 1
         CommonDialog1.FromPage = 1
        ' 以下のｿｰｽｺｰﾄﾞは、<ｷｬﾝｾﾙ>ﾎﾞﾀﾝがｸﾘｯｸされたときに実行されます。
         CommonDialog1.CancelError = True
         CommonDialog1.ShowPrinter
         If Err = 32755 Then
        '<ｷｬﾝｾﾙ>ﾎﾞﾀﾝが押されたときにｴﾗｰｺｰﾄﾞを返すようにする設定です。
        '必要なｿｰｽｺｰﾄﾞを入力してください。
         Else
             mPrtDev = Printer.DeviceName
             lb_PrtDev.Caption = mPrtDev
         End If
         On Error GoTo 0
    End If
End Sub
'+---------------------------------------------+
'+      印刷処理                               +
'+---------------------------------------------+
Private Function PrintStart() As Boolean
    PrintStart = False
    If Op_Out(0).Value Or Op_Out(1).Value Then
        lb_msg.Caption = "【 棚卸プレシート 印刷中 】"
        DoEvents
        With CisPrtCZK0020
            .cDataMaxCnt = CisDB.RecordCount
            Set .iPrtForm = CZK0020L
            .dpMSG = "棚卸プレシート　印刷中"
            If Op_Out(0).Value Then
                .iPreview = pv_Print
            Else
                .iPreview = pv_Preview
            End If
            .PrintStart
        End With
    Else
        If Not CsvPut Then Exit Function
    End If
    PrintStart = True
End Function
'****************************
'*       CSV出力処理         *
'****************************
Private Function CsvPut() As Boolean
    Dim OutFNo      As Long
    Dim OutStr      As String
    Dim OutCnt      As Long

    CsvPut = False

    On Error GoTo CsvPut_Err

'   出力ﾌｧｲﾙ存在確認
    If Dir(mCsvDir) <> "" Then
        DoEvents
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "   　" & mCsvDir & "　ﾌｧｲﾙは既に存在します。            "
            .MB_MSG(3) = "   　" & "上書きしますか？            "
            .MB_Title = "上書き確認"
            .MB_Button = Yes_No
            If Not .MBOX Then Exit Function
        End With
    End If

    lb_msg.Caption = "【 棚卸プレシートデータ 出力中 】"
    Csv_Bar.Visible = True

    DoEvents

    Call CisFrm.MousePT(11)

    With CisFrm
        Set .MeterPic = Csv_Bar
        .MeterMax = CisDB.RecordCount
    End With

'   >> 出力ﾃｷｽﾄ Open
    OutFNo = FreeFile
    Open mCsvDir For Output As #OutFNo
'   >> Head部 編集＆出力
    Call CsvEdit_Head1(OutStr)
    Print #OutFNo, OutStr
    Call CsvEdit_Head2(OutStr)
    Print #OutFNo, OutStr
'   >> ｸﾘｱ
    OutCnt = 0
    Do Until Not TPH_RDSTS
    '   >> ｶｳﾝﾄ
        OutCnt = OutCnt + 1
    '   >> Body部 編集＆出力
        Call CsvEdit_Body1(OutStr, OutCnt)
        Print #OutFNo, OutStr

        CisFrm.MeterUpdate OutCnt

    '   >> 次のﾚｺｰﾄﾞを読み込む
        Call TPHReadNext(1)
    Loop
    Call TPHClose(1)
'   >> 出力ﾃｷｽﾄ Close
    Close #OutFNo

    SaveSetting appname:=gRegKey, section:="CSV", Key:=mCsvKey, setting:=RTrim(mCsvDir)

    Call CisFrm.MousePT(1)

    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "    【  CSV出力 正常終了 】            "
        .MB_MSG(3) = "    出 力 先  = " & RTrim(mCsvDir) & "　　　　　"
        .MB_MSG(4) = "    出力件数 = " & Format(OutCnt, "#,##0") & " 件"
        .MB_Title = "【確認】"
        .MB_Button = OK
        .MBOX
    End With

    CsvPut = True
    Exit Function
CsvPut_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------------+
'+    Head部　編集処理
'+----------------------------+
Private Sub CsvEdit_Head1(CsvStr As String)
    CsvStr = ""
    CsvStr = CsvStr & Mid(RTrim(iH1_TYM), 1, 4) & "年" & Mid(RTrim(iH1_TYM), 5, 2) & "月" & " 棚卸プレシート"
End Sub
Private Sub CsvEdit_Head2(CsvStr As String)
    CsvStr = ""
    CsvStr = CsvStr & "№"
    CsvStr = CsvStr & ",在庫場所１"
    CsvStr = CsvStr & ",在庫場所２"
    CsvStr = CsvStr & ",部品区分１"
    CsvStr = CsvStr & ",部品区分２"
    CsvStr = CsvStr & ",伝票№"
    CsvStr = CsvStr & ",伝票行№"
    CsvStr = CsvStr & ",棚番"
    CsvStr = CsvStr & ",背番号"
    CsvStr = CsvStr & ",品番"
    CsvStr = CsvStr & ",理論在庫数"
    CsvStr = CsvStr & ",棚卸数"
    CsvStr = CsvStr & ",品名"
End Sub
'+----------------------------+
'+    Body部　編集処理
'+----------------------------+
Private Sub CsvEdit_Body1(CsvStr As String, CsvCnt As Long)
    With TPH
        CsvStr = ""
        CsvStr = CsvStr & CsvCnt
        CsvStr = CsvStr & "," & RTrim(.在庫場所名1)
        CsvStr = CsvStr & "," & RTrim(.在庫場所名2)
        If .部品区分 = 0 Then
            CsvStr = CsvStr & ",完成品"
        Else
            If .部品区分 = 1 Then
                CsvStr = CsvStr & ",購入品"
            Else
                If .部品区分 = 2 Then
                    CsvStr = CsvStr & ",仕掛品"
                Else
                    CsvStr = CsvStr & ","
                End If
            End If
        End If
        If .自動車ガス = 0 Then
            CsvStr = CsvStr & ",自動車"
        Else
            If .自動車ガス = 1 Then
                CsvStr = CsvStr & ",ガス"
            Else
                CsvStr = CsvStr & ","
            End If
        End If
            
        CsvStr = CsvStr & "," & Format(.伝票NO, "#########")
        CsvStr = CsvStr & "," & Format(.伝票行NO, "##")
        CsvStr = CsvStr & "," & RTrim(.棚番)
        CsvStr = CsvStr & "," & RTrim(.背番号)
        CsvStr = CsvStr & "," & RTrim(.品番)
        CsvStr = CsvStr & "," & Format(.在庫数, "###0")
        CsvStr = CsvStr & ","
        CsvStr = CsvStr & "," & RTrim(.品名)
    End With
End Sub
'/////////////////////////////////////////////////////
'+---------------------------------------------------+
'+          棚卸プレシート発行データ作成
'+---------------------------------------------------+
'/////////////////////////////////////////////////////
Private Function Create_Pure() As Boolean
    Create_Pure = False
    
    lb_msg.Caption = "【 棚卸プレシートデータ 抽出中 】"
    DoEvents
    
'''    Call CisDB.DBTran(TransBegin)
    With CisDB
        .SQL = "棚卸プレシート発行データ作成"
        .StoadoCount = 6
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = Input用
        .ParaNo = 4:    .ParaIO = Input用
        .ParaNo = 5:    .ParaIO = Input用
        .ParaNo = 6:    .ParaIO = OutPut用

        .ParaNo = 1:    .ParaValue = iH1_TYM                        ' 対象年月
        .ParaNo = 2:    .ParaValue = iH1_ZaikoBasyo                 ' 在庫場所
        If H1Op_Zaiko(0).Value Then                                 ' 理論在庫表示
            .ParaNo = 3:    .ParaValue = 0
        Else
            .ParaNo = 3:    .ParaValue = 1
        End If
                    
        .ParaNo = 4:    .ParaValue = iH1_SYMD                       ' 最終出庫日（開始）
        .ParaNo = 5:    .ParaValue = iH1_EYMD                       ' 最終出庫日（終了）

        .DBStored

        .ParaNo = 0
        If .ParaValue <> 0 Then
            CisFun.MB_Lines = 5
            CisFun.MB_MSG(2) = "    棚卸プレシート発行データ作成　異常！      "
            CisFun.MB_MSG(4) = "    Error No = " & .ParaValue
            CisFun.MB_Button = Error
            CisFun.MBOX
            End
        End If

        '.ParaNo = 5: C1lb_Input(1) = Format(CisFun.Val2(C1lb_Input(1)) + CisFun.Val2(.ParaValue), "#,###")
    End With
'''    Call CisDB.DBTran(TransCommit)
    
    If iH1_ZaikoBasyo = "" Then
        wStr = "%"
    Else
        wStr = iH1_ZaikoBasyo
    End If
    
    '   < 印刷データ読み込み >
    gSL_Select = "SELECT TPH.*,ISNULL(MS.値名称,'') 在庫場所名1,ISNULL(MT.値名称,'') 在庫場所名2" & Chr(13)
    gSL_Select = gSL_Select & "      ,ISNULL(HM.表示品番,'') 表示品番,ISNULL(HM.品名,'') 品名" & Chr(13)
    gSL_Select = gSL_Select & "  FROM 棚卸プレシート発行テーブル TPH" & Chr(13)
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 名称マスタ MS" & Chr(13)
    gSL_Select = gSL_Select & "    ON MS.区分名称 = '管理場所'" & Chr(13)
    gSL_Select = gSL_Select & "   AND MS.値 = TPH.在庫場所1" & Chr(13)
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 名称マスタ MT" & Chr(13)
    gSL_Select = gSL_Select & "    ON MT.区分名称 = '管理場所ユーザー'" & Chr(13)
    gSL_Select = gSL_Select & "   AND MT.値 = TPH.在庫場所2" & Chr(13)
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 品番マスタ HM" & Chr(13)
    gSL_Select = gSL_Select & "    ON HM.品番 = TPH.品番" & Chr(13)
    gSL_Select = gSL_Select & " WHERE TPH.発行年月 = '" & RTrim(iH1_TYM) & "'" & Chr(13)
    gSL_Select = gSL_Select & "   AND TPH.在庫場所1 LIKE '" & RTrim(wStr) & "'" & Chr(13)
    '部品区分
    If H1Op_Buhin(1).Value Then
        gSL_Select = gSL_Select & "   AND TPH.部品区分 = 0" & Chr(13)
    Else
        If H1Op_Buhin(2).Value Then
            gSL_Select = gSL_Select & "   AND TPH.部品区分 = 1" & Chr(13)
        Else
            If H1Op_Buhin(3).Value Then
                gSL_Select = gSL_Select & "   AND TPH.部品区分 = 2" & Chr(13)
            End If
        End If
    End If
    '自動車/ガス
    If H1Op_Gass(1).Value Then
        gSL_Select = gSL_Select & "   AND TPH.自動車ガス = 0" & Chr(13)
    Else
        If H1Op_Gass(2).Value Then
            gSL_Select = gSL_Select & "   AND TPH.自動車ガス = 1" & Chr(13)
        End If
    End If
    
    gSL_Select = gSL_Select & " ORDER BY TPH.伝票NO,TPH.伝票行NO"
    
    If Not TPHRead(gSL_Select, , 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    該当データがありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Call TPHClose
        Exit Function
    End If
    
    Create_Pure = True
End Function


