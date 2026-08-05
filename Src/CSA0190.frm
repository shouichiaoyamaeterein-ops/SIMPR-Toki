VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form CSA0190 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "負荷リスト発行"
   ClientHeight    =   11190
   ClientLeft      =   1605
   ClientTop       =   1650
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
      Top             =   1410
      Visible         =   0   'False
      Width           =   6660
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   585
      Left            =   5633
      Top             =   9075
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
      Height          =   6690
      Left            =   4365
      ScaleHeight     =   6900.62
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1845
      Width           =   6660
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   6540
         Index           =   1
         Left            =   75
         Top             =   45
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   11536
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
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0C0&
            Height          =   570
            Left            =   1770
            TabIndex        =   27
            Top             =   4260
            Width           =   3300
            Begin VB.OptionButton H1Op_Kosu 
               BackColor       =   &H00C0C0C0&
               Caption         =   "目標工数"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   1560
               TabIndex        =   11
               Top             =   195
               Width           =   1245
            End
            Begin VB.OptionButton H1Op_Kosu 
               BackColor       =   &H00C0FFC0&
               Caption         =   "基準工数"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   165
               TabIndex        =   10
               Top             =   195
               Value           =   -1  'True
               Width           =   1230
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0C0&
            Height          =   570
            Left            =   1770
            TabIndex        =   26
            Top             =   4860
            Width           =   3300
            Begin VB.OptionButton H1Op_OutPut 
               BackColor       =   &H00C0FFC0&
               Caption         =   "一覧表"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   165
               TabIndex        =   12
               Top             =   195
               Value           =   -1  'True
               Width           =   1170
            End
            Begin VB.OptionButton H1Op_OutPut 
               BackColor       =   &H00C0C0C0&
               Caption         =   "明細表"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   1560
               TabIndex        =   13
               Top             =   195
               Width           =   1170
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "　出　力　先"
            Height          =   600
            Left            =   60
            TabIndex        =   25
            Top             =   5445
            Width           =   6345
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
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Height          =   1170
            Left            =   1770
            TabIndex        =   24
            Top             =   3060
            Width           =   3300
            Begin VB.OptionButton H1Op_KMoto 
               BackColor       =   &H00C0C0C0&
               Caption         =   "実績"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   2175
               TabIndex        =   8
               Top             =   195
               Width           =   1065
            End
            Begin VB.OptionButton H1Op_KMoto 
               BackColor       =   &H00C0C0C0&
               Caption         =   "出荷"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   1200
               TabIndex        =   7
               Top             =   195
               Width           =   1170
            End
            Begin VB.OptionButton H1Op_KMoto 
               BackColor       =   &H00C0FFC0&
               Caption         =   "内示"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   165
               TabIndex        =   6
               Top             =   195
               Value           =   -1  'True
               Width           =   1170
            End
            Begin Cis3D_v60.CIS3D lbl_Naiji 
               Height          =   360
               Left            =   105
               Top             =   660
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   635
               ForeColor       =   16711680
               Caption         =   "内示年月"
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
            Begin CisYM_V60.CisYM iH1_YM 
               Height          =   360
               Left            =   1500
               TabIndex        =   9
               Top             =   675
               Width           =   1560
               _ExtentX        =   2646
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
               cContent        =   1
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   315
            Left            =   3690
            Top             =   885
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "～"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   18
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
         Begin CisText_V60.CisText iH1_Bumon 
            Height          =   360
            Left            =   1800
            TabIndex        =   3
            Top             =   1440
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXXX"
            MaxLength       =   7
         End
         Begin Cis3D_v60.CIS3D H1lb_Bumonnm 
            Height          =   360
            Left            =   3120
            Top             =   1440
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
                  Caption         =   "負荷リスト発行を行います"
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
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   360
            Index           =   1
            Left            =   405
            Top             =   900
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   635
            ForeColor       =   16711680
            Caption         =   "日付範囲"
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
            Top             =   1455
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "部　門"
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
         Begin Cis3D_v60.CIS3D PB_Bumon 
            Height          =   360
            Left            =   2775
            Top             =   1440
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_SDate 
            Height          =   330
            Left            =   1800
            TabIndex        =   1
            Top             =   870
            Width           =   1695
            _ExtentX        =   2990
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
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_EDate 
            Height          =   345
            Left            =   4305
            TabIndex        =   2
            Top             =   870
            Width           =   1830
            _ExtentX        =   2990
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
            Object.Width           =   1830
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CisText_V60.CisText iH1_Kikai 
            Height          =   360
            Left            =   1785
            TabIndex        =   4
            Top             =   2025
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXX"
            MaxLength       =   3
         End
         Begin Cis3D_v60.CIS3D H1lb_Kikainm 
            Height          =   360
            Left            =   2595
            Top             =   2025
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
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   300
            Left            =   525
            Top             =   2040
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "機　械"
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
         Begin Cis3D_v60.CIS3D PB_Kikai 
            Height          =   360
            Left            =   2250
            Top             =   2025
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
         Begin CisText_V60.CisText iH1_KikaiNo 
            Height          =   360
            Left            =   1785
            TabIndex        =   5
            Top             =   2625
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXX"
            MaxLength       =   3
         End
         Begin Cis3D_v60.CIS3D H1lb_KikaiNonm 
            Height          =   360
            Left            =   2595
            Top             =   2625
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
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   300
            Left            =   525
            Top             =   2640
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "機 械 №"
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
         Begin Cis3D_v60.CIS3D PB_KikaiNo 
            Height          =   360
            Left            =   2250
            Top             =   2625
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
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   300
            Index           =   1
            Left            =   510
            Top             =   3240
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "計算元"
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
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   300
            Index           =   2
            Left            =   510
            Top             =   5040
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "帳  票"
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
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   300
            Index           =   0
            Left            =   495
            Top             =   4440
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "工  数"
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
         Begin Cis3D_v60.CIS3D lb_PrtDev 
            Height          =   375
            Left            =   75
            Top             =   6075
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
            Top             =   6105
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
      Caption         =   "【 負荷リスト発行 】"
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
      TabIndex        =   17
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
   Begin Cis3D_v60.CIS3D PB_Ent 
      Height          =   585
      Left            =   8288
      Top             =   9075
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
      Left            =   6923
      Top             =   9075
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
      Top             =   3840
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
      Left            =   6908
      TabIndex        =   23
      Top             =   9750
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8288
      TabIndex        =   21
      Top             =   9750
      Width           =   1485
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5618
      TabIndex        =   20
      Top             =   9750
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
Attribute VB_Name = "CSA0190"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  負荷リスト 発行
'**       フォームID    :  CSA0190
'**       処理概要      :
'**
'**       作  成  日    :  2010/01/21
'**       変  更  日    :  2010/11/11  内示年月を指定出来るように変更
'**       変  更  日    :  2011/02/09  CSV出力時の工程数項目を修正
'**       変  更  日    :  2011/12/28  作業日報ﾃｰﾌﾞﾙ参照条件(工順)を変更
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
    Dim mUnKubunColor   As Long
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub form_load()
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
        .MSGTYPE = Left_To_Right
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

'=========================================== 2010/11/11 Start
'年月初期値･･･展開処理最新年月
    gSL_Select = "SELECT * FROM 名称マスタ "
    gSL_Select = gSL_Select & " WHERE 区分名称 = '処理内容'"
    gSL_Select = gSL_Select & " AND   値 = 'PNJ0310'"
    If SYMRead(gSL_Select, 1) Then
       iH1_YM = RTrim(SYM.英数字1)
    Else
       iH1_YM = Format(Now, "YYYYMM")
    End If
'=========================================== 2010/11/11 End

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

    H1Op_KMoto(0).Value = True
    H1Op_Kosu(0).Value = True
    H1Op_OutPut(0).Value = True
    
'   CSVﾌｧｲﾙ出力先
    mCsvName = Mid(CisFun.WinRoot, 1, 3) & "生産負荷リスト.csv"
    mCsvKey = "CSA0190"
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
                            If PB_ENT.Visible Then                          '【入力】
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
'+----------------------------------------+
'+  H1Op_KMoto   計算元
'+----------------------------------------+
Private Sub H1Op_KMoto_GotFocus(INDEX As Integer)
    H1Op_KMoto(0).Tag = INDEX
    For Each gObj In H1Op_KMoto
        If gObj.INDEX = INDEX Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj

    '================================== 2010/11/11 Start
    If INDEX = 0 Then
       lbl_Naiji.Visible = True
       iH1_YM.Visible = True
    Else
       lbl_Naiji.Visible = False
       iH1_YM.Visible = False
    End If
    '================================== 2010/11/11 End
End Sub
'+----------------------------------------+
'+  H1Op_KMoto   工数
'+----------------------------------------+
Private Sub H1Op_Kosu_GotFocus(INDEX As Integer)
    H1Op_Kosu(0).Tag = INDEX
    For Each gObj In H1Op_Kosu
        If gObj.INDEX = INDEX Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj
End Sub
'+----------------------------------------+
'+  H1Op_Output   帳票
'+----------------------------------------+
Private Sub H1Op_Output_GotFocus(INDEX As Integer)
    H1Op_OutPut(0).Tag = INDEX
    For Each gObj In H1Op_OutPut
        If gObj.INDEX = INDEX Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj
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
    If Dummy.Tag = "iH1_Bumon" Then Call PB_Bumon_Click: Exit Sub
    If Dummy.Tag = "iH1_Kikai" Then Call PB_kikai_Click: Exit Sub
    If Dummy.Tag = "iH1_KikaiNo" Then Call PB_KikaiNo_Click: Exit Sub
End Sub
'+---------------------+
'+      部門検索       +
'+---------------------+
Private Sub PB_Bumon_Click()

    RV_Left = 0:    RV_Top = 0
    RV_TorKb = 7: RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
       iH1_Bumon = RTrim(RV_TorcdK)
       Call iH1_Bumon_LostFocus
    End If
    
End Sub
'+----------------------------------------------------------------------------+
'+      機械(名称マスタ) 検索
'+----------------------------------------------------------------------------+
Private Sub PB_kikai_Click()

    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "機械"
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        iH1_Kikai = RV_Code
        Call iH1_Kikai_LostFocus
    End If
End Sub
'+----------------------------------------------------------------------------+
'+      機械No(名称マスタ) 検索
'+----------------------------------------------------------------------------+
Private Sub PB_KikaiNo_Click()

    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "機械No"
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        iH1_KikaiNo = RV_Code
        Call iH1_KikaiNo_LostFocus
    End If
End Sub
'+----------------------------------------+
'+  Op_Out   出力先
'+----------------------------------------+
Private Sub Op_Out_GotFocus(INDEX As Integer)
    Op_Out(0).Tag = INDEX
    For Each gObj In Op_Out
        If gObj.INDEX = INDEX Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj
    Select Case INDEX
        Case 0: lb_PrtDev = mPrtDev
        Case 1: lb_PrtDev = mPrtDev
        Case 2: lb_PrtDev = mCsvDir
        Case Else
            lb_PrtDev = ""
    End Select
End Sub
'+----------------------------+
'+    iH1_Bumon(GotFocus)     +
'+----------------------------+
Private Sub iH1_Bumon_GotFocus()
    Dummy.Tag = ActiveControl.Name
End Sub
'+----------------------------+
'+    iH1_Bumon(LostFocus)    +
'+----------------------------+
Private Sub iH1_Bumon_LostFocus()
    Dummy.Tag = ""
    
    H1lb_Bumonnm = ""
    If TorNmGet(RTrim(iH1_Bumon), 7) Then
        H1lb_Bumonnm = RTrim(TRM_RName)
    End If
    
End Sub
'+-----------------------------+
'+  iH1_Kikai
'+-----------------------------+
Private Sub iH1_Kikai_GotFocus()
    Dummy.Tag = ActiveControl.Name
    Dummy.Visible = True
End Sub
Private Sub iH1_Kikai_LostFocus()
    Dummy.Tag = ""

    H1lb_Kikainm = ""
    If MeisyoGet("機械", RTrim(iH1_Kikai)) Then
        H1lb_Kikainm = RTrim(SYM_Meisyo)
    End If
End Sub
'+-----------------------------+
'+  iH1_KikaiNo
'+-----------------------------+
Private Sub iH1_KikaiNo_GotFocus()
    Dummy.Tag = ActiveControl.Name
    Dummy.Visible = True
End Sub
Private Sub iH1_KikaiNo_LostFocus()
    Dummy.Tag = ""

    H1lb_KikaiNonm = ""
    If MeisyoGet("機械NO", RTrim(iH1_KikaiNo)) Then
        H1lb_KikaiNonm = RTrim(SYM_Meisyo)
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

    PB_ENT.Visible = H1Mode
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
    
    If Not Create_List Then GoTo ReturnPress_Ee

    If Not PrintStart Then GoTo ReturnPress_Ee

    lb_msg.Caption = "【  出  力  終  了  】"
    DoEvents

    Call CisFrm.MousePT(1)
    With CisFrm
        .MSGTYPE = FSize_Changes
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
    Dim WorkMSG As String
    
    Head1Chk = False
    WorkMSG = ""
    
'   日付範囲（開始）
    If Trim(iH1_SDate) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    日付範囲（開始）を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SDate.SetFocus
        Exit Function
    End If
    
    If Trim(iH1_SDate.cYear) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    日付範囲（開始 - 年）を入力して下さい          "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SDate.SetFocus
        Exit Function
    End If
    
    If Trim(iH1_SDate.cMonth) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    日付範囲（開始 - 月）を入力して下さい          "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SDate.SetFocus
        Exit Function
    End If
    
    If Trim(iH1_SDate.cDay) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    日付範囲（開始 - 日）を入力して下さい          "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SDate.SetFocus
        Exit Function
    End If
    
    If iH1_SDate.cChkResult = False Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    日付範囲（開始）を正しく入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SDate.SetFocus
        Exit Function
    End If

    '日付範囲（終了）
    If Not Trim(iH1_EDate) = "" Then
    
        If Trim(iH1_EDate.cYear) = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    日付範囲（終了 - 年）を入力して下さい          "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_EDate.SetFocus
            Exit Function
        End If
        
        If Trim(iH1_EDate.cMonth) = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    日付範囲（終了 - 月）を入力して下さい          "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_EDate.SetFocus
            Exit Function
        End If
        
        If Trim(iH1_SDate.cDay) = "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    日付範囲（終了 - 日）を入力して下さい          "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_EDate.SetFocus
            Exit Function
        End If
        
        If iH1_EDate.cChkResult = False Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    日付範囲（終了）を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_EDate.SetFocus
            Exit Function
        End If
        
        If iH1_SDate > iH1_EDate Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    日付範囲の大小関係が正しくありません。            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_EDate.SetFocus
            Exit Function
        End If
    
    End If

'   部門
    If RTrim(iH1_Bumon) <> "" Then
        If Not TorNmGet(RTrim(iH1_Bumon), 6) Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    部　門    未 登 録            "
                .MB_MSG(4) = "     【 取引先マスタ 】           "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Bumon.SetFocus
           Exit Function
        End If
    End If
    
' 機械
    If RTrim(iH1_Kikai) <> "" Then
        If Not MeisyoGet("機械", RTrim(iH1_Kikai)) Then
            With CisFun
                 .MB_Lines = 5
                 .MB_MSG(2) = "      機 械 未 登 録                "
                 .MB_MSG(4) = "      【 名称マスタ 】              "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Kikai.SetFocus
           Exit Function
        End If
    End If
    
' 機械№
    If RTrim(iH1_KikaiNo) <> "" Then
        If Not MeisyoGet("機械NO", RTrim(iH1_KikaiNo)) Then
            With CisFun
                 .MB_Lines = 5
                 .MB_MSG(2) = "     機 械 № 未 登 録                "
                 .MB_MSG(4) = "       【 名称マスタ 】               "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_KikaiNo.SetFocus
           Exit Function
        End If
    End If

    If H1Op_KMoto(0).Value Then
        If H1Op_OutPut(0).Value Then
            WorkMSG = "   　負荷一覧表（内示）の発行を行います。            "
        Else
            If H1Op_OutPut(1).Value Then
                WorkMSG = "   　負荷明細表（内示）の発行を行います。            "
            End If
        End If
    Else
        If H1Op_KMoto(1).Value Then
            If H1Op_OutPut(0).Value Then
                WorkMSG = "   　負荷一覧表（出荷）の発行を行います。            "
            Else
                If H1Op_OutPut(1).Value Then
                    WorkMSG = "   　負荷明細表（出荷）の発行を行います。            "
                End If
            End If
        Else
            If H1Op_KMoto(2).Value Then
                If H1Op_OutPut(0).Value Then
                    WorkMSG = "   　負荷一覧表（実績）の発行を行います。            "
                Else
                    If H1Op_OutPut(1).Value Then
                        WorkMSG = "   　負荷明細表（実績）の発行を行います。            "
                    End If
                End If
            End If
        End If
    End If

    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = WorkMSG
        .MB_MSG(3) = "   　よろしいですか？            "
        .MB_Title = "負荷リスト発行"
        .MB_Button = Yes_No
        If Not .MBOX Then Exit Function
    End With
    
    'DoEvents
    
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
        lb_msg.Caption = "【 負荷リスト 印刷中 】"
        DoEvents
        If H1Op_OutPut(0).Value Then
            With CisPrtCSA0190
                .cDataMaxCnt = CisDB.RecordCount
                Set .iPrtForm = CSA0190L      '一覧表
                .dpMSG = "負荷リスト　印刷中"
                If Op_Out(0).Value Then
                    .iPreview = pv_Print
                Else
                    .iPreview = pv_Preview
                End If
                .PrintStart
            End With
        Else
            If H1Op_OutPut(1).Value Then
'                With CisPrtCSA0190
                With CisPrtCSA0195
                    .cDataMaxCnt = CisDB.RecordCount
                    Set .iPrtForm = CSA0195L  '明細表
                    .dpMSG = "負荷リスト　印刷中"
                    If Op_Out(0).Value Then
                        .iPreview = pv_Print
                    Else
                        .iPreview = pv_Preview
                    End If
                    .PrintStart
                End With
            End If
        End If
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
    Dim WorkMSG     As String

    CsvPut = False
    WorkMSG = ""

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
    
    
     If H1Op_KMoto(0).Value Then
        If H1Op_OutPut(0).Value Then
            WorkMSG = "【 負荷一覧表（内示）データ 出力中 】"
        Else
            If H1Op_OutPut(1).Value Then
                WorkMSG = "【 負荷明細表（内示）データ 出力中 】"
            End If
        End If
    Else
        If H1Op_KMoto(1).Value Then
            If H1Op_OutPut(0).Value Then
                WorkMSG = "【 負荷一覧表（出荷）データ 出力中 】"
            Else
                If H1Op_OutPut(1).Value Then
                    WorkMSG = "【 負荷明細表（出荷）データ 出力中 】"
                End If
            End If
        Else
            If H1Op_KMoto(2).Value Then
                If H1Op_OutPut(0).Value Then
                    WorkMSG = "【 負荷一覧表（実績）データ 出力中 】"
                Else
                    If H1Op_OutPut(1).Value Then
                        WorkMSG = "【 負荷明細表（実績）データ 出力中 】"
                    End If
                End If
            End If
        End If
    End If

    lb_msg.Caption = WorkMSG
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
    Do Until Not HKT_RDSTS
    '   >> ｶｳﾝﾄ
        OutCnt = OutCnt + 1
    '   >> Body部 編集＆出力
        Call CsvEdit_Body1(OutStr, OutCnt)
        Print #OutFNo, OutStr

        CisFrm.MeterUpdate OutCnt

    '   >> 次のﾚｺｰﾄﾞを読み込む
        Call HKTReadNext
    Loop
    Call HKTClose
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
    
    Dim Work As String
    
    Work = ""
    CsvStr = ""
    
    If Trim(iH1_EDate) = "" Then
        Work = Mid(RTrim(iH1_SDate), 1, 4) & "年" & Mid(RTrim(iH1_SDate), 5, 2) & "月" & Mid(RTrim(iH1_SDate), 7, 2) & "日"
    Else
        Work = Mid(RTrim(iH1_SDate), 1, 4) & "年" & Mid(RTrim(iH1_SDate), 5, 2) & "月" & Mid(RTrim(iH1_SDate), 7, 2) & "日 ～ " _
               & Mid(RTrim(iH1_EDate), 1, 4) & "年" & Mid(RTrim(iH1_EDate), 5, 2) & "月" & Mid(RTrim(iH1_EDate), 7, 2) & "日"
    End If
    
    If H1Op_KMoto(0).Value Then
        If H1Op_OutPut(0).Value Then
            CsvStr = CsvStr & Work & " " & "負荷一覧表（内示）"
        Else
            CsvStr = CsvStr & Work & " " & "負荷明細表（内示）"
        End If
    Else
        If H1Op_KMoto(1).Value Then
            If H1Op_OutPut(0).Value Then
                CsvStr = CsvStr & Work & " " & "負荷一覧表（出荷）"
            Else
                CsvStr = CsvStr & Work & " " & "負荷明細表（出荷）"
            End If
        Else
            If H1Op_OutPut(0).Value Then
                CsvStr = CsvStr & Work & " " & "負荷一覧表（実績）"
            Else
                CsvStr = CsvStr & Work & " " & "負荷明細表（実績）"
            End If
        End If
    End If
    
End Sub
Private Sub CsvEdit_Head2(CsvStr As String)
    
    CsvStr = ""
    CsvStr = CsvStr & "部門"
    CsvStr = CsvStr & ",部門名"
    CsvStr = CsvStr & ",機械"
    CsvStr = CsvStr & ",機械名"
    CsvStr = CsvStr & ",区分"
    CsvStr = CsvStr & ",機械№"
    CsvStr = CsvStr & ",機械№名"
    CsvStr = CsvStr & ",工程"
    CsvStr = CsvStr & ",工程名"
    If H1Op_OutPut(1).Value Then
        CsvStr = CsvStr & ",品番"
    End If
    CsvStr = CsvStr & ",数量"
    CsvStr = CsvStr & ",工程数"
    CsvStr = CsvStr & ",基準時間"
    
End Sub
'+----------------------------+
'+    Body部　編集処理
'+----------------------------+
Private Sub CsvEdit_Body1(CsvStr As String, CsvCnt As Long)
    
    CsvStr = ""
    
    With HKT
        
        If Trim(.部門) = "" Then
            CsvStr = ",,"
        Else
            CsvStr = Replace(Trim(.部門), ":", ",")
        End If
        If Trim(.機械) = "" Then
            CsvStr = CsvStr & ",,"
        Else
            CsvStr = CsvStr & "," & Replace(Trim(.機械), ":", ",")
        End If
        CsvStr = CsvStr & "," & RTrim(.区分)
        If Trim(.機械NO) = "" Then
            CsvStr = CsvStr & ",,"
        Else
            CsvStr = CsvStr & "," & Replace(Trim(.機械NO), ":", ",")
        End If
        If Trim(.工程) = "" Then
            CsvStr = CsvStr & ",,"
        Else
            CsvStr = CsvStr & "," & Replace(Trim(.工程), ":", ",")
        End If
        If H1Op_OutPut(1).Value Then
            CsvStr = CsvStr & "," & RTrim(.品番)
        End If
        CsvStr = CsvStr & "," & Format(.数量, "########0")
'        CsvStr = CsvStr & "," & Format(.工程数, "########0")                   '2011/02/09
        CsvStr = CsvStr & "," & Format(.表示工程数, "########0")                '2011/02/09
        If H1Op_Kosu(0).Value Then
            CsvStr = CsvStr & "," & Format(.基準時間, "########0")
        Else
            CsvStr = CsvStr & "," & Format(.目標時間, "########0")
        End If

        CsvStr = Replace(CsvStr, " ", "")
    
    End With
End Sub
'/////////////////////////////////////////////////////
'+---------------------------------------------------+
'+               負荷リストデータ作成
'+---------------------------------------------------+
'/////////////////////////////////////////////////////
Private Function Create_List() As Boolean
    Dim WorkMSG As String
    Dim gsl_Where1 As String
    Dim gsl_Where2 As String
    
    WorkMSG = ""
 
    Create_List = False
    
    
     If H1Op_KMoto(0).Value Then
        If H1Op_OutPut(0).Value Then
            WorkMSG = "【 負荷一覧表（内示）データ 抽出中 】"
        Else
            If H1Op_OutPut(1).Value Then
                WorkMSG = "【 負荷明細表（内示）データ 抽出中 】"
            End If
        End If
    Else
        If H1Op_KMoto(1).Value Then
            If H1Op_OutPut(0).Value Then
                WorkMSG = "【 負荷一覧表（出荷）データ 抽出中 】"
            Else
                If H1Op_OutPut(1).Value Then
                    WorkMSG = "【 負荷明細表（出荷）データ 抽出中 】"
                End If
            End If
        Else
            If H1Op_KMoto(2).Value Then
                If H1Op_OutPut(0).Value Then
                    WorkMSG = "【 負荷一覧表（実績）データ 抽出中 】"
                Else
                    If H1Op_OutPut(1).Value Then
                        WorkMSG = "【 負荷明細表（実績）データ 抽出中 】"
                    End If
                End If
            End If
        End If
    End If
    
    lb_msg.Caption = WorkMSG
    DoEvents
    
    'ワーク削除
    gSL_Select = ""
    gSL_Select = gSL_Select & " delete from 生産負荷リストワーク "
    gSL_Select = gSL_Select & " where 作成端末 = substring(host_name(),1,20) "
    With CisDB
        '.ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    
    '   < 印刷データ読み込み >
    If H1Op_KMoto(0).Value = True Then  '内示
        If H1Op_OutPut(0).Value = True Then  '一覧表
            gSL_Select = ""
            gSL_Select = gSL_Select & " INSERT INTO 生産負荷リストワーク "
            '============================================ 2010/09/16 Start
'            gSL_Select = gSL_Select & " ( 作成端末,部門,機械,区分,機械NO,工程,工程数,基準時間1,基準時間2,数量,表示工程数,基準時間,目標時間 ) "
            gSL_Select = gSL_Select & " ( 作成端末,部門,機械,区分,機械NO,工程,工程数,基準時間1,基準時間2,数量,表示工程数,基準時間,目標時間,使用数1,使用数2,使用数3,使用数4,使用数5 ) "
            '============================================ 2010/09/16 End
            gSL_Select = gSL_Select & " Select substring(host_name(),1,20),Case When ISNULL(HKM.取引先,'') = '' AND ISNULL(TRN.略称,'') = '' Then '' " & Chr(13)
            gSL_Select = gSL_Select & "             Else ISNULL(HKM.取引先,'') + ' : ' + ISNULL(TRN.略称,'') END 部門," & Chr(13)
            gSL_Select = gSL_Select & "        Case When ISNULL(HKM.機械,'') = '' AND ISNULL(MS.値名称,'') = '' Then '' " & Chr(13)
            gSL_Select = gSL_Select & "             Else ISNULL(HKM.機械,'') + ' : ' + ISNULL(MS.値名称,'') END 機械," & Chr(13)
            gSL_Select = gSL_Select & "        Case When HKM.ロボット区分 = 0 then 'ﾛﾎﾞｯﾄ'" & Chr(13)
            gSL_Select = gSL_Select & "             Else '単発' End 区分," & Chr(13)
            gSL_Select = gSL_Select & "        Case When ISNULL(HKM.機械NO,'') = '' AND ISNULL(MT.値名称,'') = '' Then '' " & Chr(13)
            gSL_Select = gSL_Select & "             Else ISNULL(HKM.機械NO,'') + ' : ' + ISNULL(MT.値名称,'') END 機械NO," & Chr(13)
            gSL_Select = gSL_Select & "        Case When ISNULL(HKM.工程,'') = '' AND ISNULL(MY.値名称,'') = '' Then ''" & Chr(13)
            gSL_Select = gSL_Select & "             Else ISNULL(HKM.工程,'') + ' : ' + ISNULL(MY.値名称,'') END 工程," & Chr(13)
            gSL_Select = gSL_Select & "        HKM.工程数 工程数," & Chr(13)
            gSL_Select = gSL_Select & "        HKM.基準工数1 基準工数1,HKM.基準工数2 基準工数2,Sum(HNT.数量) 数量,Sum(HNT.数量) * HKM.工程数 表示工程数," & Chr(13)
            gSL_Select = gSL_Select & "        Sum(HNT.数量) * HKM.基準工数1 基準時間,Sum(HNT.数量) * HKM.基準工数2 目標時間" & Chr(13)
            '=================================================== 2010/09/16 Start
            gSL_Select = gSL_Select & "       ,Sum(HNT.数量) * ISNULL(HKM.使用数1,0) 使用数1,Sum(HNT.数量) * ISNULL(HKM.使用数2,0) 使用数2" & Chr(13)
            gSL_Select = gSL_Select & "       ,Sum(HNT.数量) * ISNULL(HKM.使用数3,0) 使用数3,Sum(HNT.数量) * ISNULL(HKM.使用数4,0) 使用数4" & Chr(13)
            gSL_Select = gSL_Select & "       ,Sum(HNT.数量) * ISNULL(HKM.使用数5,0) 使用数5" & Chr(13)
            '=================================================== 2010/09/16 End
            gSL_Select = gSL_Select & "   from 品番工順マスタ HKM" & Chr(13)
            gSL_Select = gSL_Select & "   LEFT OUTER JOIN 発注内示テーブル HNT" & Chr(13)
            gSL_Select = gSL_Select & "     On HNT.品番   = HKM.品番" & Chr(13)
            gSL_Select = gSL_Select & "    AND HNT.手配先 = HKM.取引先" & Chr(13)
            gSL_Select = gSL_Select & "    AND HNT.受入   = HKM.受入" & Chr(13)
            gSL_Select = gSL_Select & "   LEFT Outer Join 品番マスタ HNM" & Chr(13)
            gSL_Select = gSL_Select & "     On HNM.品番   = HKM.品番" & Chr(13)
            gSL_Select = gSL_Select & "   LEFT OUTER JOIN 取引先マスタ TRN" & Chr(13)
            gSL_Select = gSL_Select & "     On TRN.取引先CD   = HKM.取引先" & Chr(13)
            gSL_Select = gSL_Select & "    AND TRN.取引先区分 = '1'" & Chr(13)
            gSL_Select = gSL_Select & "    AND TRN.契約先区分 = '0'" & Chr(13)
            gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MS" & Chr(13)
            gSL_Select = gSL_Select & "     On MS.区分名称 = '機械'" & Chr(13)
            gSL_Select = gSL_Select & "    AND MS.値       = HKM.機械" & Chr(13)
            gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MT" & Chr(13)
            gSL_Select = gSL_Select & "     On MT.区分名称 = '機械NO'" & Chr(13)
            gSL_Select = gSL_Select & "    AND MT.値       = HKM.機械NO" & Chr(13)
            gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MY" & Chr(13)
            gSL_Select = gSL_Select & "     On MY.区分名称 = '工程管理'" & Chr(13)
            gSL_Select = gSL_Select & "    AND MY.値       = HKM.工程" & Chr(13)
'            gSL_Select = gSL_Select & "  Where HNT.内示年月 = '" & Mid(iH1_SDate, 1, 6) & "'" & Chr(13)        '2010/11/11
            gSL_Select = gSL_Select & "  Where HNT.内示年月 = '" & iH1_YM & "'" & Chr(13)                       '2010/11/11
            If RTrim(iH1_EDate) = "" Then
                gSL_Select = gSL_Select & "    And HNT.納期     = '" & iH1_SDate & "'" & Chr(13)
            Else
                gSL_Select = gSL_Select & "    And HNT.納期 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
            End If
            gSL_Select = gSL_Select & "    AND IsNull(HNT.品番,'') <> '' " & Chr(13)
            
        Else '明細表
            gSL_Select = ""
            gSL_Select = gSL_Select & " INSERT INTO 生産負荷リストワーク "
            '=================================================== 2010/09/16 Start
'            gSL_Select = gSL_Select & " ( 作成端末,部門,機械,区分,機械NO,工程,品番,工程数,基準時間1,基準時間2,数量,表示工程数,基準時間,目標時間 ) "
            gSL_Select = gSL_Select & " ( 作成端末,部門,機械,区分,機械NO,工程,品番,工程数,基準時間1,基準時間2,数量,表示工程数,基準時間,目標時間,使用数1,使用数2,使用数3,使用数4,使用数5 ) "
            '=================================================== 2010/09/16 End
            gSL_Select = gSL_Select & " Select substring(host_name(),1,20),Case When ISNULL(HKM.取引先,'') = '' AND ISNULL(TRN.略称,'') = '' Then '' " & Chr(13)
            gSL_Select = gSL_Select & "             Else ISNULL(HKM.取引先,'') + ' : ' + ISNULL(TRN.略称,'') END 部門," & Chr(13)
            gSL_Select = gSL_Select & "        Case When ISNULL(HKM.機械,'') = '' AND ISNULL(MS.値名称,'') = '' Then '' " & Chr(13)
            gSL_Select = gSL_Select & "             Else ISNULL(HKM.機械,'') + ' : ' + ISNULL(MS.値名称,'') END 機械," & Chr(13)
            gSL_Select = gSL_Select & "        Case When HKM.ロボット区分 = 0 then 'ﾛﾎﾞｯﾄ'" & Chr(13)
            gSL_Select = gSL_Select & "             Else '単発' End 区分," & Chr(13)
            gSL_Select = gSL_Select & "        Case When ISNULL(HKM.機械NO,'') = '' AND ISNULL(MT.値名称,'') = '' Then '' " & Chr(13)
            gSL_Select = gSL_Select & "             Else ISNULL(HKM.機械NO,'') + ' : ' + ISNULL(MT.値名称,'') END 機械NO," & Chr(13)
            gSL_Select = gSL_Select & "        Case When ISNULL(HKM.工程,'') = '' AND ISNULL(MY.値名称,'') = '' Then ''" & Chr(13)
            gSL_Select = gSL_Select & "             Else ISNULL(HKM.工程,'') + ' : ' + ISNULL(MY.値名称,'') END 工程," & Chr(13)
            gSL_Select = gSL_Select & "        ISNULL(HNM.表示品番,'') 品番,HKM.工程数 工程数," & Chr(13)
            gSL_Select = gSL_Select & "        HKM.基準工数1 基準工数1,HKM.基準工数2 基準工数2,Sum(HNT.数量) 数量,Sum(HNT.数量) * HKM.工程数 表示工程数," & Chr(13)
            gSL_Select = gSL_Select & "        Sum(HNT.数量) * HKM.基準工数1 基準時間,Sum(HNT.数量) * HKM.基準工数2 目標時間" & Chr(13)
            '=================================================== 2010/09/16 Start
            gSL_Select = gSL_Select & "       ,Sum(HNT.数量) * ISNULL(HKM.使用数1,0) 使用数1,Sum(HNT.数量) * ISNULL(HKM.使用数2,0) 使用数2" & Chr(13)
            gSL_Select = gSL_Select & "       ,Sum(HNT.数量) * ISNULL(HKM.使用数3,0) 使用数3,Sum(HNT.数量) * ISNULL(HKM.使用数4,0) 使用数4" & Chr(13)
            gSL_Select = gSL_Select & "       ,Sum(HNT.数量) * ISNULL(HKM.使用数5,0) 使用数5" & Chr(13)
            '=================================================== 2010/09/16 End
            gSL_Select = gSL_Select & "   from 品番工順マスタ HKM" & Chr(13)
            gSL_Select = gSL_Select & "   LEFT OUTER JOIN 発注内示テーブル HNT" & Chr(13)
            gSL_Select = gSL_Select & "     On HNT.品番   = HKM.品番" & Chr(13)
            gSL_Select = gSL_Select & "    AND HNT.手配先 = HKM.取引先" & Chr(13)
            gSL_Select = gSL_Select & "    AND HNT.受入   = HKM.受入" & Chr(13)
            gSL_Select = gSL_Select & "   LEFT Outer Join 品番マスタ HNM" & Chr(13)
            gSL_Select = gSL_Select & "     On HNM.品番   = HKM.品番" & Chr(13)
            gSL_Select = gSL_Select & "   LEFT OUTER JOIN 取引先マスタ TRN" & Chr(13)
            gSL_Select = gSL_Select & "     On TRN.取引先CD   = HKM.取引先" & Chr(13)
            gSL_Select = gSL_Select & "    AND TRN.取引先区分 = '1'" & Chr(13)
            gSL_Select = gSL_Select & "    AND TRN.契約先区分 = '0'" & Chr(13)
            gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MS" & Chr(13)
            gSL_Select = gSL_Select & "     On MS.区分名称 = '機械'" & Chr(13)
            gSL_Select = gSL_Select & "    AND MS.値       = HKM.機械" & Chr(13)
            gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MT" & Chr(13)
            gSL_Select = gSL_Select & "     On MT.区分名称 = '機械NO'" & Chr(13)
            gSL_Select = gSL_Select & "    AND MT.値       = HKM.機械NO" & Chr(13)
            gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MY" & Chr(13)
            gSL_Select = gSL_Select & "     On MY.区分名称 = '工程管理'" & Chr(13)
            gSL_Select = gSL_Select & "    AND MY.値       = HKM.工程" & Chr(13)
'            gSL_Select = gSL_Select & "  Where HNT.内示年月 = '" & Mid(iH1_SDate, 1, 6) & "'" & Chr(13)        '2010/11/11
            gSL_Select = gSL_Select & "  Where HNT.内示年月 = '" & iH1_YM & "'" & Chr(13)                       '2010/11/11
            If RTrim(iH1_EDate) = "" Then
                gSL_Select = gSL_Select & "    And HNT.納期     = '" & iH1_SDate & "'" & Chr(13)
            Else
                gSL_Select = gSL_Select & "    And HNT.納期 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
            End If
            gSL_Select = gSL_Select & "     AND IsNull(HNT.品番,'') <> '' " & Chr(13)
           
        End If
    Else
        If H1Op_KMoto(1).Value = True Then  '出荷
            If H1Op_OutPut(0).Value = True Then  '一覧表
                gSL_Select = ""
                gSL_Select = gSL_Select & " INSERT INTO 生産負荷リストワーク "
                '=================================================== 2010/09/16 Start
'                gSL_Select = gSL_Select & " ( 作成端末,部門,機械,区分,機械NO,工程,工程数,基準時間1,基準時間2,数量,表示工程数,基準時間,目標時間 ) "
                gSL_Select = gSL_Select & " ( 作成端末,部門,機械,区分,機械NO,工程,工程数,基準時間1,基準時間2,数量,表示工程数,基準時間,目標時間,使用数1,使用数2,使用数3,使用数4,使用数5 ) "
                '=================================================== 2010/09/16 End
                gSL_Select = gSL_Select & " Select substring(host_name(),1,20),Case When ISNULL(HKM.取引先,'') = '' AND ISNULL(TRN.略称,'') = '' Then '' " & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.取引先,'') + ' : ' + ISNULL(TRN.略称,'') END 部門," & Chr(13)
                gSL_Select = gSL_Select & "        Case When ISNULL(HKM.機械,'') = '' AND ISNULL(MS.値名称,'') = '' Then '' " & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.機械,'') + ' : ' + ISNULL(MS.値名称,'') END 機械," & Chr(13)
                gSL_Select = gSL_Select & "        Case When HKM.ロボット区分 = 0 then 'ﾛﾎﾞｯﾄ'" & Chr(13)
                gSL_Select = gSL_Select & "             Else '単発' End 区分," & Chr(13)
                gSL_Select = gSL_Select & "        Case When ISNULL(HKM.機械NO,'') = '' AND ISNULL(MT.値名称,'') = '' Then '' " & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.機械NO,'') + ' : ' + ISNULL(MT.値名称,'') END 機械NO," & Chr(13)
                gSL_Select = gSL_Select & "        Case When ISNULL(HKM.工程,'') = '' AND ISNULL(MY.値名称,'') = '' Then ''" & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.工程,'') + ' : ' + ISNULL(MY.値名称,'') END 工程," & Chr(13)
                gSL_Select = gSL_Select & "        HKM.工程数 工程数," & Chr(13)
                gSL_Select = gSL_Select & "        HKM.基準工数1 基準工数1,HKM.基準工数2 基準工数2,Sum(JKT.実績数) 数量,Sum(JKT.実績数) * HKM.工程数 表示工程数," & Chr(13)
                gSL_Select = gSL_Select & "        Sum(JKT.実績数) * HKM.基準工数1 基準時間,Sum(JKT.実績数) * HKM.基準工数2 目標時間" & Chr(13)
                '=================================================== 2010/09/16 Start
                gSL_Select = gSL_Select & "       ,Sum(JKT.実績数) * ISNULL(HKM.使用数1,0) 使用数1,Sum(JKT.実績数) * ISNULL(HKM.使用数2,0) 使用数2" & Chr(13)
                gSL_Select = gSL_Select & "       ,Sum(JKT.実績数) * ISNULL(HKM.使用数3,0) 使用数3,Sum(JKT.実績数) * ISNULL(HKM.使用数4,0) 使用数4" & Chr(13)
                gSL_Select = gSL_Select & "       ,Sum(JKT.実績数) * ISNULL(HKM.使用数5,0) 使用数5" & Chr(13)
                '=================================================== 2010/09/16 End
                gSL_Select = gSL_Select & "   from 品番工順マスタ HKM" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 実績管理テーブル JKT" & Chr(13)
                gSL_Select = gSL_Select & "     On JKT.品番   = HKM.品番" & Chr(13)
                gSL_Select = gSL_Select & "    AND JKT.実績先 = HKM.取引先" & Chr(13)
                gSL_Select = gSL_Select & "    AND JKT.受入   = HKM.受入" & Chr(13)
                gSL_Select = gSL_Select & "    AND JKT.工順   = HKM.工順" & Chr(13)
                gSL_Select = gSL_Select & "    AND JKT.実績区分 <> 1" & Chr(13)
                gSL_Select = gSL_Select & "    AND JKT.入力種類 <> 'A'" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT Outer Join 品番マスタ HNM" & Chr(13)
                gSL_Select = gSL_Select & "     On HNM.品番   = HKM.品番" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 取引先マスタ TRN" & Chr(13)
                gSL_Select = gSL_Select & "     On TRN.取引先CD   = HKM.取引先" & Chr(13)
                gSL_Select = gSL_Select & "    AND TRN.取引先区分 = '1'" & Chr(13)
                gSL_Select = gSL_Select & "    AND TRN.契約先区分 = '0'" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MS" & Chr(13)
                gSL_Select = gSL_Select & "     On MS.区分名称 = '機械'" & Chr(13)
                gSL_Select = gSL_Select & "    AND MS.値       = HKM.機械" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MT" & Chr(13)
                gSL_Select = gSL_Select & "     On MT.区分名称 = '機械NO'" & Chr(13)
                gSL_Select = gSL_Select & "    AND MT.値       = HKM.機械NO" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MY" & Chr(13)
                gSL_Select = gSL_Select & "     On MY.区分名称 = '工程管理'" & Chr(13)
                gSL_Select = gSL_Select & "    AND MY.値       = HKM.工程" & Chr(13)
                If RTrim(iH1_EDate) = "" Then
                    gSL_Select = gSL_Select & "   Where JKT.実績日 = '" & iH1_SDate & "'" & Chr(13)
                Else
                    gSL_Select = gSL_Select & "   Where JKT.実績日 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
                End If
                gSL_Select = gSL_Select & "    AND IsNull(JKT.品番,'') <> '' " & Chr(13)
                
            Else  '明細表
                gSL_Select = ""
                gSL_Select = gSL_Select & " INSERT INTO 生産負荷リストワーク "
                '============================================= 2010/09/16 Start
'                gSL_Select = gSL_Select & " ( 作成端末,部門,機械,区分,機械NO,工程,品番,工程数,基準時間1,基準時間2,数量,表示工程数,基準時間,目標時間 ) "
                gSL_Select = gSL_Select & " ( 作成端末,部門,機械,区分,機械NO,工程,品番,工程数,基準時間1,基準時間2,数量,表示工程数,基準時間,目標時間,使用数1,使用数2,使用数3,使用数4,使用数5 ) "
                '============================================= 2010/09/16 End
                gSL_Select = gSL_Select & " Select substring(host_name(),1,20),Case When ISNULL(HKM.取引先,'') = '' AND ISNULL(TRN.略称,'') = '' Then '' " & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.取引先,'') + ' : ' + ISNULL(TRN.略称,'') END 部門," & Chr(13)
                gSL_Select = gSL_Select & "        Case When ISNULL(HKM.機械,'') = '' AND ISNULL(MS.値名称,'') = '' Then '' " & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.機械,'') + ' : ' + ISNULL(MS.値名称,'') END 機械," & Chr(13)
                gSL_Select = gSL_Select & "        Case When HKM.ロボット区分 = 0 then 'ﾛﾎﾞｯﾄ'" & Chr(13)
                gSL_Select = gSL_Select & "             Else '単発' End 区分," & Chr(13)
                gSL_Select = gSL_Select & "        Case When ISNULL(HKM.機械NO,'') = '' AND ISNULL(MT.値名称,'') = '' Then '' " & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.機械NO,'') + ' : ' + ISNULL(MT.値名称,'') END 機械NO," & Chr(13)
                gSL_Select = gSL_Select & "        Case When ISNULL(HKM.工程,'') = '' AND ISNULL(MY.値名称,'') = '' Then ''" & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.工程,'') + ' : ' + ISNULL(MY.値名称,'') END 工程," & Chr(13)
                gSL_Select = gSL_Select & "        ISNULL(HNM.表示品番,'') 品番,HKM.工程数 工程数," & Chr(13)
                gSL_Select = gSL_Select & "        HKM.基準工数1 基準工数1,HKM.基準工数2 基準工数2,Sum(JKT.実績数) 数量,Sum(JKT.実績数) * HKM.工程数 表示工程数," & Chr(13)
                gSL_Select = gSL_Select & "        Sum(JKT.実績数) * HKM.基準工数1 基準時間,Sum(JKT.実績数) * HKM.基準工数2 目標時間" & Chr(13)
                '=================================================== 2010/09/16 Start
                gSL_Select = gSL_Select & "       ,Sum(JKT.実績数) * ISNULL(HKM.使用数1,0) 使用数1,Sum(JKT.実績数) * ISNULL(HKM.使用数2,0) 使用数2" & Chr(13)
                gSL_Select = gSL_Select & "       ,Sum(JKT.実績数) * ISNULL(HKM.使用数3,0) 使用数3,Sum(JKT.実績数) * ISNULL(HKM.使用数4,0) 使用数4" & Chr(13)
                gSL_Select = gSL_Select & "       ,Sum(JKT.実績数) * ISNULL(HKM.使用数5,0) 使用数5" & Chr(13)
                '=================================================== 2010/09/16 End
                gSL_Select = gSL_Select & "   from 品番工順マスタ HKM" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 実績管理テーブル JKT" & Chr(13)
                gSL_Select = gSL_Select & "     On JKT.品番   = HKM.品番" & Chr(13)
                gSL_Select = gSL_Select & "    AND JKT.実績先 = HKM.取引先" & Chr(13)
                gSL_Select = gSL_Select & "    AND JKT.受入   = HKM.受入" & Chr(13)
                gSL_Select = gSL_Select & "    AND JKT.工順   = HKM.工順" & Chr(13)
                gSL_Select = gSL_Select & "    AND JKT.実績区分 <> 1" & Chr(13)
                gSL_Select = gSL_Select & "    AND JKT.入力種類 <> 'A'" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT Outer Join 品番マスタ HNM" & Chr(13)
                gSL_Select = gSL_Select & "     On HNM.品番   = HKM.品番" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 取引先マスタ TRN" & Chr(13)
                gSL_Select = gSL_Select & "     On TRN.取引先CD   = HKM.取引先" & Chr(13)
                gSL_Select = gSL_Select & "    AND TRN.取引先区分 = '1'" & Chr(13)
                gSL_Select = gSL_Select & "    AND TRN.契約先区分 = '0'" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MS" & Chr(13)
                gSL_Select = gSL_Select & "     On MS.区分名称 = '機械'" & Chr(13)
                gSL_Select = gSL_Select & "    AND MS.値       = HKM.機械" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MT" & Chr(13)
                gSL_Select = gSL_Select & "     On MT.区分名称 = '機械NO'" & Chr(13)
                gSL_Select = gSL_Select & "    AND MT.値       = HKM.機械NO" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MY" & Chr(13)
                gSL_Select = gSL_Select & "     On MY.区分名称 = '工程管理'" & Chr(13)
                gSL_Select = gSL_Select & "    AND MY.値       = HKM.工程" & Chr(13)
                If RTrim(iH1_EDate) = "" Then
                    gSL_Select = gSL_Select & "   Where JKT.実績日 = '" & iH1_SDate & "'" & Chr(13)
                Else
                    gSL_Select = gSL_Select & "   Where JKT.実績日 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
                End If
                gSL_Select = gSL_Select & "    AND IsNull(JKT.品番,'') <> '' " & Chr(13)
            
            End If
        Else  '実績
            If H1Op_OutPut(0).Value = True Then  '一覧表
                gSL_Select = ""
                gSL_Select = gSL_Select & " INSERT INTO 生産負荷リストワーク "
                '=================================================== 2010/09/16 Start
                'gSL_Select = gSL_Select & " ( 作成端末,部門,機械,区分,機械NO,工程,工程数,基準時間1,基準時間2,数量,表示工程数,基準時間,目標時間 ) "
                gSL_Select = gSL_Select & " ( 作成端末,部門,機械,区分,機械NO,工程,工程数,基準時間1,基準時間2,数量,表示工程数,基準時間,目標時間,使用数1,使用数2,使用数3,使用数4,使用数5 ) "
                '=================================================== 2010/09/16 End
                gSL_Select = gSL_Select & " Select substring(host_name(),1,20),Case When ISNULL(HKM.取引先,'') = '' AND ISNULL(TRN.略称,'') = '' Then '' " & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.取引先,'') + ' : ' + ISNULL(TRN.略称,'') END 部門," & Chr(13)
                gSL_Select = gSL_Select & "        Case When ISNULL(HKM.機械,'') = '' AND ISNULL(MS.値名称,'') = '' Then '' " & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.機械,'') + ' : ' + ISNULL(MS.値名称,'') END 機械," & Chr(13)
                gSL_Select = gSL_Select & "        Case When HKM.ロボット区分 = 0 then 'ﾛﾎﾞｯﾄ'" & Chr(13)
                gSL_Select = gSL_Select & "             Else '単発' End 区分," & Chr(13)
                gSL_Select = gSL_Select & "        Case When ISNULL(HKM.機械NO,'') = '' AND ISNULL(MT.値名称,'') = '' Then '' " & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.機械NO,'') + ' : ' + ISNULL(MT.値名称,'') END 機械NO," & Chr(13)
                gSL_Select = gSL_Select & "        Case When ISNULL(HKM.工程,'') = '' AND ISNULL(MY.値名称,'') = '' Then ''" & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.工程,'') + ' : ' + ISNULL(MY.値名称,'') END 工程," & Chr(13)
                gSL_Select = gSL_Select & "        HKM.工程数 工程数," & Chr(13)
                gSL_Select = gSL_Select & "        HKM.基準工数1 基準工数1,HKM.基準工数2 基準工数2,Sum(SNT.加工数) 数量,Sum(SNT.加工数) * HKM.工程数 表示工程数," & Chr(13)
                gSL_Select = gSL_Select & "        Sum(SNT.加工数) * HKM.基準工数1 基準時間,Sum(SNT.加工数) * HKM.基準工数2 目標時間" & Chr(13)
                '=================================================== 2010/09/16 Start
                gSL_Select = gSL_Select & "       ,Sum(SNT.加工数) * ISNULL(HKM.使用数1,0) 使用数1,Sum(SNT.加工数) * ISNULL(HKM.使用数2,0) 使用数2" & Chr(13)
                gSL_Select = gSL_Select & "       ,Sum(SNT.加工数) * ISNULL(HKM.使用数3,0) 使用数3,Sum(SNT.加工数) * ISNULL(HKM.使用数4,0) 使用数4" & Chr(13)
                gSL_Select = gSL_Select & "       ,Sum(SNT.加工数) * ISNULL(HKM.使用数5,0) 使用数5" & Chr(13)
                '=================================================== 2010/09/16 End
                gSL_Select = gSL_Select & "   from 品番工順マスタ HKM" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 作業日報テーブル SNT" & Chr(13)
                gSL_Select = gSL_Select & "     On SNT.品番   = HKM.品番" & Chr(13)
                gSL_Select = gSL_Select & "    AND SNT.部門   = HKM.取引先" & Chr(13)
                gSL_Select = gSL_Select & "    AND SNT.受入   = HKM.受入" & Chr(13)
'''                gSL_Select = gSL_Select & "    AND SNT.工順   = HKM.工順" & Chr(13)              '2011/12/28
                gSL_Select = gSL_Select & "   LEFT Outer Join 品番マスタ HNM" & Chr(13)
                gSL_Select = gSL_Select & "     On HNM.品番   = HKM.品番" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 取引先マスタ TRN" & Chr(13)
                gSL_Select = gSL_Select & "     On TRN.取引先CD   = HKM.取引先" & Chr(13)
                gSL_Select = gSL_Select & "    AND TRN.取引先区分 = '1'" & Chr(13)
                gSL_Select = gSL_Select & "    AND TRN.契約先区分 = '0'" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MS" & Chr(13)
                gSL_Select = gSL_Select & "     On MS.区分名称 = '機械'" & Chr(13)
                gSL_Select = gSL_Select & "    AND MS.値       = HKM.機械" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MT" & Chr(13)
                gSL_Select = gSL_Select & "     On MT.区分名称 = '機械NO'" & Chr(13)
                gSL_Select = gSL_Select & "    AND MT.値       = HKM.機械NO" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MY" & Chr(13)
                gSL_Select = gSL_Select & "     On MY.区分名称 = '工程管理'" & Chr(13)
                gSL_Select = gSL_Select & "    AND MY.値       = HKM.工程" & Chr(13)
                If RTrim(iH1_EDate) = "" Then
                    gSL_Select = gSL_Select & "   Where SNT.作業日 = '" & iH1_SDate & "'" & Chr(13)
                Else
                    gSL_Select = gSL_Select & "   Where SNT.作業日 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
                End If
                gSL_Select = gSL_Select & "    AND IsNull(SNT.品番,'') <> '' " & Chr(13)

            Else                                 '明細表
                gSL_Select = ""
                gSL_Select = gSL_Select & " INSERT INTO 生産負荷リストワーク "
                '=================================================== 2010/09/16 Start
                'gSL_Select = gSL_Select & " ( 作成端末,部門,機械,区分,機械NO,工程,品番,工程数,基準時間1,基準時間2,数量,表示工程数,基準時間,目標時間 ) "
                gSL_Select = gSL_Select & " ( 作成端末,部門,機械,区分,機械NO,工程,品番,工程数,基準時間1,基準時間2,数量,表示工程数,基準時間,目標時間,使用数1,使用数2,使用数3,使用数4,使用数5 ) "
                '=================================================== 2010/09/16 End
                gSL_Select = gSL_Select & " Select substring(host_name(),1,20),Case When ISNULL(HKM.取引先,'') = '' AND ISNULL(TRN.略称,'') = '' Then '' " & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.取引先,'') + ' : ' + ISNULL(TRN.略称,'') END 部門," & Chr(13)
                gSL_Select = gSL_Select & "        Case When ISNULL(HKM.機械,'') = '' AND ISNULL(MS.値名称,'') = '' Then '' " & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.機械,'') + ' : ' + ISNULL(MS.値名称,'') END 機械," & Chr(13)
                gSL_Select = gSL_Select & "        Case When HKM.ロボット区分 = 0 then 'ﾛﾎﾞｯﾄ'" & Chr(13)
                gSL_Select = gSL_Select & "             Else '単発' End 区分," & Chr(13)
                gSL_Select = gSL_Select & "        Case When ISNULL(HKM.機械NO,'') = '' AND ISNULL(MT.値名称,'') = '' Then '' " & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.機械NO,'') + ' : ' + ISNULL(MT.値名称,'') END 機械NO," & Chr(13)
                gSL_Select = gSL_Select & "        Case When ISNULL(HKM.工程,'') = '' AND ISNULL(MY.値名称,'') = '' Then ''" & Chr(13)
                gSL_Select = gSL_Select & "             Else ISNULL(HKM.工程,'') + ' : ' + ISNULL(MY.値名称,'') END 工程," & Chr(13)
                gSL_Select = gSL_Select & "        ISNULL(HNM.表示品番,'') 品番,HKM.工程数 工程数," & Chr(13)
                gSL_Select = gSL_Select & "        HKM.基準工数1 基準工数1,HKM.基準工数2 基準工数2,Sum(SNT.加工数) 数量,Sum(SNT.加工数) * HKM.工程数 表示工程数," & Chr(13)
                gSL_Select = gSL_Select & "        Sum(SNT.加工数) * HKM.基準工数1 基準時間,Sum(SNT.加工数) * HKM.基準工数2 目標時間" & Chr(13)
                '=================================================== 2010/09/16 Start
                gSL_Select = gSL_Select & "       ,Sum(SNT.加工数) * ISNULL(HKM.使用数1,0) 使用数1,Sum(SNT.加工数) * ISNULL(HKM.使用数2,0) 使用数2" & Chr(13)
                gSL_Select = gSL_Select & "       ,Sum(SNT.加工数) * ISNULL(HKM.使用数3,0) 使用数3,Sum(SNT.加工数) * ISNULL(HKM.使用数4,0) 使用数4" & Chr(13)
                gSL_Select = gSL_Select & "       ,Sum(SNT.加工数) * ISNULL(HKM.使用数5,0) 使用数5" & Chr(13)
                '=================================================== 2010/09/16 End
                gSL_Select = gSL_Select & "   from 品番工順マスタ HKM" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 作業日報テーブル SNT" & Chr(13)
                gSL_Select = gSL_Select & "     On SNT.品番   = HKM.品番" & Chr(13)
                gSL_Select = gSL_Select & "    AND SNT.部門   = HKM.取引先" & Chr(13)
                gSL_Select = gSL_Select & "    AND SNT.受入   = HKM.受入" & Chr(13)
'''                gSL_Select = gSL_Select & "    AND SNT.工順   = HKM.工順" & Chr(13)                  '2011/12/28
                gSL_Select = gSL_Select & "   LEFT Outer Join 品番マスタ HNM" & Chr(13)
                gSL_Select = gSL_Select & "     On HNM.品番   = HKM.品番" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 取引先マスタ TRN" & Chr(13)
                gSL_Select = gSL_Select & "     On TRN.取引先CD   = HKM.取引先" & Chr(13)
                gSL_Select = gSL_Select & "    AND TRN.取引先区分 = '1'" & Chr(13)
                gSL_Select = gSL_Select & "    AND TRN.契約先区分 = '0'" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MS" & Chr(13)
                gSL_Select = gSL_Select & "     On MS.区分名称 = '機械'" & Chr(13)
                gSL_Select = gSL_Select & "    AND MS.値       = HKM.機械" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MT" & Chr(13)
                gSL_Select = gSL_Select & "     On MT.区分名称 = '機械NO'" & Chr(13)
                gSL_Select = gSL_Select & "    AND MT.値       = HKM.機械NO" & Chr(13)
                gSL_Select = gSL_Select & "   LEFT OUTER JOIN 名称マスタ MY" & Chr(13)
                gSL_Select = gSL_Select & "     On MY.区分名称 = '工程管理'" & Chr(13)
                gSL_Select = gSL_Select & "    AND MY.値       = HKM.工程" & Chr(13)
                If RTrim(iH1_EDate) = "" Then
                    gSL_Select = gSL_Select & "   Where SNT.作業日 = '" & iH1_SDate & "'" & Chr(13)
                Else
                    gSL_Select = gSL_Select & "   Where SNT.作業日 Between '" & iH1_SDate & "' AND '" & iH1_EDate & "'" & Chr(13)
                End If
                gSL_Select = gSL_Select & "    AND IsNull(SNT.品番,'') <> '' " & Chr(13)
            
            End If
        
        End If
        
    End If
    
    If H1Op_OutPut(0).Value Then '一覧表用Where句作成

        If RTrim(iH1_Bumon) <> "" Then   '部門
            gsl_Where1 = gsl_Where1 & "       AND HKM.取引先 = '" & iH1_Bumon & "'" & Chr(13)
        End If
        If RTrim(iH1_Kikai) <> "" Then   '機械
            gsl_Where1 = gsl_Where1 & "       AND HKM.機械   = '" & iH1_Kikai & "'" & Chr(13)
        End If
        If RTrim(iH1_KikaiNo) <> "" Then '機械no
            gsl_Where1 = gsl_Where1 & "       AND HKM.機械NO = '" & iH1_KikaiNo & "'" & Chr(13)
        End If
        gsl_Where1 = gsl_Where1 & " GROUP BY HKM.取引先,TRN.略称,HKM.機械,MS.値名称,HKM.ロボット区分,HKM.機械NO," & Chr(13)
        gsl_Where1 = gsl_Where1 & "          MT.値名称,HKM.工程,MY.値名称,HKM.工程数,HKM.基準工数1,HKM.基準工数2" & Chr(13)
        gsl_Where1 = gsl_Where1 & "         ,HKM.使用数1,HKM.使用数2,HKM.使用数3,HKM.使用数4,HKM.使用数5" & Chr(13)         '2010/09/16
        gsl_Where1 = gsl_Where1 & " ORDER BY HKM.取引先,HKM.機械,HKM.ロボット区分,HKM.機械NO,HKM.工程"

    Else '明細表用Where句作成
    
        If RTrim(iH1_Bumon) <> "" Then   '部門
            gsl_Where2 = gsl_Where2 & "       AND HKM.取引先 = '" & iH1_Bumon & "'" & Chr(13)
        End If
        If RTrim(iH1_Kikai) <> "" Then   '機械
            gsl_Where2 = gsl_Where2 & "       AND HKM.機械   = '" & iH1_Kikai & "'" & Chr(13)
        End If
        If RTrim(iH1_KikaiNo) <> "" Then '機械no
            gsl_Where2 = gsl_Where2 & "       AND HKM.機械NO = '" & iH1_KikaiNo & "'" & Chr(13)
        End If
        gsl_Where2 = gsl_Where2 & " GROUP BY HKM.取引先,TRN.略称,HKM.機械,MS.値名称,HKM.ロボット区分,HKM.機械NO," & Chr(13)
        gsl_Where2 = gsl_Where2 & "          MT.値名称,HKM.工程,MY.値名称,HNM.表示品番,HKM.工程数,HKM.基準工数1,HKM.基準工数2" & Chr(13)
        gsl_Where2 = gsl_Where2 & "         ,HKM.使用数1,HKM.使用数2,HKM.使用数3,HKM.使用数4,HKM.使用数5" & Chr(13)         '2010/09/16
        gsl_Where2 = gsl_Where2 & " ORDER BY HKM.取引先,HKM.機械,HKM.ロボット区分,HKM.機械NO,HKM.工程,HNM.表示品番"
    
    End If
    
    If H1Op_OutPut(0).Value Then
        gSL_Select = gSL_Select & gsl_Where1
    Else
        gSL_Select = gSL_Select & gsl_Where2
    End If
   
    With CisDB
        '.ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
        
    '表示データ
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT "
    gSL_Select = gSL_Select & "    部門,機械,区分,機械NO,工程,品番, "
    gSL_Select = gSL_Select & "    ISNULL(SUM(数量),0) 数量,ISNULL(SUM(表示工程数),0) 表示工程数,ISNULL(SUM(基準時間),0) 基準時間,ISNULL(SUM(目標時間),0) 目標時間 "
    gSL_Select = gSL_Select & "   ,ISNULL(SUM(使用数1),0) 使用数1,ISNULL(SUM(使用数2),0) 使用数2,ISNULL(SUM(使用数3),0) 使用数3,ISNULL(SUM(使用数4),0) 使用数4,ISNULL(SUM(使用数5),0) 使用数5 " '2010/09/16
    gSL_Select = gSL_Select & "FROM "
    gSL_Select = gSL_Select & "( SELECT DISTINCT  作成端末,部門,機械,区分,機械NO,工程,品番, "
    gSL_Select = gSL_Select & "                   数量,表示工程数,基準時間,目標時間,  "
    gSL_Select = gSL_Select & "                   工程数,基準時間1,基準時間2 "
    gSL_Select = gSL_Select & "                  ,使用数1,使用数2,使用数3,使用数4,使用数5 "         '2010/09/16
    gSL_Select = gSL_Select & "                FROM  生産負荷リストワーク "
    gSL_Select = gSL_Select & "                WHERE 作成端末 = SUBSTRING(HOST_NAME(),1,20) ) SD "
    gSL_Select = gSL_Select & "WHERE 作成端末 =  SUBSTRING(HOST_NAME(),1,20) "
    gSL_Select = gSL_Select & "GROUP BY "
    gSL_Select = gSL_Select & "    部門,機械,区分,機械NO,工程,品番 "
    gSL_Select = gSL_Select & "ORDER BY 部門,機械,区分,機械NO,工程,品番"

    
    If Not HKTRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　該当データはありません。            "
            .MB_Title = "負荷リスト発行"
            .MB_Button = Error
            If Not .MBOX Then Exit Function
        End With
       Call HKTClose
       Exit Function
    End If
    
    Create_List = True
End Function
