VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form COY0310 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "売上予想表"
   ClientHeight    =   11190
   ClientLeft      =   75
   ClientTop       =   1515
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   Begin VB.PictureBox Csv_Bar 
      Height          =   345
      Left            =   4380
      ScaleHeight     =   285
      ScaleWidth      =   6555
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2325
      Visible         =   0   'False
      Width           =   6615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1605
      Top             =   4590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   585
      Left            =   5625
      Top             =   8880
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
      Height          =   5430
      Left            =   4350
      ScaleHeight     =   5589.189
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6600
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2730
      Width           =   6660
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   5265
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   9287
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
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   435
            Left            =   1875
            Top             =   2580
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   767
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
            Begin VB.CheckBox ChkH1_HinKb 
               BackColor       =   &H00E0E0E0&
               Caption         =   "型費"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   1650
               TabIndex        =   4
               Top             =   105
               Value           =   1  'ﾁｪｯｸ
               Width           =   1260
            End
            Begin VB.CheckBox ChkH1_HinKb 
               BackColor       =   &H00E0E0E0&
               Caption         =   "部品"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   180
               TabIndex        =   3
               Top             =   105
               Value           =   1  'ﾁｪｯｸ
               Width           =   1260
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   330
            Left            =   675
            Top             =   1530
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   582
            ForeColor       =   16711680
            Caption         =   "契約先"
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
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   375
            Left            =   1875
            TabIndex        =   0
            Top             =   930
            Width           =   1440
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
            Left            =   75
            TabIndex        =   18
            Top             =   4140
            Width           =   6345
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "CSV"
               Height          =   240
               Index           =   2
               Left            =   4380
               TabIndex        =   11
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "画面(ﾌﾟﾚﾋﾞｭｰ)"
               Height          =   240
               Index           =   1
               Left            =   2505
               TabIndex        =   10
               Top             =   270
               Width           =   1875
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "プリンター"
               Height          =   240
               Index           =   0
               Left            =   1065
               TabIndex        =   9
               Top             =   270
               Value           =   -1  'True
               Width           =   1260
            End
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
               TabIndex        =   14
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
                  Caption         =   "売上予想表　発行を行います"
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
            Top             =   4770
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
            Top             =   4800
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
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   360
            Index           =   1
            Left            =   570
            Top             =   930
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   635
            ForeColor       =   16711680
            Caption         =   "予想年月"
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   1875
            TabIndex        =   1
            Top             =   1500
            Width           =   1020
            _ExtentX        =   1799
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
         Begin Cis3D_v60.CIS3D H1lb_Torcd 
            Height          =   360
            Left            =   3240
            Top             =   1500
            Width           =   2535
            _ExtentX        =   4471
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D PB_Torcd 
            Height          =   360
            Left            =   2895
            Top             =   1500
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
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   435
            Left            =   1875
            Top             =   3045
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   767
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
            Begin VB.CheckBox ChkH1_PrtSyu 
               BackColor       =   &H00E0E0E0&
               Caption         =   "集計表"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   180
               TabIndex        =   5
               Top             =   105
               Value           =   1  'ﾁｪｯｸ
               Width           =   1260
            End
            Begin VB.CheckBox ChkH1_PrtSyu 
               BackColor       =   &H00E0E0E0&
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
               Height          =   225
               Index           =   1
               Left            =   1650
               TabIndex        =   6
               Top             =   105
               Value           =   1  'ﾁｪｯｸ
               Width           =   1260
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   330
            Left            =   675
            Top             =   2655
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   582
            ForeColor       =   16711680
            Caption         =   "印刷種類"
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
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   330
            Left            =   675
            Top             =   2085
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   582
            ForeColor       =   16711680
            Caption         =   "工区"
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
         Begin CisText_V60.CisText iH1_Kouku 
            Height          =   375
            Left            =   1875
            TabIndex        =   2
            Top             =   2055
            Width           =   270
            _ExtentX        =   476
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
            Text            =   "X"
            MaxLength       =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   435
            Left            =   1875
            Top             =   3555
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   767
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
            Begin VB.OptionButton OpH1_HinKb 
               BackColor       =   &H00E0E0E0&
               Caption         =   "全桁"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   150
               TabIndex        =   7
               Top             =   105
               Value           =   -1  'True
               Width           =   1320
            End
            Begin VB.OptionButton OpH1_HinKb 
               BackColor       =   &H00E0E0E0&
               Caption         =   "集約"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   1650
               TabIndex        =   8
               Top             =   105
               Width           =   1425
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   330
            Left            =   675
            Top             =   3630
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   582
            ForeColor       =   16711680
            Caption         =   "品番形態"
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
      End
   End
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   375
      Left            =   0
      Top             =   -15
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 売 上 予 想 表 】"
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
      TabIndex        =   13
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
   Begin Cis3D_v60.CIS3D PB_Ent 
      Height          =   585
      Left            =   8280
      Top             =   8880
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
      Left            =   6915
      Top             =   8880
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
   Begin VB.Label PNL_Ok 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6900
      TabIndex        =   20
      Top             =   9495
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8280
      TabIndex        =   17
      Top             =   9495
      Width           =   1485
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5610
      TabIndex        =   16
      Top             =   9495
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
      TabIndex        =   15
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "COY0310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  売上予想表 発行
'**       フォームID    :  CXU0310
'**       処理概要      :
'**
'**       作  成  日    :  2004/11/05
'**       変  更  日    :  2005/03/04  By CIS - 重複起動ﾁｪｯｸで起動済の時、『実行時エラー'91'』を表示する対応（Timerﾌﾟﾛﾊﾟﾃｨの Interval=0 に修正）
'**       変  更  日    :  2010/09/01  By CIS - 車型廃止、明細に背番号追加
'**
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim mPrtDev         As String       ' ﾌﾟﾘﾝﾀ出力先
    Dim mCsvDir         As String
    Dim mCsvKey         As String
    Dim mCsvName        As String
    
    Dim mMsgType        As String
    Dim SL_Select       As String
    
    Dim mSelColor       As Long
    Dim mUnSelColor     As Long
    Dim mSavePrint      As String
    
    Dim mSaveMsg        As String
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
    
    With CisFrm
        Set .MSGObject = lb_msg
        Set .MSGTimer = Timer2
        .MSGStopCount = 10
        .MSGTYPE = Left_To_Right
        .MSGText = mSaveMsg
        .MSGInterval = 100
        .MSGRelate
    End With

    Dummy.Left = -1000
    Dummy.Enabled = False

    mSelColor = Op_Out(0).BackColor
    mUnSelColor = Op_Out(1).BackColor
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
    
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
    
    ' 対象年月初期表示値を求める
    gSL_Select = "Select Max(内示年月) 内示年月 From 売上予想テーブル "
    If UYTRead(gSL_Select, 1) Then
        iH1_YM = UYT.内示年月
    End If

'CSVﾌｧｲﾙ出力先
    mCsvName = "売上予想表"
    mCsvKey = "COY0310"
    mCsvDir = GetSetting(appname:=gRegKey, section:="CSV", _
                       Key:=mCsvKey, Default:=Mid(CisFun.WinRoot, 1, 3) & mCsvName & ".csv")
''ﾌﾟﾘﾝﾀ出力先
    Op_Out(0).Value = True
    mPrtDev = Printer.DeviceName
    lb_PrtDev = mPrtDev
    mSavePrint = mPrtDev
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 1000   ' 時刻表示間隔
    Timer2.Enabled = True    ' 時刻表示開始
    
    FormAct = Not FormAct
    
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
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
        .Gradation_Direction = Portrait
        .Gradation_Shading = Deepen
        .Gradation 60, 190, 200, 4, 2, 3
    End With
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
'    gp_PrtKB = Index
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
    If Dummy.Tag = "iH1_Torcd" Then Call PB_Torcd_Click: Exit Sub
End Sub
'+---------------------+
'+    契約先検索        +
'+---------------------+
Private Sub PB_Torcd_Click()

    RV_TorKb = 2    ' 得意先
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
'+    iH1_Torcd(GotFocus)    +
'+----------------------------+
Private Sub iH1_Torcd_GotFocus()
    Dummy.Tag = ActiveControl.Name
End Sub
'+----------------------------+
'+    iH1_Torcd(LostFocus)    +
'+----------------------------+
Private Sub iH1_Torcd_LostFocus()
    Dummy.Tag = ""
    If TorNmGet(iH1_Torcd, 2) Then
        H1lb_Torcd = TRM_RName
        If Trim(H1lb_Torcd) = "" Then
           H1lb_Torcd = TRM_Name
        End If
    Else
        H1lb_Torcd = ""
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
       Case "H1"
          H1Mode = True
       Case "B2"
          B2Mode = True
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
    
    If Not Head1Chk Then GoTo ReturnPress_Ee
        
    Call DispChange("B1")
    
    If Not PrintStart Then GoTo ReturnPress_Ee

    gStr = "【  出  力  終  了  】"
    lb_msg.Caption = gStr
    DoEvents
    
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
    With CisFrm
        .MSGTYPE = FSize_Changes
        .MSGText = gStr
    End With
    ProcHB = "H1"
    Call DispChange(ProcHB)
    Exit Sub
ReturnPress_Ee:
    lb_msg = mSaveMsg
    DoEvents
    ProcHB = "H1"
    Call DispChange(ProcHB)
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Dim wSl_Where   As String
    
    Head1Chk = False
    
    If Trim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　予想年月を入力して下さい    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If

    If Trim(iH1_Torcd) <> "" Then
        If Not TorNmGet(iH1_Torcd, 2) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "   　契約先　未登録    "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
    End If
'印刷種別
    If ChkH1_HinKb(0).Value = 0 And ChkH1_HinKb(1).Value = 0 Then
        ChkH1_HinKb(0).Value = 1: ChkH1_HinKb(1).Value = 1
    End If
    If ChkH1_PrtSyu(0).Value = 0 And ChkH1_PrtSyu(1).Value = 0 Then
        ChkH1_PrtSyu(0).Value = 1: ChkH1_PrtSyu(1).Value = 1
    End If
'データ存在確認
    UYT.件数 = 0: UYK.件数 = 0
    If ChkH1_HinKb(0).Value = 1 Then
        gSL_Select = "Select COUNT(品番) 件数 from 売上予想テーブル " & vbCr
        gSL_Select = gSL_Select & " Where 内示年月 = '" & iH1_YM & "'" & vbCr
        If Trim(iH1_Torcd) <> "" Then
           gSL_Select = gSL_Select & " And 契約先 = '" & iH1_Torcd & "'" & vbCr
        End If
        If Trim(iH1_Kouku) <> "" Then
           gSL_Select = gSL_Select & " And 工区 = '" & iH1_Kouku & "'" & vbCr
        End If
        If UYTRead(gSL_Select, 1) Then
            If UYT.件数 = 0 And ChkH1_HinKb(1).Value = 0 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "   　[部品]該当データがありません       "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_YM.SetFocus
                Exit Function
            End If
        End If
    End If
    If ChkH1_HinKb(1).Value = 1 Then
        gSL_Select = "Select COUNT(品番) 件数 from 売上予想型費テーブル " & vbCr
        gSL_Select = gSL_Select & " Where 予想年月 = '" & iH1_YM & "'" & vbCr
        If Trim(iH1_Torcd) <> "" Then
           gSL_Select = gSL_Select & " And 契約先 = '" & iH1_Torcd & "'" & vbCr
        End If
        If UYKRead(gSL_Select, 1) Then
            If UYK.件数 = 0 And ChkH1_HinKb(0).Value = 0 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "   　[型費]該当データがありません       "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_YM.SetFocus
                Exit Function
            End If
        End If
    End If
    If UYT.件数 = 0 And UYK.件数 = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　該当データがありません       "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
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
'*******************************
'*      メッセージ表示        *
'*******************************
Private Sub Timer2_Timer()
    CisFrm.MSGRelate
End Sub
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
'
    UY_NJYm = RTrim(iH1_YM)
    UY_Torcd = RTrim(iH1_Torcd)
    UY_Kouku = RTrim(iH1_Kouku)
    If OpH1_HinKb(0).Value Then
        UY_HinKeta = 0
    Else
        UY_HinKeta = 1
    End If
    
    If Op_Out(0).Value Or Op_Out(1).Value Then
        If ChkH1_HinKb(0).Value <> 0 And UYT.件数 <> 0 Then
            '*-------------------------------------*
            '*  売上予想表【部品・集計】
            '*-------------------------------------*
            If ChkH1_PrtSyu(0).Value <> 0 Then
                lb_msg.Caption = "【 売上予想集計表《 部品 》 印刷中 】"
                DoEvents
                With CisPrt
                     Set .iPrtForm = COY0310L
                     .dpMSG = "売上予想集計表【 部品 】　印刷中"
                     If Op_Out(0).Value Then
                        .iPreview = pv_Print
                     Else
                        .iPreview = pv_Preview
                     End If
                     .PrintStart
                End With
            End If
            '*-------------------------------------*
            '*  売上予想表【部品・明細】
            '*-------------------------------------*
            If ChkH1_PrtSyu(1).Value <> 0 Then
                lb_msg.Caption = "【 売上予想表《 部品 》 印刷中 】"
                DoEvents
                With CisPrt
                     Set .iPrtForm = COY0311L
                     .dpMSG = "売上予想表【 部品 】　印刷中"
                     If Op_Out(0).Value Then
                        .iPreview = pv_Print
                     Else
                        .iPreview = pv_Preview
                     End If
                     .PrintStart
                End With
            End If
        End If
        If ChkH1_HinKb(1).Value <> 0 And UYK.件数 <> 0 Then
            '*-------------------------------------*
            '*  売上予想表【型費・集計】
            '*-------------------------------------*
            If ChkH1_PrtSyu(0).Value <> 0 Then
                lb_msg.Caption = "【 売上予想集計表《 型費 》 印刷中 】"
                DoEvents
                With CisPrt
                     Set .iPrtForm = COY0312L
                     .dpMSG = "売上予想集計表【 型費 】　印刷中"
                     If Op_Out(0).Value Then
                        .iPreview = pv_Print
                     Else
                        .iPreview = pv_Preview
                     End If
                     .PrintStart
                End With
            End If
            '*-------------------------------------*
            '*  売上予想表【型費・明細】
            '*-------------------------------------*
            If ChkH1_PrtSyu(1).Value <> 0 Then
                lb_msg.Caption = "【 売上予想表《 型費 》 印刷中 】"
                DoEvents
                With CisPrt
                     Set .iPrtForm = COY0313L
                     .dpMSG = "売上予想表【 型費 】　印刷中"
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

    Dim csvNJYm1    As String
    Dim csvNJYm2    As String
    Dim csvNJYm3    As String
    Dim OutStr      As String
    Dim OutFNo      As Long
    Dim OutCnt      As Long
    Dim TotalCnt    As Long
    
    CsvPut = False
    
    ' 出力ﾌｧｲﾙ存在確認
    If Dir(mCsvDir) <> "" Then
        DoEvents
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　 " & mCsvDir & "　ﾌｧｲﾙは既に存在します。                 "
            .MB_MSG(4) = "   　 " & "上書きしますか？           "
            .MB_Title = "上書き確認"
            .MB_Button = Yes_No
            If Not .MBOX Then Exit Function
        End With
    End If
    
    lb_msg.Caption = "【 売上予想集計表（部品） 出力中 】"
    Csv_Bar.Visible = True
    
    DoEvents
        
    OutFNo = FreeFile
    Open mCsvDir For Output As #OutFNo

    Call CisFrm.MousePT(11)
    
    With CisFrm
        Set .MeterPic = Csv_Bar
    End With
    TotalCnt = 0
'*---------------------------------------*
'*  売上予想CSV出力 【部品・集計】
'*---------------------------------------*
    If ChkH1_HinKb(0).Value <> 0 And ChkH1_PrtSyu(0).Value <> 0 Then
        
        Call SQLEdit0310
        If UYTRead(gSL_Select, 0) Then
            CisFrm.MeterMax = CisDB.RecordCount
            TotalCnt = TotalCnt + CisDB.RecordCount
            
            csvNJYm1 = Mid(UYT.内示年月, 5, 2)
            csvNJYm2 = Format(DateAdd("m", 1, CDate(Left(UYT.内示年月, 4) & "/" & Mid(UYT.内示年月, 5, 2) & "/01")), "mm")
            csvNJYm3 = Format(DateAdd("m", 2, CDate(Left(UYT.内示年月, 4) & "/" & Mid(UYT.内示年月, 5, 2) & "/01")), "mm")
            
            OutStr = "--*-----------------------------------------------------------------------------------*--,"
            Print #OutFNo, OutStr
            OutStr = Left(UYT.内示年月, 4) & "年" & Mid(UYT.内示年月, 5, 2) & "月度,,※　売上予想集計表【部品】※"
            Print #OutFNo, OutStr
            OutStr = "--*-----------------------------------------------------------------------------------*--,"
            Print #OutFNo, OutStr
            
            OutStr = "№,工区,契約先,契約先名,車型,,"
            OutStr = OutStr & csvNJYm1 & "月数量," & csvNJYm1 & "月金額," & csvNJYm1 & "月前月比,"
            OutStr = OutStr & csvNJYm2 & "月数量," & csvNJYm2 & "月金額," & csvNJYm2 & "月前月比,"
            OutStr = OutStr & csvNJYm3 & "月数量," & csvNJYm3 & "月金額"
            Print #OutFNo, OutStr
            OutCnt = 0
            Do Until Not UYT_RDSTS
                With UYT
                    OutCnt = OutCnt + 1
                    OutStr = OutCnt & ","
                    OutStr = OutStr & RTrim(.工区) & ","
                    OutStr = OutStr & RTrim(.契約先) & ","
                    OutStr = OutStr & RTrim(.契約先名) & ","
                    OutStr = OutStr & RTrim(.車型) & ",,"
                    OutStr = OutStr & .数量1 & ","
                    OutStr = OutStr & .金額1 & ","
                    OutStr = OutStr & .前月比1 & ","
                    OutStr = OutStr & .数量2 & ","
                    OutStr = OutStr & .金額2 & ","
                    OutStr = OutStr & .前月比2 & ","
                    OutStr = OutStr & .数量3 & ","
                    OutStr = OutStr & .金額3 & ","
        
                    Print #OutFNo, OutStr
            
                    CisFrm.MeterUpdate OutCnt
        
                    Call UYTReadNext
                End With
            Loop
        End If
        Call UYTClose
    End If
'*---------------------------------------*
'*  売上予想CSV出力 【部品・明細】
'*---------------------------------------*
    If ChkH1_HinKb(0).Value <> 0 And ChkH1_PrtSyu(1).Value <> 0 Then
        
        Call SQLEdit0311(UY_HinKeta)
        If UYTRead(gSL_Select, 0) Then
            CisFrm.MeterMax = CisDB.RecordCount
            TotalCnt = TotalCnt + CisDB.RecordCount
            
            csvNJYm1 = Mid(UYT.内示年月, 5, 2)
            csvNJYm2 = Format(DateAdd("m", 1, CDate(Left(UYT.内示年月, 4) & "/" & Mid(UYT.内示年月, 5, 2) & "/01")), "mm")
            csvNJYm3 = Format(DateAdd("m", 2, CDate(Left(UYT.内示年月, 4) & "/" & Mid(UYT.内示年月, 5, 2) & "/01")), "mm")
            
            OutStr = "--*-----------------------------------------------------------------------------------*--,"
            Print #OutFNo, OutStr
            OutStr = Left(UYT.内示年月, 4) & "年" & Mid(UYT.内示年月, 5, 2) & "月度,,※　売上予想表【部品】※"
            Print #OutFNo, OutStr
            OutStr = "--*-----------------------------------------------------------------------------------*--,"
            Print #OutFNo, OutStr
            
            OutStr = "№,工区,契約先,契約先名,車型,品番,"
            OutStr = OutStr & csvNJYm1 & "月単価," & csvNJYm1 & "月仮区分," & csvNJYm1 & "月数量," & csvNJYm1 & "月金額," & csvNJYm1 & "月前月比,"
            OutStr = OutStr & csvNJYm2 & "月単価," & csvNJYm2 & "月仮区分," & csvNJYm2 & "月数量," & csvNJYm2 & "月金額," & csvNJYm2 & "月前月比,"
            OutStr = OutStr & csvNJYm3 & "月単価," & csvNJYm3 & "月仮区分," & csvNJYm3 & "月数量," & csvNJYm3 & "月金額"
            Print #OutFNo, OutStr
            OutCnt = 0
            Do Until Not UYT_RDSTS
                With UYT
                    OutCnt = OutCnt + 1
                    OutStr = OutCnt & ","
                    OutStr = OutStr & RTrim(.工区) & ","
                    OutStr = OutStr & RTrim(.契約先) & ","
                    OutStr = OutStr & RTrim(.契約先名) & ","
                    OutStr = OutStr & RTrim(.車型) & ","
                    OutStr = OutStr & EditHinbn(.品番, .品番編集) & ","
                    OutStr = OutStr & .単価1 & ","
                    Select Case .仮区分1
                        Case "1": gStr = "*"
                        Case "9": gStr = "未決"
                        Case Else: gStr = ""
                    End Select
                    OutStr = OutStr & gStr & ","
                    OutStr = OutStr & .数量1 & ","
                    OutStr = OutStr & .金額1 & ","
                    OutStr = OutStr & .前月比1 & ","
                    OutStr = OutStr & .単価2 & ","
                    Select Case .仮区分2
                        Case "1": gStr = "*"
                        Case "9": gStr = "未決"
                        Case Else: gStr = ""
                    End Select
                    OutStr = OutStr & gStr & ","
                    OutStr = OutStr & .数量2 & ","
                    OutStr = OutStr & .金額2 & ","
                    OutStr = OutStr & .前月比2 & ","
                    OutStr = OutStr & .単価3 & ","
                    Select Case .仮区分3
                        Case "1": gStr = "*"
                        Case "9": gStr = "未決"
                        Case Else: gStr = ""
                    End Select
                    OutStr = OutStr & gStr & ","
                    OutStr = OutStr & .数量3 & ","
                    OutStr = OutStr & .金額3 & ","
        
                    Print #OutFNo, OutStr
            
                    CisFrm.MeterUpdate OutCnt
        
                    Call UYTReadNext
                End With
            Loop
        End If
        Call UYTClose
    End If
'*---------------------------------------*
'*  売上予想CSV出力 【型費・集計】
'*---------------------------------------*
    If ChkH1_HinKb(1).Value <> 0 And ChkH1_PrtSyu(0).Value <> 0 Then
        
        Call SQLEdit0312
        If UYKRead(gSL_Select, 0) Then
            CisFrm.MeterMax = CisDB.RecordCount
            TotalCnt = TotalCnt + CisDB.RecordCount
            
            OutStr = "--*-----------------------------------------------------------------------------------*--,"
            Print #OutFNo, OutStr
            OutStr = Left(UYK.予想年月, 4) & "年" & Mid(UYK.予想年月, 5, 2) & "月度,,※　売上予想集計表【型費】※"
            Print #OutFNo, OutStr
            OutStr = "--*-----------------------------------------------------------------------------------*--,"
            Print #OutFNo, OutStr
            
            OutStr = "№,工区,契約先,契約先名,車型,,"
            For gInt = 1 To 12
                OutStr = OutStr & Format(DateAdd("m", gInt - 1, CDate(Left(UYK.開始年月, 4) & "/" & Mid(UYK.開始年月, 5, 2) & "/01")), "mm") & "月型費,"
            Next
            OutStr = OutStr & "上期型費,下期型費,期合計型費"
            Print #OutFNo, OutStr
            OutCnt = 0
            Do Until Not UYK_RDSTS
                With UYK
                    OutCnt = OutCnt + 1
                    OutStr = OutCnt & ","
                    OutStr = OutStr & RTrim(.工区) & ","
                    OutStr = OutStr & RTrim(.契約先) & ","
                    OutStr = OutStr & RTrim(.契約先名) & ","
                    OutStr = OutStr & RTrim(.車型) & ","
                    OutStr = OutStr & EditHinbn(.品番, .品番編集) & ","
                                    
                    For gInt = 1 To 12
                        OutStr = OutStr & .金額(gInt) & ","
                    Next
                    OutStr = OutStr & .上期計 & ","
                    OutStr = OutStr & .下期計 & ","
                    OutStr = OutStr & .期合計 & ""
        
                    Print #OutFNo, OutStr
            
                    CisFrm.MeterUpdate OutCnt
        
                    Call UYKReadNext
                End With
            Loop
        End If
        Call UYKClose
    End If
'*---------------------------------------*
'*  売上予想CSV出力 【型費・明細】
'*---------------------------------------*
    If ChkH1_HinKb(1).Value <> 0 And ChkH1_PrtSyu(1).Value <> 0 Then
        
        Call SQLEdit0313(UY_HinKeta)
        If UYKRead(gSL_Select, 0) Then
            CisFrm.MeterMax = CisDB.RecordCount
            TotalCnt = TotalCnt + CisDB.RecordCount
            
            OutStr = "--*-----------------------------------------------------------------------------------*--,"
            Print #OutFNo, OutStr
            OutStr = Left(UYK.予想年月, 4) & "年" & Mid(UYK.予想年月, 5, 2) & "月度,,※　売上予想表【型費】※"
            Print #OutFNo, OutStr
            OutStr = "--*-----------------------------------------------------------------------------------*--,"
            Print #OutFNo, OutStr
            
            OutStr = "№,工区,契約先,契約先名,車型,品番,"
            For gInt = 1 To 12
                OutStr = OutStr & Format(DateAdd("m", gInt - 1, CDate(Left(UYK.開始年月, 4) & "/" & Mid(UYK.開始年月, 5, 2) & "/01")), "mm") & "月型費,"
            Next
            OutStr = OutStr & "上期型費,下期型費,期合計型費"
            Print #OutFNo, OutStr
            OutCnt = 0
            Do Until Not UYK_RDSTS
                With UYK
                    OutCnt = OutCnt + 1
                    OutStr = OutCnt & ","
                    OutStr = OutStr & RTrim(.工区) & ","
                    OutStr = OutStr & RTrim(.契約先) & ","
                    OutStr = OutStr & RTrim(.契約先名) & ","
                    OutStr = OutStr & RTrim(.車型) & ","
                    OutStr = OutStr & EditHinbn(.品番, .品番編集) & ","
                                    
                    For gInt = 1 To 12
                        OutStr = OutStr & .金額(gInt) & ","
                    Next
                    OutStr = OutStr & .上期計 & ","
                    OutStr = OutStr & .下期計 & ","
                    OutStr = OutStr & .期合計 & ""
        
                    Print #OutFNo, OutStr
            
                    CisFrm.MeterUpdate OutCnt
        
                    Call UYKReadNext
                End With
            Loop
        End If
        Call UYKClose
    End If
CsvPut_Ed:

    Close #OutFNo
    SaveSetting appname:=gRegKey, section:="CSV", _
                    Key:=mCsvKey, setting:=RTrim(lb_PrtDev)
    
    Call CisFrm.MousePT(1)
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "   出力処理が終了しました。    "
        .MB_MSG(4) = "   　 " & "出力先    = " & mCsvDir
        .MB_MSG(5) = "   　 " & "出力件数 = " & StrConv(Format(TotalCnt, "##,##0"), vbWide) & " 件 "
        .MB_Title = "出力確認"
        .MB_Button = OK
        .MBOX
    End With
    CsvPut = True
    
End Function
