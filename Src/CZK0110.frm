VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BA5D7819-2B67-11D4-806A-00000EA57F9E}#3.0#0"; "CISYMD~1.OCX"
Begin VB.Form CZK0110 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "材料在庫金額一覧発行"
   ClientHeight    =   9420
   ClientLeft      =   2190
   ClientTop       =   3405
   ClientWidth     =   14520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   9420
   ScaleWidth      =   14520
   Begin VB.PictureBox Csv_Bar 
      Height          =   345
      Left            =   4260
      ScaleHeight     =   285
      ScaleWidth      =   6600
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   6660
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   585
      Left            =   5640
      Top             =   6810
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
      Height          =   4230
      Left            =   4260
      ScaleHeight     =   4340.21
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6585
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2025
      Width           =   6645
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   4065
         Index           =   1
         Left            =   75
         Top             =   45
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   7170
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
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   330
            Left            =   1620
            TabIndex        =   1
            Top             =   900
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cChkResult      =   0   'False
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "　出　力　先"
            Height          =   600
            Left            =   60
            TabIndex        =   16
            Top             =   2970
            Width           =   6345
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "プリンター"
               Height          =   240
               Index           =   0
               Left            =   1065
               TabIndex        =   6
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
               TabIndex        =   7
               Top             =   270
               Width           =   1875
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "CSV"
               Height          =   240
               Index           =   2
               Left            =   4380
               TabIndex        =   8
               Top             =   240
               Width           =   975
            End
         End
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   1620
            TabIndex        =   2
            Top             =   1395
            Width           =   1005
            _ExtentX        =   1773
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   2970
            Top             =   1395
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
               TabIndex        =   10
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
                  Caption         =   "材料在庫金額一覧の発行を行います"
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
            Top             =   3615
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
            Top             =   3645
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
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   300
            Index           =   0
            Left            =   840
            Top             =   1410
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "手配先"
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
         Begin Cis3D_v60.CIS3D PB_Torcd 
            Height          =   360
            Left            =   2625
            Top             =   1395
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
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   300
            Index           =   1
            Left            =   840
            Top             =   2025
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "材　質"
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
         Begin CisText_V60.CisText iH1_Zaist 
            Height          =   360
            Left            =   1620
            TabIndex        =   3
            Top             =   1980
            Width           =   3270
            _ExtentX        =   5768
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   300
            Left            =   3540
            Top             =   2580
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "～"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
            cAlingnment     =   1
            cBoderStyle     =   2
            cPositionX      =   20
         End
         Begin CisYMD_V60.CisYMD iH1_YMDS 
            Height          =   375
            Left            =   1620
            TabIndex        =   4
            Top             =   2535
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
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   300
            Index           =   1
            Left            =   420
            Top             =   2580
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   529
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
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   300
            Index           =   2
            Left            =   645
            Top             =   900
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
            ForeColor       =   16711680
            Caption         =   "指定年月"
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
         Begin Cis3D_v60.CIS3D PB_Zaist 
            Height          =   360
            Left            =   4845
            Top             =   1980
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
         Begin CisYMD_V60.CisYMD iH1_YMDE 
            Height          =   375
            Left            =   3840
            TabIndex        =   5
            Top             =   2535
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
      Caption         =   "【 材料在庫金額一覧発行 】"
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
      TabIndex        =   9
      Text            =   "Dummy"
      Top             =   -90
      Width           =   345
   End
   Begin Cis3D_v60.CIS3D PB_Ent 
      Height          =   585
      Left            =   8280
      Top             =   6810
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
      Top             =   6810
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
      Left            =   4260
      Top             =   4050
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Cis3D_v60.CIS3D L_TestWaku 
      Height          =   1140
      Left            =   8520
      Top             =   1920
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   2011
      Caption         =   "確認用"
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
      cFont3D         =   2
      cAlingnment     =   6
      cPositionY      =   15
      Begin VB.Label L_TestDisp 
         BorderStyle     =   1  '実線
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   2160
      End
   End
   Begin VB.Label PNL_Ok 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6900
      TabIndex        =   15
      Top             =   7485
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8280
      TabIndex        =   13
      Top             =   7485
      Width           =   1485
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5610
      TabIndex        =   12
      Top             =   7485
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
      TabIndex        =   11
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CZK0110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  材料在庫金額一覧 発行
'**       フォームID    :  CZK0110
'**       処理概要      :
'**
'**       作  成  日    :  2010/01/06
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
    
    Dim mTanka          As Double           ' 単価          ' 2010.01.12 add
    Dim mTBL_TRCd()     As String           ' 手配先        ' 2010.01.12 add
  ''Dim mTBL_YYMM()     As String           ' 年月度        ' 2010.01.12 add    ' DEL
    Dim mTBL_Tana()     As String           ' 実棚日        ' 2010.01.12 add
    Dim mTBL_ZKNo()     As Integer          ' 材料管理番号  ' 2010.01.12 add
    Dim mTBL_Kazu       As Integer          ' データ数      ' 2010.01.12 add
    Dim mHostName       As String           ' 端末名        ' 2010.01.12 add
    
'   試験用
    Dim mzaikosu        As Double           ' 2010.01.07 add
    Const mDebugFlg     As Boolean = False  ' 2010.01.07 add

'   背景色
    Dim mSelColor       As Long
    Dim mUnSelColor     As Long
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
        .MSGTYPE = Left_To_Right
        .MSGText = mSaveMsg
        .MSGInterval = 100
        .MSGRelate
    End With

'   #------------------#
'   #  端末名取得      #
'   #------------------#
    gSL_Select = "SELECT SUBSTRING(HOST_NAME(), 1, 20) AS 処理端末"
    Call ZKWRead(gSL_Select)
    mHostName = ZKW.処理端末
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
     
    L_TestWaku.Visible = mDebugFlg          ' 2010.01.08 add : デバッグモード表示
    L_TestDisp.Visible = mDebugFlg          ' 2010.01.07 add : デバッグモード表示
    iH1_YM = Format(Now(), "yyyyMM")        ' 2010.01.07 add
    Call iH1_YM.SetFocus                    ' 2010.01.06 add
    
'   CSVﾌｧｲﾙ出力先
    mCsvName = Mid(CisFun.WinRoot, 1, 3) & "材料在庫金額一覧表.csv"
    mCsvKey = "COK0530"
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
        .Gradation 200, 130, 200, 4, 2, 3
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
    If Dummy.Tag = "iH1_Torcd" Then Call PB_Torcd_Click: Exit Sub
    If Dummy.Tag = "iH1_Zaist" Then Call PB_Zaist_Click: Exit Sub
End Sub
'+---------------------+
'+    取引先検索        +
'+---------------------+
Private Sub PB_Torcd_Click()
''''    RV_Left = 0
''''    RV_Top = 0
''''    RV_Call = "C"
''''    RV_TorKb = 8
'''''
''''    CKK0025.Show vbModal
''''    Unload CKK0025
''''    Set CKK0025 = Nothing
''''
''''    iH1_Torcd.SetFocus
''''
''''    If RV_Rtn Then
''''        iH1_Torcd = RV_TorcdK
''''        H1lb_Tornm = RV_TorRName
''''    End If
'---------------------------------------------------' 2010.01.06 add start (上記分と差替)
    RV_Left = 0                 ' 取引先検索
    RV_Top = 0
    RV_TorKb = 1                '   発注
    RV_TorKb2 = 0               '   社内非限定

    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
       iH1_Torcd = RV_Torcd
       Call iH1_Torcd_LostFocus
    End If
'---------------------------------------------------' 2010.01.06 add end   (上記分と差替)
End Sub
'---------------------------------------------------' 2010.01.06 add start
'+----------------------+
'+       材質検索       +
'+----------------------+
Private Sub PB_Zaist_Click()
    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 1
    RVI_SK = 9
    RVI_Tehai = ""
    RVI_Torcd = Trim(iH1_Torcd)

    CKK0077.Show vbModal
    Unload CKK0077
    Set CKK0077 = Nothing

    If RV_Rtn Then
       iH1_Zaist = RV_Zaist
       Call iH1_Zaist.SetFocus
    End If
End Sub
'---------------------------------------------------' 2010.01.06 add end

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
    H1lb_Tornm = ""
    If RTrim(iH1_Torcd) <> "" Then
        If TorNmGet(iH1_Torcd, 0) Then  ' 2010.01.06 upd 8-->0
            H1lb_Tornm = TRM_RName
        End If
    End If
End Sub
'+----------------------------+
'+    iH1_Zaist(GotFocus)    +
'+----------------------------+
Private Sub iH1_Zaist_GotFocus()
    Dummy.Tag = ActiveControl.Name
End Sub
'+----------------------------+
'+    iH1_Zaist(LostFocus)    +
'+----------------------------+
Private Sub iH1_Zaist_LostFocus()
    Dummy.Tag = ""
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
    Head1Chk = False

    Dim wCnt1   As Integer              ' 2010.01.12 add
    Dim wCnt2   As Integer              ' 2010.01.12 add
    
'   指定年月
    If Trim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    指定年月を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If
    If Not iH1_YM.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    指定年月を正しく入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If

'   手配先
    If Trim(iH1_Torcd) <> "" Then
        If Not TorNmGet(iH1_Torcd, 9) Then      ' 2010.01.06 upd : 8-->0
        With CisFun
            .MB_Lines = 4                       ' 2010.01.07 upd : 2-->4
            .MB_MSG(2) = "    手配先 未登録                     "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
    End If
    
'-----------------------------------------------------------------------------------------------------------' 2010.01.07 add start
'   材質
    If Trim(iH1_Zaist) <> "" Then
        gSL_Select = "SELECT * FROM 材料マスタ WHERE 材質 = '" & Trim(iH1_Zaist) & "'"
        If Not ZRMRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    材質 未登録                       "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
            iH1_Zaist.SetFocus
            Exit Function
        End If
    End If

'   最終出庫日
    If Trim(iH1_YMDS) <> "" And Not iH1_YMDS.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    最終出庫日(開始)を正しく入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YMDS.SetFocus
        Exit Function
    End If
    If Trim(iH1_YMDE) <> "" And Not iH1_YMDE.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    最終出庫日(終了)を正しく入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YMDE.SetFocus
        Exit Function
    End If
'-----------------------------------------------------------------------------------------------------------' 2010.01.07 add end

    lb_msg.Caption = "【 材料在庫金額一覧データ 抽出中 】"
    DoEvents
    
    With CisDB
        '+---------- ワークの削除 -----------+
        .SQL = ""
        .SQL = .SQL & " DELETE FROM 材料在庫金額一覧ワーク"
        .SQL = .SQL & "  WHERE 処理端末 = SUBSTRING(HOST_NAME(), 1, 20)"
        Call CisDB.DBExec
    End With
    
'---------------------------------------------------------------------------------------------------------------' 2010.01.07 add start
''''gSL_Select = "SELECT * FROM 材料在庫マスタ WHERE 年月度 = '" & Trim(iH1_YM) & "'"
''''If Trim(iH1_Zaist) <> "" Then
''''    gSL_Select = gSL_Select & "   AND 材質 = '" & Trim(iH1_Zaist) & "'"
''''End If
    
    gSL_Select = ""
'    gSL_Select = gSL_Select & "SELECT ZZ.年月度, ZZ.材料管理番号, ZZ.材質, ZZ.板厚, ZZ.幅, ZZ.長さ, ZZ.在庫数,"
'    gSL_Select = gSL_Select & "       ZZ.実棚日, ZT.取引先 AS 変更事由   , ZZ.SIZE, ZZ.作成区分"    ' 取引先を変更事由で代用する
'    gSL_Select = gSL_Select & " FROM  材料在庫マスタ ZZ"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料取引先マスタ ZT"
'    gSL_Select = gSL_Select & "            ON  ZT.材料管理番号 = zz.材料管理番号"
'    gSL_Select = gSL_Select & "            AND ZT.材質         = ZZ.材質"
'    gSL_Select = gSL_Select & "            AND ZT.板厚         = ZZ.板厚"
'    gSL_Select = gSL_Select & "            AND ZT.幅           = ZZ.幅"
'    gSL_Select = gSL_Select & "            AND ZT.長さ         = ZZ.長さ"
'    gSL_Select = gSL_Select & " WHERE ZZ.年月度 = '" & Trim(iH1_YM) & "'"
'    If Trim(iH1_Zaist) <> "" Then gSL_Select = gSL_Select & "   AND ZZ.材質   = '" & Trim(iH1_Zaist) & "'"
'    If Trim(iH1_Torcd) <> "" Then gSL_Select = gSL_Select & "   AND ZT.取引先 = '" & Trim(iH1_Torcd) & "'"
'    gSL_Select = gSL_Select & " ORDER BY ZT.材料管理番号"
    gSL_Select = gSL_Select & "SELECT ZZ.年月度 AS 納品書形態名, ZT.材料管理番号, ZT.材質, ZT.板厚, ZT.幅, ZT.長さ, ZZ.在庫数 AS 在庫管理数値,"
    gSL_Select = gSL_Select & "       ZZ.実棚日 AS 受領書形態名, ZT.取引先      , ZM.表示寸法, ZZ.作成区分 AS 在庫管理区分"
    gSL_Select = gSL_Select & " FROM  材料取引先マスタ ZT"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料在庫マスタ ZZ"
    gSL_Select = gSL_Select & "            ON  ZT.材料管理番号 = ZZ.材料管理番号"
'    gSL_Select = gSL_Select & "            AND ZT.材質         = ZZ.材質"
'    gSL_Select = gSL_Select & "            AND ZT.板厚         = ZZ.板厚"
'    gSL_Select = gSL_Select & "            AND ZT.幅           = ZZ.幅"
'    gSL_Select = gSL_Select & "            AND ZT.長さ         = ZZ.長さ"
    gSL_Select = gSL_Select & "            AND ZZ.年月度       = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料マスタ ZM"
    gSL_Select = gSL_Select & "            ON  ZM.材料管理番号 = ZT.材料管理番号"
'    gSL_Select = gSL_Select & " WHERE ZZ.年月度 = '" & Trim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " WHERE ISNULL(ZT.棚卸区分,0) = 0 "
    If Trim(iH1_Zaist) <> "" Then gSL_Select = gSL_Select & "   AND ZZ.材質   = '" & Trim(iH1_Zaist) & "'"
    If Trim(iH1_Torcd) <> "" Then gSL_Select = gSL_Select & "   AND ZT.取引先 = '" & Trim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " ORDER BY ZT.材料管理番号"

    If Not ZTMRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    該当データがありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Call ZTMClose
        Exit Function
    
    Else
        Do Until Not ZTM_RDSTS
           '在庫年月(納品書形態名を使用)にて在庫ﾏｽﾀ存在判断
           If RTrim(ZTM.納品書形態名) = "" Then
            ' 在庫マスタ[無]のとき
               Call ZaikoKeisan1
           Else
            ' 在庫マスタ[有]のとき
               mTBL_Kazu = CisDB.RecordCount
               ReDim mTBL_ZKNo(mTBL_Kazu)
               ReDim mTBL_Tana(mTBL_Kazu)
               ReDim mTBL_TRCd(mTBL_Kazu)
               For wCnt2 = 0 To mTBL_Kazu - 1
''                   mTBL_ZKNo(wCnt1) = ZTM.材料管理番号
''                   mTBL_Tana(wCnt1) = ZTM.受領書形態名     ' 実棚日は「受領書形態名」で代用
''                   mTBL_TRCd(wCnt1) = ZTM.取引先
                   mTBL_ZKNo(wCnt2) = ZTM.材料管理番号
                   mTBL_Tana(wCnt2) = ZTM.受領書形態名     ' 実棚日は「受領書形態名」で代用
                   mTBL_TRCd(wCnt2) = ZTM.取引先
                   Call ZTMReadNext
               Next
               Call ZaikoKeisan2
           End If
           
           Call ZTMReadNext
        Loop
    End If
    
' 旧ロジック
'    If Not ZTMRead(gSL_Select) Then
'    ' 在庫マスタ[無]のとき
'        Call ZaikoKeisan1
'    Else
'    ' 在庫マスタ[有]のとき
'        mTBL_Kazu = CisDB.RecordCount
'        ReDim mTBL_ZKNo(mTBL_Kazu)
'        ReDim mTBL_Tana(mTBL_Kazu)
'        ReDim mTBL_TRCd(mTBL_Kazu)
'        For wCnt2 = 0 To mTBL_Kazu - 1
'            mTBL_ZKNo(wCnt1) = ZTM.材料管理番号
'            mTBL_Tana(wCnt1) = ZTM.受領書形態名     ' 実棚日は「受領書形態名」で代用
'            mTBL_TRCd(wCnt1) = ZTM.取引先
'            Call ZTMReadNext
'        Next
'        Call ZaikoKeisan2
'    End If
'---------------------------------------------------------------------------------------------------------------' 2010.01.07 add end

'---テスト--------------------------------------------------------------------------------------------------' 2010.01.07 add start
GoTo jump0091
Dim wTestCnt As Integer
Dim wHostNam As String
gSL_Select = "SELECT SUBSTRING(HOST_NAME(), 1, 20) AS 処理端末"
Call ZKWRead(gSL_Select)
wHostNam = ZKW.処理端末
With ZKW
For wTestCnt = 1 To 90
.処理端末 = wHostNam '"FMV-D5255-PC"
.材質 = "材質" & Format(wTestCnt, "00")
.寸法 = "寸法" & Format(wTestCnt, "00")
.手配先名 = "手配先" & Format(wTestCnt, "00")
.仕入単価 = 100 + wTestCnt
.前月末在庫数 = 1000 + wTestCnt
.前月末在庫金額 = 1100 + wTestCnt
.当月入庫数 = 2000 + wTestCnt
.当月入庫金額 = 2100 + wTestCnt
.当月出庫数 = 3000 + wTestCnt
.当月出庫金額 = 3100 + wTestCnt
.調整数 = 4000 + wTestCnt
.調整金額 = 4100 + wTestCnt
.現在在庫数 = 5000 + wTestCnt
.現在在庫金額 = 5100 + wTestCnt
.最終入庫日 = "201001" & Format(wTestCnt, "00")
Call ZKWInsert
Next
End With
jump0091:
'---テスト--------------------------------------------------------------------------------------------------' 2010.01.07 add end
       
  ''材料在庫金額一覧ワーク確認
    gSL_Select = "SELECT * FROM 材料在庫金額一覧ワーク"                             ' 2010.01.07 add
    gSL_Select = gSL_Select & " WHERE 処理端末 = SUBSTRING(HOST_NAME(), 1, 20)"     ' 2010.01.07 add
    gSL_Select = gSL_Select & " ORDER BY 材質,寸法,手配先 "
    If Not ZKWRead(gSL_Select, , 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    該当データがありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If

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
        lb_msg.Caption = "【 材料在庫金額一覧表 印刷中 】"
        DoEvents
        CZK0110_YM = Mid(iH1_YM, 1, 4) & "年" & Mid(iH1_YM, 5, 2) & "月"
        With CisPrt
            .cDataMaxCnt = CisDB.RecordCount
            Set .iPrtForm = CZK0110L
            .dpMSG = "材料在庫金額一覧表　印刷中"
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

    lb_msg.Caption = "【 材料在庫金額一覧データ 出力中 】"
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
    Do Until Not ZKW_RDSTS
    '   >> ｶｳﾝﾄ
        OutCnt = OutCnt + 1
    '   >> Body部 編集＆出力
        Call CsvEdit_Body1(OutStr, OutCnt)
        Print #OutFNo, OutStr

        CisFrm.MeterUpdate OutCnt

    '   >> 次のﾚｺｰﾄﾞを読み込む
        Call ZKWReadNext(1)
    Loop
    Call ZKWClose(1)
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
    CsvStr = CsvStr & Mid(RTrim(iH1_YM), 1, 4) & "年" & Mid(RTrim(iH1_YM), 5, 2) & "月        ※ 材料在庫金額一覧表 ※,,"
  ''CsvStr = CsvStr & ",,※ 材料在庫金額一覧表 ※"
    CsvStr = CsvStr & ",,,,,,,,,,," & "発行日：" & Format(Now(), "yyyy/MM/dd")
End Sub
Private Sub CsvEdit_Head2(CsvStr As String)
    CsvStr = ""
    CsvStr = CsvStr & ",,,,<前月末在庫>,,<当月入庫>,,<当月出庫>,,<調整数>,,<現在在庫>" & vbCrLf
    CsvStr = CsvStr & "材質"
    CsvStr = CsvStr & ",寸法"
    CsvStr = CsvStr & ",手配先"
    CsvStr = CsvStr & ",仕入単価"
    CsvStr = CsvStr & ",数量"       ' 前月末在庫
    CsvStr = CsvStr & ",金額"       ' 前月末在庫
    CsvStr = CsvStr & ",数量"       ' 当月入庫
    CsvStr = CsvStr & ",金額"       ' 当月入庫
    CsvStr = CsvStr & ",数量"       ' 当月出庫
    CsvStr = CsvStr & ",金額"       ' 当月出庫
    CsvStr = CsvStr & ",数量"       ' 調整数
    CsvStr = CsvStr & ",金額"       ' 調整数
    CsvStr = CsvStr & ",数量"       ' 現在在庫
    CsvStr = CsvStr & ",金額"       ' 現在在庫
    CsvStr = CsvStr & ",最終入庫日"
End Sub
'+----------------------------+
'+    Body部　編集処理
'+----------------------------+
Private Sub CsvEdit_Body1(CsvStr As String, CsvCnt As Long)
    With ZKW
        CsvStr = RTrim(.材質)
        CsvStr = CsvStr & "," & RTrim(.寸法)
        CsvStr = CsvStr & "," & RTrim(.手配先名)
        CsvStr = CsvStr & "," & Format(.仕入単価, "#####0.00")
        CsvStr = CsvStr & "," & Format(.前月末在庫数, "#####0")
        CsvStr = CsvStr & "," & Format(.前月末在庫金額, "#####0")
        CsvStr = CsvStr & "," & Format(.当月入庫数, "#####0")
        CsvStr = CsvStr & "," & Format(.当月入庫金額, "#####0")
        CsvStr = CsvStr & "," & Format(.当月出庫数, "#####0")
        CsvStr = CsvStr & "," & Format(.当月出庫金額, "#####0")
        CsvStr = CsvStr & "," & Format(.調整数, "#####0")
        CsvStr = CsvStr & "," & Format(.調整金額, "#####0")
        CsvStr = CsvStr & "," & Format(.現在在庫数, "#####0")
        CsvStr = CsvStr & "," & Format(.現在在庫金額, "#####0")
        CsvStr = CsvStr & "," & " " & Format(.最終入庫日, "0000/00/00")
    End With
End Sub

'-----------------------------------------------------------------------------------------------------------' 2010.01.07 add start
'+----------------------------+
'+  材料在庫マスタ無の処理
'+----------------------------+
Private Function ZaikoKeisan1() As Boolean
    ZaikoKeisan1 = False
        
    Dim wSuIn   As Double
    Dim wSuOut  As Double
    wSuIn = 0:  wSuOut = 0
    
    On Error GoTo ZaikoKeisan1_Err
    ' [代用項目メモ]　納品書形態名：年月度 ／ 在庫管理数値：在庫数 ／ 受領書形態名：実棚日     ／ 在庫管理区分：作成区分
    '                 実績金額0 　：出庫数 ／ 実績金額5 　：単価   ／ 実績金額2 　：前月末在庫

'    gSL_Select = ""
'    gSL_Select = gSL_Select & "SELECT TOP(1)"
'    gSL_Select = gSL_Select & "       ISNULL( ZNT.入荷数, 0 ) + ISNULL( ZZT.実績数, 0 ) AS 入荷数"
'    gSL_Select = gSL_Select & " FROM  材料入荷テーブル ZN"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
'    gSL_Select = gSL_Select & "           (SELECT NK.材料管理番号, SUM(NK.入荷数) AS 入荷数 FROM 材料入荷テーブル NK"
'    gSL_Select = gSL_Select & "             WHERE SUBSTRING(NK.入荷日, 1, 6) >= '" & Mid(iH1_YM, 1, 6) & "'"
'    gSL_Select = gSL_Select & "               AND NK.材料管理番号 =  " & ZRM.材料管理番号 'ZZM.材料管理番号
'    gSL_Select = gSL_Select & "             GROUP BY NK.材料管理番号) ZNT"
'    gSL_Select = gSL_Select & "         ON ZNT.材料管理番号 = ZN.材料管理番号"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
'    gSL_Select = gSL_Select & "           (SELECT ZZ.実績区分, SUM(ZZ.実績数) AS 実績数 FROM 材料実績管理テーブル ZZ"
'    gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZZ.実績日, 1,6) >= '" & Mid(iH1_YM, 1, 6) & "'"
'    gSL_Select = gSL_Select & "               AND ZZ.材質     = '" & Trim(ZRM.材質) & "'"   'Trim(ZZM.材質) & "'"
'    gSL_Select = gSL_Select & "               AND ZZ.板厚     =  " & ZRM.板厚               'ZZM.板厚
'    gSL_Select = gSL_Select & "               AND ZZ.幅       =  " & ZRM.幅                 'ZZM.幅
'    gSL_Select = gSL_Select & "               AND ZZ.長さ     =  " & ZRM.長さ               'ZZM.長さ
'    gSL_Select = gSL_Select & "               AND ZZ.実績区分 =  2"
'    gSL_Select = gSL_Select & "             GROUP BY ZZ.実績区分) ZZT"
'    gSL_Select = gSL_Select & "         ON ZZT.実績区分 = 2"
'    gSL_Select = gSL_Select & " WHERE SUBSTRING(ZN.入荷日, 1, 6) >= '" & Mid(iH1_YM, 1, 6) & "'"
'    gSL_Select = gSL_Select & "   AND ZN.材料管理番号 =  " & ZRM.材料管理番号 'ZZM.材料管理番号
'    If ZYTRead(gSL_Select) Then wSuIn = ZYT.入荷数
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT ZT.材料管理番号 , ZT.取引先 AS 仕入先, TM.略称 AS 仕入先名"
    gSL_Select = gSL_Select & "      ,ZT.材質, ZT.板厚, ZT.幅 , ZT.長さ, ZM.表示寸法"
    gSL_Select = gSL_Select & "      ,ZNT1.入荷数,  ZZT1.実績数"
    gSL_Select = gSL_Select & "      ,ISNULL(ZNT1.入荷数,0) + ISNULL(ZZT1.実績数,0) AS 入荷数"
    gSL_Select = gSL_Select & "      ,ZNT2.入荷日 AS 入荷入荷日,ZZT2.実績日 AS 実績入荷日"
    gSL_Select = gSL_Select & "      ,CASE WHEN ZNT2.入荷日 >= ZZT2.実績日  THEN ZNT2.入荷日 ELSE ZZT2.実績日 END 入荷日"
    gSL_Select = gSL_Select & "      ,ISNULL(ZSR1.使用量,0) + ISNULL(ZZT3.実績数出,0) AS 実績金額0" ' 使用量
    gSL_Select = gSL_Select & "      ,ISNULL(ZR.在庫数,0) AS 実績金額2"                             ' 前月末在庫
    gSL_Select = gSL_Select & " FROM  材料取引先マスタ ZT"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料在庫マスタ ZZ"
    gSL_Select = gSL_Select & "        ON  ZZ.材質   <> ''"
''''gSL_Select = gSL_Select & "        AND ZZ.棚卸区分= 0"              ' 2010.01.18 add
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
    gSL_Select = gSL_Select & "           (SELECT ZN1.材料管理番号,SUM(ZN1.入荷数) AS 入荷数 FROM 材料入荷テーブル ZN1"
    gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZN1.入荷日, 1, 6) = '" & iH1_YM & "'"
    gSL_Select = gSL_Select & "             GROUP BY ZN1.材料管理番号) ZNT1"
    gSL_Select = gSL_Select & "        ON  ZNT1.材料管理番号 = ZZ.材料管理番号"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
''    gSL_Select = gSL_Select & "           (SELECT TOP(1)ZN2.* FROM 材料入荷テーブル ZN2"
''    gSL_Select = gSL_Select & "             ORDER BY ZN2.入荷日 DESC) ZNT2"
    gSL_Select = gSL_Select & "           (SELECT max(ZN2.入荷日) AS 入荷日,材料管理番号 FROM 材料入荷テーブル ZN2"
    gSL_Select = gSL_Select & "             GROUP BY 材料管理番号) ZNT2"
    gSL_Select = gSL_Select & "        ON  ZNT2.材料管理番号 = ZZ.材料管理番号"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
    gSL_Select = gSL_Select & "           (SELECT ZZ1.実績区分,ZZ1.材質,ZZ1.板厚,ZZ1.幅,ZZ1.長さ,SUM(ZZ1.実績数) AS 実績数 FROM 材料実績管理テーブル ZZ1"
    gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZZ1.実績日, 1, 6) = '" & iH1_YM & "'"
''''gSL_Select = gSL_Select & "               AND ZZ1.実績日 < '20091231'"                  ' 入庫計算時は不要===============
    gSL_Select = gSL_Select & "             GROUP BY ZZ1.実績区分, ZZ1.材質, ZZ1.板厚, ZZ1.幅, ZZ1.長さ) ZZT1"
''  gSL_Select = gSL_Select & "        ON  ZZT1.実績区分 = 2"           ' 2010.01.18 del
    gSL_Select = gSL_Select & "        ON  ZZT1.実績区分 = 1"           ' 2010.01.18 add
    gSL_Select = gSL_Select & "        AND ZZT1.材質     = ZZ.材質"
    gSL_Select = gSL_Select & "        AND ZZT1.板厚     = ZZ.板厚"
    gSL_Select = gSL_Select & "        AND ZZT1.幅       = ZZ.幅"
    gSL_Select = gSL_Select & "        AND ZZT1.長さ     = ZZ.長さ"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
    gSL_Select = gSL_Select & "           (SELECT TOP(1)ZZ2.* FROM 材料実績管理テーブル ZZ2"
    gSL_Select = gSL_Select & "             ORDER BY ZZ2.実績日 DESC) ZZT2"
''  gSL_Select = gSL_Select & "        ON  ZZT2.実績区分 = 2"           ' 2010.01.18 del
    gSL_Select = gSL_Select & "        ON  ZZT2.実績区分 = 1"           ' 2010.01.18 add
    gSL_Select = gSL_Select & "        AND ZZT2.材質     = ZZ.材質"
    gSL_Select = gSL_Select & "        AND ZZT2.板厚     = ZZ.板厚"
    gSL_Select = gSL_Select & "        AND ZZT2.幅       = ZZ.幅"
    gSL_Select = gSL_Select & "        AND ZZT2.長さ     = ZZ.長さ"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TM"
    gSL_Select = gSL_Select & "        ON  TM.取引先CD = ZT.取引先"
    gSL_Select = gSL_Select & "        AND TM.取引先区分 = 1"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料マスタ ZM"
    gSL_Select = gSL_Select & "        ON  ZM.材質   = ZT.材質"
    gSL_Select = gSL_Select & "        AND ZM.板厚   = ZT.板厚"
    gSL_Select = gSL_Select & "        AND ZM.幅     = ZT.幅"
    gSL_Select = gSL_Select & "        AND ZM.長さ   = ZT.長さ"
    '---------------------------------------------------' 2010.01.14 add start
''    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
''    gSL_Select = gSL_Select & "           (SELECT TOP(1) * FROM 材料在庫履歴マスタ WHERE 実棚日='' ORDER BY 入力日付 DESC) ZR"
''    gSL_Select = gSL_Select & "        ON  ZR.年月度       = '" & Trim(iH1_YM) & "'"
''    gSL_Select = gSL_Select & "        AND ZR.材料管理番号 = ZT.材料管理番号"
''    gSL_Select = gSL_Select & "        AND ZR.材質         = ZT.材質"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT 入力日付,年月度,材料管理番号,材質,在庫数"
    gSL_Select = gSL_Select & "                         FROM  材料在庫履歴マスタ WHERE 実棚日=''"
    gSL_Select = gSL_Select & "                         GROUP BY 入力日付,年月度,材料管理番号,材質,在庫数) ZR"
    gSL_Select = gSL_Select & "        ON  ZR.年月度       = '" & Trim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "        AND ZR.材料管理番号 = ZT.材料管理番号"
    gSL_Select = gSL_Select & "        AND ZR.材質         = ZT.材質"
    '---------------------------------------------------' 2010.01.14 add end
    '---------------------------------------------------' 2010.01.12 add start
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
    gSL_Select = gSL_Select & "           (SELECT ZS1.材料管理番号,SUM(ZS1.使用量) AS 使用量 FROM 材料使用量テーブル ZS1"
    gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZS1.作業日, 1, 6) = '" & iH1_YM & "'"
    

''''If Trim(iH1_YMDS) <> "" Then gSL_Select = gSL_Select & "               AND ZS1.作業日 <= '" & iH1_YMDS & "'"
    If Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) = "" Then
       gSL_Select = gSL_Select & "               AND (ZS1.作業日 >= '" & iH1_YMDS & "' AND"
       gSL_Select = gSL_Select & "                    ISNULL(ZS1.使用量,0) = 0)"
    ElseIf Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) <> "" Then
       gSL_Select = gSL_Select & "               AND (ZS1.作業日 >= '" & iH1_YMDS & "' AND"
       gSL_Select = gSL_Select & "                    ZS1.作業日 <= '" & iH1_YMDE & "' AND"
       gSL_Select = gSL_Select & "                    ISNULL(ZS1.使用量,0) = 0)"
    ElseIf Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) = "" Then
       gSL_Select = gSL_Select & "               AND (ZS1.作業日 <= '" & iH1_YMDE & "' AND"
       gSL_Select = gSL_Select & "                    ISNULL(ZS1.使用量,0) = 0)"
    End If

    gSL_Select = gSL_Select & "             GROUP BY ZS1.材料管理番号) ZSR1"
    gSL_Select = gSL_Select & "              ON  ZSR1.材料管理番号 = ZZ.材料管理番号"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
    gSL_Select = gSL_Select & "           (SELECT ZZ3.実績区分,ZZ3.入力種類,ZZ3.材質,ZZ3.板厚,ZZ3.幅,ZZ3.長さ,SUM(ZZ3.実績数) AS 実績数出 FROM 材料実績管理テーブル ZZ3"    ' 2010.01.18 upd :入力種類を追加
    gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZZ3.実績日, 1, 6) = '" & iH1_YM & "'"
    

''''If Trim(iH1_YMDS) <> "" Then gSL_Select = gSL_Select & "               AND ZZ3.実績日 <= '" & iH1_YMDS & "'"
    If Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) = "" Then
       gSL_Select = gSL_Select & "               AND (ZZ3.実績日 >= '" & iH1_YMDS & "' AND"
       gSL_Select = gSL_Select & "                    ISNULL(ZZ3.実績数,0) = 0)"
    ElseIf Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) <> "" Then
       gSL_Select = gSL_Select & "               AND (ZZ3.実績日 >= '" & iH1_YMDS & "' AND"
       gSL_Select = gSL_Select & "                    ZZ3.実績日 <= '" & iH1_YMDE & "' AND"
       gSL_Select = gSL_Select & "                    ISNULL(ZZ3.実績数,0) = 0)"
    ElseIf Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) = "" Then
       gSL_Select = gSL_Select & "               AND (ZZ3.実績日 <= '" & iH1_YMDE & "' AND"
       gSL_Select = gSL_Select & "                    ISNULL(ZZ3.実績数,0) = 0)"
            End If

    gSL_Select = gSL_Select & "             GROUP BY ZZ3.実績区分, ZZ3.入力種類, ZZ3.材質, ZZ3.板厚, ZZ3.幅, ZZ3.長さ) ZZT3"    ' 2010.01.18 upd : 入力種類を追加
''  gSL_Select = gSL_Select & "        ON  ZZT3.実績区分 = 3"       ' 2010.01.18 del
    gSL_Select = gSL_Select & "        ON  ZZT3.実績区分 <> 1"      ' 2010.01.18 add
    gSL_Select = gSL_Select & "        AND ZZT3.入力種類 <> 'A'"    ' 2010.01.18 add
    gSL_Select = gSL_Select & "        AND ZZT3.材質     = ZZ.材質"
    gSL_Select = gSL_Select & "        AND ZZT3.板厚     = ZZ.板厚"
    gSL_Select = gSL_Select & "        AND ZZT3.幅       = ZZ.幅"
    gSL_Select = gSL_Select & "        AND ZZT3.長さ     = ZZ.長さ"
    '---------------------------------------------------' 2010.01.12 add end
    gSL_Select = gSL_Select & " WHERE ZT.材料管理番号 = ZZ.材料管理番号"
    If Trim(iH1_Torcd) <> "" Then gSL_Select = gSL_Select & "   AND ZT.取引先       = '" & Trim(iH1_Torcd) & "'"
    If Trim(iH1_Zaist) <> "" Then gSL_Select = gSL_Select & "   AND ZT.材質         = '" & Trim(iH1_Zaist) & "'"
    gSL_Select = gSL_Select & "   AND ZT.棚卸区分 = 0"  ' 2010.01.18 add
    gSL_Select = gSL_Select & " GROUP BY ZT.材料管理番号, ZT.取引先, ZT.材質, ZT.板厚, ZT.幅, ZT.長さ"
    gSL_Select = gSL_Select & "         ,ZZ.材料管理番号, TM.略称  , ZM.表示寸法"
    gSL_Select = gSL_Select & "         ,ZNT1.入荷数, ZNT2.入荷日"
    gSL_Select = gSL_Select & "         ,ZZT1.材質  , ZZT1.板厚, ZZT1.幅, ZZT1.長さ,ZZT1.実績区分,ZZT1.実績数"
    gSL_Select = gSL_Select & "         ,ZZT2.材質  , ZZT2.板厚, ZZT2.幅, ZZT2.長さ,ZZT2.実績日"
    gSL_Select = gSL_Select & "         ,ZSR1.使用量,ZZT3.実績数出,ZR.在庫数"  ' 2010.01.12 add

    If ZYTRead(gSL_Select, , 0) Then
        wSuIn = ZYT.入荷数
        wSuOut = ZYT.実績金額0
    Else
        GoTo ZaikoKeisan1_Err
    End If

''''Call ItemsClearZYT  ' 項目内容確認用
''''GoTo jump0011
'''''    gSL_Select = ""
'''''    gSL_Select = gSL_Select & "SELECT TOP(1)"
'''''    gSL_Select = gSL_Select & "       ISNULL( ZST.入荷数, 0 ) + ISNULL( ZZT.実績数, 0 ) AS 使用量"
'''''    gSL_Select = gSL_Select & " FROM  材料使用量テーブル ZS"
'''''    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
'''''    gSL_Select = gSL_Select & "           (SELECT SR.材料管理番号, SUM(SR.使用量) AS 入荷数 FROM 材料使用量テーブル SR"
'''''    gSL_Select = gSL_Select & "             WHERE SUBSTRING(SR.作業日, 1, 6) >= '" & Mid(iH1_YM, 1, 6) & "'"
'''''    gSL_Select = gSL_Select & "               AND SR.材料管理番号 =  " & ZRM.材料管理番号   'ZZM.材料管理番号       ' mZRKanriNo
'''''    gSL_Select = gSL_Select & "             GROUP BY SR.材料管理番号) ZST"
'''''    gSL_Select = gSL_Select & "         ON ZST.材料管理番号 = ZS.材料管理番号"
'''''    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
'''''    gSL_Select = gSL_Select & "           (SELECT ZZ.実績区分, SUM(ZZ.実績数) AS 実績数 FROM 材料実績管理テーブル ZZ"
'''''    gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZZ.実績日, 1, 6) >= '" & Mid(iH1_YM, 1, 6) & "'"
'''''    gSL_Select = gSL_Select & "               AND ZZ.材質     = '" & Trim(ZRM.材質) & "'"   'Trim(ZZM.材質) & "'"
'''''    gSL_Select = gSL_Select & "               AND ZZ.板厚     =  " & ZRM.板厚               'ZZM.板厚
'''''    gSL_Select = gSL_Select & "               AND ZZ.幅       =  " & ZRM.幅                 'ZZM.幅
'''''    gSL_Select = gSL_Select & "               AND ZZ.長さ     =  " & ZRM.長さ               'ZZM.長さ
'''''    gSL_Select = gSL_Select & "               AND ZZ.実績区分 =  3"
'''''    gSL_Select = gSL_Select & "             GROUP BY ZZ.実績区分) ZZT"
'''''    gSL_Select = gSL_Select & "         ON ZZT.実績区分 = 3"
'''''    gSL_Select = gSL_Select & " WHERE SUBSTRING(ZS.作業日, 1, 6) >= '" & Mid(iH1_YM, 1, 6) & "'"
'''''    gSL_Select = gSL_Select & "   AND ZS.材料管理番号 =  " & ZRM.材料管理番号 'ZZM.材料管理番号                     ' mZRKanriNo
''''    gSL_Select = ""
''''    gSL_Select = gSL_Select & "SELECT ZT.材料管理番号 , ZT.取引先 AS 仕入先, TM.略称 AS 部門名"
''''    gSL_Select = gSL_Select & "      ,ZT.材質, ZT.板厚, ZT.幅 , ZT.長さ, ZM.表示寸法 AS 表示材料"
''''    gSL_Select = gSL_Select & "      ,ZSR1.使用量,  ZZT1.実績数"
''''    gSL_Select = gSL_Select & "      ,ISNULL(ZSR1.使用量,0) + ISNULL(ZZT1.実績数,0) AS 使用量"
''''    gSL_Select = gSL_Select & " FROM  材料取引先マスタ ZT"
''''    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料在庫マスタ ZZ"
''''    gSL_Select = gSL_Select & "        ON  ZZ.材質   <> ''"
''''    gSL_Select = gSL_Select & "        AND ZZ.棚卸区分= 0"
''''    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
''''    gSL_Select = gSL_Select & "           (SELECT ZS1.材料管理番号,SUM(ZS1.使用量) AS 使用量 FROM 材料使用量テーブル ZS1"
''''    gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZS1.作業日, 1, 6) = '" & iH1_YM & "'" ''200912'"
''''    If Trim(iH1_YMDS) <> "" Then
''''       gSL_Select = gSL_Select & "               AND ZS1.作業日 < '" & Trim(iH1_YMDS) & "'"
''''    End If
''''    gSL_Select = gSL_Select & "             GROUP BY ZS1.材料管理番号) ZSR1"
''''    gSL_Select = gSL_Select & "         ON ZSR1.材料管理番号 = ZZ.材料管理番号"
''''    gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
''''    gSL_Select = gSL_Select & "           (SELECT ZZ1.実績区分,ZZ1.材質,ZZ1.板厚,ZZ1.幅,ZZ1.長さ,SUM(ZZ1.実績数) AS 実績数 FROM 材料実績管理テーブル ZZ1"
''''    gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZZ1.実績日, 1, 6) = '" & iH1_YM & "'" ''200912'"
''''    If Trim(iH1_YMDS) <> "" Then
''''       gSL_Select = gSL_Select & "               AND ZZ1.実績日 < '" & Trim(iH1_YMDS) & "'"
''''    End If
''''    gSL_Select = gSL_Select & "             GROUP BY ZZ1.実績区分, ZZ1.材質, ZZ1.板厚, ZZ1.幅, ZZ1.長さ) ZZT1"
''''    gSL_Select = gSL_Select & "         ON  ZZT1.実績区分 = 3"
''''    gSL_Select = gSL_Select & "         AND ZZT1.材質     = ZZ.材質"
''''    gSL_Select = gSL_Select & "         AND ZZT1.板厚     = ZZ.板厚"
''''    gSL_Select = gSL_Select & "         AND ZZT1.幅       = ZZ.幅"
''''    gSL_Select = gSL_Select & "         AND ZZT1.長さ     = ZZ.長さ"
''''    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TM"
''''    gSL_Select = gSL_Select & "        ON  TM.取引先 = ZT.取引先"
''''    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料マスタ ZM"
''''    gSL_Select = gSL_Select & "        ON  ZM.材質   = ZT.材質"
''''    gSL_Select = gSL_Select & "        AND ZM.板厚   = ZT.板厚"
''''    gSL_Select = gSL_Select & "        AND ZM.幅     = ZT.幅"
''''    gSL_Select = gSL_Select & "        AND ZM.長さ   = ZT.長さ"
''''    gSL_Select = gSL_Select & " WHERE ZT.材料管理番号 = ZZ.材料管理番号"
''''    gSL_Select = gSL_Select & " GROUP BY ZT.材料管理番号, ZT.取引先, ZT.材質, ZT.板厚, ZT.幅, ZT.長さ"
''''    gSL_Select = gSL_Select & "         ,ZZ.材料管理番号, TM.略称  , ZM.表示寸法"
''''    gSL_Select = gSL_Select & "         ,ZSR1.使用量"
''''    gSL_Select = gSL_Select & "         ,ZZT1.材質  , ZZT1.板厚, ZZT1.幅, ZZT1.長さ,ZZT1.実績区分,ZZT1.実績数"
''''''''----現状：年月の特定の仕方が不明、手配先、材質が条件に入っていない----''''
''''    If ZSRRead(gSL_Select, , 1) Then wSuOut = ZSR.使用量
''''''''Call ItemsClearZSR  ' 項目内容確認用
''''
''''jump0011:
''''    '---確認表示用（デバッグ時のみ）
''''    mzaikosu = wSuIn - wSuOut
''''    If mDebugFlg Then L_TestDisp.Caption = "入庫 = " & Format(wSuIn, "00000000.000") _
''''                                & vbCrLf & "出庫 = " & Format(wSuOut, "00000000.000") _
''''                                & vbCrLf & "在庫 = " & Format(mzaikosu, "00000000.000")

'---ワーク作成--------
    Do While ZYT_RDSTS
        Call ItemsClearZKW
        Call GetShiireTanka(ZYT.材料管理番号, ZYT.仕入先)
        With ZKW
            .処理端末 = mHostName
            .材料管理番号 = ZYT.材料管理番号
            .材質 = ZYT.材質
            .寸法 = ZYT.表示寸法
            .手配先 = ZYT.仕入先
            .手配先名 = ZYT.仕入先名
            .仕入単価 = mTanka
            .前月末在庫数 = ZYT.実績金額2
            .前月末在庫金額 = .前月末在庫数 * .仕入単価
            .当月入庫数 = ZYT.入荷数
            .当月入庫金額 = .当月入庫数 * .仕入単価
            .当月出庫数 = ZYT.実績金額0
            .当月出庫金額 = .当月出庫数 * .仕入単価
            .調整数 = 0
            .調整金額 = 0
            .現在在庫数 = .前月末在庫数 + .当月入庫数 - .当月出庫数
            .現在在庫金額 = .現在在庫数 * .仕入単価
            .最終入庫日 = ZYT.入荷日
            Call ZKWInsert(1)
            Call ZYTReadNext(0)
        End With
    Loop

ZaikoKeisan1_ED:
    On Error GoTo 0
    ZaikoKeisan1 = True
    Exit Function
ZaikoKeisan1_Err:
    On Error GoTo 0
    Exit Function
End Function

'+----------------------------+
'+  材料在庫マスタ有の処理
'+----------------------------+
Private Function ZaikoKeisan2() As Boolean
    ZaikoKeisan2 = False
    
    Dim wCnt1   As Integer
    Dim wNendo  As String   ' 年月度
    Dim wJTDay  As String   ' 実棚日
    Dim wBango  As Long     ' 材料管理番号
    Dim wSuIn   As Double
    Dim wSuOut  As Double
    wSuIn = 0:  wSuOut = 0
    
    On Error GoTo ZaikoKeisan2_Err
    ' [代用項目メモ]　納品書形態名：年月度 ／ 在庫管理数値：在庫数 ／ 受領書形態名：実棚日     ／ 在庫管理区分：作成区分
    '                 実績金額0 　：出庫数 ／ 実績金額5 　：単価   ／ 実績金額2 　：前月末在庫

    For wCnt1 = 0 To mTBL_Kazu - 1
      ''wNendo = mTBL_YYMM(wCnt1)  'ZTM.納品書形態名 ' 年月度  ' ZZM.年月度
        wJTDay = mTBL_Tana(wCnt1)  'ZTM.受領書形態名 ' 実棚日  ' ZZM.実棚日
        wBango = mTBL_ZKNo(wCnt1)
        Call GetShiireTanka(wBango, mTBL_TRCd(wCnt1))
        
        If Trim(wJTDay) = "" Then
        ' 実棚日が空白の場合の処理
            gSL_Select = ""
            gSL_Select = gSL_Select & "SELECT ZT.材料管理番号 , ZT.取引先 AS 仕入先, TM.略称 AS 仕入先名"
            gSL_Select = gSL_Select & "      ,ZT.材質, ZT.板厚, ZT.幅 , ZT.長さ, ZM.表示寸法      ,ZNT1.入荷数,  ZZT1.実績数, ZZ.在庫数"
            gSL_Select = gSL_Select & "      ,ISNULL(ZNT1.入荷数,0) + ISNULL(ZZT1.実績数,0) AS 入荷数"  ' + ISNULL(ZZ.在庫数,0)
            gSL_Select = gSL_Select & "      ,ZNT2.入荷日 AS 入荷日,ZZT2.実績日 AS 実績入荷日"
          ''gSL_Select = gSL_Select & "      ,CASE WHEN ZNT2.入荷日 >= ZZT2.実績日  THEN ZNT2.入荷日 ELSE ZZT2.実績日 END 入荷日"
            gSL_Select = gSL_Select & "      ,ISNULL(ZSR1.使用量,0) + ISNULL(ZZT3.実績数出,0) AS 実績金額0"
            gSL_Select = gSL_Select & "      ," & Str(mTanka) & " AS 実績金額5, ISNULL(ZR.在庫数,0) AS 実績金額2"
            gSL_Select = gSL_Select & "      ,ZR.入力日付"
            gSL_Select = gSL_Select & "  FROM  材料取引先マスタ ZT"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料在庫マスタ ZZ"
          ''gSL_Select = gSL_Select & "      --ON  ZZ.材質   <> ''"
            gSL_Select = gSL_Select & "        ON  ZZ.材料管理番号 = " & Str(wBango)
            gSL_Select = gSL_Select & "        AND ZZ.年月度 = '" & iH1_YM & "'"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
            gSL_Select = gSL_Select & "           (SELECT ZN1.材料管理番号,SUM(ZN1.入荷数) AS 入荷数 FROM 材料入荷テーブル ZN1"
            gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZN1.入荷日, 1, 6) = '" & iH1_YM & "'"
            gSL_Select = gSL_Select & "             GROUP BY ZN1.材料管理番号) ZNT1"
            gSL_Select = gSL_Select & "        ON  ZNT1.材料管理番号 = Zt.材料管理番号"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
''            gSL_Select = gSL_Select & "           (SELECT TOP(1) ZN2.* FROM 材料入荷テーブル ZN2"
''            gSL_Select = gSL_Select & "             ORDER BY ZN2.入荷日 DESC) ZNT2"
            gSL_Select = gSL_Select & "           (SELECT max(ZN2.入荷日) AS 入荷日,材料管理番号 FROM 材料入荷テーブル ZN2"
            gSL_Select = gSL_Select & "             GROUP BY 材料管理番号) ZNT2"
            gSL_Select = gSL_Select & "        ON  ZNT2.材料管理番号 = Zt.材料管理番号"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
            gSL_Select = gSL_Select & "           (SELECT ZZ1.実績区分,ZZ1.材質,ZZ1.板厚,ZZ1.幅,ZZ1.長さ,SUM(ZZ1.実績数) AS 実績数 FROM 材料実績管理テーブル ZZ1"
            gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZZ1.実績日, 1, 6) = '" & iH1_YM & "'"
            gSL_Select = gSL_Select & "             GROUP BY ZZ1.実績区分, ZZ1.材質, ZZ1.板厚, ZZ1.幅, ZZ1.長さ) ZZT1"
''          gSL_Select = gSL_Select & "        ON  ZZT1.実績区分 = 2"           ' 2010.01.18 del
            gSL_Select = gSL_Select & "        ON  ZZT1.実績区分 = 1"           ' 2010.01.18 add
            gSL_Select = gSL_Select & "        AND ZZT1.材質     = ZZ.材質"
            gSL_Select = gSL_Select & "        AND ZZT1.板厚     = ZZ.板厚"
            gSL_Select = gSL_Select & "        AND ZZT1.幅       = ZZ.幅"
            gSL_Select = gSL_Select & "        AND ZZT1.長さ     = ZZ.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
            gSL_Select = gSL_Select & "           (SELECT TOP(1)ZZ2.* FROM 材料実績管理テーブル ZZ2"
            gSL_Select = gSL_Select & "             ORDER BY ZZ2.実績日 DESC) ZZT2"
''          gSL_Select = gSL_Select & "        ON  ZZT2.実績区分 = 2"           ' 2010.01.18 del
            gSL_Select = gSL_Select & "        ON  ZZT2.実績区分 = 1"           ' 2010.01.18 add
            gSL_Select = gSL_Select & "        AND ZZT2.材質     = ZZ.材質"
            gSL_Select = gSL_Select & "        AND ZZT2.板厚     = ZZ.板厚"
            gSL_Select = gSL_Select & "        AND ZZT2.幅       = ZZ.幅"
            gSL_Select = gSL_Select & "        AND ZZT2.長さ     = ZZ.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TM"
            gSL_Select = gSL_Select & "        ON  TM.取引先CD = ZT.取引先"
            gSL_Select = gSL_Select & "        AND TM.取引先区分 = 1"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料マスタ ZM"
            gSL_Select = gSL_Select & "        ON  ZM.材質   = ZT.材質"
            gSL_Select = gSL_Select & "        AND ZM.板厚   = ZT.板厚"
            gSL_Select = gSL_Select & "        AND ZM.幅     = ZT.幅"
            gSL_Select = gSL_Select & "        AND ZM.長さ   = ZT.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
            gSL_Select = gSL_Select & "           (SELECT ZS1.材料管理番号,SUM(ZS1.使用量) AS 使用量 FROM 材料使用量テーブル ZS1"
            gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZS1.作業日, 1, 6) = '" & iH1_YM & "'"


''''        If Trim(iH1_YMDS) <> "" Then gSL_Select = gSL_Select & "               AND ZS1.作業日 <= '" & iH1_YMDS & "'"
            If Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) = "" Then
               gSL_Select = gSL_Select & "               AND (ZS1.作業日 >= '" & iH1_YMDS & "' AND"
               gSL_Select = gSL_Select & "                    ISNULL(ZS1.使用量,0) = 0)"
            ElseIf Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) <> "" Then
               gSL_Select = gSL_Select & "               AND (ZS1.作業日 >= '" & iH1_YMDS & "' AND"
               gSL_Select = gSL_Select & "                    ZS1.作業日 <= '" & iH1_YMDE & "' AND"
               gSL_Select = gSL_Select & "                    ISNULL(ZS1.使用量,0) = 0)"
            ElseIf Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) = "" Then
               gSL_Select = gSL_Select & "               AND (ZS1.作業日 <= '" & iH1_YMDE & "' AND"
               gSL_Select = gSL_Select & "                    ISNULL(ZS1.使用量,0) = 0)"
            End If
            
            gSL_Select = gSL_Select & "             GROUP BY ZS1.材料管理番号) ZSR1              ON  ZSR1.材料管理番号 = ZZ.材料管理番号"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
            gSL_Select = gSL_Select & "           (SELECT ZZ3.実績区分,ZZ3.入力種類,ZZ3.材質,ZZ3.板厚,ZZ3.幅,ZZ3.長さ,SUM(ZZ3.実績数) AS 実績数出 FROM 材料実績管理テーブル ZZ3"    ' 2010.01.18 upd :入力種類を追加
            gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZZ3.実績日, 1, 6) = '" & iH1_YM & "'"
            

''''        If Trim(iH1_YMDS) <> "" Then gSL_Select = gSL_Select & "               AND ZZ3.実績日 <= '" & iH1_YMDS & "'"
            If Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) = "" Then
               gSL_Select = gSL_Select & "               AND (ZZ3.実績日 >= '" & iH1_YMDS & "' AND"
               gSL_Select = gSL_Select & "                    ISNULL(ZZ3.実績数,0) = 0)"
            ElseIf Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) <> "" Then
               gSL_Select = gSL_Select & "               AND (ZZ3.実績日 >= '" & iH1_YMDS & "' AND"
               gSL_Select = gSL_Select & "                    ZZ3.実績日 <= '" & iH1_YMDE & "' AND"
               gSL_Select = gSL_Select & "                    ISNULL(ZZ3.実績数,0) = 0)"
            ElseIf Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) = "" Then
               gSL_Select = gSL_Select & "               AND (ZZ3.実績日 <= '" & iH1_YMDE & "' AND"
               gSL_Select = gSL_Select & "                    ISNULL(ZZ3.実績数,0) = 0)"
            End If
            
            gSL_Select = gSL_Select & "             GROUP BY ZZ3.実績区分, ZZ3.入力種類, ZZ3.材質, ZZ3.板厚, ZZ3.幅, ZZ3.長さ) ZZT3"    ' 2010.01.18 upd :入力種類を追加
''          gSL_Select = gSL_Select & "        ON  ZZT3.実績区分 = 3"           ' 2010.01.18 del
            gSL_Select = gSL_Select & "        ON  ZZT3.実績区分 <> 1"          ' 2010.01.18 add
            gSL_Select = gSL_Select & "        AND ZZT3.入力種類 <> 'A'"        ' 2010.01.18 add
            gSL_Select = gSL_Select & "        AND ZZT3.材質     = ZZ.材質"
            gSL_Select = gSL_Select & "        AND ZZT3.板厚     = ZZ.板厚"
            gSL_Select = gSL_Select & "        AND ZZT3.幅       = ZZ.幅"
            gSL_Select = gSL_Select & "        AND ZZT3.長さ     = ZZ.長さ"
            '---------------------------------------------------' 2010.01.14 add start
'            gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
'            gSL_Select = gSL_Select & "           (SELECT TOP(1) * FROM 材料在庫履歴マスタ WHERE 実棚日='' ORDER BY 入力日付 DESC) ZR"
'            gSL_Select = gSL_Select & "        ON  ZR.年月度       = '" & Trim(iH1_YM) & "'"
'            gSL_Select = gSL_Select & "        AND ZR.材料管理番号 = ZT.材料管理番号"
'            gSL_Select = gSL_Select & "        AND ZR.材質         = ZT.材質"
'            gSL_Select = gSL_Select & "        AND ZR.板厚         = ZT.板厚"
'            gSL_Select = gSL_Select & "        AND ZR.幅           = ZT.幅"
'            gSL_Select = gSL_Select & "        AND ZR.長さ         = ZT.長さ"
            gSL_Select = gSL_Select & "    LEFT OUTER JOIN ( SELECT ZR1.入力日付,ZR1.材料管理番号,ZR1.材質,ZR1.板厚,ZR1.幅,ZR1.長さ,ZR1.在庫数"
            gSL_Select = gSL_Select & "     FROM 材料在庫履歴マスタ ZR1"
            gSL_Select = gSL_Select & "     LEFT OUTER JOIN ( SELECT MAX(入力日付) 入力日付,材質,板厚,幅,長さ"
            gSL_Select = gSL_Select & "                 From 材料在庫履歴マスタ"
'            gSL_Select = gSL_Select & "                WHERE 年月度 = '201001'"
            gSL_Select = gSL_Select & "                WHERE 年月度 = '" & RTrim(iH1_YM) & "'"
            gSL_Select = gSL_Select & "                  GROUP BY 材質,板厚,幅,長さ ) ZR2"
            gSL_Select = gSL_Select & "        ON ZR1.入力日付 = ZR2.入力日付"
            gSL_Select = gSL_Select & "        AND ZR1.材質 = ZR2.材質"
            gSL_Select = gSL_Select & "        AND ZR1.板厚 = ZR2.板厚"
            gSL_Select = gSL_Select & "        AND ZR1.幅 = ZR2.幅"
            gSL_Select = gSL_Select & "        AND ZR1.長さ = ZR2.長さ"
'            gSL_Select = gSL_Select & "    WHERE ZR1.年月度 = '201001'"
            gSL_Select = gSL_Select & "    WHERE ZR1.年月度 = '" & RTrim(iH1_YM) & "'"
            gSL_Select = gSL_Select & "    AND ISNULL(ZR2.入力日付 ,'') <> '' ) ZR"
            gSL_Select = gSL_Select & "        ON ZR.材料管理番号 = ZT.材料管理番号"
            gSL_Select = gSL_Select & "        AND ZR.材質         = ZT.材質        AND ZR.板厚         = ZT.板厚        AND ZR.幅           = ZT.幅"
            gSL_Select = gSL_Select & "        AND ZR.長さ         = ZT.長さ"
            '---------------------------------------------------' 2010.01.14 add end
            gSL_Select = gSL_Select & " WHERE ZT.材料管理番号 = ZZ.材料管理番号"
            If Trim(iH1_Torcd) <> "" Then gSL_Select = gSL_Select & "   AND ZT.取引先       = '" & Trim(iH1_Torcd) & "'"
            If Trim(iH1_Zaist) <> "" Then gSL_Select = gSL_Select & "   AND ZT.材質         = '" & Trim(iH1_Zaist) & "'"
            gSL_Select = gSL_Select & "   AND ZT.棚卸区分= 0"   ' 2010.01.18 add
            gSL_Select = gSL_Select & " GROUP BY ZT.材料管理番号, ZT.取引先, ZT.材質, ZT.板厚, ZT.幅, ZT.長さ"
            gSL_Select = gSL_Select & "         ,ZZ.材料管理番号, TM.略称  , ZM.表示寸法"
            gSL_Select = gSL_Select & "         ,ZNT1.入荷数, ZNT2.入荷日"
            gSL_Select = gSL_Select & "         ,ZZT1.材質  , ZZT1.板厚, ZZT1.幅, ZZT1.長さ,ZZT1.実績区分,ZZT1.実績数"
            gSL_Select = gSL_Select & "         ,ZZT2.材質  , ZZT2.板厚, ZZT2.幅, ZZT2.長さ,ZZT2.実績日"
            gSL_Select = gSL_Select & "         ,ZSR1.使用量, ZZT3.実績数出, ZZ.在庫数, ZR.在庫数"
            gSL_Select = gSL_Select & "         ,ZR.入力日付"
            gSL_Select = gSL_Select & " ORDER BY ZR.入力日付 DESC"
        Else
        ' 実棚日が入力済の場合の処理
            gSL_Select = ""
            gSL_Select = gSL_Select & "SELECT ZT.材料管理番号 , ZT.取引先 AS 仕入先, TM.略称 AS 仕入先名"
            gSL_Select = gSL_Select & "      ,ZT.材質, ZT.板厚, ZT.幅 , ZT.長さ, ZM.表示寸法      ,ZNT1.入荷数,  ZZT1.実績数, ZZ.在庫数"
            gSL_Select = gSL_Select & "      ,ISNULL(ZNT1.入荷数,0) + ISNULL(ZZT1.実績数,0) + ISNULL(ZZ.在庫数,0) AS 入荷数"
            gSL_Select = gSL_Select & "      ,ZNT2.入荷日 AS 入荷日,ZZT2.実績日 AS 実績入荷日"
            ''gSL_Select = gSL_Select & "      ,CASE WHEN ZNT2.入荷日 >= ZZT2.実績日  THEN ZNT2.入荷日 ELSE ZZT2.実績日 END 入荷日"
            gSL_Select = gSL_Select & "      ,ISNULL(ZSR1.使用量,0) + ISNULL(ZZT3.実績数出,0) AS 実績金額0"
            gSL_Select = gSL_Select & "      ," & Str(mTanka) & " AS 実績金額5, ISNULL(ZR.在庫数,0) AS 実績金額2"
            gSL_Select = gSL_Select & "  FROM  材料取引先マスタ ZT"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料在庫マスタ ZZ"
          ''gSL_Select = gSL_Select & "      --ON  ZZ.材質   <> ''"
            gSL_Select = gSL_Select & "        ON  ZZ.材料管理番号 = " & Str(wBango)
            gSL_Select = gSL_Select & "        AND ZZ.年月度 = '" & iH1_YM & "'"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
            gSL_Select = gSL_Select & "           (SELECT ZN1.材料管理番号,SUM(ZN1.入荷数) AS 入荷数 FROM 材料入荷テーブル ZN1"
            gSL_Select = gSL_Select & "             WHERE '" & Str(wJTDay) & "' < ZN1.入荷日"
            gSL_Select = gSL_Select & "             GROUP BY ZN1.材料管理番号) ZNT1"
            gSL_Select = gSL_Select & "         ON ZNT1.材料管理番号 = ZZ.材料管理番号"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
''            gSL_Select = gSL_Select & "           (SELECT TOP(1)ZN2.* FROM 材料入荷テーブル ZN2"
''            gSL_Select = gSL_Select & "             ORDER BY ZN2.入荷日 DESC) ZNT2"
            gSL_Select = gSL_Select & "           (SELECT max(ZN2.入荷日) AS 入荷日,材料管理番号 FROM 材料入荷テーブル ZN2"
            gSL_Select = gSL_Select & "             GROUP BY 材料管理番号) ZNT2"
            gSL_Select = gSL_Select & "         ON ZNT2.材料管理番号 = ZZ.材料管理番号"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
            gSL_Select = gSL_Select & "           (SELECT ZZ1.実績区分,ZZ1.材質,ZZ1.板厚,ZZ1.幅,ZZ1.長さ,SUM(ZZ1.実績数) AS 実績数 FROM 材料実績管理テーブル ZZ1"
            gSL_Select = gSL_Select & "             WHERE '" & wJTDay & "' < ZZ1.実績日"
            gSL_Select = gSL_Select & "             GROUP BY ZZ1.実績区分, ZZ1.材質, ZZ1.板厚, ZZ1.幅, ZZ1.長さ) ZZT1"
''          gSL_Select = gSL_Select & "         ON  ZZT1.実績区分 = 2"          ' 2010.01.18 del
            gSL_Select = gSL_Select & "         ON  ZZT1.実績区分 = 1"          ' 2010.01.18 add
            gSL_Select = gSL_Select & "         AND ZZT1.材質     = ZZ.材質"
            gSL_Select = gSL_Select & "         AND ZZT1.板厚     = ZZ.板厚"
            gSL_Select = gSL_Select & "         AND ZZT1.幅       = ZZ.幅"
            gSL_Select = gSL_Select & "         AND ZZT1.長さ     = ZZ.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
            gSL_Select = gSL_Select & "           (SELECT TOP(1)ZZ2.* FROM 材料実績管理テーブル ZZ2"
            gSL_Select = gSL_Select & "             ORDER BY ZZ2.実績日 DESC) ZZT2"
''          gSL_Select = gSL_Select & "         ON  ZZT2.実績区分 = 2"          ' 2010.01.18 del
            gSL_Select = gSL_Select & "         ON  ZZT2.実績区分 = 1"          ' 2010.01.18 add
            gSL_Select = gSL_Select & "         AND ZZT2.材質     = ZZ.材質"
            gSL_Select = gSL_Select & "         AND ZZT2.板厚     = ZZ.板厚"
            gSL_Select = gSL_Select & "         AND ZZT2.幅       = ZZ.幅"
            gSL_Select = gSL_Select & "         AND ZZT2.長さ     = ZZ.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TM"
            gSL_Select = gSL_Select & "        ON  TM.取引先CD = ZT.取引先"
            gSL_Select = gSL_Select & "        AND TM.取引先区分 = 1"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料マスタ ZM"
            gSL_Select = gSL_Select & "        ON  ZM.材質   = ZT.材質"
            gSL_Select = gSL_Select & "        AND ZM.板厚   = ZT.板厚"
            gSL_Select = gSL_Select & "        AND ZM.幅     = ZT.幅"
            gSL_Select = gSL_Select & "        AND ZM.長さ   = ZT.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
            gSL_Select = gSL_Select & "           (SELECT ZS1.材料管理番号,SUM(ZS1.使用量) AS 使用量 FROM 材料使用量テーブル ZS1"
''            gSL_Select = gSL_Select & "             WHERE '" & wJTDay & "' < ZS1.作業日"
            gSL_Select = gSL_Select & "             WHERE SUBSTRING(ZS1.作業日, 1, 6) = '" & iH1_YM & "'"
            

''''        If Trim(iH1_YMDS) <> "" Then gSL_Select = gSL_Select & "               AND ZS1.作業日 <= '" & Trim(iH1_YMDS) & "'"
            If Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) = "" Then
               gSL_Select = gSL_Select & "               AND (ZS1.作業日 >= '" & iH1_YMDS & "' AND"
               gSL_Select = gSL_Select & "                    ISNULL(ZS1.使用量,0) = 0)"
            ElseIf Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) <> "" Then
               gSL_Select = gSL_Select & "               AND (ZS1.作業日 >= '" & iH1_YMDS & "' AND"
               gSL_Select = gSL_Select & "                    ZS1.作業日 <= '" & iH1_YMDE & "' AND"
               gSL_Select = gSL_Select & "                    ISNULL(ZS1.使用量,0) = 0)"
            ElseIf Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) = "" Then
               gSL_Select = gSL_Select & "               AND (ZS1.作業日 <= '" & iH1_YMDE & "' AND"
               gSL_Select = gSL_Select & "                    ISNULL(ZS1.使用量,0) = 0)"
            End If
            
            gSL_Select = gSL_Select & "             GROUP BY ZS1.材料管理番号) ZSR1              ON  ZSR1.材料管理番号 = ZZ.材料管理番号"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
            gSL_Select = gSL_Select & "           (SELECT ZZ3.実績区分,ZZ3.入力種類,ZZ3.材質,ZZ3.板厚,ZZ3.幅,ZZ3.長さ,SUM(ZZ3.実績数) AS 実績数出 FROM 材料実績管理テーブル ZZ3"    ' 2010.01.18 upd :入力種類を追加
            gSL_Select = gSL_Select & "             WHERE '" & wJTDay & "'< ZZ3.実績日"
            

''''        If Trim(iH1_YMDS) <> "" Then gSL_Select = gSL_Select & "               AND ZZ3.実績日 <= '" & Trim(iH1_YMDS) & "'"
            If Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) = "" Then
               gSL_Select = gSL_Select & "               AND (ZZ3.実績日 >= '" & iH1_YMDS & "' AND"
               gSL_Select = gSL_Select & "                    ISNULL(ZZ3.実績数,0) = 0)"
            ElseIf Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) <> "" Then
               gSL_Select = gSL_Select & "               AND (ZZ3.実績日 >= '" & iH1_YMDS & "' AND"
               gSL_Select = gSL_Select & "                    ZZ3.実績日 <= '" & iH1_YMDE & "' AND"
               gSL_Select = gSL_Select & "                    ISNULL(ZZ3.実績数,0) = 0)"
            ElseIf Trim(iH1_YMDS) <> "" And Trim(iH1_YMDE) = "" Then
               gSL_Select = gSL_Select & "               AND (ZZ3.実績日 <= '" & iH1_YMDE & "' AND"
               gSL_Select = gSL_Select & "                    ISNULL(ZZ3.実績数,0) = 0)"
            End If
            
            gSL_Select = gSL_Select & "             GROUP BY ZZ3.実績区分, ZZ3.入力種類, ZZ3.材質, ZZ3.板厚, ZZ3.幅, ZZ3.長さ) ZZT3"    ' 2010.01.18 upd :入力種類を追加
''          gSL_Select = gSL_Select & "              ON  ZZT3.実績区分 = 3"         ' 2010.01.18 del
            gSL_Select = gSL_Select & "              ON  ZZT3.実績区分 <> 1"        ' 2010.01.18 add
            gSL_Select = gSL_Select & "              AND ZZT3.入力種類 <> 'A'"      ' 2010.01.18 add
            gSL_Select = gSL_Select & "              AND ZZT3.材質     = ZZ.材質"
            gSL_Select = gSL_Select & "              AND ZZT3.板厚     = ZZ.板厚"
            gSL_Select = gSL_Select & "              AND ZZT3.幅       = ZZ.幅"
            gSL_Select = gSL_Select & "              AND ZZT3.長さ     = ZZ.長さ"
            '---------------------------------------------------' 2010.01.14 add start
'            gSL_Select = gSL_Select & "       LEFT OUTER JOIN"
'            gSL_Select = gSL_Select & "           (SELECT TOP(1) * FROM 材料在庫履歴マスタ WHERE 実棚日<> '' ORDER BY 入力日付 DESC) ZR"
'            gSL_Select = gSL_Select & "              ON  ZR.年月度       = '" & Trim(iH1_YM) & "'"
'            gSL_Select = gSL_Select & "              AND ZR.材料管理番号 = ZT.材料管理番号"
'            gSL_Select = gSL_Select & "              AND ZR.材質         = ZT.材質"
'            gSL_Select = gSL_Select & "              AND ZR.板厚         = ZT.板厚"
'            gSL_Select = gSL_Select & "              AND ZR.幅           = ZT.幅"
'            gSL_Select = gSL_Select & "              AND ZR.長さ         = ZT.長さ"
            gSL_Select = gSL_Select & "    LEFT OUTER JOIN ( SELECT ZR1.入力日付,ZR1.材料管理番号,ZR1.材質,ZR1.板厚,ZR1.幅,ZR1.長さ,ZR1.在庫数"
            gSL_Select = gSL_Select & "     FROM 材料在庫履歴マスタ ZR1"
            gSL_Select = gSL_Select & "     LEFT OUTER JOIN ( SELECT MAX(入力日付) 入力日付,材質,板厚,幅,長さ"
            gSL_Select = gSL_Select & "                 From 材料在庫履歴マスタ"
'            gSL_Select = gSL_Select & "                WHERE 年月度 = '201001'"
            gSL_Select = gSL_Select & "                WHERE 年月度 = '" & RTrim(iH1_YM) & "'"
            gSL_Select = gSL_Select & "                  GROUP BY 材質,板厚,幅,長さ ) ZR2"
            gSL_Select = gSL_Select & "        ON ZR1.入力日付 = ZR2.入力日付"
            gSL_Select = gSL_Select & "        AND ZR1.材質 = ZR2.材質"
            gSL_Select = gSL_Select & "        AND ZR1.板厚 = ZR2.板厚"
            gSL_Select = gSL_Select & "        AND ZR1.幅 = ZR2.幅"
            gSL_Select = gSL_Select & "        AND ZR1.長さ = ZR2.長さ"
'            gSL_Select = gSL_Select & "    WHERE ZR1.年月度 = '201001'"
            gSL_Select = gSL_Select & "    WHERE ZR1.年月度 = '" & RTrim(iH1_YM) & "'"
            gSL_Select = gSL_Select & "    AND ISNULL(ZR2.入力日付 ,'') <> '' ) ZR"
            gSL_Select = gSL_Select & "        ON ZR.材料管理番号 = ZT.材料管理番号"
            gSL_Select = gSL_Select & "        AND ZR.材質         = ZT.材質        AND ZR.板厚         = ZT.板厚        AND ZR.幅           = ZT.幅"
            gSL_Select = gSL_Select & "        AND ZR.長さ         = ZT.長さ"
            '---------------------------------------------------' 2010.01.14 add end
            gSL_Select = gSL_Select & " WHERE ZT.材料管理番号 = ZZ.材料管理番号"
            If Trim(iH1_Torcd) <> "" Then gSL_Select = gSL_Select & "   AND ZT.取引先       = '" & Trim(iH1_Torcd) & "'"
            If Trim(iH1_Zaist) <> "" Then gSL_Select = gSL_Select & "   AND ZT.材質         = '" & Trim(iH1_Zaist) & "'"
            gSL_Select = gSL_Select & "   AND ZT.棚卸区分= 0"   ' 2010.01.18 add
            gSL_Select = gSL_Select & " GROUP BY ZT.材料管理番号, ZT.取引先, ZT.材質, ZT.板厚, ZT.幅, ZT.長さ"
            gSL_Select = gSL_Select & "         ,ZZ.材料管理番号, TM.略称  , ZM.表示寸法"
            gSL_Select = gSL_Select & "         ,ZNT1.入荷数, ZNT2.入荷日"
            gSL_Select = gSL_Select & "         ,ZZT1.材質  , ZZT1.板厚, ZZT1.幅, ZZT1.長さ,ZZT1.実績区分,ZZT1.実績数"
            gSL_Select = gSL_Select & "         ,ZZT2.材質  , ZZT2.板厚, ZZT2.幅, ZZT2.長さ,ZZT2.実績日"
            gSL_Select = gSL_Select & "         ,ZSR1.使用量, ZZT3.実績数出, ZZ.在庫数, ZR.在庫数"
            gSL_Select = gSL_Select & "         ,ZR.入力日付"
            gSL_Select = gSL_Select & " ORDER BY ZR.入力日付 DESC"
        End If
        
        If ZYTRead(gSL_Select, , 0) Then
            wSuIn = ZYT.入荷数
            wSuOut = ZYT.実績金額0
        Else
            GoTo ZaikoKeisan2_Err
        End If
    
''''        '---確認表示用（デバッグ時のみ）
''''        mzaikosu = wSuIn - wSuOut
''''        If mDebugFlg Then L_TestDisp.Caption = "入庫 = " & Format(wSuIn, "00000000.000") _
''''                                    & vbCrLf & "出庫 = " & Format(wSuOut, "00000000.000") _
''''                                    & vbCrLf & "在庫 = " & Format(mzaikosu, "00000000.000")
        '---ワーク作成--------
        Do While ZYT_RDSTS
            Call ItemsClearZKW
            With ZKW
                .処理端末 = mHostName
                .材料管理番号 = ZYT.材料管理番号
                .材質 = ZYT.材質
                .寸法 = ZYT.表示寸法
                .手配先 = ZYT.仕入先
                .手配先名 = ZYT.仕入先名
                .仕入単価 = ZYT.実績金額5                   ' 単価 mTanka
                .前月末在庫数 = ZYT.実績金額2
                .前月末在庫金額 = .前月末在庫数 * .仕入単価
                .当月入庫数 = ZYT.入荷数
                .当月入庫金額 = .当月入庫数 * .仕入単価
                .当月出庫数 = ZYT.実績金額0
                .当月出庫金額 = .当月出庫数 * .仕入単価
                .調整数 = 0
                .調整金額 = 0
                .現在在庫数 = .前月末在庫数 + .当月入庫数 - .当月出庫数
                .現在在庫金額 = .現在在庫数 * .仕入単価
                .最終入庫日 = ZYT.入荷日
                Call ZKWInsert(1)
                Call ZYTReadNext(0)
            End With
        Loop
    Next

ZaikoKeisan2_ED:
    On Error GoTo 0
    ZaikoKeisan2 = True
    Exit Function
ZaikoKeisan2_Err:
    On Error GoTo 0
    Exit Function
End Function

'+--------------------------------------+
'+  ストアドを実行して仕入単価を求める  +
'+--------------------------------------+
Private Function GetShiireTanka(ZNo As Long, TCd As String) As Boolean
    GetShiireTanka = False
    With CisDB
        .SQL = "材料購入単価決定"
        .StoadoCount = 9                '10
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用      ' 材料管理番号
        .ParaNo = 2: .ParaIO = Input用      ' 取引先
        .ParaNo = 3: .ParaIO = Input用      ' 契約先
        .ParaNo = 4: .ParaIO = Input用      ' 単価決定日付
        .ParaNo = 5: .ParaIO = OutPut用     ' 決定単価
        .ParaNo = 6: .ParaIO = OutPut用     ' 仮区分
        .ParaNo = 7: .ParaIO = OutPut用
        .ParaNo = 8: .ParaIO = OutPut用     ' 適用開始年月日(決定単価)
        .ParaNo = 9: .ParaIO = OutPut用     ' 履歴NO
''''''''.ParaNo = 10: .ParaIO = Return用
    '
        .ParaNo = 1: .ParaValue = ZNo
        .ParaNo = 2: .ParaValue = TCd
        .ParaNo = 3: .ParaValue = TCd
        .ParaNo = 4: .ParaValue = iH1_YM & "01"
    '
        If Not .DBStored Then
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(2) = "    材料購入単価の獲得に失敗しました            "
            CisFun.MB_Title = "単価決定処理"
            CisFun.MB_Button = Error
            CisFun.MBOX
            Exit Function
        Else
            .ParaNo = 5: mTanka = .ParaValue    ' 単価を取得する
        End If
    End With
    GetShiireTanka = True
End Function
'-----------------------------------------------------------------------------------------------------------' 2010.01.07 add end





