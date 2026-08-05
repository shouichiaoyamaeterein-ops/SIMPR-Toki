VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BA5D7819-2B67-11D4-806A-00000EA57F9E}#3.0#0"; "CISYMD~1.OCX"
Begin VB.Form COK0530 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "仕入明細表(集計)　発行"
   ClientHeight    =   9420
   ClientLeft      =   3405
   ClientTop       =   3615
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   6660
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
      Top             =   7290
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
      Height          =   4305
      Left            =   4260
      ScaleHeight     =   4418.27
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6600
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2565
      Width           =   6660
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   4140
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   7303
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   3120
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   1800
            TabIndex        =   5
            Top             =   2625
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
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   405
            Left            =   1815
            Top             =   1515
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   714
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
            Begin VB.OptionButton H1Op_OutPut 
               BackColor       =   &H00C0C0C0&
               Caption         =   "契約先"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   1350
               TabIndex        =   3
               Top             =   90
               Width           =   1095
            End
            Begin VB.OptionButton H1Op_OutPut 
               BackColor       =   &H00C0FFC0&
               Caption         =   "仕入先"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   75
               TabIndex        =   2
               Top             =   90
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   360
            Left            =   3765
            Top             =   930
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   635
            ForeColor       =   16711680
            Caption         =   "～"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   15.75
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
         Begin CisYMD_V60.CisYMD iH1_SYMD 
            Height          =   390
            Left            =   1815
            TabIndex        =   0
            Top             =   930
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
            Object.Height          =   390
            Object.Width           =   1875
            cSize           =   -1  'True
            cChkResult      =   0   'False
         End
         Begin VB.ComboBox H1Comb_Jisya 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1815
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   4
            Top             =   2100
            Width           =   1395
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "　出　力　先"
            Height          =   600
            Left            =   75
            TabIndex        =   15
            Top             =   3060
            Width           =   6345
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
               TabIndex        =   11
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
                  Caption         =   "仕入明細表(集計)　発行を行います"
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
            Top             =   3690
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
            Top             =   3720
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
            Left            =   405
            Top             =   900
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   635
            ForeColor       =   16711680
            Caption         =   "対象日付"
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
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   315
            Left            =   510
            Top             =   2130
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "工　　区"
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
         Begin CisYMD_V60.CisYMD iH1_EYMD 
            Height          =   375
            Left            =   4350
            TabIndex        =   1
            Top             =   930
            Width           =   1920
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
            Object.Width           =   1920
            cSize           =   -1  'True
            cChkResult      =   0   'False
         End
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   360
            Index           =   0
            Left            =   405
            Top             =   1500
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   635
            ForeColor       =   16711680
            Caption         =   "出力単位"
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
            Top             =   2640
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   529
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
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D PB_Torcd 
            Height          =   360
            Left            =   2775
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
      Caption         =   "【 仕入明細表(集計)　発行 】"
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
      TabIndex        =   10
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
   Begin Cis3D_v60.CIS3D PB_Ent 
      Height          =   585
      Left            =   8280
      Top             =   7290
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
      Top             =   7290
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6900
      TabIndex        =   17
      Top             =   7965
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8280
      TabIndex        =   14
      Top             =   7965
      Width           =   1485
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5610
      TabIndex        =   13
      Top             =   7965
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
      TabIndex        =   12
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "COK0530"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  仕入明細表 発行
'**       フォームID    :  COK0530
'**       処理概要      :
'**
'**       作  成  日    :  2009/02/19
'**       変  更  日    :  2009/03/14 材料直送支給の変更
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

    Dim CisComb1        As Cis_Comb.CisCombo
'   背景色
    Dim mSelColor       As Long
    Dim mUnSelColor     As Long
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

    If gLogoDisp Then
       BG_Logo(0).Caption = gLogo
       CisFrm.BGLDisp BG_Logo, 10
    End If

    '+------------------------+
    '+     ｺﾝﾎﾞﾎﾞｯｸｽのｾｯﾄ     +
    '+------------------------+
    Set CisComb1 = New CisCombo
    With CisComb1
        Set .Connect = CisDB
        Set .CombObj = H1Comb_Jisya
        .Control = "名称マスタ"
        .ComboName = "自社工区"
        .Ess = False
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With

    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

    Call CisComb1.ComboSet

    H1Op_OutPut(0).Value = True
    
'   CSVﾌｧｲﾙ出力先
    mCsvName = Mid(CisFun.WinRoot, 1, 3) & "仕入明細表.csv"
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
End Sub
'+---------------------+
'+    取引先検索        +
'+---------------------+
Private Sub PB_Torcd_Click()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 8
'
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    iH1_Torcd.SetFocus

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
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
'+----------------------------------------+
'+  H1Op_Out
'+----------------------------------------+
Private Sub H1Op_OutPut_GotFocus(Index As Integer)
    H1Op_OutPut(0).Tag = Index
    For Each gObj In H1Op_OutPut
        If gObj.Index = Index Then
            gObj.BackColor = mSelColor
        Else
            gObj.BackColor = mUnSelColor
        End If
    Next gObj
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
        If TorNmGet(iH1_Torcd, 8) Then
            H1lb_Tornm = TRM_RName
        End If
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

'   開始日
    If Trim(iH1_SYMD) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    開始年月日を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SYMD.SetFocus
        Exit Function
    End If
    If Not iH1_SYMD.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    開始年月日を正しく入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SYMD.SetFocus
        Exit Function
    End If
'   終了日
    If Trim(iH1_EYMD) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    終了年月日を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_EYMD.SetFocus
        Exit Function
    End If
    If Not iH1_EYMD.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    終了年月日を正しく入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_EYMD.SetFocus
        Exit Function
    End If

'   契約先
    If Trim(iH1_Torcd) <> "" Then
        If Not TorNmGet(iH1_Torcd, 8) Then
        With CisFun
            .MB_Lines = 2
            .MB_MSG(2) = "    契約先 未登録         "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
    End If
    
    lb_msg.Caption = "【 仕入明細データ 抽出中 】"
    DoEvents

    RV_SYmd = RTrim(iH1_SYMD)
    RV_EYmd = RTrim(iH1_EYMD)

    COK0530_Kojyo = RTrim(H1Comb_Jisya.Text)
    If RTrim(COK0530_Kojyo) = "" Then
       COK0530_Kojyo = "合算"
    End If
    
    COK0530_KTornm = RTrim(H1lb_Tornm)
    
    If H1Op_OutPut(0) Then COK0530_OutPut = 0 Else COK0530_OutPut = 1
    
    With CisDB
        '+---------- ワークの削除 -----------+
        .SQL = ""
        .SQL = .SQL & " DELETE FROM 仕入明細書ワーク"
        .SQL = .SQL & "  WHERE 処理端末 = SUBSTRING(HOST_NAME(), 1, 20)"
        Call CisDB.DBExec
    End With
    
    gSL_Select = "INSERT INTO 仕入明細書ワーク " & vbCrLf
    '********* 入荷テーブル：外注先費用
    gSL_Select = gSL_Select & " SELECT  SUBSTRING(HOST_NAME(),1,20) 処理端末,1 区分,'外注先費用' 区分名称,"
    gSL_Select = gSL_Select & "         ISNULL(HTM.契約先, '') 契約先, ISNULL(HTM.契約先,'') + ':' + ISNULL(TR.略称, '') 契約先名,"
    gSL_Select = gSL_Select & "         NY.品番,0 材料管理番号, NY.仕入先, ISNULL(NY.仕入先,'') + ':' + ISNULL(TR1.略称, '') 仕入先名 ,"
    gSL_Select = gSL_Select & "         SUM(入荷数) 入荷数,"
    gSL_Select = gSL_Select & "         CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.単価, 0)"
    gSL_Select = gSL_Select & "              WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.単価, 0)"
    gSL_Select = gSL_Select & "              WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.単価, 0)"
    gSL_Select = gSL_Select & "              WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.単価, 0)"
    gSL_Select = gSL_Select & "              ELSE 0"
    gSL_Select = gSL_Select & "         END 単価,"
    gSL_Select = gSL_Select & "         CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.有償支給単価, 0)"
    gSL_Select = gSL_Select & "              WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.有償支給単価, 0)"
    gSL_Select = gSL_Select & "              WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.有償支給単価, 0)"
    gSL_Select = gSL_Select & "              WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.有償支給単価, 0)"
    gSL_Select = gSL_Select & "              ELSE 0"
    gSL_Select = gSL_Select & "         END 有償支給,"
    gSL_Select = gSL_Select & "         CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.外注材料費, 0)"
    gSL_Select = gSL_Select & "              WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.外注材料費, 0)"
    gSL_Select = gSL_Select & "              WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.外注材料費, 0)"
    gSL_Select = gSL_Select & "              WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.外注材料費, 0)"
    gSL_Select = gSL_Select & "              ELSE 0"
    gSL_Select = gSL_Select & "         END 外注材料,"
    gSL_Select = gSL_Select & "         CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.外注購入費, 0)"
    gSL_Select = gSL_Select & "              WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.外注購入費, 0)"
    gSL_Select = gSL_Select & "              WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.外注購入費, 0)"
    gSL_Select = gSL_Select & "              WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.外注購入費, 0)"
    gSL_Select = gSL_Select & "              ELSE 0"
    gSL_Select = gSL_Select & "         END 外注購入,"
    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT.有償支給単価, 0)"
    gSL_Select = gSL_Select & "                  WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT2.有償支給単価, 0)"
    gSL_Select = gSL_Select & "                  WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT3.有償支給単価, 0)"
    gSL_Select = gSL_Select & "                  WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT4.有償支給単価, 0)"
    gSL_Select = gSL_Select & "                  ELSE 0"
    gSL_Select = gSL_Select & "              END  有償支給計,"
    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT.外注材料費, 0)"
    gSL_Select = gSL_Select & "                  WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT2.外注材料費, 0)"
    gSL_Select = gSL_Select & "                  WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT3.外注材料費, 0)"
    gSL_Select = gSL_Select & "                  WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT4.外注材料費, 0)"
    gSL_Select = gSL_Select & "                  ELSE 0"
    gSL_Select = gSL_Select & "              END 外注材料計,"
    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT.外注購入費, 0)"
    gSL_Select = gSL_Select & "                  WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT2.外注購入費, 0)"
    gSL_Select = gSL_Select & "                  WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT3.外注購入費, 0)"
    gSL_Select = gSL_Select & "                  WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT4.外注購入費, 0)"
    gSL_Select = gSL_Select & "                  ELSE 0"
    gSL_Select = gSL_Select & "              END 外注購入計,"
    gSL_Select = gSL_Select & "             0 実績換算値,"
    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT.単価, 0)"
    gSL_Select = gSL_Select & "                  WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT2.単価, 0)"
    gSL_Select = gSL_Select & "                  WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT3.単価, 0)"
    gSL_Select = gSL_Select & "                  WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                      Sum (入荷数) * IsNull(KT4.単価, 0)"
    gSL_Select = gSL_Select & "                  ELSE 0"
    gSL_Select = gSL_Select & "              END 金額,"
    gSL_Select = gSL_Select & " "
    gSL_Select = gSL_Select & "         ISNULL(SY.値, '') 自社工区, NY.受入"
    gSL_Select = gSL_Select & "    FROM 入荷テーブル NY"
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 発注テーブル HT"
    gSL_Select = gSL_Select & "              ON HT.発注管理NO = NY.発注管理NO"
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "              ON HN.品番 = HT.品番"
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 品番取引先マスタ HTM"
    gSL_Select = gSL_Select & "              ON HTM.品番 = HT.品番"
    gSL_Select = gSL_Select & "             AND HTM.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "             AND HTM.受入 = HT.受入"
    gSL_Select = gSL_Select & "             AND HTM.品目 = 1"
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "              ON TR.取引先CD = HTM.契約先"
    gSL_Select = gSL_Select & "             AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "              ON TR1.取引先CD = NY.仕入先"
    gSL_Select = gSL_Select & "             AND TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "              ON SY.区分名称 = '自社工区'"
    gSL_Select = gSL_Select & "             AND SY.英数字1 = NY.受入"
    gSL_Select = gSL_Select & "             AND SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 購入単価マスタ KT"
    gSL_Select = gSL_Select & "              ON KT.品番 = HT.品番"
    gSL_Select = gSL_Select & "             AND KT.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "             AND KT.契約先 = HTM.契約先"
    gSL_Select = gSL_Select & "             AND KT.仮単価区分 = ''"
    gSL_Select = gSL_Select & "             AND KT.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT1"
    gSL_Select = gSL_Select & "                               Where KT1.品番 = HT.品番"
    gSL_Select = gSL_Select & "                                 AND   KT1.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "                                 AND   KT1.仮単価区分 = ''"
    gSL_Select = gSL_Select & "                                 AND ((KT1.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                 AND   KT1.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                  OR   KT1.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                 AND ((KT1.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                 AND   KT1.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                  OR   KT1.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 購入単価マスタ KT2"
    gSL_Select = gSL_Select & "              ON KT2.品番 = HT.品番"
    gSL_Select = gSL_Select & "             AND KT2.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "             AND KT2.契約先 = HTM.契約先"
    gSL_Select = gSL_Select & "             AND KT2.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "             AND KT2.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT12"
    gSL_Select = gSL_Select & "                               Where KT12.品番 = HT.品番"
    gSL_Select = gSL_Select & "                                 AND   KT12.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "                                 AND   KT12.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "                                 AND ((KT12.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                 AND   KT12.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                  OR   KT12.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                 AND ((KT12.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                 AND   KT12.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                  OR   KT12.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 購入単価マスタ KT3"
    gSL_Select = gSL_Select & "              ON KT3.品番 = HT.品番"
    gSL_Select = gSL_Select & "             AND KT3.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "             AND KT3.契約先 = ''"
    gSL_Select = gSL_Select & "             AND KT3.仮単価区分 = ''"
    gSL_Select = gSL_Select & "             AND KT3.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT13"
    gSL_Select = gSL_Select & "                               Where KT13.品番 = HT.品番"
    gSL_Select = gSL_Select & "                                 AND   KT13.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "                                 AND   KT13.仮単価区分 = ''"
    gSL_Select = gSL_Select & "                                 AND ((KT13.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                 AND   KT13.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                  OR   KT13.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                 AND ((KT13.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                 AND   KT13.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                  OR   KT13.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "         LEFT OUTER JOIN 購入単価マスタ KT4"
    gSL_Select = gSL_Select & "              ON KT4.品番 = HT.品番"
    gSL_Select = gSL_Select & "             AND KT4.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "             AND KT4.契約先 = ''"
    gSL_Select = gSL_Select & "             AND KT4.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "             AND KT4.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT14"
    gSL_Select = gSL_Select & "                               Where KT14.品番 = HT.品番"
    gSL_Select = gSL_Select & "                                 AND   KT14.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "                                 AND   KT14.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "                                 AND ((KT14.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                 AND   KT14.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                  OR   KT14.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                 AND ((KT14.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                 AND   KT14.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                  OR   KT14.適用終了日付 = ''))"
'    gSL_Select = gSL_Select & "         Where (HTM.手配区分 = 0 Or HTM.手配区分 = 1) " '** 購入 Or 直送のみ
    gSL_Select = gSL_Select & "         Where (HTM.手配区分 = 0 Or ( HTM.手配区分 = 1 AND HTM.取引先 = HTM.契約先 ) ) " '** 購入 Or 直送(仕入先・契約先同一)のみ  2009/03/14
    gSL_Select = gSL_Select & "         AND   ( HN.部品区分 = 0 OR HN.部品区分 = 1 ) " '** 部品 or 外注加工
    gSL_Select = gSL_Select & "         AND   NY.入荷日 >= '" & RTrim(iH1_SYMD) & "'"
    gSL_Select = gSL_Select & "         AND   NY.入荷日 <= '" & RTrim(iH1_EYMD) & "'"
    If RTrim(iH1_Torcd) <> "" Then
       gSL_Select = gSL_Select & "         AND   HTM.契約先 = '" & RTrim(iH1_Torcd) & "'"
    End If
    
    ' 自社工区
    Call CisComb1.CodeGet
    If RTrim(CisComb1.Code) <> "" Then
       gSL_Select = gSL_Select & "         AND   SY.値 = " & CisFun.Val2(CisComb1.Code)
    End If
    
    gSL_Select = gSL_Select & "  GROUP BY TR.略称, HTM.契約先,KT.単価, KT2.単価, KT3.単価, KT4.単価,"
    gSL_Select = gSL_Select & "           KT.有償支給単価 ,KT2.有償支給単価 ,KT3.有償支給単価 ,KT4.有償支給単価 ,"
    gSL_Select = gSL_Select & "           KT.外注材料費 ,KT2.外注材料費 ,KT3.外注材料費 ,KT4.外注材料費 ,"
    gSL_Select = gSL_Select & "           KT.外注購入費 ,KT2.外注購入費 ,KT3.外注購入費 ,KT4.外注購入費 ,"
    gSL_Select = gSL_Select & "           ISNULL(SY.値, ''), NY.受入,       NY.品番, NY.仕入先,TR1.略称"

    
    gSL_Select = gSL_Select & " UNION ALL "
    
    
    '*****入荷テーブル：補助材
    gSL_Select = gSL_Select & "SELECT  SUBSTRING(HOST_NAME(),1,20) 処理端末,2 区分,'補助材' 区分名称,"
    gSL_Select = gSL_Select & "        ISNULL(HTM.契約先, '') 契約先, ISNULL(HTM.契約先,'') + ':' + ISNULL(TR.略称, '') 契約先名,"
    gSL_Select = gSL_Select & "        NY.品番,0 材料管理番号, NY.仕入先, ISNULL(NY.仕入先,'') + ':' + ISNULL(TR1.略称, '') 仕入先名 ,"
    gSL_Select = gSL_Select & "        SUM(入荷数) 入荷数,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.単価, 0)"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "        END 単価,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.有償支給単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.有償支給単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.有償支給単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.有償支給単価, 0)"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "        END 有償支給,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.外注材料費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.外注材料費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.外注材料費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.外注材料費, 0)"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "        END 外注材料,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.外注購入費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.外注購入費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.外注購入費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.外注購入費, 0)"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "        END 外注購入,"
    gSL_Select = gSL_Select & "            CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT.有償支給単価, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT2.有償支給単価, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT3.有償支給単価, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT4.有償支給単価, 0)"
    gSL_Select = gSL_Select & "                 ELSE 0"
    gSL_Select = gSL_Select & "             END  有償支給計,"
    gSL_Select = gSL_Select & "            CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT.外注材料費, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT2.外注材料費, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT3.外注材料費, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT4.外注材料費, 0)"
    gSL_Select = gSL_Select & "                 ELSE 0"
    gSL_Select = gSL_Select & "             END 外注材料計,"
    gSL_Select = gSL_Select & "            CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT.外注購入費, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT2.外注購入費, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT3.外注購入費, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT4.外注購入費, 0)"
    gSL_Select = gSL_Select & "                 ELSE 0"
    gSL_Select = gSL_Select & "             END 外注購入計,"
    gSL_Select = gSL_Select & "            0 実績換算値,"
    gSL_Select = gSL_Select & "            CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT.単価, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT2.単価, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT3.単価, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT4.単価, 0)"
    gSL_Select = gSL_Select & "                 ELSE 0"
    gSL_Select = gSL_Select & "             END 金額,"

    gSL_Select = gSL_Select & "        ISNULL(SY.値, '') 自社工区, NY.受入"
    gSL_Select = gSL_Select & "   FROM 入荷テーブル NY"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 発注テーブル HT"
    gSL_Select = gSL_Select & "             ON HT.発注管理NO = NY.発注管理NO"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "             ON HN.品番 = HT.品番"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番取引先マスタ HTM"
    gSL_Select = gSL_Select & "             ON HTM.品番 = HT.品番"
    gSL_Select = gSL_Select & "            AND HTM.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "            AND HTM.受入 = HT.受入"
    gSL_Select = gSL_Select & "            AND HTM.品目 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "             ON TR.取引先CD = HTM.契約先"
    gSL_Select = gSL_Select & "            AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "             ON TR1.取引先CD = NY.仕入先"
    gSL_Select = gSL_Select & "            AND TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "             ON SY.区分名称 = '自社工区'"
    gSL_Select = gSL_Select & "            AND SY.英数字1 = NY.受入"
    gSL_Select = gSL_Select & "            AND SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT"
    gSL_Select = gSL_Select & "             ON KT.品番 = HT.品番"
    gSL_Select = gSL_Select & "            AND KT.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "            AND KT.契約先 = HTM.契約先"
    gSL_Select = gSL_Select & "            AND KT.仮単価区分 = ''"
    gSL_Select = gSL_Select & "            AND KT.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT1"
    gSL_Select = gSL_Select & "                              Where KT1.品番 = HT.品番"
    gSL_Select = gSL_Select & "                                AND   KT1.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT1.仮単価区分 = ''"
    gSL_Select = gSL_Select & "                                AND ((KT1.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT1.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT1.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT1.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT1.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT1.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT2"
    gSL_Select = gSL_Select & "             ON KT2.品番 = HT.品番"
    gSL_Select = gSL_Select & "            AND KT2.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "            AND KT2.契約先 = HTM.契約先"
    gSL_Select = gSL_Select & "            AND KT2.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "            AND KT2.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT12"
    gSL_Select = gSL_Select & "                              Where KT12.品番 = HT.品番"
    gSL_Select = gSL_Select & "                                AND   KT12.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT12.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "                                AND ((KT12.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT12.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT12.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT12.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT12.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT12.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT3"
    gSL_Select = gSL_Select & "             ON KT3.品番 = HT.品番"
    gSL_Select = gSL_Select & "            AND KT3.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "            AND KT3.契約先 = ''"
    gSL_Select = gSL_Select & "            AND KT3.仮単価区分 = ''"
    gSL_Select = gSL_Select & "            AND KT3.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT13"
    gSL_Select = gSL_Select & "                              Where KT13.品番 = HT.品番"
    gSL_Select = gSL_Select & "                                AND   KT13.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT13.仮単価区分 = ''"
    gSL_Select = gSL_Select & "                                AND ((KT13.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT13.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT13.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT13.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT13.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT13.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT4"
    gSL_Select = gSL_Select & "             ON KT4.品番 = HT.品番"
    gSL_Select = gSL_Select & "            AND KT4.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "            AND KT4.契約先 = ''"
    gSL_Select = gSL_Select & "            AND KT4.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "            AND KT4.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT14"
    gSL_Select = gSL_Select & "                              Where KT14.品番 = HT.品番"
    gSL_Select = gSL_Select & "                                AND   KT14.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT14.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "                                AND ((KT14.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT14.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT14.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT14.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT14.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT14.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        Where (HTM.手配区分 = 0 Or HTM.手配区分 = 1) "  '** 購入 Or 直送のみ
    gSL_Select = gSL_Select & "        AND   HN.部品区分 = 2 "                         '** 補助材"
    gSL_Select = gSL_Select & "        AND   NY.入荷日 >= '" & RTrim(iH1_SYMD) & "'"
    gSL_Select = gSL_Select & "        AND   NY.入荷日 <= '" & RTrim(iH1_EYMD) & "'"
    If RTrim(iH1_Torcd) <> "" Then
       gSL_Select = gSL_Select & "         AND   HTM.契約先 = '" & RTrim(iH1_Torcd) & "'"
    End If
    
    ' 自社工区
    Call CisComb1.CodeGet
    If RTrim(CisComb1.Code) <> "" Then
       gSL_Select = gSL_Select & "     AND   SY.値 = " & CisFun.Val2(CisComb1.Code)
    End If
    
    gSL_Select = gSL_Select & " GROUP BY TR.略称, HTM.契約先,KT.単価, KT2.単価, KT3.単価, KT4.単価,"
    gSL_Select = gSL_Select & "          KT.有償支給単価 ,KT2.有償支給単価 ,KT3.有償支給単価 ,KT4.有償支給単価 ,"
    gSL_Select = gSL_Select & "          KT.外注材料費 ,KT2.外注材料費 ,KT3.外注材料費 ,KT4.外注材料費 ,"
    gSL_Select = gSL_Select & "          KT.外注購入費 ,KT2.外注購入費 ,KT3.外注購入費 ,KT4.外注購入費 ,"
    gSL_Select = gSL_Select & "          ISNULL(SY.値, ''), NY.受入,       NY.品番, NY.仕入先,TR1.略称"

    
    gSL_Select = gSL_Select & " UNION ALL "
    
    
    '***** 入荷テーブル：有償支給
    gSL_Select = gSL_Select & "SELECT  SUBSTRING(HOST_NAME(),1,20) 処理端末,3 区分,'有償支給' 区分名称,"
    gSL_Select = gSL_Select & "        ISNULL(HTM.契約先, '') 契約先, ISNULL(HTM.契約先,'') + ':' + ISNULL(TR.略称, '') 契約先名,"
    gSL_Select = gSL_Select & "        NY.品番,0 材料管理番号, NY.仕入先, ISNULL(NY.仕入先,'') + ':' + ISNULL(TR1.略称, '') 仕入先名 ,"
    gSL_Select = gSL_Select & "        SUM(入荷数) 入荷数,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.単価, 0)"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "        END 単価,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.有償支給単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.有償支給単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.有償支給単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.有償支給単価, 0)"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "        END 有償支給,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.外注材料費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.外注材料費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.外注材料費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.外注材料費, 0)"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "        END 外注材料,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.外注購入費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.外注購入費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.外注購入費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.外注購入費, 0)"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "        END 外注購入,"
'    gSL_Select = gSL_Select & "            CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT.有償支給単価, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT2.有償支給単価, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT3.有償支給単価, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT4.有償支給単価, 0)"
'    gSL_Select = gSL_Select & "                 ELSE 0"
'    gSL_Select = gSL_Select & "             END  有償支給計,"
'    gSL_Select = gSL_Select & "            CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT.外注材料費, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT2.外注材料費, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT3.外注材料費, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT4.外注材料費, 0)"
'    gSL_Select = gSL_Select & "                 ELSE 0"
'    gSL_Select = gSL_Select & "             END 外注材料計,"
'    gSL_Select = gSL_Select & "            CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT.外注購入費, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT2.外注購入費, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT3.外注購入費, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT4.外注購入費, 0)"
'    gSL_Select = gSL_Select & "                 ELSE 0"
'    gSL_Select = gSL_Select & "             END 外注購入計,"
    gSL_Select = gSL_Select & "            0 有償支給計,"
    gSL_Select = gSL_Select & "            0 外注材料計,"
    gSL_Select = gSL_Select & "            0 外注購入計,"
    gSL_Select = gSL_Select & "            0 実績換算値,"
    gSL_Select = gSL_Select & "            CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT.単価, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT2.単価, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT3.単価, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * IsNull(KT4.単価, 0)"
    gSL_Select = gSL_Select & "                 ELSE 0"
    gSL_Select = gSL_Select & "             END 金額,"

    gSL_Select = gSL_Select & "        ISNULL(SY.値, '') 自社工区, NY.受入"
    gSL_Select = gSL_Select & "   FROM 入荷テーブル NY"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 発注テーブル HT"
    gSL_Select = gSL_Select & "             ON HT.発注管理NO = NY.発注管理NO"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "             ON HN.品番 = HT.品番"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番取引先マスタ HTM"
    gSL_Select = gSL_Select & "             ON HTM.品番 = HT.品番"
    gSL_Select = gSL_Select & "            AND HTM.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "            AND HTM.受入 = HT.受入"
    gSL_Select = gSL_Select & "            AND HTM.品目 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "             ON TR.取引先CD = HTM.契約先"
    gSL_Select = gSL_Select & "            AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "             ON TR1.取引先CD = NY.仕入先"
    gSL_Select = gSL_Select & "            AND TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "             ON SY.区分名称 = '自社工区'"
    gSL_Select = gSL_Select & "            AND SY.英数字1 = NY.受入"
    gSL_Select = gSL_Select & "            AND SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT"
    gSL_Select = gSL_Select & "             ON KT.品番 = HT.品番"
    gSL_Select = gSL_Select & "            AND KT.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "            AND KT.契約先 = HTM.契約先"
    gSL_Select = gSL_Select & "            AND KT.仮単価区分 = ''"
    gSL_Select = gSL_Select & "            AND KT.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT1"
    gSL_Select = gSL_Select & "                              Where KT1.品番 = HT.品番"
    gSL_Select = gSL_Select & "                                AND   KT1.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT1.仮単価区分 = ''"
    gSL_Select = gSL_Select & "                                AND ((KT1.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT1.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT1.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT1.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT1.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT1.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT2"
    gSL_Select = gSL_Select & "             ON KT2.品番 = HT.品番"
    gSL_Select = gSL_Select & "            AND KT2.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "            AND KT2.契約先 = HTM.契約先"
    gSL_Select = gSL_Select & "            AND KT2.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "            AND KT2.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT12"
    gSL_Select = gSL_Select & "                              Where KT12.品番 = HT.品番"
    gSL_Select = gSL_Select & "                                AND   KT12.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT12.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "                                AND ((KT12.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT12.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT12.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT12.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT12.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT12.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT3"
    gSL_Select = gSL_Select & "             ON KT3.品番 = HT.品番"
    gSL_Select = gSL_Select & "            AND KT3.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "            AND KT3.契約先 = ''"
    gSL_Select = gSL_Select & "            AND KT3.仮単価区分 = ''"
    gSL_Select = gSL_Select & "            AND KT3.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT13"
    gSL_Select = gSL_Select & "                              Where KT13.品番 = HT.品番"
    gSL_Select = gSL_Select & "                                AND   KT13.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT13.仮単価区分 = ''"
    gSL_Select = gSL_Select & "                                AND ((KT13.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT13.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT13.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT13.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT13.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT13.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT4"
    gSL_Select = gSL_Select & "             ON KT4.品番 = HT.品番"
    gSL_Select = gSL_Select & "            AND KT4.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "            AND KT4.契約先 = ''"
    gSL_Select = gSL_Select & "            AND KT4.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "            AND KT4.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT14"
    gSL_Select = gSL_Select & "                              Where KT14.品番 = HT.品番"
    gSL_Select = gSL_Select & "                                AND   KT14.取引先 = HT.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT14.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "                                AND ((KT14.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT14.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT14.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT14.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT14.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT14.適用終了日付 = ''))"
'    gSL_Select = gSL_Select & "        Where HTM.手配区分 = 3"                              '** 受給のみ
    gSL_Select = gSL_Select & "        Where (HTM.手配区分 = 3 Or ( HTM.手配区分 = 1 AND HTM.取引先 <> HTM.契約先 ) )" '** 受給 Or 直送(仕入先・契約先不一致)のみ  2009/03/14
    gSL_Select = gSL_Select & "        AND   NY.入荷日 >= '" & RTrim(iH1_SYMD) & "'"
    gSL_Select = gSL_Select & "        AND   NY.入荷日 <= '" & RTrim(iH1_EYMD) & "'"
    If RTrim(iH1_Torcd) <> "" Then
       gSL_Select = gSL_Select & "         AND   HTM.契約先 = '" & RTrim(iH1_Torcd) & "'"
    End If
    
    ' 自社工区
    Call CisComb1.CodeGet
    If RTrim(CisComb1.Code) <> "" Then
       gSL_Select = gSL_Select & "     AND   SY.値 = " & CisFun.Val2(CisComb1.Code)
    End If
    
    gSL_Select = gSL_Select & " GROUP BY TR.略称, HTM.契約先,KT.単価, KT2.単価, KT3.単価, KT4.単価,"
    gSL_Select = gSL_Select & "          KT.有償支給単価 ,KT2.有償支給単価 ,KT3.有償支給単価 ,KT4.有償支給単価 ,"
    gSL_Select = gSL_Select & "          KT.外注材料費 ,KT2.外注材料費 ,KT3.外注材料費 ,KT4.外注材料費 ,"
    gSL_Select = gSL_Select & "          KT.外注購入費 ,KT2.外注購入費 ,KT3.外注購入費 ,KT4.外注購入費 ,"
    gSL_Select = gSL_Select & "          ISNULL(SY.値, ''), NY.受入,       NY.品番, NY.仕入先,TR1.略称"

    
    gSL_Select = gSL_Select & " UNION ALL "
    
    
    '***** 受給検収テーブル：有償支給
    gSL_Select = gSL_Select & "SELECT  SUBSTRING(HOST_NAME(),1,20) 処理端末,3 区分,'有償支給' 区分名称,"
    gSL_Select = gSL_Select & "        ISNULL(HTM.契約先, '') 契約先, ISNULL(HTM.契約先,'') + ':' + ISNULL(TR.略称, '') 契約先名,"
    gSL_Select = gSL_Select & "        NY.品番,0 材料管理番号, NY.仕入先,ISNULL(NY.仕入先,'') + ':' + ISNULL(TR1.略称, '') 仕入先名 ,"
    gSL_Select = gSL_Select & "        SUM(検収数) 入荷数,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.単価, 0)"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "        END 単価,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.有償支給単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.有償支給単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.有償支給単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.有償支給単価, 0)"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "        END 有償支給,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.外注材料費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.外注材料費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.外注材料費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.外注材料費, 0)"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "        END 外注材料,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.外注購入費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.外注購入費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.外注購入費, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.外注購入費, 0)"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "        END 外注購入,"
'    gSL_Select = gSL_Select & "            CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT.有償支給単価, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT2.有償支給単価, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT3.有償支給単価, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT4.有償支給単価, 0)"
'    gSL_Select = gSL_Select & "                 ELSE 0"
'    gSL_Select = gSL_Select & "             END  有償支給計,"
'    gSL_Select = gSL_Select & "            CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT.外注材料費, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT2.外注材料費, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT3.外注材料費, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT4.外注材料費, 0)"
'    gSL_Select = gSL_Select & "                 ELSE 0"
'    gSL_Select = gSL_Select & "             END 外注材料計,"
'    gSL_Select = gSL_Select & "            CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT.外注購入費, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT2.外注購入費, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT3.外注購入費, 0)"
'    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
'    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT4.外注購入費, 0)"
'    gSL_Select = gSL_Select & "                 ELSE 0"
'    gSL_Select = gSL_Select & "             END 外注購入計,"
    gSL_Select = gSL_Select & "            0 有償支給計,"
    gSL_Select = gSL_Select & "            0 外注材料計,"
    gSL_Select = gSL_Select & "            0 外注購入計,"
    gSL_Select = gSL_Select & "            0 実績換算値,"
    gSL_Select = gSL_Select & "            CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT.単価, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT2.単価, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT3.単価, 0)"
    gSL_Select = gSL_Select & "                 WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (検収数) * IsNull(KT4.単価, 0)"
    gSL_Select = gSL_Select & "                 ELSE 0"
    gSL_Select = gSL_Select & "             END 金額,"

    gSL_Select = gSL_Select & "        ISNULL(SY.値, '') 自社工区, NY.受入"
    gSL_Select = gSL_Select & "   FROM 受給検収テーブル NY"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "             ON HN.品番 = NY.品番"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番取引先マスタ HTM"
    gSL_Select = gSL_Select & "             ON HTM.品番 = NY.品番"
    gSL_Select = gSL_Select & "            AND HTM.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND HTM.受入 = NY.受入"
    gSL_Select = gSL_Select & "            AND HTM.品目 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "             ON TR.取引先CD = HTM.契約先"
    gSL_Select = gSL_Select & "            AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "             ON TR1.取引先CD = NY.仕入先"
    gSL_Select = gSL_Select & "            AND TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "             ON SY.区分名称 = '自社工区'"
    gSL_Select = gSL_Select & "            AND SY.英数字1 = NY.受入"
    gSL_Select = gSL_Select & "            AND SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT"
    gSL_Select = gSL_Select & "             ON KT.品番 = NY.品番"
    gSL_Select = gSL_Select & "            AND KT.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND KT.契約先 = HTM.契約先"
    gSL_Select = gSL_Select & "            AND KT.仮単価区分 = ''"
    gSL_Select = gSL_Select & "            AND KT.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT1"
    gSL_Select = gSL_Select & "                              Where KT1.品番 = NY.品番"
    gSL_Select = gSL_Select & "                                AND   KT1.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT1.仮単価区分 = ''"
    gSL_Select = gSL_Select & "                                AND ((KT1.適用開始日付 < NY.検収日"
    gSL_Select = gSL_Select & "                                AND   KT1.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT1.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT1.適用終了日付 > NY.検収日"
    gSL_Select = gSL_Select & "                                AND   KT1.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT1.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT2"
    gSL_Select = gSL_Select & "             ON KT2.品番 = NY.品番"
    gSL_Select = gSL_Select & "            AND KT2.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND KT2.契約先 = HTM.契約先"
    gSL_Select = gSL_Select & "            AND KT2.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "            AND KT2.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT12"
    gSL_Select = gSL_Select & "                              Where KT12.品番 = NY.品番"
    gSL_Select = gSL_Select & "                                AND   KT12.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT12.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "                                AND ((KT12.適用開始日付 < NY.検収日"
    gSL_Select = gSL_Select & "                                AND   KT12.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT12.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT12.適用終了日付 > NY.検収日"
    gSL_Select = gSL_Select & "                                AND   KT12.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT12.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT3"
    gSL_Select = gSL_Select & "             ON KT3.品番 = NY.品番"
    gSL_Select = gSL_Select & "            AND KT3.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND KT3.契約先 = ''"
    gSL_Select = gSL_Select & "            AND KT3.仮単価区分 = ''"
    gSL_Select = gSL_Select & "            AND KT3.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT13"
    gSL_Select = gSL_Select & "                              Where KT13.品番 = NY.品番"
    gSL_Select = gSL_Select & "                                AND   KT13.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT13.仮単価区分 = ''"
    gSL_Select = gSL_Select & "                                AND ((KT13.適用開始日付 < NY.検収日"
    gSL_Select = gSL_Select & "                                AND   KT13.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT13.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT13.適用終了日付 > NY.検収日"
    gSL_Select = gSL_Select & "                                AND   KT13.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT13.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT4"
    gSL_Select = gSL_Select & "             ON KT4.品番 = NY.品番"
    gSL_Select = gSL_Select & "            AND KT4.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND KT4.契約先 = ''"
    gSL_Select = gSL_Select & "            AND KT4.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "            AND KT4.履歴NO = (SELECT   MAX(履歴NO) FROM 購入単価マスタ KT14"
    gSL_Select = gSL_Select & "                              Where KT14.品番 = NY.品番"
    gSL_Select = gSL_Select & "                                AND   KT14.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT14.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "                                AND ((KT14.適用開始日付 < NY.検収日"
    gSL_Select = gSL_Select & "                                AND   KT14.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT14.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT14.適用終了日付 > NY.検収日"
    gSL_Select = gSL_Select & "                                AND   KT14.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT14.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        Where HTM.手配区分 = 3 "                             '** 受給のみ
    gSL_Select = gSL_Select & "        AND   NY.検収日 >= '" & RTrim(iH1_SYMD) & "'"
    gSL_Select = gSL_Select & "        AND   NY.検収日 <= '" & RTrim(iH1_EYMD) & "'"
    If RTrim(iH1_Torcd) <> "" Then
       gSL_Select = gSL_Select & "         AND   HTM.契約先 = '" & RTrim(iH1_Torcd) & "'"
    End If

    ' 自社工区
    Call CisComb1.CodeGet
    If RTrim(CisComb1.Code) <> "" Then
       gSL_Select = gSL_Select & "     AND   SY.値 = " & CisFun.Val2(CisComb1.Code)
    End If
    
    gSL_Select = gSL_Select & " GROUP BY TR.略称, HTM.契約先,KT.単価, KT2.単価, KT3.単価, KT4.単価,"
    gSL_Select = gSL_Select & "          KT.有償支給単価 ,KT2.有償支給単価 ,KT3.有償支給単価 ,KT4.有償支給単価 ,"
    gSL_Select = gSL_Select & "          KT.外注材料費 ,KT2.外注材料費 ,KT3.外注材料費 ,KT4.外注材料費 ,"
    gSL_Select = gSL_Select & "          KT.外注購入費 ,KT2.外注購入費 ,KT3.外注購入費 ,KT4.外注購入費 ,"
    gSL_Select = gSL_Select & "          ISNULL(SY.値, ''), NY.受入,       NY.品番, NY.仕入先,TR1.略称"
    
    
    gSL_Select = gSL_Select & " UNION ALL "
    
    
    '***** 材料入荷テーブル：材料（自給）
    gSL_Select = gSL_Select & "SELECT  SUBSTRING(HOST_NAME(),1,20) 処理端末,4 区分,'材料' 区分名称,"
    gSL_Select = gSL_Select & "        ISNULL(HTM.契約先, '') 契約先, ISNULL(HTM.契約先,'') + ':' + ISNULL(TR.略称, '') 契約先名,"
    gSL_Select = gSL_Select & "        '' 品番,NY.材料管理番号, NY.仕入先,ISNULL(NY.仕入先,'') + ':' + ISNULL(TR1.略称, '') 仕入先名 ,"
    gSL_Select = gSL_Select & "        SUM(入荷数) 入荷数,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.単価, 0)"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "        END 単価,"
    gSL_Select = gSL_Select & "        0 有償支給,"
    gSL_Select = gSL_Select & "        0 外注材料,"
    gSL_Select = gSL_Select & "        0 外注購入,"
    gSL_Select = gSL_Select & "        0 有償支給計,"
    gSL_Select = gSL_Select & "        0 外注材料計,"
    gSL_Select = gSL_Select & "        0 外注購入計,"
    gSL_Select = gSL_Select & "        HTM.実績換算値,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(HTM.実績換算値,0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * HTM.実績換算値 * IsNull(KT.単価, 0)"
    gSL_Select = gSL_Select & "                 ELSE SUM(入荷数) * ISNULL(KT.単価,0) END"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(HTM.実績換算値,0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * HTM.実績換算値 * IsNull(KT2.単価, 0)"
    gSL_Select = gSL_Select & "                 ELSE SUM(入荷数) * ISNULL(KT2.単価,0)  END"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(HTM.実績換算値,0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * HTM.実績換算値 * IsNull(KT3.単価, 0)"
    gSL_Select = gSL_Select & "                 ELSE SUM(入荷数) * ISNULL(KT3.単価,0) END"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(HTM.実績換算値,0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * HTM.実績換算値 * IsNull(KT4.単価, 0)"
    gSL_Select = gSL_Select & "                 ELSE SUM(入荷数) * ISNULL(KT4.単価,0) END"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "         END 金額,"

    gSL_Select = gSL_Select & "        ISNULL(SY.値, '') 自社工区, NY.受入"
    gSL_Select = gSL_Select & "   FROM 材料入荷テーブル NY"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ HN"
    gSL_Select = gSL_Select & "             ON HN.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料取引先マスタ HTM"
    gSL_Select = gSL_Select & "             ON HTM.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "            AND HTM.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND HTM.受入 = NY.受入"
    gSL_Select = gSL_Select & "            AND HTM.品目 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "             ON TR.取引先CD = HTM.契約先"
    gSL_Select = gSL_Select & "            AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "             ON TR1.取引先CD = NY.仕入先"
    gSL_Select = gSL_Select & "            AND TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "             ON SY.区分名称 = '自社工区'"
    gSL_Select = gSL_Select & "            AND SY.英数字1 = NY.受入"
    gSL_Select = gSL_Select & "            AND SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料購入単価マスタ KT"
    gSL_Select = gSL_Select & "             ON KT.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "            AND KT.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND KT.契約先 = HTM.契約先"
    gSL_Select = gSL_Select & "            AND KT.仮単価区分 = ''"
    gSL_Select = gSL_Select & "            AND KT.履歴NO = (SELECT   MAX(履歴NO) FROM 材料購入単価マスタ KT1"
    gSL_Select = gSL_Select & "                              Where KT1.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "                                AND   KT1.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT1.仮単価区分 = ''"
    gSL_Select = gSL_Select & "                                AND ((KT1.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT1.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT1.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT1.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT1.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT1.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料購入単価マスタ KT2"
    gSL_Select = gSL_Select & "             ON KT2.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "            AND KT2.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND KT2.契約先 = HTM.契約先"
    gSL_Select = gSL_Select & "            AND KT2.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "            AND KT2.履歴NO = (SELECT   MAX(履歴NO) FROM 材料購入単価マスタ KT12"
    gSL_Select = gSL_Select & "                              Where KT12.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "                                AND   KT12.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT12.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "                                AND ((KT12.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT12.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT12.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT12.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT12.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT12.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料購入単価マスタ KT3"
    gSL_Select = gSL_Select & "             ON KT3.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "            AND KT3.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND KT3.契約先 = ''"
    gSL_Select = gSL_Select & "            AND KT3.仮単価区分 = ''"
    gSL_Select = gSL_Select & "            AND KT3.履歴NO = (SELECT   MAX(履歴NO) FROM 材料購入単価マスタ KT13"
    gSL_Select = gSL_Select & "                              Where KT13.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "                                AND   KT13.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT13.仮単価区分 = ''"
    gSL_Select = gSL_Select & "                                AND ((KT13.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT13.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT13.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT13.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT13.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT13.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料購入単価マスタ KT4"
    gSL_Select = gSL_Select & "             ON KT4.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "            AND KT4.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND KT4.契約先 = ''"
    gSL_Select = gSL_Select & "            AND KT4.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "            AND KT4.履歴NO = (SELECT   MAX(履歴NO) FROM 材料購入単価マスタ KT14"
    gSL_Select = gSL_Select & "                              Where KT14.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "                                AND   KT14.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT14.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "                                AND ((KT14.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT14.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT14.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT14.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT14.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT14.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        WHERE  ( HTM.手配区分 = 0                               " '*** 自給/直送支給(仕入先と契約先が同一のみ)
'    gSL_Select = gSL_Select & "           OR    HTM.手配区分 = 1 )"
    gSL_Select = gSL_Select & "           OR  ( HTM.手配区分 = 1 AND HTM.取引先 = HTM.契約先 ) )"         '2009/03/14
    gSL_Select = gSL_Select & "        AND   NY.入荷日 >= '" & RTrim(iH1_SYMD) & "'"
    gSL_Select = gSL_Select & "        AND   NY.入荷日 <= '" & RTrim(iH1_EYMD) & "'"
    If RTrim(iH1_Torcd) <> "" Then
       gSL_Select = gSL_Select & "         AND   HTM.契約先 = '" & RTrim(iH1_Torcd) & "'"
    End If
    
    ' 自社工区
    Call CisComb1.CodeGet
    If RTrim(CisComb1.Code) <> "" Then
       gSL_Select = gSL_Select & "     AND   SY.値 = " & CisFun.Val2(CisComb1.Code)
    End If
    
    gSL_Select = gSL_Select & " GROUP BY TR.略称, HTM.契約先,KT.単価, KT2.単価, KT3.単価, KT4.単価,"
    gSL_Select = gSL_Select & "          ISNULL(SY.値, ''), NY.受入,NY.材料管理番号, NY.仕入先,TR1.略称,HTM.実績換算値"

    
    gSL_Select = gSL_Select & " UNION ALL "
    
    
    '***** 材料入荷テーブル：材料有償仕入（受給）
    gSL_Select = gSL_Select & "SELECT  SUBSTRING(HOST_NAME(),1,20) 処理端末,3 区分,'有償支給' 区分名称,"
    'gSL_Select = gSL_Select & "SELECT  SUBSTRING(HOST_NAME(),1,20) 処理端末,4 区分,'材料' 区分名称,"
    gSL_Select = gSL_Select & "        ISNULL(HTM.契約先, '') 契約先, ISNULL(HTM.契約先,'') + ':' + ISNULL(TR.略称, '') 契約先名,"
    gSL_Select = gSL_Select & "        '' 品番,NY.材料管理番号, NY.仕入先,ISNULL(NY.仕入先,'') + ':' + ISNULL(TR1.略称, '') 仕入先名 ,"
    gSL_Select = gSL_Select & "        SUM(入荷数) 入荷数,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN ISNULL(KT.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN ISNULL(KT2.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN ISNULL(KT3.単価, 0)"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN ISNULL(KT4.単価, 0)"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "        END 単価,"
    gSL_Select = gSL_Select & "        0 有償支給,"
    gSL_Select = gSL_Select & "        0 外注材料,"
    gSL_Select = gSL_Select & "        0 外注購入,"
    gSL_Select = gSL_Select & "        0 有償支給計,"
    gSL_Select = gSL_Select & "        0 外注材料計,"
    gSL_Select = gSL_Select & "        0 外注購入計,"
    gSL_Select = gSL_Select & "        HTM.実績換算値,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(HTM.実績換算値,0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * HTM.実績換算値 * IsNull(KT.単価, 0)"
    gSL_Select = gSL_Select & "                 ELSE SUM(入荷数) * ISNULL(KT.単価,0) END"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(HTM.実績換算値,0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * HTM.実績換算値 * IsNull(KT2.単価, 0)"
    gSL_Select = gSL_Select & "                 ELSE SUM(入荷数) * ISNULL(KT2.単価,0)  END"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(HTM.実績換算値,0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * HTM.実績換算値 * IsNull(KT3.単価, 0)"
    gSL_Select = gSL_Select & "                 ELSE SUM(入荷数) * ISNULL(KT3.単価,0) END"
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN"
    gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(HTM.実績換算値,0) <> 0 THEN"
    gSL_Select = gSL_Select & "                     Sum (入荷数) * HTM.実績換算値 * IsNull(KT4.単価, 0)"
    gSL_Select = gSL_Select & "                 ELSE SUM(入荷数) * ISNULL(KT4.単価,0) END"
    gSL_Select = gSL_Select & "             ELSE 0"
    gSL_Select = gSL_Select & "         END 金額,"

    gSL_Select = gSL_Select & "        ISNULL(SY.値, '') 自社工区, NY.受入"
    gSL_Select = gSL_Select & "   FROM 材料入荷テーブル NY"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ HN"
    gSL_Select = gSL_Select & "             ON HN.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料取引先マスタ HTM"
    gSL_Select = gSL_Select & "             ON HTM.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "            AND HTM.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND HTM.受入 = NY.受入"
    gSL_Select = gSL_Select & "            AND HTM.品目 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "             ON TR.取引先CD = HTM.契約先"
    gSL_Select = gSL_Select & "            AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "             ON TR1.取引先CD = NY.仕入先"
    gSL_Select = gSL_Select & "            AND TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "             ON SY.区分名称 = '自社工区'"
    gSL_Select = gSL_Select & "            AND SY.英数字1 = NY.受入"
    gSL_Select = gSL_Select & "            AND SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料購入単価マスタ KT"
    gSL_Select = gSL_Select & "             ON KT.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "            AND KT.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND KT.契約先 = HTM.契約先"
    gSL_Select = gSL_Select & "            AND KT.仮単価区分 = ''"
    gSL_Select = gSL_Select & "            AND KT.履歴NO = (SELECT   MAX(履歴NO) FROM 材料購入単価マスタ KT1"
    gSL_Select = gSL_Select & "                              Where KT1.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "                                AND   KT1.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT1.仮単価区分 = ''"
    gSL_Select = gSL_Select & "                                AND ((KT1.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT1.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT1.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT1.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT1.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT1.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料購入単価マスタ KT2"
    gSL_Select = gSL_Select & "             ON KT2.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "            AND KT2.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND KT2.契約先 = HTM.契約先"
    gSL_Select = gSL_Select & "            AND KT2.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "            AND KT2.履歴NO = (SELECT   MAX(履歴NO) FROM 材料購入単価マスタ KT12"
    gSL_Select = gSL_Select & "                              Where KT12.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "                                AND   KT12.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT12.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "                                AND ((KT12.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT12.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT12.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT12.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT12.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT12.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料購入単価マスタ KT3"
    gSL_Select = gSL_Select & "             ON KT3.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "            AND KT3.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND KT3.契約先 = ''"
    gSL_Select = gSL_Select & "            AND KT3.仮単価区分 = ''"
    gSL_Select = gSL_Select & "            AND KT3.履歴NO = (SELECT   MAX(履歴NO) FROM 材料購入単価マスタ KT13"
    gSL_Select = gSL_Select & "                              Where KT13.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "                                AND   KT13.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT13.仮単価区分 = ''"
    gSL_Select = gSL_Select & "                                AND ((KT13.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT13.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT13.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT13.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT13.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT13.適用終了日付 = ''))"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料購入単価マスタ KT4"
    gSL_Select = gSL_Select & "             ON KT4.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "            AND KT4.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "            AND KT4.契約先 = ''"
    gSL_Select = gSL_Select & "            AND KT4.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "            AND KT4.履歴NO = (SELECT   MAX(履歴NO) FROM 材料購入単価マスタ KT14"
    gSL_Select = gSL_Select & "                              Where KT14.材料管理番号 = NY.材料管理番号"
    gSL_Select = gSL_Select & "                                AND   KT14.取引先 = NY.仕入先"
    gSL_Select = gSL_Select & "                                AND   KT14.仮単価区分 = '1'"
    gSL_Select = gSL_Select & "                                AND ((KT14.適用開始日付 < NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT14.適用開始日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT14.適用開始日付 = '')"
    gSL_Select = gSL_Select & "                                AND ((KT14.適用終了日付 > NY.入荷日"
    gSL_Select = gSL_Select & "                                AND   KT14.適用終了日付 <> '')"
    gSL_Select = gSL_Select & "                                 OR   KT14.適用終了日付 = ''))"
'    gSL_Select = gSL_Select & "        Where HTM.手配区分 = 3 "                                  '*** 受給
    gSL_Select = gSL_Select & "        Where ( HTM.手配区分 = 3 "                                 '*** 受給/直送支給(仕入先と契約先が違う)
    gSL_Select = gSL_Select & "           OR ( HTM.手配区分 = 1 AND HTM.取引先 <> HTM.契約先 ) )"         '2009/03/14
    gSL_Select = gSL_Select & "        AND   NY.入荷日 >= '" & RTrim(iH1_SYMD) & "'"
    gSL_Select = gSL_Select & "        AND   NY.入荷日 <= '" & RTrim(iH1_EYMD) & "'"
    If RTrim(iH1_Torcd) <> "" Then
       gSL_Select = gSL_Select & "         AND   HTM.契約先 = '" & RTrim(iH1_Torcd) & "'"
    End If
    
    ' 自社工区
    Call CisComb1.CodeGet
    If RTrim(CisComb1.Code) <> "" Then
       gSL_Select = gSL_Select & "     AND   SY.値 = " & CisFun.Val2(CisComb1.Code)
    End If
    
    gSL_Select = gSL_Select & " GROUP BY TR.略称, HTM.契約先,KT.単価, KT2.単価, KT3.単価, KT4.単価,"
    gSL_Select = gSL_Select & "          ISNULL(SY.値, ''), NY.受入,NY.材料管理番号, NY.仕入先,TR1.略称,HTM.実績換算値"

    
    gSL_Select = gSL_Select & " UNION ALL "
    
    
    '***** 購入実績テーブル：画面入力データ
    gSL_Select = gSL_Select & "SELECT  SUBSTRING(HOST_NAME(),1,20) 処理端末,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(SY1.数字3,0) = 1 THEN 1"
    gSL_Select = gSL_Select & "             WHEN ISNULL(SY1.数字3,0) = 2 THEN 2"
    gSL_Select = gSL_Select & "             WHEN ISNULL(SY1.数字3,0) = 3 THEN 3"
    gSL_Select = gSL_Select & "             WHEN ISNULL(SY1.数字3,0) = 4 THEN 4"
    gSL_Select = gSL_Select & "             ELSE 2 END 区分,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(SY1.数字3,0) = 1 THEN '外注先費用'"
    gSL_Select = gSL_Select & "             WHEN ISNULL(SY1.数字3,0) = 2 THEN '補助材'"
    gSL_Select = gSL_Select & "             WHEN ISNULL(SY1.数字3,0) = 3 THEN '有償支給'"
    gSL_Select = gSL_Select & "             WHEN ISNULL(SY1.数字3,0) = 4 THEN '材料'"
    gSL_Select = gSL_Select & "             Else '補助材' END 区分名称,"
    gSL_Select = gSL_Select & "        ISNULL(NY.契約先, '') 契約先, ISNULL(NY.契約先,'') + ':' + ISNULL(TR.略称, '') 契約先名,"
    gSL_Select = gSL_Select & "        NY.品番,0 材料管理番号,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(NY.仕入先,'') <> '' THEN NY.仕入先 ELSE NY.契約先 END 仕入先 ,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(NY.仕入先,'') <> '' THEN ISNULL(NY.仕入先,'') + ':' + ISNULL(TR1.略称, '') ELSE ISNULL(NY.契約先,'') + ':' + ISNULL(TR.略称, '') END 仕入先名 ,"
    gSL_Select = gSL_Select & "        SUM(数量) 入荷数,"
    gSL_Select = gSL_Select & "        NY.単価,"
    gSL_Select = gSL_Select & "        0 有償支給,"
    gSL_Select = gSL_Select & "        0 外注材料,"
    gSL_Select = gSL_Select & "        0 外注購入,"
    gSL_Select = gSL_Select & "        0 有償支給計,"
    gSL_Select = gSL_Select & "        0 外注材料計,"
    gSL_Select = gSL_Select & "        0 外注購入計,"
    gSL_Select = gSL_Select & "        0 実績換算値,"
    gSL_Select = gSL_Select & "        SUM(NY.金額) 金額,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(SY.値, '') <> '' THEN ISNULL(SY.値, '') ELSE '0' END 自社工区,"
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(NY.受入,'') <> '' THEN ISNULL(NY.受入,'') ELSE '1' END 受入"
    gSL_Select = gSL_Select & "   FROM 購入実績テーブル NY"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "             ON TR.取引先CD = NY.契約先"
    gSL_Select = gSL_Select & "            AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "             ON TR1.取引先CD = NY.仕入先"
    gSL_Select = gSL_Select & "            AND TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "             ON SY.区分名称 = '自社工区'"
    gSL_Select = gSL_Select & "            AND SY.英数字1 = NY.受入"
    gSL_Select = gSL_Select & "            AND SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "             ON SY1.区分名称 = '買掛識別子'"
    gSL_Select = gSL_Select & "            AND SY1.値 = NY.識別子"
    gSL_Select = gSL_Select & "            AND SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & "        WHERE  ISNULL(NY.入力区分,'') = ''"
    gSL_Select = gSL_Select & "          AND  ISNULL(NY.発注管理NO,0) = 0"
    gSL_Select = gSL_Select & "          AND  NY.検収日 >= '" & RTrim(iH1_SYMD) & "'"
    gSL_Select = gSL_Select & "          AND  NY.検収日 <= '" & RTrim(iH1_EYMD) & "'"
    If RTrim(iH1_Torcd) <> "" Then
       gSL_Select = gSL_Select & "         AND   NY.契約先 = '" & RTrim(iH1_Torcd) & "'"
    End If
    
    ' 自社工区
    Call CisComb1.CodeGet
    If RTrim(CisComb1.Code) <> "" Then
       gSL_Select = gSL_Select & "     AND   SY.値 = " & CisFun.Val2(CisComb1.Code)
    End If
    
    gSL_Select = gSL_Select & " GROUP BY TR.略称, NY.契約先,NY.単価,"
    gSL_Select = gSL_Select & "          ISNULL(SY.値, ''), NY.受入,NY.品番, NY.仕入先,TR1.略称,NY.識別子,SY1.数字3"
    
    
    With CisDB
        '+---------- ワークにデータ出力 -----------+
        .SQL = gSL_Select
        Call CisDB.DBExec
    End With
    
        
    '*** ワークデータを集計
    If H1Op_OutPut(0) Then
       gSL_Select = "SELECT 区分,区分名称,仕入先,仕入先名,"
    Else
       gSL_Select = "SELECT 区分,区分名称,契約先 仕入先,契約先名 仕入先名,"
    End If
    gSL_Select = gSL_Select & " SUM(有償支給計) 有償支給計,"
    gSL_Select = gSL_Select & " SUM(外注材料計) 外注材料計,"
    gSL_Select = gSL_Select & " SUM(外注購入計) 外注購入計,"
    gSL_Select = gSL_Select & " SUM(金額) 金額"
    gSL_Select = gSL_Select & "  From 仕入明細書ワーク"
    gSL_Select = gSL_Select & " Where 金額 <> 0"
    gSL_Select = gSL_Select & "   And 処理端末 = SUBSTRING(HOST_NAME(),1,20)"
    If H1Op_OutPut(0) Then
       gSL_Select = gSL_Select & " GROUP BY 区分,区分名称,仕入先,仕入先名"
       gSL_Select = gSL_Select & " ORDER BY 区分,仕入先"
    Else
       gSL_Select = gSL_Select & " GROUP BY 区分,区分名称,契約先,契約先名"
       gSL_Select = gSL_Select & " ORDER BY 区分,契約先"
    End If
       
    If Not SMWRead(gSL_Select, , 1) Then
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
        lb_msg.Caption = "【 仕入明細表 印刷中 】"
        DoEvents
        With CisPrtCOK0530
            .cDataMaxCnt = CisDB.RecordCount
            Set .iPrtForm = COK0530L
            .dpMSG = "仕入明細表　印刷中"
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

    lb_msg.Caption = "【 仕入明細データ 出力中 】"
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
    Do Until Not SMW_RDSTS
    '   >> ｶｳﾝﾄ
        OutCnt = OutCnt + 1
    '   >> Body部 編集＆出力
        Call CsvEdit_Body1(OutStr, OutCnt)
        Print #OutFNo, OutStr

        CisFrm.MeterUpdate OutCnt

    '   >> 次のﾚｺｰﾄﾞを読み込む
        Call SMWReadNext(1)
    Loop
    Call SMWClose(1)
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
    CsvStr = CsvStr & Mid(RTrim(iH1_SYMD), 1, 4) & "年" & Mid(RTrim(iH1_SYMD), 5, 2) & "月" & "度仕入明細表"
End Sub
Private Sub CsvEdit_Head2(CsvStr As String)
    CsvStr = ""
    CsvStr = CsvStr & "№"
    CsvStr = CsvStr & ",区分"
    If COK0530_OutPut = 0 Then
       CsvStr = CsvStr & ",仕入先"
    Else
       CsvStr = CsvStr & ",契約先"
    End If
    CsvStr = CsvStr & ",有償支給"
    CsvStr = CsvStr & ",外注材料"
    CsvStr = CsvStr & ",外注購入"
    CsvStr = CsvStr & ",金額"
End Sub
'+----------------------------+
'+    Body部　編集処理
'+----------------------------+
Private Sub CsvEdit_Body1(CsvStr As String, CsvCnt As Long)
    With SMW
        CsvStr = ""
        CsvStr = CsvStr & CsvCnt
        CsvStr = CsvStr & "," & RTrim(.区分名称)
        CsvStr = CsvStr & "," & RTrim(.仕入先名)
        CsvStr = CsvStr & "," & Format(.有償支給計, "########0.00")
        CsvStr = CsvStr & "," & Format(.外注材料計, "########0.00")
        CsvStr = CsvStr & "," & Format(.外注購入計, "########0.00")
        CsvStr = CsvStr & "," & Format(.金額, "########0.00")
    End With
End Sub


