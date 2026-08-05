VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BA5D7819-2B67-11D4-806A-00000EA57F9E}#3.0#0"; "CISYMD~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CSJ0100 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "プレス製造指示書発行"
   ClientHeight    =   10125
   ClientLeft      =   -1185
   ClientTop       =   1650
   ClientWidth     =   15375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10125
   ScaleWidth      =   15375
   Begin VB.PictureBox Csv_Bar 
      Height          =   345
      Left            =   4357
      ScaleHeight     =   285
      ScaleWidth      =   6600
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1935
      Visible         =   0   'False
      Width           =   6660
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   585
      Left            =   5715
      Top             =   8505
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
      Height          =   5760
      Left            =   4357
      ScaleHeight     =   5932.66
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6585
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2310
      Width           =   6645
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   5595
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   9869
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
            Height          =   2085
            Left            =   1815
            Top             =   2265
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   3678
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
            Begin vsFlexLib.vsFlexArray vsGrid1 
               Height          =   2025
               Left            =   30
               TabIndex        =   4
               Top             =   30
               Width           =   1485
               _Version        =   196608
               _ExtentX        =   2619
               _ExtentY        =   3572
               _StockProps     =   228
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   1
               ConvInfo        =   1418783674
               FormatString    =   "  |順位|選択"
               Rows            =   6
               Cols            =   4
               BackColor       =   12648447
               HighLight       =   2
               SelectionMode   =   1
            End
         End
         Begin CisYMD_V60.CisYMD iH1_SYMD 
            Height          =   330
            Left            =   1800
            TabIndex        =   1
            Top             =   930
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
            cChkResult      =   0   'False
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "　出　力　先"
            Height          =   600
            Left            =   60
            TabIndex        =   15
            Top             =   4545
            Width           =   6345
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "プリンター"
               Height          =   240
               Index           =   0
               Left            =   1065
               TabIndex        =   5
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
               TabIndex        =   6
               Top             =   270
               Width           =   1875
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "CSV"
               Height          =   240
               Index           =   2
               Left            =   4380
               TabIndex        =   7
               Top             =   240
               Width           =   975
            End
         End
         Begin CisText_V60.CisText iH1_Kikai 
            Height          =   360
            Left            =   1800
            TabIndex        =   3
            Top             =   1755
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
         Begin Cis3D_v60.CIS3D H1lb_Kikai 
            Height          =   360
            Left            =   2610
            Top             =   1755
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
               TabIndex        =   9
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
                  Caption         =   "プレス製造指示書の発行を行います"
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
            Top             =   5175
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
            Top             =   5205
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
            Height          =   390
            Index           =   1
            Left            =   405
            Top             =   900
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   688
            ForeColor       =   16711680
            Caption         =   "指 示 日"
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
            Height          =   315
            Index           =   0
            Left            =   540
            Top             =   1770
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "機　　械"
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
         Begin Cis3D_v60.CIS3D PB_Look 
            Height          =   360
            Index           =   1
            Left            =   2265
            Top             =   1755
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
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   390
            Index           =   0
            Left            =   405
            Top             =   2205
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   688
            ForeColor       =   16711680
            Caption         =   "優先順位"
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
         Begin Cis3D_v60.CIS3D H1lb_Busyo 
            Height          =   360
            Left            =   3120
            Top             =   1335
            Width           =   2820
            _ExtentX        =   4974
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
            cPositionX      =   50
         End
         Begin CisText_V60.CisText iH1_Busyo 
            Height          =   360
            Left            =   1800
            TabIndex        =   2
            Top             =   1335
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
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   420
            Index           =   2
            Left            =   390
            Top             =   1305
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   741
            ForeColor       =   16711680
            Caption         =   "部　　署"
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
         Begin Cis3D_v60.CIS3D PB_Look 
            Height          =   360
            Index           =   0
            Left            =   2775
            Top             =   1335
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
      Caption         =   "【 プレス製造指示書発行 】"
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
      Left            =   8370
      Top             =   8505
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
      Left            =   7005
      Top             =   8505
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
      Left            =   4350
      Top             =   5235
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
      Left            =   6990
      TabIndex        =   14
      Top             =   9180
      Width           =   1485
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8370
      TabIndex        =   12
      Top             =   9180
      Width           =   1485
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5700
      TabIndex        =   11
      Top             =   9180
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
      TabIndex        =   10
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CSJ0100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  プレス製造指示書 発行
'**       フォームID    :  CSJ0100
'**       処理概要      :
'**
'**       作  成  日    :  2010/01/27 By CIS
'**       変  更  日    :  2010/05/31 By CIS 優先順位を複数選択可能に変更
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim CisVsGrid       As New CisVsGrid3
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
    
    Dim wSumK           As Long
    Dim wSumP           As Long
    Dim wJyun           As String
    Dim wBusyo          As String
    Dim wKikai          As String
    Dim wYusen          As String       ' 2010/05/31 Insert
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
'------------------ 2010/05/31 Insert Start
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 3
        .InitGet = False
        .Init
    End With
    Call GridInit
    Call GridSet
'------------------ 2010/05/31 Insert End
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
    
'   CSVﾌｧｲﾙ出力先
    mCsvName = Mid(CisFun.WinRoot, 1, 3) & "プレス製造指示書.csv"
    mCsvKey = "CSJ0100"
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

    Select Case Key_Code   '【検索】
        Case vbKeyF7:
                            If PB_Look(0).Tag <> "" Then
                                Select Case PB_Look(0).Tag
                                    Case "iH1_Busyo": PB_Look_Click (0)
                                    Case "iH1_Kikai": PB_Look_Click (1)
                                End Select
                            End If
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
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click(INDEX As Integer)
    If INDEX = 0 Then Call Look_Busyo: Exit Sub
    If INDEX = 1 Then Call Look_Kikai: Exit Sub
End Sub
'+----------------------------+
'+       取引先検索           +
'+----------------------------+
Private Sub Look_Busyo()
    iH1_Busyo.SetFocus

    RV_Left = 0:    RV_Top = 0
    RV_TorKb = 7: RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
       iH1_Busyo = RTrim(RV_TorcdK)
       Call iH1_Busyo_LostFocus
    End If

End Sub
'+----------------------------------------------------------------------------+
'+      機械(名称マスタ) 検索
'+----------------------------------------------------------------------------+
Private Sub Look_Kikai()
    iH1_Kikai.SetFocus
    
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
'+-----------------------------+
'+  iH1_Busyo
'+-----------------------------+
Private Sub iH1_Busyo_GotFocus()
    PB_Look(0).Tag = ActiveControl.Name
End Sub
Private Sub iH1_Busyo_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look(0).Tag = ""
    End If
    
    H1lb_Busyo = ""
    If TorNmGet(RTrim(iH1_Busyo), 7) Then
        H1lb_Busyo = RTrim(TRM_RName)
    End If

End Sub
'+-----------------------------+
'+  iH1_Kikai
'+-----------------------------+
Private Sub iH1_Kikai_GotFocus()
    PB_Look(0).Tag = ActiveControl.Name
End Sub
Private Sub iH1_Kikai_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look(0).Tag = ""
    End If

    H1lb_Kikai = ""
    If MeisyoGet("機械", RTrim(iH1_Kikai)) Then
        H1lb_Kikai = RTrim(SYM_Meisyo)
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
    Dim YM0 As String
    Dim YM1 As String
    Dim YM2 As String


'   開始日
    If Trim(iH1_SYMD) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    指示日を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SYMD.SetFocus
        Exit Function
    End If
    
    '
    YM0 = Mid(iH1_SYMD, 1, 6)
        
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add RTrim(YM0), M, 1, ResultDate: YM1 = .Date_Result
        .Date_Add RTrim(YM0), M, 2, ResultDate: YM2 = .Date_Result
    End With
    
    If Not iH1_SYMD.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    指示日を正しく入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_SYMD.SetFocus
        Exit Function
    End If
    
    If iH1_Busyo <> "" Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT *"
        gSL_Select = gSL_Select & "  FROM 取引先マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先CD = '" & Trim(iH1_Busyo) & "'"
        gSL_Select = gSL_Select & "   AND 取引先区分 = 1"
        gSL_Select = gSL_Select & "   AND 社内区分 = 1"
        If Not TRMRead(gSL_Select, , 0) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    部署未登録！            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Busyo.SetFocus
            Exit Function
        End If
    End If
    
    If iH1_Kikai <> "" Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT *"
        gSL_Select = gSL_Select & "  FROM 名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = '機械'"
        gSL_Select = gSL_Select & "   AND 値 = '" & Trim(iH1_Kikai) & "'"
        If Not SYMRead(gSL_Select, , 0) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    機械未登録！            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Kikai.SetFocus
            Exit Function
        End If
    End If
    
    With CisFun
        .Where_CreateKB = NewDocument
        Call .Where_Create(Numeric, "HK.日報出力", "0")
        Call .Where_Create(Numeric, "TH.プレススポット区分", "0")
        
'------------------------------------------------------- 2010/05/31 Delete Start
'        If Trim(iH1_Jyun) <> "" Then
'            Call .Where_Create(AlphaMe, "HK.優先順位", Trim(iH1_Jyun))
'        End If
'------------------------------------------------------- 2010/05/31 Delete End
        
        If Trim(iH1_Busyo) <> "" Then
            Call .Where_Create(AlphaMe, "HK.取引先", Trim(iH1_Busyo))
        End If
        If Trim(iH1_Kikai) <> "" Then
            Call .Where_Create(AlphaMe, "HK.機械", Trim(iH1_Kikai))
        End If
        gStr = .Where_Phrase
        
        If gStr = "" Then
            gStr = " WHERE HK.日報出力 = 0" & vbCrLf
            gStr = gStr & "   AND TH.プレススポット区分 = 0" & vbCrLf
            
'            gStr = gStr & "   AND ISNULL(HM.表示品番,'') <> ''" & vbCrLf
        Else
            gStr = gStr & "   AND HK.日報出力 = 0" & vbCrLf
            gStr = gStr & "   AND TH.プレススポット区分 = 0" & vbCrLf
            
'            gStr = gStr & "   AND ISNULL(HM.表示品番,'') <> ''" & vbCrLf
        End If
'------------------------------------------------------- 2010/05/31 Insert Start
        wYusen = ""
        For gInt = 1 To vsGrid1.Rows - 1
            If vsGrid1.TextMatrix(gInt, CisVsGrid.SelectCol) <> "" Then
                If wYusen = "" Then
                    wYusen = wYusen + "'" + vsGrid1.TextMatrix(gInt, 1) + "'"
                Else
                    wYusen = wYusen + ",'" + vsGrid1.TextMatrix(gInt, 1) + "'"
                End If
            End If
        Next gInt
        If Trim(wYusen) <> "" Then
            gStr = gStr & "   AND HK.優先順位 IN (" + wYusen + ")"
        End If
'------------------------------------------------------- 2010/05/31 Insert End
    End With
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT ISNULL(TH.略称,'') 取引先名,ISNULL(SY1.値名称,'') 機械名,ISNULL(SY2.値名称,'') 機械NO名" & vbCrLf
    gSL_Select = gSL_Select & "      ,ISNULL(HM.表示品番,'') 表示品番,ISNULL(TM.略称,'') 得意先名,HK.優先順位,HK.加工数" & vbCrLf
    gSL_Select = gSL_Select & "      ,(HK.加工数 * HK.工程数) パンチ数,(ISNULL(ZM.材質,'') + ISNULL(ZM.表示寸法,'')) 材料名" & vbCrLf
    gSL_Select = gSL_Select & "      ,ISNULL(HN1.内示当月,'') 内示当月,ISNULL(HN2.内示翌月,'') 内示翌月,ISNULL(HN3.内示翌々月,'') 内示翌々月" & vbCrLf
    gSL_Select = gSL_Select & "      ,HK.備考,HK.品番,HK.取引先,HK.受入,ZH.材料管理番号,HK.機械" & vbCrLf
    gSL_Select = gSL_Select & "  FROM 品番工順マスタ HK" & vbCrLf
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 品番マスタ HM" & vbCrLf
    gSL_Select = gSL_Select & "    ON HM.品番 = HK.品番" & vbCrLf
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 取引先マスタ TM" & vbCrLf
    gSL_Select = gSL_Select & "    ON TM.取引先CD = HM.得意先" & vbCrLf
    gSL_Select = gSL_Select & "   AND TM.取引先区分 = 0" & vbCrLf
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 取引先マスタ TH" & vbCrLf
    gSL_Select = gSL_Select & "    ON TH.取引先CD = HK.取引先" & vbCrLf
    gSL_Select = gSL_Select & "   AND TH.取引先区分 = 1" & vbCrLf
    gSL_Select = gSL_Select & "   AND TH.社内区分 = 1" & vbCrLf
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN (" & vbCrLf
    gSL_Select = gSL_Select & "                    SELECT MAX(材料管理番号) 材料管理番号,使用品番" & vbCrLf
    gSL_Select = gSL_Select & "                      FROM 材料品番マスタ" & vbCrLf
    gSL_Select = gSL_Select & "                     GROUP BY 使用品番" & vbCrLf
    gSL_Select = gSL_Select & "                   ) ZH" & vbCrLf
    gSL_Select = gSL_Select & "    ON ZH.使用品番 = HK.品番" & vbCrLf
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 材料マスタ ZM" & vbCrLf
    gSL_Select = gSL_Select & "    ON ZM.材料管理番号 = ZH.材料管理番号" & vbCrLf
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN (" & vbCrLf
    gSL_Select = gSL_Select & "                    SELECT 内示年月,品番,手配先,受入,計画年月,SUM(数量) 内示当月" & vbCrLf
    gSL_Select = gSL_Select & "                      FROM 発注内示テーブル" & vbCrLf
'    gSL_Select = gSL_Select & "                     WHERE 内示年月 = CONVERT(VarChar(6), GETDATE(), 112)" & vbCrLf
'    gSL_Select = gSL_Select & "                       AND 計画年月 = CONVERT(VARCHAR(6),GETDATE(),112)" & vbCrLf
    gSL_Select = gSL_Select & "                     WHERE 内示年月 = '" & YM0 & "'" & vbCrLf
    gSL_Select = gSL_Select & "                       AND 計画年月 = '" & YM0 & "'" & vbCrLf
    gSL_Select = gSL_Select & "                     GROUP BY 内示年月,品番,手配先,受入,計画年月" & vbCrLf
    gSL_Select = gSL_Select & "                   ) HN1" & vbCrLf
'    gSL_Select = gSL_Select & "    ON HN1.内示年月 = CONVERT(VARCHAR(6),GETDATE(),112)" & vbCrLf
'    gSL_Select = gSL_Select & "   AND HN1.計画年月 = CONVERT(VARCHAR(6),GETDATE(),112)" & vbCrLf
    gSL_Select = gSL_Select & "    ON HN1.内示年月 = '" & YM0 & "'" & vbCrLf
    gSL_Select = gSL_Select & "   AND HN1.計画年月 = '" & YM0 & "'" & vbCrLf
    gSL_Select = gSL_Select & "   AND HN1.品番 = HK.品番" & vbCrLf
    gSL_Select = gSL_Select & "   AND HN1.手配先 = HK.取引先" & vbCrLf
    gSL_Select = gSL_Select & "   AND HN1.受入 = HK.受入" & vbCrLf
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN (" & vbCrLf
    gSL_Select = gSL_Select & "                    SELECT 内示年月,品番,手配先,受入,計画年月,SUM(数量) 内示翌月" & vbCrLf
    gSL_Select = gSL_Select & "                      FROM 発注内示テーブル" & vbCrLf
'    gSL_Select = gSL_Select & "                     WHERE 内示年月 = CONVERT(VarChar(6), GETDATE(), 112)" & vbCrLf
'    gSL_Select = gSL_Select & "                       AND 計画年月 = CONVERT(VARCHAR(6),DATEADD(M,1,GETDATE()),112)" & vbCrLf
    gSL_Select = gSL_Select & "                     WHERE 内示年月 = '" & YM0 & "'" & vbCrLf
    gSL_Select = gSL_Select & "                       AND 計画年月 = '" & YM1 & "'" & vbCrLf
    gSL_Select = gSL_Select & "                     GROUP BY 内示年月,品番,手配先,受入,計画年月" & vbCrLf
    gSL_Select = gSL_Select & "                   ) HN2" & vbCrLf
'    gSL_Select = gSL_Select & "    ON HN2.内示年月 = CONVERT(VARCHAR(6),GETDATE(),112)" & vbCrLf
'    gSL_Select = gSL_Select & "   AND HN2.計画年月 = CONVERT(VARCHAR(6),DATEADD(M,1,GETDATE()),112)" & vbCrLf
    gSL_Select = gSL_Select & "    ON HN2.内示年月 = '" & YM0 & "'" & vbCrLf
    gSL_Select = gSL_Select & "   AND HN2.計画年月 = '" & YM1 & "'" & vbCrLf
    gSL_Select = gSL_Select & "   AND HN2.品番 = HK.品番" & vbCrLf
    gSL_Select = gSL_Select & "   AND HN2.手配先 = HK.取引先" & vbCrLf
    gSL_Select = gSL_Select & "   AND HN2.受入 = HK.受入" & vbCrLf
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN (" & vbCrLf
    gSL_Select = gSL_Select & "                    SELECT 内示年月,品番,手配先,受入,計画年月,SUM(数量) 内示翌々月" & vbCrLf
    gSL_Select = gSL_Select & "                      FROM 発注内示テーブル" & vbCrLf
'    gSL_Select = gSL_Select & "                     WHERE 内示年月 = CONVERT(VarChar(6), GETDATE(), 112)" & vbCrLf
'    gSL_Select = gSL_Select & "                       AND 計画年月 = CONVERT(VARCHAR(6),DATEADD(M,2,GETDATE()),112)" & vbCrLf
    gSL_Select = gSL_Select & "                     WHERE 内示年月 = '" & YM0 & "'" & vbCrLf
    gSL_Select = gSL_Select & "                       AND 計画年月 = '" & YM2 & "'" & vbCrLf
    gSL_Select = gSL_Select & "                     GROUP BY 内示年月,品番,手配先,受入,計画年月" & vbCrLf
    gSL_Select = gSL_Select & "                   ) HN3" & vbCrLf
'    gSL_Select = gSL_Select & "    ON HN3.内示年月 = CONVERT(VARCHAR(6),GETDATE(),112)" & vbCrLf
'    gSL_Select = gSL_Select & "   AND HN3.計画年月 = CONVERT(VARCHAR(6),DATEADD(M,2,GETDATE()),112)" & vbCrLf
    gSL_Select = gSL_Select & "    ON HN3.内示年月 = '" & YM0 & "'" & vbCrLf
    gSL_Select = gSL_Select & "   AND HN3.計画年月 = '" & YM2 & "'" & vbCrLf
    gSL_Select = gSL_Select & "   AND HN3.品番 = HK.品番" & vbCrLf
    gSL_Select = gSL_Select & "   AND HN3.手配先 = HK.取引先" & vbCrLf
    gSL_Select = gSL_Select & "   AND HN3.受入 = HK.受入" & vbCrLf
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 名称マスタ SY1" & vbCrLf
    gSL_Select = gSL_Select & "    ON SY1.区分名称 = '機械'" & vbCrLf
    gSL_Select = gSL_Select & "   AND SY1.値 = HK.機械" & vbCrLf
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 名称マスタ SY2" & vbCrLf
    gSL_Select = gSL_Select & "    ON SY2.区分名称 = '機械NO'" & vbCrLf
    gSL_Select = gSL_Select & "   AND SY2.値 = HK.機械NO " & vbCrLf
    gSL_Select = gSL_Select & gStr
'-------------------------------------------------------------------------------------------- 2010/05/31 Update Start
'    gSL_Select = gSL_Select & " ORDER BY CASE WHEN ISNULL(HK.優先順位,'') = '' THEN 1 ELSE 0 END" & vbCrLf
'    gSL_Select = gSL_Select & "         ,HK.優先順位,HK.取引先,HK.機械,HK.機械NO,HK.品番" & vbCrLf
    gSL_Select = gSL_Select & " ORDER BY HK.取引先,CASE WHEN ISNULL(HK.優先順位,'') = '' THEN 1 ELSE 0 END" & vbCrLf
    gSL_Select = gSL_Select & "         ,HK.優先順位,HK.機械,HK.機械NO,HK.品番" & vbCrLf
'-------------------------------------------------------------------------------------------- 2010/05/31 Update End
       
    If Not HKTRead(gSL_Select, , 1) Then
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
        lb_msg.Caption = "【 プレス製造指示書 印刷中 】"
        DoEvents
        With CisPrtCSJ0100
            gSYMD = iH1_SYMD
            .cDataMaxCnt = CisDB.RecordCount
            Set .iPrtForm = CSJ0100L
            .dpMSG = "プレス製造指示書　印刷中"
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
    Call CsvEdit_Head3(OutStr)
    Print #OutFNo, OutStr
    Call CsvEdit_Head4(OutStr)
    Print #OutFNo, OutStr
'   >> ｸﾘｱ
    OutCnt = 0
    wSumK = 0
    wSumP = 0
    wJyun = HKT.優先順位
    wBusyo = HKT.取引先
    wKikai = HKT.機械
    Do Until Not HKT_RDSTS
    '   >> ｶｳﾝﾄ
        OutCnt = OutCnt + 1
    '   >> Body部 編集＆出力
        Call CsvEdit_Body1(OutStr, OutCnt)
        Print #OutFNo, OutStr
        wSumK = wSumK + HKT.加工数
        wSumP = wSumP + HKT.パンチ数
        
        CisFrm.MeterUpdate OutCnt

    '   >> 次のﾚｺｰﾄﾞを読み込む
        Call HKTReadNext(1)
        
        If wJyun <> HKT.優先順位 Or wBusyo <> HKT.取引先 Or wKikai = HKT.機械 Or Not HKT_RDSTS Then
            Call CsvEdit_Body2(OutStr, OutCnt)
            Print #OutFNo, OutStr
            wJyun = HKT.優先順位
            wBusyo = HKT.取引先
            wKikai = HKT.機械
            wSumK = 0
            wSumP = 0
        End If
    Loop
    Call HKTClose(1)
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
    CsvStr = CsvStr & "※ プレス製造指示書 ※"
End Sub
Private Sub CsvEdit_Head2(CsvStr As String)
    CsvStr = ""
    CsvStr = CsvStr & CisFun.Mid2(iH1_SYMD, 1, 4) & "年" & CisFun.Mid2(iH1_SYMD, 5, 2) & "月" & CisFun.Mid2(iH1_SYMD, 7, 2) & "日"
End Sub
Private Sub CsvEdit_Head3(CsvStr As String)
    CsvStr = ""
    CsvStr = CsvStr & ",,,,,,,,,,,,[ 内示数 ]"
End Sub
Private Sub CsvEdit_Head4(CsvStr As String)
    CsvStr = ""
    CsvStr = CsvStr & "№"
    CsvStr = CsvStr & ",部署"
    CsvStr = CsvStr & ",機械"
    CsvStr = CsvStr & ",機械No"
    CsvStr = CsvStr & ",プレス品番"
    CsvStr = CsvStr & ",得意先"
    CsvStr = CsvStr & ",SD(親)"
    CsvStr = CsvStr & ",優先順位"
    CsvStr = CsvStr & ",加工数"
    CsvStr = CsvStr & ",ﾊﾟﾝﾁ数"
    CsvStr = CsvStr & ",材料"
    CsvStr = CsvStr & ",材料在庫"
    CsvStr = CsvStr & ",当月"
    CsvStr = CsvStr & ",翌月"
    CsvStr = CsvStr & ",翌々月"
    CsvStr = CsvStr & ",プレス順(備考)"
    
End Sub
'+----------------------------+
'+    Body部　編集処理
'+----------------------------+
Private Sub CsvEdit_Body1(CsvStr As String, CsvCnt As Long)
    With HKT
        CsvStr = ""
        CsvStr = CsvStr & CsvCnt
        CsvStr = CsvStr & "," & RTrim(.取引先名)
        CsvStr = CsvStr & "," & RTrim(.機械名)
        CsvStr = CsvStr & "," & RTrim(.機械NO名)
        CsvStr = CsvStr & "," & RTrim(.表示品番)
        CsvStr = CsvStr & "," & RTrim(.得意先名)
        Call GetSD(.品番)                      ' 構成問合せ
        Call SelectKOL(.品番)
        If Not KOLRead(gSL_Select, , 0) Then
            CsvStr = CsvStr & ","
        Else
            CsvStr = CsvStr & "," & RTrim(KOL.表示品番)
        End If
        CsvStr = CsvStr & "," & RTrim(.優先順位)
        CsvStr = CsvStr & "," & Format(.加工数, "#")
        CsvStr = CsvStr & "," & Format(.パンチ数, "#")
        CsvStr = CsvStr & "," & RTrim(.材料名)
        Call GetZaikoSu(.材料管理番号)
        CsvStr = CsvStr & "," & Format(.材料在庫量, "#")
        CsvStr = CsvStr & "," & Format(.内示当月, "#")
        CsvStr = CsvStr & "," & Format(.内示翌月, "#")
        CsvStr = CsvStr & "," & Format(.内示翌々月, "#")
        CsvStr = CsvStr & "," & RTrim(.備考)
    End With
End Sub
Private Sub CsvEdit_Body2(CsvStr As String, CsvCnt As Long)
    CsvStr = ""
    CsvStr = CsvStr & ",,,,,[ 数量計 ],,"
    CsvStr = CsvStr & "," & Format(wSumK, "#")
    CsvStr = CsvStr & "," & Format(wSumP, "#")
End Sub
'**************************************************************************************************<< 2010/05/31 Insert Start >>
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "")
        Call .FixedSet(中中, 中中, あり, 10, 1, "順位")

'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 2, "Select")

'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1")

'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'******************************************'
'*****        行選択（クリック）       *****
'******************************************'
Private Sub VSGrid1_Click()
    
    With vsGrid1
        .Redraw = False
        For gLong = .Row To .RowSel
            If Trim(.TextMatrix(gLong, CisVsGrid.NoSpaceCol)) <> "" Then
                gStr = Trim(.TextMatrix(gLong, CisVsGrid.SelectCol))
                .Row = gLong
                If gStr = "X" Then
                   Call CisVsGrid.vsColor(消去個別)
                Else
                   Call CisVsGrid.vsColor(個別, mSelColor)
                End If
            End If
        Next gLong
        .Redraw = True
    End With

End Sub
'******************************************'
'*****        行選択（スペースｷｰ)      *****
'******************************************'
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    Call VSGrid1_Click
End Sub
'******************************************'
'*****       グリッド内容セット       *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long

    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT 優先順位"
    gSL_Select = gSL_Select & "  FROM 品番工順マスタ"
    gSL_Select = gSL_Select & " WHERE 優先順位 <> ''"
    gSL_Select = gSL_Select & " GROUP BY 優先順位"
    gSL_Select = gSL_Select & " ORDER BY 優先順位"

    Call HKTRead(gSL_Select)
    
    ' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避

' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not HKT_RDSTS
        With vsGrid1
            ID = ID + 1
            
            .TextMatrix(ID, 1) = HKT.優先順位
        End With
        Call HKTReadNext
    Loop
    Call HKTClose
    
    Call CisVsGrid.ReNumber
    
    With vsGrid1
         .Row = 1: .RowSel = .Rows - 1
         Call VSGrid1_Click
         
         .Row = 1
         .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    
End Sub
'**************************************************************************************************<< 2010/05/31 Insert End >>
