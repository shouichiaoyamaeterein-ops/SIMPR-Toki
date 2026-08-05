VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CNJ1110 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "内示数量チェック"
   ClientHeight    =   11190
   ClientLeft      =   60
   ClientTop       =   1530
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   Begin VB.PictureBox Csv_Bar 
      Height          =   345
      Left            =   4290
      ScaleHeight     =   285
      ScaleWidth      =   6555
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2160
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
      Top             =   7770
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
      Height          =   4920
      Left            =   4260
      ScaleHeight     =   5058.373
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   6600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2565
      Width           =   6660
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   4740
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   8361
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
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   315
            Index           =   3
            Left            =   2400
            Top             =   1470
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "％"
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   855
            Left            =   1980
            Top             =   2640
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   1508
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
            Begin VB.OptionButton Op_Naiyo 
               BackColor       =   &H00C0C0C0&
               Caption         =   "前回内示　　 ←→ 当月内示"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   1
               Left            =   120
               TabIndex        =   4
               Top             =   480
               Width           =   2760
            End
            Begin VB.OptionButton Op_Naiyo 
               BackColor       =   &H00C0FFC0&
               Caption         =   "前回翌月内示 ←→ 当月内示"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   0
               Left            =   120
               TabIndex        =   3
               Top             =   120
               Width           =   2760
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   315
            Index           =   2
            Left            =   570
            Top             =   1470
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "許容率"
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
         Begin CisText_V60.CisText iH1_Kyoyo 
            Height          =   375
            Left            =   1980
            TabIndex        =   1
            Top             =   1440
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   661
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
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   315
            Index           =   0
            Left            =   570
            Top             =   2055
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
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
            Left            =   1980
            TabIndex        =   0
            Top             =   900
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
            Caption         =   " 数量チェックリスト 出力先 "
            Height          =   600
            Left            =   75
            TabIndex        =   14
            Top             =   3540
            Width           =   6345
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
                  Caption         =   "内示数量チェックを行います"
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
            Top             =   4170
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
            Top             =   4200
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
            Height          =   345
            Index           =   1
            Left            =   570
            Top             =   900
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
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
            cAlingnment     =   1
            cBoderStyle     =   2
         End
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   1980
            TabIndex        =   2
            Top             =   2025
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
            Left            =   3345
            Top             =   2025
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
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D PB_Torcd 
            Height          =   360
            Left            =   3000
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
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   315
            Index           =   1
            Left            =   570
            Top             =   2910
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "工  区  "
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
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 内示数量チェック 】"
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
      Top             =   -105
      Width           =   345
   End
   Begin Cis3D_v60.CIS3D PB_Ent 
      Height          =   585
      Left            =   8280
      Top             =   7770
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
      Top             =   7770
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
      TabIndex        =   16
      Top             =   8445
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
      Top             =   8445
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
      Top             =   8445
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
Attribute VB_Name = "CNJ1110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  内示数量チェック
'**       フォームID    :  CNJ1110
'**       処理概要      :
'**
'**       作  成  日    :  2009/10/14
'**       変  更  日    :
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
    
    Dim mMaeDate        As String       ' 2009.10.14 add : １ヶ月前の年月を格納
    Dim mErrIchi        As Integer      ' 2009.10.14 add : 入力エラー位置：0=無／1=対象年月／2=許容率／3=契約先
    
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
    
    Call Op_Naiyo_Click(0)
    
    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
' 対象年月初期表示値を求める
    iH1_YM.Text = Format(Now(), "YYYYMM")

'-----------------------------------------------------------------------' 2009.10.15 add start
' 名称マスタから許容率を読み込む
    gSL_Select = "SELECT 数字1 FROM 名称マスタ" & _
                 " WHERE 区分名称= '内示数量チェック許容率' AND 値 = 1"
    If SYMRead(gSL_Select, 0) Then
       iH1_Kyoyo.Text = Format(SYM.数字1, "#0")
    Else
       Call ItemsClearSYM       '------ 名称マスタになければ新しく作成する (50%)
       With SYMCreate
            .区分名称 = "内示数量チェック許容率"
            .値 = 1
            .値名称 = "内示数量チェック許容率"
            .区分桁数 = 1
            .英数字1 = "許容率：％"
            .数字1 = 50
            iH1_Kyoyo.Text = Format(.数字1, "#0")
       End With
       Call SYMInitCreate
'       Call SYMInsert
    End If
'-----------------------------------------------------------------------' 2009.10.15 add end

'CSVﾌｧｲﾙ出力先
    mCsvName = "内示数量チェックリスト"
    mCsvKey = "CNJ1110"
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
        .Gradation 60, 190, 200, 4, 2, 3
    End With
End Sub
'-------------------------------------------' 2009.10.14 add start
'+----------------------------------------+
'+  Op_Naiyo    内容選択
'+----------------------------------------+
Private Sub Op_Naiyo_Click(Index As Integer)
    Op_Naiyo(0).Tag = Index
    Op_Naiyo(Index).Value = True
    Op_Naiyo(Index).BackColor = mSelColor
    Op_Naiyo(1 - Index).BackColor = mUnSelColor
End Sub
'-------------------------------------------' 2009.10.14 add end

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
    
    If Not Head1Chk Then GoTo ReturnPress_Ee
        
    Call DispChange("B1")
    
    If Not PrintStart Then GoTo ReturnPress_Ee

    gStr = "【  出  力  終  了  】"
    lb_msg.Caption = gStr
    DoEvents
    
'-----------------------------------' 2009.10.14 add start
    gSL_Select = "SELECT * FROM 名称マスタ " & _
                 " WHERE 区分名称 = '内示数量チェック許容率'" & _
                 "   AND 値       = '1'"
    On Error Resume Next
    Call SYMRead(gSL_Select)
    SYM.数字1 = CisFun.Val2(Trim(iH1_Kyoyo.Text))
    Call SYMUpdate  ' Call SYMRead(gSL_Select)
    On Error GoTo 0
'-----------------------------------' 2009.10.14 add end
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
    With CisFrm
        .MSGTYPE = FSize_Changes
        .MSGText = gStr
    End With
    ProcHB = "B2"
    Call DispChange(ProcHB)
    Exit Sub
ReturnPress_Ee:
    lb_msg = mSaveMsg
    DoEvents
    ProcHB = "H1"
    Call DispChange(ProcHB)
'-----------------------------------' 2009.10.14 add start
    Select Case mErrIchi
        Case 1:
            Call iH1_YM.SetFocus
        Case 2:
            Call iH1_Kyoyo.SetFocus
        Case 3:
            Call iH1_Torcd.SetFocus
    End Select
'-----------------------------------' 2009.10.14 add end
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Dim wSl_Where   As String
    Dim wKyoyo      As Currency
    
    
    Head1Chk = False
    mErrIchi = 0        ' エラー位置：無し
    
    If Trim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   対象年月を入力して下さい    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        mErrIchi = 1    ' iH1_YM.SetFocus
        Exit Function
    End If
'-------------------------------------------' 2009.10.14 add start
    mMaeDate = DateAdd("m", -1, Mid(iH1_YM.Text, 1, 4) & "/" & Mid(iH1_YM, 5, 2) & "/01")
    mMaeDate = Mid(mMaeDate, 1, 4) & Mid(mMaeDate, 6, 2)    ' 比較する月を求める

 Dim strWk As Integer
    strWk = CisFun.Val2(iH1_Kyoyo.Text)
    If strWk = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   許容率を入力して下さい    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        mErrIchi = 2    'iH1_Kyoyo.SetFocus
        Exit Function
    End If
'-------------------------------------------' 2009.10.14 add end

    If Trim(iH1_Torcd) <> "" Then
        If Not TorNmGet(iH1_Torcd, 2) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "   　契約先　未登録    "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            mErrIchi = 3    'iH1_Torcd.SetFocus
            Exit Function
        End If
    End If

    lb_msg.Caption = "【 内示数量チェック 抽出中 】"
    DoEvents
    
    wKyoyo = CisFun.Val2(iH1_Kyoyo) / 100
    
   'gSL_Select = "SELECT N1.*"
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT N1.品番, N1.納入先, N1.受入, N1.契約先, N1.内示数"
    gSL_Select = gSL_Select & "      ,N2.内示数 翌々月内示数, N2.翌月内示数 翌月内示数, HT.背番号 背番号"
    gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(HM.表示品番, '') <> ''"
    gSL_Select = gSL_Select & "            THEN HM.表示品番"
    gSL_Select = gSL_Select & "            ELSE N1.品番"
    gSL_Select = gSL_Select & "            END  表示品番"
    gSL_Select = gSL_Select & "      ,T1.略称 契約先名, T2.略称 納入先名"
    gSL_Select = gSL_Select & "      ,cast((cast(N1.内示数 as decimal(9,2))/cast("
    If Op_Naiyo(0).Value = True Then
       gSL_Select = gSL_Select & "N2.翌月内示数"
    Else
       gSL_Select = gSL_Select & "N2.内示数"
    End If
    gSL_Select = gSL_Select & " as decimal(9,2))) as decimal(9,2)) 内示比率"
    gSL_Select = gSL_Select & "  FROM 内示テーブル N1"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 内示テーブル N2"
    gSL_Select = gSL_Select & "            ON  N2.内示年月 = '" & mMaeDate & "'"
    gSL_Select = gSL_Select & "            AND N2.納入先   = N1.納入先"
    gSL_Select = gSL_Select & "            AND N2.品番     = N1.品番"
    gSL_Select = gSL_Select & "            AND N2.受入     = N1.受入"
    gSL_Select = gSL_Select & "            AND N2.内示区分 = N1.内示区分"
    gSL_Select = gSL_Select & "            AND N2.契約先   = N1.契約先"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 品番取引先マスタ HT"
    gSL_Select = gSL_Select & "            ON  HT.品番     = N1.品番"
    gSL_Select = gSL_Select & "            AND HT.取引先   = N1.納入先"
    gSL_Select = gSL_Select & "            AND HT.受入     = N1.受入"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 品番マスタ HM"
    gSL_Select = gSL_Select & "            ON  HM.品番     = N1.品番"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ T1"
    gSL_Select = gSL_Select & "            ON  T1.取引先CD     = N1.契約先"
    gSL_Select = gSL_Select & "            AND T1.取引先区分 = 0"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ T2"
    gSL_Select = gSL_Select & "            ON  T2.取引先CD     = N1.納入先"
    gSL_Select = gSL_Select & "            AND T2.取引先区分 = 0"
    gSL_Select = gSL_Select & " WHERE N1.内示年月   =  '" & iH1_YM.Text & "'"
    If Trim(iH1_Torcd.Text) <> "" Then
       gSL_Select = gSL_Select & "   AND N1.契約先     =  '" & Trim(iH1_Torcd.Text) & "'"
    End If
    gSL_Select = gSL_Select & "   AND N1.内示数     <> 0"
    If Op_Naiyo(0).Value = True Then
       gSL_Select = gSL_Select & "   AND N2.翌月内示数 <> 0"
      'gSL_Select = gSL_Select & "   AND(N1.内示数 / N2.翌月内示数)"
       gSL_Select = gSL_Select & "   AND ( cast((cast(N1.内示数 as decimal(9,2)) / " & _
                                             " cast(N2.翌月内示数 as decimal(9,2))) " & _
                                              "as decimal(9,2))"
       gSL_Select = gSL_Select & "    > " & Format(1 + wKyoyo, "0.00")
       gSL_Select = gSL_Select & " OR "
       gSL_Select = gSL_Select & "         cast((cast(N1.内示数 as decimal(9,2)) / " & _
                                             " cast(N2.翌月内示数 as decimal(9,2))) " & _
                                              "as decimal(9,2))"
       gSL_Select = gSL_Select & "    < " & Format(1 - wKyoyo, "0.00") & " )"
    Else
       gSL_Select = gSL_Select & "   AND N2.内示数     <> 0"
      'gSL_Select = gSL_Select & "   AND(N1.内示数 / N2.内示数)"
       gSL_Select = gSL_Select & "   AND ( cast((cast(N1.内示数 as decimal(9,2)) / " & _
                                             " cast(N2.内示数 as decimal(9,2))) " & _
                                              "as decimal(9,2))"
'       gSL_Select = gSL_Select & "    < 0." & Format(CisFun.Val2(iH1_Kyoyo.Text), "00")
       gSL_Select = gSL_Select & "    > " & Format(1 + wKyoyo, "0.00")
       gSL_Select = gSL_Select & " OR "
       gSL_Select = gSL_Select & "         cast((cast(N1.内示数 as decimal(9,2)) / " & _
                                             " cast(N2.内示数 as decimal(9,2))) " & _
                                              "as decimal(9,2))"
'       gSL_Select = gSL_Select & "    > 1." & Format(CisFun.Val2(iH1_Kyoyo.Text), "00") & " )"
       gSL_Select = gSL_Select & "    < " & Format(1 - wKyoyo, "0.00") & " )"
    End If
'''    gSL_Select = gSL_Select & " > 0." & Format(CisFun.Val2(iH1_Kyoyo.Text), "00")
    
    
    ' 今回または前回のいずれかがゼロの場合のデータを結合する
    gSL_Select = gSL_Select & " UNION "
    gSL_Select = gSL_Select & "SELECT N1.品番, N1.納入先, N1.受入, N1.契約先, N1.内示数"
    gSL_Select = gSL_Select & "      ,N2.内示数 翌々月内示数, N2.翌月内示数 翌月内示数, HT.背番号 背番号"
    gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(HM.表示品番, '') <> ''"
    gSL_Select = gSL_Select & "            THEN HM.表示品番"
    gSL_Select = gSL_Select & "            ELSE N1.品番"
    gSL_Select = gSL_Select & "            END  表示品番"
    gSL_Select = gSL_Select & "      ,T1.略称 契約先名, T2.略称 納入先名"
    gSL_Select = gSL_Select & "      ,0  内示比率"
    gSL_Select = gSL_Select & "  FROM 内示テーブル N1"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 内示テーブル N2"
    gSL_Select = gSL_Select & "            ON  N2.内示年月 = '" & mMaeDate & "'"
    gSL_Select = gSL_Select & "            AND N2.納入先   = N1.納入先"
    gSL_Select = gSL_Select & "            AND N2.品番     = N1.品番"
    gSL_Select = gSL_Select & "            AND N2.受入     = N1.受入"
    gSL_Select = gSL_Select & "            AND N2.内示区分 = N1.内示区分"
    gSL_Select = gSL_Select & "            AND N2.契約先   = N1.契約先"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 品番取引先マスタ HT"
    gSL_Select = gSL_Select & "            ON  HT.品番     = N1.品番"
    gSL_Select = gSL_Select & "            AND HT.取引先   = N1.納入先"
    gSL_Select = gSL_Select & "            AND HT.受入     = N1.受入"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 品番マスタ HM"
    gSL_Select = gSL_Select & "            ON  HM.品番     = N1.品番"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ T1"
    gSL_Select = gSL_Select & "            ON  T1.取引先CD     = N1.契約先"
    gSL_Select = gSL_Select & "            AND T1.取引先区分 = 0"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ T2"
    gSL_Select = gSL_Select & "            ON  T2.取引先CD     = N1.納入先"
    gSL_Select = gSL_Select & "            AND T2.取引先区分 = 0"
    gSL_Select = gSL_Select & " WHERE N1.内示年月   =  '" & iH1_YM.Text & "'"
    If Trim(iH1_Torcd.Text) <> "" Then
       gSL_Select = gSL_Select & "   AND N1.契約先     =  '" & Trim(iH1_Torcd.Text) & "'"
    End If
    If Op_Naiyo(0).Value = True Then
       gSL_Select = gSL_Select & "   AND ((N1.内示数 =  0 AND N2.翌月内示数 <> 0) OR (N1.内示数 <> 0 AND N2.翌月内示数 = 0))"
    Else
       gSL_Select = gSL_Select & "   AND ((N1.内示数 =  0 AND N2.内示数     <> 0) OR (N1.内示数 <> 0 AND N2.内示数     = 0))"
    End If
    '
    gSL_Select = gSL_Select & " ORDER BY N1.契約先, N1.納入先, N1.品番 "
        
    If Not NJTRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　該当データがありません       "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        mErrIchi = 1    ' iH1_YM.SetFocus
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
    
    ' 印刷処理に渡す情報をセット（他モジュール用の変数を流用）
    RV_YM = iH1_YM.Text                                         ' 年月
    If Op_Naiyo(0).Value = True Then
       RV_TorKb = 0                                             ' 内容区分
       Rv_LongStr = "前回翌月内示 ←→ 当月内示"                ' チェック内容
    Else
       RV_TorKb = 1
       Rv_LongStr = "前回内示 ←→ 当月内示"
    End If
    RV_Code = Format(CisFun.Val2(Trim(iH1_Kyoyo.Text)), "#0")   ' 許容率
    
    If Op_Out(0).Value Or Op_Out(1).Value Then
        lb_msg.Caption = "【 内示数量チェックリスト 印刷中 】"
        DoEvents
        
        With CisPrt
             Set .iPrtForm = CNJ1110L
             .dpMSG = "内示数量チェックリスト　印刷中"
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

    Dim OutStr      As String
    Dim OutFNo      As Long
    Dim OutCnt      As Long
    
    Dim Moji_F      As String
    Dim Moji_B      As String
    Moji_F = "=" & Chr(34): Moji_B = Chr(34) & ","
    
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
    
    lb_msg.Caption = "【 内示数量チェックリスト 出力中 】"
    Csv_Bar.Visible = True
    
    DoEvents
        
    OutFNo = FreeFile
    Open mCsvDir For Output As #OutFNo

    Call CisFrm.MousePT(11)
    
    With CisFrm
        Set .MeterPic = Csv_Bar
        .MeterMax = CisDB.RecordCount
    End With
        
    ' 表題
    OutStr = Mid(RV_YM, 1, 4) & "月" & Mid(RV_YM, 5, 2) & "月度　内示数量チェックリスト"
    OutStr = OutStr & ",,,,,,,,発行日：" & Format((Now), "yyyy/mm/dd") & vbCrLf
    OutStr = OutStr & "チェック内容：" & Rv_LongStr & ",,,,,,,,許容率：" & RV_Code & "%"
    Print #OutFNo, OutStr
    
    ' 見出し
    OutStr = "契約先,契約先名,納入先,納入先名,受入,背番号,品番,前回数量,今回数量,数量差,誤差率"
    Print #OutFNo, OutStr
    
    ' 内容
    OutCnt = 0
    Do Until Not NJT_RDSTS
        With NJT
            OutCnt = OutCnt + 1
            OutStr = ""
            OutStr = OutStr & Moji_F & .契約先 & Moji_B
            OutStr = OutStr & .契約先名 & ","
            OutStr = OutStr & Moji_F & .納入先 & Moji_B
            OutStr = OutStr & .納入先名 & ","
            OutStr = OutStr & Moji_F & .受入 & Moji_B
            OutStr = OutStr & Moji_F & .背番号 & Moji_B
            OutStr = OutStr & Moji_F & .表示品番 & Moji_B
            If Op_Naiyo(0).Value = True Then
               OutStr = OutStr & .翌月内示数 & "," & _
                                 .内示数 & "," & _
                                 .内示数 - .翌月内示数 & ","
            Else
               OutStr = OutStr & .翌々月内示数 & "," & _
                                 .内示数 & "," & _
                                 .内示数 - .翌々月内示数 & ","
            End If
            OutStr = OutStr & Moji_F & Format(.内示比率, "##,####0.00") & Chr(34)
            
            Print #OutFNo, OutStr
            CisFrm.MeterUpdate OutCnt
            Call NJTReadNext
        End With
    Loop

CsvPut_Ed:
    Call NJTClose

    Close #OutFNo
    SaveSetting appname:=gRegKey, section:="CSV", _
                    Key:=mCsvKey, setting:=RTrim(lb_PrtDev)
    
    Call CisFrm.MousePT(1)
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "   出力処理が終了しました。    "
        .MB_MSG(4) = "   　 " & "出力先    = " & mCsvDir
        .MB_MSG(5) = "   　 " & "出力件数 = " & StrConv(Format(CisDB.RecordCount, "##,##0"), vbWide) & " 件 "
        .MB_Title = "出力確認"
        .MB_Button = OK
        .MBOX
    End With
    CsvPut = True
    
End Function

