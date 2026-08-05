VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form COK0060 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "受給検収入力[一括]"
   ClientHeight    =   10980
   ClientLeft      =   1710
   ClientTop       =   1815
   ClientWidth     =   15210
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
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
   ScaleHeight     =   10980
   ScaleWidth      =   15210
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
      Left            =   120
      ScaleHeight     =   870
      ScaleWidth      =   10905
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1050
      Width           =   10965
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Index           =   2
         Left            =   5670
         Top             =   60
         Width           =   585
         _ExtentX        =   1032
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
            Left            =   90
            TabIndex        =   11
            Top             =   330
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
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   780
         Left            =   2145
         Top             =   60
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   1376
         BackColor       =   16761087
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   30
            TabIndex        =   10
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
         End
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1020
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
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   780
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1376
         BackColor       =   16761087
         ForeColor       =   16711680
         Caption         =   "検 収 日"
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_Ymd 
            Height          =   345
            Left            =   60
            TabIndex        =   9
            Top             =   330
            Width           =   1950
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
            Object.Width           =   1950
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
      End
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   780
         Left            =   6255
         Top             =   60
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "品　　　番"
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
            Left            =   30
            TabIndex        =   12
            Top             =   330
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   635
            cWildeCard      =   -1  'True
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
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   780
         Left            =   9465
         Top             =   60
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "納入番号"
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
         Begin CisText_V60.CisText iH1_NouNo 
            Height          =   360
            Left            =   105
            TabIndex        =   13
            Top             =   330
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            cWildeCard      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXXXXX"
            MaxLength       =   9
         End
      End
   End
   Begin VB.PictureBox DspKensu_Back 
      BackColor       =   &H00000000&
      Height          =   405
      Left            =   12690
      ScaleHeight     =   345
      ScaleWidth      =   2325
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3690
      Width           =   2385
      Begin Cis3D_v60.CIS3D lbKensu 
         Height          =   345
         Left            =   15
         Top             =   0
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   609
         Caption         =   " 表示件数"
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
         cAlingnment     =   1
         Begin Cis3D_v60.CIS3D Dsp_Kensu 
            Height          =   270
            Left            =   1290
            Top             =   30
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   476
            BackColor       =   16777152
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -60
         End
      End
   End
   Begin VB.Timer Timer4 
      Left            =   210
      Top             =   810
   End
   Begin VB.PictureBox B2_Area1 
      Height          =   6135
      Left            =   105
      ScaleHeight     =   6075
      ScaleWidth      =   14940
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4200
      Width           =   15000
      Begin Cis3D_v60.CIS3D BLB_DD 
         Height          =   5940
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   14835
         _ExtentX        =   26167
         _ExtentY        =   10478
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   5850
            Left            =   15
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   30
            Width           =   14775
            _Version        =   196608
            _ExtentX        =   26061
            _ExtentY        =   10319
            _StockProps     =   228
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "№ |仕入先              |受|検収日     |納入番号  |背番    |品　番　　　　　　　　　 |赤/黒|収容数  |枚　数 |端　数 |入荷数 　||"
            Rows            =   16
            Cols            =   20
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   360
            SelectionMode   =   1
         End
         Begin CisText_V60.CisText B_TabDummy 
            Height          =   360
            Left            =   135
            TabIndex        =   22
            Top             =   105
            Width           =   240
            _ExtentX        =   423
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
            Text            =   "X"
            MaxLength       =   1
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
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
      Height          =   1920
      Left            =   120
      ScaleHeight     =   1860
      ScaleWidth      =   10155
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2145
      Width           =   10215
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   885
         Left            =   5175
         Top             =   45
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "検 収 日"
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
         Begin CisYMD_With_Btn.CisYMDwB iB1_Ymd 
            Height          =   345
            Left            =   405
            TabIndex        =   2
            Top             =   390
            Width           =   1905
            _ExtentX        =   2831
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
            Object.Width           =   1905
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   885
         Left            =   8085
         Top             =   45
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "納入番号"
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
         Begin CisText_V60.CisText iB1_NouNo 
            Height          =   360
            Left            =   330
            TabIndex        =   3
            Top             =   390
            Width           =   1230
            _ExtentX        =   2170
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
            Text            =   "XXXXXXXXX"
            MaxLength       =   9
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   885
         Left            =   8715
         Top             =   945
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "数 量"
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
         Begin CisText_V60.CisText iB1_Suryo 
            Height          =   360
            Left            =   75
            TabIndex        =   8
            Top             =   390
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   635
            cFormat         =   "#,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "#,###,###"
            MaxLength       =   9
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "#######"
            cILength        =   7
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   885
         Index           =   0
         Left            =   4725
         Top             =   945
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "赤黒"
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
         Begin VB.CheckBox B1Chk_Aka 
            BackColor       =   &H00C0C0C0&
            Caption         =   "黒"
            Height          =   345
            Left            =   75
            TabIndex        =   6
            Top             =   390
            Width           =   555
         End
      End
      Begin Cis3D_v60.CIS3D BLB_Zaist 
         Height          =   885
         Left            =   1320
         Top             =   945
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "品　番"
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
         cPositionX      =   90
         cPositionY      =   30
         Begin CisText_V60.CisText iB1_Hinbn 
            Height          =   360
            Left            =   105
            TabIndex        =   5
            Top             =   390
            Width           =   3210
            _ExtentX        =   5662
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   885
         Index           =   0
         Left            =   60
         Top             =   945
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1561
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iB1_Seban 
            Height          =   360
            Left            =   105
            TabIndex        =   4
            Top             =   390
            Width           =   1080
            _ExtentX        =   1905
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
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   885
         Left            =   4455
         Top             =   45
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1561
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
         Begin CisText_V60.CisText iB1_Ukeir 
            Height          =   360
            Left            =   165
            TabIndex        =   1
            Top             =   375
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
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   885
         Left            =   60
         Top             =   45
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "仕 入 先"
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
         Begin Cis3D_v60.CIS3D B1lb_Tornm 
            Height          =   360
            Left            =   1215
            Top             =   360
            Width           =   2880
            _ExtentX        =   5080
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin CisText_V60.CisText iB1_Torcd 
            Height          =   360
            Left            =   225
            TabIndex        =   0
            Top             =   360
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
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   885
         Index           =   3
         Left            =   5445
         Top             =   945
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1561
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Syuyo 
            Height          =   360
            Left            =   75
            Top             =   390
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   635
            BackColor       =   16777152
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
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   885
         Index           =   4
         Left            =   6495
         Top             =   945
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "枚 数"
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
         Begin CisText_V60.CisText iB1_Maisu 
            Height          =   375
            Left            =   75
            TabIndex        =   7
            Top             =   390
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   661
            cFormat         =   "###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###,###"
            MaxLength       =   7
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "######"
            cILength        =   6
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   885
         Index           =   5
         Left            =   7590
         Top             =   945
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "端 数"
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
         Begin Cis3D_v60.CIS3D B1lb_Hasu 
            Height          =   360
            Left            =   90
            Top             =   390
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   635
            BackColor       =   16777152
            ForeColor       =   0
            Caption         =   "###,###"
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
      End
   End
   Begin VB.Timer Timer3 
      Left            =   1980
      Top             =   810
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
      Height          =   1365
      Left            =   11565
      ScaleHeight     =   1305
      ScaleWidth      =   3510
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   615
      Width           =   3570
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   915
         Top             =   660
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Left            =   1785
         Top             =   660
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Left            =   2640
         Top             =   660
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Left            =   1770
         Top             =   30
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   900
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   45
         Top             =   660
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_MNT 
         Height          =   585
         Left            =   45
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   8
         BtnCaption      =   "修正"
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
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   990
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   540
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   930
      Top             =   660
   End
   Begin VB.Timer Timer2 
      Left            =   1440
      Top             =   720
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【　受給検収入力[一括]  】"
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
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'なし
         Height          =   345
         Index           =   1
         Left            =   10620
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   30
         Width           =   3675
         Begin Cis3D_v60.CIS3D CIS3D15 
            Height          =   345
            Index           =   1
            Left            =   2460
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   8421631
            ForeColor       =   0
            Caption         =   "削除"
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
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Delsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   8421631
               ForeColor       =   0
               Caption         =   "ZZ9"
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
         Begin Cis3D_v60.CIS3D UC_3D6 
            Height          =   345
            Left            =   0
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   8454143
            ForeColor       =   0
            Caption         =   "追加"
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
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Addsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   8454143
               ForeColor       =   0
               Caption         =   "ZZ9"
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
         Begin Cis3D_v60.CIS3D UC_3D7 
            Height          =   345
            Left            =   1230
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   16777088
            ForeColor       =   0
            Caption         =   "修正"
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
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Mntsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   16777088
               ForeColor       =   0
               Caption         =   "ZZ9"
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14400
         Top             =   90
         Width           =   885
         _ExtentX        =   1561
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
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   8070
         Top             =   60
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
End
Attribute VB_Name = "COK0060"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************'
'**
'**       フォーム名    :   受給検収入力
'**       フォームID    :   COK0060
'**       処理概要      :
'**
'**       作  成  日    :   2006/09/12  By CIS SIMPR-A
'**       変  更  日    :   2008/01/18  ｷﾉｯｸｽよりSIMPR-Aへ移行
'**       変  更  日    :   2008/04/11  背番号8桁
'**
'********************************************************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    Dim CisVsGrid1      As New CisVsGrid3

    Dim SaveRow         As Integer
    Dim RowColFlg       As Boolean
    Dim AddRow          As Integer
    Dim SaveKanriNo     As Long
    Dim AddSyoriKB      As String

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
    Dim Save_Torcd      As String
    Dim Save_Ukeir      As String
    Dim Save_YMD        As String
    Dim Save_NouNo      As String
    Dim Save_Seban      As String
    Dim Save_Hinbn      As String
    Dim Save_Aka        As Byte
    
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
'
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
'
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
    CisFrm.TimeSet_TitleName = gSysName & "[ " & App.EXEName & " ]"
    CisFrm.TimeSet_TitleSpace = 130 - CisFun.Len2(gSysName) - CisFun.Len2(App.EXEName)
'
    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔(基本)
    Timer2.Enabled = False   ' 処理区分点滅禁止(基本)
'
    '+--------------------------------+
    '+  グリッド初期設定
    '+--------------------------------+
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 360
        .PatanMax = 1
        .SelectCol = 18
        .InitGet = False
        .Init
    End With
    Call GridInit1
    
    AddRow = 0
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    SyoriKB = "ADD"
    ProcHB = "B1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    Call SyoriName(SyoriKB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
'
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    FormAct = Not FormAct
'
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
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    '【新規】
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_Can_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
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
Private Sub PB_Can_Click()
    Call CisFrm.MousePT(11)
   
    If ProcHB = "H1" Then
        Call HeadBodyClear("H1")
        ProcHB = "B1"
        SyoriKB = "ADD"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        GoTo PB_CAN_Ed
    End If
            
    If ProcHB = "B2" Then
        Call CisVsGrid1.vsColor(消去個別)
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        Call HeadBodyClear("B")
        GoTo PB_CAN_Ed
    End If

    If ProcHB = "B1" Then
        Call CisVsGrid1.vsColor(消去個別)
        ProcHB = "B2"
        Call DispChange(ProcHB)
        GoTo PB_CAN_Ed
    End If

PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新　規  キ ー(F1)        +
'+----------------------------+
Private Sub PB_ADD_Click()
    
    If SyoriKB <> "ADD" Then
        If ProcHB = "H1" Then
            AddRow = 0
            Call GridInit1
            ProcHB = "B1"
            SyoriKB = "ADD"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
            Call HeadBodyClear("B")
            RowColFlg = True
            Call CisVsGrid1.vsColor(消去個別)
            RowColFlg = False
            iB1_Torcd.SetFocus
        End If
    End If

End Sub
'+----------------------------+
'+    修  正  キ ー(F2)        +
'+----------------------------+
Private Sub PB_MNT_Click()
    
    If SyoriKB <> "MNT" Then
        If ProcHB = "B1" Then
            ProcHB = "H1"
            SyoriKB = "MNT"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
            RowColFlg = True
            Call HeadBodyClear("B")
            Call HeadBodyClear("H")
            Call CisVsGrid1.vsColor(消去個別)
            RowColFlg = False
        Else
            SyoriKB = "MNT"
            Call SyoriName(SyoriKB)
        End If
            
    End If

End Sub
'+----------------------------+
'+    削  除  キ ー(F3)        +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        If ProcHB = "B1" Then
            ProcHB = "H1"
            SyoriKB = "DEL"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
            RowColFlg = True
            Call HeadBodyClear("B")
            Call HeadBodyClear("H")
            Call CisVsGrid1.vsColor(消去個別)
            RowColFlg = False
        Else
            SyoriKB = "DEL"
            Call SyoriName(SyoriKB)
        End If
        
    End If

End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd(0): Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_Ukeir(0): Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn(0): Exit Sub
'
    If PB_Look.Tag = "iB1_Ymd" Then
        iB1_Ymd.ShowCalender
        Exit Sub
    End If
    If PB_Look.Tag = "iH1_Ymd" Then
        iH1_Ymd.ShowCalender
        Exit Sub
    End If
    If PB_Look.Tag = "iB1_Torcd" Then Call Look_Torcd(1): Exit Sub
    If PB_Look.Tag = "iB1_Ukeir" Then Call Look_Ukeir(1): Exit Sub
    If PB_Look.Tag = "iB1_Hinbn" Then Call Look_Hinbn(1): Exit Sub
    If PB_Look.Tag = "iB1_Seban" Then Call Look_Hinbn(1): Exit Sub

End Sub
'+----------------------------+
'+       仕入先検索           +
'+----------------------------+
Private Sub Look_Torcd(KB As Byte)

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 1
    RV_TorKb1 = 0
    
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
        If KB = 0 Then
           iH1_Torcd = RTrim(RV_TorcdK)
           Call iH1_Torcd_LostFocus
        Else
           iB1_Torcd = RTrim(RV_TorcdK)
           Call iB1_Torcd_LostFocus
        End If
    End If
End Sub
'+----------------------------+
'+       品番検索              +
'+----------------------------+
Private Sub Look_Hinbn(KB As Byte)

    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 1
    RVI_SK = 9
    If KB = 0 Then
        RVI_Torcd = RTrim(iH1_Torcd)
        RVI_Ukeir = RTrim(iH1_Ukeir)
    Else
        RVI_Torcd = RTrim(iB1_Torcd)
        RVI_Ukeir = RTrim(iB1_Ukeir)
    End If
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        If KB = 0 Then
            iH1_Hinbn = RTrim(RV_Hinbn)
        Else
            iB1_Hinbn = RTrim(RV_Hinbn)
            iB1_Seban = RTrim(RV_Seban)
        End If
    End If
End Sub
'+----------------------------+
'+     納入先受入検索          +
'+----------------------------+
Private Sub Look_Ukeir(KB As Byte)

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 8
    If RTrim(PB_Look.Tag) = "iH1_Ukeir" Then
        RV_Torcd = RTrim(iH1_Torcd)
    Else
        RV_Torcd = RTrim(iB1_Torcd)
    End If
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        If KB = 0 Then
            iH1_Torcd = RTrim(RV_TorcdK)
            Call iH1_Torcd_LostFocus
            iH1_Ukeir = RTrim(RV_Ukeir)
        Else
            iB1_Torcd = RTrim(RV_TorcdK)
            Call iB1_Torcd_LostFocus
            iB1_Ukeir = RTrim(RV_Ukeir)
        End If
    End If
End Sub
'+---( H1 )----------------------------------------------------------------------------
'+-----------------------------+
'+  iH1_Ymd
'+-----------------------------+
Private Sub iH1_Ymd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ymd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
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
    If TorNmGet(iH1_Torcd, 8) Then
        H1lb_Tornm = TRM.略称
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
'+---( B1 )----------------------------------------------------------------------------
'+-----------------------------+
'+  iB1_Torcd
'+-----------------------------+
Private Sub iB1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_Tornm = ""
    If TorNmGet(iB1_Torcd, 8) Then
        B1lb_Tornm = TRM.略称
    End If
End Sub
'+-----------------------------+
'+  iB1_Ukeir
'+-----------------------------+
Private Sub iB1_Ukeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Ymd
'+-----------------------------+
Private Sub iB1_Ymd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Ymd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Seban
'+-----------------------------+
Private Sub iB1_Seban_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Seban_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    Call Seban_Set
End Sub
'+-----------------------------+
'+  iB1_Hinbn
'+-----------------------------+
Private Sub iB1_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    Call Hinbn_Set
End Sub
'+-----------------------------+
'+  B1Chk_Aka
'+-----------------------------+
Private Sub B1Chk_Aka_Click()
    If B1Chk_Aka.Value = 0 Then
        B1Chk_Aka.Caption = "黒"
    Else
        B1Chk_Aka.Caption = "赤"
    End If
    
    If B1Chk_Aka.Value = 0 Then
        If CisFun.Val2(iB1_Suryo) < 0 Then
           iB1_Suryo = Format(CisFun.Val2(iB1_Suryo) * -1, "###,###")
        End If
    
        If CisFun.Val2(iB1_Maisu) < 0 Then
           iB1_Maisu = Format(CisFun.Val2(iB1_Maisu) * -1, "###,###")
        End If
    End If

End Sub
'+-----------------------------+
'+  iB1_Maisu
'+-----------------------------+
Private Sub iB1_Maisu_LostFocus()
    If CisFun.Val2(iB1_Maisu) <> 0 Then
        If CisFun.Val2(B1lb_Syuyo) <> 0 Then
            gLong = (CisFun.Val2(B1lb_Syuyo) * CisFun.Val2(iB1_Maisu))
            iB1_Suryo = Format(gLong, "######")
        End If
    Else
        iB1_Suryo = ""
    End If
    B1lb_Hasu = ""
End Sub
'+-----------------------------+
'+  iB1_Suryo
'+-----------------------------+
Private Sub iB1_Suryo_LostFocus()
    Call Suryo_Set
End Sub
'*------------------------------------------*'
'*    背番号による品番・収容数の表示          *
'*------------------------------------------*'
Private Sub Seban_Set()
    If RTrim(iB1_Seban) = "" Then Exit Sub
    Call ItemsClearHTM
    gSL_Select = "SELECT * FROM 品番取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 背番号 = '" & RTrim(iB1_Seban) & "'"
    gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(iB1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(iB1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   AND 品目 = 1"
    Call HTMRead(gSL_Select, 0)
    Call HTMClose
    If CisDB.RecordCount = 0 Then Exit Sub
    If CisDB.RecordCount > 1 Then Exit Sub
    iB1_Hinbn = RTrim(HTM.品番)
'収容数表示
    Call Syuyo_Set
End Sub
'*------------------------------------------*'
'*    品番による背番号・収容数の表示          *
'*------------------------------------------*'
Private Sub Hinbn_Set()
    Call ItemsClearHTM
    gSL_Select = "SELECT * FROM 品番取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iB1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(iB1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(iB1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   AND 品目 = 1"
    If Not HTMRead(gSL_Select, 1) Then
        B1lb_Syuyo = ""
        Exit Sub
    End If
    iB1_Seban = RTrim(HTM.背番号)
    iB1_Hinbn = RTrim(HTM.品番)
'収容数表示
    Call Syuyo_Set
End Sub
'*------------------------------------------*'
'*    収容数の表示                           *
'*------------------------------------------*'
Private Sub Syuyo_Set()
    Dim wNYmd   As String
    Dim wNBin   As String
    
    wNYmd = RTrim(iH1_Ymd)
    If RTrim(HTM.変更日) <> "" Then
        If RTrim(HTM.変更日) > RTrim(wNYmd) Then
            B1lb_Syuyo = Format(HTM.収容数1, "#,###")
        Else
            B1lb_Syuyo = Format(HTM.収容数2, "#,###")
        End If
    Else
        B1lb_Syuyo = Format(HTM.収容数1, "#,###")
    End If
    Call Suryo_Set
End Sub
'*------------------------------------------*'
'*    枚数・端数・発注数の再表示              *
'*------------------------------------------*'
Private Sub Suryo_Set()
    B1lb_Hasu = ""
    If CisFun.Val2(B1lb_Syuyo) <> 0 Then
        If RTrim(iB1_Suryo) = "" Then
            If CisFun.Val2(iB1_Maisu) <> 0 Then
                gLong = (CisFun.Val2(B1lb_Syuyo) * CisFun.Val2(iB1_Maisu))
                iB1_Suryo = Format(gLong, "######")
            End If
        Else
            gCnt1 = CisFun.Val2(iB1_Suryo) \ CisFun.Val2(B1lb_Syuyo)
            gCnt2 = CisFun.Val2(iB1_Suryo) Mod CisFun.Val2(B1lb_Syuyo)
            If gCnt2 <> 0 Then
                gCnt1 = gCnt1 + 1
            End If
            iB1_Maisu = Format(gCnt1, "#####")
            B1lb_Hasu = Format(gCnt2, "###,###")
        End If
    Else
        iB1_Maisu = ""
        B1lb_Hasu = ""
    End If
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
'+------------------------------------------------------------------+
'+      基本情報
'+------------------------------------------------------------------+
Private Sub SyoriName(SYKB As String)
    Dim w_Color     As Long
    w_Color = &H80000008 '初期色ｾｯﾄ(黒)
   
'   H2LB_Size_B.Enabled = False
   
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
           w_Color = &HFF&
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

    If ProcHB = "B2" Then
       PB_ADD.Visible = False
       PB_MNT.Visible = False
       PB_DEL.Visible = False
    End If
    
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
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
          AddRow = 0
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
       Case "B2"
          B2Mode = True
          B2Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    H1_Area1.Visible = H1Mode
    If SyoriKB <> "ADD" Then
        H1_Area1.Visible = True
        DspKensu_Back.Visible = True
    Else
        DspKensu_Back.Visible = False
    End If
    
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B2_Area1.Enabled = B2Mode
    B2_Area1.BackColor = B2Color

    PB_CAN.Visible = B2Mode
    If SyoriKB <> "ADD" And Pro = "B1" Then
       PB_CAN.Visible = True
    End If
    
    PB_ADD.Visible = Not B2Mode
    PB_MNT.Visible = Not B2Mode
    PB_DEL.Visible = Not B2Mode
    
    If SyoriKB = "MNT" And Pro = "B1" Then
        PB_ADD.Visible = False
        PB_MNT.Visible = False
        PB_DEL.Visible = False
    End If
    
    If Pro = "B2" Then
       vsGrid1.SetFocus
    End If
    
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
    
    On Error GoTo HeadBodyClear_ER
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
              '  .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
                .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    
    
    If HeadBody = "B1" Or HeadBody = "B" Or HeadBody = "ALL" Then
        B1Chk_Aka.Value = 0
        B1Chk_Aka.Caption = "黒"
    End If
    If HeadBody = "B2" Or HeadBody = "B" Or HeadBody = "ALL" Then
        CisVsGrid1.Clear
        Call CisVsGrid1.Clear
        Dsp_Kensu = ""
    End If
    Exit Sub
HeadBodyClear_ER:
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        
        ProcHB = "B2"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        If ProcHB = "B2" Then
            Call VSGrid1_Click
        End If
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut(0) Then
            GoTo ReturnPress_Ed
        End If
        If SyoriKB = "ADD" Then
            ProcHB = "B1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
            iB1_Torcd.SetFocus
        Else
            If CisFun.Val2(Dsp_Kensu) <> 0 Then
               ProcHB = "B2"
               Call DispChange(ProcHB)
               Call SyoriName(SyoriKB)
               Call VSGrid1_Click
               Call vsGrid1.SetFocus
            Else
                ProcHB = "B1"
                SyoriKB = "ADD"
                Call DispChange(ProcHB)
                Call SyoriName(SyoriKB)
            End If
        End If
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B2" Then
        If Not Body2Chk Then
           vsGrid1.SetFocus
           GoTo ReturnPress_Ed
        End If
' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then
                RowColFlg = True
                Call CisVsGrid1.vsColor(消去個別)
                RowColFlg = False
                vsGrid1.SetFocus
                GoTo ReturnPress_Ed
            End If
            Call HeadBodyClear("B1")
            If CisFun.Val2(Dsp_Kensu) <> 0 Then
                ProcHB = "B2"
                Call DispChange(ProcHB)
                Call SyoriName(SyoriKB)
                Call VSGrid1_Click
                Call vsGrid1.SetFocus
            Else
                ProcHB = "B1"
                SyoriKB = "ADD"
                Call DispChange(ProcHB)
                Call SyoriName(SyoriKB)
            End If
        Else
            ProcHB = "B1"
            Call DispChange(ProcHB)
        End If
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
'
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'　検収日
    If Trim(iH1_Ymd) = "" Then
        CisFun.MB_MSG(2) = "    【検収日】必須入力          "
        CisFun.MBOX
       iH1_Ymd.SetFocus
       Exit Function
    End If
    If Not iH1_Ymd.cChkResult Then
        CisFun.MB_MSG(2) = "    正しい検収日を入力して下さい            "
        CisFun.MBOX
       iH1_Ymd.SetFocus
       Exit Function
    End If
'　仕入先
    If Trim(iH1_Torcd) = "" Then
        CisFun.MB_MSG(2) = "    【仕入先】必須入力            "
        CisFun.MBOX
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If Not TorNmGet(iH1_Torcd, 8) Then
        CisFun.MB_MSG(2) = "     仕入先　未登録！            "
        CisFun.MB_MSG(4) = "    （取引先マスタ）            "
        CisFun.MBOX
       iH1_Torcd.SetFocus
       Exit Function
    End If
'受給検収テーブル検索
    If Not DBInput("INV") Then
       Exit Function
'        SyoriKB = "ADD"
'        ProcHB = "B1"
    Else
'        SyoriKB = "MNT"
        ProcHB = "B2"
    End If
    
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    
    Body1Chk = False
        
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
    ' 仕入先
    Call iB1_Torcd_LostFocus
    If RTrim(iB1_Torcd) = "" Then
        CisFun.MB_MSG(2) = "     仕入先を入力して下さい            "
        CisFun.MBOX
        iB1_Torcd.SetFocus
        Exit Function
    End If
    If Not TorNmGet(iB1_Torcd, 8) Then
        CisFun.MB_MSG(2) = "     仕入先 未登録！            "
        CisFun.MB_MSG(4) = "    （取引先マスタ）            "
        CisFun.MBOX
        iB1_Torcd.SetFocus
        Exit Function
    End If
    ' 受入
    If RTrim(iB1_Ukeir) = "" Then
        CisFun.MB_MSG(2) = "     受入を入力して下さい            "
        CisFun.MBOX
        iB1_Ukeir.SetFocus
        Exit Function
    End If
    TUM.取引先CD = RTrim(iB1_Torcd)
    TUM.受入 = RTrim(iB1_Ukeir)
    TUM.取引先区分 = 1
    If Not TUMRead("", 1) Then
        CisFun.MB_MSG(2) = "    受入　未登録！（取引先マスタ）            "
        CisFun.MBOX
        iB1_Ukeir.SetFocus
        Exit Function
    End If
    ' 検収日
    If Trim(iB1_Ymd) = "" Then
        CisFun.MB_MSG(2) = "    検収日を入力して下さい            "
        CisFun.MBOX
       iB1_Ymd.SetFocus
       Exit Function
    End If
    If Not iB1_Ymd.cChkResult Then
        CisFun.MB_MSG(2) = "    正しい検収日を入力して下さい            "
        CisFun.MBOX
       iB1_Ymd.SetFocus
       Exit Function
    End If
    If RTrim(iB1_Ymd) > Format(Now(), "yyyymmdd") Then
        CisFun.MB_MSG(1) = "     検収日が未来の日付です。     "
        CisFun.MB_MSG(3) = "     　　　よろしいですか？      "
        CisFun.MB_Button = OK_CAN
        If Not CisFun.MBOX Then
            iB1_Ymd.SetFocus
            Exit Function
        End If
        CisFun.MB_MSG(1) = "        "
        CisFun.MB_MSG(3) = "        "
        CisFun.MB_Button = Error
    End If
    gInt = Kdou_Check(iB1_Ymd)
    If gInt = 1 Then
    '休日
        CisFun.MB_MSG(1) = "       検収日  は非稼働日です。     "
        CisFun.MB_MSG(3) = "     　　よろしいですか？      "
        CisFun.MB_Button = OK_CAN
        If Not CisFun.MBOX Then
            iB1_Ymd.SetFocus
            Exit Function
        End If
        CisFun.MB_MSG(1) = "        "
        CisFun.MB_MSG(3) = "        "
        CisFun.MB_Button = Error
    End If
    If gInt = 2 Then
    'カレンダ未登録
        CisFun.MB_MSG(1) = "     " & RTrim(iB1_Ymd.cYear) & "年" & RTrim(iB1_Ymd.cMonth) & "月のカレンダ未登録！     "
        CisFun.MB_MSG(3) = "     　　(　カレンダマスタ　)　"
        CisFun.MBOX
        iB1_Ymd.SetFocus
        Exit Function
    End If
    
    
    '　納入番号
    If RTrim(iB1_NouNo) = "" Then
        CisFun.MB_MSG(2) = "     納入番号を入力して下さい            "
        CisFun.MBOX
        iB1_NouNo.SetFocus
        Exit Function
    End If
    
    ' 背番号半角チェック
    If RTrim(RTrim(iB1_Seban)) <> "" And Not HalfSizeChk(RTrim(iB1_Seban)) Then
        With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "    背番号は半角にて入力して下さい            "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
        End With
        iB1_Seban.SetFocus
        Exit Function
    End If
    
    '　品番
    If RTrim(iB1_Hinbn) = "" Then
        CisFun.MB_MSG(2) = "     品番を入力して下さい            "
        CisFun.MBOX
        iB1_Hinbn.SetFocus
        Exit Function
    End If
    
    ' 品番取引先マスタ存在チェック
    gSL_Select = "SELECT * FROM 品番取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iB1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(iB1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(iB1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   AND 品目 = 1"
    If Not HTMRead(gSL_Select, 1) Then
        CisFun.MB_MSG(2) = "    品番取引先情報未登録      "
        CisFun.MBOX
        iB1_Seban.SetFocus
        Exit Function
    Else
        iB1_Seban = HTM.背番号
    End If
    Call Hinbn_Set
    
    ' 数量
    If CisFun.Val2(iB1_Suryo) = 0 Then
        CisFun.MB_MSG(2) = "     数量を入力して下さい            "
        CisFun.MBOX
        iB1_Suryo.SetFocus
        Exit Function
    End If
    
    'データ重複チェック
    If SyoriKB = "ADD" Then
        gSL_Select = "SELECT * FROM 受給検収テーブル"
        gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iB1_Hinbn) & "'"
        gSL_Select = gSL_Select & "   AND 仕入先 = '" & RTrim(iB1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(iB1_Ukeir) & "'"
        gSL_Select = gSL_Select & "   AND 納入番号 = '" & RTrim(iB1_NouNo) & "'"
        gSL_Select = gSL_Select & "   AND 検収日 = '" & RTrim(iB1_Ymd) & "'"
        gSL_Select = gSL_Select & "   AND 赤黒区分 = " & B1Chk_Aka.Value
'        gSL_Select = gSL_Select & "   AND 作成種別 = 'I'"
        If JYKRead(gSL_Select, 1) Then
            CisFun.MB_MSG(2) = "    仕入先･受入･検収日･納入番号･品番・赤黒が重複しています。            "
            CisFun.MBOX
            iB1_Torcd.SetFocus
            Exit Function
        End If
    Else
        If SyoriKB = "MNT" Then
            If RTrim(Save_Torcd) <> RTrim(iB1_Torcd) Or _
               RTrim(Save_Ukeir) <> RTrim(iB1_Ukeir) Or _
               RTrim(Save_YMD) <> RTrim(iB1_Ymd) Or _
               RTrim(Save_NouNo) <> RTrim(iB1_NouNo) Or _
               RTrim(Save_Seban) <> RTrim(iB1_Seban) Or _
               RTrim(Save_Hinbn) <> RTrim(iB1_Hinbn) Or _
               Save_Aka <> B1Chk_Aka.Value Then
                gSL_Select = "SELECT * FROM 受給検収テーブル"
                gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iB1_Hinbn) & "'"
                gSL_Select = gSL_Select & "   AND 仕入先 = '" & RTrim(iB1_Torcd) & "'"
                gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(iB1_Ukeir) & "'"
                gSL_Select = gSL_Select & "   AND 納入番号 = '" & RTrim(iB1_NouNo) & "'"
                gSL_Select = gSL_Select & "   AND 検収日 = '" & RTrim(iB1_Ymd) & "'"
                gSL_Select = gSL_Select & "   AND 赤黒区分 = " & B1Chk_Aka.Value
'                gSL_Select = gSL_Select & "   AND 作成種別 = 'I'"
                If JYKRead(gSL_Select, 1) Then
                    CisFun.MB_MSG(2) = "    仕入先･受入･検収日･納入番号･品番・赤黒が重複しています。            "
                    CisFun.MBOX
                    iB1_Torcd.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    
    Body1Chk = True
End Function
'****************************
'*      Body2ﾁｪｯｸ処理
'****************************
Private Function Body2Chk() As Boolean
    
    Body2Chk = False
        
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
    ' 選択
    If CisFun.Val2(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(16))) = 0 Then Exit Function
    
    RowColFlg = True
    If SyoriKB = "MNT" Then
        Call CisVsGrid1.vsColor(個別, &HC0FFC0)
    Else
        Call CisVsGrid1.vsColor(個別, &HFFC0FF)
    End If
    RowColFlg = False
    
    Body2Chk = True
End Function
'***********************************************************************************
'*
'*       データベース入力 ＆ 表示
'*
'***********************************************************************************
Private Function DBInput(InputKey As String) As Boolean
    Dim wSl_Where       As String
        
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("B")
    
    gSL_Select = "SELECT JK.*,ISNULL(TR.略称,'') 仕入先名,ISNULL(HN.表示品番,'') 表示品番,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(JK.赤黒区分,0) = 0 THEN '' ELSE '赤' END 赤黒 "
    gSL_Select = gSL_Select & " FROM 受給検収テーブル JK"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR "
    gSL_Select = gSL_Select & "   ON TR.取引先区分 = 1 "
    gSL_Select = gSL_Select & "   AND TR.取引先CD = JK.仕入先 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT "
    gSL_Select = gSL_Select & "   ON HT.品番 = JK.品番 "
    gSL_Select = gSL_Select & "   AND HT.取引先 = JK.仕入先 "
    gSL_Select = gSL_Select & "   AND HT.受入 = JK.受入 "
    gSL_Select = gSL_Select & "   AND HT.品目 = 1 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "   ON HN.品番 = JK.品番 "
        
    'Where句作成
    CisFun.Where_CreateKB = NewDocument
'    CisFun.Where_Create AlphaMe, "作成種別", "I"
    CisFun.Where_Create AlphaMe, "検収日", RTrim(iH1_Ymd)
    CisFun.Where_Create AlphaMe, "仕入先", RTrim(iH1_Torcd)
    CisFun.Where_Create AlphaMe, "JK.受入", RTrim(iH1_Ukeir)
    CisFun.Where_Create AlphaMe, "JK.品番", RTrim(iH1_Hinbn)
    CisFun.Where_Create AlphaMe, "納入番号", RTrim(iH1_NouNo)
    wSl_Where = CisFun.Where_Phrase
    
    gSL_Select = gSL_Select & wSl_Where
    
    gSL_Select = gSL_Select & " ORDER BY JK.仕入先,JK.受入,JK.検収日,JK.納入番号,JK.背番号,JK.品番,JK.赤黒区分 "
    
    If Not JYKRead(gSL_Select, 0) Then
'        If ProcHB <> "B2" Then
        If ProcHB = "H1" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     対象データが存在しません        "
                .MB_Button = OK
                .MBOX
            End With
        End If
        GoTo DBInput_Ed
    Else
        Call Grid1Set
        DBInput = True
    End If

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------------+
'+      削 除 処 理            +
'+----------------------------+
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
        .SQL = "DELETE FROM 受給検収テーブル "
        .SQL = .SQL & " WHERE 受給管理NO = " & CisFun.Val2(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(16))) & ""
        .DBExec

        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")
        
        On Error GoTo 0
    End With
    
'Grid再表示
    Call vsGrid1_ReSetDisp
    
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut(KB As Byte)
    DBPut = False
    Dim AddFlg  As Boolean

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        If SyoriKB = "ADD" Then
            .MB_MSG(2) = "     追加処理を行います      "
            .MB_Title = "追加処理"
        Else
            .MB_MSG(2) = "     修正処理を行います      "
            .MB_Title = "修正処理"
        End If
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err

'
    SaveKanriNo = 0
    
    If SyoriKB = "ADD" Then
        With CisDB
            .SQL = "番号獲得"
            .StoadoCount = 3
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = OutPut用
            .ParaNo = 3: .ParaIO = Return用

            .ParaNo = 1: .ParaValue = 19

            If Not .DBStored Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "     受給管理NOの採番に失敗しました    "
                    .MB_Title = "採番処理"
                    .MB_Button = Error
                    If Not .MBOX Then GoTo DBPut_Err
                End With
            Else
                .ParaNo = 2: SaveKanriNo = .ParaValue
            End If
        End With
    Else
        SaveKanriNo = CisFun.Val2(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(16)))
    End If
'
    Call ItemsClearJYK
    With JYK
        .受給管理NO = SaveKanriNo
        .受給管理行NO = 1
        .検収日 = RTrim(iB1_Ymd)
        .納入日 = RTrim(iB1_Ymd)
        .仕入先 = RTrim(iB1_Torcd)
        .受入 = RTrim(iB1_Ukeir)
        .納入番号 = RTrim(iB1_NouNo)
        .契約先 = HTM.契約先
        .代行契約先 = HTM.代行契約先
        .品番 = RTrim(iB1_Hinbn)
        .背番号 = RTrim(iB1_Seban)
        .赤黒区分 = B1Chk_Aka.Value
        .枚数 = CisFun.Val2(iB1_Maisu)
        .収容数 = CisFun.Val2(B1lb_Syuyo)
        .検収数 = CisFun.Val2(iB1_Suryo)
        .赤黒区分 = B1Chk_Aka.Value
        If B1Chk_Aka.Value = 0 Then
            .検収数 = CisFun.Val2(iB1_Suryo)
        Else
            .検収数 = CisFun.Val2(iB1_Suryo) * -1
        End If
        .作成種別 = "I"
    End With

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 追加処理 )
    If SyoriKB = "ADD" Then
        Call JYKInsert
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
        Call TextToGrid
    Else
' ( 更新処理 )
        Call JYKUpdate
        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
        'Call TextToGrid
    End If
'
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

'Grid再表示
    If SyoriKB = "MNT" Then
       Call vsGrid1_ReSetDisp
    End If
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        ReDim SColZ(21)

        Call .FixedSet(中中, 右中, あり, 3, 0, "No")
        '*-*-*-*-* 2008/04/11
        'Call .FixedSet(中中, 左中, あり, 20, 1, "仕 入 先")
        Call .FixedSet(中中, 左中, あり, 16, 1, "仕 入 先")
        '*-*-*-*-* 2008/04/11
        Call .FixedSet(中中, 中中, あり, 2, 2, "受")
        Call .FixedSet(中中, 中中, あり, 10, 3, "検収日")
        Call .FixedSet(中中, 中中, あり, 9, 4, "納入番号")
        '*-*-*-* 2008/04/11 Update Start
        'Call .FixedSet(中中, 左中, あり, 4, 5, "背番")
        Call .FixedSet(中中, 左中, あり, 8, 5, "背番")
        '*-*-*-* 2008/04/11 Update End
        Call .FixedSet(中中, 左中, あり, 30, 6, "品      番")
        Call .FixedSet(中中, 中中, あり, 4, 7, "赤黒")
        Call .FixedSet(中中, 右中, あり, 6, 8, "収容数")
        Call .FixedSet(中中, 右中, あり, 7, 9, "枚  数")
        Call .FixedSet(中中, 右中, あり, 7, 10, "端　数")
        Call .FixedSet(中中, 右中, あり, 9, 11, "入荷数")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 中中, あり, 0, 12, "X仕入先")
        Call .FixedSet(中中, 中中, あり, 0, 13, "X検収日")
        Call .FixedSet(中中, 中中, あり, 0, 14, "X品番")
        Call .FixedSet(中中, 中中, あり, 0, 15, "X赤黒")
        Call .FixedSet(中中, 中中, あり, 0, 16, "X検収管理NO")
        Call .FixedSet(中中, 中中, あり, 0, 17, "X検収管理行NO")
        Call .FixedSet(中中, 左中, あり, 0, 18, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp

    End With

End Sub
'****************************************************************************************'
'*      グリッド内容セット
'****************************************************************************************'
Private Sub Grid1Set()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
    Dsp_Kensu = Format(Ttl_Kensu, "#,###")
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid1.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
            .Rows = CisVsGrid1.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    
    ID = 0
    
    Do Until Not JYK_RDSTS
        With vsGrid1
            ID = ID + 1
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            '仕入先・名
            If RTrim(JYK.仕入先名) <> "" Then
                .TextMatrix(ID, GridCol(1)) = RTrim(JYK.仕入先) & ":" & RTrim(JYK.仕入先名)
            Else
                .TextMatrix(ID, GridCol(1)) = RTrim(JYK.仕入先)
            End If
            '受入
            .TextMatrix(ID, GridCol(2)) = RTrim(JYK.受入)
            ' 検収日
            If Trim(JYK.検収日) <> "" Then
               With CisFun
                    .DateE_BefVal = JYK.検収日
                    .DateEdit
                    gStr = .DateE_AfVal
               End With
               .TextMatrix(ID, GridCol(3)) = gStr
            End If
            '納入番号
            .TextMatrix(ID, GridCol(4)) = RTrim(JYK.納入番号)
            '背番号
            .TextMatrix(ID, GridCol(5)) = RTrim(JYK.背番号)
            '品番
            .TextMatrix(ID, GridCol(6)) = RTrim(JYK.表示品番)
            '赤黒
            .TextMatrix(ID, GridCol(7)) = RTrim(JYK.赤黒)
            '収容数
            .TextMatrix(ID, GridCol(8)) = CisFun.RSetFld(JYK.収容数, 7, "#,###")
            '枚数
            .TextMatrix(ID, GridCol(9)) = CisFun.RSetFld(JYK.枚数, 7, "#,###")
            '端数
            gLong = 0
            If JYK.収容数 <> 0 Then gLong = JYK.検収数 Mod JYK.収容数
            .TextMatrix(ID, GridCol(10)) = CisFun.RSetFld(gLong, 7, "#,###")
            '検収数
            .TextMatrix(ID, GridCol(11)) = CisFun.RSetFld(JYK.検収数, 9, "#,###")
            
            '非表示エリア
            .TextMatrix(ID, GridCol(12)) = RTrim(JYK.仕入先)
            .TextMatrix(ID, GridCol(13)) = RTrim(JYK.検収日)
            .TextMatrix(ID, GridCol(14)) = RTrim(JYK.品番)
            .TextMatrix(ID, GridCol(15)) = RTrim(JYK.赤黒区分)
            .TextMatrix(ID, GridCol(16)) = JYK.受給管理NO
            .TextMatrix(ID, GridCol(17)) = JYK.受給管理行NO
            .TextMatrix(ID, GridCol(18)) = ""
            
        End With
        
        Call JYKReadNext
    Loop
    
    Call JYKClose
    
    With vsGrid1
        .Row = 1: .Col = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  グリッドより入力域へセット
'+---------------------------------------------------------------------------------------+
Private Sub GridToText()
    
    With vsGrid1
        If Trim(.TextMatrix(.Row, CisVsGrid1.FixedGet(16))) = "" Then Exit Sub
        
        iB1_Torcd = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(12)))
        Call iB1_Torcd_LostFocus
        iB1_Ukeir = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(2)))
        iB1_Ymd = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(13)))
        iB1_NouNo = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(4)))
        iB1_Seban = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(5)))
        iB1_Hinbn = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(14)))
        If CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(15))) = 0 Then
            B1Chk_Aka.Value = 0
        Else
            B1Chk_Aka.Value = 1
        End If
        B1lb_Syuyo = Format(CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(8))), "##,###")
        iB1_Maisu = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(9)))
        B1lb_Hasu = Format(CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(10))), "###,###")
        iB1_Suryo = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(11)))
        
        If CisFun.Val2(iB1_Suryo) < 0 Then
           iB1_Suryo = Format(CisFun.Val2(iB1_Suryo) * -1, "###,###")
        End If
        
        '退避エリア
        Save_Torcd = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(12)))
        Save_Ukeir = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(2)))
        Save_YMD = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(13)))
        Save_NouNo = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(4)))
        Save_Seban = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(5)))
        Save_Hinbn = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(14)))
        Save_Aka = B1Chk_Aka.Value
    
    End With

End Sub
'+---------------------------------------------------------------------------------------+
'+  入力域からグリッド最終行へセット
'+---------------------------------------------------------------------------------------+
Private Sub TextToGrid()
    
    AddRow = AddRow + 1
    
    If AddRow >= vsGrid1.Rows Then
       vsGrid1.Rows = vsGrid1.Rows + 1
       vsGrid1.TopRow = AddRow
    End If
    
    With vsGrid1
        '仕入先
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(12)) = iB1_Torcd
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(1)) = RTrim(iB1_Torcd) & "：" & RTrim(B1lb_Tornm)
        '受入
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(2)) = iB1_Ukeir
        '検収日
        CisFun.DateE_BefVal = RTrim(iB1_Ymd)
        CisFun.DateEdit
        gStr = CisFun.DateE_AfVal
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(3)) = gStr
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(13)) = RTrim(iB1_Ymd)
        '納入番号
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(4)) = iB1_NouNo
        '背番号
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(5)) = iB1_Seban
        '品番
        gSL_Select = ""
        HNM.品番 = RTrim(iB1_Hinbn)
        Call HNMRead(gSL_Select, 1)
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(6)) = HNM.表示品番
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(14)) = iB1_Hinbn
        '赤黒
        If B1Chk_Aka.Value = 1 Then
           .TextMatrix(AddRow, CisVsGrid1.FixedGet(7)) = "赤"
        Else
           .TextMatrix(AddRow, CisVsGrid1.FixedGet(7)) = ""
        End If
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(15)) = B1Chk_Aka.Value
        '収容数
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(8)) = Format(CisFun.Val2(B1lb_Syuyo), "##,###")
        '枚数
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(9)) = Format(CisFun.Val2(iB1_Maisu), "###,###")
        '端数
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(10)) = Format(CisFun.Val2(B1lb_Hasu), "###,###")
        '数量
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(11)) = Format(CisFun.Val2(iB1_Suryo), "#,###,###")
        
        
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(16)) = SaveKanriNo
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(17)) = 1
        .TextMatrix(AddRow, CisVsGrid1.FixedGet(18)) = ""
        
    
    End With

End Sub
'+---------------------------------------+
'+  グリッド (Click) サイズ情報
'+---------------------------------------+
Private Sub VSGrid1_Click()
    With vsGrid1
        If ProcHB = "B2" Then
            Call HeadBodyClear("B1")
            Call GridToText
        End If
    End With
End Sub
'+---------------------------------------+
'+  グリッド (DblClick)
'+---------------------------------------+
Private Sub vsGrid1_DblClick()
    
    With vsGrid1
        If ProcHB = "B2" Then
            If Trim(.TextMatrix(.Row, CisVsGrid1.FixedGet(16))) <> "" Then
               Call ReturnPress
            End If
        End If
    End With
End Sub
'+---------------------------------------+
'+  グリッド (GotFocus)
'+---------------------------------------+
Private Sub vsGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace And Trim(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(16))) <> "" Then
       Call ReturnPress
    End If
End Sub
'+---------------------------------------+
'+  グリッド (RowColChange)
'+---------------------------------------+
Private Sub vsGrid1_RowColChange()
    If Not FormAct Then Exit Sub
    If ProcHB <> "B2" Then Exit Sub
    If RowColFlg Then Exit Sub
    Call VSGrid1_Click
End Sub
'+---------------------------------------+
'+  グリッド 再表示
'+---------------------------------------+
Private Sub vsGrid1_ReSetDisp()
    
    SaveRow = vsGrid1.Row
    RowColFlg = True
    
    Call DBInput("INV")
    
    RowColFlg = False
    If SaveRow > vsGrid1.Rows - 1 Then
        SaveRow = SaveRow - 1
    End If
    If SyoriKB <> "ADD" Then
        vsGrid1.Row = SaveRow
    End If
    If vsGrid1.RowIsVisible(SaveRow) = False Then
        vsGrid1.TopRow = SaveRow
    End If

End Sub




