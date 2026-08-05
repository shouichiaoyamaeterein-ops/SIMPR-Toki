VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CisYMDwB_6.0.ocx"
Begin VB.Form CKN0410 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "かんばん流動管理問合せ"
   ClientHeight    =   10980
   ClientLeft      =   1785
   ClientTop       =   1575
   ClientWidth     =   15330
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10980
   ScaleWidth      =   15330
   WindowState     =   2  '最大化
   Begin VB.PictureBox PNL_Count 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   495
      Left            =   12270
      ScaleHeight     =   495
      ScaleWidth      =   2520
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1590
      Width           =   2520
      Begin Cis3D_v60.CIS3D UC_3D6 
         Height          =   435
         Left            =   30
         Top             =   30
         Width           =   2460
         _ExtentX        =   4339
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
         Begin Cis3D_v60.CIS3D H1lb_Prtsu 
            Height          =   315
            Left            =   1350
            Top             =   60
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BackColor       =   16761024
            ForeColor       =   0
            Caption         =   "ZZ,ZZ9"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
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
   Begin Cis3D_v60.CIS3D PopupSMenu 
      Height          =   1170
      Left            =   5370
      Top             =   4155
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   2064
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
      Begin Cis3D_v60.CIS3D POP_Menu 
         Height          =   285
         Index           =   0
         Left            =   15
         Top             =   15
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
         ForeColor       =   0
         Caption         =   "有効"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D POP_Menu 
         Height          =   285
         Index           =   1
         Left            =   15
         Top             =   300
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
         ForeColor       =   0
         Caption         =   "無効"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D POP_Menu 
         Height          =   285
         Index           =   2
         Left            =   15
         Top             =   585
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
         Caption         =   "発行"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D POP_Menu 
         Height          =   285
         Index           =   3
         Left            =   15
         Top             =   870
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
         Caption         =   "クリア"
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
         cAlingnment     =   7
      End
   End
   Begin Cis3D_v60.CIS3D B1_Area1 
      Height          =   7065
      Left            =   450
      Top             =   3420
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   12462
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
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   6915
         Left            =   75
         Top             =   75
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   12197
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
            Height          =   6720
            Left            =   90
            TabIndex        =   14
            Top             =   105
            Width           =   3525
            _Version        =   196608
            _ExtentX        =   6218
            _ExtentY        =   11853
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
            FormatString    =   "№ |流動日付  |発行|指示|||"
            Rows            =   21
            BackColor       =   12648447
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
            AllowBigSelection=   0   'False
         End
         Begin vsFlexLib.vsFlexArray vsGrid2 
            Height          =   6720
            Left            =   3630
            TabIndex        =   15
            Top             =   105
            Width           =   3540
            _Version        =   196608
            _ExtentX        =   6244
            _ExtentY        =   11853
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
            FormatString    =   "№ |流動日付  |発行|指示|||"
            Rows            =   21
            BackColor       =   12648447
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
            AllowBigSelection=   0   'False
         End
         Begin vsFlexLib.vsFlexArray vsGrid3 
            Height          =   6720
            Left            =   7185
            TabIndex        =   16
            Top             =   105
            Width           =   3780
            _Version        =   196608
            _ExtentX        =   6667
            _ExtentY        =   11853
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
            FormatString    =   "№ |流動日付  |発行|指示|||"
            Rows            =   21
            BackColor       =   12648447
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
            AllowBigSelection=   0   'False
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
      Height          =   1215
      Left            =   450
      ScaleHeight     =   1155
      ScaleWidth      =   14295
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2100
      Width           =   14355
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1095
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1931
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
         cPositionX      =   30
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Seban 
            Height          =   360
            Left            =   135
            TabIndex        =   3
            Top             =   390
            Width           =   1095
            _ExtentX        =   1931
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1065
         Index           =   2
         Left            =   1395
         Top             =   60
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   1879
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
            Left            =   105
            TabIndex        =   4
            Top             =   405
            Width           =   3180
            _ExtentX        =   5609
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1065
         Index           =   3
         Left            =   4755
         Top             =   60
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   1879
         ForeColor       =   16711680
         Caption         =   "取 引 先"
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
            Left            =   105
            TabIndex        =   5
            Top             =   405
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
            Height          =   345
            Left            =   1095
            Top             =   405
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   609
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXX"
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1065
         Index           =   4
         Left            =   8265
         Top             =   60
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1879
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
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   360
            Left            =   105
            TabIndex        =   6
            Top             =   405
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1065
         Index           =   5
         Left            =   8880
         Top             =   60
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1879
         ForeColor       =   16711680
         Caption         =   "流 動 日 付"
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
         cPositionX      =   120
         cPositionY      =   30
         Begin CisYMD_With_Btn.CisYMDwB iH1_SRYmd 
            Height          =   330
            Left            =   570
            TabIndex        =   7
            Top             =   315
            Width           =   1845
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
            Object.Width           =   1845
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
         Begin Cis3D_v60.CIS3D CIS3D32 
            Height          =   330
            Left            =   90
            Top             =   315
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   582
            Caption         =   "開始"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cPositionX      =   15
         End
         Begin Cis3D_v60.CIS3D CIS3D34 
            Height          =   330
            Left            =   90
            Top             =   660
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   582
            Caption         =   "終了"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cPositionX      =   15
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_ERYmd 
            Height          =   330
            Left            =   570
            TabIndex        =   8
            Top             =   660
            Width           =   1845
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
            Object.Width           =   1845
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1065
         Index           =   6
         Left            =   11400
         Top             =   60
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1879
         ForeColor       =   16711680
         Caption         =   "発行区分"
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
         Begin VB.CheckBox ChkH1_HKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "未流動"
            Height          =   300
            Index           =   3
            Left            =   870
            TabIndex        =   12
            Top             =   645
            Width           =   1035
         End
         Begin VB.CheckBox ChkH1_HKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "滞留"
            Height          =   300
            Index           =   2
            Left            =   870
            TabIndex        =   11
            Top             =   345
            Width           =   810
         End
         Begin VB.CheckBox ChkH1_HKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "無効"
            Height          =   300
            Index           =   1
            Left            =   75
            TabIndex        =   10
            Top             =   645
            Width           =   810
         End
         Begin VB.CheckBox ChkH1_HKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "有効"
            Height          =   300
            Index           =   0
            Left            =   75
            TabIndex        =   9
            Top             =   345
            Width           =   810
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1065
         Index           =   7
         Left            =   13335
         Top             =   60
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   1879
         ForeColor       =   16711680
         Caption         =   "発行№ "
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
         cPositionX      =   60
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Hno 
            Height          =   360
            Left            =   210
            TabIndex        =   13
            Top             =   405
            Width           =   510
            _ExtentX        =   900
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
      Left            =   7125
      ScaleHeight     =   660
      ScaleWidth      =   7605
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   660
      Width           =   7665
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   5070
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
         Left            =   5910
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
         Left            =   6750
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
         Left            =   3390
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
         Left            =   2550
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
         Left            =   1710
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
      Begin CisBtn_60.CisBtn PB_LOK 
         Height          =   585
         Left            =   870
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
         BtnType         =   12
         BtnCaption      =   "照会"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16711935
         BtnShadowColor  =   16777215
         PFCaption       =   "( F4 )"
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
         PFBackColor     =   16744703
      End
      Begin CisBtn_60.CisBtn PB_MNT 
         Height          =   585
         Left            =   30
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
      Begin CisBtn_60.CisBtn PB_Prt 
         Height          =   585
         Left            =   4230
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
      IMEMode         =   3  'ｵﾌ固定
      Left            =   960
      TabIndex        =   0
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
   Begin Cis3D_v60.CIS3D B1_Area2 
      Height          =   7065
      Left            =   11700
      Top             =   3420
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   12462
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
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   6915
         Left            =   30
         Top             =   75
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   12197
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
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   300
            Left            =   240
            Top             =   525
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            Caption         =   "サイクル"
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
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   375
            Left            =   15
            Top             =   15
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            BackColor       =   0
            ForeColor       =   65535
            Caption         =   "【かんばん情報】"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
         End
         Begin vsFlexLib.vsFlexArray vsGrid4 
            Height          =   1980
            Left            =   45
            TabIndex        =   17
            Top             =   2535
            Width           =   2850
            _Version        =   196608
            _ExtentX        =   5027
            _ExtentY        =   3492
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
            FormatString    =   "№|納 入 日  |便|枚数|||"
            Rows            =   6
            Cols            =   7
            BackColor       =   12648384
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
         End
         Begin Cis3D_v60.CIS3D B_Syuyo 
            Height          =   300
            Index           =   1
            Left            =   240
            Top             =   1425
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            Caption         =   "収容数１"
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
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   300
            Left            =   240
            Top             =   825
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
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
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   375
            Left            =   15
            Top             =   2190
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            BackColor       =   0
            ForeColor       =   65535
            Caption         =   "【発注情報】"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   390
            Left            =   0
            Top             =   5280
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   688
            BackColor       =   0
            ForeColor       =   65535
            Caption         =   "【かんばん情報】"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D12 
            Height          =   330
            Left            =   150
            Top             =   5760
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   582
            Caption         =   "発行枚数"
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
         Begin Cis3D_v60.CIS3D B_Hkbn 
            Height          =   330
            Index           =   1
            Left            =   150
            Top             =   6090
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   582
            BackColor       =   16761024
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
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   15
         End
         Begin Cis3D_v60.CIS3D B_Hkbn 
            Height          =   330
            Index           =   2
            Left            =   150
            Top             =   6420
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   582
            BackColor       =   16761087
            Caption         =   "無効"
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
            cPositionX      =   15
         End
         Begin Cis3D_v60.CIS3D B_Hkbn 
            Height          =   330
            Index           =   3
            Left            =   1515
            Top             =   6090
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   582
            Caption         =   "滞留"
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
            cPositionX      =   15
         End
         Begin Cis3D_v60.CIS3D B_Hkbn 
            Height          =   330
            Index           =   0
            Left            =   1515
            Top             =   6420
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   582
            Caption         =   "未流動"
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
            cPositionX      =   15
         End
         Begin Cis3D_v60.CIS3D B1lb_Cycle 
            Height          =   300
            Left            =   1425
            Top             =   525
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            BackColor       =   14737632
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B1lb_Syuyo 
            Height          =   300
            Index           =   1
            Left            =   1425
            Top             =   1425
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            BackColor       =   14737632
            Caption         =   "#,###"
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
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_KaiSu 
            Height          =   300
            Left            =   1425
            Top             =   825
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            BackColor       =   14737632
            Caption         =   "#,###"
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
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_Maisu 
            Height          =   315
            Index           =   5
            Left            =   1515
            Top             =   5760
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            BackColor       =   14737632
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_Maisu 
            Height          =   330
            Index           =   1
            Left            =   930
            Top             =   6090
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
            BackColor       =   14737632
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_Maisu 
            Height          =   330
            Index           =   2
            Left            =   930
            Top             =   6420
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
            BackColor       =   14737632
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_Maisu 
            Height          =   330
            Index           =   3
            Left            =   2295
            Top             =   6090
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
            BackColor       =   14737632
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_Maisu 
            Height          =   330
            Index           =   0
            Left            =   2295
            Top             =   6420
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   582
            BackColor       =   14737632
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D CIS3D27 
            Height          =   300
            Left            =   60
            Top             =   4530
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   529
            Caption         =   "発注枚数合計"
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
         Begin Cis3D_v60.CIS3D CIS3D28 
            Height          =   300
            Left            =   60
            Top             =   4830
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   529
            Caption         =   "発注予定枚数"
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
         Begin Cis3D_v60.CIS3D B1lb_HCSu 
            Height          =   300
            Left            =   2025
            Top             =   4530
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            BackColor       =   14737632
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_HYSu 
            Height          =   300
            Left            =   2025
            Top             =   4830
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            BackColor       =   14737632
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B_Syuyo 
            Height          =   300
            Index           =   2
            Left            =   240
            Top             =   1725
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            BackColor       =   16777152
            Caption         =   "収容数２"
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
         Begin Cis3D_v60.CIS3D B1lb_Syuyo 
            Height          =   300
            Index           =   2
            Left            =   1425
            Top             =   1725
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            BackColor       =   16777152
            Caption         =   "#,###"
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
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   300
            Index           =   0
            Left            =   240
            Top             =   1125
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            Caption         =   "変 更 日"
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
         Begin Cis3D_v60.CIS3D B1lb_SYmd 
            Height          =   300
            Left            =   1425
            Top             =   1125
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            BackColor       =   14737632
            Caption         =   "YYYY/MM/DD"
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
      Caption         =   "【 かんばん流動管理問合せ  】"
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14355
         Top             =   90
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
         Left            =   11865
         Top             =   30
         Width           =   2490
         _ExtentX        =   4392
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
   End
End
Attribute VB_Name = "CKN0410"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   かんばん流動管理問合せ
'**       フォームID    :   CKN0410
'**       処理概要      :
'**
'**       作  成  日    :   2004/01/17  By CIS
'**       変  更  日    :   2004/02/19  By CIS - 収容数の背景色を変更可能にする対応
'**       変  更  日    :   2004/07/09  By CIS - マウス右クリック時の処理修正
'**       　  　  　                               > 1.マウス右クリック行を対象に変更
'**       　  　  　                               > 2.最下位行をクリック時にパネル下位部分が表示不可になる
'**       　  　  　                               > 3.スクロール後に実行するとパネル表示位置がずれる
'**       変  更  日    :   2005/03/01  By CIS - 最終発行連番を流動管理テーブルより表示に変更
'**       変  更  日    :   2005/03/02  By CIS - 最終発行連番は品番取引先マスタorかんばん流動管理テーブルの値が大きい方を使用
'**       変  更  日    :   2005/12/21  By CIS - 収容数の背景色を変更可能にする対応の不具合修正
'**       変  更  日    :   2006/12/20  By CIS - 版数の表示
'**       変  更  日    :   2008/04/11  By CIS - 背番号８桁対応
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl

    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKB                 As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM                 As String       ' 処理名称(新規･修正･削除 等)
    Dim CisVsGrid1              As New CisVsGrid3
    Dim CisVsGrid2              As New CisVsGrid3
    Dim CisVsGrid3              As New CisVsGrid3
    Dim CisVsGrid4              As New CisVsGrid3
    Dim wHCSu                   As Long         ' 発注枚数合計
    Dim Hkbn_BackColor(3)       As Long         ' 発行区分の背景色
    Dim Syuyo_BackColor(2)      As Long         ' 収容数の背景色
    
    Dim mRType                  As Integer
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

    Select Case App.Title
        Case "PKN0420"
            mRType = 1
        Case "PKN0430"
            mRType = 2
        Case Else
            mRType = 0
    End Select
    Dummy.Left = -1000
    Dummy.Enabled = False

    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止

'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
'   流動日付
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 0
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 2
        .SelectCol = 8
        .InitGet = False
        .Init
    End With
    With CisVsGrid2
        Set .GridObj = vsGrid2
        .NoSpaceCol = 0
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 2
        .SelectCol = 8
        .InitGet = False
        .Init
    End With
    With CisVsGrid3
        Set .GridObj = vsGrid3
        .NoSpaceCol = 0
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 2
        .SelectCol = 8
        .InitGet = False
        .Init
    End With
    Call GridInit
'   発注情報
    With CisVsGrid4
        Set .GridObj = vsGrid4
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 10
        .InitGet = False
        .Init
    End With
    Call GridInit4

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = "LOK"
    Call SyoriName(SyoriKB)
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

'   <><><> 背景色の退避（発行区分） <><><>
    For gInt = 0 To 3
        Select Case B_Hkbn(gInt).BackColor
            Case &HC0C0C0:  Hkbn_BackColor(gInt) = &HC0FFFF
            Case Else:      Hkbn_BackColor(gInt) = B_Hkbn(gInt).BackColor
        End Select
    Next gInt
'   <><><> 背景色の退避（収容数） <><><>
    For gInt = 1 To 2
        Select Case B_Syuyo(gInt).BackColor
            Case &HC0C0C0:  Syuyo_BackColor(gInt) = &HC0C0C0
            Case Else:      Syuyo_BackColor(gInt) = B_Syuyo(gInt).BackColor
        End Select
    Next gInt

    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
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
    Key_Acc = False

    Select Case Key_Code
           Case vbKeyF2:        If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF4:        If PB_LOK.Visible Then Call PB_LOK_Click    '【照会】
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF8:        If PB_Prt.Visible Then Call PB_Prt_Click    '【印刷】
           Case vbKeyPageUp:
                                If ProcHB = "H1" Then
                                    If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
                                Else
                                    Exit Function
                                End If
           Case vbKeyPageDown:
                                If ProcHB = "H1" Then
                                    If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
                                Else
                                    Exit Function
                                End If
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_Can_Click    '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      '【入力】
                                    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                    Dummy.Enabled = False
                                End If
           Case Else:           Exit Function
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
       Call HeadBodyClear("B")
       Call SyoriName(SyoriKB)
       iH1_Seban.SetFocus
       GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If ProcHB = "H1" Then
        If SyoriKB <> "MNT" Then
            SyoriKB = "MNT"
            Call SyoriName(SyoriKB)
        End If
        iH1_Seban.SetFocus
    End If
End Sub
'+----------------------------+
'+    照  会  キ ー(F4)       +
'+----------------------------+
Private Sub PB_LOK_Click()
    If ProcHB = "H1" Then
        If SyoriKB <> "LOK" Then
            SyoriKB = "LOK"
            Call SyoriName(SyoriKB)
        End If
        iH1_Seban.SetFocus
    End If
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    If ProcHB = "H1" Then
        Call CisFrm.MousePT(11)

        Call SqlSelectSet("Back")
        Call DBInput("Back")

        iH1_Seban.SetFocus

        Call CisFrm.MousePT(1)
    End If
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    If ProcHB = "H1" Then
        Call CisFrm.MousePT(11)

        Call SqlSelectSet("Next")
        Call DBInput("Next")

        iH1_Seban.SetFocus

        Call CisFrm.MousePT(1)
    End If
End Sub
'+----------------------------+
'+    印　刷  キ ー(F8)       +
'+----------------------------+
Private Sub PB_Prt_Click()
    If ProcHB = "H1" Then
        kn_Call = "C"
        kn_PrtKB = 0

        CKN0220.Show vbModal
        Unload CKN0220
        Set CKN0220 = Nothing

        Call DBSelect_KHT       ' 件数表示
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Seban" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_Ukeir: Exit Sub
'   日付
    If PB_Look.Tag = "iH1_SRYmd" Then iH1_SRYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_ERYmd" Then iH1_ERYmd.ShowCalender: Exit Sub
End Sub
'+----------------------------+
'+      品番検索              +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RVI_Hinmoku = 1

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Seban = RV_Seban
        iH1_Hinbn = RV_Hinbn
        iH1_Torcd = RV_TorcdK
        iH1_Ukeir = RV_Ukeir
        H1lb_Tornm = ""
        If TorNmGet(iH1_Torcd, 9) Then
           H1lb_Tornm = TRM_RName
        End If
    End If
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
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+----------------------------+
'+      取引先受入検索        +
'+----------------------------+
Private Sub Look_Ukeir()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 9            ' [発注]手配先/契約先/社内検索
    RV_Torcd = iH1_Torcd
    RV_Kojyo = ""

    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
        iH1_Ukeir = RV_Ukeir
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
    If TorNmGet(iH1_Torcd, 9) Then
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
'+  iH1_SRYmd
'+-----------------------------+
Private Sub iH1_SRYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_SRYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_ERYmd
'+-----------------------------+
Private Sub iH1_ERYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_ERYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  ChkH1_HKbn
'+-----------------------------+
Private Sub ChkH1_HKbn_Click(Index As Integer)
    If ChkH1_HKbn(Index) = 1 Then
        Select Case Index
            Case 1
                CisVsGrid1.PatanNo = 0
                CisVsGrid2.PatanNo = 0
                CisVsGrid3.PatanNo = 0
            Case Else
                CisVsGrid1.PatanNo = 1
                CisVsGrid2.PatanNo = 1
                CisVsGrid3.PatanNo = 1
        End Select
        Call CisVsGrid1.LayOutChg
        Call CisVsGrid2.LayOutChg
        Call CisVsGrid3.LayOutChg
        If Index = 1 Then
            If ChkH1_HKbn(0) = 1 Then ChkH1_HKbn(0) = 0
            If ChkH1_HKbn(2) = 1 Then ChkH1_HKbn(2) = 0
            If ChkH1_HKbn(3) = 1 Then ChkH1_HKbn(3) = 0
        Else
            If ChkH1_HKbn(1) = 1 Then ChkH1_HKbn(1) = 0
        End If
    Else
        Select Case Index
            Case 1
                CisVsGrid1.PatanNo = 1
                CisVsGrid2.PatanNo = 1
                CisVsGrid3.PatanNo = 1
                Call CisVsGrid1.LayOutChg
                Call CisVsGrid2.LayOutChg
                Call CisVsGrid3.LayOutChg
        End Select
    End If
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
    Static w_Color      As Long
    w_Color = &H80000008 '初期色ｾｯﾄ(黒)
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
        Case Else
            SyoriNM = ""
    End Select
    For Each gObj In Me.Controls
        If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
        If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
'       * PF-Key 表示設定 *
        If gObj.Name = "PB_ADD" Or _
           gObj.Name = "PB_MNT" Or _
           gObj.Name = "PB_DEL" Or _
           gObj.Name = "PB_LOK" Then
           gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
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
    B1_Area2.Enabled = B1Mode: B1_Area2.BackColor = B1Color

    PB_LOK.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_Look.Visible = False
    PB_Prt.Visible = False
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode

    PopupSMenu.Visible = False
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
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = "": .Enabled = True
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
    If HeadBody = "B" Then
        Call CisVsGrid1.Clear
        Call CisVsGrid2.Clear
        Call CisVsGrid3.Clear
        Call CisVsGrid4.Clear
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
        If SyoriKB = "MNT" Then
            If Not Body1Chk Then GoTo ReturnPress_Ed
            If Not DBPut Then GoTo ReturnPress_Ed
        End If
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
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

'   背番号
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

    gSL_Select = "select * from 品番取引先マスタ"
    gSL_Select = gSL_Select & " where 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   and 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and 背番号 = '" & RTrim(iH1_Seban) & "'"
    If Not HTMRead(gSL_Select, 1, 1) Then
        gSL_Select = ""
        gSL_Select = gSL_Select & "Select"
        gSL_Select = gSL_Select & "  htm.背番号"
        gSL_Select = gSL_Select & ", htm.取引先"
        gSL_Select = gSL_Select & ", htm.受入"
        gSL_Select = gSL_Select & ", htm.品番"
        gSL_Select = gSL_Select & ", hnm.表示品番"
        gSL_Select = gSL_Select & ", hnm.社内背番"
'▽▽------------------------------------------------------<< 2005/03/02 >> Update Start
'▽▽------------------------------------------------------<< 2005/03/01 >> Update Start
'       gSL_Select = gSL_Select & ", htm.最終発行連番"
'       gSL_Select = gSL_Select & ", isnull(krt.発行NO, 0) 最終発行連番"
        gSL_Select = gSL_Select & ", case when htm.最終発行連番 >= isnull(krt.発行NO, 0)"
        gSL_Select = gSL_Select & "       then htm.最終発行連番"
        gSL_Select = gSL_Select & "       else krt.発行NO"
        gSL_Select = gSL_Select & "  end 最終発行連番"
'△△------------------------------------------------------<< 2005/03/01 >> Update End
'△△------------------------------------------------------<< 2005/03/02 >> Update End
        gSL_Select = gSL_Select & ", rtrim(trm.略称) 取引先名"
        gSL_Select = gSL_Select & ", htm.手配区分"
        gSL_Select = gSL_Select & " from 品番取引先マスタ htm"
        gSL_Select = gSL_Select & "      left outer join 品番マスタ hnm"
        gSL_Select = gSL_Select & "        on hnm.品番 = htm.品番"
        gSL_Select = gSL_Select & "      left outer join 取引先マスタ trm"
        gSL_Select = gSL_Select & "        on trm.取引先CD = htm.取引先"
        gSL_Select = gSL_Select & "       and trm.取引先区分 = 1"
'▽▽------------------------------------------------------<< 2005/03/01 >> Insert Start
        gSL_Select = gSL_Select & "      left outer join"
        gSL_Select = gSL_Select & " (select 取引先"
        gSL_Select = gSL_Select & "       , 受入"
        gSL_Select = gSL_Select & "       , 品番"
        gSL_Select = gSL_Select & "       , max(発行NO) 発行NO"
        gSL_Select = gSL_Select & "    from かんばん流動管理テーブル"
        gSL_Select = gSL_Select & "   group by 取引先"
        gSL_Select = gSL_Select & "          , 受入"
        gSL_Select = gSL_Select & "          , 品番"
        gSL_Select = gSL_Select & " ) krt"
        gSL_Select = gSL_Select & "    on krt.取引先 = htm.取引先"
        gSL_Select = gSL_Select & "   and krt.受入 = htm.受入"
        gSL_Select = gSL_Select & "   and krt.品番 = htm.品番"
'△△------------------------------------------------------<< 2005/03/01 >> Insert End
        gSL_Select = gSL_Select & " where htm.背番号 = '" & iH1_Seban & "'"
        gSL_Select = gSL_Select & "   and htm.品目 = 1"
        gSL_Select = gSL_Select & "   and htm.SK区分 = 0"
       'gSL_Select = gSL_Select & "   and htm.手配区分 <> '2'"
        gSL_Select = gSL_Select & " order by"
        gSL_Select = gSL_Select & "  htm.品番"
        gSL_Select = gSL_Select & ", htm.取引先"
        gSL_Select = gSL_Select & ", htm.受入"
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
            Else
                iH1_Torcd = RTrim(HTM.取引先)
                iH1_Ukeir = RTrim(HTM.受入)
                iH1_Hinbn = RTrim(HTM.品番)
            End If
            Call HTMClose(1)
        Else
            Call HTMClose(1)

'           品番
            If Trim(iH1_Hinbn) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    品番を入力して下さい            "
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_Hinbn.SetFocus
                Exit Function
            End If
'           取引先
            If Trim(iH1_Torcd) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    取引先を入力して下さい            "
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_Torcd.SetFocus
                Exit Function
            End If
            If Not TorNmGet(iH1_Torcd, 9) Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    取引先　未登録！（取引先マスタ）            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
               iH1_Torcd.SetFocus
               Exit Function
            End If
'           受入
            If Trim(iH1_Ukeir) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    受入を入力して下さい            "
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_Ukeir.SetFocus
                Exit Function
            End If
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
'           <<< 品番取引先マスタ >>>
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    品番取引先マスタ未登録            "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
            iH1_Seban.SetFocus
            Exit Function

        End If
    End If

'   流動日付
    If Trim(iH1_SRYmd) <> "" Then
        If CisFun.Val2(iH1_SRYmd.cDay) = 0 Then
            iH1_ERYmd = ""
        Else
            If Not iH1_SRYmd.cChkResult Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    流動日付（開始）を正しく入力して下さい            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_SRYmd.SetFocus
                Exit Function
            End If
        End If
    End If
    If Trim(iH1_ERYmd) <> "" Then
        If Not iH1_ERYmd.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    流動日付（終了）を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_ERYmd.SetFocus
            Exit Function
        End If
    End If
    If Trim(iH1_SRYmd) <> "" And Trim(iH1_ERYmd) <> "" Then
       If iH1_SRYmd > iH1_ERYmd Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    流動日付を正しく入力して下さい（開始＞終了）            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iH1_SRYmd.SetFocus
          Exit Function
       End If
    End If

    Call SqlSelectSet("INV")
    If Not DBInput("INV") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    該当データが存在しません｡            "
            .MB_MSG(3) = "    再度抽出条件を入力してください       "
            .MB_Title = "抽出処理"
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If

    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False

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
    If Not KRTRead(gSL_Select, 1) Then
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
        iH1_Seban = RTrim(KRT.背番号)
        iH1_Hinbn = RTrim(KRT.品番)
        iH1_Torcd = RTrim(KRT.取引先)
        iH1_Ukeir = RTrim(KRT.受入)
        H1lb_Tornm = RTrim(KRT.取引先名)
        Select Case InputKey
            Case "INV": If Not BodySet("ALL") Then GoTo DBInput_Ed
            Case Else:  If Not BodySet("") Then GoTo DBInput_Ed
        End Select
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
Private Sub SqlSelectSet(InputKey As String)
'
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  htm.品番"
    gSL_Select = gSL_Select & ", htm.取引先"
    gSL_Select = gSL_Select & ", htm.受入"
    gSL_Select = gSL_Select & ", htm.背番号"
    gSL_Select = gSL_Select & ", htm.サイクル"
    gSL_Select = gSL_Select & ", htm.変更日"
    gSL_Select = gSL_Select & ", htm.収容数1"
    gSL_Select = gSL_Select & ", htm.収容数2"
    gSL_Select = gSL_Select & ", trm.略称 取引先名"
    gSL_Select = gSL_Select & ", isnull(kkm.回転枚数, 0) 回転枚数"
    gSL_Select = gSL_Select & ", sum(case when krt.発行区分 = 0 then 1 else 0 end) 未流件数"
    gSL_Select = gSL_Select & ", sum(case when krt.発行区分 = 1 then 1 else 0 end) 有効件数"
    gSL_Select = gSL_Select & ", sum(case when krt.発行区分 = 2 then 1 else 0 end) 無効件数"
    gSL_Select = gSL_Select & ", sum(case when krt.発行区分 = 3 then 1 else 0 end) 滞留件数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  かんばん流動管理テーブル krt"
    gSL_Select = gSL_Select & ", 品番取引先マスタ htm"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = htm.取引先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    gSL_Select = gSL_Select & "  left outer join かんばん回転枚数テーブル kkm"
    gSL_Select = gSL_Select & "    on kkm.年月 = '" & Format(Date, "yyyymm") & "'"
    gSL_Select = gSL_Select & "   and kkm.仕入先 = htm.取引先"
    gSL_Select = gSL_Select & "   and kkm.品番 = htm.品番"
    gSL_Select = gSL_Select & "   and kkm.受入 = htm.受入"
    gSL_Select = gSL_Select & " where htm.取引先 = krt.取引先"
    gSL_Select = gSL_Select & "   and htm.受入 = krt.受入"
    gSL_Select = gSL_Select & "   and htm.品番 = krt.品番"
    gSL_Select = gSL_Select & "   and"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
    'gSL_Select = gSL_Select & " ( htm.背番号 + space(5 - len(htm.背番号))"
    gSL_Select = gSL_Select & " ( htm.背番号 + space(8 - len(htm.背番号))"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
    gSL_Select = gSL_Select & " + htm.品番 + space(25 - len(htm.品番))"
    gSL_Select = gSL_Select & " + htm.取引先 + space(7 - len(htm.取引先))"
    gSL_Select = gSL_Select & " + htm.受入 + space(2 - len(htm.受入))"
    gSL_Select = gSL_Select & " ) "
    Select Case InputKey
        Case "Back": gSL_Select = gSL_Select & "<"
        Case "Next": gSL_Select = gSL_Select & ">"
        Case Else:   gSL_Select = gSL_Select & "="
    End Select
    gSL_Select = gSL_Select & " ( '"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE SART
    'gSL_Select = gSL_Select & iH1_Seban & Space(5 - CisFun.Len2(iH1_Seban))
    gSL_Select = gSL_Select & iH1_Seban & Space(8 - CisFun.Len2(iH1_Seban))
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
    gSL_Select = gSL_Select & iH1_Hinbn & Space(25 - CisFun.Len2(iH1_Hinbn))
    gSL_Select = gSL_Select & iH1_Torcd & Space(7 - CisFun.Len2(iH1_Torcd))
    gSL_Select = gSL_Select & iH1_Ukeir & Space(2 - CisFun.Len2(iH1_Ukeir))
    gSL_Select = gSL_Select & "' )"
    gSL_Select = gSL_Select & " group by"
    gSL_Select = gSL_Select & "  htm.品番"
    gSL_Select = gSL_Select & ", htm.取引先"
    gSL_Select = gSL_Select & ", htm.受入"
    gSL_Select = gSL_Select & ", htm.背番号"
    gSL_Select = gSL_Select & ", htm.サイクル"
    gSL_Select = gSL_Select & ", htm.収容数1"
    gSL_Select = gSL_Select & ", htm.収容数2"
    gSL_Select = gSL_Select & ", htm.変更日"
    gSL_Select = gSL_Select & ", trm.略称"
    gSL_Select = gSL_Select & ", kkm.回転枚数"
    Select Case InputKey
        Case "Back"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  htm.背番号 desc"
            gSL_Select = gSL_Select & ", htm.品番 desc"
            gSL_Select = gSL_Select & ", htm.取引先 desc"
            gSL_Select = gSL_Select & ", htm.受入 desc"
        Case "Next"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  htm.背番号 asc"
            gSL_Select = gSL_Select & ", htm.品番 asc"
            gSL_Select = gSL_Select & ", htm.取引先 asc"
            gSL_Select = gSL_Select & ", htm.受入 asc"
    End Select
'
End Sub
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Function BodySet(SetType As String) As Boolean
    BodySet = False

    With KRT
'       かんばん情報
        If Trim(.サイクル) <> "" Then
            B1lb_Cycle = CisFun.RSetFld(CisFun.Mid2(.サイクル, 1, 1), 1, "#") _
                 & "-" & CisFun.RSetFld(CisFun.Mid2(.サイクル, 2, 2), 2, "#") _
                 & "-" & CisFun.RSetFld(CisFun.Mid2(.サイクル, 4, 2), 2, "#")
        End If
        B1lb_KaiSu = Format(.回転枚数, "#,###")
        B1lb_Syuyo(1) = Format(.収容数1, "#,###")
        If Trim(.変更日) <> "" Then
            B1lb_SYmd = Format(.変更日, "0000/00/00")
            B1lb_Syuyo(2) = Format(.収容数2, "#,###")
        End If
'       かんばん情報
        B1lb_Maisu(5) = Format(.有効件数 + .無効件数 + .滞留件数 + .未流件数, "###")
        B1lb_Maisu(1) = Format(.有効件数, "###")
        B1lb_Maisu(2) = Format(.無効件数, "###")
        B1lb_Maisu(3) = Format(.滞留件数, "###")
        B1lb_Maisu(0) = Format(.未流件数, "###")
    End With

    If SetType = "ALL" Then
'       かんばん流動管理テーブル
        gSL_Select = ""
        gSL_Select = gSL_Select & "select"
        gSL_Select = gSL_Select & "  krt.発行NO"
        gSL_Select = gSL_Select & ", krt.流動日付"
        gSL_Select = gSL_Select & ", krt.発行区分"
        gSL_Select = gSL_Select & ", krt.有効指示日"
        gSL_Select = gSL_Select & ", krt.無効指示日"
        gSL_Select = gSL_Select & ", krt.回収日"
        gSL_Select = gSL_Select & ", krt.収容区分"
        gSL_Select = gSL_Select & ", krt.版数"
        gSL_Select = gSL_Select & " from かんばん流動管理テーブル krt"
        gSL_Select = gSL_Select & " where krt.取引先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   and krt.受入 = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & "   and krt.品番 = '" & RTrim(iH1_Hinbn) & "'"
        If CisFun.Val2(iH1_Hno) <> 0 Then
            gSL_Select = gSL_Select & "   and krt.発行NO >= " & CisFun.Val2(iH1_Hno) & ""
        End If
'       >> 流動日付
        If Trim(iH1_SRYmd) <> "" And CisFun.Val2(iH1_SRYmd.cDay) = 0 Then
        '   > 年月指定
            gSL_Select = gSL_Select & "   and substring(krt.流動日付, 1, 6) = '" & iH1_SRYmd & "'"
        Else
            '> 開始日のみ
            If Trim(iH1_SRYmd) <> "" And Trim(iH1_ERYmd) = "" Then
                gSL_Select = gSL_Select & "   and krt.流動日付 = '" & iH1_SRYmd & "'"
            End If
            '> 終了日のみ
            If Trim(iH1_SRYmd) = "" And Trim(iH1_ERYmd) <> "" Then
                gSL_Select = gSL_Select & "   and krt.流動日付 <= '" & iH1_ERYmd & "'"
            End If
            '> 開始日＆終了日
            If Trim(iH1_SRYmd) <> "" And Trim(iH1_ERYmd) <> "" Then
                gSL_Select = gSL_Select & "   and krt.流動日付 between '" & iH1_SRYmd & "'" _
                                                               & " and '" & iH1_ERYmd & "'"
            End If
        End If
        If ChkH1_HKbn(0) = 1 Or _
           ChkH1_HKbn(1) = 1 Or _
           ChkH1_HKbn(2) = 1 Or _
           ChkH1_HKbn(3) = 1 Then
            gStr = " "
            gSL_Select = gSL_Select & "   and krt.発行区分 in"
            gSL_Select = gSL_Select & " ("
            If ChkH1_HKbn(0) = 1 Then gSL_Select = gSL_Select & gStr & "1": gStr = ","
            If ChkH1_HKbn(1) = 1 Then gSL_Select = gSL_Select & gStr & "2": gStr = ","
            If ChkH1_HKbn(2) = 1 Then gSL_Select = gSL_Select & gStr & "3": gStr = ","
            If ChkH1_HKbn(3) = 1 Then gSL_Select = gSL_Select & gStr & "0": gStr = ","
            gSL_Select = gSL_Select & " )"
        End If
        gSL_Select = gSL_Select & " order by"
        gSL_Select = gSL_Select & "  krt.発行NO"
        If KRTRead(gSL_Select, , 1) Then
            Call GridSet1
        Else
            Call KRTClose
            Exit Function
        End If
        Call KRTClose
    End If

'   発注テーブル（入荷テーブル）
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  hct.納入日"
    gSL_Select = gSL_Select & ", hct.納入便"
    gSL_Select = gSL_Select & ", hct.枚数"
    gSL_Select = gSL_Select & " from 発注テーブル hct"
    gSL_Select = gSL_Select & " where hct.仕入先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and hct.受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and hct.品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   and not exists"
    gSL_Select = gSL_Select & " ("
    gSL_Select = gSL_Select & " select nyt.発注管理NO"
    gSL_Select = gSL_Select & "   from 入荷テーブル nyt"
    gSL_Select = gSL_Select & "  where nyt.発注管理NO = hct.発注管理NO"
    gSL_Select = gSL_Select & " )"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  hct.納入日"
    gSL_Select = gSL_Select & ", hct.納入便"
    If HCTRead(gSL_Select, , 1) Then
        wHCSu = 0
        Call GridSet4
        B1lb_HCSu = Format(wHCSu, "#,###")
    End If
    Call HCTClose
'   発注かんばん予定テーブル
    gSL_Select = ""
    gSL_Select = gSL_Select & "select sum(hky.枚数) 枚数"
    gSL_Select = gSL_Select & " from 発注かんばん予定テーブル hky"
    gSL_Select = gSL_Select & " where hky.仕入先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and hky.受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and hky.品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   and hky.処理区分 <> 2"
    If HKYRead(gSL_Select, , 1) Then
        B1lb_HYSu = Format(HKY.枚数, "#,###")
    End If
    Call HKYClose

    BodySet = True
End Function
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True

    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "    更新処理を行います。            "
        .MB_MSG(4) = "     よろしいですか？            "
        .MB_Title = "更新処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err

    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 更新処理 )
    Dim ID          As Long
    Call DBSet_KHT
    Call DBDelete_KHT
'   [1]
    With vsGrid1
        For ID = 1 To .Rows - 1
            If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(0))) = "" Then Exit For
            If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(6))) = "発行" Then
                If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(5))) <> "発行" Then
                    Call DBDelete_HKTNo(CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(0))))
                End If
            End If
            Select Case Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(5)))
                Case "有効": Call DBUpdate_KRT(CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(0))), 1)
                Case "無効": Call DBUpdate_KRT(CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(0))), 2)
                Case "発行": Call DBInsert_KHT(CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(0))), 0)
            End Select
        Next ID
    End With
'   [2]
    With vsGrid2
        For ID = 1 To .Rows - 1
            If Trim(.TextMatrix(ID, CisVsGrid2.FixedGet(0))) = "" Then Exit For
            If Trim(.TextMatrix(ID, CisVsGrid2.FixedGet(6))) = "発行" Then
                If Trim(.TextMatrix(ID, CisVsGrid2.FixedGet(5))) <> "発行" Then
                    Call DBDelete_HKTNo(CisFun.Val2(.TextMatrix(ID, CisVsGrid2.FixedGet(0))))
                End If
            End If
            Select Case Trim(.TextMatrix(ID, CisVsGrid2.FixedGet(5)))
                Case "有効": Call DBUpdate_KRT(CisFun.Val2(.TextMatrix(ID, CisVsGrid2.FixedGet(0))), 1)
                Case "無効": Call DBUpdate_KRT(CisFun.Val2(.TextMatrix(ID, CisVsGrid2.FixedGet(0))), 2)
                Case "発行": Call DBInsert_KHT(CisFun.Val2(.TextMatrix(ID, CisVsGrid2.FixedGet(0))), 0)
            End Select
        Next ID
    End With
'   [3]
    With vsGrid3
        For ID = 1 To .Rows - 1
            If Trim(.TextMatrix(ID, CisVsGrid3.FixedGet(0))) = "" Then Exit For
            If Trim(.TextMatrix(ID, CisVsGrid3.FixedGet(6))) = "発行" Then
                If Trim(.TextMatrix(ID, CisVsGrid3.FixedGet(5))) <> "発行" Then
                    Call DBDelete_HKTNo(CisFun.Val2(.TextMatrix(ID, CisVsGrid3.FixedGet(0))))
                End If
            End If
            Select Case Trim(.TextMatrix(ID, CisVsGrid3.FixedGet(5)))
                Case "有効": Call DBUpdate_KRT(CisFun.Val2(.TextMatrix(ID, CisVsGrid3.FixedGet(0))), 1)
                Case "無効": Call DBUpdate_KRT(CisFun.Val2(.TextMatrix(ID, CisVsGrid3.FixedGet(0))), 2)
                Case "発行": Call DBInsert_KHT(CisFun.Val2(.TextMatrix(ID, CisVsGrid3.FixedGet(0))), 0)
            End Select
        Next ID
    End With

    CisDB.DBTran (TransCommit)      ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

    Call DBSelect_KHT       ' 件数表示

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+------------------------------------+
'+      更新処理 - 指示：有効＆無効
'+------------------------------------+
Private Function DBUpdate_KRT(HNo As Integer, HKbn As Integer, Optional DBNo As Integer) As Boolean
    DBUpdate_KRT = False

    With KRT
        .取引先 = RTrim(iH1_Torcd)
        .受入 = RTrim(iH1_Ukeir)
        .品番 = RTrim(iH1_Hinbn)
        .発行NO = HNo
        .発行区分 = HKbn
        Select Case .発行区分
            Case 1
                .有効指示日 = Format(Date, "yyyymmdd")
                .無効指示日 = ""
            Case 2
                .有効指示日 = ""
                .無効指示日 = Format(Date, "yyyymmdd")
        End Select
        .回収日 = ""
        .更新者 = gTanto

        gSL_Select = ""
        gSL_Select = gSL_Select & "update かんばん流動管理テーブル"
        gSL_Select = gSL_Select & "   set 発行区分 = " & .発行区分 & ""
        gSL_Select = gSL_Select & "     , 有効指示日 = '" & RTrim(.有効指示日) & "'"
        gSL_Select = gSL_Select & "     , 無効指示日 = '" & RTrim(.無効指示日) & "'"
        gSL_Select = gSL_Select & "     , 回収日 = '" & RTrim(.回収日) & "'"
        gSL_Select = gSL_Select & "     , 更新日 = getdate()"
        gSL_Select = gSL_Select & "     , 更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & "     , 更新端末 = substring(host_name(),1,20)"
        gSL_Select = gSL_Select & " where 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   and 発行NO = " & .発行NO & ""
    End With
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With

    DBUpdate_KRT = True
End Function
'+------------------------------------+
'+      索引処理 - 指示：発行
'+------------------------------------+
Private Function DBSet_KHT()
    Dim wTSeban(4)  As String
    Dim wTorCD(4)   As String
    Dim wTorKJ(4)   As String
    
    DBSet_KHT = False

    Call ItemsClearKHT
    With KHT
        .品番 = RTrim(iH1_Hinbn)
        .取引先 = RTrim(iH1_Torcd)
        .受入 = RTrim(iH1_Ukeir)
        .自社CD = Kn_Jisya
        .自社名1 = CisFun.Mid2(Kn_JisyaNM, 1, 20)
        .自社名2 = CisFun.Mid2(Kn_JisyaKJNM, 1, 14)
        .発行区分 = 1
        .作成者 = gTanto
    End With
'   品番取引先情報(得意先情報)
    gSL_Select = "Select *"
    gSL_Select = gSL_Select & " From 品番取引先マスタ "
    gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(KHT.品番) & "'"
    gSL_Select = gSL_Select & " And   品目 = 0"
    gSL_Select = gSL_Select & " And   SK区分 = 0"
    gSL_Select = gSL_Select & " Order By 背番号 "
    If HTMRead(gSL_Select) Then
       gCnt1 = 1
       Do Until Not HTM_RDSTS
           If Trim(HTM.背番号) <> "" Then
              wTSeban(gCnt1) = Mid(Trim(HTM.背番号), 1, 3)
              wTorCD(gCnt1) = HTM.取引先
              gCnt1 = gCnt1 + 1
              If gCnt1 > 4 Then Exit Do
           End If
           Call HTMReadNext
       Loop
       For gCnt1 = 1 To 4
           If Trim(wTorCD(gCnt1)) <> "" Then
              gSL_Select = "Select * From 取引先マスタ "
              gSL_Select = gSL_Select & " Where 取引先区分 = 0"
              gSL_Select = gSL_Select & "   AND 取引先CD = '" & wTorCD(gCnt1) & "'"
              If TRMRead(gSL_Select, 1) Then
                 wTorKJ(gCnt1) = CisFun.Mid2(TRM.工場名, 1, 6)
              End If
           End If
           Select Case gCnt1
                Case 1
                    KHT.得意先背番1 = wTSeban(gCnt1)
                    KHT.得意先工場名1 = wTorKJ(gCnt1)
                Case 2
                    KHT.得意先背番2 = wTSeban(gCnt1)
                    KHT.得意先工場名2 = wTorKJ(gCnt1)
                Case 3
                    KHT.得意先背番3 = wTSeban(gCnt1)
                    KHT.得意先工場名3 = wTorKJ(gCnt1)
                Case 4
                    KHT.得意先背番4 = wTSeban(gCnt1)
                    KHT.得意先工場名4 = wTorKJ(gCnt1)
           End Select
       Next gCnt1
    End If
    Call HTMClose
'>> 品番取引先マスタ
    gSL_Select = ""
    gSL_Select = gSL_Select & "select htm.背番号"
    gSL_Select = gSL_Select & "     , htm.品目"
    gSL_Select = gSL_Select & "     , htm.手配区分"
    gSL_Select = gSL_Select & "     , htm.納入先"
    gSL_Select = gSL_Select & "     , htm.納入先受入"
    gSL_Select = gSL_Select & "     , htm.サイクル"
    gSL_Select = gSL_Select & "     , htm.収容数1"
    gSL_Select = gSL_Select & "     , htm.収容器1"
    gSL_Select = gSL_Select & "     , htm.変更日"
    gSL_Select = gSL_Select & "     , htm.収容数2"
    gSL_Select = gSL_Select & "     , htm.収容器2"
    gSL_Select = gSL_Select & "     , htm.置場"
    gSL_Select = gSL_Select & "     , htm.基準工数"
    gSL_Select = gSL_Select & "     , htm.アドレス"
    gSL_Select = gSL_Select & "     , htm.備考"
    gSL_Select = gSL_Select & "     , isnull(sym.値名称, '') 手配区分名"
    gSL_Select = gSL_Select & "  from 品番取引先マスタ htm"
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sym"
    gSL_Select = gSL_Select & "    on sym.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "   and sym.区分タイプ = ''"
    gSL_Select = gSL_Select & "   and sym.値 = htm.手配区分"
    gSL_Select = gSL_Select & " where htm.品番 = '" & KHT.品番 & "'"
    gSL_Select = gSL_Select & "   and htm.取引先 = '" & KHT.取引先 & "'"
    gSL_Select = gSL_Select & "   and htm.受入 = '" & KHT.受入 & "'"
    If HTMRead(gSL_Select, 1) Then
        KHT.背番号 = HTM.背番号
        KHT.品目 = HTM.品目
        KHT.手配区分 = HTM.手配区分
        KHT.手配区分名 = CisFun.Mid2(HTM.手配区分名, 1, 4)
        KHT.納入先 = HTM.納入先
        KHT.納入先受入 = HTM.納入先受入
        If Trim(HTM.サイクル) <> "" Then
            KHT.サイクル = CisFun.RSetFld(CisFun.Mid2(HTM.サイクル, 1, 1), 1, "#") _
                   & "-" & CisFun.RSetFld(CisFun.Mid2(HTM.サイクル, 2, 2), 2, "#") _
                   & "-" & CisFun.RSetFld(CisFun.Mid2(HTM.サイクル, 4, 2), 2, "#")
        End If
        KHT.置場 = HTM.置場
        KHT.アドレス = HTM.アドレス
        KHT.備考 = HTM.備考
        If KHT.手配区分 = 2 Then
        '   << 社内の時 >>
            KHT.基準工数 = HTM.基準工数
        End If
    End If
'>> 品番マスタ
'   ■■■■■■■■■■　品番マスタより　■■■■■■■■■■
    gSL_Select = ""
    gSL_Select = gSL_Select & "select hnm.表示品番"
    gSL_Select = gSL_Select & "     , hnm.品名"
    gSL_Select = gSL_Select & "     , hnm.工程"
    gSL_Select = gSL_Select & "     , hnm.次工程"
    gSL_Select = gSL_Select & "     , hnm.IMAGE"
    gSL_Select = gSL_Select & "     , hnm.実数区分"
    gSL_Select = gSL_Select & "     , hnm.車型"
    gSL_Select = gSL_Select & "     , hnm.担当"
    gSL_Select = gSL_Select & "  from 品番マスタ hnm"
    gSL_Select = gSL_Select & " where hnm.品番 = '" & KHT.品番 & "'"
    If HNMRead(gSL_Select, 1) Then
        KHT.表示品番 = HNM.表示品番
        KHT.品名 = HNM.品名
        KHT.IMAGE = HNM.IMAGE
        KHT.車型 = HNM.車型
        KHT.実数区分 = HNM.実数区分
        If KHT.手配区分 = 2 Then
        '   << 社内の時 >>
            KHT.工程 = HNM.工程
            KHT.次工程 = HNM.次工程
            If KHT.工程 <> "" Then
                If MeisyoGet("工程管理", KHT.工程) Then KHT.工程名 = SYM_Meisyo
            End If
            If KHT.次工程 <> "" Then
                If MeisyoGet("工程管理", KHT.次工程) Then KHT.次工程名 = SYM_Meisyo
            End If
        End If
        If mRType = 2 Then
           ' 担当者名
           Call MeisyoGet("担当管理", HNM.担当)
           KHT.区分 = CisFun.Mid2(SYM_Meisyo, 1, 4)
        End If
    End If
'>> 取引先マスタ（取引先）
    gSL_Select = ""
    gSL_Select = gSL_Select & "select trm.略称"
    gSL_Select = gSL_Select & "     , trm.工場名"
    gSL_Select = gSL_Select & "     , trm.取引先"
    gSL_Select = gSL_Select & "     , trm.工場CD"
    gSL_Select = gSL_Select & "  from 取引先マスタ trm"
    gSL_Select = gSL_Select & " where trm.取引先CD = '" & KHT.取引先 & "'"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    If TRMRead(gSL_Select, 1) Then
        KHT.取引先CD = TRM.取引先
        If Trim(TRM.工場CD) <> "" Then
           KHT.取引先CD = TRM.取引先 & "-" & TRM.工場CD
        End If
        KHT.取引先名 = TRM.略称
        KHT.取引先名1 = CisFun.Mid2(TRM.略称, 1, 10)
        KHT.取引先名2 = CisFun.Mid2(TRM.略称, 11, 10)
        KHT.取引先工場名 = TRM.工場名
        KHT.取引先工場名1 = CisFun.Mid2(TRM.工場名, 1, 10)
        KHT.取引先工場名2 = CisFun.Mid2(TRM.工場名, 11, 10)
    End If
    ' 取引先受入情報
    gSL_Select = "Select * From 取引先受入マスタ "
    gSL_Select = gSL_Select & " Where 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & KHT.取引先 & "'"
    gSL_Select = gSL_Select & "   AND 受入 = '" & KHT.受入 & "'"
    If TUMRead(gSL_Select, 1) Then
       KHT.受入名 = TUM.受入名称
    End If
'>> 取引先マスタ（納入先）
    gSL_Select = ""
    gSL_Select = gSL_Select & "select trm.略称,trm.取引先,trm.工場CD"
    gSL_Select = gSL_Select & "  from 取引先マスタ trm"
    gSL_Select = gSL_Select & " where trm.取引先CD = '" & KHT.納入先 & "'"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    If TRMRead(gSL_Select, 1) Then
        KHT.納入先CD = TRM.取引先
        If Trim(TRM.工場CD) <> "" Then
           KHT.納入先CD = TRM.取引先 & "-" & TRM.工場CD
        End If
        KHT.納入先名 = TRM.略称
        KHT.納入先名1 = CisFun.Mid2(TRM.略称, 1, 10)
        KHT.納入先名2 = CisFun.Mid2(TRM.略称, 11, 10)
    End If
    If KHT.手配区分 = "1" Then
       KHT.納入先固定 = "直送先"
    End If
    ' 納入時間情報
    gSL_Select = "Select * From 発注サイクルマスタ"
    gSL_Select = gSL_Select & " Where 仕入先 = '" & RTrim(KHT.取引先) & "'"
    gSL_Select = gSL_Select & " And   サイクル = '" & RTrim(HTM.サイクル) & "'"
    gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHT.受入) & "'"
    gSL_Select = gSL_Select & " And   行識別 = 0"
    If HCMRead(gSL_Select) Then
       gCnt1 = 1
       Do Until Not HCM_RDSTS
           If Trim(HCM.入荷時間) <> "" And HCM.入荷時間 <> "0000" Then
              gStr = Mid(HCM.入荷時間, 1, 2) & ":" & Mid(HCM.入荷時間, 3, 2)
              Select Case gCnt1
                    Case 1: KHT.納入時間1 = gStr
                    Case 2: KHT.納入時間2 = gStr
                    Case 3: KHT.納入時間3 = gStr
                    Case 4: KHT.納入時間4 = gStr
                    Case 5: KHT.納入時間5 = gStr
                    Case 6: KHT.納入時間6 = gStr
                    Case 7: KHT.納入時間7 = gStr
                    Case 8: KHT.納入時間8 = gStr
              End Select
              gCnt1 = gCnt1 + 1
              If gCnt1 > 8 Then Exit Do
           End If
           Call HCMReadNext
       Loop
    End If
    Call HCMClose
    ' 発行タイプ別
    If mRType = 1 Then
       Select Case HTM.手配区分
          Case "1": KHT.区分 = "直送"
          Case "3": KHT.区分 = "自給"
          Case Else
            KHT.区分 = "自給"
       End Select
       If Mid(KHT.取引先, 1, 4) < "7000" Then
          gSL_Select = "Select * From 名称マスタ"
          gSL_Select = gSL_Select & " Where 区分名称 = 'かんばん発行ホッパ指定'"
          gSL_Select = gSL_Select & " And   区分タイプ = ''"
          gSL_Select = gSL_Select & " And   数字3 = 1"
          gSL_Select = gSL_Select & " Order By 値"
          If Not SYMRead(gSL_Select, 1) Then
             SYM.数字1 = 1: SYM.数字2 = 3
          End If
          KHT.発行種類 = 10
          KHT.StHNo = SYM.数字1 - 1
          KHT.EdHNo = SYM.数字2 - 1
          KHT.Stack1 = KHT.StHNo
          KHT.Stack2 = KHT.StHNo
       Else
          gSL_Select = "Select * From 名称マスタ"
          gSL_Select = gSL_Select & " Where 区分名称 = 'かんばん発行ホッパ指定'"
          gSL_Select = gSL_Select & " And   区分タイプ = ''"
          gSL_Select = gSL_Select & " And   数字3 = 2"
          gSL_Select = gSL_Select & " Order By 値"
          If Not SYMRead(gSL_Select, 1) Then
             SYM.数字1 = 4: SYM.数字2 = 4
          End If
          KHT.発行種類 = 11
          KHT.StHNo = SYM.数字1 - 1
          KHT.EdHNo = SYM.数字2 - 1
          KHT.Stack1 = KHT.StHNo
          KHT.Stack2 = KHT.StHNo
       End If
    End If
    DBSet_KHT = True
End Function
'+------------------------------------+
'+      作成処理 - 指示：発行
'+------------------------------------+
Private Function DBInsert_KHT(HNo As Integer, HKbn As Integer, Optional DBNo As Integer) As Boolean
    DBInsert_KHT = False

    With KHT
        .発行NO = HNo
        .発行区分 = HKbn
    End With
'   かんばん流動管理テーブルの参照
    gSL_Select = ""
    gSL_Select = gSL_Select & "select krt.版数"
    gSL_Select = gSL_Select & "     , krt.収容区分"
    gSL_Select = gSL_Select & "  from かんばん流動管理テーブル krt"
    gSL_Select = gSL_Select & " where krt.取引先 = '" & RTrim(KHT.取引先) & "'"
    gSL_Select = gSL_Select & "   and krt.受入 = '" & RTrim(KHT.受入) & "'"
    gSL_Select = gSL_Select & "   and krt.品番 = '" & RTrim(KHT.品番) & "'"
    gSL_Select = gSL_Select & "   and krt.発行NO = " & KHT.発行NO & ""
    gSL_Select = gSL_Select & "   and not exists"
    gSL_Select = gSL_Select & " ("
    gSL_Select = gSL_Select & " select kht.版数"
    gSL_Select = gSL_Select & "   from かんばん発行テーブル kht"
    gSL_Select = gSL_Select & "  where kht.かんばん種類 = " & KHT.かんばん種類 & ""
    gSL_Select = gSL_Select & "    and kht.背番号 = '" & RTrim(KHT.背番号) & "'"
    gSL_Select = gSL_Select & "    and kht.品番 = '" & RTrim(KHT.品番) & "'"
    gSL_Select = gSL_Select & "    and kht.取引先 = '" & RTrim(KHT.取引先) & "'"
    gSL_Select = gSL_Select & "    and kht.受入 = '" & RTrim(KHT.受入) & "'"
    gSL_Select = gSL_Select & "    and kht.発行NO = " & KHT.発行NO & ""
    gSL_Select = gSL_Select & "    and kht.発行済 = 0"
    gSL_Select = gSL_Select & " )"
    If Not KRTRead(gSL_Select, 1) Then
       With CisDB
            .SQL = "DELETE かんばん発行テーブル"
            .SQL = .SQL & "  where かんばん種類 = " & KHT.かんばん種類 & ""
            .SQL = .SQL & "    and 背番号 = '" & RTrim(KHT.背番号) & "'"
            .SQL = .SQL & "    and 品番 = '" & RTrim(KHT.品番) & "'"
            .SQL = .SQL & "    and 取引先 = '" & RTrim(KHT.取引先) & "'"
            .SQL = .SQL & "    and 受入 = '" & RTrim(KHT.受入) & "'"
            .SQL = .SQL & "    and 発行NO = " & KHT.発行NO & ""
            If Not .DBExec Then Exit Function
       End With
    End If
'   かんばん発行テーブルの作成
    With KHT
        .版数 = KRT.版数
        .収容区分 = KRT.収容区分
        Select Case KRT.収容区分
            Case 1
                .収容器 = HTM.収容器1
                .収容数 = HTM.収容数1
            Case 2:
                .収容器 = HTM.収容器2
                .収容数 = HTM.収容数2
        End Select
    End With
    Call KHTInsert

    DBInsert_KHT = True
End Function
'+------------------------------------+
'+      削除処理 - 指示：発行
'+------------------------------------+
Private Function DBDelete_KHT(Optional DBNo As Integer) As Boolean
    DBDelete_KHT = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "delete かんばん発行テーブル"
    gSL_Select = gSL_Select & " where かんばん種類 = " & KHT.かんばん種類 & ""
    gSL_Select = gSL_Select & "   and 発行済 = 1"
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With

    DBDelete_KHT = True
End Function

Private Function DBDelete_HKTNo(rno As Integer) As Boolean
    DBDelete_HKTNo = False
    With CisDB
         .SQL = "DELETE かんばん発行テーブル"
         .SQL = .SQL & "  where かんばん種類 = " & KHT.かんばん種類 & ""
         .SQL = .SQL & "    and 背番号 = '" & RTrim(KHT.背番号) & "'"
         .SQL = .SQL & "    and 品番 = '" & RTrim(KHT.品番) & "'"
         .SQL = .SQL & "    and 取引先 = '" & RTrim(KHT.取引先) & "'"
         .SQL = .SQL & "    and 受入 = '" & RTrim(KHT.受入) & "'"
         .SQL = .SQL & "    and 発行NO = " & rno & ""
         If Not .DBExec Then Exit Function
    End With
    DBDelete_HKTNo = True
End Function
'+------------------------------------+
'+      表示処理 - 指示：発行
'+------------------------------------+
Private Function DBSelect_KHT(Optional DBNo As Integer) As Boolean
    DBSelect_KHT = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select count(*) 件数"
    gSL_Select = gSL_Select & "  from かんばん発行テーブル"
    gSL_Select = gSL_Select & "  where かんばん種類 = 0"
    gSL_Select = gSL_Select & "    and 発行済 = 0"
    If KHTRead(gSL_Select, 1) Then
        H1lb_Prtsu = Format(KHT.件数, "#,###")
    End If

    DBSelect_KHT = True
End Function



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
Private Sub GridInit()
'[1]
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 中中, あり, 10, 1, "流動日付")
        Call .FixedSet(中中, 中中, あり, 6, 2, "発行")
        Call .FixedSet(中中, 中中, あり, 10, 3, "無効指示日")
        Call .FixedSet(中中, 中中, あり, 6, 4, "回収日")
        Call .FixedSet(中中, 中中, あり, 4, 5, "指示")
        Call .FixedSet(中中, 右中, あり, 3, 6, "発行№")
        Call .FixedSet(中中, 右中, あり, 2, 7, "版")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 8, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,5,7")
        Call .PatanSet(1, "0,3,4,5,7")
        .InitDisp
    End With
'[2]
    With CisVsGrid2
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 中中, あり, 10, 1, "流動日付")
        Call .FixedSet(中中, 中中, あり, 6, 2, "発行")
        Call .FixedSet(中中, 中中, あり, 10, 3, "無効指示日")
        Call .FixedSet(中中, 中中, あり, 6, 4, "回収日")
        Call .FixedSet(中中, 中中, あり, 4, 5, "指示")
        Call .FixedSet(中中, 右中, あり, 3, 6, "発行№")
        Call .FixedSet(中中, 右中, あり, 2, 7, "版")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 8, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,5,7")
        Call .PatanSet(1, "0,3,4,5,7")
        .InitDisp
    End With
'[3]
    With CisVsGrid3
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 中中, あり, 10, 1, "流動日付")
        Call .FixedSet(中中, 中中, あり, 6, 2, "発行")
        Call .FixedSet(中中, 中中, あり, 10, 3, "無効指示日")
        Call .FixedSet(中中, 中中, あり, 6, 4, "回収日")
        Call .FixedSet(中中, 中中, あり, 4, 5, "指示")
        Call .FixedSet(中中, 右中, あり, 3, 6, "発行№")
        Call .FixedSet(中中, 右中, あり, 2, 7, "版")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 8, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,5,7")
        Call .PatanSet(1, "0,3,4,5,7")
        .InitDisp
    End With
End Sub
'***********************************************'
'*****      グリッド内容セット - 流動日付
'***********************************************'
Private Sub GridSet1()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    If Ttl_Kensu > 60 Then
        With vsGrid3
           .Redraw = False
            If Ttl_Kensu + 1 < CisVsGrid3.DispMax Then
                .Rows = CisVsGrid3.DispMax
            Else
                .Rows = Ttl_Kensu + 1 - 40
            End If
        End With
    End If
    ID = 0
    Do Until Not KRT_RDSTS
        ID = ID + 1

        Select Case ID
            Case 1 To 20:  Call KrtInfoToGrid1(ID)
            Case 21 To 40: Call KrtInfoToGrid2(ID - 20)
            Case Else:     Call KrtInfoToGrid3(ID - 40)
        End Select

        Call KRTReadNext(1)
    Loop
    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    With vsGrid2
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    With vsGrid3
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'+----------------------------------------------+
'+  かんばん流動管理テーブルよりグリッドに転送
'+----------------------------------------------+
'[1]
Private Sub KrtInfoToGrid1(SetRow As Long)
    With vsGrid1
       'No
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(0)) = KRT.発行NO
       '流動日付
        If RTrim(KRT.流動日付) <> "" Then
           .TextMatrix(SetRow, CisVsGrid1.FixedGet(1)) = Format(KRT.流動日付, "0000/00/00")
        End If
       '発行
        Select Case KRT.発行区分
            Case 0:    .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = "未流動"
            Case 1:    .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = "有効"
            Case 2:    .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = "無効"
            Case 3:    .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = "滞留"
            Case Else: .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = ""
        End Select
       '無効指示日
        If RTrim(KRT.無効指示日) <> "" Then
           .TextMatrix(SetRow, CisVsGrid1.FixedGet(3)) = Format(KRT.無効指示日, "0000/00/00")
        End If
       '回収日
        If RTrim(KRT.回収日) <> "" Then
           .TextMatrix(SetRow, CisVsGrid1.FixedGet(4)) = CisFun.Mid2(Format(KRT.回収日, "0000/00/00"), 6, 5)
        End If
       '回収日
        If CheckKHT Then
           .TextMatrix(SetRow, CisVsGrid1.FixedGet(5)) = "発行"
           .TextMatrix(SetRow, CisVsGrid1.FixedGet(6)) = "発行"
        End If
       '--------------------------------------------------------------- 2006/12/20 追加 (Start)
       '版数
        If KRT.版数 <> 0 Then
           .TextMatrix(SetRow, CisVsGrid1.FixedGet(7)) = KRT.版数
        End If
       '--------------------------------------------------------------- 2006/12/20 追加 (End)
       '<><><> 背景色のｾｯﾄ <><><>
        Select Case KRT.発行区分
            Case 0 To 3
                .Row = SetRow
                For gInt = 1 To 5
                    .Col = gInt
                    .CellBackColor = Hkbn_BackColor(KRT.発行区分)
                Next gInt
        End Select
        Select Case KRT.収容区分
            Case 1 To 2
                .Row = SetRow
                .Col = 0
                .CellBackColor = Syuyo_BackColor(KRT.収容区分)
        End Select
    End With
End Sub
'[2]
Private Sub KrtInfoToGrid2(SetRow As Long)
    With vsGrid2
       'No
        .TextMatrix(SetRow, CisVsGrid2.FixedGet(0)) = KRT.発行NO
       '流動日付
        If RTrim(KRT.流動日付) <> "" Then
           .TextMatrix(SetRow, CisVsGrid2.FixedGet(1)) = Format(KRT.流動日付, "0000/00/00")
        End If
       '発行
        Select Case KRT.発行区分
            Case 0:    .TextMatrix(SetRow, CisVsGrid2.FixedGet(2)) = "未流動"
            Case 1:    .TextMatrix(SetRow, CisVsGrid2.FixedGet(2)) = "有効"
            Case 2:    .TextMatrix(SetRow, CisVsGrid2.FixedGet(2)) = "無効"
            Case 3:    .TextMatrix(SetRow, CisVsGrid2.FixedGet(2)) = "滞留"
            Case Else: .TextMatrix(SetRow, CisVsGrid2.FixedGet(2)) = ""
        End Select
       '無効指示日
        If RTrim(KRT.無効指示日) <> "" Then
           .TextMatrix(SetRow, CisVsGrid2.FixedGet(3)) = Format(KRT.無効指示日, "0000/00/00")
        End If
       '回収日
        If RTrim(KRT.回収日) <> "" Then
           .TextMatrix(SetRow, CisVsGrid2.FixedGet(4)) = CisFun.Mid2(Format(KRT.回収日, "0000/00/00"), 6, 5)
        End If
       '回収日
        If CheckKHT Then
           .TextMatrix(SetRow, CisVsGrid2.FixedGet(5)) = "発行"
           .TextMatrix(SetRow, CisVsGrid2.FixedGet(6)) = "発行"
        End If
       '--------------------------------------------------------------- 2006/12/20 追加 (Start)
       '版数
        If KRT.版数 <> 0 Then
           .TextMatrix(SetRow, CisVsGrid1.FixedGet(7)) = KRT.版数
        End If
       '--------------------------------------------------------------- 2006/12/20 追加 (End)
       '<><><> 背景色のｾｯﾄ <><><>
        Select Case KRT.発行区分
            Case 0 To 3
                .Row = SetRow
                For gInt = 1 To 5
                    .Col = gInt
                    .CellBackColor = Hkbn_BackColor(KRT.発行区分)
                Next gInt
        End Select
    '------------------------------------------- 2005/12/21 Insert
        Select Case KRT.収容区分
            Case 1 To 2
                .Row = SetRow
                .Col = 0
                .CellBackColor = Syuyo_BackColor(KRT.収容区分)
        End Select
    '------------------------------------------- 2005/12/21 Insert End
    End With
End Sub
'[3]
Private Sub KrtInfoToGrid3(SetRow As Long)
    With vsGrid3
       'No
        .TextMatrix(SetRow, CisVsGrid3.FixedGet(0)) = KRT.発行NO
       '流動日付
        If RTrim(KRT.流動日付) <> "" Then
           .TextMatrix(SetRow, CisVsGrid3.FixedGet(1)) = Format(KRT.流動日付, "0000/00/00")
        End If
       '発行
        Select Case KRT.発行区分
            Case 0:    .TextMatrix(SetRow, CisVsGrid3.FixedGet(2)) = "未流動"
            Case 1:    .TextMatrix(SetRow, CisVsGrid3.FixedGet(2)) = "有効"
            Case 2:    .TextMatrix(SetRow, CisVsGrid3.FixedGet(2)) = "無効"
            Case 3:    .TextMatrix(SetRow, CisVsGrid3.FixedGet(2)) = "滞留"
            Case Else: .TextMatrix(SetRow, CisVsGrid3.FixedGet(2)) = ""
        End Select
       '無効指示日
        If RTrim(KRT.無効指示日) <> "" Then
           .TextMatrix(SetRow, CisVsGrid3.FixedGet(3)) = Format(KRT.無効指示日, "0000/00/00")
        End If
       '回収日
        If RTrim(KRT.回収日) <> "" Then
           .TextMatrix(SetRow, CisVsGrid3.FixedGet(4)) = CisFun.Mid2(Format(KRT.回収日, "0000/00/00"), 6, 5)
        End If
       '回収日
        If CheckKHT Then
           .TextMatrix(SetRow, CisVsGrid3.FixedGet(5)) = "発行"
           .TextMatrix(SetRow, CisVsGrid3.FixedGet(6)) = "発行"
        End If
       '--------------------------------------------------------------- 2006/12/20 追加 (Start)
       '版数
        If KRT.版数 <> 0 Then
           .TextMatrix(SetRow, CisVsGrid1.FixedGet(7)) = KRT.版数
        End If
       '--------------------------------------------------------------- 2006/12/20 追加 (End)
       '<><><> 背景色のｾｯﾄ <><><>
        Select Case KRT.発行区分
            Case 0 To 3
                .Row = SetRow
                For gInt = 1 To 5
                    .Col = gInt
                    .CellBackColor = Hkbn_BackColor(KRT.発行区分)
                Next gInt
        End Select
    '------------------------------------------- 2005/12/21 Insert
        Select Case KRT.収容区分
            Case 1 To 2
                .Row = SetRow
                .Col = 0
                .CellBackColor = Syuyo_BackColor(KRT.収容区分)
        End Select
    '------------------------------------------- 2005/12/21 Insert End
    End With
End Sub
'+----------------------------------------------+
'+  かんばん発行テーブルの存在ﾁｪｯｸ
'+----------------------------------------------+
Private Function CheckKHT() As Boolean
    gSL_Select = ""
    gSL_Select = gSL_Select & "select kht.発行NO"
    gSL_Select = gSL_Select & "   from かんばん発行テーブル kht"
    gSL_Select = gSL_Select & "  where kht.かんばん種類 = 0"
    gSL_Select = gSL_Select & "    and kht.背番号 = '" & RTrim(iH1_Seban) & "'"
    gSL_Select = gSL_Select & "    and kht.品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "    and kht.取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "    and kht.受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "    and kht.発行NO = " & KRT.発行NO & ""
    gSL_Select = gSL_Select & "    and kht.発行済 = 0"
    If TRMRead(gSL_Select, 1) Then
        CheckKHT = True
    Else
        CheckKHT = False
    End If
End Function
'
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ - 発注情報
'+----------------------------------------------------------+
Private Sub GridInit4()
    With CisVsGrid4
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 2, 0, "№")
        Call .FixedSet(中中, 中中, あり, 10, 1, "納 入 日")
        Call .FixedSet(中中, 中中, あり, 2, 2, "便")
        Call .FixedSet(中中, 右中, あり, 4, 3, "枚数")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
       'Call .FixedSet(中中, 左中, なし, 0, 4, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3")
        .InitDisp
    End With
End Sub
'***********************************************'
'*****      グリッド内容セット - 発注情報
'***********************************************'
Private Sub GridSet4()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid4
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid4.DispMax Then
            .Rows = CisVsGrid4.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not HCT_RDSTS
        ID = ID + 1

        With vsGrid4
            .TextMatrix(ID, CisVsGrid4.FixedGet(0)) = ID
            .TextMatrix(ID, CisVsGrid4.FixedGet(1)) = Format(HCT.納入日, "0000/00/00")
            .TextMatrix(ID, CisVsGrid4.FixedGet(2)) = HCT.納入便
            .TextMatrix(ID, CisVsGrid4.FixedGet(3)) = Format(HCT.枚数, "###")
            wHCSu = wHCSu + HCT.枚数
        End With

        Call HCTReadNext(1)
    Loop
    With vsGrid4
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'
'+---------------------------------------+
'+  グリッド (MouseDown)
'+---------------------------------------+
'[1]
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RowNoS      As Long
    Dim RowNoE      As Long
    Dim RowNo       As Long

    If SyoriKB = "MNT" Then
        With vsGrid1
            PopupSMenu.Visible = False

            If .MouseRow = 0 Then Exit Sub

            If Button = vbRightButton Then
'*----------------------------------------------------------------------* << 2004/07/09 >> Insert Start
                .Row = .MouseRow
'*----------------------------------------------------------------------* << 2004/07/09 >> Insert End
                If .RowSel > .Row Then
                    RowNoS = .Row
                    RowNoE = .RowSel
                Else
                    RowNoS = .RowSel
                    RowNoE = .Row
                End If
                If .TextMatrix(RowNoS, CisVsGrid1.FixedGet(0)) = "" Then Exit Sub
                If .TextMatrix(RowNoE, CisVsGrid1.FixedGet(0)) = "" Then
                    For RowNo = RowNoE To RowNoS Step -1
                        If .TextMatrix(RowNo, CisVsGrid1.FixedGet(0)) <> "" Then Exit For
                        RowNoE = RowNoE - 1
                    Next RowNo
                    .Row = RowNoS
                    .RowSel = RowNoE
                End If

                PopupSMenu.Left = B1_Area1.Left + 2160 + .Left
'*----------------------------------------------------------------------* << 2004/07/09 >> Update Start
               'PopupSMenu.Top = B1_Area1.Top + 540 + RowNoE * .CellHeight
                Select Case RowNoE
                    Case 1 To 16: PopupSMenu.Top = B1_Area1.Top + RowNoE * .CellHeight + 540
                    Case Else:    PopupSMenu.Top = B1_Area1.Top + RowNoS * .CellHeight - 960
                End Select
'*----------------------------------------------------------------------* << 2004/07/09 >> Update End
                PopupSMenu.Visible = True
                PopupSMenu.Tag = 1
                .SetFocus
            End If
        End With
    End If
End Sub
'[2]
Private Sub VSGrid2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RowNoS      As Long
    Dim RowNoE      As Long
    Dim RowNo       As Long

    If SyoriKB = "MNT" Then
        With vsGrid2
            PopupSMenu.Visible = False

            If .MouseRow = 0 Then Exit Sub

            If Button = vbRightButton Then
'*----------------------------------------------------------------------* << 2004/07/09 >> Insert Start
                .Row = .MouseRow
'*----------------------------------------------------------------------* << 2004/07/09 >> Insert End
                If .RowSel > .Row Then
                    RowNoS = .Row
                    RowNoE = .RowSel
                Else
                    RowNoS = .RowSel
                    RowNoE = .Row
                End If
                If .TextMatrix(RowNoS, CisVsGrid2.FixedGet(0)) = "" Then Exit Sub
                If .TextMatrix(RowNoE, CisVsGrid2.FixedGet(0)) = "" Then
                    For RowNo = RowNoE To RowNoS Step -1
                        If .TextMatrix(RowNo, CisVsGrid2.FixedGet(0)) <> "" Then Exit For
                        RowNoE = RowNoE - 1
                    Next RowNo
                    .Row = RowNoS
                    .RowSel = RowNoE
                End If

                PopupSMenu.Left = B1_Area1.Left + 2160 + .Left
'*----------------------------------------------------------------------* << 2004/07/09 >> Update Start
               'PopupSMenu.Top = B1_Area1.Top + 540 + RowNoE * .CellHeight
                Select Case RowNoE
                    Case 1 To 16: PopupSMenu.Top = B1_Area1.Top + RowNoE * .CellHeight + 540
                    Case Else:    PopupSMenu.Top = B1_Area1.Top + RowNoS * .CellHeight - 960
                End Select
'*----------------------------------------------------------------------* << 2004/07/09 >> Update End
                PopupSMenu.Visible = True
                PopupSMenu.Tag = 2
                .SetFocus
            End If
        End With
    End If
End Sub
'[3]
Private Sub VSGrid3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RowNoS      As Long
    Dim RowNoE      As Long
    Dim RowNo       As Long

    If SyoriKB = "MNT" Then
        With vsGrid3
            PopupSMenu.Visible = False

            If .MouseRow = 0 Then Exit Sub

            If Button = vbRightButton Then
'*----------------------------------------------------------------------* << 2004/07/09 >> Insert Start
                .Row = .MouseRow
'*----------------------------------------------------------------------* << 2004/07/09 >> Insert End
                If .RowSel > .Row Then
                    RowNoS = .Row
                    RowNoE = .RowSel
                Else
                    RowNoS = .RowSel
                    RowNoE = .Row
                End If
                If .TextMatrix(RowNoS, CisVsGrid3.FixedGet(0)) = "" Then Exit Sub
                If .TextMatrix(RowNoE, CisVsGrid3.FixedGet(0)) = "" Then
                    For RowNo = RowNoE To RowNoS Step -1
                        If .TextMatrix(RowNo, CisVsGrid3.FixedGet(0)) <> "" Then Exit For
                        RowNoE = RowNoE - 1
                    Next RowNo
                    .Row = RowNoS
                    .RowSel = RowNoE
                End If

                PopupSMenu.Left = B1_Area1.Left + 2160 + .Left
'*----------------------------------------------------------------------* << 2004/07/09 >> Update Start
               'PopupSMenu.Top = B1_Area1.Top + 540 + .MouseRow * .CellHeight
                RowNoS = RowNoS - .TopRow + 1: If RowNoS < 1 Then RowNoS = 1
                RowNoE = RowNoE - .TopRow + 1
                Select Case RowNoE
                    Case 1 To 16: PopupSMenu.Top = B1_Area1.Top + RowNoE * .CellHeight + 540
                    Case Else:    PopupSMenu.Top = B1_Area1.Top + RowNoS * .CellHeight - 960
                End Select
'*----------------------------------------------------------------------* << 2004/07/09 >> Update End
                PopupSMenu.Visible = True
                PopupSMenu.Tag = 3
                .SetFocus
            End If
        End With
    End If
End Sub
'+---------------------------------------+
'+  POP_Menu (Click)
'+---------------------------------------+
Private Sub POP_Menu_Click(Index As Integer)
    Dim RowNoS      As Long
    Dim RowNoE      As Long
    Dim RowNo       As Long
    Dim MsgText     As String

    Select Case Index
        Case 0: MsgText = "有効"
        Case 1: MsgText = "無効"
        Case 2: MsgText = "発行"
        Case 3: MsgText = ""
    End Select
    Select Case PopupSMenu.Tag
        Case 1
            With vsGrid1
                If .RowSel > .Row Then
                    RowNoS = .Row
                    RowNoE = .RowSel
                Else
                    RowNoS = .RowSel
                    RowNoE = .Row
                End If
                For RowNo = RowNoS To RowNoE
                    If .TextMatrix(RowNo, CisVsGrid1.FixedGet(0)) <> "" Then
                        If .TextMatrix(RowNo, CisVsGrid1.FixedGet(2)) <> MsgText Then
                            .TextMatrix(RowNo, CisVsGrid1.FixedGet(5)) = MsgText
                        Else
                            .TextMatrix(RowNo, CisVsGrid1.FixedGet(5)) = ""
                        End If
                    End If
                Next RowNo
            End With
        Case 2
            With vsGrid2
                If .RowSel > .Row Then
                    RowNoS = .Row
                    RowNoE = .RowSel
                Else
                    RowNoS = .RowSel
                    RowNoE = .Row
                End If
                For RowNo = RowNoS To RowNoE
                    If .TextMatrix(RowNo, CisVsGrid2.FixedGet(0)) <> "" Then
                        If .TextMatrix(RowNo, CisVsGrid2.FixedGet(2)) <> MsgText Then
                            .TextMatrix(RowNo, CisVsGrid2.FixedGet(5)) = MsgText
                        Else
                            .TextMatrix(RowNo, CisVsGrid2.FixedGet(5)) = ""
                        End If
                    End If
                Next RowNo
            End With
        Case 3
            With vsGrid3
                If .RowSel > .Row Then
                    RowNoS = .Row
                    RowNoE = .RowSel
                Else
                    RowNoS = .RowSel
                    RowNoE = .Row
                End If
                For RowNo = RowNoS To RowNoE
                    If .TextMatrix(RowNo, CisVsGrid3.FixedGet(0)) <> "" Then
                        If .TextMatrix(RowNo, CisVsGrid3.FixedGet(2)) <> MsgText Then
                            .TextMatrix(RowNo, CisVsGrid3.FixedGet(5)) = MsgText
                        Else
                            .TextMatrix(RowNo, CisVsGrid3.FixedGet(5)) = ""
                        End If
                    End If
                Next RowNo
            End With
    End Select
    PopupSMenu.Visible = False
End Sub

