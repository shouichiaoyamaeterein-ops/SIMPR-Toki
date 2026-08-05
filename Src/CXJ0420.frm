VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CXJ0420 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "受注計画状況照会"
   ClientHeight    =   11190
   ClientLeft      =   195
   ClientTop       =   375
   ClientWidth     =   15360
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
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   5340
      Top             =   4380
      Visible         =   0   'False
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   1640
      BackColor       =   16761024
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
      cBoderWidth     =   5
      Begin Cis3D_v60.CIS3D Message 
         Height          =   525
         Left            =   240
         Top             =   210
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   926
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "【抽出中】しばらくお待ち下さい"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         cBoderWidth     =   2
         cAlingnment     =   7
         cBoderStyle     =   1
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
      Height          =   2580
      Left            =   105
      ScaleHeight     =   2520
      ScaleWidth      =   4395
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1095
      Width           =   4455
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   795
         Left            =   2385
         Top             =   1665
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "自社工区"
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
         Begin VB.ComboBox H1Comb_Koku 
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
            Left            =   330
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   15
            Top             =   360
            Width           =   1335
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   795
         Left            =   3855
         Top             =   870
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "受"
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
            Height          =   375
            Left            =   60
            TabIndex        =   3
            Top             =   345
            Width           =   375
            _ExtentX        =   661
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
            Text            =   "XX"
            MaxLength       =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   795
         Left            =   30
         Top             =   1665
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "指 定 日"
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
            Left            =   285
            TabIndex        =   4
            Top             =   360
            Width           =   1800
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
            Object.Width           =   1800
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
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   795
         Left            =   1125
         Top             =   75
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "品   番"
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
            Height          =   375
            Left            =   15
            TabIndex        =   1
            Top             =   360
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   661
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
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   795
         Left            =   30
         Top             =   870
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "納  入  先"
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
            Height          =   375
            Left            =   60
            TabIndex        =   2
            Top             =   360
            Width           =   990
            _ExtentX        =   1746
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1050
            Top             =   360
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   635
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
      Begin Cis3D_v60.CIS3D CIS3D38 
         Height          =   795
         Left            =   30
         Top             =   75
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1402
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
         Begin CisText_V60.CisText iH1_Seban 
            Height          =   360
            Left            =   15
            TabIndex        =   0
            Top             =   360
            Width           =   1065
            _ExtentX        =   1879
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
            Text            =   "XXXXXXXX"
            MaxLength       =   8
            IMEMode         =   2
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   9660
      Left            =   4635
      ScaleHeight     =   9600
      ScaleWidth      =   10440
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1095
      Width           =   10500
      Begin vsFlexLib.vsFlexArray vsGrid 
         Height          =   9525
         Left            =   75
         TabIndex        =   5
         Top             =   45
         Width           =   10290
         _Version        =   196608
         _ExtentX        =   18150
         _ExtentY        =   16801
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "日付    | |１日      |２日      |３日      |４日       |５日    |６日    |７日      |"
         Rows            =   30
         Cols            =   11
         BackColor       =   16777152
         BackColorFixed  =   16777152
         BackColorBkg    =   12632256
         TextStyle       =   4
         TextStyleFixed  =   3
         FocusRect       =   2
         HighLight       =   0
         ScrollBars      =   0
         RowHeightMin    =   320
         FillStyle       =   1
         GridColor       =   14737632
         SelectionMode   =   2
         BackColorAlternate=   16777152
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
      Height          =   660
      Left            =   10575
      ScaleHeight     =   600
      ScaleWidth      =   4500
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   390
      Width           =   4560
      Begin CisBtn_60.CisBtn PB_Chg 
         Height          =   585
         Left            =   945
         Top             =   0
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
         BtnType         =   10
         BtnCaption      =   "切替"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8388736
         BtnShadowColor  =   16777215
         PFCaption       =   "( F6 )"
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
         PFBackColor     =   8388736
      End
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1830
         Top             =   0
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
         Left            =   2715
         Top             =   0
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
         Left            =   3600
         Top             =   0
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
         Left            =   60
         Top             =   0
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
      TabIndex        =   6
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
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 受注計画状況照会 】"
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
         Left            =   14430
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   11910
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   7035
      Left            =   105
      ScaleHeight     =   6975
      ScaleWidth      =   4395
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3720
      Width           =   4455
      Begin Cis3D_v60.CIS3D CIS3D19 
         Height          =   1200
         Left            =   2730
         Top             =   4305
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   2117
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
         Begin Cis3D_v60.CIS3D B1lb_SSuryo 
            Height          =   315
            Index           =   4
            Left            =   170
            Top             =   435
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   16777215
            Caption         =   "##,###,###"
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
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   1200
         Left            =   2730
         Top             =   3105
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   2117
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
         Begin Cis3D_v60.CIS3D B1lb_SSuryo 
            Height          =   315
            Index           =   3
            Left            =   170
            Top             =   435
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   12640511
            Caption         =   "##,###,###"
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
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   1200
         Left            =   2730
         Top             =   1905
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   2117
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
         Begin Cis3D_v60.CIS3D B1lb_SSuryo 
            Height          =   315
            Index           =   2
            Left            =   170
            Top             =   435
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   12648447
            Caption         =   "##,###,###"
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
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   1200
         Left            =   1125
         Top             =   705
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   2117
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
         Begin Cis3D_v60.CIS3D B1lb_TSyoyo 
            Height          =   315
            Index           =   1
            Left            =   165
            Top             =   435
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   12648384
            Caption         =   "##,###,###"
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
      Begin Cis3D_v60.CIS3D Grid_RowHead 
         Height          =   1200
         Index           =   1
         Left            =   45
         Tag             =   "内示数"
         Top             =   705
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   2117
         ForeColor       =   16711680
         Caption         =   "内示数"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Grid_RowHead 
         Height          =   1200
         Index           =   2
         Left            =   45
         Tag             =   "受注数"
         Top             =   1905
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   2117
         ForeColor       =   16711680
         Caption         =   "受注数"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Grid_RowHead 
         Height          =   1200
         Index           =   3
         Left            =   45
         Tag             =   "出荷数"
         Top             =   3105
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   2117
         ForeColor       =   16711680
         Caption         =   "出荷数"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Grid_RowHead 
         Height          =   1200
         Index           =   4
         Left            =   45
         Tag             =   "検収数"
         Top             =   4305
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   2117
         ForeColor       =   16711680
         Caption         =   "検収数"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D BLB_D 
         Height          =   2010
         Index           =   17
         Left            =   5415
         Top             =   360
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   3545
         BackColor       =   0
         ForeColor       =   16744703
         Caption         =   "品番情報"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   300
      End
      Begin Cis3D_v60.CIS3D BLB_D 
         Height          =   660
         Index           =   18
         Left            =   1110
         Top             =   30
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   1164
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "当月合計"
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
         cFont3D         =   1
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   1200
         Left            =   2730
         Top             =   705
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   2117
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
         Begin Cis3D_v60.CIS3D B1lb_SSuryo 
            Height          =   315
            Index           =   1
            Left            =   170
            Top             =   435
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   12648384
            Caption         =   "##,###,###"
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
      Begin Cis3D_v60.CIS3D BLB_D 
         Height          =   660
         Index           =   0
         Left            =   2730
         Top             =   30
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   1164
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "指定日合計"
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
         cFont3D         =   1
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   1200
         Left            =   1125
         Top             =   1905
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   2117
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
         Begin Cis3D_v60.CIS3D B1lb_TSyoyo 
            Height          =   315
            Index           =   2
            Left            =   170
            Top             =   435
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   12648447
            Caption         =   "##,###,###"
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
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   1200
         Left            =   1125
         Top             =   3105
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   2117
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
         Begin Cis3D_v60.CIS3D B1lb_TSyoyo 
            Height          =   315
            Index           =   3
            Left            =   170
            Top             =   435
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   12640511
            Caption         =   "##,###,###"
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
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   1200
         Left            =   1125
         Top             =   4305
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   2117
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
         Begin Cis3D_v60.CIS3D B1lb_TSyoyo 
            Height          =   315
            Index           =   4
            Left            =   170
            Top             =   435
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   16777215
            Caption         =   "##,###,###"
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
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   660
         Left            =   45
         Top             =   30
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   1164
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
      End
      Begin Cis3D_v60.CIS3D Grid_RowHead 
         Height          =   570
         Index           =   0
         Left            =   45
         Top             =   5505
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   1005
         ForeColor       =   16711680
         Caption         =   "前月内示数"
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
         cAlingnment     =   1
         cPositionX      =   500
      End
      Begin Cis3D_v60.CIS3D CIS3D20 
         Height          =   840
         Left            =   45
         Top             =   6075
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   1482
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
         Begin Cis3D_v60.CIS3D B1lb_YSyoyo 
            Height          =   315
            Left            =   450
            Top             =   270
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   16777152
            Caption         =   "##,###,###"
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
      Begin Cis3D_v60.CIS3D Grid_RowHead 
         Height          =   570
         Index           =   6
         Left            =   2205
         Top             =   5505
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   1005
         ForeColor       =   16711680
         Caption         =   "翌月内示数"
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
         cAlingnment     =   1
         cPositionX      =   500
      End
      Begin Cis3D_v60.CIS3D CIS3D39 
         Height          =   840
         Left            =   2205
         Top             =   6075
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   1482
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
         Begin Cis3D_v60.CIS3D B1lb_YYSyoyo 
            Height          =   315
            Left            =   450
            Top             =   255
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            BackColor       =   16777152
            Caption         =   "##,###,###"
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
   End
   Begin Cis3D_v60.CIS3D DispSuryo 
      Height          =   600
      Left            =   105
      Top             =   420
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1058
      BackColor       =   16711680
      ForeColor       =   16777215
      Caption         =   "数 量"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   21.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   9375
      Top             =   120
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label ColorToday 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFC0C0&
      Caption         =   "当日色"
      Height          =   330
      Left            =   9375
      TabIndex        =   14
      Top             =   540
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label ColorSata 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FF8080&
      Caption         =   "土曜色"
      Height          =   330
      Left            =   8205
      TabIndex        =   13
      Top             =   540
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label ColorSunday 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0C0FF&
      Caption         =   "日曜色"
      Height          =   315
      Left            =   7050
      TabIndex        =   12
      Top             =   540
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label DayColor 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00E0E0E0&
      Caption         =   "対象外色"
      Height          =   300
      Left            =   5835
      TabIndex        =   11
      Top             =   540
      Visible         =   0   'False
      Width           =   1050
   End
End
Attribute VB_Name = "CXJ0420"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   受注計画状況照会
'**       フォームID    :   CXJ0420
'**       処理概要      :
'**
'**       作  成  日    :   2004/03/19  By CIS
'**       変  更  日    :   2004.10.27  By CIS  納入先名称獲得時の取引先区分修正(LostFocus)
'**       変  更  日    :   2008/04/10  By CIS  背番号8桁
'**       変  更  日    :   2008/05/02  By CIS  【金額】表示追加
'**       変  更  日    :   2008/06/11  By CIS  条件に【工区】追加
'**       変  更  日    :   2009/09/25  By CIS  受注データ読み込み時の分納回数廃止
'**       変  更  日    :   2014/02/04  By CIS WIN7対応
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    
    Dim mWeekStart              As Integer
    Dim mSaturdayColor          As Long
    Dim mSundayColor            As Long

    Dim mDayRow(31)             As Integer
    Dim mDayCol(31)             As Integer

    Dim mSuryo(5, 31)           As Long
    Dim mMaisu(5, 31)           As Long
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/02 INSERT START
    Dim mKingak(5, 31)          As Currency
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/02 INSERT END
    Dim mZaiko(31)              As Long
    Dim mNyuka(31)              As Long
    Dim mKensyu(31)             As Long
    Dim mZaikoNew               As Long
    
    Dim mSuryoK(31)             As Long
    Dim mSuryoA(31)             As Long
    
    Dim mSetType                As Integer
    Dim mCellSet                As Boolean
    Dim mSetDisp                As Integer
    Dim CisComboKK              As CisCombo     ' 2008/06/11 ADD
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

    CisDB.TimeOut = 3600
    
    If Not IniGet Then End
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
    
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Call CisFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/11 INSERT START
    '+----------------------+
    '+  コンボの設定(工区)  +
    '+----------------------+
    ' 自社工区
    Set CisComboKK = New CisCombo
    With CisComboKK
        Set .Connect = CisDB
        Set .CombObj = H1Comb_Koku
        .Control = "名称マスタ"
        .ComboName = "自社工区"
        .IniPath = gIniExe
        .SetSpace = 50
        .Ess = False
        .ComboInit
        .ComboSet
    End With
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/11 INSERT END
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    
    iH1_Ymd = Format(Date, "YYYYMM")
    ' 曜日開始を求める
    gSL_Select = "select 数字1 From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '曜日開始管理'"
    gSL_Select = gSL_Select & "   And 値 = '1'"
    If Not SYMRead(gSL_Select, 1) Then
       SYM.数字1 = 0
    End If
    
    mWeekStart = SYM.数字1 + 1

    mSaturdayColor = &HFFC0C0
    mSundayColor = &HC0C0FF

    Call GridInit
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    mSetType = 2
    Call PB_Chg_Click
' ADOｺﾝﾄﾛｰﾙ　接続情報　セット
    Adodc1.ConnectionString = CisDB.ConnectString
    '-----2014/02/04 EDIT START-----
    gStr = CisDB.ConnectString
    gConnect = "server=" & CisDB.Server & ";database=" & CisDB.DBName & ";User ID=" & CisDB.User & ";password=" & CisDB.PassWord(GetSetting(gRegKey, "Environment", "DrowsSap", ""))
    
    gStr = Replace(gStr, "Use Procedure", "Extended Properties=" & gConnect & ";Use Procedure")
    
    Adodc1.ConnectionString = gStr
    '-----2014/02/04 EDIT END-------
    
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
           Case vbKeyF6:        If PB_Chg.Visible Then Call PB_Chg_Click    '【切替】
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
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
    Call CisDB.DBDISConnect                            ' ﾃﾞｰﾀﾍﾞｰｽ切断
    Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
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
    
    If ProcHB = "H1" Then
        Call HeadBodyClear("H")
        iH1_Ymd = Format(Date, "YYYYMM")
        iH1_Seban.SetFocus
        GoTo PB_CAN_Ed
    End If
    
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        Call GridInit
        iH1_Seban.SetFocus
        GoTo PB_CAN_Ed
    End If
    
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    切　替  キ ー(F6)       +
'+----------------------------+
Private Sub PB_Chg_Click()
    Dim SaveRow As Long
    Dim SaveCol As Long
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/02 UPDATE START
'    If mSetType = 0 Then
'       mSetType = 1
'       DispSuryo = "枚 数"
'       PB_Chg.BtnCaption = "数量"
'    Else
'       mSetType = 0
'       DispSuryo = "数 量"
'       PB_Chg.BtnCaption = "枚数"
'    End If
    Select Case mSetType
           Case 0
                mSetType = 1
                DispSuryo = "枚 数"
                PB_Chg.BtnCaption = "金額"
           Case 1
                mSetType = 2
                DispSuryo = "金 額"
                PB_Chg.BtnCaption = "数量"
           Case 2
                mSetType = 0
                DispSuryo = "数 量"
                PB_Chg.BtnCaption = "枚数"
    End Select
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/02 UPDATE END
    SaveRow = vsGrid.Row
    SaveCol = vsGrid.Col
    Call GridValueSet
    vsGrid.Row = SaveRow
    vsGrid.Col = SaveCol
    
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Ymd" Then iH1_Ymd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_SSTorcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd(0): Exit Sub
    If PB_Look.Tag = "iH1_Seban" Then Call Look_Hinbn(1): Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn(0): Exit Sub
End Sub
'+-----------------------------+
'+       取引先検索             +
'+-----------------------------+
Private Sub Look_SSTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 3
    
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
Private Sub Look_UTorcd(KB As Integer)
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_Torcd = iH1_Torcd
    RV_TorKb = 3
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
'+       品番検索              +
'+-----------------------------+
Private Sub Look_Hinbn(KB As Integer)
    RV_Left = 0
    RV_Top = 0
        
    RVI_Hinmoku = 0 ' 受注
    RVI_SK = 9
    RV_Call = ""
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        Select Case KB
            Case 1: iH1_Seban = RV_Seban
            Case Else
                iH1_Hinbn = RV_Hinbn
        End Select
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
        gSL_Select = gSL_Select & "   And 品目 = 0"
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
    
'------------------------------------< 2004.10.27 MNT START >--------------------------
'
'    If TorNmGet(iH1_Torcd, 8) Then
    If TorNmGet(iH1_Torcd, 3) Then
        H1lb_Tornm = TRM_RName
    End If
'
'------------------------------------< 2004.10.27 MNT END   >--------------------------
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
    
    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    
    PB_Look.Visible = H1Mode
    
'   ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
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
            If .Name Like wbc_Enm & "*Comb*" Then .ListIndex = 0
        End With
    Next gObj
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

ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Dim SetItem     As String
    Dim SetDayStr   As String
    Dim YSuryo      As Long
    Dim YYSuryo     As Long
    Dim wWhere      As String
    Dim wJEndDay    As Long
    
    Head1Chk = False
    
    '======= 指定日必須チェック ========
    If Trim(iH1_Ymd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    指定日（年月又は年月日）を入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ymd.SetFocus
        Exit Function
    End If
    '======= 指定日 年必須チェック ========
    If CisFun.Val2(iH1_Ymd.cYear) = 0 Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    指定日（年）を入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ymd.SetFocus
       Exit Function
    End If
    '======= 指定日 月必須チェック ========
    If CisFun.Val2(iH1_Ymd.cMonth) = 0 Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    指定日（月）を入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ymd.SetFocus
       Exit Function
    End If
    '======= 指定日 日必須チェック ========
    If CisFun.Val2(iH1_Ymd.cDay) <> 0 Then
       If Not iH1_Ymd.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    指定日を正しい日付にて入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ymd.SetFocus
        Exit Function
       End If
    End If
    
    '*** グリッド初期化処理 ***
    Call GridInit
    
    '*** 配列の初期化 ***
    Erase mSuryo
    Erase mMaisu
    Erase mNyuka
    Erase mKensyu
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/02 INSERT START
    Erase mKingak
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/02 INSERT END
    
    PNL_Wait.Visible = True
    DoEvents
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/02 INSERT START
    If Not DBInput Then GoTo Head1Chk_Ed
    Message.BackColor = &HFFC0FF
    DoEvents
    If Not DBInputNaiji Then GoTo Head1Chk_Ed
    Message.BackColor = &HFFC0C0
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/02 INSERT END
        
    
    Head1Chk = True

Head1Chk_Ed:
    PNL_Wait.Visible = False
End Function

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/02 INSERT START
'::::::::::::::::::::::::::::::::::::::
'::                                  ::
'::       　 データ表示              ::
'::                                  ::
'::::::::::::::::::::::::::::::::::::::
Private Function DBInput() As Boolean
    Dim SetItem     As String
    Dim wWhere      As String
    
    DBInput = False
    
    '***********************
    '*     Where文作成     *
    '***********************
    With CisFun
         .Where_CreateKB = NewDocument
        ' 納入日
        .Where_Create AlphaMe, "Substring(納入日,1,6)", Mid(iH1_Ymd, 1, 6)
        ' 品番
        .Where_Create AlphaMe, "JY.品番", iH1_Hinbn
        ' 背番号
        .Where_Create AlphaMe, "HTM.背番号", iH1_Seban
        ' 仕入先
        .Where_Create AlphaMe, "JY.納入先", iH1_Torcd
        ' 受入
        .Where_Create AlphaMe, "JY.受入", iH1_Ukeir
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/11 INSERT START
        ' 工区
        Call CisComboKK.CodeGet
        .Where_Create AlphaMe, "ISNULL(HTM.自社工区, '')", CisComboKK.Code
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/11 INSERT END
    
        wWhere = .Where_Phrase
    
    End With
    
    SetItem = ""
    If Trim(iH1_Seban) <> "" Then
       SetItem = "HTM.背番号,"
    End If
    If Trim(iH1_Hinbn) <> "" Then
       SetItem = SetItem & "JY.品番,"
    End If
    If Trim(iH1_Torcd) <> "" Then
       SetItem = SetItem & "JY.納入先,"
    End If
    If Trim(iH1_Ukeir) <> "" Then
       SetItem = SetItem & "JY.受入,"
    End If
'*************************************   受注テーブルより出力   *****************************************
    gSL_Select = ""
    gSL_Select = gSL_Select & " Select 2 件数," & SetItem & "納入日,SUM(枚数) 枚数,Sum(納入数) 納入数," & vbCrLf
    gSL_Select = gSL_Select & "     CASE WHEN ISNULL(JT.単価,0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "          CEILING(( SUM(JY.納入数) * JT.単価 ) / 1000)" & vbCrLf
    gSL_Select = gSL_Select & "     WHEN ISNULL(JT2.単価,0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "          CEILING(( SUM(JY.納入数) * JT2.単価 ) / 1000)" & vbCrLf
    gSL_Select = gSL_Select & "     Else" & vbCrLf
    gSL_Select = gSL_Select & "       0" & vbCrLf
    gSL_Select = gSL_Select & "     END 金額, JY.品番" & vbCrLf
    gSL_Select = gSL_Select & "   From 受注テーブル JY WITH (INDEX = 計画対比1)" & vbCrLf
    gSL_Select = gSL_Select & "        Left Outer Join 受注単価マスタ JT" & vbCrLf
    gSL_Select = gSL_Select & "             ON  JT.品番 = JY.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.取引先 = JY.納入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.適用開始日付 <= JY.納入日" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND (JT.適用終了日付 = '' Or (JT.適用終了日付 <> ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.適用終了日付 >= JY.納入日 ))" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.履歴NO = ( SELECT MAX(JT1.履歴NO) FROM  受注単価マスタ JT1" & vbCrLf
    gSL_Select = gSL_Select & "                                 Where  JT1.品番 = JY.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                   AND  JT1.取引先 = JY.納入先" & vbCrLf
    gSL_Select = gSL_Select & "                                   AND  JT1.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "                                   AND  JT1.適用開始日付 <= JY.納入日" & vbCrLf
    gSL_Select = gSL_Select & "                                   And (JT1.適用終了日付 = '' Or (JT1.適用終了日付 <> ''" & vbCrLf
    gSL_Select = gSL_Select & "                                   And  JT1.適用終了日付 >= JY.納入日 )) )" & vbCrLf
    gSL_Select = gSL_Select & "        Left Outer Join 受注単価マスタ JT2" & vbCrLf
    gSL_Select = gSL_Select & "             ON  JT2.品番 = JY.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.取引先 = JY.納入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.適用開始日付 <= JY.納入日" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.仮単価区分 = '1'" & vbCrLf
    gSL_Select = gSL_Select & "            AND (JT2.適用終了日付 = '' Or (JT2.適用終了日付 <> ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.適用終了日付 >= JY.納入日 ))" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.履歴NO = ( SELECT MAX(JT3.履歴NO) FROM  受注単価マスタ JT3" & vbCrLf
    gSL_Select = gSL_Select & "                                 Where  JT3.品番 = JY.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                   AND  JT3.取引先 = JY.納入先" & vbCrLf
    gSL_Select = gSL_Select & "                                   AND  JT3.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "                                   AND  JT3.適用開始日付 <= JY.納入日" & vbCrLf
    gSL_Select = gSL_Select & "                                   And (JT3.適用終了日付 = '' Or (JT3.適用終了日付 <> ''" & vbCrLf
    gSL_Select = gSL_Select & "                                   And  JT3.適用終了日付 >= JY.納入日 )) )" & vbCrLf
    gSL_Select = gSL_Select & "        Left Outer Join 品番マスタ HM" & vbCrLf
    gSL_Select = gSL_Select & "             ON  HM.品番 = JY.品番" & vbCrLf
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/11 INSERT START
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番取引先マスタ HTM" & vbCrLf
    gSL_Select = gSL_Select & "             ON HTM.品番   = JY.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.取引先 = JY.納入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.受入   = JY.受入" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.品目   = 0" & vbCrLf
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/11 INSERT END
    gSL_Select = gSL_Select & CisFun.Where_Phrase & vbCrLf
'---    gSL_Select = gSL_Select & "    And 分納回数 = 0" & vbCrLf           '2009/09/25
    gSL_Select = gSL_Select & "  Group By " & SetItem & "納入日,JT.単価,JT2.単価, JY.品番" & vbCrLf

    gSL_Select = gSL_Select & " Union " & vbCrLf
'*************************************   出荷テーブルより出力   *****************************************
    gSL_Select = gSL_Select & " Select 3 件数," & SetItem & "出荷日,SUM(出荷枚数) 枚数,Sum(出荷数) 出荷数," & vbCrLf
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(JT.単価,0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                  CEILING(( SUM(SK.出荷数) * JT.単価 ) / 1000)" & vbCrLf
    gSL_Select = gSL_Select & "             WHEN ISNULL(JT2.単価,0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                  CEILING(( SUM(SK.出荷数) * JT2.単価 ) / 1000)" & vbCrLf
    gSL_Select = gSL_Select & "        Else 0" & vbCrLf
    gSL_Select = gSL_Select & "        END 金額, JY.品番" & vbCrLf
    gSL_Select = gSL_Select & "   From 出荷テーブル SK WITH (INDEX = PK_出荷テーブル)" & vbCrLf
    gSL_Select = gSL_Select & "        Left Outer Join 受注テーブル JY " & vbCrLf
    gSL_Select = gSL_Select & "             ON JY.出荷指示書NO   = SK.出荷指示書NO" & vbCrLf
    gSL_Select = gSL_Select & "            And JY.出荷指示書行NO = SK.出荷指示書行NO" & vbCrLf
    gSL_Select = gSL_Select & "        Left Outer Join 受注単価マスタ JT" & vbCrLf
    gSL_Select = gSL_Select & "             ON  JT.品番 = JY.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.取引先 = JY.納入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.適用開始日付 <= SK.出荷日" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND (JT.適用終了日付 = '' Or (JT.適用終了日付 <> ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.適用終了日付 >= SK.出荷日 ))" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.履歴NO = ( SELECT MAX(JT1.履歴NO) FROM  受注単価マスタ JT1" & vbCrLf
    gSL_Select = gSL_Select & "                                Where  JT1.品番 = JY.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                  AND  JT1.取引先 = JY.納入先" & vbCrLf
    gSL_Select = gSL_Select & "                                  AND  JT1.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "                                  AND  JT1.適用開始日付 <= SK.出荷日" & vbCrLf
    gSL_Select = gSL_Select & "                                  And (JT1.適用終了日付 = '' Or (JT1.適用終了日付 <> ''" & vbCrLf
    gSL_Select = gSL_Select & "                                  And  JT1.適用終了日付 >= SK.出荷日 )) )" & vbCrLf
    gSL_Select = gSL_Select & "        Left Outer Join 受注単価マスタ JT2" & vbCrLf
    gSL_Select = gSL_Select & "             ON  JT2.品番 = JY.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.取引先 = JY.納入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.適用開始日付 <= SK.出荷日" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.仮単価区分 = '1'" & vbCrLf
    gSL_Select = gSL_Select & "            AND (JT2.適用終了日付 = '' Or (JT2.適用終了日付 <> ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.適用終了日付 >= SK.出荷日 ))" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.履歴NO = ( SELECT MAX(JT3.履歴NO) FROM  受注単価マスタ JT3" & vbCrLf
    gSL_Select = gSL_Select & "                                 Where  JT3.品番 = JY.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                   AND  JT3.取引先 = JY.納入先" & vbCrLf
    gSL_Select = gSL_Select & "                                   AND  JT3.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "                                   AND  JT3.適用開始日付 <= SK.出荷日" & vbCrLf
    gSL_Select = gSL_Select & "                                   And (JT3.適用終了日付 = '' Or (JT3.適用終了日付 <> ''" & vbCrLf
    gSL_Select = gSL_Select & "                                   And  JT3.適用終了日付 >= SK.出荷日 )) )" & vbCrLf
    gSL_Select = gSL_Select & "        Left Outer Join 品番マスタ HM" & vbCrLf
    gSL_Select = gSL_Select & "             ON  HM.品番 = JY.品番" & vbCrLf
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/11 INSERT START
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番取引先マスタ HTM" & vbCrLf
    gSL_Select = gSL_Select & "             ON HTM.品番   = JY.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.取引先 = JY.納入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.受入   = JY.受入" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.品目   = 0" & vbCrLf
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/11 INSERT END
    gSL_Select = gSL_Select & Replace(wWhere, "納入日", "出荷日") & vbCrLf
    gSL_Select = gSL_Select & "  Group By " & SetItem & "出荷日,JT.単価,JT2.単価, JY.品番" & vbCrLf

    gSL_Select = gSL_Select & " Union " & vbCrLf
'*************************************   検収テーブルより出力   *****************************************
    gSL_Select = gSL_Select & " Select 4 件数," & SetItem & "検収日," & vbCrLf
    gSL_Select = gSL_Select & "        Sum(CASE WHEN 赤黒区分 <> 0 THEN KN.枚数 * -1 ELSE KN.枚数 END) 枚数," & vbCrLf
    gSL_Select = gSL_Select & "        Sum(CASE WHEN 赤黒区分 <> 0 THEN KN.検収数 * -1 ELSE KN.検収数 END) 検収数," & vbCrLf
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(JT.単価,0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                  CEILING(( Sum(CASE WHEN 赤黒区分 <> 0 THEN 検収数 * -1 ELSE 検収数 END)  * JT.単価 ) / 1000)" & vbCrLf
    gSL_Select = gSL_Select & "             WHEN ISNULL(JT2.単価,0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                  CEILING(( Sum(CASE WHEN 赤黒区分 <> 0 THEN 検収数 * -1 ELSE 検収数 END)  * JT2.単価 ) / 1000)" & vbCrLf
    gSL_Select = gSL_Select & "        Else 0" & vbCrLf
    gSL_Select = gSL_Select & "        END 金額, JY.品番" & vbCrLf
    gSL_Select = gSL_Select & "   From 検収テーブル KN WITH (INDEX = 計画対比)" & vbCrLf
    gSL_Select = gSL_Select & "        Left Outer Join 品番マスタ HM" & vbCrLf
    gSL_Select = gSL_Select & "             ON HM.品番 = KN.品番" & vbCrLf
    gSL_Select = gSL_Select & "        Left Outer Join 受注テーブル JY " & vbCrLf
    gSL_Select = gSL_Select & "             ON JY.出荷指示書NO   = KN.指示書NO" & vbCrLf
    gSL_Select = gSL_Select & "            And JY.出荷指示書行NO = KN.指示書行NO" & vbCrLf
    gSL_Select = gSL_Select & "        Left Outer Join 受注単価マスタ JT" & vbCrLf
    gSL_Select = gSL_Select & "             ON  JT.品番 = KN.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.取引先 = KN.納入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.適用開始日付 <= KN.検収日" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND (JT.適用終了日付 = '' Or (JT.適用終了日付 <> ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.適用終了日付 >= KN.検収日 ))" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT.履歴NO = ( SELECT MAX(JT1.履歴NO) FROM  受注単価マスタ JT1" & vbCrLf
    gSL_Select = gSL_Select & "                                 Where  JT1.品番 = KN.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                   AND  JT1.取引先 = KN.納入先" & vbCrLf
    gSL_Select = gSL_Select & "                                   AND  JT1.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "                                   AND  JT1.適用開始日付 <= KN.検収日" & vbCrLf
    gSL_Select = gSL_Select & "                                   And (JT1.適用終了日付 = '' Or (JT1.適用終了日付 <> ''" & vbCrLf
    gSL_Select = gSL_Select & "                                   And  JT1.適用終了日付 >= KN.検収日 )) )" & vbCrLf
    gSL_Select = gSL_Select & "        Left Outer Join 受注単価マスタ JT2" & vbCrLf
    gSL_Select = gSL_Select & "             ON  JT2.品番 = KN.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.取引先 = KN.納入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.適用開始日付 <= KN.検収日" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.仮単価区分 = '1'" & vbCrLf
    gSL_Select = gSL_Select & "            AND (JT2.適用終了日付 = '' Or (JT2.適用終了日付 <> ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.適用終了日付 >= KN.検収日 ))" & vbCrLf
    gSL_Select = gSL_Select & "            AND  JT2.履歴NO = ( SELECT MAX(JT3.履歴NO) FROM  受注単価マスタ JT3" & vbCrLf
    gSL_Select = gSL_Select & "                                  Where  JT3.品番 = KN.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                    AND  JT3.取引先 = KN.納入先" & vbCrLf
    gSL_Select = gSL_Select & "                                    AND  JT3.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "                                    AND  JT3.適用開始日付 <= KN.検収日" & vbCrLf
    gSL_Select = gSL_Select & "                                    And (JT3.適用終了日付 = '' Or (JT3.適用終了日付 <> ''" & vbCrLf
    gSL_Select = gSL_Select & "                                    And  JT3.適用終了日付 >= KN.検収日 )) )" & vbCrLf
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/11 INSERT START
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番取引先マスタ HTM" & vbCrLf
    gSL_Select = gSL_Select & "             ON HTM.品番   = JY.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.取引先 = JY.納入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.受入   = JY.受入" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.品目   = 0" & vbCrLf
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/11 INSERT END
    gSL_Select = gSL_Select & Replace(wWhere, "納入日", "検収日") & vbCrLf
    gSL_Select = gSL_Select & "  Group By " & SetItem & "検収日,JT.単価,JT2.単価, JY.品番" & vbCrLf
    
    If Not JYTRead(gSL_Select) Then
        Call JYTClose
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    該当データが存在しません｡            "
            .MB_MSG(3) = "    再度抽出条件を入力してください       "
            .MB_Title = "抽出処理"
            .MB_Button = Error
            .MBOX
        End With
        iH1_Seban.SetFocus
        Exit Function
    End If
    
    '********************************
    '**        抽出情報退避        **
    '********************************
    Do Until Not JYT_RDSTS
        
        gCnt2 = CisFun.Val2(Mid(JYT.納入日, 7, 2))
        If gCnt2 <> 0 Then
           mSuryo(JYT.件数, gCnt2) = mSuryo(JYT.件数, gCnt2) + JYT.納入数
           mMaisu(JYT.件数, gCnt2) = mMaisu(JYT.件数, gCnt2) + JYT.枚数
           mKingak(JYT.件数, gCnt2) = mKingak(JYT.件数, gCnt2) + JYT.金額
           If JYT.件数 = 3 Then
              mNyuka(gCnt2) = mNyuka(gCnt2) + JYT.納入数
           End If
           If JYT.件数 = 4 Then
              mKensyu(gCnt2) = mKensyu(gCnt2) + JYT.納入数
           End If
        End If

        Call JYTReadNext
    Loop
    Call JYTClose
    
    
    DBInput = True
    
End Function

'::::::::::::::::::::::::::::::::::::::::::::
'::                                        ::
'::       　 データ表示(内示)              ::
'::                                        ::
'::::::::::::::::::::::::::::::::::::::::::::
Private Function DBInputNaiji() As Boolean
    Dim wKari       As String
    Dim wTanka      As Currency
    Dim wKingak     As Currency
    Dim wLpCnt      As Integer
    Dim XStr        As String
    
    DBInputNaiji = False
    gSL_Select = ""
    gSL_Select = gSL_Select & " Select "
    For wLpCnt = 1 To 31
        XStr = Format(wLpCnt, "00")
        gSL_Select = gSL_Select & "        内示数量" & XStr & "," & vbCrLf
        gSL_Select = gSL_Select & "        CASE WHEN ISNULL(JT" & XStr & ".単価,0) <> 0 THEN " & vbCrLf
        gSL_Select = gSL_Select & "                      CEILING(( HN.内示数量" & XStr & " * JT" & XStr & ".単価 ) / 1000) " & vbCrLf
        gSL_Select = gSL_Select & "                 WHEN ISNULL(KJT" & XStr & ".単価,0) <> 0 THEN " & vbCrLf
        gSL_Select = gSL_Select & "                      CEILING(( HN.内示数量" & XStr & " * KJT" & XStr & ".単価 ) / 1000) " & vbCrLf
        gSL_Select = gSL_Select & "        Else 0 END 内示金額" & XStr & "," & vbCrLf
    Next wLpCnt
    gSL_Select = gSL_Select & "   HTM.背番号, HN.翌月内示数" & vbCrLf
    gSL_Select = gSL_Select & "   From 内示テーブル HN" & vbCrLf
    gSL_Select = gSL_Select & "        Left Outer Join 品番取引先マスタ HTM" & vbCrLf
    gSL_Select = gSL_Select & "             ON  HN.品番 = HTM.品番" & vbCrLf
    gSL_Select = gSL_Select & "            And HN.納入先 = HTM.取引先" & vbCrLf
    gSL_Select = gSL_Select & "            And HN.受入 = HTM.受入" & vbCrLf
    For wLpCnt = 1 To 31
        XStr = Format(wLpCnt, "00")
        gSL_Select = gSL_Select & "        Left Outer Join 受注単価マスタ JT" & XStr & vbCrLf
        gSL_Select = gSL_Select & "             ON  JT" & XStr & ".品番 = HN.品番" & vbCrLf
        gSL_Select = gSL_Select & "            AND  JT" & XStr & ".取引先 = HN.納入先" & vbCrLf
        gSL_Select = gSL_Select & "            AND  JT" & XStr & ".適用開始日付 <= HN.内示年月 + '" & XStr & "'" & vbCrLf
        gSL_Select = gSL_Select & "            AND  JT" & XStr & ".仮単価区分 = ''" & vbCrLf
        gSL_Select = gSL_Select & "            AND (JT" & XStr & ".適用終了日付 = '' " & vbCrLf
        gSL_Select = gSL_Select & "          Or (JT" & XStr & ".適用終了日付 <> ''" & vbCrLf
        gSL_Select = gSL_Select & "            AND  JT" & XStr & ".適用終了日付 >= HN.内示年月 + '" & XStr & "' ))" & vbCrLf
        gSL_Select = gSL_Select & "            AND  JT" & XStr & ".履歴NO = (SELECT MAX(SJT" & XStr & ".履歴NO) FROM  受注単価マスタ SJT" & XStr & vbCrLf
        gSL_Select = gSL_Select & "                                 Where  SJT" & XStr & ".品番 = HN.品番" & vbCrLf
        gSL_Select = gSL_Select & "                                   AND  SJT" & XStr & ".取引先 = HN.納入先" & vbCrLf
        gSL_Select = gSL_Select & "                                   AND  SJT" & XStr & ".仮単価区分 = ''" & vbCrLf
        gSL_Select = gSL_Select & "                                   AND  SJT" & XStr & ".適用開始日付 <= HN.内示年月 + '" & XStr & "'" & vbCrLf
        gSL_Select = gSL_Select & "                                   And (SJT" & XStr & ".適用終了日付 = '' Or (SJT" & XStr & ".適用終了日付 <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                                   And  SJT" & XStr & ".適用終了日付 >= HN.内示年月 + '" & XStr & "' )) )" & vbCrLf
        gSL_Select = gSL_Select & "        Left Outer Join 受注単価マスタ KJT" & XStr & vbCrLf
        gSL_Select = gSL_Select & "             ON  KJT" & XStr & ".品番 = HN.品番" & vbCrLf
        gSL_Select = gSL_Select & "            AND  KJT" & XStr & ".取引先 = HN.納入先" & vbCrLf
        gSL_Select = gSL_Select & "            AND  KJT" & XStr & ".適用開始日付 <= HN.内示年月 + '" & XStr & "'" & vbCrLf
        gSL_Select = gSL_Select & "            AND  KJT" & XStr & ".仮単価区分 = '1'" & vbCrLf
        gSL_Select = gSL_Select & "            AND (KJT" & XStr & ".適用終了日付 = '' Or (KJT" & XStr & ".適用終了日付 <> ''" & vbCrLf
        gSL_Select = gSL_Select & "            AND  KJT" & XStr & ".適用終了日付 >= HN.内示年月 + '" & XStr & "' ))" & vbCrLf
        gSL_Select = gSL_Select & "            AND  KJT" & XStr & ".履歴NO = ( SELECT MAX(SKJT" & XStr & ".履歴NO) FROM  受注単価マスタ SKJT" & XStr & vbCrLf
        gSL_Select = gSL_Select & "                                  Where  SKJT" & XStr & ".品番 = HN.品番" & vbCrLf
        gSL_Select = gSL_Select & "                                    AND  SKJT" & XStr & ".取引先 = HN.納入先" & vbCrLf
        gSL_Select = gSL_Select & "                                    AND  SKJT" & XStr & ".仮単価区分 = ''" & vbCrLf
        gSL_Select = gSL_Select & "                                    AND  SKJT" & XStr & ".適用開始日付 <= HN.内示年月 + '" & XStr & "'" & vbCrLf
        gSL_Select = gSL_Select & "                                    And (SKJT" & XStr & ".適用終了日付 = '' Or (SKJT" & XStr & ".適用終了日付 <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                                    And  SKJT" & XStr & ".適用終了日付 >= HN.内示年月 + '" & XStr & "' )) )" & vbCrLf
    Next wLpCnt
    
    '***********************
    '*     Where文作成     *
    '***********************
    With CisFun
         .Where_CreateKB = NewDocument
        ' 内示年月
        .Where_Create AlphaMe, "HN.内示年月", Mid(iH1_Ymd, 1, 6)
        ' 品番
        If Trim(iH1_Seban) = "" Then
           .Where_Create AlphaMe, "HN.品番", iH1_Hinbn
        Else
           .Where_Create AlphaMe, "HN.品番", iH1_Hinbn
        End If
        ' 背番号
        .Where_Create AlphaMe, "HTM.背番号", iH1_Seban
        ' 仕入先
        .Where_Create AlphaMe, "HN.納入先", iH1_Torcd
        ' 受入
        .Where_Create AlphaMe, "HN.受入", iH1_Ukeir
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/11 INSERT START
        ' 工区
        Call CisComboKK.CodeGet
        .Where_Create AlphaMe, "ISNULL(HTM.自社工区, '')", CisComboKK.Code
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/11 INSERT END

        gSL_Select = gSL_Select & .Where_Phrase

    End With
    
    'CisFun.ClipPut gSL_Select, False
    
    If NJTRead(gSL_Select) Then
        Do Until Not NJT_RDSTS
            For gCnt2 = 1 To 31
                If gCnt2 <> 0 Then
                   mSuryo(1, gCnt2) = mSuryo(1, gCnt2) + NJT.内示数量(gCnt2)
                   wKingak = CCur(CisFun.HasuProc(CDbl(NJT.内示金額(gCnt2)), 切上げ))
                 
                   mKingak(1, gCnt2) = mKingak(1, gCnt2) + wKingak
                End If
            Next gCnt2
            B1lb_YYSyoyo = Format(CisFun.Val2(B1lb_YYSyoyo) + NJT.翌月内示数, "#,##0")

            Call NJTReadNext
        Loop
    End If
    Call NJTClose

    gSL_Select = "Select HN.*,HTM.背番号 From 内示テーブル HN"
    gSL_Select = gSL_Select & " Left Outer Join 品番取引先マスタ HTM"
    gSL_Select = gSL_Select & "     ON  HN.品番 = HTM.品番"
    gSL_Select = gSL_Select & "     And HN.納入先 = HTM.取引先"
    gSL_Select = gSL_Select & "     And HN.受入 = HTM.受入"
    With CisFun
         .Where_CreateKB = NewDocument
        ' 内示年月
        .Date_RsFormat = fYM
        .Date_Add iH1_Ymd, M, -1
        gStr = .Date_Result

        ' 内示年月
        .Where_Create AlphaMe, "HN.内示年月", Mid(gStr, 1, 6)
        ' 品番
        If Trim(iH1_Seban) = "" Then
           .Where_Create AlphaMe, "HN.品番", iH1_Hinbn
        Else
           .Where_Create AlphaMe, "HN.品番", iH1_Hinbn
        End If
        ' 背番号
        .Where_Create AlphaMe, "HTM.背番号", iH1_Seban
        ' 仕入先
        .Where_Create AlphaMe, "HN.納入先", iH1_Torcd
        ' 受入
        .Where_Create AlphaMe, "HN.受入", iH1_Ukeir
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/11 INSERT START
        ' 工区
        Call CisComboKK.CodeGet
        .Where_Create AlphaMe, "ISNULL(HTM.自社工区, '')", CisComboKK.Code
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/11 INSERT END

        gSL_Select = gSL_Select & .Where_Phrase

    End With
    
    If NJTRead(gSL_Select) Then
        Do Until Not NJT_RDSTS
            B1lb_YSyoyo = Format(CisFun.Val2(B1lb_YSyoyo) + NJT.内示数, "#,##0")

            Call NJTReadNext
        Loop
    End If
    Call NJTClose
        
    Call GridValueSet
    
    With vsGrid
        If CisFun.Val2(iH1_Ymd.cDay) = 0 Then
           If Mid(iH1_Ymd, 1, 6) <> Format(Date, "YYYYMM") Then
              gCnt1 = 1
           Else
              gCnt1 = CisFun.Val2(Format(Date, "dd"))
           End If
        Else
           gCnt1 = CisFun.Val2(iH1_Ymd.cDay)
        End If
        mCellSet = True
        .Row = mDayRow(gCnt1)
        .Col = mDayCol(gCnt1)
        mCellSet = False
    End With
    
    DBInputNaiji = True
End Function
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/02 INSERT END

'**************************************
'*                                    *
'*          グリッド処理              *
'*                                    *
'**************************************
Private Sub GridInit()
    Dim SetColor        As Boolean
    
    mCellSet = True
    
    With vsGrid
        .Redraw = False
        .Rows = 1: .Cols = 1
        .Rows = 30: .Cols = 9
        
        .Row = 0
        For gCnt2 = 0 To .Cols - 1
            .Col = gCnt2
            .ColAlignment(gCnt2) = flexAlignCenterCenter
            If gCnt2 > 0 Then
                If gCnt2 = 1 Then
                   .ColWidth(gCnt2) = 30
                Else
                   .ColWidth(gCnt2) = 1300
                End If
            End If
        Next gCnt2
        
        Call DaySet
        
        For gCnt1 = 0 To .Rows - 1
            .Row = gCnt1: .Col = 0
            If gCnt1 Mod 5 = 0 Then
                .Text = "日 付"
                For gCnt2 = 0 To .Cols - 1
                    .Col = gCnt2
                Next gCnt2
            Else
                .Text = Grid_RowHead(gCnt1 Mod 5).Tag
                .CellBackColor = B1lb_TSyoyo(gCnt1 Mod 5).BackColor
            End If

            For gCnt2 = 2 To .Cols - 1
                .Col = gCnt2
                SetColor = True
                If gCnt1 Mod 5 = 0 Then
                   SetColor = False
                End If
                If gCnt1 < 5 Then
                   If Trim(.TextMatrix(0, gCnt2)) = "" Then
                      .CellBackColor = DayColor.BackColor
                      SetColor = False
                   End If
                End If
                If gCnt1 > 20 Then
                   If Trim(.TextMatrix(20, gCnt2)) = "" Then
                      .CellBackColor = DayColor.BackColor
                      SetColor = False
                   End If
                End If
                If gCnt1 > 25 Then
                   If Trim(.TextMatrix(25, gCnt2)) = "" Then
                      .CellBackColor = DayColor.BackColor
                      SetColor = False
                   End If
                End If
                If SetColor Then
                   .CellBackColor = B1lb_TSyoyo(gCnt1 Mod 5).BackColor
                   .CellAlignment = flexAlignRightCenter
                End If
                
            Next gCnt2
        Next gCnt1
        
        .Row = 1: .Col = 2
        .Redraw = True
    End With
    mCellSet = False

End Sub

Private Sub DaySet()
    Dim wDayCol         As Integer
    Dim SetDay          As Integer
    Dim EndDay          As Integer
    Dim wKado           As Boolean
    
    wDayCol = 1
    SetDay = 1
    
'       *  カレンダ部日付 セット  *
    gSL_Select = "Select * From カレンダマスタ"
    gSL_Select = gSL_Select & " Where 取引先区分 = 1"
    gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   年 = '" & Format(iH1_Ymd.cYear, "0000") & "'"
    gSL_Select = gSL_Select & " And   月 = '" & Format(iH1_Ymd.cMonth, "00") & "'"
    gSL_Select = gSL_Select & " And   受入 = '" & RTrim(iH1_Ukeir) & "'"
    If Not CLMRead(gSL_Select, 1) Then
        Call ItemsClearCLM
        CLM.稼動区分 = String(31, "0")
        CLM.夜勤区分 = String(31, "0")
        With CisFun
            .Date_RsFormat = fDD
            .Date_End iH1_Ymd
            CLM.末日 = .Val2(.Date_Result)
            CLM.稼動日数 = CLM.末日
        End With
    End If
    
    EndDay = CLM.末日
    gDate = DateValue(iH1_Ymd.cYear & "/" & iH1_Ymd.cMonth & "/01")
    gInt = Weekday(gDate, mWeekStart)
    Erase mDayRow
    Erase mDayCol
    
    With vsGrid
        For gCnt1 = 0 To .Rows - 1 Step 5
            .Row = gCnt1
            For gCnt2 = 1 To .Cols - 2
                If gCnt2 >= gInt Then
                   .Col = gCnt2 + 1
                   If SetDay <= EndDay Then
                      gDate = DateValue(iH1_Ymd.cYear & "/" & iH1_Ymd.cMonth & "/" & SetDay)
                      gCnt3 = Weekday(gDate, mWeekStart)
                      .CellBackColor = H1lb_Tornm.BackColor
                      
                      wKado = False
                      If Mid(CLM.稼動区分, SetDay, 1) = "0" Or _
                         Mid(CLM.夜勤区分, SetDay, 1) = "0" Then
                         wKado = True
                      End If
                      If Not wKado Then
                         .CellBackColor = ColorSunday.BackColor
                      End If
                      
                      If mWeekStart = 2 And gCnt3 = 7 And Not wKado Then
                         .CellBackColor = ColorSunday.BackColor
                      End If
                      If mWeekStart = 2 And gCnt3 = 6 And Not wKado Then
                         .CellBackColor = ColorSata.BackColor
                      End If
                      If mWeekStart = 1 And gCnt3 = 1 And Not wKado Then
                         .CellBackColor = ColorSunday.BackColor
                      End If
                      If mWeekStart = 1 And gCnt3 = 7 And Not wKado Then
                         .CellBackColor = ColorSata.BackColor
                      End If
                      
                      If Date = gDate And CisFun.Val2(iH1_Ymd.cDay) = 0 Then
                         .CellBackColor = ColorToday.BackColor
                      End If
                      If CisFun.Val2(iH1_Ymd.cDay) <> 0 And CisFun.Val2(iH1_Ymd.cDay) = SetDay Then
                         .CellBackColor = ColorToday.BackColor
                      End If
                      
                      .TextMatrix(gCnt1, gCnt2 + 1) = SetDay & " (" & WeekdayName(gCnt2, True, mWeekStart) & ")"
                      mDayRow(SetDay) = gCnt1
                      mDayCol(SetDay) = gCnt2 + 1

                      SetDay = SetDay + 1
                      gInt = 0
                   Else
                      .CellBackColor = DayColor.BackColor
                   End If
                Else
                    .CellBackColor = DayColor.BackColor
                End If
            Next gCnt2
        Next gCnt1
    End With
End Sub
Private Sub GridValueSet()
    Dim TSuryo(5)    As Long
    Dim SSuryo(5)    As Long
    Dim SetSu        As Long
    
    Call GridInit
    
    mCellSet = True
            
    With vsGrid
        .Redraw = False
        For gCnt1 = 1 To 31
            If mDayRow(gCnt1) = 0 And mDayCol(gCnt1) = 0 Then Exit For
            gDate = DateValue(iH1_Ymd.cYear & "/" & iH1_Ymd.cMonth & "/" & gCnt1)
            .Col = mDayCol(gCnt1)
            For gCnt2 = 1 To 4
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/07 UPDATE START
'                If mSetType = 0 Then
'                    SetSu = mSuryo(gCnt2, gCnt1)
'                    If gCnt2 = 4 Then
'                       If mSetDisp = 1 Then
'                          SetSu = mSuryoK(gCnt1)
'                       End If
'                       If mSetDisp = 2 Then
'                          SetSu = mSuryoA(gCnt1)
'                       End If
'                    End If
'                    .TextMatrix(mDayRow(gCnt1) + gCnt2, mDayCol(gCnt1)) = _
'                                            Format(SetSu, "#,###")
'                Else
'                    SetSu = mMaisu(gCnt2, gCnt1)
'                    .TextMatrix(mDayRow(gCnt1) + gCnt2, mDayCol(gCnt1)) = _
'                                            Format(SetSu, "#,###")
'                End If
                If mSetType = 0 Then
                    SetSu = mSuryo(gCnt2, gCnt1)
                    If gCnt2 = 4 Then
                       If mSetDisp = 1 Then
                          SetSu = mSuryoK(gCnt1)
                       End If
                       If mSetDisp = 2 Then
                          SetSu = mSuryoA(gCnt1)
                       End If
                    End If
                    .TextMatrix(mDayRow(gCnt1) + gCnt2, mDayCol(gCnt1)) = _
                                            Format(SetSu, "#,###")
                End If
                If mSetType = 1 Then
                    SetSu = mMaisu(gCnt2, gCnt1)
                    .TextMatrix(mDayRow(gCnt1) + gCnt2, mDayCol(gCnt1)) = _
                                            Format(SetSu, "#,###")
                End If
                If mSetType = 2 Then
                    SetSu = mKingak(gCnt2, gCnt1)
                    .TextMatrix(mDayRow(gCnt1) + gCnt2, mDayCol(gCnt1)) = _
                                            Format(SetSu, "#,###")
                End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/07 UPDATE END
                .Row = mDayRow(gCnt1) + gCnt2
                If Date = gDate And CisFun.Val2(iH1_Ymd.cDay) = 0 Then
                   .CellBackColor = ColorToday.BackColor
                End If
                If CisFun.Val2(iH1_Ymd.cDay) <> 0 And CisFun.Val2(iH1_Ymd.cDay) = gCnt1 Then
                   .CellBackColor = ColorToday.BackColor
                End If
                
                TSuryo(gCnt2) = TSuryo(gCnt2) + SetSu
                
                If Date >= gDate And CisFun.Val2(iH1_Ymd.cDay) = 0 Then
                   SSuryo(gCnt2) = SSuryo(gCnt2) + SetSu
                End If
                If CisFun.Val2(iH1_Ymd.cDay) <> 0 And CisFun.Val2(iH1_Ymd.cDay) >= gCnt1 Then
                   SSuryo(gCnt2) = SSuryo(gCnt2) + SetSu
                End If
                
            Next gCnt2
        Next gCnt1
        
        For gCnt1 = 1 To 4
            B1lb_TSyoyo(gCnt1) = Format(TSuryo(gCnt1), "#,###")
            B1lb_SSuryo(gCnt1) = Format(SSuryo(gCnt1), "#,###")
        Next gCnt1
        .Redraw = True
    End With
    
    mCellSet = False
    
End Sub
Private Function GetDay() As Integer
    Dim wCnt        As Long
    GetDay = 0
    With vsGrid
        For wCnt = 1 To 31
            If mDayRow(wCnt) = (.Row \ 5) * 5 And _
               mDayCol(wCnt) = .Col Then
               GetDay = wCnt
               Exit Function
            End If
        Next wCnt
    End With
End Function
Private Sub vsGrid_DblClick()
    Dim SetDay      As Integer
    Dim SetType     As Integer
    Dim CallFlg     As Boolean
    
    
    SetDay = GetDay
    SetType = vsGrid.Row Mod 5
    
    CallFlg = False
    With vsGrid
        If SetType <> 4 Then
            If mSuryo(2, SetDay) <> 0 Or _
               mNyuka(SetDay) <> 0 Then
               CallFlg = True
            End If
        Else
            If mKensyu(SetDay) <> 0 Then
               CallFlg = True
            End If
        End If
    End With
    
    If SetDay = 0 Or Not CallFlg Then
       Exit Sub
    End If
        
    Call CisFrm.MousePT(1)
    
    RV_SYmd = Mid(iH1_Ymd, 1, 6) & Format(SetDay, "00")
    RV_Hinbn = RTrim(iH1_Hinbn)
    RV_Seban = RTrim(iH1_Seban)
    RV_Torcd = RTrim(iH1_Torcd)
    RV_Ukeir = RTrim(iH1_Ukeir)
    Rv_Call3 = "K"
    Rv_CallRtn = 0
    If SetType = 3 Then
       Rv_CallRtn = 1
    End If
    
    CXJ0420.Hide
    
    Set Rv_CallObj = CXJ0420
    Set Rv_CallObj2 = CXJ0425
    
    Rv_Call3Rtn = SetType
    
    CXJ0425.Show
    
    
    If SetType <> 4 Then
       CXJ0410.Show
    Else
       CXK0410.Show
    End If
'    Unload CXH0410
'    Set CXH0410 = Nothing
'
'    Call CisFrm.MousePT(11)
'
'    CXH0420.Show
End Sub



