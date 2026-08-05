VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Begin VB.Form CXJ0310 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "検収入力"
   ClientHeight    =   11115
   ClientLeft      =   855
   ClientTop       =   960
   ClientWidth     =   15240
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
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  '最大化
   Begin VB.PictureBox H5_Area1 
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
      Height          =   1275
      Left            =   8475
      ScaleHeight     =   1215
      ScaleWidth      =   1905
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1470
      Visible         =   0   'False
      Width           =   1965
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   1110
         Left            =   75
         Top             =   60
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "バーコード"
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
         Begin CisText_V60.CisText iH5_Bar 
            Height          =   360
            Left            =   105
            TabIndex        =   9
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
            cFormat         =   "000000000000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "000000000000"
            MaxLength       =   12
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "############"
            cILength        =   12
            cAutoEject      =   0   'False
         End
      End
   End
   Begin PdqcommLib.PDQComm PDQComm1 
      Height          =   480
      Left            =   1260
      TabIndex        =   19
      Top             =   660
      Width           =   480
      _Version        =   196610
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.26
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FastScroll      =   0   'False
      Object.Height          =   0
      Object.Width           =   0
   End
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   4830
      Top             =   4290
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
   Begin Cis3D_v60.CIS3D CIS3D1 
      Height          =   1365
      Left            =   10635
      Top             =   1440
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2408
      BackColor       =   0
      ForeColor       =   65535
      Caption         =   ""
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
      cFont3DColor2   =   12632256
      cAlingnment     =   7
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   420
         Left            =   45
         Top             =   45
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   741
         Caption         =   "表示件数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B1lb_Kensu 
            Height          =   360
            Left            =   1155
            Top             =   30
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
               Charset         =   0
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
         Height          =   420
         Left            =   45
         Top             =   465
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   741
         Caption         =   "検 収 済"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B1lb_KKensu 
            Height          =   390
            Left            =   1170
            Top             =   15
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   688
            BackColor       =   12648384
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
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
         Height          =   420
         Left            =   45
         Top             =   885
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   741
         ForeColor       =   0
         Caption         =   "未 検 収"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B1lb_BKensu 
            Height          =   345
            Left            =   1170
            Top             =   30
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            BackColor       =   16761087
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
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
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   420
         Left            =   2160
         Top             =   45
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   741
         ForeColor       =   0
         Caption         =   "未 処 理"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B1lb_MKensu 
            Height          =   345
            Left            =   1170
            Top             =   15
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            BackColor       =   16777215
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
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
      Begin Cis3D_v60.CIS3D CIS3D16 
         Height          =   420
         Left            =   2160
         Top             =   465
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   741
         Caption         =   "検 収 済"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D TTL_KSKensu 
            Height          =   390
            Left            =   1170
            Top             =   15
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   688
            BackColor       =   12648384
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
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
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   420
         Left            =   2160
         Top             =   885
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   741
         ForeColor       =   0
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
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
      Height          =   750
      Left            =   8730
      ScaleHeight     =   690
      ScaleWidth      =   6150
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   525
      Width           =   6210
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   3525
         Top             =   45
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
         Left            =   4380
         Top             =   45
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
         Left            =   5235
         Top             =   45
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
         Top             =   45
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
      Begin CisBtn_60.CisBtn PB_CHG 
         Height          =   585
         Left            =   1815
         Top             =   45
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
      Begin CisBtn_60.CisBtn PB_UPD 
         Height          =   585
         Left            =   2670
         Top             =   45
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
         BtnCaption      =   "更新"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   0
         BtnShadowColor  =   16777215
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
         PFForeColor     =   65535
         PFShadowColor   =   0
         PFBackColor     =   0
      End
      Begin CisBtn_60.CisBtn PB_Bun 
         Height          =   585
         Left            =   945
         Top             =   45
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
         BtnCaption      =   "個別"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   192
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
         PFBackColor     =   192
      End
   End
   Begin Cis3D_v60.CIS3D H4_Area1 
      Height          =   1320
      Left            =   495
      Top             =   3750
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2328
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
      Begin VB.TextBox QRDummy 
         Height          =   330
         Left            =   3645
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   30
         Width           =   240
      End
      Begin Cis3D_v60.CIS3D QRMsg 
         Height          =   1170
         Left            =   90
         Top             =   75
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   2064
         BackColor       =   12648384
         ForeColor       =   16711680
         Caption         =   "QRリーダー読取"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   7
         cAlingnment     =   7
      End
   End
   Begin VB.PictureBox H2_Area1 
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
      Height          =   1275
      Left            =   4440
      ScaleHeight     =   1215
      ScaleWidth      =   4200
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   450
      Width           =   4260
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   1110
         Left            =   2250
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "納 番"
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
         Begin CisText_V60.CisText iH2_Nouno 
            Height          =   360
            Left            =   195
            TabIndex        =   7
            Top             =   480
            Width           =   780
            _ExtentX        =   1376
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
            Text            =   "XXXXX"
            MaxLength       =   5
            cDataReplace    =   1
            cFaZero         =   0
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1110
         Left            =   60
         Top             =   60
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "納 入 日"
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
         Begin CisYMD_With_Btn.CisYMDwB iH2_NYmd 
            Height          =   360
            Left            =   165
            TabIndex        =   6
            Top             =   495
            Width           =   1860
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
            Object.Height          =   360
            Object.Width           =   1860
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
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   1110
         Left            =   3465
         Top             =   60
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "頁"
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
         Begin CisText_V60.CisText iH2_NPage 
            Height          =   360
            Left            =   135
            TabIndex        =   8
            Top             =   480
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
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
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
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
      Height          =   1275
      Left            =   390
      ScaleHeight     =   1215
      ScaleWidth      =   2265
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1455
      Width           =   2325
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   1110
         Left            =   60
         Top             =   60
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   1958
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_NYmd 
            Height          =   360
            Left            =   120
            TabIndex        =   0
            Top             =   450
            Width           =   1860
            _ExtentX        =   3069
            _ExtentY        =   582
            cCalender       =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Object.Height          =   360
            Object.Width           =   1860
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   7725
      Left            =   390
      ScaleHeight     =   7665
      ScaleWidth      =   14505
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2850
      Width           =   14565
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   7560
         Left            =   60
         Top             =   45
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   13335
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
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   7350
            Left            =   90
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   105
            Width           =   14265
            _Version        =   196608
            _ExtentX        =   25162
            _ExtentY        =   12965
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
            FormatString    =   "№|処理 |納　期 　　-便|出荷日  　　-便|納入先 |受 |納 番 |頁|受注区||品　番　　　　　　　|納品書管理NO|納品書頁|Select|"
            Rows            =   24
            Cols            =   16
            BackColor       =   12648447
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   300
            GridLines       =   3
            SelectionMode   =   1
            BackColorAlternate=   12648447
         End
      End
   End
   Begin VB.TextBox Dummy 
      Height          =   330
      IMEMode         =   3  'ｵﾌ固定
      Left            =   90
      TabIndex        =   12
      Text            =   "Dummy"
      Top             =   420
      Width           =   240
   End
   Begin VB.Timer Timer1 
      Left            =   900
      Top             =   -1380
   End
   Begin VB.Timer Timer2 
      Left            =   1170
      Top             =   -1380
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 検収入力 [一括)]  】"
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
      cPositionX      =   3000
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   661
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "XXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cBoderWidth     =   3
         cFont3D         =   2
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14340
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
         Left            =   11850
         Top             =   45
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
   Begin VB.PictureBox H3_Area1 
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
      Height          =   1275
      Left            =   3315
      ScaleHeight     =   1215
      ScaleWidth      =   6750
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1455
      Width           =   6810
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   1110
         Left            =   60
         Top             =   60
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "      出 荷 日    - 便"
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
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   345
            Index           =   1
            Left            =   90
            Top             =   315
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   609
            ForeColor       =   0
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   345
            Index           =   2
            Left            =   90
            Top             =   675
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   609
            ForeColor       =   0
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin CisText_V60.CisText iH3_SNBin 
            Height          =   375
            Left            =   2475
            TabIndex        =   2
            Top             =   300
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            cFormat         =   "00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "00"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin CisText_V60.CisText iH3_ENBin 
            Height          =   375
            Left            =   2475
            TabIndex        =   4
            Top             =   660
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            cFormat         =   "00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "00"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin CisYMD_With_Btn.CisYMDwB iH3_ENYmd 
            Height          =   360
            Left            =   615
            TabIndex        =   3
            Top             =   660
            Width           =   1860
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
            Object.Height          =   360
            Object.Width           =   1860
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CisYMD_With_Btn.CisYMDwB iH3_SNYmd 
            Height          =   360
            Left            =   615
            TabIndex        =   1
            Top             =   300
            Width           =   1860
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
            Object.Height          =   360
            Object.Width           =   1860
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   1110
         Left            =   3015
         Top             =   60
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "納 入 先"
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
         Begin CisText_V60.CisText iH3_Torcd 
            Height          =   375
            Left            =   90
            TabIndex        =   5
            Top             =   450
            Width           =   1005
            _ExtentX        =   1773
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
         Begin Cis3D_v60.CIS3D H3lb_Tornm 
            Height          =   360
            Left            =   1080
            Top             =   450
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
   End
   Begin Cis3D_v60.CIS3D ErrMsg 
      Height          =   465
      Left            =   3315
      Top             =   945
      Visible         =   0   'False
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   820
      BackColor       =   16761087
      Caption         =   "エラー"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   14.25
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
      cPositionX      =   90
   End
End
Attribute VB_Name = "CXJ0310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   入荷実績入力
'**       フォームID    :   CXH0210
'**       処理概要      :
'**
'**       作  成  日    :   2004/02/06  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim CisVsGrid               As New CisVsGrid3
    Dim SL_Select               As String
    Dim mSL_Select              As String
    Dim mSL_Group               As String
    Dim mSL_Order               As String
    
    Dim mHeadProc               As String
    
    Dim mKanCnt                 As Long
    Dim mBunCnt                 As Long
    Dim mMiCnt                  As Long
    
    Dim QRReadFlg               As Boolean
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    If Rv_NCall = "" Then
    '   #-------------------#
    '   # 二 重 起 動 防 止  #
    '   #-------------------#
        If Not Dupli_Start Then End
    '   #-------------------#
    '   # 初 期 内 容 取 得  #
    '   #-------------------#
        If Not IniGet Then End
    End If
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    QRDummy.Left = -1000
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 600    '
    Timer2.Enabled = True    '

    H2_Area1.Move H3_Area1.Left, H3_Area1.Top
    H4_Area1.Move H3_Area1.Left, H3_Area1.Top
    H5_Area1.Move H3_Area1.Left, H3_Area1.Top

    Set QRr_Object = PDQComm1
    Set QRr_MsgObject = QRMsg
    Set QRr_ErrMsgObj = ErrMsg
    '   ( QR Reader 環境 )
    Call QRInit
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = VsGrid1
        .NoSpaceCol = 2
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 14
        .InitGet = True
        .Init
    End With
    Call GridInit
    
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
    
    With CisFun
        .INI_FileID = gIniExe & gIniName
        .INI_Section = "実績入力"
        .INI_Key = "入荷"
        .INI_Default = ""
        
        .GetIni
    
        If RTrim$(.INI_String) = "" Then
            .INI_String = "1"
            .PutIni
        End If
        
        gInt = .Val2(.INI_String)
    End With
    
    If gInt <> 0 Then gInt = 1
    

    iH1_NYmd = Format(Now(), "YYYYMMDD")

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    
    K_Sykbnm.Tag = "2"
    Call PB_CHG_Click

'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    If Rv_NCall = "" Then
       Me.Show
    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    mSL_Select = "SELECT HC.納品書管理NO,HC.納入日,HC.納入便,"
    mSL_Select = mSL_Select & "  ISNULL(NK.入荷日,'') 入荷日,ISNULL(NK.入荷便,'') 入荷便,"
    mSL_Select = mSL_Select & "  HC.仕入先,HC.受入,HC.納入番号,HC.納品書頁,"
    mSL_Select = mSL_Select & "  HC.発注区分,HC.手配区分,"
    mSL_Select = mSL_Select & "  ISNULL(TR.略称,'') 仕入先名,"
    mSL_Select = mSL_Select & "  ISNULL(SY1.値名称,'') 発注区分名,"
    mSL_Select = mSL_Select & "  ISNULL(SY2.値名称,'') 手配区分名,"
    mSL_Select = mSL_Select & "  Count(表示品番) 件数,"
    mSL_Select = mSL_Select & "  MAX(HN.表示品番) 表示品番,"
    mSL_Select = mSL_Select & "  MAX(HN.実数区分) 実数区分 "
    mSL_Select = mSL_Select & " From 発注テーブル HC"
    mSL_Select = mSL_Select & "  Left Outer Join 入荷テーブル NK"
    mSL_Select = mSL_Select & "     On NK.発注管理NO = HC.発注管理NO"
    mSL_Select = mSL_Select & "  Left Outer Join 品番マスタ HN"
    mSL_Select = mSL_Select & "     ON HN.品番 = HC.品番"
    mSL_Select = mSL_Select & "  Left Outer Join 取引先マスタ TR"
    mSL_Select = mSL_Select & "     ON TR.取引先区分 = 1"
    mSL_Select = mSL_Select & "     And TR.取引先CD = HC.仕入先"
    mSL_Select = mSL_Select & "  Left Outer Join 名称マスタ SY1"
    mSL_Select = mSL_Select & "     ON SY1.区分名称 = '発注区分'"
    mSL_Select = mSL_Select & "     And SY1.区分タイプ = ''"
    mSL_Select = mSL_Select & "     And SY1.値 = HC.発注区分"
    mSL_Select = mSL_Select & "  Left Outer Join 名称マスタ SY2"
    mSL_Select = mSL_Select & "     ON SY2.区分名称 = '手配区分'"
    mSL_Select = mSL_Select & "     And SY2.区分タイプ = ''"
    mSL_Select = mSL_Select & "     And SY2.値 = HC.手配区分"
    
    
    mSL_Group = " Group by HC.納品書管理NO,HC.納入日,HC.納入便,NK.入荷日,NK.入荷便,"
    mSL_Group = mSL_Group & "HC.仕入先,HC.受入,HC.納入番号,HC.納品書頁,"
    mSL_Group = mSL_Group & "HC.発注区分 , HC.手配区分, TR.略称, SY1.値名称, SY2.値名称"

    mSL_Order = " Order by HC.納入日,HC.納入便,HC.仕入先,HC.受入,HC.納入番号,HC.納品書頁,"
    mSL_Order = mSL_Order & "HC.発注区分,HC.手配区分"

    ' 入荷実績処理ワーク削除
    With CisDB
        .SQL = "Delete 入荷実績処理ワーク"
        .SQL = .SQL & "    Where 処理区分 = 9"
        .DBExec
    End With
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
        Case vbKeyF1:       If PB_UPD.Visible Then Call PB_UPD_Click        '【更新】
        Case vbKeyF4:       If PB_Bun.Visible Then Call PB_Bun_Click        '【個別】
        Case vbKeyF6:       If PB_CHG.Visible Then Call PB_CHG_Click        '【切替】
'        Case vbKeyF9:       If PB_QR.Visible Then Call PB_QR_Click          '【読取】
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click      '【検索】
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click        '【取消】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click        '【終了】
        Case vbKeyReturn
                            If PB_Ent.Visible Then                          '【入力】
                                Call PB_ENT_Click
                            End If
        Case Else:          Exit Function
    End Select
    
    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    If Rv_NCall = "" Then
        gSL_Select = "Select Count(*) 件数 From 入荷実績処理ワーク"
        gSL_Select = gSL_Select & " Where 処理端末 = Host_Name()"
        gSL_Select = gSL_Select & "   And 処理区分 = 0"
        If NJWRead(gSL_Select, 1) Then
          If NJW.件数 <> 0 Then
          ' 発注処理
            On Error Resume Next
            CisDB.PassLink
            With CisFun
                .ShellApl = gIniExe & "PXH0110.exe"
                .ShellAplPara = "J"
                .ShellTask
            End With
            On Error GoTo 0
          End If
        End If
        
        Call CisDB.DBDISConnect                         ' ﾃﾞｰﾀﾍﾞｰｽ切断
        Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
        End
    Else
        Me.Hide
    End If
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
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "B1" Then
        Call HeadBodyClear("B")
        ProcHB = mHeadProc
        Call DispChange(ProcHB)
        ' 表示件数カウント初期化
        mKanCnt = 0
        mBunCnt = 0
        mMiCnt = 0
        GoTo PB_CAN_Ed
    End If
    
    If ProcHB = "H2" Or ProcHB = "H3" Or ProcHB = "H4" Or ProcHB = "H5" Then
        ProcHB = "H1"
        Call HeadBodyClear("H2")
        Call HeadBodyClear("H3")
        Call HeadBodyClear("H5")
        Call HeadBodyClear("B")
        Call DispChange(ProcHB)
        GoTo PB_CAN_Ed
    End If
    
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+---------------------+
'+    一括画面切替     +
'+---------------------+
Private Sub PB_Bun_Click()
    Rv_NCall = "K2"
    Rv_NTtlSu = 0
        
    Timer1.Enabled = False
    Call QR_Comm_Close
    
    CXH0220.Show vbModal
    Unload CXH0220
    Set CXH0220 = Nothing
    
    Timer1.Enabled = True
    Rv_NCall = ""
    
    '   ( QR Reader 環境 )
    Set QRr_Object = PDQComm1
    Set QRr_MsgObject = QRMsg
    Set QRr_ErrMsgObj = ErrMsg
    Call QRInit
    Call DispChange(ProcHB)
    
End Sub
'+------------------+
'+    切替   (F6)   +
'+------------------+
Private Sub PB_CHG_Click()
    gLong = CisFun.Val2(K_Sykbnm.Tag)
    
    gLong = gLong + 1
    If gLong > 4 Then gLong = 1
    
    gStr = ""
    gStr = "H3"
    If IRN_NKRead = 1 Then
       gStr = gStr & "H4"
    Else
       gStr = gStr & "  "
    End If
    If IRN_NKReadBar = 1 Then
       gStr = gStr & "H5"
    Else
       gStr = gStr & "  "
    End If
    gStr = gStr & "H2"
    
    mHeadProc = ""
    For gInt = 1 To 4
        mHeadProc = Mid(gStr, gLong + (gLong - 1), 2)
        If Trim(mHeadProc) <> "" Then
           K_Sykbnm.Tag = gLong
           Exit For
        End If
        gLong = gLong + 1
        If gLong > 4 Then gLong = 1
    Next gInt
    
    Select Case mHeadProc
        Case "H3"
            K_Sykbnm.Caption = "【 一覧 】"
            H2_Area1.Visible = False
            H3_Area1.Visible = True
            H4_Area1.Visible = False
            H5_Area1.Visible = False
        Case "H4"
            K_Sykbnm.Caption = "【 QRﾘｰﾀﾞｰ 】"
            H2_Area1.Visible = False
            H3_Area1.Visible = False
            H4_Area1.Visible = True
            H5_Area1.Visible = False
        Case "H5"
            K_Sykbnm.Caption = "【 ﾊﾞｰｺｰﾄﾞ 】"
            H2_Area1.Visible = False
            H3_Area1.Visible = False
            H4_Area1.Visible = False
            H5_Area1.Visible = True
        Case Else
            K_Sykbnm.Caption = "【 納番 】"
            H2_Area1.Visible = True
            H3_Area1.Visible = False
            H4_Area1.Visible = False
            H5_Area1.Visible = False
    End Select
    If ProcHB <> "H1" Then
       ProcHB = mHeadProc
       Call DispChange(ProcHB)
       Select Case mHeadProc
            Case "H2": iH2_NYmd.SetFocus
            Case "H3"
                iH3_SNYmd.SetFocus
                Call HeadBodyClear("H2")
                Call HeadBodyClear("H3")
                Call HeadBodyClear("H5")
                Call HeadBodyClear("B")
                ' 表示件数カウント初期化
                mKanCnt = 0
                mBunCnt = 0
                mMiCnt = 0
       End Select
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_NYmd" Then iH1_NYmd.ShowCalender
    If PB_Look.Tag = "iH2_NYmd" Then iH2_NYmd.ShowCalender
    If PB_Look.Tag = "iH3_SNYmd" Then iH3_SNYmd.ShowCalender
    If PB_Look.Tag = "iH3_ENYmd" Then iH3_ENYmd.ShowCalender
    If PB_Look.Tag = "iH3_Torcd" Then Call Look_TORCD(1)
End Sub
'+-------------------------+
'+    取　引　先　検　索    +
'+-------------------------+
Private Sub Look_TORCD(pKB As Byte)

    RV_Left = 0:    RV_Top = 0
    RV_TorKb = 8
    
    CKK0025.Show vbModal

    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH3_Torcd = RTrim$(RV_TorcdK)
        Call iH3_Torcd_LostFocus
    End If
    
    iH3_Torcd.SetFocus
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*--------------------------------------*
'*      入 荷 日    ( iH1_NYmd )        *
'*--------------------------------------*
Private Sub iH1_NYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_NYmd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*----------------------------------------------*
'*      納 入 日  ( iH2_NYmd )       *
'*----------------------------------------------*
Private Sub iH2_NYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH2_NYmd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*----------------------------------------------*
'*      ( 開 始 ) 納 入 日  ( iH3_SNYmd )       *
'*----------------------------------------------*
Private Sub iH3_SNYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH3_SNYmd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*----------------------------------------------*
'*      ( 終 了 ) 納 入 日  ( iH3_ENYmd )       *
'*----------------------------------------------*
Private Sub iH3_ENYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH3_ENYmd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*----------------------------------*
'*      取引先  ( iH3_Torcd )       *
'*----------------------------------*
Private Sub iH3_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH3_Torcd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    TRM.略称 = ""
    If RTrim$(iH3_Torcd) = "" Then
        H3lb_Tornm = ""
    Else
        Call TorNmGet(RTrim$(iH3_Torcd), 8)
        H3lb_Tornm = RTrim$(TRM.略称)
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = mHeadProc
        Select Case mHeadProc
            Case "H3"
               Call HeadBodyClear("H3")
               iH3_ENYmd = iH1_NYmd
               iH3_ENBin = iH1_NBin
               If Trim(iH3_ENYmd) = "" Then iH3_ENYmd = Format(Now(), "YYYYMMDD")
        End Select
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "H2" Then
       If Not Head2Chk Then GoTo ReturnPress_Ed
       GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "H3" Then
        If Not Head3Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        VsGrid1.SetFocus
        GoTo ReturnPress_Ed
    End If

    If ProcHB = "H4" Then
       If Not Head4Chk Then GoTo ReturnPress_Ed
       GoTo ReturnPress_Ed
    End If
    If ProcHB = "H5" Then
       If Not Head5Chk Then GoTo ReturnPress_Ed
       GoTo ReturnPress_Ed
    End If
    
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
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
    Dim H1Mode          As Boolean
    Dim H1Color         As Long
    Dim H2Mode          As Boolean
    Dim H2Color         As Long
    Dim H3Mode          As Boolean
    Dim H3Color         As Long
    Dim H4Mode          As Boolean
    Dim H4Color         As Long
    Dim H5Mode          As Boolean
    Dim H5Color         As Long
    Dim B1Mode          As Boolean
    Dim B1Color         As Long
    
    H1Mode = False:     H1Color = gPLostSel
    H2Mode = False:     H2Color = gPLostSel
    H3Mode = False:     H3Color = gPLostSel
    H4Mode = False:     H4Color = gPLostSel
    H5Mode = False:     H5Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "H2":      H2Mode = True:  H2Color = gPGotSel
        Case "H3":      H3Mode = True:  H3Color = gPGotSel
        Case "H4":      H4Mode = True:  H4Color = gPGotSel
        Case "H5":      H5Mode = True:  H5Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    H2_Area1.Enabled = H2Mode
    H2_Area1.BackColor = H2Color
    H3_Area1.Enabled = H3Mode
    H3_Area1.BackColor = H3Color
    H4_Area1.Enabled = H4Mode
    H4_Area1.BackColor = H4Color
    H5_Area1.Enabled = H5Mode
    H5_Area1.BackColor = H5Color
    
    B1_Area1.Enabled = B1Mode Or H2Mode Or H3Mode Or H4Mode Or H5Mode
    B1_Area1.BackColor = B1Color
    
    PB_CAN.Visible = Not H1Mode
    PB_CHG.Visible = Not B1Mode
    PB_Look.Visible = H1Mode
    PB_Bun.Visible = H1Mode
    
    PB_UPD.Visible = B1Mode
    If H2Mode Or H4Mode Or H5Mode Then
        If mBunCnt <> 0 Or mKanCnt <> 0 Then
            PB_UPD.Visible = True
        End If
    End If
    
    PB_Ent.Visible = Not (H4Mode Or B1Mode)
    
    If ProcHB = "H1" Then
       B1lb_Kensu = ""
       B1lb_KKensu = ""
       B1lb_BKensu = ""
       B1lb_MKensu = ""
       PB_Look.Visible = False
    End If
'
    If ProcHB = "H4" Then
       Call QR_Comm_Open
    Else
       Call QR_Comm_Close
    End If

' ﾀﾞﾐｰ項目(入力禁止)
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
        End With
    Next gObj
    If HeadBody = "B" Then
       CisVsGrid.Clear
    End If
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
'( 入荷日 )
    If RTrim$(iH1_NYmd) = "" And CisFun.Val2(iH1_NBin) <> 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    入荷日を入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_NYmd.SetFocus
        Exit Function
    Else
        If RTrim$(iH1_NYmd) <> "" And Not iH1_NYmd.cChkResult Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    正しい日付を入力して下さい         "
                 .MB_Button = Error
                 .MBOX
             End With
             iH1_NYmd.SetFocus
             Exit Function
         End If
    End If
    
'( 入荷便 )
    If RTrim$(iH1_NYmd) <> "" And CisFun.Val2(iH1_NBin) = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    入荷便を入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_NBin.SetFocus
        Exit Function
    End If
    ' 表示件数カウント初期化
    mKanCnt = 0
    mBunCnt = 0
    mMiCnt = 0
    
    Head1Chk = True
End Function
'*-----------------------------------------*
'*      HEAD2ﾁｪｯｸ処理 ( 納入日～頁 )       *
'*-----------------------------------------*
Private Function Head2Chk() As Boolean
    Head2Chk = False
    
'( 納入日 )
    If RTrim$(iH2_NYmd) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  納入日を入力して下さい  "
            .MB_Button = Error
            .MBOX
        End With
        iH2_NYmd.SetFocus
        Exit Function
    Else
        If Not iH2_NYmd.cChkResult Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    正しい日付を入力して下さい         "
                 .MB_Button = Error
                 .MBOX
             End With
             iH2_NYmd.SetFocus
             Exit Function
         End If
    End If
    
'( 納入番号 )
    If RTrim$(iH2_Nouno) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  納番を入力して下さい  "
            .MB_Button = Error
            .MBOX
        End With
        iH2_Nouno.SetFocus
        Exit Function
    End If
    
'( 納品書頁 )
    If CisFun.Val2(iH2_NPage) = 0 Then iH2_NPage = " 1"
    
'( データ確認 )
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where HC.納入番号 = '" & RTrim$(iH2_Nouno) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   HC.納品書頁 =  " & CisFun.Val2(iH2_NPage) & vbCrLf
    gSL_Select = gSL_Select & " And   HC.納入日   = '" & RTrim$(iH2_NYmd) & "'" & vbCrLf
    gSL_Select = gSL_Select & mSL_Group
    
    If Not HCTRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  該当データがありません  "
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
    
    If Not GridSetKobetsu Then Exit Function
    
    Head2Chk = True
End Function
'*-----------------------------------------*
'*      HEAD3ﾁｪｯｸ処理 ( 納入日～実績 )     *
'*-----------------------------------------*
Private Function Head3Chk() As Boolean
    Dim DateStr As String
    Dim DateSt  As String
    Dim DateEd  As String
    Dim BinSt   As String
    Dim BinEd   As String
    
    Head3Chk = False
    
    gSL_Select = mSL_Select
    With CisFun
         .Where_CreateKB = NewDocument
         ' 日付
         DateStr = "納入"
         DateSt = "": DateEd = "": BinSt = "": BinEd = ""
         ' 日付 月指定
         If Trim(iH3_SNYmd) <> "" And CisFun.Val2(iH3_SNYmd.cDay) = 0 Then
            .Where_Create AlphaMe, "Substring(" & DateStr & "日,1,6)", iH3_SNYmd
            .Where_Create AlphaMe, DateStr & "便", Format(iH3_SNBin, "00")
         End If
         ' 日付開始編集
         If Trim(iH3_SNYmd) <> "" Then
            DateSt = iH3_SNYmd
         End If
         If CisFun.Val2(iH3_SNBin) <> 0 Then
            BinSt = Format(iH3_SNBin, "00")
         End If
         ' 日付終了編集
         If Trim(iH3_ENYmd) <> "" Then
            DateEd = iH3_ENYmd
         End If
         If CisFun.Val2(iH3_ENBin) <> 0 Then
            BinEd = Format(iH3_ENBin, "00")
         End If
         ' 日付開始のみ
         If Trim(DateSt) <> "" And Trim(DateEd) = "" And CisFun.Val2(iH3_SNYmd.cDay) <> 0 Then
            If Trim(BinSt) = "" Then
               .Where_Create AlphaMe, DateStr & "日", DateSt
            Else
               .Where_Create AlphaMe, DateStr & "日+" & DateStr & "便", DateSt & BinSt
            End If
         End If
         ' 日付終了のみ
         If Trim(DateSt) = "" And Trim(DateEd) <> "" Then
            If Trim(BinEd) = "" Then
               .Where_Create AlphaMe, DateStr & "日", DateEd, 以下
            Else
               .Where_Create AlphaMe, DateStr & "日+" & DateStr & "便", DateEd & BinEd, 以下
            End If
         End If
         ' 日付開始・終了
         If Trim(DateSt) <> "" And Trim(DateEd) <> "" Then
            If Trim(BinSt) = "" And Trim(BinEd) = "" Then
               .Where_Create AlphaMe, DateStr & "日", DateSt, Between, , DateEd
            Else
               If Trim(BinSt) = "" Then BinSt = "00"
               If Trim(BinEd) = "" Then BinEd = "99"
               .Where_Create AlphaMe, DateStr & "日+" & DateStr & "便", DateSt & BinSt, Between, , DateEd & BinEd
            End If
         End If
         ' 仕入先
         .Where_Create AlphaMe, "HC.仕入先", iH3_Torcd
         ' 未入荷のみ
         .Where_Create AlphaMe, "ISNULL(入荷日,'')", "", 等しい, Create
         
         gSL_Select = gSL_Select & .Where_Phrase
    End With
    
    gSL_Select = gSL_Select & mSL_Group & mSL_Order
    
    PNL_Wait.Visible = True
    Message = "【抽出中】しばらくお待ち下さい"
    DoEvents
    
    If Not HCTRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  該当データがありません  "
            .MB_Button = Error
            .MBOX
        End With
        PNL_Wait.Visible = False
        Exit Function
    End If
    
    Call GridSet
    
    PNL_Wait.Visible = False
    
    Head3Chk = True
End Function
'*-----------------------------------------*
'*      HEAD4ﾁｪｯｸ処理 ( ＱＲ )             *
'*-----------------------------------------*
Private Function Head4Chk() As Boolean
    Head4Chk = False
    
    
    Head4Chk = True
End Function
'*-----------------------------------------*
'*      HEAD5ﾁｪｯｸ処理 ( バーコード )       *
'*-----------------------------------------*
Private Function Head5Chk() As Boolean
    Head5Chk = False
    
'( バーコード )
    If CisFun.Val2(iH5_Bar) = 0 Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  バーコードを入力して下さい  "
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
    
    
'( データ確認 )
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where HC.納品書管理NO = " & CisFun.Val2(Mid(iH5_Bar, 1, 10))
    gSL_Select = gSL_Select & "   And HC.納品書頁 = " & CisFun.Val2(Mid(iH5_Bar, 11, 2))
    gSL_Select = gSL_Select & mSL_Group
    
    If Not HCTRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  該当データがありません  "
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
    
    If Not GridSetKobetsu Then Exit Function
    
    Head5Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
'    If iB1_???? = "" Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    ???を入力して下さい         "
'            .MB_Button = Error
'            .MBOX
'        End With
'        iB1_???.SetFocus
'        Exit Function
'    End If
    
'   < 印刷データ確認 >
'    gSL_Select = "SELECT * FROM ??テーブル"
'    gSL_Select = gSL_Select & " WHERE Key = '" & RTrim(iH1_NYmd) & "'"
    
'    If Not NKTRead(gSL_Select, 1) Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    該当データがありません            "
'            .MB_Button = Error
'            .MBOX
'        End With
''        iB1_???.SetFocus
'        Exit Function
'    End If
    
    Body1Chk = True
End Function
'****************************
'*      ボディ内容セット
'****************************
Private Function BodySet()

End Function
'*--------------------------*
'*      処　理　開　始      *
'*--------------------------*
Private Function Proc_Start() As Boolean
    Proc_Start = False
    
' << 印刷処理開始 >>
'    With CisPrt
''        Set .iPrtForm = LXX0000
    
'        .dpMSG = "【  印 刷 中  】しばらくお待ち下さい"
'        .dpWidth = 0
'        .iPreview = pv_Preview
'        .PrintStart
'    End With
    
    Proc_Start = True
End Function
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆
'☆                         入　　荷　　実　　績　　更　　新　　(  F1  )
'☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
Private Sub PB_UPD_Click()
    
    gCnt1 = 0
    With VsGrid1
        For gCnt2 = 0 To .Rows - 1
            If Trim(.TextMatrix(gCnt2, 1)) = "完納" Or _
               Trim(.TextMatrix(gCnt2, 1)) = "分納" Then
               gCnt1 = 1
               Exit For
            End If
        Next gCnt2
    End With
        
    If gCnt1 = 0 Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "      処理を選択して下さい               "
            .MB_Button = OK_CAN
            .MBOX
            Exit Sub
        End With
    End If
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "      入荷実績更新を行います。               "
        .MB_MSG(4) = "          よろしいですか？                      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Sub
    End With

    On Error GoTo PB_UPD_Er:
    

    PNL_Wait.Visible = True
    Message = "【実績更新中】しばらくお待ち下さい"
    DoEvents
    '+---------------
    '+   完納処理
    '+---------------
    Call CisDB.DBTran(TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    With VsGrid1
        .Redraw = False
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, 1)) = "完納" Then
               With CisDB
                    .SQL = "Insert Into 入荷テーブル  ( " & vbCrLf
                    .SQL = .SQL & "発注管理NO,納入番号,納品書頁," & vbCrLf
                    .SQL = .SQL & "納品書行,分納区分,仕入先,受入," & vbCrLf
                    .SQL = .SQL & "品番,背番号,入荷日," & vbCrLf
                    .SQL = .SQL & "入荷便,収容数,入荷枚数," & vbCrLf
                    .SQL = .SQL & "入荷端数,入荷数,分割数," & vbCrLf
                    .SQL = .SQL & "挽回日,挽回便,分割納品書NO," & vbCrLf
                    .SQL = .SQL & "検収区分,検収年月,検収処理日," & vbCrLf
                    .SQL = .SQL & "作成日,作成端末,作成者 )" & vbCrLf
                    
                    .SQL = .SQL & "Select 発注管理NO,納入番号,納品書頁," & vbCrLf
                    .SQL = .SQL & "納品書行,0,仕入先,受入," & vbCrLf
                    .SQL = .SQL & "品番,背番号," & vbCrLf
                    If Trim(iH1_NYmd) <> "" Then
                       .SQL = .SQL & "'" & iH1_NYmd & "'," & vbCrLf
                    Else
                       .SQL = .SQL & "納入日," & vbCrLf
                    End If
                    .SQL = .SQL & "Case When 納入便 = '' Then '' Else " & vbCrLf
                    If Trim(iH1_NYmd) <> "" Then
                       .SQL = .SQL & "'" & iH1_NBin & "' End," & vbCrLf
                    Else
                       .SQL = .SQL & "納入便 End," & vbCrLf
                    End If
                    .SQL = .SQL & "収容数,枚数,端数,発注数,0," & vbCrLf
                    .SQL = .SQL & "'','','','','',''," & vbCrLf
                    .SQL = .SQL & "Getdate(),Host_Name(),'" & gTanto & "'" & vbCrLf
                    .SQL = .SQL & " From 発注テーブル "
                    .SQL = .SQL & " Where 納品書管理NO = " & CisFun.Val2(VsGrid1.TextMatrix(gLong, 12))
                    .SQL = .SQL & "   And 納品書頁 = " & CisFun.Val2(VsGrid1.TextMatrix(gLong, 13))
                    
                    If Not .DBExec Then GoTo PB_UPD_Er
                    
               
                    .SQL = "Insert Into 入荷実績処理ワーク  ( "
                    .SQL = .SQL & "処理端末,発注管理NO,処理区分,作成日 )"
                    .SQL = .SQL & " Select Host_Name(),発注管理NO,0,Getdate() "
                    .SQL = .SQL & " From 発注テーブル "
                    .SQL = .SQL & " Where 納品書管理NO = " & CisFun.Val2(VsGrid1.TextMatrix(gLong, 12))
                    .SQL = .SQL & "   And 納品書頁 = " & CisFun.Val2(VsGrid1.TextMatrix(gLong, 13))
                    
                    If Not .DBExec Then GoTo PB_UPD_Er
               End With
               .TextMatrix(gLong, 1) = "完済"
               TTL_KSKensu = Format(CisFun.Val2(TTL_KSKensu) + 1, "#,###")
               .Row = gLong
               Call CisVsGrid.vsColor(消去個別)
               DoEvents
            End If
        Next gLong
        .Row = 1: .Col = 1: .TopRow = 1
        .Redraw = True
    End With
    Call CisDB.DBTran(TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ終了
    ' 分納処理
    With VsGrid1
        .Redraw = False
        Erase Rv_HNo
        Erase Rv_HPage
        Erase Rv_NExec
        Erase Rv_NRow
        
        Rv_NTtlSu = 0
        
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, 1)) = "分納" Then
               If CisFun.Val2(VsGrid1.TextMatrix(gLong, 12)) <> 0 Then
                  Rv_NTtlSu = Rv_NTtlSu + 1
                  ReDim Preserve Rv_HNo(Rv_NTtlSu)
                  ReDim Preserve Rv_HPage(Rv_NTtlSu)
                  ReDim Preserve Rv_NExec(Rv_NTtlSu)
                  ReDim Preserve Rv_NRow(Rv_NTtlSu)
                  Rv_HNo(Rv_NTtlSu) = CisFun.Val2(VsGrid1.TextMatrix(gLong, 12))
                  Rv_HPage(Rv_NTtlSu) = CisFun.Val2(VsGrid1.TextMatrix(gLong, 13))
                  Rv_NRow(Rv_NTtlSu) = gLong
               Else
                  .TextMatrix(gLong, 1) = ""
                  mBunCnt = mBunCnt - 1
                  mMiCnt = mMiCnt + 1
                  B1lb_BKensu = Format(mBunCnt, "#,###")
                  B1lb_MKensu = Format(mMiCnt, "#,###")
                  .Row = gLong
                  Call CisVsGrid.vsColor(消去個別)
               End If
            End If
        Next gLong
        .Redraw = True
    End With
    
    If Rv_NTtlSu <> 0 Then
        Rv_NCall = "K"
        Rv_NYmd = iH1_NYmd
        Rv_NBin = iH1_NBin
        
        If ProcHB = "H4" Then
           Call QR_Comm_Open
        Else
           Call QR_Comm_Close
        End If
        
        CXH0220.Show vbModal
        Unload CXH0220
        Set CXH0220 = Nothing
        Rv_NCall = ""
        
        '   ( QR Reader 環境 )
        Set QRr_Object = PDQComm1
        Set QRr_MsgObject = QRMsg
        Set QRr_ErrMsgObj = ErrMsg
        Call QRInit
        
        If ProcHB = "H4" Then
           Call QR_Comm_Open
        Else
           Call QR_Comm_Close
        End If
    
        With VsGrid1
            .Redraw = False
            For gLong = 1 To UBound(Rv_NRow)
                .Row = Rv_NRow(gLong)
                If Rv_NExec(gLong) Then
                   TTL_BSKensu = Format(CisFun.Val2(TTL_BSKensu) + 1, "#,###")
                   .TextMatrix(Rv_NRow(gLong), 1) = "分済"
                Else
                   .TextMatrix(Rv_NRow(gLong), 1) = ""
                   mBunCnt = mBunCnt - 1
                   mMiCnt = mMiCnt + 1
                   B1lb_BKensu = Format(mBunCnt, "#,###")
                   B1lb_MKensu = Format(mMiCnt, "#,###")
                End If
                Call CisVsGrid.vsColor(消去個別)
            Next gLong
            .Row = 1: .Col = 1: .TopRow = 1
            .Redraw = True
       End With
    End If
PB_UPD_Ed:
    PNL_Wait.Visible = False
    PB_UPD.Visible = False
'    Timer1.Enabled = True
    Exit Sub
PB_UPD_Er:
    VsGrid1.Redraw = True
    PNL_Wait.Visible = False
    Call CisDB.DBTran(TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ破棄

End Sub
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
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 中中, あり, 5, 1, "処理")
        Call .FixedSet(中中, 中中, あり, 13, 2, "納入日-便")
        Call .FixedSet(中中, 中中, あり, 13, 3, "入荷日-便")
        Call .FixedSet(中中, 左中, あり, 27, 4, "仕入先")
        Call .FixedSet(中中, 中中, あり, 2, 5, "受")
        Call .FixedSet(中中, 中中, あり, 5, 6, "納 番")
        Call .FixedSet(中中, 右中, あり, 2, 7, "頁")
        Call .FixedSet(中中, 中中, あり, 5, 8, "発注")
        Call .FixedSet(中中, 中中, あり, 5, 9, "手配")
        Call .FixedSet(中中, 中中, あり, 2, 10, "実")
        Call .FixedSet(中中, 左中, あり, 25, 11, "品　番")
        Call .FixedSet(中中, 左中, あり, 0, 12, "納品書管理NO")
        Call .FixedSet(中中, 左中, あり, 0, 13, "納品書頁")

'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 14, "Select")

'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11")

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
    
    With VsGrid1
        .Redraw = False
        For gLong = .Row To .RowSel
            If Trim(.TextMatrix(gLong, CisVsGrid.NoSpaceCol)) <> "" Then
                gStr = Trim(.TextMatrix(gLong, CisVsGrid.SelectCol))
                .Row = gLong
                Select Case Trim(.TextMatrix(gLong, 1))
                    Case ""
                        If gStr = "X" Then
                           Call CisVsGrid.vsColor(消去個別)
                        End If
                        If Trim(.TextMatrix(gLong, 10)) = "○" Then
                            Call CisVsGrid.vsColor(個別, B1lb_BKensu.BackColor)
                            .TextMatrix(gLong, 1) = "分納"
                            mBunCnt = mBunCnt + 1
                        Else
                            Call CisVsGrid.vsColor(個別, B1lb_KKensu.BackColor)
                            mKanCnt = mKanCnt + 1
                            .TextMatrix(gLong, 1) = "完納"
                        End If
                        mMiCnt = mMiCnt - 1
                    Case "完納"
                        If gStr = "X" Then
                           Call CisVsGrid.vsColor(消去個別)
                        End If
                        If Right(.TextMatrix(gLong, 6), 1) = "9" Then
                           .TextMatrix(gLong, 1) = ""
                           mKanCnt = mKanCnt - 1
                           mMiCnt = mMiCnt + 1
                        Else
                           Call CisVsGrid.vsColor(個別, B1lb_BKensu.BackColor)
                           .TextMatrix(gLong, 1) = "分納"
                           mKanCnt = mKanCnt - 1
                           mBunCnt = mBunCnt + 1
                        End If
                    Case "分納"
                        If gStr = "X" Then
                           Call CisVsGrid.vsColor(消去個別)
                        End If
                        .TextMatrix(gLong, 1) = ""
                        mBunCnt = mBunCnt - 1
                        mMiCnt = mMiCnt + 1
                End Select
            End If
        Next gLong
        .Redraw = True
    End With
    
    B1lb_KKensu = Format(mKanCnt, "#,###")
    B1lb_BKensu = Format(mBunCnt, "#,###")
    B1lb_MKensu = Format(mMiCnt, "#,###")
    
    If mBunCnt <> 0 Or mKanCnt <> 0 Then
        PB_UPD.Visible = True
    Else
        PB_UPD.Visible = False
    End If

End Sub
'******************************************'
'*****        行選択（スペースｷｰ)      *****
'******************************************'
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    Call VSGrid1_Click
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid1
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not HCT_RDSTS
        With VsGrid1
            ID = ID + 1
       
            Call GridMove(ID)
        End With
        Call HCTReadNext
    Loop
    Call HCTClose
    
    Call CisVsGrid.ReNumber
    
    With VsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With

    B1lb_KKensu = Format(mKanCnt, "#,###")
    B1lb_BKensu = Format(mBunCnt, "#,###")
    B1lb_MKensu = Format(mMiCnt, "#,###")
End Sub
'+------------------------------------+
'+      グリッドセット（個別)
'+------------------------------------+
Private Function GridSetKobetsu() As Boolean
    GridSetKobetsu = False
    
    If HCT.入荷日 <> "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "      入荷済みデータです              "
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
    With VsGrid1
        .Redraw = False
        
        For gCnt1 = 1 To .Rows - 1
            gCnt2 = CisFun.Val2(.TextMatrix(gCnt1, 12))
            gCnt3 = CisFun.Val2(.TextMatrix(gCnt1, 13))
            If gCnt2 <> 0 And gCnt2 = HCT.納品書管理NO And gCnt3 = HCT.納品書頁 Then
               With CisFun
                   .MB_Lines = 5
                   .MB_MSG(2) = "      表示済みです。               "
                   .MB_Button = Error
                   .MBOX
               End With
               .Redraw = True
               Exit Function
            End If
        Next gCnt1
        
        
        If Trim(.TextMatrix(.Rows - 1, 2)) <> "" Then
           .Rows = .Rows + 1
        End If

        Call GridMove(.Rows - 1)
        .RowPosition(.Rows - 1) = 1
        
        Call CisVsGrid.ReNumber
        
        .Redraw = True
    
    End With

    B1lb_KKensu = Format(mKanCnt, "#,###")
    B1lb_BKensu = Format(mBunCnt, "#,###")
    B1lb_MKensu = Format(mMiCnt, "#,###")
    B1lb_Kensu = Format(CisFun.Val2(B1lb_Kensu) + 1, "#,###")
    If mBunCnt <> 0 Or mKanCnt <> 0 Then
        PB_UPD.Visible = True
    Else
        PB_UPD.Visible = False
    End If

    GridSetKobetsu = True
End Function
'+------------------------------------+
'+      グリッド項目転送
'+------------------------------------+
Private Sub GridMove(mRow As Long)

    With VsGrid1
        
        With CisFun
            .DateE_BefVal = HCT.納入日
            .DateEdit
            gStr = CisFun.DateE_AfVal
            If Trim(HCT.納入便) <> "" Then
                gStr = gStr & "-" & HCT.納入便
            Else
                gStr = gStr & "   "
            End If
        End With
        .TextMatrix(mRow, 2) = gStr
        
        If Trim(iH1_NYmd) = "" Then
            gStr = "============="
        Else
            With CisFun
                .DateE_BefVal = iH1_NYmd
                .DateEdit
                gStr = CisFun.DateE_AfVal
                If Trim(HCT.納入便) <> "" Then
                    gStr = gStr & "-" & Format(iH1_NBin, "00")
                Else
                    gStr = gStr & "   "
                End If
            End With
        End If
        .TextMatrix(mRow, 3) = gStr
        
        If Trim(HCT.仕入先名) = "" Then
           gStr = HCT.仕入先
        Else
           gStr = RTrim(HCT.仕入先) & Space(7 - Len(HCT.仕入先)) & ":"
           gStr = gStr & HCT.仕入先名
        End If
        .TextMatrix(mRow, 4) = gStr
        .TextMatrix(mRow, 5) = HCT.受入
        .TextMatrix(mRow, 6) = HCT.納入番号
        .TextMatrix(mRow, 7) = HCT.納品書頁
        .TextMatrix(mRow, 8) = HCT.発注区分名
        .TextMatrix(mRow, 9) = HCT.手配区分名
        
        If Mid(HCT.納入番号, 5, 1) <> "0" Then
           .TextMatrix(mRow, 11) = HCT.表示品番
        End If
        
        If HCT.実数区分 = 0 Or Right(HCT.納入番号, 1) = "9" Then
           .TextMatrix(mRow, 1) = "完納"
           mKanCnt = mKanCnt + 1
           .Row = mRow
           Call CisVsGrid.vsColor(個別, B1lb_KKensu.BackColor)
        Else
           .TextMatrix(mRow, 10) = "○"
           .Row = mRow
           .TextMatrix(mRow, 1) = "分納"
           mBunCnt = mBunCnt + 1
           Call CisVsGrid.vsColor(個別, B1lb_BKensu.BackColor)
        End If
    
        .TextMatrix(mRow, 12) = HCT.納品書管理NO
        .TextMatrix(mRow, 13) = HCT.納品書頁
    
    End With
End Sub
'*=======================================================================================================*
'*=======================================================================================================*
'*
'*                      Ｑ  Ｒ  リ  ー  ダ  (　Ｃｏｍｍ　)  処  理
'*
'*=======================================================================================================*
'*=======================================================================================================*
'+--------------------------+
'+      読  取  処  理      +
'+--------------------------+
Private Sub PDQComm1_OnComm()
    
    If QR_CommEvent <> "OK" Then Exit Sub
    
    With QRDataArea
        If .バーコード種類 <> "B" Then
            If .ステータス <> 0 Then
               Call QRReadMsg("E")
               Exit Sub
            End If
            ' 部品のみ有効
            If .バーコード種類 <> "D" Then
               .エラー内容 = "部品の受領書でない"
               Call QRReadMsg("E")
               Exit Sub
            End If
            ' 納品書のみ有効
            If .カード <> "7" Then
               .エラー内容 = "カードが対象外です"
               Call QRReadMsg("E")
               Exit Sub
            End If
        End If
    End With
'( データ確認 )
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where HC.納品書管理NO = " & QRDataArea.納品書管理NO
    gSL_Select = gSL_Select & "   And HC.納品書頁 = " & QRDataArea.納品書頁
    gSL_Select = gSL_Select & mSL_Group
    
    If Not HCTRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  該当データがありません  "
            .MB_Button = Error
            .MBOX
        End With
        Exit Sub
    End If
    
    If Not GridSetKobetsu Then Exit Sub
    
End Sub
