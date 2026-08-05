VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "Cis3D_6.0.ocx"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CisBtn_6.0.ocx"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CisText_6.0.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CisYMDwB_6.0.ocx"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Begin VB.Form CXJ0210 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "出荷実績入力"
   ClientHeight    =   11115
   ClientLeft      =   105
   ClientTop       =   345
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
   Begin PdqcommLib.PDQComm PDQComm1 
      Height          =   480
      Left            =   195
      TabIndex        =   16
      Top             =   945
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
   Begin Cis3D_v60.CIS3D H4_Area1 
      Height          =   1320
      Left            =   0
      Top             =   2280
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
         Left            =   3660
         TabIndex        =   15
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
         BackColor       =   16761024
         ForeColor       =   16711680
         Caption         =   "QRﾀｯﾁﾘｰﾀﾞ読取"
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
      Left            =   1815
      ScaleHeight     =   1215
      ScaleWidth      =   1905
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3525
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
            Left            =   210
            TabIndex        =   14
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   635
            cFormat         =   "0000000000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0000000000"
            MaxLength       =   10
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##########"
            cILength        =   10
            cAutoEject      =   0   'False
         End
      End
   End
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   4905
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
      Left            =   10965
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
         Caption         =   "完納件数"
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
         Caption         =   "分納件数"
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
         Caption         =   "完 納 済"
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
         Caption         =   "分 納 済"
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
         Begin Cis3D_v60.CIS3D TTL_BSKensu 
            Height          =   345
            Left            =   1170
            Top             =   15
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
      Left            =   9120
      ScaleHeight     =   690
      ScaleWidth      =   6105
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   525
      Width           =   6165
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   3508
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
         Left            =   4370
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
      Begin CisBtn_60.CisBtn PB_UPD 
         Height          =   585
         Left            =   2646
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
         Left            =   922
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
      Begin CisBtn_60.CisBtn PB_CHG 
         Height          =   585
         Left            =   1784
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
      Left            =   165
      ScaleHeight     =   1215
      ScaleWidth      =   2625
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1455
      Width           =   2685
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   1110
         Left            =   60
         Top             =   60
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "    出 荷 日  - 便"
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_SYmd 
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
         Begin CisText_V60.CisText iH1_SBin 
            Height          =   375
            Left            =   2010
            TabIndex        =   1
            Top             =   450
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
      End
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   7725
      Left            =   165
      ScaleHeight     =   7665
      ScaleWidth      =   15060
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2850
      Width           =   15120
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   7560
         Left            =   60
         Top             =   45
         Width           =   14970
         _ExtentX        =   26405
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
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   105
            Width           =   14790
            _Version        =   196608
            _ExtentX        =   26088
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
            FormatString    =   "№|処理 |納入日 　　-便|出荷日  　　-便|納入先 |受 |出荷指示書№|受注区|||納品書管理NO|納品書頁|出荷管理部署|Select|"
            Rows            =   24
            Cols            =   17
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
      Left            =   165
      TabIndex        =   8
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
      Caption         =   "【 出荷実績入力 [一括)]  】"
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
         Width           =   2970
         _ExtentX        =   5239
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
      Left            =   3165
      ScaleHeight     =   1215
      ScaleWidth      =   6750
      TabIndex        =   11
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
         Caption         =   "      納 入 日    - 便"
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
            TabIndex        =   3
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
            TabIndex        =   5
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
            TabIndex        =   4
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
            TabIndex        =   2
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
            TabIndex        =   6
            Top             =   450
            Width           =   1005
            _ExtentX        =   1773
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
      Left            =   3165
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
Attribute VB_Name = "CXJ0210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   出荷実績入力
'**       フォームID    :   CXJ0210
'**       処理概要      :
'**
'**       作  成  日    :   2004/03/27  By CIS
'**       変  更  日    :   2004/04/16  By CIS - 自社伝票発行＆自社伝票ファイル出力の追加
'**       変  更  日    :   2004/04/21  By CIS - QRﾀｯﾁﾘｰﾀﾞｰ機能追加
'**       変  更  日    :   2004/06/28  By CIS - 受領書・出荷時発行の対応
'**       変  更  日    :   2006/11/28  By CIS - カレンダマスタ読込時 取引先区分＝発注を受注に
'**       変  更  日    :   2009/05/18  By CIS - Toppsﾃﾞｰﾀ非対象
'**       変  更  日    :   2009/09/21  By CIS - 土岐津Ver
'**       変  更  日    :   2009/12/15  By CIS - 土岐津Ver(出庫ﾃﾞｰﾀ(実績管理ﾃｰﾌﾞﾙ)作成)
'**       変  更  日    :   2010/04/22  By CIS - 納入先、ﾜｲﾙﾄﾞｶｰﾄﾞ許可
'**       変  更  日    :   2010/12/28  By CIS - 桁溢れｴﾗｰ対応(wInt → wLong)
'**       変  更  日    :   2011/02/04  By CIS - 12/28修正ﾊﾞｸﾞ対応
'**       変  更  日    :   2011/10/17  By CIS - 収容数０対応（実績管理テーブル）
'**       変  更  日    :   2020/01/10  By CIS - 棚卸ﾌﾟﾚｼｰﾄ区分条件廃止
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
    
    Dim wSJNo                   As Long         '>出荷指示書NO
    Dim wEndChk                 As Boolean

    Dim wInt                    As Integer
    Dim wLong                   As Long         '2010/12/28
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
        gConnectCount = 4
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

    H4_Area1.Move H3_Area1.Left, H3_Area1.Top
    H5_Area1.Move H3_Area1.Left, H3_Area1.Top

'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 2
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 9
        .InitGet = False
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
    
    iH1_SYmd = Format(Now(), "YYYYMMDD")
    K_Sykbnm.Caption = "【 一覧 】"
    
    Set QT10QSR_Comm = PDQComm1
    Call QT10QSRInfoGet
    
    If QT10QSR_UseFlg Then
       IRN_NKRead = 1
    Else
       IRN_NKRead = 0
    End If
    With CisFun
        .INI_Section = "Apparatus"
        .INI_FileID = gIniExe & gIniName
        .INI_Key = "5"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "ｷｰﾎﾞｰﾄﾞｲﾝﾀｰﾌｪｲｽ"
           .PutIni
        End If
        .INI_Section = .INI_String
        ' '----- 使用/未使用 -----
        .INI_Default = ""
        .INI_Key = "Use"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "Yes"
           .PutIni
        End If
        If UCase(Trim(.INI_String)) = "N" Or UCase(Trim(.INI_String)) = "NO" Then
           IRN_NKReadBar = 0
        Else
           IRN_NKReadBar = 1
        End If
    End With
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    K_Sykbnm.Tag = "3"
    Call PB_CHG_Click
    Call DispChange(ProcHB)

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

'*--------------------------------------------------<< 2004/06/28 Insert Start >>
    gCHT_DelKbn53 = 0
'*--------------------------------------------------<< 2004/06/28 Insert End >>
    If InitProc Then        ' 初期処理（データ削除）
        wEndChk = True
    Else
        wEndChk = False
        Call PB_END_Click
    End If

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
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click      '【検索】
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click        '【取消】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click        '【終了】
        Case vbKeyReturn
                            If PB_ENT.Visible Then                          '【入力】
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
'   【 後方処理 】
    '   #### 終了処理 ####
'*---------------------------------------------<< 2004/04/16 Update Start >>
'       If wEndChk Then Call EndProc
        Call EndProc
'*---------------------------------------------<< 2004/04/16 Update End >>
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

    If ProcHB = "H3" Or ProcHB = "H4" Or ProcHB = "H5" Then
        ProcHB = "H1"
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
    Rv_STtlSu = 0
        
    Timer1.Enabled = False
    
    CXJ0220.Show vbModal
    Unload CXJ0220
    Set CXJ0220 = Nothing
    
    Timer1.Enabled = True
    Rv_NCall = ""
    
    Set QT10QSR_Comm = PDQComm1
    Call QT10QSRInfoGet
    
    Call DispChange(ProcHB)
    
End Sub
'+------------------+
'+    切替   (F6)   +
'+------------------+
Private Sub PB_CHG_Click()
    gLong = CisFun.Val2(K_Sykbnm.Tag)
    
    gLong = gLong + 1
    If gLong > 3 Then gLong = 1
    
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
    
    mHeadProc = ""
    For gInt = 1 To 3
        mHeadProc = Mid(gStr, gLong + (gLong - 1), 2)
        If Trim(mHeadProc) <> "" Then
           K_Sykbnm.Tag = gLong
           Exit For
        End If
        gLong = gLong + 1
        If gLong > 3 Then gLong = 1
    Next gInt
    
    H3_Area1.Visible = False
    H4_Area1.Visible = False
    H5_Area1.Visible = False
    
    Select Case mHeadProc
        Case "H4"
            K_Sykbnm.Caption = "【 QRﾀｯﾁﾘｰﾀﾞ 】"
            H4_Area1.Visible = True
        Case "H5"
            K_Sykbnm.Caption = "【 ﾊﾞｰｺｰﾄﾞ 】"
            H5_Area1.Visible = True
        Case Else
            K_Sykbnm.Caption = "【 一覧 】"
            H3_Area1.Visible = True
    End Select
    If ProcHB <> "H1" Then
       ProcHB = mHeadProc
       Call DispChange(ProcHB)
       Select Case mHeadProc
            Case "H3"
                iH3_SNYmd.SetFocus
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
    If PB_Look.Tag = "iH1_SYmd" Then iH1_SYmd.ShowCalender
    If PB_Look.Tag = "iH3_SNYmd" Then iH3_SNYmd.ShowCalender
    If PB_Look.Tag = "iH3_ENYmd" Then iH3_ENYmd.ShowCalender
    If PB_Look.Tag = "iH3_Torcd" Then Call Look_TORCD(1)
End Sub
'+-------------------------+
'+    取　引　先　検　索    +
'+-------------------------+
Private Sub Look_TORCD(pKB As Byte)

    RV_Left = 0:    RV_Top = 0
    RV_TorKb = 3
    
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
'*      出 荷 日    ( iH1_SYmd )        *
'*--------------------------------------*
Private Sub iH1_SYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_SYmd_LostFocus()
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
        Call TorNmGet(RTrim$(iH3_Torcd), 3)
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
               iH3_ENYmd = iH1_SYmd
               iH3_ENBin = iH1_SBin
               If Trim(iH3_ENYmd) = "" Then iH3_ENYmd = Format(Now(), "YYYYMMDD")
        End Select
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If

    If ProcHB = "H3" Then
        If Not Head3Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        vsGrid1.SetFocus
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
    Dim H3Mode          As Boolean
    Dim H3Color         As Long
    Dim H4Mode          As Boolean
    Dim H4Color         As Long
    Dim H5Mode          As Boolean
    Dim H5Color         As Long
    Dim B1Mode          As Boolean
    Dim B1Color         As Long
    
    H1Mode = False:     H1Color = gPLostSel
    H3Mode = False:     H3Color = gPLostSel
    H4Mode = False:     H4Color = gPLostSel
    H5Mode = False:     H5Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "H3":      H3Mode = True:  H3Color = gPGotSel
        Case "H4":      H4Mode = True:  H4Color = gPGotSel
        Case "H5":      H5Mode = True:  H5Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    H3_Area1.Enabled = H3Mode
    H3_Area1.BackColor = H3Color
    H4_Area1.Enabled = H4Mode
    H4_Area1.BackColor = H4Color
    H5_Area1.Enabled = H5Mode
    H5_Area1.BackColor = H5Color
    
    B1_Area1.Enabled = B1Mode Or H3Mode Or H4Mode Or H5Mode
    B1_Area1.BackColor = B1Color
    
    PB_CAN.Visible = Not H1Mode
    PB_Look.Visible = H1Mode
    PB_Bun.Visible = H1Mode
    
    PB_ENT.Visible = Not B1Mode
    
    PB_UPD.Visible = B1Mode
    If H3Mode Or H4Mode Or H5Mode Then
        If mBunCnt <> 0 Or mKanCnt <> 0 Then
            PB_UPD.Visible = True
        End If
    End If
    
    If ProcHB = "H1" Then
       B1lb_Kensu = ""
       B1lb_KKensu = ""
       B1lb_BKensu = ""
       B1lb_MKensu = ""
       PB_Look.Visible = False
    End If
    
    If IRN_NKRead = 0 And IRN_NKReadBar = 0 Then
       PB_CHG.Visible = False
    Else
       PB_CHG.Visible = H1Mode Or H3Mode Or H4Mode Or H5Mode
    End If

    If ProcHB = "H4" Then
       Call QT10QSR_Open
       QRMsg.BackColor = &HC0FFC0
       QRMsg.Caption = "QRﾀｯﾁﾘｰﾀﾞ読取中"
    Else
       Call QT10QSR_Close
       QRMsg.BackColor = &HFFC0C0
       QRMsg.Caption = "QRﾀｯﾁﾘｰﾀﾞ読取"
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
            If .Name Like "i" & wbc_Enm & "_*" Then .text = "": .Enabled = True
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
    
'( 出荷日 )
    If RTrim$(iH1_SYmd) = "" And CisFun.Val2(iH1_SBin) <> 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    出荷日を入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_SYmd.SetFocus
        Exit Function
    Else
        If RTrim$(iH1_SYmd) <> "" And Not iH1_SYmd.cChkResult Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    正しい日付を入力して下さい         "
                 .MB_Button = Error
                 .MBOX
             End With
             iH1_SYmd.SetFocus
             Exit Function
         End If
    End If
    
'( 出荷便 )
    If RTrim$(iH1_SYmd) <> "" And CisFun.Val2(iH1_SBin) = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    出荷便を入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_SBin.SetFocus
        Exit Function
    End If
    
    If RTrim$(iH1_SYmd) <> "" Then
'----------------------------------------------------------------2006/11/28 Update Start
        'gInt = Kdou_Check(iH1_SYmd)
        gInt = Kdou_Check(iH1_SYmd, 0)
'----------------------------------------------------------------2006/11/28 End
        If gInt = 1 Then
        '休日
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(1) = "     出荷日が非稼働日です。     "
            CisFun.MB_MSG(3) = "     　よろしいですか？      "
            CisFun.MB_Button = OK_CAN
            If Not CisFun.MBOX Then
                iH1_SYmd.SetFocus
                Exit Function
            End If
        End If
        If gInt = 2 Then
        'カレンダ未登録
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(1) = "     " & RTrim(iH1_SYmd.cYear) & "年" & RTrim(iH1_SYmd.cMonth) & "月のカレンダ未登録！     "
            CisFun.MB_MSG(3) = "     　　(　カレンダマスタ　)　"
            CisFun.MB_Button = Error
            CisFun.MBOX
            iH1_SYmd.SetFocus
            Exit Function
        End If
    End If
    
    ' 表示件数カウント初期化
    mKanCnt = 0
    mBunCnt = 0
    mMiCnt = 0
    
    Head1Chk = True
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

    Call SqlSelectSet
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
         .Where_Create AlphaMe, "jyt.納入先", iH3_Torcd
         ' 未入荷のみ
         .Where_Create Numeric, "出荷実績区分", 0, 等しい, Create

         ' Toppsﾃﾞｰﾀ非対象
         .Where_Create AlphaMe, "jyt.作成種別", "T", 等しくない                 '2009/05/18
            

         gSL_Select = gSL_Select & .Where_Phrase
    End With

    gSL_Select = gSL_Select & mSL_Group & mSL_Order

    PNL_Wait.Visible = True
    Message = "【抽出中】しばらくお待ち下さい"
    DoEvents

    If Not JYTRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "    該当データがありません            "
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
    Call SqlSelectSet
    
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where jyt.出荷指示書NO = " & CisFun.Val2(iH5_Bar)
    gSL_Select = gSL_Select & "   And jyt.出荷実績区分 = 0"
    gSL_Select = gSL_Select & mSL_Group
    
    If Not JYTRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "    該当データがないか出荷済みです            "
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If

    Call GridSetKobetsu
    
    Head5Chk = True
End Function
'+---------------------------
'+      SQL文ｾｯﾄ処理
'+---------------------------
Private Sub SqlSelectSet()

    mSL_Select = ""
    mSL_Select = mSL_Select & "select"
    mSL_Select = mSL_Select & "  jyt.出荷指示書NO"
    mSL_Select = mSL_Select & ", jyt.納入日"
    mSL_Select = mSL_Select & ", jyt.納入便"
    mSL_Select = mSL_Select & ", jyt.納入先"
    mSL_Select = mSL_Select & ", isnull(trm.略称, '') 納入先名"
    mSL_Select = mSL_Select & ", jyt.受入"
    mSL_Select = mSL_Select & ", isnull(sym.値名称, '') 受注区分名"
'    mSL_Select = mSL_Select & ", ISNULL(TR2.略称,'') 出荷管理部署名"
    mSL_Select = mSL_Select & ", '' 出荷管理部署名"
    mSL_Select = mSL_Select & " from"
    mSL_Select = mSL_Select & "  受注テーブル jyt"
    mSL_Select = mSL_Select & "  left outer join 取引先マスタ trm"
    mSL_Select = mSL_Select & "    on trm.取引先CD = jyt.納入先"
    mSL_Select = mSL_Select & "   and trm.取引先区分 = 0"
    mSL_Select = mSL_Select & "  left outer join 名称マスタ sym"
    mSL_Select = mSL_Select & "    on sym.区分名称 = '受注区分'"
    mSL_Select = mSL_Select & "   and sym.値 = jyt.受注区分"
    mSL_Select = mSL_Select & "   AND sym.区分タイプ = ''"
    mSL_Select = mSL_Select & "  Left Outer Join 品番取引先マスタ HT"
    mSL_Select = mSL_Select & "    ON  HT.品番 = jyt.品番"
    mSL_Select = mSL_Select & "    And HT.取引先 = jyt.納入先"
    mSL_Select = mSL_Select & "    And HT.受入 = jyt.受入"
    mSL_Select = mSL_Select & "    And HT.品目 = 0"
'    mSL_Select = mSL_Select & "  Left Outer Join 取引先マスタ TR2"
'    mSL_Select = mSL_Select & "    ON TR2.取引先区分 = 1"
'    mSL_Select = mSL_Select & "    And TR2.取引先CD = HT.出荷管理部署"
'    mSL_Select = mSL_Select & "    And TR2.社内区分 = 1"

    mSL_Group = ""
    mSL_Group = mSL_Group & " group by"
    mSL_Group = mSL_Group & "  jyt.出荷指示書NO"
    mSL_Group = mSL_Group & ", jyt.納入日"
    mSL_Group = mSL_Group & ", jyt.納入便"
    mSL_Group = mSL_Group & ", jyt.納入先"
    mSL_Group = mSL_Group & ", trm.略称"
    mSL_Group = mSL_Group & ", jyt.受入"
    mSL_Group = mSL_Group & ", sym.値名称"
'    mSL_Group = mSL_Group & ", TR2.略称"

    mSL_Order = ""
    mSL_Order = mSL_Order & " order by"
    mSL_Order = mSL_Order & "  jyt.納入日"
    mSL_Order = mSL_Order & ", jyt.納入便"
    mSL_Order = mSL_Order & ", jyt.納入先"
    mSL_Order = mSL_Order & ", jyt.受入"
    mSL_Order = mSL_Order & ", jyt.出荷指示書NO"

End Sub
'
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆
'☆                         出　　荷　　実　　績　　更　　新　　(  F1  )
'☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
Private Sub PB_UPD_Click()

'   >> 完納・分納の選択ﾁｪｯｸ
    gCnt1 = 0
    With vsGrid1
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
        .MB_MSG(2) = "      出荷実績更新を行います。               "
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
'    Call CisDB.DBTran(TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始   2009/12/15

    With vsGrid1
        .Redraw = False
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, 1)) = "完納" Then
                wSJNo = CisFun.Val2(vsGrid1.TextMatrix(gLong, 6))
            
            '   >> 出荷テーブル出力
                If Not DBInsert_SKT Then GoTo PB_UPD_Er
            '   >> 受注テーブル更新
                If Not DBUpdate_JYT Then GoTo PB_UPD_Er

            '   >> 実績管理テーブル(出庫ﾃﾞｰﾀ)出力
                If Not SyukoInsert Then GoTo PB_UPD_Er
                
'*--------------------------------------------------<< 2004/06/28 Insert Start >>
            '   >> 初回実行の判定
                If gCHT_DelKbn53 = 0 Then
                    gCHT_DelKbn53 = 1
                '   >> 帳票出力テーブル削除
                    If Not DBDelete_CHT Then GoTo PB_UPD_Er
                End If
            '   >> 帳票出力テーブル出力
                If Not DBInsert_CHT(wSJNo) Then GoTo PB_UPD_Er
'*--------------------------------------------------<< 2004/06/28 Insert End >>
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
'    Call CisDB.DBTran(TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ終了         2009/12/15
    ' 分納処理
    With vsGrid1
        .Redraw = False
        Erase Rv_SJNo
        Erase Rv_SExec
        Erase Rv_SRow
        
        Rv_STtlSu = 0
        
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, 1)) = "分納" Then
               If CisFun.Val2(vsGrid1.TextMatrix(gLong, 6)) <> 0 Then
                  
                  Rv_STtlSu = Rv_STtlSu + 1
                  ReDim Preserve Rv_SJNo(Rv_STtlSu)
                  ReDim Preserve Rv_SExec(Rv_STtlSu)
                  ReDim Preserve Rv_SRow(Rv_STtlSu)
                  
                  Rv_SJNo(Rv_STtlSu) = CisFun.Val2(vsGrid1.TextMatrix(gLong, 6))
                  Rv_SRow(Rv_STtlSu) = gLong
               
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
    
    If Rv_STtlSu <> 0 Then
        Rv_NCall = "K"
        Rv_NYmd = iH1_SYmd
        Rv_NBin = iH1_SBin
        
        CXJ0220.Show vbModal
        Unload CXJ0220
        Set CXJ0220 = Nothing
        Rv_NCall = ""
        
        With vsGrid1
            .Redraw = False
            For gLong = 1 To UBound(Rv_SRow)
                .Row = Rv_SRow(gLong)
                If Rv_SExec(gLong) Then
                   TTL_BSKensu = Format(CisFun.Val2(TTL_BSKensu) + 1, "#,###")
                   .TextMatrix(Rv_SRow(gLong), 1) = "分済"
                Else
                   .TextMatrix(Rv_SRow(gLong), 1) = ""
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
    
    Set QT10QSR_Comm = PDQComm1
    Call QT10QSRInfoGet

'    Timer1.Enabled = True
    Exit Sub
PB_UPD_Er:
    vsGrid1.Redraw = True
    PNL_Wait.Visible = False
    Call CisDB.DBTran(TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ破棄

End Sub
'+--------------------------------
'+      登録処理：出荷テーブル
'+--------------------------------
Private Function DBInsert_SKT(Optional DBNo As Integer = 0) As Boolean
    DBInsert_SKT = False
    gSL_Select = ""
    gSL_Select = gSL_Select & "insert into 出荷テーブル"
    gSL_Select = gSL_Select & " ("
    gSL_Select = gSL_Select & "  出荷指示書NO, 出荷指示書行NO"
    gSL_Select = gSL_Select & ", 出荷日, 出荷便"
    gSL_Select = gSL_Select & ", 出荷枚数, 出荷端数"
    gSL_Select = gSL_Select & ", 出荷数, 分納数"
    gSL_Select = gSL_Select & ", 照合年月, 照合処理日, 照合結果"
    gSL_Select = gSL_Select & ", 実績年月, 実績処理日, 実績結果"
    gSL_Select = gSL_Select & ", 作成日, 作成端末, 作成者"
    gSL_Select = gSL_Select & ") "
    gSL_Select = gSL_Select & "select 出荷指示書NO, 出荷指示書行NO"
    If Trim(iH1_SYmd) <> "" Then
        gSL_Select = gSL_Select & ", '" & iH1_SYmd & "'"
        gSL_Select = gSL_Select & ", '" & iH1_SBin & "'"
    Else
        gSL_Select = gSL_Select & ", 納入日, 納入便"
    End If
    gSL_Select = gSL_Select & ", 枚数, 端数, 納入数, 0 分納数"
    gSL_Select = gSL_Select & ", '' 照合年月, '' 照合処理日, '' 照合結果"
    gSL_Select = gSL_Select & ", '' 実績年月, '' 実績処理日, '' 実績結果"
    gSL_Select = gSL_Select & ", getdate(), substring(host_name(),1,20), '" & gTanto & "'"
    gSL_Select = gSL_Select & "  from 受注テーブル"
    gSL_Select = gSL_Select & " where 出荷指示書NO = " & wSJNo
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBInsert_SKT = True
End Function
'=================================================================== 2009/12/15 Start
'+-------------------------------------------------
'+      登録処理：実績管理テーブル---出庫ﾃﾞｰﾀ作成
'+-------------------------------------------------
Private Function SyukoInsert(Optional DBNo As Integer = 0) As Boolean
    SyukoInsert = False
    
    gSL_Select = "SELECT * FROM 受注テーブル "
    gSL_Select = gSL_Select & " WHERE 出荷指示書NO = " & wSJNo
    If JYTRead(gSL_Select) Then
        Do Until Not JYT_RDSTS
           '構成品番情報の取得
            With CisDB
                .SQL = "構成問合せ情報抽出"
                .StoadoCount = 4
                .ParaNo = 0: .ParaIO = Return用
                .ParaNo = 1: .ParaIO = Input用
                .ParaNo = 2: .ParaIO = Input用
                .ParaNo = 3: .ParaIO = Input用
                .ParaNo = 4: .ParaIO = Input用
            
                .ParaNo = 1: .ParaValue = RTrim(JYT.品番)
                .ParaNo = 2: .ParaValue = 0
                .ParaNo = 3: .ParaValue = 0
                .ParaNo = 4: .ParaValue = 0
            
                If Not .DBStored Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(2) = " [出庫作成]構成マスタ情報の抽出に失敗しました。 "
                        .MB_Title = "抽出処理"
                        .MB_Button = Error
                        If .MBOX Then Exit Function
                    End With
                End If
            End With
            
            '構成情報読み込み
            gSL_Select = "SELECT * FROM 構成問合せワーク "
            gSL_Select = gSL_Select & " WHERE 処理端末 = Host_name() "
'            gSL_Select = gSL_Select & " AND   レベル <> 0 "
            gSL_Select = gSL_Select & " ORDER BY 順位 "
            If KOLRead(gSL_Select, , 1) Then
                Do Until Not KOL_RDSTS
                   '品番工順ﾏｽﾀ読み込み･･･棚卸ﾌﾟﾚｼｰﾄ発行[有]のみ対象
                   gSL_Select = "SELECT * FROM 品番工順マスタ"
                   gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(KOL.子品番) & "'"
                   'gSL_Select = gSL_Select & " AND 棚卸プレシート発行 = 0"                             '2020/01/10 comment
                   gSL_Select = gSL_Select & " ORDER BY 品番,取引先,受入,工順 "
                   If HKTRead(gSL_Select, , 2) Then
                       Do Until Not HKT_RDSTS
                           Call ItemsClearJKT
                           
                           With CisDB
                               .SQL = "番号獲得"
                               .StoadoCount = 3
                               .ParaNo = 0: .ParaIO = Return用
                               .ParaNo = 1: .ParaIO = Input用
                               .ParaNo = 2: .ParaIO = OutPut用
                               .ParaNo = 3: .ParaIO = Return用
                           '
                               .ParaNo = 1: .ParaValue = 6
                           '
                               If Not .DBStored Then
                                   CisFun.MB_Lines = 4
                                   CisFun.MB_MSG(2) = "    実績管理NOの獲得に失敗しました            "
                                   CisFun.MB_Title = "採番処理"
                                   CisFun.MB_Button = Error
                                   CisFun.MBOX
                                   Exit Function
                               Else
                                   .ParaNo = 2: JKT.実績管理NO = .ParaValue
                               End If
                           End With
                           JKT.実績区分 = 4
                           JKT.納入番号 = JYT.納入番号
                           JKT.納品書頁 = JYT.納品書頁
                           JKT.納品書行 = JYT.納品書行
                           JKT.分納区分 = 0
                           JKT.実績先 = HKT.取引先
                           JKT.受入 = HKT.受入
                           JKT.品番 = HKT.品番
                           If Trim(iH1_SYmd) <> "" Then
                               JKT.実績日 = iH1_SYmd
                               JKT.実績便 = iH1_SBin
                           Else
                               JKT.実績日 = JYT.納入日
                               JKT.実績便 = JYT.納入便
                           End If
                           If HKT.収容数 <> 0 Then
                              JKT.収容数 = HKT.収容数
                           Else
                              JKT.収容数 = JYT.収容数
                           End If
                                                       
'                           wInt = JYT.納入数 * KOL.構成数              '2010/12/28
                           wLong = JYT.納入数 * KOL.構成数              '2010/12/28
                           

'                           JKT.実績枚数 = CisFun.HasuProc(JYT.納入数 / JKT.収容数, 切上げ)
'                           JKT.実績端数 = JYT.納入数 Mod JKT.収容数
'                           JKT.実績数 = JYT.納入数
'                           JKT.実績枚数 = CisFun.HasuProc(wInt / JKT.収容数, 切上げ)       '2010/12/28
                                                       

'==================================== 2011/10/17 Start
'                           JKT.実績枚数 = CisFun.HasuProc(wLong / JKT.収容数, 切上げ)       '2010/12/28
                           If JKT.収容数 <> 0 Then
                               JKT.実績枚数 = CisFun.HasuProc(wLong / JKT.収容数, 切上げ)
                           Else
                               JKT.実績枚数 = 0
                           End If
'==================================== 2011/10/17 End


'                           JKT.実績端数 = wInt Mod JKT.収容数                              '2011/02/04
'                           JKT.実績数 = wInt                                               '2011/02/04

'==================================== 2011/10/17 Start
'                           JKT.実績端数 = wLong Mod JKT.収容数
'                           JKT.実績数 = wLong
                           If JKT.収容数 <> 0 Then
                               JKT.実績端数 = wLong Mod JKT.収容数
                               JKT.実績数 = wLong
                           Else
                               JKT.実績端数 = 0
                               JKT.実績数 = wLong
                           End If
'==================================== 2011/10/17 End
                           
                           JKT.有無償区分 = 1
                           JKT.入力種類 = "S"
                           JKT.工順 = HKT.工順
                           JKT.出荷指示書NO = JYT.出荷指示書NO
                           JKT.出荷指示書行NO = JYT.出荷指示書行NO
                           JKT.作成者 = "CXJ0210"
                           
                           If KOL.レベル = 0 Then
                              JKT.親品番 = 1
                           Else
                              JKT.親品番 = 0
                           End If
                           
                           Call JKTInsert
                                                       
                           Call HKTReadNext(2)
                      Loop
                      Call HKTClose(2)
                   End If
                   
                   Call KOLReadNext(1)
                Loop
                Call KOLClose(1)
            End If
            
           Call JYTReadNext
       Loop
       
       Call JYTClose
    End If
    
    SyukoInsert = True
End Function
'=================================================================== 2009/12/15 End
'+--------------------------------
'+      更新処理：受注テーブル
'+--------------------------------
Private Function DBUpdate_JYT(Optional DBNo As Integer = 0) As Boolean
    DBUpdate_JYT = False
    gSL_Select = ""
    gSL_Select = gSL_Select & "update 受注テーブル"
    gSL_Select = gSL_Select & "   set 出荷実績区分 = 1"
    gSL_Select = gSL_Select & " where 出荷指示書NO = " & wSJNo
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBUpdate_JYT = True
End Function
'*--------------------------------------------------<< 2004/06/28 Insert Start >>
'+--------------------------------
'+      削除処理：帳票出力テーブル
'+--------------------------------
Private Function DBDelete_CHT(Optional DBNo As Integer = 0) As Boolean
    DBDelete_CHT = False
    gSL_Select = ""
    gSL_Select = gSL_Select & "delete"
    gSL_Select = gSL_Select & "  from 帳票出力テーブル"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and 帳票区分 in (53)"
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBDelete_CHT = True
End Function
'+--------------------------------
'+      登録処理：帳票出力テーブル
'+--------------------------------
Private Function DBInsert_CHT(SJNo As Long, Optional DBNo As Integer = 0) As Boolean
    DBInsert_CHT = False
    gSL_Select = ""
    gSL_Select = gSL_Select & "insert into 帳票出力テーブル"
    gSL_Select = gSL_Select & " (端末番号, 帳票区分,管理文字,検索情報,管理NO,管理行NO,補助番号,出力先,作成日) "
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  substring(host_name(),1,20)"
    gSL_Select = gSL_Select & ", 53"
    gSL_Select = gSL_Select & ", jyt.納入番号"
    gSL_Select = gSL_Select & ", '6'"
    gSL_Select = gSL_Select & ", jyt.納品書管理NO"
    gSL_Select = gSL_Select & ", 0"
    gSL_Select = gSL_Select & ", jyt.納品書頁"
    gSL_Select = gSL_Select & ", 0"
    gSL_Select = gSL_Select & ", getdate()"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  受注テーブル jyt"
    gSL_Select = gSL_Select & ", 品番取引先マスタ htm"
    gSL_Select = gSL_Select & ", 出荷テーブル skt"
    gSL_Select = gSL_Select & " where jyt.出荷指示書NO = " & SJNo
    gSL_Select = gSL_Select & "   and htm.品番 = jyt.品番"
    gSL_Select = gSL_Select & "   and htm.取引先 = jyt.納入先"
    gSL_Select = gSL_Select & "   and htm.受入 = jyt.受入"
    gSL_Select = gSL_Select & "   and htm.品目 = 0"
    gSL_Select = gSL_Select & "   and htm.受領書発行 = 1"
    gSL_Select = gSL_Select & "   and skt.出荷指示書NO = jyt.出荷指示書NO"
    gSL_Select = gSL_Select & "   and skt.出荷指示書行NO = jyt.出荷指示書行NO"
    gSL_Select = gSL_Select & "   and not exists"
    gSL_Select = gSL_Select & "("
    gSL_Select = gSL_Select & "select 'X'"
    gSL_Select = gSL_Select & "  from 帳票出力テーブル cht"
    gSL_Select = gSL_Select & " where cht.端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and cht.帳票区分 = 53"
    gSL_Select = gSL_Select & "   and cht.管理文字 = jyt.納入番号"
    gSL_Select = gSL_Select & "   and cht.検索情報 = '6'"
    gSL_Select = gSL_Select & "   and cht.管理NO = jyt.納品書管理NO"
    gSL_Select = gSL_Select & "   and cht.補助番号 = jyt.納品書頁"
    gSL_Select = gSL_Select & ")"
    gSL_Select = gSL_Select & " group by"
    gSL_Select = gSL_Select & "  jyt.納品書管理NO"
    gSL_Select = gSL_Select & ", jyt.納入番号"
    gSL_Select = gSL_Select & ", jyt.納品書頁"
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBInsert_CHT = True
End Function
'*--------------------------------------------------<< 2004/06/28 Insert End >>

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
        Call .FixedSet(中中, 右中, あり, 4, 0, "№")
        Call .FixedSet(中中, 中中, あり, 6, 1, "処理")
        Call .FixedSet(中中, 中中, あり, 14, 2, "納入日-便")
        Call .FixedSet(中中, 中中, あり, 14, 3, "出荷日-便")
        Call .FixedSet(中中, 左中, あり, 34, 4, "納入先")
        Call .FixedSet(中中, 中中, あり, 4, 5, "受入")
        Call .FixedSet(中中, 中中, あり, 14, 6, "出荷指示書№")
        Call .FixedSet(中中, 中中, あり, 9, 7, "受注区分")
        Call .FixedSet(中中, 左中, あり, 14, 8, "出荷管理部署")

'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 9, "Select")

'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8")

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
                Select Case Trim(.TextMatrix(gLong, 1))
                    Case ""
                        If gStr = "X" Then
                           Call CisVsGrid.vsColor(消去個別)
                        End If
                        Call CisVsGrid.vsColor(個別, B1lb_KKensu.BackColor)
                        mKanCnt = mKanCnt + 1
                        .TextMatrix(gLong, 1) = "完納"
                        mMiCnt = mMiCnt - 1
                    Case "完納"
                        If gStr = "X" Then
                           Call CisVsGrid.vsColor(消去個別)
                        End If
                        Call CisVsGrid.vsColor(個別, B1lb_BKensu.BackColor)
                        .TextMatrix(gLong, 1) = "分納"
                        mKanCnt = mKanCnt - 1
                        mBunCnt = mBunCnt + 1
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
    With vsGrid1
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not JYT_RDSTS
        With vsGrid1
            ID = ID + 1
       
            Call GridMove(ID)
        End With
        Call JYTReadNext
    Loop
    Call JYTClose
    
    Call CisVsGrid.ReNumber
    
    With vsGrid1
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
    
''''''''''    If JYT.入荷日 <> "" Then
''''''''''        With CisFun
''''''''''            .MB_Lines = 5
''''''''''            .MB_MSG(2) = "      入荷済みデータです              "
''''''''''            .MB_Button = Error
''''''''''            .MBOX
''''''''''        End With
''''''''''        Exit Function
''''''''''    End If
    With vsGrid1
        .Redraw = False
        
        For gCnt1 = 1 To .Rows - 1
            gCnt2 = CisFun.Val2(.TextMatrix(gCnt1, 6))
            If gCnt2 <> 0 And gCnt2 = JYT.出荷指示書NO Then
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

    With vsGrid1
        
        With CisFun
            .DateE_BefVal = JYT.納入日
            .DateEdit
            gStr = CisFun.DateE_AfVal
            If Trim(JYT.納入便) <> "" Then
                gStr = gStr & "-" & JYT.納入便
            Else
                gStr = gStr & "   "
            End If
        End With
        .TextMatrix(mRow, 2) = gStr
        
        If Trim(iH1_SYmd) = "" Then
            gStr = "============="
        Else
            With CisFun
                .DateE_BefVal = iH1_SYmd
                .DateEdit
                gStr = CisFun.DateE_AfVal
                If Trim(JYT.納入便) <> "" Then
                    gStr = gStr & "-" & Format(iH1_SBin, "00")
                Else
                    gStr = gStr & "   "
                End If
            End With
        End If
        .TextMatrix(mRow, 3) = gStr
        
        If Trim(JYT.納入先名) = "" Then
           gStr = JYT.納入先
        Else
           gStr = RTrim(JYT.納入先) & Space(7 - Len(JYT.納入先)) & ":"
           gStr = gStr & JYT.納入先名
        End If
        .TextMatrix(mRow, 4) = gStr
        .TextMatrix(mRow, 5) = JYT.受入
        .TextMatrix(mRow, 6) = Format(JYT.出荷指示書NO, String(10, "0"))
        .TextMatrix(mRow, 7) = RTrim$(JYT.受注区分名)
        
        .TextMatrix(mRow, 8) = RTrim$(JYT.出荷管理部署名)
        
        .TextMatrix(mRow, 1) = "完納"
        mKanCnt = mKanCnt + 1
        .Row = mRow
        Call CisVsGrid.vsColor(個別, B1lb_KKensu.BackColor)
    
    End With
End Sub

'******************************************************************
'*      初期処理
'******************************************************************
Private Function InitProc() As Boolean
    InitProc = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select *"
    gSL_Select = gSL_Select & "  from 受注入力ワーク"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and 累積区分 = 0"
    If JNWRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 6
            .MB_MSG(1) = "    当端末処理分の受注未累積データが存在します。            "
            .MB_MSG(4) = "      『出荷実績入力（一括）』処理を終了しますか？"
            .MB_MSG(5) = "      いいえの場合は、未累積データを削除し続行します"
            .MB_Title = "未累積データの確認"
            .MB_Button = Yes_No
            If .MBOX Then Exit Function
        End With
    End If

    With CisDB
        On Error GoTo InitProc_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
'       << 受注入力テーブル >>
        .SQL = ""
        .SQL = .SQL & "delete"
        .SQL = .SQL & "  from 受注入力ワーク"
        .SQL = .SQL & " where 端末番号 = substring(host_name(),1,20)"
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        On Error GoTo 0
    End With

    InitProc = True
    Exit Function
InitProc_Err:
    CisFun.ErrorBox
    End
End Function
'==========================================================='
'   終了処理（入力データがあれば後方処理を行う）
'==========================================================='
Private Sub EndProc()

'   << 入力件数確認 >>
    gSL_Select = ""
    gSL_Select = gSL_Select & "select count(*) 件数"
    gSL_Select = gSL_Select & "  from 受注入力ワーク"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and"
    gSL_Select = gSL_Select & " (     出荷指示書NO = 0"
    gSL_Select = gSL_Select & "    or 納入番号 = ''"
    gSL_Select = gSL_Select & "    or 累積区分 = 0"
    gSL_Select = gSL_Select & " )"
    If JNWRead(gSL_Select, 1) Then
        If JNW.件数 > 0 Then
        '   出荷指示NO採番
            CXJ0110.Show vbModal
            Unload CXJ0110
            Set CXJ0110 = Nothing
            DoEvents
        '   出荷指示書発行
            Call CXJ0120LPrint(0)
'*---------------------------------------------<< 2004/04/16 Insert Start >>
        '   自社伝票発行
            Call CXJ0140LPrint(0)
        '   自社伝票ﾌｧｲﾙ出力
            Call CXJ0150TFile(0)
'*---------------------------------------------<< 2004/04/16 Insert End >>
'*---------------------------------------------<< 2004/06/28 Insert Start >>
        Else
            If gCHT_DelKbn53 = 1 Then
            '   自社伝票発行
                Call CXJ0140LPrint(0)
            End If
'*---------------------------------------------<< 2004/06/28 Insert End >>
        End If
    End If

End Sub
'+------------------------------+
'+  ﾊﾞｰｺｰﾄﾞ読取処理             *
'+------------------------------+
Private Sub PDQComm1_OnComm()
    
    gStr = QT10QSR_Event
    If gStr = "ERR" Then Exit Sub
    If gStr <> "RCV" Then Exit Sub
    
    If Not QT10QSR_RCVProc Then Exit Sub    ' CR が存在しない時抜ける

'( データ確認 )
    If Trim(QT10QSR_ReadData) = "" Then Exit Sub
    
    Call SqlSelectSet
    
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where jyt.出荷指示書NO = " & CisFun.Val2(QT10QSR_ReadData)
    gSL_Select = gSL_Select & "   And jyt.出荷実績区分 = 0"
    gSL_Select = gSL_Select & mSL_Group
    
    QT10QSR_ReadData = ""
    
    If Not JYTRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "    該当データがないか出荷済みです。            "
            .MB_Button = Error
            .MBOX
        End With
        Exit Sub
    End If

    Call GridSetKobetsu
    
End Sub

