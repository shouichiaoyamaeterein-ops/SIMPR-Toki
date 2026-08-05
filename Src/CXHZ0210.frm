VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Object = "{490AF35D-478A-11D4-AEC2-00000EA57F9E}#4.0#0"; "CISMSG~2.OCX"
Begin VB.Form CXHZ0210 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "材料入荷実績入力"
   ClientHeight    =   11115
   ClientLeft      =   3990
   ClientTop       =   4335
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  '最大化
   Begin VB.PictureBox PNL_Start 
      BackColor       =   &H00000000&
      Height          =   825
      Left            =   390
      ScaleHeight     =   765
      ScaleWidth      =   8235
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   540
      Visible         =   0   'False
      Width           =   8295
      Begin CisMsgRelate_v60.CISMsgRelate ProcMsg 
         Height          =   645
         Left            =   75
         Top             =   60
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   1138
         BackColor       =   65535
         Caption         =   "ホッパリーダの【ｽﾀｰﾄ】キーを押して下さい"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   15.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         mrInterval      =   50
         mrStopCnt       =   40
      End
   End
   Begin Cis3D_v60.CIS3D H6_Area1 
      Height          =   1320
      Left            =   465
      Top             =   5415
      Visible         =   0   'False
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
      Begin VB.TextBox QRRDummy 
         Height          =   330
         Left            =   3645
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   30
         Width           =   240
      End
      Begin Cis3D_v60.CIS3D QRRMsg 
         Height          =   1170
         Left            =   90
         Top             =   75
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   2064
         BackColor       =   16777152
         ForeColor       =   16711680
         Caption         =   "QRﾎｯﾊﾟﾘｰﾀﾞ読取"
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
      Left            =   10245
      ScaleHeight     =   1215
      ScaleWidth      =   1905
      TabIndex        =   19
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
            TabIndex        =   10
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
      TabIndex        =   20
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
      Left            =   8730
      ScaleHeight     =   690
      ScaleWidth      =   6150
      TabIndex        =   16
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
         Left            =   3660
         TabIndex        =   11
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
      TabIndex        =   14
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
            TabIndex        =   8
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
            TabIndex        =   7
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
            TabIndex        =   9
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
      ScaleWidth      =   2625
      TabIndex        =   18
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
         Caption         =   "    入 荷 日  - 便"
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
         Begin CisText_V60.CisText iH1_NBin 
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
      Left            =   390
      ScaleHeight     =   7665
      ScaleWidth      =   14505
      TabIndex        =   15
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
            TabIndex        =   12
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
            FormatString    =   "№|処理 |納　期 　　-便|入荷日  　　-便|仕入先 |受 |納 番 |頁|発注区|手配区|実|品　番　　　　　　　|納品書管理NO|納品書頁|Select|"
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
      TabIndex        =   13
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
      BackColor       =   8438015
      ForeColor       =   0
      Caption         =   "【 材料・入荷実績入力 [一括)]  】"
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
         Width           =   2865
         _ExtentX        =   5054
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
         BackColor       =   8438015
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
         cBoderColor1    =   8438015
         cBoderColor2    =   8438015
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
         BackColor       =   8438015
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
      TabIndex        =   17
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
         Begin CisText_V60.CisText iH3_Torcd 
            Height          =   375
            Left            =   90
            TabIndex        =   6
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
   Begin PdqcommLib.PDQComm PDQComm2 
      Height          =   480
      Left            =   1785
      TabIndex        =   22
      Top             =   630
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
End
Attribute VB_Name = "CXHZ0210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   材料入荷実績入力
'**       フォームID    :   CXHZ0210
'**       処理概要      :
'**
'**       作  成  日    :   2004/02/23  By CIS
'**       変  更  日    :   2009/01/19  By CIS 手配区分：直送時　直送向け支給伝票発行
'**       変  更  日    :   2014/02/04  By CIS 項目名間違?
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
    QRRDummy.Left = -1000
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 600    '
    Timer2.Enabled = True    '

    H2_Area1.Move H3_Area1.Left, H3_Area1.Top
    H4_Area1.Move H3_Area1.Left, H3_Area1.Top
    H5_Area1.Move H3_Area1.Left, H3_Area1.Top
    H6_Area1.Move H3_Area1.Left, H3_Area1.Top

    Set QRr_Object = PDQComm1
    Set QRr_MsgObject = QRMsg
    Set QRr_ErrMsgObj = ErrMsg
    Set TX2400_StartMsg = PNL_Start
    Set TX2400_ReadObj = QRRMsg
    Set TX2400_ErObj = ErrMsg
    '   ( QR Reader 環境 )
    Call QRInit
    Call TX2400InfoGet
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
    
    K_Sykbnm.Tag = "3"
    Call PB_CHG_Click

'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    If Rv_NCall = "" Then
       Me.Show
    End If

    '================================ 2009/01/16 Start
    Call DBDelete_CHT(1)            '支給伝票データ削除
    '================================ 2009/01/16 End

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
    mSL_Select = mSL_Select & "  Count(HN.材質) 件数,"
    mSL_Select = mSL_Select & "  MAX(HN.材質) 材質,"
    mSL_Select = mSL_Select & "  MAX(HN.実数区分) 実数区分 "
    mSL_Select = mSL_Select & " From 材料発注テーブル HC"
    mSL_Select = mSL_Select & "  Left Outer Join 材料入荷テーブル NK"
    mSL_Select = mSL_Select & "     On NK.発注管理NO = HC.発注管理NO"
    mSL_Select = mSL_Select & "  Left Outer Join 材料マスタ HN"
    mSL_Select = mSL_Select & "     ON HN.材料管理番号 = HC.材料管理番号"
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

    ' 材料入荷実績処理ワーク削除
    With CisDB
        .SQL = "Delete 材料入荷実績処理ワーク"
        .SQL = .SQL & "    Where 処理区分 = 9"
        .DBExec
    End With
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
        gSL_Select = "Select Count(*) 件数 From 材料入荷実績処理ワーク"
        gSL_Select = gSL_Select & " Where 処理端末 = Host_Name()"
        gSL_Select = gSL_Select & "   And 処理区分 = 0"
        If ZJWRead(gSL_Select, 1) Then
          If ZJW.件数 <> 0 Then
          ' 発注処理
            On Error Resume Next
            CisDB.PassLink
            With CisFun
                .ShellApl = gIniExe & "PXHZ0110.exe"
                .ShellAplPara = "J"
                .ShellTask
            End With
            On Error GoTo 0
          
            '================================================ 2009/01/19 Start
            Call Prt_Nou                            '支給伝票発行
            '================================================ 2009/01/19 End
          
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
    
    If ProcHB = "H2" Or ProcHB = "H3" Or ProcHB = "H4" Or ProcHB = "H5" Or ProcHB = "H6" Then
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
    Call TX2400Close
    
    CXHZ0220.Show vbModal
    Unload CXHZ0220
    Set CXHZ0220 = Nothing
    
    Timer1.Enabled = True
    Rv_NCall = ""
    
    '   ( QR Reader 環境 )
    Set QRr_Object = PDQComm1
    Set QRr_MsgObject = QRMsg
    Set QRr_ErrMsgObj = ErrMsg
    Set TX2400_StartMsg = PNL_Start
    Set TX2400_ReadObj = QRRMsg
    Set TX2400_ErObj = ErrMsg
    Call QRInit
    Call TX2400InfoGet
    Call DispChange(ProcHB)
    
End Sub
'+------------------+
'+    切替   (F6)   +
'+------------------+
Private Sub PB_CHG_Click()
    gLong = CisFun.Val2(K_Sykbnm.Tag)
    
    gLong = gLong + 1
    If gLong > 5 Then gLong = 1
    
    gStr = ""
    gStr = "H3"
    If IRN_NKRead = 1 Then
       gStr = gStr & "H4"
    Else
       gStr = gStr & "  "
    End If
    If IRN_NKReadHP = 1 Then
       gStr = gStr & "H6"
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
    For gInt = 1 To 5
        mHeadProc = Mid(gStr, gLong + (gLong - 1), 2)
        If Trim(mHeadProc) <> "" Then
           K_Sykbnm.Tag = gLong
           Exit For
        End If
        gLong = gLong + 1
        If gLong > 5 Then gLong = 1
    Next gInt
    
    H2_Area1.Visible = False
    H3_Area1.Visible = False
    H4_Area1.Visible = False
    H5_Area1.Visible = False
    H6_Area1.Visible = False
    
    Select Case mHeadProc
        Case "H3"
            K_Sykbnm.Caption = "【 一覧 】"
            H3_Area1.Visible = True
        Case "H4"
            K_Sykbnm.Caption = "【 QRﾀｯﾁﾘｰﾀﾞ 】"
            H4_Area1.Visible = True
        Case "H5"
            K_Sykbnm.Caption = "【 ﾊﾞｰｺｰﾄﾞ 】"
            H5_Area1.Visible = True
        Case "H6"
            K_Sykbnm.Caption = "【QRﾎｯﾊﾟﾘｰﾀﾞ】"
            H6_Area1.Visible = True
        Case Else
            K_Sykbnm.Caption = "【 納番 】"
            H2_Area1.Visible = True
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
    Dim H6Mode          As Boolean
    Dim H6Color         As Long
    Dim B1Mode          As Boolean
    Dim B1Color         As Long
    
    H1Mode = False:     H1Color = gPLostSel
    H2Mode = False:     H2Color = gPLostSel
    H3Mode = False:     H3Color = gPLostSel
    H4Mode = False:     H4Color = gPLostSel
    H5Mode = False:     H5Color = gPLostSel
    H6Mode = False:     H6Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "H2":      H2Mode = True:  H2Color = gPGotSel
        Case "H3":      H3Mode = True:  H3Color = gPGotSel
        Case "H4":      H4Mode = True:  H4Color = gPGotSel
        Case "H5":      H5Mode = True:  H5Color = gPGotSel
        Case "H6":      H6Mode = True:  H6Color = gPGotSel
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
    H6_Area1.Enabled = H6Mode
    H6_Area1.BackColor = H6Color
    
    B1_Area1.Enabled = B1Mode Or H2Mode Or H3Mode Or H4Mode Or H5Mode Or H6Mode
    B1_Area1.BackColor = B1Color
    
    PB_CAN.Visible = Not H1Mode
    PB_CHG.Visible = Not B1Mode
    PB_Look.Visible = H1Mode
    PB_Bun.Visible = H1Mode
    
    PB_UPD.Visible = B1Mode
    If H2Mode Or H4Mode Or H5Mode Or H6Mode Then
        If mBunCnt <> 0 Or mKanCnt <> 0 Then
            PB_UPD.Visible = True
        End If
    End If
    
    PB_Ent.Visible = Not (H4Mode Or H6Mode Or B1Mode)
    
    If ProcHB = "H1" Then
       B1lb_Kensu = ""
       B1lb_KKensu = ""
       B1lb_BKensu = ""
       B1lb_MKensu = ""
       PB_Look.Visible = False
    End If
'
    Set QRr_Object = PDQComm1
    Set TX2400_Comm = PDQComm2
    Call QR_Comm_Close
    Call TX2400Close
    If ProcHB = "H4" Then
       Call QR_Comm_Open
    End If
    If ProcHB = "H6" Then
       If TX2400_ReadAuto Then
          ProcMsg.Caption = "用紙をホッパーにセットして下さい。"
       Else
          ProcMsg.Caption = "用紙をﾎｯﾊﾟｰにｾｯﾄし【ｽﾀｰﾄ】キーを押して下さい。"
       End If
       PNL_Start.Tag = ""
       Call TX2400OpenAuto
    End If
    PNL_Start.Visible = H6Mode
    Call TX2400QRRReadMsg("")

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
    If RTrim$(iH1_NYmd) <> "" Then
        gInt = Kdou_Check(iH1_NYmd)
        If gInt = 1 Then
        '休日
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(1) = "     入荷日が非稼働日です。     "
            CisFun.MB_MSG(3) = "     　よろしいですか？      "
            CisFun.MB_Button = OK_CAN
            If Not CisFun.MBOX Then
                iH1_NYmd.SetFocus
                Exit Function
            End If
        End If
        If gInt = 2 Then
        'カレンダ未登録
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(1) = "     " & RTrim(iH1_NYmd.cYear) & "年" & RTrim(iH1_NYmd.cMonth) & "月のカレンダ未登録！     "
            CisFun.MB_MSG(3) = "     　　(　カレンダマスタ　)　"
            CisFun.MB_Button = Error
            CisFun.MBOX
            iH1_NYmd.SetFocus
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
    
    If Not ZCTRead(gSL_Select) Then
        Call ZCTClose
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  該当データがありません  "
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
    Call ZCTClose
    If CisDB.RecordCount > 1 Then
        gSL_Select = "Select HC.*,HN.表示寸法 From 材料発注テーブル HC "
        gSL_Select = gSL_Select & " Left Outer Join 材料マスタ HN "
        gSL_Select = gSL_Select & "   ON HN.材料管理番号 = HC.材料管理番号 "
        gSL_Select = gSL_Select & " Left Outer Join 材料入荷テーブル NY "
        gSL_Select = gSL_Select & "   On NY.発注管理NO = HC.発注管理NO "
        gSL_Select = gSL_Select & " Where HC.納入番号 = '" & RTrim$(iH2_Nouno) & "'"
        gSL_Select = gSL_Select & " And   HC.納品書頁 =  " & CisFun.Val2(iH2_NPage)
        gSL_Select = gSL_Select & " And   HC.納入日   = '" & RTrim$(iH2_NYmd) & "'"
        gSL_Select = gSL_Select & " And   Isnull(NY.入荷日,'') = ''"
        gSL_Select = gSL_Select & " Order By HC.材質,HN.表示寸法"
        Call ZCTRead(gSL_Select)
        If CisDB.RecordCount = 1 Then
           iH2_Nouno.Tag = ZCT.納品書管理NO
           Call ZCTClose
        Else
            CXHZ0215.Show vbModal
            Unload CXHZ0215
            Set CXHZ0215 = Nothing
            If RV_Rtn Then
               iH2_Nouno.Tag = Rvc_NKNo
            Else
               With CisFun
                   .MB_Lines = 5
                   .MB_MSG(2) = "  伝票の選択がされませんでした  "
                   .MB_Button = Error
                   .MBOX
               End With
               Exit Function
            End If
        End If
        
        gSL_Select = mSL_Select
        
        gSL_Select = gSL_Select & " Where HC.納品書管理NO = " & CisFun.Val2(iH2_Nouno.Tag) & vbCrLf
        gSL_Select = gSL_Select & " And   HC.納入番号 = '" & RTrim$(iH2_Nouno) & "'" & vbCrLf
        gSL_Select = gSL_Select & " And   HC.納品書頁 =  " & CisFun.Val2(iH2_NPage) & vbCrLf
        gSL_Select = gSL_Select & " And   HC.納入日   = '" & RTrim$(iH2_NYmd) & "'" & vbCrLf
        gSL_Select = gSL_Select & mSL_Group
        If Not ZCTRead(gSL_Select, 1) Then
           With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "  該当データがありません  "
                .MB_Button = Error
                .MBOX
            End With
            Exit Function
       End If
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
    
    If Not ZCTRead(gSL_Select) Then
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
    
    If Not ZCTRead(gSL_Select, 1) Then
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
    Set QRr_Object = PDQComm1
    Set TX2400_Comm = PDQComm2
    Call QR_Comm_Close
    Call TX2400Close
    
    Call CisDB.DBTran(TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    With VsGrid1
        .Redraw = False
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, 1)) = "完納" Then
               With CisDB
                    .SQL = "Insert Into 材料入荷テーブル  ( " & vbCrLf
                    .SQL = .SQL & "発注管理NO,納入番号,納品書頁," & vbCrLf
                    .SQL = .SQL & "納品書行,分納区分,仕入先,受入," & vbCrLf
                    .SQL = .SQL & "材質,板厚,幅,長さ,材料管理番号,背番号,入荷日," & vbCrLf
                    .SQL = .SQL & "入荷便,収容数,入荷枚数," & vbCrLf
                    .SQL = .SQL & "入荷端数,入荷数,分割数," & vbCrLf
                    .SQL = .SQL & "挽回日,挽回便,入荷備考,分割納品書NO," & vbCrLf
                    .SQL = .SQL & "検収区分,検収年月,検収処理日," & vbCrLf
                    .SQL = .SQL & "作成日,作成端末,作成者 )" & vbCrLf
                    
                    .SQL = .SQL & "Select 発注管理NO,納入番号,納品書頁," & vbCrLf
                    .SQL = .SQL & "納品書行,0,仕入先,受入," & vbCrLf
                    .SQL = .SQL & "材質,板厚,幅,長さ,材料管理番号,背番号," & vbCrLf
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
                    .SQL = .SQL & "'','','','','','',''," & vbCrLf
                    .SQL = .SQL & "Getdate(),Host_Name(),'" & gTanto & "'" & vbCrLf
                    .SQL = .SQL & " From 材料発注テーブル "
                    .SQL = .SQL & " Where 納品書管理NO = " & CisFun.Val2(VsGrid1.TextMatrix(gLong, 12))
                    .SQL = .SQL & "   And 納品書頁 = " & CisFun.Val2(VsGrid1.TextMatrix(gLong, 13))
                    
                    If Not .DBExec Then GoTo PB_UPD_Er
                    
               
                    .SQL = "Insert Into 材料入荷実績処理ワーク  ( "
                    .SQL = .SQL & "処理端末,発注管理NO,処理区分,作成日 )"
                    .SQL = .SQL & " Select Host_Name(),発注管理NO,0,Getdate() "
                    .SQL = .SQL & " From 材料発注テーブル "
                    .SQL = .SQL & " Where 納品書管理NO = " & CisFun.Val2(VsGrid1.TextMatrix(gLong, 12))
                    .SQL = .SQL & "   And 納品書頁 = " & CisFun.Val2(VsGrid1.TextMatrix(gLong, 13))
                    
                    If Not .DBExec Then GoTo PB_UPD_Er
               
                    .SQL = "UPDATE 材料発注帳票発行テーブル SET "
                    .SQL = .SQL & " 納入日 = 入荷日,"
                    .SQL = .SQL & " 納入便 = 入荷便"
                    .SQL = .SQL & "   From 材料発注帳票発行テーブル NY"
                    .SQL = .SQL & "          Left Outer Join 材料発注テーブル HC"
                    .SQL = .SQL & "            On HC.納品書管理NO = NY.納品書管理NO"
                    .SQL = .SQL & "            And HC.納入番号 = NY.納入番号"
                    .SQL = .SQL & "            And HC.納品書頁 = NY.頁NO"
                    .SQL = .SQL & "         Left Outer Join 材料入荷テーブル NC"
                    .SQL = .SQL & "             ON NC.発注管理NO = HC.発注管理NO"
                    .SQL = .SQL & "          Where HC.納品書管理NO = " & CisFun.Val2(VsGrid1.TextMatrix(gLong, 12))
                    .SQL = .SQL & "            And HC.納品書頁 = " & CisFun.Val2(VsGrid1.TextMatrix(gLong, 13))
                    .SQL = .SQL & "            And NY.カード = '6'"
                    .SQL = .SQL & "            And NY.発行区分 = 1"
               
                    If Not .DBExec Then GoTo PB_UPD_Er
               
                    '================================================ 2009/01/19 Start
                    If Not Tyokuso_Sy(gLong) Then Exit Sub         '実績管理テーブル追加/帳票発行データ追加
                    '================================================ 2009/01/19 End
                    
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
        
        CXHZ0220.Show vbModal
        Unload CXHZ0220
        Set CXHZ0220 = Nothing
        Rv_NCall = ""
        
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
    '   ( QR Reader 環境 )
    Set QRr_Object = PDQComm1
    Set QRr_MsgObject = QRMsg
    Set QRr_ErrMsgObj = ErrMsg
    Set TX2400_StartMsg = PNL_Start
    Set TX2400_ReadObj = QRRMsg
    Set TX2400_ErObj = ErrMsg
    Call QRInit
    Call TX2400InfoGet
    
    Call DispChange(ProcHB)
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
        Call .FixedSet(中中, 左中, あり, 25, 11, "材　　質")
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
    Do Until Not ZCT_RDSTS
        With VsGrid1
            ID = ID + 1
       
            Call GridMove(ID)
        End With
        Call ZCTReadNext
    Loop
    Call ZCTClose
    
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
Private Function GridSetKobetsu(Optional MsgType As Boolean = False) As Boolean
    GridSetKobetsu = False
    
    If ZCT.入荷日 <> "" Then
        If Not MsgType Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "      入荷済みデータです              "
                .MB_Button = Error
                .MBOX
            End With
        Else
            QRDataArea.エラー内容 = "入荷済み伝票"
        End If
        Exit Function
    End If
    With VsGrid1
        .Redraw = False
        
        For gCnt1 = 1 To .Rows - 1
            gCnt2 = CisFun.Val2(.TextMatrix(gCnt1, 12))
            gCnt3 = CisFun.Val2(.TextMatrix(gCnt1, 13))
            If gCnt2 <> 0 And gCnt2 = ZCT.納品書管理NO And gCnt3 = ZCT.納品書頁 Then
               If Not MsgType Then
                  With CisFun
                       .MB_Lines = 5
                       .MB_MSG(2) = "      表示済みです。               "
                       .MB_Button = Error
                       .MBOX
                   End With
                Else
                    QRDataArea.エラー内容 = "表示済み伝票"
                End If
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
            .DateE_BefVal = ZCT.納入日
            .DateEdit
            gStr = CisFun.DateE_AfVal
            If Trim(ZCT.納入便) <> "" Then
                gStr = gStr & "-" & ZCT.納入便
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
                If Trim(ZCT.納入便) <> "" Then
                    gStr = gStr & "-" & Format(iH1_NBin, "00")
                Else
                    gStr = gStr & "   "
                End If
            End With
        End If
        .TextMatrix(mRow, 3) = gStr
        
        If Trim(ZCT.仕入先名) = "" Then
           gStr = ZCT.仕入先
        Else
           gStr = RTrim(ZCT.仕入先) & Space(7 - Len(ZCT.仕入先)) & ":"
           gStr = gStr & ZCT.仕入先名
        End If
        .TextMatrix(mRow, 4) = gStr
        .TextMatrix(mRow, 5) = ZCT.受入
        .TextMatrix(mRow, 6) = ZCT.納入番号
        .TextMatrix(mRow, 7) = ZCT.納品書頁
        .TextMatrix(mRow, 8) = ZCT.発注区分名
        .TextMatrix(mRow, 9) = ZCT.手配区分名
        
        If Mid(ZCT.納入番号, 5, 1) <> "0" Then
           .TextMatrix(mRow, 11) = ZCT.材質
        End If
        
        If ZCT.実数区分 = 0 Or Right(ZCT.納入番号, 1) = "9" Then
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
    
        .TextMatrix(mRow, 12) = ZCT.納品書管理NO
        .TextMatrix(mRow, 13) = ZCT.納品書頁
    
    End With
End Sub
'*=======================================================================================================*
'*=======================================================================================================*
'*
'*                      Ｑ  Ｒ  リ  ー  ダ  (　Ｃｏｍｍ　)  処  理
'*
'*=======================================================================================================*
'*=======================================================================================================*
'+------------------------------------+
'+      読  取  処  理(QRﾀｯﾁﾘｰﾀﾞ)     +
'+------------------------------------+
Private Sub PDQComm1_OnComm()
    
    If QR_CommEvent <> "OK" Then Exit Sub
    
    With QRDataArea
        If .バーコード種類 <> "B" Then
            If .ステータス <> 0 Then
               Call QRReadMsg("E")
               Exit Sub
            End If
            ' 部品のみ有効
            If .バーコード種類 <> "Z" Then
               .エラー内容 = "材料の納品書でない"
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
    
    If Not ZCTRead(gSL_Select, 1) Then
       QRDataArea.エラー内容 = "該当データなし"
       Call QRReadMsg("E")
       Exit Sub
    End If
'        With CisFun
'            .MB_Lines = 5
'            .MB_MSG(2) = "  該当データがありません  "
'            .MB_Button = Error
'            .MBOX
'        End With
'        Exit Sub
'    End If
    
    If Not GridSetKobetsu(True) Then
       Call QRReadMsg("E")
       Exit Sub
    End If
    
End Sub
'+------------------------------------+
'+      読  取  処  理(QRﾎｯﾊﾟﾘｰﾀﾞ)     +
'+------------------------------------+
Private Sub PDQComm2_OnComm()

    Call TX2400Event
'   * エラー *
    If TX2400_Event = "ERR" Then
         QRDataArea.エラー内容 = TX2400_EventMsg
         Call TX2400QRRReadMsg("E")
         TX2400_ReadFlg = True
         Exit Sub            ' 通信エラー処理
    End If

    If TX2400_Event <> "RCV" Then Exit Sub   ' 受信ﾃﾞｰﾀ編集
    
    If Not TX2400SendProc Then Exit Sub    ' CR が存在しない時抜ける
    ' ｽﾃｰﾀｽﾁｪｯｸ
    If Not TX2400STSCheck Then
        If TX2400_STSNo <> 0 Then
           QRDataArea.エラー内容 = TX2400_STSMsg
        End If
        PNL_Start.Tag = ""
        GoTo PDQComm2_OnComm_ED
    End If
    
    If TX2400_STSNo = 1 Then
       Call TX2400OpenAuto
       Exit Sub
    End If
    If TX2400_STSNo = 2 Then
       Call TX2400ReadAuto
       Exit Sub
    End If
    
    If Trim(TX2400_CommData) = "" Then Exit Sub
    
    Call TX2400QRRReadMsg("R")
    
    ' 読取内容編集
    QRr_ReadData = TX2400_CommData
    If Not QRReadCheck Then
       TX2400_STSNo = 5
       GoTo PDQComm2_OnComm_ED
    End If
    
    With QRDataArea
        If .バーコード種類 <> "B" Then
            If .ステータス <> 0 Then
               TX2400_STSNo = 5
               GoTo PDQComm2_OnComm_ED
            End If
            ' 部品のみ有効
            If .バーコード種類 <> "Z" Then
               .エラー内容 = "材料の納品書でない"
               TX2400_STSNo = 5
               GoTo PDQComm2_OnComm_ED
            End If
            ' 納品書のみ有効
            If .カード <> "7" Then
               .エラー内容 = "カードが対象外です"
               TX2400_STSNo = 5
               GoTo PDQComm2_OnComm_ED
            End If
        End If
    End With
'( データ確認 )
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where HC.納品書管理NO = " & QRDataArea.納品書管理NO
    gSL_Select = gSL_Select & "   And HC.納品書頁 = " & QRDataArea.納品書頁
    gSL_Select = gSL_Select & mSL_Group
    
    If Not ZCTRead(gSL_Select, 1) Then
       QRDataArea.エラー内容 = "該当データなし"
       TX2400_STSNo = 5
       GoTo PDQComm2_OnComm_ED
    End If
    
    If Not GridSetKobetsu(True) Then
       TX2400_STSNo = 5
       GoTo PDQComm2_OnComm_ED
    End If
    
    Call TX2400Stack(True)
    
    Exit Sub
PDQComm2_OnComm_ED:
    TX2400_ReadFlg = True
    Call TX2400QRRReadMsg("E")
    If TX2400_STSNo = 1 Then
       Call TX2400QRRReadMsg("")
       PNL_Start.Visible = True
       DoEvents
       Call TX2400ReadAuto
    End If
    If TX2400_STSNo = 2 Then
       Call TX2400Stack(False)
    End If
    If TX2400_STSNo = 4 Then
       Call TX2400Stack(False)
    End If
    If TX2400_STSNo = 5 Then
       Call TX2400Stack(False)
    End If
    If TX2400_STSNo = 3 Then
       PNL_Start.Visible = True
       DoEvents
    End If
End Sub
'====================================================================== 2009/01/19 Start
'**************************************************
'* 【直送支給処理】
'*  　手配区分：直送 存在時,実績管理データの作成/伝票発行
'**************************************************
Private Function Tyokuso_Sy(wGyo As Long)
    Tyokuso_Sy = False
    
    Dim wNouNo      As String
    Dim wGyoNo      As Byte
    Dim wSL_Select  As String
    
    '納入番号採番用キー
    Dim Key_KTorcd As String
    Dim Key_Torcd  As String
    Dim Key_Uke    As String
    Dim Key_HCKb   As Byte
    Dim Key_NYmd   As String
    Dim Key_NBin   As String
    Dim Key_TEKb   As String
    Dim Key_SKKb   As Byte
    Dim Key_NTorcd  As String
    Dim Key_NUke    As String
        
    gSL_Select = "SELECT ZC.契約先,ZC.仕入先,ZC.受入,ZC.発注区分,ZC.納入先,ZC.納入先受入,ZN.入荷日 納入日,ZN.入荷便 納入便, " & vbCr
    gSL_Select = gSL_Select & " ZC.背番号 , ZC.収容数, ZC.枚数, ZC.端数, ZC.発注数 ," & vbCr
    gSL_Select = gSL_Select & " ZC.手配区分 , ZC.SK区分, ZC.材質, ZC.板厚, ZC.幅, ZC.長さ " & vbCr
    gSL_Select = gSL_Select & " FROM 材料発注テーブル ZC " & vbCr
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料取引先マスタ ZT " & vbCr
    gSL_Select = gSL_Select & " ON  ZN.材料管理番号 = ZT.材料管理番号 " & vbCr
    gSL_Select = gSL_Select & " AND ZN.仕入先 = ZT.取引先 " & vbCr
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料入荷テーブル ZN " & vbCr
    gSL_Select = gSL_Select & " ON  ZN.材料管理番号 = ZT.材料管理番号 " & vbCr
    gSL_Select = gSL_Select & " AND ZN.受入 = ZN.受入 " & vbCr
    gSL_Select = gSL_Select & " WHERE ISNULL(ZT.手配区分,0) = 1 " & vbCr
    gSL_Select = gSL_Select & " AND   処理端末 = HOST_NAME() " & vbCr
    gSL_Select = gSL_Select & " AND   処理区分 = 1 " & vbCr
    gSL_Select = gSL_Select & " AND   ISNULL(ZN.納入先,'') <> '' " & vbCr
    gSL_Select = gSL_Select & " AND   ISNULL(ZT.直送伝票発行区分,0) = 0 " & vbCr
    gSL_Select = gSL_Select & " AND   納品書管理NO = " & CisFun.Val2(VsGrid1.TextMatrix(wGyo, 12))
    gSL_Select = gSL_Select & " AND   納品書頁 = " & CisFun.Val2(VsGrid1.TextMatrix(wGyo, 13))
    gSL_Select = gSL_Select & " ORDER BY ZC.契約先,ZC.仕入先,ZC.受入,ZC.発注区分,ZC.納入先,ZC.納入日,ZC.納入便,"
    gSL_Select = gSL_Select & "          ZC.手配区分 , ZC.SK区分, ZC.材質, ZC.板厚, ZC.幅, ZC.長さ"
    
    wNouNo = ""
    wGyoNo = 0
    If ZCTRead(gSL_Select) Then
       If Not Tyokuso_Nouno_Get Then Exit Function            '納入番号獲得
       wNouNo = gStr
       
       Do Until Not ZNG_RDSTS
          wGyoNo = wGyoNo + 1
          
'          If (wGyoNo > 3) Or _
'             (Key_KTorcd <> ZNG.契約先 Or Key_Torcd <> ZNG.仕入先 Or Key_Uke <> ZNG.受入 Or _
'              Key_HCKb <> ZNG.発注区分 Or Key_NTorcd <> ZNG.納入先 Or Key_NUke <> ZNG.納入先受入 Or _
'              Key_NYmd <> ZNG.納入日 Or Key_NBin <> ZNG.納入便 Or Key_TEKb <> ZNG.手配区分 Or Key_SKKb <> ZNG.SK区分) Then
          If (wGyoNo > 3) Or _
             (Key_KTorcd <> ZCT.契約先 Or Key_Torcd <> ZCT.仕入先 Or Key_Uke <> ZCT.受入 Or _
              Key_HCKb <> ZCT.発注区分 Or Key_NTorcd <> ZCT.納入先 Or _
              Key_NYmd <> ZCT.納入日 Or Key_NBin <> ZCT.納入便 Or Key_TEKb <> ZCT.手配区分 Or Key_SKKb <> ZCT.SK区分) Then
             If Not Tyokuso_Nouno_Get Then Exit Function      '納入番号獲得
             wNouNo = gStr
             wGyoNo = 1
          End If
          
          Call ZJK_Add(wNouNo, wGyoNo)    ' 実績管理テーブル登録
          Call DBInsert_CHT               ' 帳票出力テーブル登録
          
          '材料入荷テーブルに支給納入番号/実績管理NOを更新
          wSL_Select = "UPDATE 材料入荷テーブル SET 支給納入番号 = '" & RTrim(wNouNo) & "',"
          wSL_Select = wSL_Select & " 支給実績管理NO = " & ZJK.実績管理NO
          wSL_Select = wSL_Select & " WHERE 発注管理NO = " & ZCT.発注管理NO
          With CisDB
              .SQL = wSL_Select
              If Not .DBExec Then Exit Function
          End With
          
          Key_KTorcd = ZCT.契約先
          Key_Torcd = ZCT.仕入先
          Key_Uke = ZCT.受入
          Key_HCKb = ZCT.発注区分
          Key_NTorcd = ZCT.納入先
          Key_NUke = ZCT.納入先受入
          Key_NYmd = ZCT.納入日
          Key_NBin = ZCT.納入便
          Key_TEKb = ZCT.手配区分
          Key_SKKb = ZCT.SK区分
          
          Call ZCTReadNext
       Loop
    
    End If
    
    Tyokuso_Sy = True
End Function
Private Function Tyokuso_Nouno_Get()
    Tyokuso_Nouno_Get = False
    
    With CisDB
    '   <><><> 納入番号の獲得 <><><>
        .SQL = "材料実績納入番号獲得"
        .StoadoCount = 2
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = OutPut用
        .ParaNo = 2: .ParaIO = OutPut用
    '
        If Not .DBStored Then
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(2) = "    材料実績納入番号の獲得に失敗しました            "
            CisFun.MB_Title = "採番処理"
            CisFun.MB_Button = Error
            CisFun.MBOX
            Exit Function
        Else
            .ParaNo = 1: gStr = .ParaValue & "0"
        End If
    End With

    Tyokuso_Nouno_Get = True
    
End Function
Private Sub ZJK_Add(NouNo As String, ID As Byte)
    
    With CisDB
    '   <><><> 実績管理番号の獲得 <><><>
        .SQL = "番号獲得"
        .StoadoCount = 3
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = OutPut用
        .ParaNo = 3: .ParaIO = Return用
    '
        .ParaNo = 1: .ParaValue = 10
    '
        If Not .DBStored Then
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(2) = "    材料実績管理番号の獲得に失敗しました            "
            CisFun.MB_Title = "採番処理"
            CisFun.MB_Button = Error
            CisFun.MBOX
            Exit Sub
        Else
            .ParaNo = 2: ZJK.実績管理NO = .ParaValue
        End If
    End With

    With ZJK
        .実績区分 = 3
        .実績先 = ZCT.納入先
        .受入 = ZCT.納入先受入
        '.実績日 = RTrim(iH1_Ymd)               '2014/02/04
        .実績日 = RTrim(iH1_NYmd)               '2014/02/04
        '.実績便 = RTrim(iH1_Bin)               '2014/02/04
        .実績便 = RTrim(iH1_NBin)               '2014/02/04
        .納入番号 = RTrim(NouNo)
        
        .納品書行 = ID
        .材質 = ZCT.材質
        .板厚 = ZCT.板厚
        .幅 = ZCT.幅
        .長さ = ZCT.長さ
        .背番号 = ZCT.背番号
        .収容数 = ZCT.収容数
        .実績枚数 = ZCT.枚数
        .実績端数 = ZCT.端数
        .実績数 = ZCT.発注数
        .有無償区分 = 1
        .入力種類 = "A"             '自動作成
    
        .作成者 = gTanto
    End With
    Call ZJKInsert

End Sub
'+------------------------------------+
'+      追加処理 - 帳票出力テーブル
'+------------------------------------+
Private Function DBInsert_CHT() As Boolean
    DBInsert_CHT = False

    Call ItemsClearCHT
    With CHT
         .端末番号 = "' + substring(host_name(),1,20) + '"
         .帳票区分 = 302
        '.管理文字 = ""
         .検索情報 = ZJK.納入番号
         .管理NO = ZJK.実績管理NO
         .管理行NO = 0
        '.補助番号 = 0
        '.出力先 = 0
     End With
    Call CHTInsert
    
    DBInsert_CHT = True
DBInsert_CHT_Err:
End Function
'+------------------------------------+
'+      削除処理 - 帳票出力テーブル
'+------------------------------------+
Private Function DBDelete_CHT(Optional DBNo As Integer) As Boolean
    DBDelete_CHT = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "delete 帳票出力テーブル"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20) "
    gSL_Select = gSL_Select & "   And 帳票区分 = 302 "

    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With

    DBDelete_CHT = True
End Function
'+----------------------------+
'+    印　刷  処　理            +
'+----------------------------+
Private Sub Prt_Nou()
''Exit Sub '試験用
    gSL_Select = "SELECT 検索情報 "
    gSL_Select = gSL_Select & " From 帳票出力テーブル "
    gSL_Select = gSL_Select & " Where 端末番号 =  HOST_NAME() "
    gSL_Select = gSL_Select & "   And 帳票区分 =  302 "
    gSL_Select = gSL_Select & "   And 検索情報 <> '' "
    gSL_Select = gSL_Select & " Group By 検索情報 "
    
    If CHTRead(gSL_Select) Then             ' 帳票出力テーブルに該当データがあれば印刷する
       CXHZ0530_PrtCnt = CisDB.RecordCount  ' 印刷枚数
       Call CXHZ0530LPrint
    End If
    Call CHTClose
End Sub
'====================================================================== 2009/01/19 End





