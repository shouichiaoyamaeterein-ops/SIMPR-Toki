VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CisYMDwB_6.0.ocx"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Object = "{490AF35D-478A-11D4-AEC2-00000EA57F9E}#4.0#0"; "CISMSG~1.OCX"
Begin VB.Form CXH0220 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "入荷実績入力 [個別]"
   ClientHeight    =   10290
   ClientLeft      =   1755
   ClientTop       =   4575
   ClientWidth     =   14580
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
   ScaleHeight     =   10290
   ScaleWidth      =   14580
   WindowState     =   2  '最大化
   Begin VB.PictureBox PNL_Start 
      BackColor       =   &H00000000&
      Height          =   585
      Left            =   3360
      ScaleHeight     =   525
      ScaleWidth      =   8235
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   1335
      Visible         =   0   'False
      Width           =   8295
      Begin CisMsgRelate_v60.CISMsgRelate ProcMsg 
         Height          =   405
         Left            =   75
         Top             =   60
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   714
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
   Begin Cis3D_v60.CIS3D H4_Area1 
      Height          =   1320
      Left            =   240
      Top             =   2760
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
         TabIndex        =   55
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
         Caption         =   "QRﾎｯﾊﾟﾘｰﾀﾞ読取中"
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
   Begin Cis3D_v60.CIS3D H3_Area1 
      Height          =   1305
      Left            =   105
      Top             =   4410
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   2302
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
         Left            =   0
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   0
         Width           =   240
      End
      Begin Cis3D_v60.CIS3D QRMsg 
         Height          =   1155
         Left            =   90
         Top             =   90
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   2037
         BackColor       =   12648384
         ForeColor       =   16711680
         Caption         =   "QRリーダー読取中"
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
   Begin Cis3D_v60.CIS3D ErrMsg 
      Height          =   465
      Left            =   1140
      Top             =   3225
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
   Begin PdqcommLib.PDQComm PDQComm1 
      Height          =   480
      Left            =   825
      TabIndex        =   54
      Top             =   1275
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
   Begin Cis3D_v60.CIS3D Back_BunCount 
      Height          =   795
      Left            =   1095
      Top             =   1140
      Visible         =   0   'False
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   1402
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
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   705
         Left            =   45
         Top             =   45
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1244
         ForeColor       =   128
         Caption         =   "分納処理件数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   60
         Begin Cis3D_v60.CIS3D Bun_Count 
            Height          =   630
            Left            =   1950
            Top             =   30
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   1111
            Caption         =   "ZZZ / ZZZ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   18
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
      Height          =   1110
      Left            =   4170
      ScaleHeight     =   1050
      ScaleWidth      =   2025
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   2085
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   930
         Index           =   1
         Left            =   75
         Top             =   60
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1640
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
         Begin CisText_V60.CisText iH2_Bar 
            Height          =   360
            Left            =   150
            TabIndex        =   3
            Top             =   420
            Width           =   1590
            _ExtentX        =   2805
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
   Begin Cis3D_v60.CIS3D CIS3D1 
      Height          =   375
      Left            =   9510
      Top             =   0
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   661
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
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   345
         Left            =   60
         Top             =   15
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
         Caption         =   "実績枚数"
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
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D lb_JMai 
            Height          =   270
            Left            =   975
            Top             =   30
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   476
            BackColor       =   12648384
            Caption         =   "#,###"
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
            cAlingnment     =   7
            cBoderStyle     =   1
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
      Height          =   1110
      Left            =   1110
      ScaleHeight     =   1050
      ScaleWidth      =   3735
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2025
      Width           =   3795
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   960
         Left            =   2115
         Top             =   45
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   1693
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
         Begin CisText_V60.CisText iH1_Nouno 
            Height          =   360
            Left            =   120
            TabIndex        =   1
            Top             =   420
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
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   960
         Left            =   3105
         Top             =   45
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   1693
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
         Begin CisText_V60.CisText iH1_Page 
            Height          =   360
            Left            =   75
            TabIndex        =   2
            Top             =   420
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
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   960
         Left            =   75
         Top             =   45
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1693
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_NYmd 
            Height          =   360
            Left            =   90
            TabIndex        =   0
            Top             =   420
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
      Left            =   8115
      ScaleHeight     =   660
      ScaleWidth      =   6045
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   570
      Width           =   6105
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   3465
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
         Left            =   4320
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
         Left            =   5175
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
         Left            =   45
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
         Left            =   2610
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
         Left            =   1755
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
      Begin CisBtn_60.CisBtn PB_CHG 
         Height          =   585
         Left            =   900
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
   Begin VB.PictureBox B1_Area1 
      Height          =   6075
      Left            =   1095
      ScaleHeight     =   6015
      ScaleWidth      =   13065
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   3735
      Width           =   13125
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   5895
         Left            =   60
         Top             =   45
         Width           =   12945
         _ExtentX        =   22834
         _ExtentY        =   10398
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
         Begin Cis3D_v60.CIS3D CIS3D17 
            Height          =   4845
            Left            =   8130
            Top             =   975
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   8546
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   1
            cPositionY      =   30
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   19
               Left            =   15
               Top             =   4380
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   9
                  Left            =   30
                  TabIndex        =   43
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   18
               Left            =   15
               Top             =   3945
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   8
                  Left            =   30
                  TabIndex        =   39
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   17
               Left            =   15
               Top             =   3510
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   7
                  Left            =   30
                  TabIndex        =   35
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   16
               Left            =   15
               Top             =   3075
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   6
                  Left            =   30
                  TabIndex        =   31
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   15
               Left            =   15
               Top             =   2640
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   5
                  Left            =   30
                  TabIndex        =   27
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   14
               Left            =   15
               Top             =   2205
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   4
                  Left            =   30
                  TabIndex        =   23
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   13
               Left            =   15
               Top             =   1770
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   3
                  Left            =   30
                  TabIndex        =   19
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   12
               Left            =   15
               Top             =   1335
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   2
                  Left            =   30
                  TabIndex        =   15
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   11
               Left            =   15
               Top             =   900
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   1
                  Left            =   30
                  TabIndex        =   11
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D18 
               Height          =   450
               Left            =   1065
               Top             =   15
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   794
               ForeColor       =   16711680
               Caption         =   "分割納入数"
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
            Begin Cis3D_v60.CIS3D CIS3D19 
               Height          =   450
               Left            =   2415
               Top             =   15
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   794
               ForeColor       =   16711680
               Caption         =   " 完納予定日  - 便"
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   0
               Left            =   1065
               Top             =   465
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   0
                  Left            =   195
                  TabIndex        =   8
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   0
               Left            =   2415
               Top             =   465
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   0
                  Left            =   45
                  TabIndex        =   9
                  Top             =   60
                  Width           =   1770
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
                  Object.Width           =   1770
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
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   0
                  Left            =   1830
                  TabIndex        =   10
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   1
               Left            =   1065
               Top             =   900
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   1
                  Left            =   195
                  TabIndex        =   12
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   1
               Left            =   2415
               Top             =   900
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   1
                  Left            =   1830
                  TabIndex        =   14
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
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
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   1
                  Left            =   60
                  TabIndex        =   13
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   2
               Left            =   1065
               Top             =   1335
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   2
                  Left            =   195
                  TabIndex        =   16
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   2
               Left            =   2415
               Top             =   1335
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   2
                  Left            =   1830
                  TabIndex        =   18
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
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
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   2
                  Left            =   60
                  TabIndex        =   17
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   3
               Left            =   1065
               Top             =   1770
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   3
                  Left            =   195
                  TabIndex        =   20
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   3
               Left            =   2415
               Top             =   1770
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   3
                  Left            =   1830
                  TabIndex        =   22
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
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
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   3
                  Left            =   60
                  TabIndex        =   21
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   4
               Left            =   1065
               Top             =   2205
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   4
                  Left            =   195
                  TabIndex        =   24
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   4
               Left            =   2415
               Top             =   2205
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   4
                  Left            =   1830
                  TabIndex        =   26
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
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
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   4
                  Left            =   60
                  TabIndex        =   25
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   5
               Left            =   1065
               Top             =   2640
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   5
                  Left            =   195
                  TabIndex        =   28
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   5
               Left            =   2415
               Top             =   2640
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   5
                  Left            =   1830
                  TabIndex        =   30
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
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
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   5
                  Left            =   60
                  TabIndex        =   29
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   6
               Left            =   1065
               Top             =   3075
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   6
                  Left            =   195
                  TabIndex        =   32
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   6
               Left            =   2415
               Top             =   3075
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   6
                  Left            =   1830
                  TabIndex        =   34
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
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
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   6
                  Left            =   60
                  TabIndex        =   33
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   7
               Left            =   1065
               Top             =   3510
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   7
                  Left            =   195
                  TabIndex        =   36
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   7
               Left            =   2415
               Top             =   3510
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   7
                  Left            =   1830
                  TabIndex        =   38
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
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
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   7
                  Left            =   60
                  TabIndex        =   37
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   8
               Left            =   1065
               Top             =   3945
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   8
                  Left            =   195
                  TabIndex        =   40
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   8
               Left            =   2415
               Top             =   3945
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   8
                  Left            =   1830
                  TabIndex        =   42
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
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
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   8
                  Left            =   60
                  TabIndex        =   41
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   9
               Left            =   1065
               Top             =   4380
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_Suryo 
                  Height          =   360
                  Index           =   9
                  Left            =   195
                  TabIndex        =   44
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   435
               Index           =   9
               Left            =   2415
               Top             =   4380
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_BBin 
                  Height          =   360
                  Index           =   9
                  Left            =   1830
                  TabIndex        =   46
                  Top             =   45
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
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
               Begin CisYMD_With_Btn.CisYMDwB iB1_BYmd 
                  Height          =   330
                  Index           =   9
                  Left            =   60
                  TabIndex        =   45
                  Top             =   45
                  Width           =   1770
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
                  Object.Width           =   1770
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
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   450
               Left            =   15
               Top             =   15
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   794
               ForeColor       =   16711680
               Caption         =   "入荷数"
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   435
               Index           =   10
               Left            =   15
               Top             =   465
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   767
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               Begin CisText_V60.CisText iB1_NSuryo 
                  Height          =   360
                  Index           =   0
                  Left            =   30
                  TabIndex        =   7
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   635
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
         End
         Begin vsFlexLib.vsFlexArray VsGrid1 
            Height          =   4860
            Left            =   90
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   960
            Width           =   7920
            _Version        =   196608
            _ExtentX        =   13970
            _ExtentY        =   8572
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
            FormatString    =   "№|品　番                   |背番号 |収容数 |箱 数  |発注数  |分割納入数|完納予定日|予定便|"
            Rows            =   11
            BackColor       =   12648447
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   0
            RowHeightMin    =   440
            GridLines       =   3
            SelectionMode   =   1
            BackColorAlternate=   12648447
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   735
            Left            =   10425
            Top             =   135
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   1296
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "   入 荷 日  - 便"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   1
            Begin CisText_V60.CisText iB1_SBin 
               Height          =   375
               Left            =   1890
               TabIndex        =   6
               Top             =   285
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
            Begin CisYMD_With_Btn.CisYMDwB iB1_SYmd 
               Height          =   330
               Left            =   75
               TabIndex        =   5
               Top             =   300
               Width           =   1770
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
               Object.Width           =   1770
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
      TabIndex        =   47
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
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 入荷実績入力 [個別]  】"
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14280
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
         Left            =   11790
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
   End
   Begin VB.PictureBox Back_HInfo 
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
      Height          =   1110
      Left            =   5670
      ScaleHeight     =   1050
      ScaleWidth      =   8490
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   2025
      Width           =   8550
      Begin Cis3D_v60.CIS3D Back_Page 
         Height          =   960
         Left            =   7980
         Top             =   60
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   1693
         ForeColor       =   0
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
         Begin Cis3D_v60.CIS3D B1lb_Page 
            Height          =   360
            Left            =   75
            Top             =   375
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "99"
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
         End
      End
      Begin Cis3D_v60.CIS3D Back_Nounox 
         Height          =   960
         Left            =   7050
         Top             =   60
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1693
         ForeColor       =   0
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
         Begin Cis3D_v60.CIS3D B1lb_Nouno 
            Height          =   360
            Left            =   75
            Top             =   390
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXX"
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
            cPositionX      =   60
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   960
         Left            =   3615
         Top             =   60
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1693
         ForeColor       =   0
         Caption         =   " 納 入 日 - 便"
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
         Begin Cis3D_v60.CIS3D B1lb_NYmd 
            Height          =   360
            Left            =   90
            Top             =   405
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "9999/99/99"
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
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D B1lb_NBin 
            Height          =   360
            Left            =   1470
            Top             =   405
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "99"
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
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   960
         Left            =   60
         Top             =   60
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   1693
         ForeColor       =   0
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
            Left            =   60
            Top             =   420
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXX:XXXXXXXXX1XXXXXXXXX2"
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
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   960
         Left            =   5490
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1693
         ForeColor       =   0
         Caption         =   "発注区"
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
         Begin Cis3D_v60.CIS3D B1lb_HKbn 
            Height          =   360
            Left            =   75
            Top             =   405
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXX"
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
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   960
         Left            =   6285
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   1693
         ForeColor       =   0
         Caption         =   "手配区"
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
         Begin Cis3D_v60.CIS3D B1lb_TKbn 
            Height          =   360
            Left            =   75
            Top             =   390
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXX"
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
            cAlingnment     =   7
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
   End
   Begin PdqcommLib.PDQComm PDQComm2 
      Height          =   480
      Left            =   90
      TabIndex        =   56
      Top             =   1710
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
Attribute VB_Name = "CXH0220"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   入荷実績入力 [個別]
'**       フォームID    :   CXH0220
'**       処理概要      :
'**
'**       作  成  日    :   2004/02/06  By CIS
'**       変  更  日    :   2004/04/08  By CIS - 更新時の確認メッセージ修正
'**       変  更  日    :   2005/02/23  By CIS ﾁｹｯﾄ対応
'**       変  更  日    :   2005/03/04  By CIS - 完納予定日･便自動算出
'**       変  更  日    :   2005/06/17  By CIS - 実数区分=1の時、出荷数＋分納数＞受注数をエラーとする
'**       変  更  日    :   2005/08/31  By CIS JAMA対応
'**       変  更  日    :   2005/11/17  By CIS オプション追加（稼働日カレンダー）
'**                                             発注日算出=自社カレンダー、納入日算出=仕入先カレンダー
'**       変  更  日    :   2007/07/02  By CIS 先行納入可に変更
'**       変  更  日    :   2007/07/13  By CIS 仕入先･受入カレンダー対応
'**       変  更  日    :   2008/04/11  By CIS 背番号８桁対応
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    
    Dim mSL_Select              As String
    
    Dim CisVsGrid               As New CisVsGrid3
    
    Dim mHeadProc               As String
    Dim mSaveHeadProc           As String
    
    Dim mBunCount               As Long
    '+-------------------------------------------- 2005/03/04 Insert
    Dim mBYMD(10)               As String
    Dim mBBin(10)               As String
    Dim mAutoInput              As Boolean
    '+-------------------------------------------- 2005/03/04 Insert  End
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
        '   ( QR Reader 環境 )
        Set QRr_Object = PDQComm1
        Set QRr_MsgObject = QRMsg
        Set QRr_ErrMsgObj = ErrMsg
        Set TX2400_StartMsg = PNL_Start
        Set TX2400_ReadObj = QRRMsg
        Set TX2400_ErObj = ErrMsg
        Call QRInit
        Call TX2400InfoGet
    End If
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    QRDummy.Left = -1000
    QRRDummy.Left = -1000
    
    H2_Area1.Move H1_Area1.Left, H1_Area1.Top
    H3_Area1.Move H1_Area1.Left, H1_Area1.Top
    H4_Area1.Move H1_Area1.Left, H1_Area1.Top
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = VsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 440
        .PatanMax = 1
        .SelectCol = 10
        .InitGet = False
        .Init
    End With
    Call GridInit
    
'   #-------------------#
'   #  項 目 初 期 化   #
'   #-------------------#
    Call CisFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    K_Sykbnm.Tag = "4"
    Call PB_CHG_Click
    
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    If Rv_NCall = "" Then
       Me.Show
    Else
    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = Not FormAct

'システムオプション獲得     2005/11/17追加
    Call GetSYSOption(33)
'
    mSL_Select = "SELECT HC.納品書管理NO,HC.納品書行,HC.納入日,HC.納入便,"
    mSL_Select = mSL_Select & "  ISNULL(NK.入荷日,'') 入荷日,ISNULL(NK.入荷便,'') 入荷便,"
    mSL_Select = mSL_Select & "  ISNULL(NK.挽回日,'') 挽回日,ISNULL(NK.挽回便,'') 挽回便,"
    mSL_Select = mSL_Select & "  ISNULL(NK.分割数,0) 分割数,ISNULL(NK.入荷数,0) 入荷数,"
    mSL_Select = mSL_Select & "  HC.仕入先,HC.受入,HC.納入番号,HC.納品書頁,"
    mSL_Select = mSL_Select & "  HC.発注区分,HC.手配区分,HC.SK区分,"
    mSL_Select = mSL_Select & "  ISNULL(TR.略称,'') 仕入先名,"
    mSL_Select = mSL_Select & "  ISNULL(SY1.値名称,'') 発注区分名,"
    mSL_Select = mSL_Select & "  ISNULL(SY2.値名称,'') 手配区分名,"
    mSL_Select = mSL_Select & "  HN.表示品番 表示品番,HC.背番号,"
    mSL_Select = mSL_Select & "  HC.収容数,HC.枚数,HC.発注数,HN.実数区分 "
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


    If Rv_NCall = "K" Then
       K_Sykbnm.Tag = "4"
       Call PB_CHG_Click
       Back_BunCount.Visible = True
       mBunCount = 1
       
       Call ContinuChk
       
    End If

    ' 入荷実績処理ワーク削除
    If Rv_NCall = "" Then
        With CisDB
            .SQL = "Delete 入荷実績処理ワーク"
            .SQL = .SQL & "    Where 処理区分 = 9"
            .DBExec
        End With
    End If

    '+---------------------------------------------------- 2005/03/04 Insert
    '完納予定日･便算出方法
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 16
    Call IRNRead("", 1)
    If Mid(IRN.oyian, 1, 1) = "1" Then
       mAutoInput = True
    Else
       mAutoInput = False
    End If
    '+---------------------------------------------------- 2005/03/04 Insert End
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Rv_NCall = "" Then
       Cancel = 1
    End If
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
            Case vbKeyF6:       If PB_CHG.Visible Then Call PB_CHG_Click        '【切替】
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyPageUp:    If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown:  If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                                If PB_Ent.Visible Then                      '【入力】
                                    Call PB_ENT_Click
                                End If
           Case Else:           Exit Function
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
                .ShellTask NoFreeze
            End With
            On Error GoTo 0
          End If
        End If
        
        Call CisDB.DBDISConnect                            ' ﾃﾞｰﾀﾍﾞｰｽ切断
        Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
        End
    Else
        Timer1.Enabled = False
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
        
        If Rv_NCall = "K" Then
           mBunCount = mBunCount + 1
           If ContinuChk Then
              GoTo PB_CAN_Ed
           Else
              Call PB_END_Click
              Exit Sub
           End If
        End If
        ProcHB = mHeadProc
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        
        If ProcHB = "H4" Then
            Call TX2400OpenAuto
        End If
        
        GoTo PB_CAN_Ed
    End If
    
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+------------------+
'+    切替   (F6)   +
'+------------------+
Private Sub PB_CHG_Click()
    gLong = CisFun.Val2(K_Sykbnm.Tag)
    
    gLong = gLong + 1
    If gLong > 4 Then gLong = 1
    
    mSaveHeadProc = ""
    mSaveHeadProc = "H1"
    If IRN_NKRead = 1 Then
       mSaveHeadProc = mSaveHeadProc & "H3"
    Else
       mSaveHeadProc = mSaveHeadProc & "  "
    End If
'    If IRN_NKReadHP = 1 Then
'       mSaveHeadProc = mSaveHeadProc & "H4"
'    Else
       mSaveHeadProc = mSaveHeadProc & "  "
'    End If
    If IRN_NKReadBar = 1 Then
       mSaveHeadProc = mSaveHeadProc & "H2"
    Else
       mSaveHeadProc = mSaveHeadProc & "  "
    End If
    
    mHeadProc = ""
    gInt = 1
    Do Until Trim(mHeadProc) <> ""
        mHeadProc = Mid(mSaveHeadProc, gLong + (gLong - 1), 2)
        If Trim(mHeadProc) <> "" Then
           K_Sykbnm.Tag = gLong
           Exit Do
        End If
        gLong = gLong + 1
        If gLong > 4 Then gLong = 1
    Loop
    
    H1_Area1.Visible = False
    H2_Area1.Visible = False
    H3_Area1.Visible = False
    H4_Area1.Visible = False
    Select Case mHeadProc
        Case "H2"
            K_Sykbnm.Caption = "【 ﾊﾞｰｺｰﾄﾞ 】"
            H2_Area1.Visible = True
        Case "H3"
            K_Sykbnm.Caption = "【 QRﾀｯﾁﾘｰﾀﾞ 】"
            H3_Area1.Visible = True
            '   ( QR Reader 環境 )
            Set QRr_Object = PDQComm1
            Set QRr_MsgObject = QRMsg
            Set QRr_ErrMsgObj = ErrMsg
            Call QRInit
        Case "H4"
            K_Sykbnm.Caption = "【QRﾎｯﾊﾟﾘｰﾀﾞ】"
            H4_Area1.Visible = True
            '   ( QR Reader 環境 )
            Set QRr_Object = PDQComm2
            Set QRr_MsgObject = QRRMsg
            Set QRr_ErrMsgObj = ErrMsg
            Set TX2400_StartMsg = PNL_Start
            Set TX2400_ReadObj = QRRMsg
            Call TX2400InfoGet
        Case Else
            K_Sykbnm.Caption = "【 納番 】"
            H1_Area1.Visible = True
    End Select
    ProcHB = mHeadProc
    Call DispChange(ProcHB)
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = mSL_Select
    gSL_Select = gSL_Select & " WHERE  HC.納入日 + HC.納入番号 + Cast(HC.納品書頁 as Char(2)) < "
    gSL_Select = gSL_Select & " '" & RTrim$(iH1_NYmd) & RTrim$(iH1_Nouno) & Format(iH1_Page, "0") & "'"
    gSL_Select = gSL_Select & " And ISNULL(入荷日,'') = '' "
    gSL_Select = gSL_Select & " ORDER BY HC.納入日 + HC.納入番号 + Cast(HC.納品書頁 as Char(2)) DESC "
    
    If Not HCTRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    先頭レコードです            "
            .MB_Title = "【前頁】"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Back_ED
    End If
    
    iH1_NYmd = HCT.納入日
    iH1_Nouno = HCT.納入番号
    iH1_Page = HCT.納品書頁
    
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where HC.納入番号 = '" & RTrim$(iH1_Nouno) & "'"
    gSL_Select = gSL_Select & " And   HC.納品書頁 =  " & CisFun.Val2(iH1_Page)
    gSL_Select = gSL_Select & " And   HC.納入日   = '" & RTrim$(iH1_NYmd) & "'"
    
    Call DBInput("Back")
PB_Back_ED:
    
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = mSL_Select
    gSL_Select = gSL_Select & " WHERE  HC.納入日 + HC.納入番号 + Cast(HC.納品書頁 as Char(2)) > "
    gSL_Select = gSL_Select & " '" & RTrim$(iH1_NYmd) & RTrim$(iH1_Nouno) & Format(iH1_Page, "0") & "'"
    gSL_Select = gSL_Select & " And ISNULL(入荷日,'') = '' "
    gSL_Select = gSL_Select & " ORDER BY HC.納入日 + HC.納入番号 + Cast(HC.納品書頁 as Char(2)) ASC "
    
    If Not HCTRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    最終レコードです            "
            .MB_Title = "【次頁】"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Next_ED
    End If
    
    iH1_NYmd = HCT.納入日
    iH1_Nouno = HCT.納入番号
    iH1_Page = HCT.納品書頁
    
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where HC.納入番号 = '" & RTrim$(iH1_Nouno) & "'"
    gSL_Select = gSL_Select & " And   HC.納品書頁 =  " & CisFun.Val2(iH1_Page)
    gSL_Select = gSL_Select & " And   HC.納入日   = '" & RTrim$(iH1_NYmd) & "'"
    
    Call DBInput("Next")
PB_Next_ED:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_NYmd" Then iH1_NYmd.ShowCalender
    If PB_Look.Tag = "iB1_SYmd" Then iB1_SYmd.ShowCalender
    If PB_Look.Tag Like "iB1_BYmd*" Then iB1_BYmd(CisFun.Val2(Mid(PB_Look.Tag, 9))).ShowCalender
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*----------------------------------------------*
'*      納 入 日  ( iH1_NYmd )       *
'*----------------------------------------------*
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
'*      出 荷 日  ( iB1_SYmd )       *
'*----------------------------------------------*
Private Sub iB1_SYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_SYmd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*----------------------------------------------*
'*       入荷数  ( iB1_NSuryo )                 *
'*----------------------------------------------*
Private Sub iB1_NSuryo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim ID      As Integer
    If KeyCode = vbKeyDown Then
       ID = Index + 1
       Do Until ID > 9
          If iB1_NSuryo(ID).Visible Then
             iB1_NSuryo(ID).SetFocus
             Exit Do
          End If
          ID = ID + 1
       Loop
       Exit Sub
    End If
    If KeyCode = vbKeyUp Then
       ID = Index - 1
       Do Until ID < 0
          If iB1_NSuryo(ID).Visible Then
             iB1_NSuryo(ID).SetFocus
             Exit Do
          End If
          ID = ID - 1
       Loop
       Exit Sub
    End If
End Sub
Private Sub iB1_NSuryo_LostFocus(Index As Integer)
    If Trim(VsGrid1.TextMatrix(Index + 1, 6)) <> "" Then Exit Sub
'    If CisFun.Val2(iB1_Suryo(Index).Tag) = 0 Then
       iB1_Suryo(Index) = CisFun.RSetFld(CisFun.Val2(iB1_NSuryo(Index).Tag) - CisFun.Val2(iB1_NSuryo(Index)), 0, iB1_Suryo(0).cFormat)
'    End If
'+--------------------------------------------------------------- 2005/03/04 Insert
    If CisFun.Val2(iB1_Suryo(Index)) = 0 Then Exit Sub
' 挽回日を自動算出
    If GetBYMD(True, Index) Then
       If Trim(iB1_BYmd(Index) & iB1_BBin(Index)) = "" Then
          iB1_BYmd(Index) = mBYMD(Index)
          If iB1_BBin(Index).Visible Then
             iB1_BBin(Index) = mBBin(Index)
          End If
       End If
    End If
'+--------------------------------------------------------------- 2005/03/04 Insert End
End Sub
'*----------------------------------------------*
'*       分割納入数  ( iB1_Suryo )              *
'*----------------------------------------------*
Private Sub iB1_Suryo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim ID      As Integer
    If KeyCode = vbKeyDown Then
       ID = Index + 1
       Do Until ID > 9
          If iB1_Suryo(ID).Visible Then
             iB1_Suryo(ID).SetFocus
             Exit Do
          End If
          ID = ID + 1
       Loop
       Exit Sub
    End If
    If KeyCode = vbKeyUp Then
       ID = Index - 1
       Do Until ID < 0
          If iB1_Suryo(ID).Visible Then
             iB1_Suryo(ID).SetFocus
             Exit Do
          End If
          ID = ID - 1
       Loop
       Exit Sub
    End If
End Sub
Private Sub iB1_Suryo_LostFocus(Index As Integer)
    If Trim(VsGrid1.TextMatrix(Index + 1, 6)) <> "" Then Exit Sub
'    If CisFun.Val2(iB1_Suryo(Index)) <> 0 Then
       iB1_NSuryo(Index) = CisFun.RSetFld(CisFun.Val2(iB1_NSuryo(Index).Tag) - CisFun.Val2(iB1_Suryo(Index)), 0, iB1_NSuryo(0).cFormat)
'    End If
'+--------------------------------------------------------------- 2005/03/04 Insert
    If CisFun.Val2(iB1_Suryo(Index)) = 0 Then Exit Sub
' 挽回日を自動算出
    If GetBYMD(True, Index) Then
       If Trim(iB1_BYmd(Index) & iB1_BBin(Index)) = "" Then
          iB1_BYmd(Index) = mBYMD(Index)
          If iB1_BBin(Index).Visible Then
             iB1_BBin(Index) = mBBin(Index)
          End If
       End If
    End If
'+--------------------------------------------------------------- 2005/03/04 Insert End
End Sub
'*----------------------------------------------*
'*       完納予定日  ( iB1_BYmd )       *
'*----------------------------------------------*
Private Sub iB1_BYmd_GotFocus(Index As Integer)
    PB_Look.Tag = ActiveControl.Name & Index
    PB_Look.Visible = True
End Sub
Private Sub iB1_BYmd_LostFocus(Index As Integer)
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
Private Sub iB1_BYmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim ID      As Integer
    If KeyCode = vbKeyDown Then
       ID = Index + 1
       Do Until ID > 9
          If iB1_BYmd(ID).Visible Then
             iB1_BYmd(ID).SetFocus
             Exit Do
          End If
          ID = ID + 1
       Loop
       Exit Sub
    End If
    If KeyCode = vbKeyUp Then
       ID = Index - 1
       Do Until ID < 0
          If iB1_BYmd(ID).Visible Then
             iB1_BYmd(ID).SetFocus
             Exit Do
          End If
          ID = ID - 1
       Loop
       Exit Sub
    End If
End Sub
'*----------------------------------------------*
'*       完納予定便  ( iB1_BBin )               *
'*----------------------------------------------*
Private Sub iB1_BBin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim ID      As Integer
    If KeyCode = vbKeyDown Then
       ID = Index + 1
       Do Until ID > 9
          If iB1_BBin(ID).Visible Then
             iB1_BBin(ID).SetFocus
             Exit Do
          End If
          ID = ID + 1
       Loop
       Exit Sub
    End If
    If KeyCode = vbKeyUp Then
       ID = Index - 1
       Do Until ID < 0
          If iB1_BBin(ID).Visible Then
             iB1_BBin(ID).SetFocus
             Exit Do
          End If
          ID = ID - 1
       Loop
       Exit Sub
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
    Dim H1Mode          As Boolean
    Dim H1Color         As Long
    Dim H2Mode          As Boolean
    Dim H2Color         As Long
    Dim H3Mode          As Boolean
    Dim H3Color         As Long
    Dim H4Mode          As Boolean
    Dim H4Color         As Long
    Dim B1Mode          As Boolean
    Dim B1Color         As Long
    
    H1Mode = False:     H1Color = gPLostSel
    H2Mode = False:     H2Color = gPLostSel
    H3Mode = False:     H3Color = gPLostSel
    H4Mode = False:     H4Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "H2":      H2Mode = True:  H2Color = gPGotSel
        Case "H3":      H3Mode = True:  H3Color = gPGotSel
        Case "H4":      H4Mode = True:  H4Color = gPGotSel
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
    
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    
    PB_CAN.Visible = B1Mode
    PB_CHG.Visible = Not B1Mode
    If Trim(mSaveHeadProc) = "H1" Then
       PB_CHG.Visible = False
    End If
    PB_BACK.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    
    If Rv_NCall = "" Then
    Else
       PB_END.BtnCaption = "戻る"
    End If
'
    Set QRr_Object = PDQComm1
    Set TX2400_Comm = PDQComm2
    Call QR_Comm_Close
    Call TX2400Close
    If ProcHB = "H3" Then
       Call QR_Comm_Open
    End If
    If ProcHB = "H4" Then
       If TX2400_ReadAuto Then
          ProcMsg.Caption = "用紙をホッパーにセットして下さい。"
       Else
          ProcMsg.Caption = "用紙をﾎｯﾊﾟｰにｾｯﾄし【ｽﾀｰﾄ】キーを押して下さい。"
       End If
       PNL_Start.Tag = ""
       Call TX2400OpenAuto
    End If
    PNL_Start.Visible = H4Mode
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
    If ProcHB = "H2" Then
        If Not Head2Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If

    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        
        If Rv_NCall = "K" Then
           mBunCount = mBunCount + 1
           If ContinuChk Then
              GoTo ReturnPress_Ed
           Else
              Call PB_END_Click
           End If
        End If
        
        Call HeadBodyClear("B")
        ProcHB = mHeadProc
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
    Head1Chk = False
    
'( 納入日 )
    If RTrim$(iH1_NYmd) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  納入日を入力して下さい  "
            .MB_Button = Error
            .MBOX
        End With
        iH1_NYmd.SetFocus
        Exit Function
    Else
        If Not iH1_NYmd.cChkResult Then
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
    
'( 納入番号 )
    If RTrim$(iH1_Nouno) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  納番を入力して下さい  "
            .MB_Button = Error
            .MBOX
        End With
        iH1_Nouno.SetFocus
        Exit Function
    End If
    
'( 納品書頁 )
    If CisFun.Val2(iH1_Page) = 0 Then iH1_Page = " 1"
    
    If Rv_NCall = "" Or Rv_NCall = "K2" Or CisFun.Val2(iH1_Nouno.Tag) = 0 Then
        gSL_Select = "Select HC.納品書管理NO From 発注テーブル HC"
        gSL_Select = gSL_Select & " Where HC.納入番号 = '" & RTrim$(iH1_Nouno) & "'"
        gSL_Select = gSL_Select & " And   HC.納品書頁 =  " & CisFun.Val2(iH1_Page)
        gSL_Select = gSL_Select & " And   HC.納入日   = '" & RTrim$(iH1_NYmd) & "'"
        gSL_Select = gSL_Select & " Group By HC.納品書管理NO"
        If Not HCTRead(gSL_Select) Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "  該当データがありません  "
                .MB_Button = Error
                .MBOX
            End With
            Exit Function
        End If
        If CisDB.RecordCount = 1 Then
           iH1_Nouno.Tag = HCT.納品書管理NO
           Call HCTClose
        Else
            gSL_Select = "Select HC.*,HN.表示品番 From 発注テーブル HC "
            gSL_Select = gSL_Select & " Left Outer Join 品番マスタ HN "
            gSL_Select = gSL_Select & "   ON HN.品番 = HC.品番 "
            gSL_Select = gSL_Select & " Left Outer Join 入荷テーブル NY "
            gSL_Select = gSL_Select & "   On NY.発注管理NO = HC.発注管理NO "
            gSL_Select = gSL_Select & " Where HC.納入番号 = '" & RTrim$(iH1_Nouno) & "'"
            gSL_Select = gSL_Select & " And   HC.納品書頁 =  " & CisFun.Val2(iH1_Page)
            gSL_Select = gSL_Select & " And   HC.納入日   = '" & RTrim$(iH1_NYmd) & "'"
            gSL_Select = gSL_Select & " And   Isnull(NY.入荷日,'') = ''"
            gSL_Select = gSL_Select & " Order By HC.品番"
            Call HCTRead(gSL_Select)
            If CisDB.RecordCount = 1 Then
               iH1_Nouno.Tag = HCT.納品書管理NO
               Call HCTClose
            Else
                CXH0215.Show vbModal
                Unload CXH0215
                Set CXH0215 = Nothing
                If RV_Rtn Then
                   iH1_Nouno.Tag = Rvc_NKNo
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
        End If
    End If
'( データ確認 )
    gSL_Select = mSL_Select

    gSL_Select = gSL_Select & " Where HC.納品書管理NO = " & CisFun.Val2(iH1_Nouno.Tag)
    gSL_Select = gSL_Select & "   And HC.納品書頁 = " & CisFun.Val2(iH1_Page)
    
    If Not DBInput("INV") Then Exit Function
    
    Head1Chk = True
End Function
'****************************
'*      HEAD2ﾁｪｯｸ処理
'****************************
Private Function Head2Chk() As Boolean
    Head2Chk = False
    
'( バーコード )
    If CisFun.Val2(iH2_Bar) = 0 Then
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
    
    gSL_Select = gSL_Select & " Where HC.納品書管理NO = " & CisFun.Val2(Mid(iH2_Bar, 1, 10))
    gSL_Select = gSL_Select & "   And HC.納品書頁 = " & CisFun.Val2(Mid(iH2_Bar, 11, 2))
    
    If Not DBInput("INV") Then Exit Function
    
    Head2Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    If iB1_SYmd = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    入荷日を入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iB1_SYmd.SetFocus
        Exit Function
    End If
    If Not iB1_SYmd.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    入荷日を正しい日付にて入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iB1_SYmd.SetFocus
        Exit Function
    End If
    If iB1_SBin.Visible And Trim(B1lb_HKbn.Tag) = "1" Then
        If CisFun.Val2(iB1_SBin) = 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    入荷便を入力して下さい                 "
                .MB_Button = Error
                .MBOX
            End With
            iB1_SBin.SetFocus
            Exit Function
        End If
        If B1lb_NYmd.Tag & B1lb_NBin > iB1_SYmd & iB1_SBin Then

'---   ---   ---   ---   ---   ---   ---   < 2007/07/02 MNT >
'            With CisFun
'                .MB_Lines = 4
'                .MB_MSG(2) = "    入荷日・便は納入日・便以降にて入力して下さい         "
'                .MB_Button = Error
'                .MBOX
'            End With
'            iB1_SYmd.SetFocus
'            Exit Function

            With CisFun
                .MB_Lines = 4
                .MB_MSG(1) = "　入荷日・便＞納入日・便となります　　"
                .MB_MSG(3) = "　　　よろしいですか？"
                .MB_Button = Yes_No
                If Not .MBOX Then
                    iB1_SYmd.SetFocus
                    Exit Function
                End If
            End With
'---   ---   ---   ---   ---   ---   ---   < 2007/07/02 MNT END >

        End If
    Else
        If B1lb_NYmd.Tag > iB1_SYmd Then

'---   ---   ---   ---   ---   ---   ---   < 2007/07/02 MNT >
'            With CisFun
'                .MB_Lines = 4
'                .MB_MSG(2) = "    入荷日は納入日以降にて入力して下さい         "
'                .MB_Button = Error
'                .MBOX
'            End With
'            iB1_SYmd.SetFocus
'            Exit Function

            With CisFun
                .MB_Lines = 4
                .MB_MSG(1) = "　入荷日＞納入日となります　　"
                .MB_MSG(3) = "　　　よろしいですか？"
                .MB_Button = Yes_No
                If Not .MBOX Then
                    iB1_SYmd.SetFocus
                    Exit Function
                End If
            End With
'---   ---   ---   ---   ---   ---   ---   < 2007/07/02 MNT END >

        End If
    End If
    If gSYSOp_CLKb = 1 Then
        gInt = Kdou_Check(iB1_SYmd, 1, HCT.仕入先, HCT.受入)                         '2007/07/13MNT
    Else
        gInt = Kdou_Check(iB1_SYmd)
    End If
    If gInt = 1 Then
    '休日
        CisFun.MB_Lines = 4
        CisFun.MB_MSG(1) = "     入荷日が非稼働日です。     "
        CisFun.MB_MSG(3) = "     　よろしいですか？      "
        CisFun.MB_Button = OK_CAN
        If Not CisFun.MBOX Then
            iB1_SYmd.SetFocus
            Exit Function
        End If
    End If
    If gInt = 2 Then
    'カレンダ未登録
        CisFun.MB_Lines = 4
        CisFun.MB_MSG(1) = "     " & RTrim(iB1_SYmd.cYear) & "年" & RTrim(iB1_SYmd.cMonth) & "月のカレンダ未登録！     "
        CisFun.MB_MSG(3) = "     　　(　カレンダマスタ　)　"
        CisFun.MB_Button = Error
        CisFun.MBOX
        iB1_SYmd.SetFocus
        Exit Function
    End If
    
    Set gObj = Nothing
    
    For gCnt1 = 0 To 9
        
    
        If iB1_Suryo(gCnt1).Visible Then
           gLong = CisFun.Val2(VsGrid1.TextMatrix(gCnt1 + 1, 5))
           
           If Right(B1lb_Nouno, 1) = "9" Then
              If CisFun.Val2(iB1_Suryo(gCnt1)) <> 0 Then
                 CisFun.MB_Lines = 4
                 CisFun.MB_MSG(2) = "     分納回数が９回を超える分納処理はできません            "
                 CisFun.MB_Button = Error
                 CisFun.MBOX
                 iB1_Suryo(gCnt1).SetFocus
                 Exit Function
              End If
           End If
           If CisFun.Val2(iB1_NSuryo(gCnt1)) = 0 And _
              CisFun.Val2(iB1_Suryo(gCnt1)) = 0 Then
              gStr = "入荷数又は分割納入数を入力して下さい"
              Set gObj = iB1_NSuryo(gCnt1)
              Exit For
           End If
           If CisFun.Val2(iB1_NSuryo(gCnt1)) > gLong And _
              Trim(VsGrid1.TextMatrix(gCnt1 + 1, 6)) = "" Then
              gStr = "入荷数は発注数以下にて指定して下さい"
              Set gObj = iB1_NSuryo(gCnt1)
              Exit For
           End If
           If CisFun.Val2(iB1_Suryo(gCnt1)) > gLong And _
              Trim(VsGrid1.TextMatrix(gCnt1 + 1, 6)) = "" Then
              gStr = "分割納入数は発注数以下にて指定して下さい"
              Set gObj = iB1_Suryo(gCnt1)
              Exit For
           End If
'-----------------------------------------------------------------------------▼2005.06.17 Add▼------
           If CisFun.Val2(iB1_Suryo(gCnt1)) + CisFun.Val2(iB1_NSuryo(gCnt1)) > gLong And _
              Trim(VsGrid1.TextMatrix(gCnt1 + 1, 6)) <> "" Then
              gStr = "入荷数と分割納入数の合計は発注数以下にて指定して下さい"
              Set gObj = iB1_Suryo(gCnt1)
              Exit For
           End If
'-----------------------------------------------------------------------------▲2005.06.17 Add▲------
           If Trim(iB1_BYmd(gCnt1)) <> "" And _
              CisFun.Val2(iB1_Suryo(gCnt1)) = 0 Then
               gStr = "完納予定日入力時、分割納入数を入力して下さい"
               Set gObj = iB1_Suryo(gCnt1)
               Exit For
            End If
            
           If CisFun.Val2(iB1_Suryo(gCnt1)) <> 0 Then
              '+--------------------------------------------------- 2005/03/04 Insert
              ' 挽回日を自動算出
              If GetBYMD(True, CisFun.Val2(gCnt1)) Then
                
                 If Trim(iB1_BYmd(gCnt1) & iB1_BBin(gCnt1)) <> "" And _
                    iB1_BYmd(gCnt1) & iB1_BBin(gCnt1) <> mBYMD(gCnt1) & mBBin(gCnt1) Then
                    With CisFun
                         .MB_Lines = 4
                         If iB1_BBin(gCnt1).Visible Then
                            .MB_MSG(1) = "       " & StrConv(gCnt1 + 1, vbWide) & "行目 完納予定日・便が自動計算値と異なります。            "
                         Else
                            .MB_MSG(1) = "       " & StrConv(gCnt1 + 1, vbWide) & "行目 完納予定日が自動計算値と異なります。            "
                         End If
                         .MB_MSG(3) = "             　   よろしいですか？          "
                         .MB_Button = OK_CAN
                         If Not .MBOX Then
                           iB1_BYmd(gCnt1).SetFocus
                           Exit Function
                        End If
                    End With
                 Else
                    iB1_BYmd(gCnt1) = mBYMD(gCnt1)
                    If iB1_BBin(gCnt1).Visible Then
                       iB1_BBin(gCnt1) = mBBin(gCnt1)
                    End If
                 End If
              Else
                 If mAutoInput Then
                    If iB1_SBin.Visible And Trim(B1lb_HKbn.Tag) = "1" Then  ' かんばんはエラー
                       Exit Function
                    End If
                 End If
              End If
              
              '+--------------------------------------------------- 2005/03/04 Insert End
              If Trim(iB1_BYmd(gCnt1)) = "" Then
                 gStr = "分割納入数入力時、完納予定日を入力して下さい"
                 Set gObj = iB1_BYmd(gCnt1)
                 Exit For
              End If
              If Not iB1_BYmd(gCnt1).cChkResult Then
                 gStr = "完納予定日を正しい日付にて入力して下さい"
                 Set gObj = iB1_BYmd(gCnt1)
                 Exit For
              End If
              
                If gSYSOp_CLKb = 1 Then
                    gInt = Kdou_Check(iB1_BYmd(gCnt1), 1, HCT.仕入先, HCT.受入)     '2007/07/13MNT
                Else
                    gInt = Kdou_Check(iB1_BYmd(gCnt1))
                End If
              If gInt = 1 Then
                '休日
                 CisFun.MB_Lines = 4
                 CisFun.MB_MSG(1) = "       " & StrConv(gCnt1 + 1, vbWide) & "行目 完納予定日が非稼働日です。     "
                 CisFun.MB_MSG(3) = "     　   よろしいですか？          "
                 CisFun.MB_Button = OK_CAN
                 If Not CisFun.MBOX Then
                    iB1_BYmd(gCnt1).SetFocus
                    Exit Function
                 End If
              End If
              If gInt = 2 Then
                'カレンダ未登録
                 CisFun.MB_Lines = 4
                 CisFun.MB_MSG(1) = "     " & RTrim(iB1_BYmd(gCnt1).cYear) & "年" & RTrim(iB1_BYmd(gCnt1).cMonth) & "月のカレンダ未登録！     "
                 CisFun.MB_MSG(3) = "     　　(　カレンダマスタ　)　"
                 CisFun.MB_Button = Error
                 CisFun.MBOX
                 iB1_BYmd(gCnt1).SetFocus
                 Exit Function
              End If
              If iB1_BBin(gCnt1).Visible Then
                 If Trim(B1lb_HKbn.Tag) = "1" And CisFun.Val2(iB1_BBin(gCnt1)) = 0 Then
                    gStr = "分割納入数入力時、完納予定便を入力して下さい"
                    Set gObj = iB1_BBin(gCnt1)
                    Exit For
                 End If
                 If iB1_SYmd & iB1_SBin > iB1_BYmd(gCnt1) & iB1_BBin(gCnt1) Then
                    gStr = "完納予定日・便は入荷日・便以降にて入力して下さい"
                    Set gObj = iB1_BYmd(gCnt1)
                    Exit For
                 End If
              Else
                 If iB1_SYmd > iB1_BYmd(gCnt1) Then
                    gStr = "完納予定日は入荷日以降にて入力して下さい"
                    Set gObj = iB1_BYmd(gCnt1)
                    Exit For
                 End If
              End If
           End If
        End If
    Next gCnt1
    
    If Not gObj Is Nothing Then
        With CisFun
            .MB_Lines = 3
            .MB_MSG(2) = "          " & StrConv(gCnt1 + 1, vbWide) & "行目 " & gStr & "                "
            .MB_Button = Error
            .MBOX
        End With
        gObj.SetFocus
        Exit Function
    End If
    Body1Chk = True
End Function
'+---------------------------------------+
'+  一括からの連携処理
'+---------------------------------------+
Private Function ContinuChk() As Boolean
    ContinuChk = False
    If mBunCount > Rv_NTtlSu Then Exit Function
    Do Until Rv_HNo(mBunCount) <> 0
       If mBunCount >= Rv_NTtlSu Then
          Call PB_END_Click
          Exit Function
       End If
       mBunCount = mBunCount + 1
    Loop
    
    Bun_Count = CisFun.RSetFld(mBunCount, 0, "###") & " / " & _
                CisFun.RSetFld(Rv_NTtlSu, 0, "###")
    gSL_Select = "Select 納品書管理NO,納入番号,納入日,納品書頁 "
    gSL_Select = gSL_Select & " From 発注テーブル "
    gSL_Select = gSL_Select & " Where 納品書管理NO = " & Rv_HNo(mBunCount)
    gSL_Select = gSL_Select & "   And 納品書頁 = " & Rv_HPage(mBunCount)
    If Not HCTRead(gSL_Select, 1) Then
       Call PB_END_Click
    End If
    
    iH1_NYmd = HCT.納入日
    iH1_Nouno.Tag = HCT.納品書管理NO
    iH1_Nouno = HCT.納入番号
    iH1_Page = CisFun.RSetFld(HCT.納品書頁, 0, iH1_Page.cFormat)
    ProcHB = "H1"
    Call PB_ENT_Click
    
    ContinuChk = True

End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    
    Call HeadBodyClear("B")
    
    On Error GoTo DBInput_Err:
    If Not HCTRead(gSL_Select, 0) Then
        If InputKey = "Next" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    最終レコードです            "
                .MB_Title = "【次頁】"
                .MB_Button = OK
                .MBOX
            End With
            GoTo DBInput_Ed
        End If
        If InputKey = "Back" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    先頭レコードです            "
                .MB_Title = "【前頁】"
                .MB_Button = OK
                .MBOX
            End With
        End If
        If InputKey = "INV" Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "  該当データがありません  "
                .MB_Button = Error
                .MBOX
            End With
        End If
        If InputKey = "QR" Then
           QRDataArea.エラー内容 = "該当データなし"
        End If
        GoTo DBInput_Ed
    Else
        Call BodySet(InputKey)
        
        If HCT.入荷日 <> "" Then
            If InputKey = "QR" Then
               QRDataArea.エラー内容 = "入荷済みデータ。"
               Exit Function
            End If
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "      入荷済みデータです              "
                .MB_Button = Error
                .MBOX
            End With
            Exit Function
        End If
        
        DBInput = True
    End If
    
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Sub BodySet(SetType As String)
'
    With HCT
        iB1_SYmd.Tag = HCT.納品書管理NO
        iB1_SBin.Tag = HCT.納品書頁
        
        B1lb_Nouno = HCT.納入番号
        B1lb_Page = HCT.納品書頁
    ' 入荷日･初期表示
        If SetType = "INV" Or SetType = "QR" Then
            iB1_SYmd = .納入日
            If Rv_NCall = "K" And Trim(Rv_NYmd) <> "" Then
               iB1_SYmd = Rv_NYmd
            End If
            If Trim(HCT.納入便) <> "" Then
               iB1_SBin.Visible = True
               iB1_SBin = .納入便
               If Rv_NCall = "K" And Trim(Rv_NBin) <> "" Then
                  iB1_SBin = Rv_NBin
               End If
            Else
               iB1_SBin.Visible = False
            End If
        End If
        
        If Trim(.仕入先名) = "" Then
           gStr = .仕入先
        Else
           gStr = RTrim(.仕入先) & ":" & .仕入先名
        End If
        B1lb_Tornm = gStr
        
        B1lb_NYmd.Tag = HCT.納入日
        With CisFun
            .DateE_BefVal = HCT.納入日
            .DateEdit
            B1lb_NYmd = CisFun.DateE_AfVal
            If Trim(HCT.納入便) <> "" Then
                B1lb_NBin = HCT.納入便
            End If
        End With
        B1lb_HKbn = .発注区分名
        B1lb_HKbn.Tag = .発注区分
        B1lb_TKbn = .手配区分名
    End With
    
    Call GridSet(SetType)
    
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'*-------------------------------------------------*
'*      デ ー タ 更 新  確 認 メ ッ セ ー ジ
'*-------------------------------------------------*
Private Function CheckMessage() As Boolean
    CheckMessage = False
    
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     入荷実績処理を行います      "
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Title = "実績更新処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    CheckMessage = True
    
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Static AddFlg               As Boolean
    
'   ( 確認メッセージ )
    If Not CheckMessage Then Exit Function
    
    gSL_Select = "Select * From 発注テーブル"
    gSL_Select = gSL_Select & " Where 納品書管理NO = " & CisFun.Val2(iB1_SYmd.Tag)
    gSL_Select = gSL_Select & "   And 納品書頁 = " & CisFun.Val2(iB1_SBin.Tag)
    
    If Not HCTRead(gSL_Select, , 1) Then
       Call HCTClose
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     発注データ異常      "
            .MB_Title = "更新エラー"
            .MB_Button = OK
            .MBOX
            Exit Function
       End With
    End If
    
    CisDB.ConnectNo = 0
    Call CisDB.DBTran(TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
    
    Do Until Not HCT_RDSTS
        
        AddFlg = False
        
        gSL_Select = "Select * From 入荷テーブル "
        gSL_Select = gSL_Select & " Where 発注管理NO = " & HCT.発注管理NO
        If Not NYTRead(gSL_Select, 1) Then
           Call ItemsClearNYT
           NYT.発注管理NO = HCT.発注管理NO
           AddFlg = True
        End If
        
        With NYT
                        
            .納入番号 = HCT.納入番号
            .納品書頁 = HCT.納品書頁
            .納品書行 = HCT.納品書行
            .仕入先 = HCT.仕入先
            .受入 = HCT.受入
            .品番 = HCT.品番
            .背番号 = HCT.背番号
            .収容数 = HCT.収容数
            
            .入荷日 = iB1_SYmd
            If iB1_SBin.Visible Then
                .入荷便 = Format(iB1_SBin, "00")
            Else
                .入荷便 = ""
            End If
            gLong = HCT.納品書行 - 1
            .入荷端数 = 0
            .分割数 = 0
            .挽回日 = ""
            .挽回便 = ""
            If CisFun.Val2(iB1_Suryo(gLong)) = 0 Then
               .入荷数 = CisFun.Val2(iB1_NSuryo(gLong))
            Else
               .入荷数 = CisFun.Val2(iB1_NSuryo(gLong))
               .分割数 = CisFun.Val2(iB1_Suryo(gLong))
               .挽回日 = iB1_BYmd(gLong)
               If iB1_BBin(gLong).Visible Then
                  .挽回便 = Format(iB1_BBin(gLong), "00")
               End If
            End If
            gCnt1 = 0
            If HCT.収容数 <> 0 Then
               .入荷枚数 = NYT.入荷数 \ HCT.収容数
               .入荷端数 = NYT.入荷数 Mod HCT.収容数
               If .入荷端数 <> 0 Then
                  .入荷枚数 = .入荷枚数 + 1
               End If
            End If
        End With
    
        If AddFlg Then
           NYT.作成者 = gTanto
           Call NYTInsert
        Else
           NYT.更新者 = gTanto
           Call NYTUpdate
        End If
        
        If CisFun.Val2(iB1_Suryo(gLong)) <> 0 Then
           If Not HPrtUpdNyuka Then GoTo DBPut_Err
        End If
        ' 入荷実績処理ワーク 出力
        Call ItemsClearNJW
        With NJW
            .処理端末 = CisFun.WSID
            .発注管理NO = HCT.発注管理NO
        End With
        Call NJWInsert
        
        Call HCTReadNext(1)
    Loop
    
    Call HCTClose(1)
    
    CisDB.ConnectNo = 0
    Call CisDB.DBTran(TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    lb_JMai = Format(CisFun.Val2(lb_JMai) + 1, "#,###")
    
    If Rv_NCall = "K" Then
       Rv_NExec(mBunCount) = True
    End If
    
    
    On Error GoTo 0
    
    DBPut = True
    Exit Function
    
DBPut_Err:
    Call HCTClose(1)
    
    CisDB.ConnectNo = 0
    Call CisDB.DBTran(TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ破棄
    CisFun.ErrorBox
    End
End Function
'+-----------------------------------+
'+      入荷情報　受領書 変更（分納分)
'+-----------------------------------+
Private Function HPrtUpdNyuka() As Boolean
    HPrtUpdNyuka = False
    
    With HCT
        gSL_Select = "Update 発注帳票発行テーブル Set "
        gSL_Select = gSL_Select & " 納入日 = '" & RTrim(NYT.入荷日) & "',"
        gSL_Select = gSL_Select & " 納入便 = '" & RTrim(NYT.入荷便) & "',"
        gSL_Select = gSL_Select & "箱数" & .納品書行 & " = " & NYT.入荷枚数 & ","
        gSL_Select = gSL_Select & "端数" & .納品書行 & " = " & NYT.入荷端数 & ","
        gSL_Select = gSL_Select & "納入数" & .納品書行 & " = " & NYT.入荷数 & ""
        
        gSL_Select = gSL_Select & " Where 納品書管理NO = " & .納品書管理NO
        gSL_Select = gSL_Select & "   And 頁NO = " & .納品書頁
        gSL_Select = gSL_Select & "   And カード = '6'"
        With CisDB
            .SQL = gSL_Select
            If Not .DBExec Then Exit Function
        End With
        
        gSL_Select = "Update 発注帳票発行テーブル Set "
        gSL_Select = gSL_Select & "箱総数 = "
        For gInt = 1 To 10
            gSL_Select = gSL_Select & "Isnull(箱数" & gInt & ",0)"
            If gInt <> 10 Then
               gSL_Select = gSL_Select & " + "
            End If
        Next gInt
        gSL_Select = gSL_Select & " Where 納品書管理NO = " & .納品書管理NO
        gSL_Select = gSL_Select & "   And 頁NO = " & .納品書頁
        gSL_Select = gSL_Select & "   And カード = '6'"
        With CisDB
            .SQL = gSL_Select
            If Not .DBExec Then Exit Function
        End With
    End With

    HPrtUpdNyuka = True

End Function
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 0, 0, "№")
        Call .FixedSet(中中, 左中, あり, 0, 1, "品　番")
        Call .FixedSet(中中, 左中, あり, 0, 2, "背番号")
        Call .FixedSet(中中, 右中, あり, 0, 3, "収容数")
        Call .FixedSet(中中, 右中, あり, 0, 4, "箱 数")
        Call .FixedSet(中中, 右中, あり, 0, 5, "発注数")
        Call .FixedSet(中中, 中中, あり, 4, 6, "実数")
        Call .FixedSet(中中, 左中, あり, 0, 7, "完納予定日")
        Call .FixedSet(中中, 左中, あり, 0, 8, "予定便")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
'        Call .FixedSet(中中, 左中, なし, 0,9,"Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet(SetType As String)
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
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
            If HCT.納品書行 > 10 Then Exit Do
            'No
            .TextMatrix(HCT.納品書行, 0) = HCT.納品書行
            '品番
            .TextMatrix(HCT.納品書行, 1) = HCT.表示品番
            ' 背番号
            .TextMatrix(HCT.納品書行, 2) = " " & HCT.背番号
            ' 収容数
            .TextMatrix(HCT.納品書行, 3) = Format(HCT.収容数, "#,###")
            ' 枚数
            .TextMatrix(HCT.納品書行, 4) = Format(HCT.枚数, "#,###")
            ' 発注数
            .TextMatrix(HCT.納品書行, 5) = Format(HCT.発注数, "#,###")
            ' 実数区分
            If HCT.実数区分 = 1 Then
               .TextMatrix(HCT.納品書行, 6) = "○"
            End If
            
            If SetType = "INV" Or SetType = "QR" Then
               iB1_NSuryo(HCT.納品書行 - 1) = CisFun.RSetFld(HCT.発注数, 0, iB1_NSuryo(0).cFormat)
               iB1_NSuryo(HCT.納品書行 - 1).Tag = HCT.発注数
               iB1_Suryo(HCT.納品書行 - 1).Tag = HCT.実数区分
            End If
                        
            If Trim(HCT.入荷日) <> "" Then
               iB1_NSuryo(HCT.納品書行 - 1) = CisFun.RSetFld(HCT.入荷数, 0, iB1_NSuryo(0).cFormat)
               ' 入荷日･便
               iB1_SYmd = HCT.入荷日
               If Trim(HCT.入荷便) <> "" Then
                  iB1_SBin = HCT.入荷便
               End If
               ' 分割納入数
               iB1_Suryo(HCT.納品書行 - 1) = CisFun.RSetFld(HCT.分割数, 0, iB1_Suryo(ID - 1).cFormat)
               ' 完納予定日･便
               iB1_BYmd(HCT.納品書行 - 1) = HCT.挽回日
               If Trim(HCT.挽回便) <> "" Then
                   iB1_BBin(HCT.納品書行 - 1) = HCT.挽回便
               End If
            End If
            
            If Trim(HCT.納入便) <> "" Then
               iB1_BBin(HCT.納品書行 - 1).Visible = True
            Else
               iB1_BBin(HCT.納品書行 - 1).Visible = False
            End If
        End With
        Call HCTReadNext
    Loop
    Call HCTClose
    With VsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    
    For Each gObj In iB1_Suryo
        gStr = Trim(VsGrid1.TextMatrix(gObj.Index + 1, 1))
        If gStr = "" Then
           gObj.Visible = False
           iB1_NSuryo(gObj.Index).Visible = False
           iB1_BYmd(gObj.Index).Visible = False
           iB1_BBin(gObj.Index).Visible = False
        Else
           gObj.Visible = True
           iB1_NSuryo(gObj.Index).Visible = True
           iB1_BYmd(gObj.Index).Visible = True
           If iB1_SBin.Visible Then
              iB1_BBin(gObj.Index).Visible = True
           Else
              iB1_BBin(gObj.Index).Visible = False
           End If
        End If
    Next gObj
    
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
    
    gStr = QR_CommEvent
    If gStr = "ER" Then GoTo PDQComm1_OnComm_ED
    If gStr <> "OK" Then Exit Sub
    
    With QRDataArea
'+----------------------------------------------------------- 2005/02/23 Insert
        If .ステータス <> 0 Then
           Call QRReadMsg("E")
           Exit Sub
        End If
        If .バーコード種類 <> "B" And .バーコード種類 <> "J" Then
            ' 納品書のみ有効
            If .カード <> "7" And .カード <> "0" Then
               .エラー内容 = "カードが対象外です"
               Call QRReadMsg("E")
               Exit Sub
            End If
        End If
        
        If .バーコード種類 <> "J" Then
            ' 部品のみ有効
            If .部品区分 <> "D" Then
               .エラー内容 = "部品の納品書でない"
               Call QRReadMsg("E")
               Exit Sub
            End If
        '( データ確認 )
            gSL_Select = mSL_Select
            
            gSL_Select = gSL_Select & " Where HC.納品書管理NO = " & QRDataArea.納品書管理NO
            gSL_Select = gSL_Select & "   And HC.納品書頁 = " & QRDataArea.納品書頁
            
            If Not DBInput("QR") Then
               Call QRReadMsg("E")
               Exit Sub
            End If
            
            iH1_NYmd = HCT.納入日
            iH1_Nouno = HCT.納入番号
            iH1_Page = HCT.納品書頁
        Else
            ' 支給品のみ有効
            
'------------------------------------------------------------------ 2005/08/31 Delete
'            If typQRAfter.カード丼 <> "1" And typQRAfter.カード丼 <> "5" Then
'------------------------------------------------------------------ 2005/08/31 Delete End
'------------------------------------------------------------------ 2005/08/31 Insert
            If typQRAfter.カード丼 <> "1" And typQRAfter.カード丼 <> "5" _
             And typQRAfter.カード丼 <> "7" And typQRAfter.カード丼 <> "8" Then
'------------------------------------------------------------------ 2005/08/31 Insert End
               .エラー内容 = "カードが対象外です"
               Call QRReadMsg("E")
               Exit Sub
            End If
            QRr_ReadData = ""
        
            iH1_NYmd = typQRAfter.納入日
            iH1_Nouno = typQRAfter.納入番号
            iH1_Page = typQRAfter.ページNO
            ProcHB = "H1"
            If Not Head1Chk Then
               ProcHB = "H3"
               Exit Sub
            End If
        End If
    End With
'+----------------------------------------------------------- 2005/02/23 Insert End
'+----------------------------------------------------------- 2005/02/23 Delete
'        If .バーコード種類 <> "B" Then
'            If .ステータス <> 0 Then GoTo PDQComm1_OnComm_ED
'            ' 部品のみ有効
'            If .バーコード種類 <> "D" Then
'               .エラー内容 = "部品の納品書でない"
'               Call QRReadMsg("E")
'               GoTo PDQComm1_OnComm_ED
'            End If
'            ' 納品書のみ有効
'            If .カード <> "7" Then
'               .エラー内容 = "カードが対象外です"
'               Call QRReadMsg("E")
'               GoTo PDQComm1_OnComm_ED
'            End If
'        End If
'    End With
''( データ確認 )
'    gSL_Select = mSL_Select
'
'    gSL_Select = gSL_Select & " Where HC.納品書管理NO = " & QRDataArea.納品書管理NO
'    gSL_Select = gSL_Select & "   And HC.納品書頁 = " & QRDataArea.納品書頁
'
'    If Not DBInput("QR") Then
'       Call QRReadMsg("E")
'       Exit Sub
'    End If
'+----------------------------------------------------------- 2005/02/23 Delete End
    
    ProcHB = "B1"
    Call DispChange(ProcHB)
    Exit Sub
PDQComm1_OnComm_ED:
    Call HeadBodyClear("B")
    Call QRReadMsg("E")
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
'       Call TX2400OpenAuto
       Exit Sub
    End If
    If TX2400_STSNo = 2 Then
'       Call TX2400ReadAuto
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
            If .バーコード種類 <> "D" Then
               .エラー内容 = "部品の納品書でない"
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
    
    If Not DBInput("QR") Then
       TX2400_STSNo = 5
       GoTo PDQComm2_OnComm_ED
    End If
    
    Call TX2400Stack(True)
    
    ProcHB = "B1"
    Call DispChange(ProcHB)
    
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
'+--------------------------------------------------------------------------- 2005/03/04 Insert
'+----------------------------------------------------------------------+
'+      挽回日自動計算
'+              IndexFlg : True/指定行のみ行う False:すべての行を行う
'+              Index    : IndexFlg=Trueの時の指定行
'+----------------------------------------------------------------------+
Private Function GetBYMD(IndexFlg As Boolean, Index As Integer) As Boolean
    Dim wSt As Integer
    Dim wEd As Integer
    Dim wId As Integer
    
    Erase mBYMD
    Erase mBBin
    GetBYMD = False
    
    If Not mAutoInput Then Exit Function
    
    If Not iB1_SBin.Visible Then Exit Function
    
    wSt = 0: wEd = 9
    If IndexFlg Then
       wSt = Index: wEd = Index
    End If
    
    For wId = wSt To wEd
            
        With VsGrid1
             If Trim(.TextMatrix(wId + 1, 1)) <> "" Then
                gSL_Select = "Select * From 発注テーブル "
                gSL_Select = gSL_Select & " Where 納品書管理NO = " & CisFun.Val2(iB1_SYmd.Tag)
                gSL_Select = gSL_Select & "   And 納品書頁 = " & CisFun.Val2(iB1_SBin.Tag)
                gSL_Select = gSL_Select & "   And 納品書行 = " & wId + 1
                If HCTRead(gSL_Select, 1, 1) Then
                    If Not BankaiYmd(wId) Then Exit Function
                End If
             End If
        End With
    Next wId
    GetBYMD = True
End Function
Private Function BankaiYmd(Index As Integer) As Boolean
    Dim wErr1 As Long
    Dim wErr2 As Long
    
    BankaiYmd = False
    
    On Error GoTo BankaiYmd_Er:
    
    ' 指示
    If HCT.SK区分 = 1 Then Exit Function
''       gSL_Select = "Select 発注リードタイム From 品番取引先マスタ "
''       gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(HCT.品番) & "'"
''       gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(HCT.仕入先) & "'"
''       gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(HCT.受入) & "'"
''       If Not HTMRead(gSL_Select, 1, 1) Then
''          HTM.発注リードタイム = 1
''       End If
''       If HTM.発注リードタイム = 0 Then
''          HTM.発注リードタイム = 1
''       End If
''
''       On Error GoTo BankaiYmd_Er
''       gStr = ""
''       With CisDB
''            .SQL = "日付加算"
''            .StoadoCount = 4
''            .ParaNo = 0: .ParaIO = Return用
''            .ParaNo = 1: .ParaIO = Input用      ' 対象日付
''            .ParaNo = 2: .ParaIO = Input用      ' 加算日数
''            .ParaNo = 3: .ParaIO = OutPut用     ' 結果日付
''            .ParaNo = 4: .ParaIO = OutPut用     ' ﾘﾀｰﾝｺｰﾄﾞ
''
''            .ParaNo = 1: .ParaValue = iB1_SYmd
''            .ParaNo = 2: .ParaValue = HTM.発注リードタイム
''
''            If Not .DBStored Then GoTo BankaiYmd_Er
''
''            ' ストアド処理エラー
''            .ParaNo = 0: wErr1 = CisFun.Val2(.ParaValue)
''            If wErr1 <> 0 Then GoTo BankaiYmd_Er
''
''            ' エラー番号
''            .ParaNo = 4: wErr2 = CisFun.Val2(.ParaValue)
''            If wErr2 <> 0 Then GoTo BankaiYmd_Er
''            ' 結果日付
''            .ParaNo = 3: mBYMD(Index) = Trim(.ParaValue)
''            If iB1_BBin(Index).Visible Then mBBin(Index) = "01"
''       End With
''       Exit Function
''    End If
    ' かんばん
    If Not Cycle_Check(Index) Then
        With CisFun
            .MB_Button = Error
            .MBOX
            Exit Function
        End With
    End If
    BankaiYmd = True
    Exit Function
BankaiYmd_Er:
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "    " & Index + 1 & "行目  完納予定日算出エラー              "
        .MB_Button = Error
        .MBOX
    End With
End Function
'*====================================================*
'*      発注サイクルより納入日算出
'*====================================================*
Private Function Cycle_Check(Index As Integer) As Boolean
    With CisFun
        .MB_Lines = 5
        .MB_Button = Error
    End With
    Cycle_Check = False
'発注サイクルマスタの存在チェック
    gSL_Select = "SELECT * FROM 発注サイクルマスタ"
    gSL_Select = gSL_Select & " WHERE 仕入先 = '" & RTrim(HCT.仕入先) & "'"
    gSL_Select = gSL_Select & "   AND (受入 = '" & RTrim(HCT.受入) & "'"
    gSL_Select = gSL_Select & "   OR  受入 = '')"
    gSL_Select = gSL_Select & "   AND サイクル = '" & RTrim(HCT.サイクル) & "'"
    gSL_Select = gSL_Select & "   AND 振出便 = '" & RTrim(iB1_SBin) & "'"
    If Not HCMRead(gSL_Select, 1) Then
        CisFun.MB_MSG(1) = "  　  発注サイクル　未登録！            "
        CisFun.MB_MSG(3) = "    （　発注サイクルマスタ　）            "
        Exit Function
    End If
'納入日・便を求める
    With CisDB
        .SQL = "便算出"
        .StoadoCount = 15
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = Input用
        .ParaNo = 3: .ParaIO = Input用
        .ParaNo = 4: .ParaIO = Input用
        .ParaNo = 5: .ParaIO = Input用
        .ParaNo = 6: .ParaIO = Input用
        .ParaNo = 7: .ParaIO = Input用
        .ParaNo = 8: .ParaIO = Input用
        .ParaNo = 9: .ParaIO = OutPut用             'リターン
        .ParaNo = 10: .ParaIO = OutPut用             '納入日
        .ParaNo = 11: .ParaIO = OutPut用            '納入便
        .ParaNo = 12: .ParaIO = OutPut用            '発注日
        .ParaNo = 13: .ParaIO = OutPut用            '発注時間
        .ParaNo = 14: .ParaIO = OutPut用            '発行日
        .ParaNo = 15: .ParaIO = OutPut用            '発行時間
        
        .ParaNo = 1: .ParaValue = "+"
        .ParaNo = 2: .ParaValue = RTrim(HCT.仕入先)
        .ParaNo = 3: .ParaValue = RTrim(HCT.サイクル)
        .ParaNo = 4: .ParaValue = RTrim(HCT.受入)
        .ParaNo = 5: .ParaValue = RTrim(iB1_SYmd)
        .ParaNo = 6: .ParaValue = RTrim(iB1_SBin)
        .ParaNo = 7: .ParaValue = 0
        .ParaNo = 8: .ParaValue = 0
'
        If Not .DBStored Then
            MsgBox "ストアドエラー "
            Exit Function
        Else
            .ParaNo = 9
            If .ParaValue <> 0 Then GoTo Cycle_Check_Err
            .ParaNo = 10: mBYMD(Index) = .ParaValue
            .ParaNo = 11: If iB1_BBin(Index).Visible Then mBBin(Index) = .ParaValue
        End If
    End With
'
    Cycle_Check = True
    Exit Function
Cycle_Check_Err:
    Select Case CisDB.ParaValue
        Case 1: CisFun.MB_MSG(2) = "    " & Index + 1 & "行目【　カレンダマスタ　未登録！】            "
        Case 2: CisFun.MB_MSG(2) = "    " & Index + 1 & "行目【　発注サイクルマスタ　未登録！】            "
        Case 3: CisFun.MB_MSG(2) = "    " & Index + 1 & "行目【　翌月カレンダマスタ　未登録！】            "
        Case 4: CisFun.MB_MSG(2) = "    " & Index + 1 & "行目【　前月カレンダマスタ　未登録！】            "
        Case 5: CisFun.MB_MSG(2) = "    " & Index + 1 & "行目【　発注サイクルマスタ　未登録！】            "
    End Select
End Function

'+--------------------------------------------------------------------------- 2005/03/04 Insert End





