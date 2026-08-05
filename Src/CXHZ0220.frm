VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Begin VB.Form CXHZ0220 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "çﬁóøÅEì¸â◊é¿ê—ì¸óÕ [å¬ï ]"
   ClientHeight    =   11085
   ClientLeft      =   270
   ClientTop       =   1755
   ClientWidth     =   15960
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
   PaletteMode     =   1  'Z µ∞¿ﬁ∞
   ScaleHeight     =   11085
   ScaleWidth      =   15960
   WindowState     =   2  'ç≈ëÂâª
   Begin Cis3D_v60.CIS3D H4_Area1 
      Height          =   1320
      Left            =   2190
      Top             =   7185
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2328
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         TabIndex        =   0
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
         Caption         =   "QRŒØ ﬂÿ∞¿ﬁì«éÊíÜ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr Çoñæí©"
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
      Left            =   2430
      Top             =   5925
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   2302
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         TabIndex        =   1
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
         Caption         =   "QRÉäÅ[É_Å[ì«éÊíÜ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr Çoñæí©"
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
      Height          =   375
      Left            =   105
      Top             =   3135
      Visible         =   0   'False
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   661
      BackColor       =   16761087
      Caption         =   "ÉGÉâÅ["
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ñæí©"
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
      TabIndex        =   65
      Top             =   1275
      Width           =   480
      _Version        =   196610
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      Left            =   135
      Top             =   1140
      Visible         =   0   'False
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   1402
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Caption         =   "ï™î[èàóùåèêî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   5040
      ScaleHeight     =   1050
      ScaleWidth      =   2025
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   570
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
         Caption         =   "ÉoÅ[ÉRÅ[Éh"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            TabIndex        =   5
            Top             =   420
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   635
            cFormat         =   "000000000000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         BackColor       =   8438015
         Caption         =   "é¿ê—ñáêî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   120
      ScaleHeight     =   1050
      ScaleWidth      =   3735
      TabIndex        =   59
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
         Caption         =   "î[ î‘"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            TabIndex        =   3
            Top             =   420
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Caption         =   "ï≈"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            TabIndex        =   4
            Top             =   420
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
            cFormat         =   "##"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Caption         =   "î[ ì¸ ì˙"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            TabIndex        =   2
            Top             =   420
            Width           =   1860
            _ExtentX        =   2990
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   9180
      ScaleHeight     =   660
      ScaleWidth      =   6045
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   795
      Width           =   6105
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   3465
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   3
         BtnCaption      =   "éÊè¡"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   2
         BtnCaption      =   "ì¸óÕ"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   1
         BtnCaption      =   "èIóπ"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   11
         BtnCaption      =   "åüçı"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   5
         BtnCaption      =   "ëOêi"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   6
         BtnCaption      =   "å„ëﬁ"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   10
         BtnCaption      =   "êÿë÷"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Height          =   7215
      Left            =   90
      ScaleHeight     =   7155
      ScaleWidth      =   15135
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   3525
      Width           =   15195
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   7050
         Left            =   60
         Top             =   45
         Width           =   15030
         _ExtentX        =   26511
         _ExtentY        =   12435
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Height          =   6015
            Left            =   8370
            Top             =   885
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   10610
            ForeColor       =   16711680
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   555
               Index           =   29
               Left            =   5670
               Top             =   5445
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Biko 
                  Height          =   360
                  Index           =   9
                  Left            =   15
                  TabIndex        =   57
                  Top             =   90
                  Width           =   870
                  _ExtentX        =   1535
                  _ExtentY        =   635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXX"
                  MaxLength       =   6
                  IMEMode         =   4
                  cDataReplace    =   1
                  cFbComma        =   0
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   28
               Left            =   5670
               Top             =   4890
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Biko 
                  Height          =   360
                  Index           =   8
                  Left            =   15
                  TabIndex        =   52
                  Top             =   90
                  Width           =   870
                  _ExtentX        =   1535
                  _ExtentY        =   635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXX"
                  MaxLength       =   6
                  IMEMode         =   4
                  cDataReplace    =   1
                  cFbComma        =   0
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   27
               Left            =   5670
               Top             =   4335
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Biko 
                  Height          =   360
                  Index           =   7
                  Left            =   15
                  TabIndex        =   47
                  Top             =   90
                  Width           =   870
                  _ExtentX        =   1535
                  _ExtentY        =   635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXX"
                  MaxLength       =   6
                  IMEMode         =   4
                  cDataReplace    =   1
                  cFbComma        =   0
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   26
               Left            =   5670
               Top             =   3780
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Biko 
                  Height          =   360
                  Index           =   6
                  Left            =   15
                  TabIndex        =   42
                  Top             =   90
                  Width           =   870
                  _ExtentX        =   1535
                  _ExtentY        =   635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXX"
                  MaxLength       =   6
                  IMEMode         =   4
                  cDataReplace    =   1
                  cFbComma        =   0
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   25
               Left            =   5670
               Top             =   3225
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Biko 
                  Height          =   360
                  Index           =   5
                  Left            =   15
                  TabIndex        =   37
                  Top             =   90
                  Width           =   870
                  _ExtentX        =   1535
                  _ExtentY        =   635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXX"
                  MaxLength       =   6
                  IMEMode         =   4
                  cDataReplace    =   1
                  cFbComma        =   0
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   24
               Left            =   5670
               Top             =   2670
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Biko 
                  Height          =   360
                  Index           =   4
                  Left            =   15
                  TabIndex        =   32
                  Top             =   90
                  Width           =   870
                  _ExtentX        =   1535
                  _ExtentY        =   635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXX"
                  MaxLength       =   6
                  IMEMode         =   4
                  cDataReplace    =   1
                  cFbComma        =   0
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   23
               Left            =   5670
               Top             =   2115
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Biko 
                  Height          =   360
                  Index           =   3
                  Left            =   15
                  TabIndex        =   27
                  Top             =   90
                  Width           =   870
                  _ExtentX        =   1535
                  _ExtentY        =   635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXX"
                  MaxLength       =   6
                  IMEMode         =   4
                  cDataReplace    =   1
                  cFbComma        =   0
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   22
               Left            =   5670
               Top             =   1560
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Biko 
                  Height          =   360
                  Index           =   2
                  Left            =   15
                  TabIndex        =   22
                  Top             =   90
                  Width           =   870
                  _ExtentX        =   1535
                  _ExtentY        =   635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXX"
                  MaxLength       =   6
                  IMEMode         =   4
                  cDataReplace    =   1
                  cFbComma        =   0
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   21
               Left            =   5670
               Top             =   1005
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Biko 
                  Height          =   360
                  Index           =   1
                  Left            =   15
                  TabIndex        =   17
                  Top             =   90
                  Width           =   870
                  _ExtentX        =   1535
                  _ExtentY        =   635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXX"
                  MaxLength       =   6
                  IMEMode         =   4
                  cDataReplace    =   1
                  cFbComma        =   0
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   20
               Left            =   5670
               Top             =   450
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Biko 
                  Height          =   360
                  Index           =   0
                  Left            =   15
                  TabIndex        =   12
                  Top             =   90
                  Width           =   870
                  _ExtentX        =   1535
                  _ExtentY        =   635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXX"
                  MaxLength       =   6
                  IMEMode         =   4
                  cDataReplace    =   1
                  cFbComma        =   0
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D15 
               Height          =   435
               Left            =   5670
               Top             =   15
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   767
               ForeColor       =   16711680
               Caption         =   "îıçl"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Begin Cis3D_v60.CIS3D CIS3D18 
               Height          =   435
               Left            =   1725
               Top             =   15
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   767
               ForeColor       =   16711680
               Caption         =   "ï™äÑî[ì¸êî"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   435
               Left            =   3435
               Top             =   15
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   767
               ForeColor       =   16711680
               Caption         =   " äÆî[ó\íËì˙  - ï÷"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   555
               Index           =   0
               Left            =   1725
               Top             =   450
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   9
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   555
               Index           =   0
               Left            =   3435
               Top             =   450
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   10
                  Top             =   90
                  Width           =   1770
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                     Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
                  Left            =   1800
                  TabIndex        =   11
                  Top             =   90
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   555
               Index           =   1
               Left            =   1725
               Top             =   1005
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   14
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   555
               Index           =   1
               Left            =   3435
               Top             =   1005
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   1800
                  TabIndex        =   16
                  Top             =   90
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   30
                  TabIndex        =   15
                  Top             =   90
                  Width           =   1770
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                     Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
               Height          =   555
               Index           =   2
               Left            =   1725
               Top             =   1560
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   19
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   555
               Index           =   2
               Left            =   3435
               Top             =   1560
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   1800
                  TabIndex        =   21
                  Top             =   90
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   30
                  TabIndex        =   20
                  Top             =   90
                  Width           =   1770
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                     Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
               Height          =   555
               Index           =   3
               Left            =   1725
               Top             =   2115
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   24
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   555
               Index           =   3
               Left            =   3435
               Top             =   2115
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   1800
                  TabIndex        =   26
                  Top             =   90
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   30
                  TabIndex        =   25
                  Top             =   90
                  Width           =   1770
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                     Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
               Height          =   555
               Index           =   4
               Left            =   1725
               Top             =   2670
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   29
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   555
               Index           =   4
               Left            =   3435
               Top             =   2670
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   1800
                  TabIndex        =   31
                  Top             =   90
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   30
                  TabIndex        =   30
                  Top             =   90
                  Width           =   1770
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                     Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
               Height          =   555
               Index           =   5
               Left            =   1725
               Top             =   3225
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   34
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   555
               Index           =   5
               Left            =   3435
               Top             =   3225
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   1800
                  TabIndex        =   36
                  Top             =   90
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   30
                  TabIndex        =   35
                  Top             =   90
                  Width           =   1770
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                     Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
               Height          =   555
               Index           =   6
               Left            =   1725
               Top             =   3780
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   39
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   555
               Index           =   6
               Left            =   3435
               Top             =   3780
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   1800
                  TabIndex        =   41
                  Top             =   90
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   30
                  TabIndex        =   40
                  Top             =   90
                  Width           =   1770
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                     Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
               Height          =   555
               Index           =   7
               Left            =   1725
               Top             =   4335
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   44
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   555
               Index           =   7
               Left            =   3435
               Top             =   4335
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   1800
                  TabIndex        =   46
                  Top             =   90
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   30
                  TabIndex        =   45
                  Top             =   90
                  Width           =   1770
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                     Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
               Height          =   555
               Index           =   8
               Left            =   1725
               Top             =   4890
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   30
                  TabIndex        =   49
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   555
               Index           =   8
               Left            =   3435
               Top             =   4890
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   1800
                  TabIndex        =   51
                  Top             =   90
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   30
                  TabIndex        =   50
                  Top             =   90
                  Width           =   1770
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                     Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
               Height          =   555
               Index           =   9
               Left            =   1725
               Top             =   5445
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   54
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   555
               Index           =   9
               Left            =   3435
               Top             =   5445
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   1800
                  TabIndex        =   56
                  Top             =   90
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
                  cFormat         =   "00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   30
                  TabIndex        =   55
                  Top             =   90
                  Width           =   1770
                  _ExtentX        =   2990
                  _ExtentY        =   582
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                     Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
               Height          =   435
               Left            =   15
               Top             =   15
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   767
               ForeColor       =   16711680
               Caption         =   "ì¸â◊êî"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   555
               Index           =   19
               Left            =   15
               Top             =   5445
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   53
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   18
               Left            =   15
               Top             =   4890
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   48
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   17
               Left            =   15
               Top             =   4335
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   43
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   16
               Left            =   15
               Top             =   3780
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   38
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   15
               Left            =   15
               Top             =   3225
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   33
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   14
               Left            =   15
               Top             =   2670
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   28
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   13
               Left            =   15
               Top             =   2115
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   23
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   12
               Left            =   15
               Top             =   1560
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   18
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   11
               Left            =   15
               Top             =   1005
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   13
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   10
               Left            =   15
               Top             =   450
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   979
               ForeColor       =   16711680
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Left            =   15
                  TabIndex        =   8
                  Top             =   90
                  Width           =   1680
                  _ExtentX        =   2963
                  _ExtentY        =   635
                  cFormat         =   "########0.000"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "########0.000"
                  MaxLength       =   13
                  cDataType       =   2
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFcDecimal      =   3
                  cFdAutoFormat   =   1
                  cGFormat        =   "########0.000"
                  cILength        =   6
               End
            End
         End
         Begin vsFlexLib.vsFlexArray VsGrid1 
            Height          =   6030
            Left            =   45
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   870
            Width           =   8310
            _Version        =   196608
            _ExtentX        =   14658
            _ExtentY        =   10636
            _StockProps     =   228
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "áÇ|çﬁ éø                   |îwî‘çÜ |é˚óeêî |î† êî  |î≠íçêî  |ï™äÑî[ì¸êî|äÆî[ó\íËì˙|ó\íËï÷|"
            Rows            =   11
            BackColor       =   12648447
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   0
            RowHeightMin    =   550
            GridLines       =   3
            SelectionMode   =   1
            BackColorAlternate=   12648447
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   735
            Left            =   12390
            Top             =   45
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   1296
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "   ì¸ â◊ ì˙  - ï÷"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               TabIndex        =   7
               Top             =   285
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   661
               cFormat         =   "00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               TabIndex        =   6
               Top             =   300
               Width           =   1770
               _ExtentX        =   2990
               _ExtentY        =   582
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'µÃå≈íË
      Left            =   960
      TabIndex        =   58
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
      BackColor       =   8438015
      ForeColor       =   0
      Caption         =   "Åy çﬁóøÅEì¸â◊é¿ê—ì¸óÕ [å¬ï ]  Åz"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Left            =   11790
         Top             =   30
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         BackColor       =   8438015
         Caption         =   "XXXXXXXXX1XXXXXXXXX2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   6720
      ScaleHeight     =   1050
      ScaleWidth      =   8490
      TabIndex        =   62
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
         Caption         =   "ï≈"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Caption         =   "î[ î‘"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Caption         =   " î[ ì¸ ì˙ - ï÷"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Caption         =   "éd ì¸ êÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Caption         =   "î≠íçãÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Caption         =   "éËîzãÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      TabIndex        =   66
      Top             =   1710
      Width           =   480
      _Version        =   196610
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
Attribute VB_Name = "CXHZ0220"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   çﬁóøì¸â◊é¿ê—ì¸óÕ [å¬ï ]
'**       ÉtÉHÅ[ÉÄID    :   CXHZ0220
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2004/02/24  By CIS
'**       ïœ  çX  ì˙    :   2004/04/08  By CIS - çXêVéûÇÃämîFÉÅÉbÉZÅ[ÉWèCê≥
'**       ïœ  çX  ì˙    :   2005/06/17  By CIS - é¿êîãÊï™=1ÇÃéûÅAèoâ◊êîÅ{ï™î[êîÅÑéÛíçêîÇÉGÉâÅ[Ç∆Ç∑ÇÈ
'**       ïœ  çX  ì˙    :   2005/11/17  By CIS ÉIÉvÉVÉáÉìí«â¡Åiâ“ì≠ì˙ÉJÉåÉìÉ_Å[Åj
'**                                             î≠íçì˙éZèo=é©é–ÉJÉåÉìÉ_Å[ÅAî[ì¸ì˙éZèo=édì¸êÊÉJÉåÉìÉ_Å[
'**       ïœ  çX  ì˙    :   2008/04/11  By CIS - îwî‘çÜÇWåÖëŒâû
'**       ïœ  çX  ì˙    :   2009/01/16  By CIS - 2005/06/17ÇÃïœçXÇñ≥å¯
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB                  As String * 2   ' ì¸óÕàÊêÿë÷
    
    Dim mSL_Select              As String
    
    Dim CisVsGrid               As New CisVsGrid3
    
    Dim mHeadProc               As String
    Dim mSaveHeadProc           As String
    
    Dim mBunCount               As Long
    
    Dim mCurren                 As Currency
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås    ÇkÇèÇÅÇÑÉCÉxÉìÉg    Åt              +
'+----------------------------------------------------------+
Private Sub Form_Load()

    If Rv_NCall = "" Then
    '   #-------------------#
    '   # ìÒ èd ãN ìÆ ñh é~  #
    '   #-------------------#
        If Not Dupli_Start Then End
    '   #-------------------#
    '   # èâ ä˙ ì‡ óe éÊ ìæ  #
    '   #-------------------#
        If Not IniGet Then End
        '   ( QR Reader ä¬ã´ )
        Set QRr_Object = PDQComm1
        Set QRr_MsgObject = QRMsg
        Set QRr_ErrMsgObj = ErrMsg
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
    
    Timer1.Interval = 1000   ' éûçèï\é¶ä‘äu
    Timer1.Enabled = True    ' éûçèï\é¶äJén
    Timer2.Interval = 500    ' èàóùãÊï™ì_ñ≈ä‘äu
    Timer2.Enabled = False   ' èàóùãÊï™ì_ñ≈ã÷é~
'   *-------------------*
'   *   ∏ﬁÿØƒﬁèâä˙èàóù
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
'   #  çÄ ñ⁄ èâ ä˙ âª   #
'   #-------------------#
    Call CisFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
    
'   #------------------#
'   #    èâä˙ì¸óÕê›íË   #
'   #------------------#
    K_Sykbnm.Tag = "4"
    Call PB_CHG_Click
    
'   #-------------------#
'   # âÊñ íÜâõà íuï\é¶   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    If Rv_NCall = "" Then
       Me.Show
    Else
    End If
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    ÅsÇ`ÇÉÇîÇâÇñÇÅÇîÇÖ ÉCÉxÉìÉgÅt                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = Not FormAct
'ÉVÉXÉeÉÄÉIÉvÉVÉáÉìälìæ     2005/11/17í«â¡
    Call GetSYSOption(33)

    mSL_Select = "SELECT HC.î[ïièëä«óùNO,HC.î[ïièëçs,HC.î[ì¸ì˙,HC.î[ì¸ï÷,"
    mSL_Select = mSL_Select & "  ISNULL(NK.ì¸â◊ì˙,'') ì¸â◊ì˙,ISNULL(NK.ì¸â◊ï÷,'') ì¸â◊ï÷,"
    mSL_Select = mSL_Select & "  ISNULL(NK.î“âÒì˙,'') î“âÒì˙,ISNULL(NK.î“âÒï÷,'') î“âÒï÷,"
    mSL_Select = mSL_Select & "  ISNULL(NK.ï™äÑêî,0) ï™äÑêî,ISNULL(NK.ì¸â◊êî,0) ì¸â◊êî,"
    mSL_Select = mSL_Select & "  HC.édì¸êÊ,HC.éÛì¸,HC.î[ì¸î‘çÜ,HC.î[ïièëï≈,"
    mSL_Select = mSL_Select & "  HC.î≠íçãÊï™,HC.éËîzãÊï™,HC.SKãÊï™,"
    mSL_Select = mSL_Select & "  ISNULL(TR.ó™èÃ,'') édì¸êÊñº,"
    mSL_Select = mSL_Select & "  ISNULL(SY1.ílñºèÃ,'') î≠íçãÊï™ñº,"
    mSL_Select = mSL_Select & "  ISNULL(SY2.ílñºèÃ,'') éËîzãÊï™ñº,"
    mSL_Select = mSL_Select & "  HC.çﬁéø,HC.îwî‘çÜ,HN.ï\é¶ê°ñ@,"
    mSL_Select = mSL_Select & "  HC.é˚óeêî,HC.ñáêî,HC.î≠íçêî,HN.é¿êîãÊï™ "
    mSL_Select = mSL_Select & " From çﬁóøî≠íçÉeÅ[ÉuÉã HC"
    mSL_Select = mSL_Select & "  Left Outer Join çﬁóøì¸â◊ÉeÅ[ÉuÉã NK"
    mSL_Select = mSL_Select & "     On NK.î≠íçä«óùNO = HC.î≠íçä«óùNO"
    mSL_Select = mSL_Select & "  Left Outer Join çﬁóøÉ}ÉXÉ^ HN"
    mSL_Select = mSL_Select & "     ON HN.çﬁóøä«óùî‘çÜ = HC.çﬁóøä«óùî‘çÜ"
    mSL_Select = mSL_Select & "  Left Outer Join éÊà¯êÊÉ}ÉXÉ^ TR"
    mSL_Select = mSL_Select & "     ON TR.éÊà¯êÊãÊï™ = 1"
    mSL_Select = mSL_Select & "     And TR.éÊà¯êÊCD = HC.édì¸êÊ"
    mSL_Select = mSL_Select & "  Left Outer Join ñºèÃÉ}ÉXÉ^ SY1"
    mSL_Select = mSL_Select & "     ON SY1.ãÊï™ñºèÃ = 'î≠íçãÊï™'"
    mSL_Select = mSL_Select & "     And SY1.ãÊï™É^ÉCÉv = ''"
    mSL_Select = mSL_Select & "     And SY1.íl = HC.î≠íçãÊï™"
    mSL_Select = mSL_Select & "  Left Outer Join ñºèÃÉ}ÉXÉ^ SY2"
    mSL_Select = mSL_Select & "     ON SY2.ãÊï™ñºèÃ = 'éËîzãÊï™'"
    mSL_Select = mSL_Select & "     And SY2.ãÊï™É^ÉCÉv = ''"
    mSL_Select = mSL_Select & "     And SY2.íl = HC.éËîzãÊï™"


    If Rv_NCall = "K" Then
       K_Sykbnm.Tag = "4"
       Call PB_CHG_Click
       Back_BunCount.Visible = True
       mBunCount = 1
       
       Call ContinuChk
       
    End If

    ' ì¸â◊é¿ê—èàóùÉèÅ[ÉNçÌèú
    If Rv_NCall = "" Then
        With CisDB
            .SQL = "Delete çﬁóøì¸â◊é¿ê—èàóùÉèÅ[ÉN"
            .SQL = .SQL & "    Where èàóùãÊï™ = 9"
            .DBExec
        End With
    End If
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås QueryUnloadÉCÉxÉìÉg Åt                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Rv_NCall = "" Then
        Cancel = 1
    End If
    Call PB_END_Click
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås ÇjÇÖÇôÇcÇèÇóÇéÉCÉxÉìÉg Åt                 +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄÅs ÉtÉ@ÉìÉNÉVÉáÉìÉLÅ[ópÉAÉNÉZÉâÉåÅ[É^ Åt         +
'+-------------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False
    
    Select Case Key_Code
            Case vbKeyF6:       If PB_CHG.Visible Then Call PB_CHG_Click        'Åyêÿë÷Åz
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  'ÅyåüçıÅz
           Case vbKeyPageUp:    If PB_BACK.Visible Then Call PB_Back_Click  'Åyå„ëﬁÅz
           Case vbKeyPageDown:  If PB_NEXT.Visible Then Call PB_Next_Click  'ÅyëOêiÅz
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click    'ÅyéÊè¡Åz
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    'ÅyèIóπÅz
           Case vbKeyReturn
                                If PB_Ent.Visible Then                      'Åyì¸óÕÅz
                                    Call PB_ENT_Click
                                End If
           Case Else:           Exit Function
    End Select
    
    Key_Acc = True
End Function
'+---------------------+
'+    èI  óπ  ÉL Å[     +
'+---------------------+
Private Sub PB_END_Click()
    If Rv_NCall = "" Then
        gSL_Select = "Select Count(*) åèêî From çﬁóøì¸â◊é¿ê—èàóùÉèÅ[ÉN"
        gSL_Select = gSL_Select & " Where èàóùí[ññ = Host_Name()"
        gSL_Select = gSL_Select & "   And èàóùãÊï™ = 0"
        If ZJWRead(gSL_Select, 1) Then
          If ZJW.åèêî <> 0 Then
          ' î≠íçèàóù
            On Error Resume Next
            CisDB.PassLink
            With CisFun
                .ShellApl = gIniExe & "PXHZ0110.exe"
                .ShellAplPara = "J"
                .ShellTask
            End With
            On Error GoTo 0
          End If
        End If
        
        Call CisDB.DBDISConnect                            ' √ﬁ∞¿Õﬁ∞Ωêÿíf
        Call CisFrm.UnLoadDisp                          ' Ã´∞—µÃﬁºﬁ™∏ƒè¡ãé
        End
    Else
        Timer1.Enabled = False
        Me.Hide
    End If
End Sub
'+----------------------------+
'+    ì¸  óÕ  ÉL Å[(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+    éÊ  è¡  ÉL Å[(Escape)    +
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
           End If
        End If
        ProcHB = mHeadProc
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        GoTo PB_CAN_Ed
    End If
    
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+------------------+
'+    êÿë÷   (F6)   +
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
            K_Sykbnm.Caption = "Åy  ﬁ∞∫∞ƒﬁ Åz"
            H2_Area1.Visible = True
        Case "H3"
            K_Sykbnm.Caption = "Åy QR¿Ø¡ÿ∞¿ﬁ Åz"
            H3_Area1.Visible = True
            '   ( QR Reader ä¬ã´ )
            Set QRr_Object = PDQComm1
            Set QRr_MsgObject = QRMsg
            Set QRr_ErrMsgObj = ErrMsg
            Call QRInit
        Case "H4"
            K_Sykbnm.Caption = "ÅyQRŒØ ﬂÿ∞¿ﬁÅz"
            H4_Area1.Visible = True
            '   ( QR Reader ä¬ã´ )
            Set QRr_Object = PDQComm2
            Set QRr_MsgObject = QRRMsg
            Set QRr_ErrMsgObj = ErrMsg
            Call TX2400InfoGet
        Case Else
            K_Sykbnm.Caption = "Åy î[î‘ Åz"
            H1_Area1.Visible = True
    End Select
    ProcHB = mHeadProc
    Call DispChange(ProcHB)
End Sub
'+----------------------------+
'+    å„  ëﬁ  ÉL Å[(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = mSL_Select
    gSL_Select = gSL_Select & " WHERE  HC.î[ì¸ì˙ + HC.î[ì¸î‘çÜ + Cast(HC.î[ïièëï≈ as Char(2)) < "
    gSL_Select = gSL_Select & " '" & RTrim$(iH1_NYmd) & RTrim$(iH1_Nouno) & Format(iH1_Page, "0") & "'"
    gSL_Select = gSL_Select & " And ISNULL(ì¸â◊ì˙,'') = '' "
    gSL_Select = gSL_Select & " ORDER BY HC.î[ì¸ì˙ + HC.î[ì¸î‘çÜ + Cast(HC.î[ïièëï≈ as Char(2)) DESC "
    
    If Not ZCTRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    êÊì™ÉåÉRÅ[ÉhÇ≈Ç∑            "
            .MB_Title = "ÅyëOï≈Åz"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Back_ED
    End If
    
    iH1_NYmd = ZCT.î[ì¸ì˙
    iH1_Nouno = ZCT.î[ì¸î‘çÜ
    iH1_Page = ZCT.î[ïièëï≈
    
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where HC.î[ì¸î‘çÜ = '" & RTrim$(iH1_Nouno) & "'"
    gSL_Select = gSL_Select & " And   HC.î[ïièëï≈ =  " & CisFun.Val2(iH1_Page)
    gSL_Select = gSL_Select & " And   HC.î[ì¸ì˙   = '" & RTrim$(iH1_NYmd) & "'"
    
    Call DBInput("Back")
PB_Back_ED:
    
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    ëO  êi  ÉL Å[(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = mSL_Select
    gSL_Select = gSL_Select & " WHERE  HC.î[ì¸ì˙ + HC.î[ì¸î‘çÜ + Cast(HC.î[ïièëï≈ as Char(2)) > "
    gSL_Select = gSL_Select & " '" & RTrim$(iH1_NYmd) & RTrim$(iH1_Nouno) & Format(iH1_Page, "0") & "'"
    gSL_Select = gSL_Select & " And ISNULL(ì¸â◊ì˙,'') = '' "
    gSL_Select = gSL_Select & " ORDER BY HC.î[ì¸ì˙ + HC.î[ì¸î‘çÜ + Cast(HC.î[ïièëï≈ as Char(2)) ASC "
    
    If Not ZCTRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ç≈èIÉåÉRÅ[ÉhÇ≈Ç∑            "
            .MB_Title = "Åyéüï≈Åz"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Next_ED
    End If
    
    iH1_NYmd = ZCT.î[ì¸ì˙
    iH1_Nouno = ZCT.î[ì¸î‘çÜ
    iH1_Page = ZCT.î[ïièëï≈
    
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where HC.î[ì¸î‘çÜ = '" & RTrim$(iH1_Nouno) & "'"
    gSL_Select = gSL_Select & " And   HC.î[ïièëï≈ =  " & CisFun.Val2(iH1_Page)
    gSL_Select = gSL_Select & " And   HC.î[ì¸ì˙   = '" & RTrim$(iH1_NYmd) & "'"
    
    Call DBInput("Next")
PB_Next_ED:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_NYmd" Then iH1_NYmd.ShowCalender
    If PB_Look.Tag = "iB1_SYmd" Then iB1_SYmd.ShowCalender
    If PB_Look.Tag Like "iB1_BYmd*" Then iB1_BYmd(CisFun.Val2(Mid(PB_Look.Tag, 9))).ShowCalender
End Sub
'*******************************
'*      É^ÉCÉgÉãÉoÅ[ï“èW        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*----------------------------------------------*
'*      î[ ì¸ ì˙  ( iH1_NYmd )       *
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
'*      èo â◊ ì˙  ( iB1_SYmd )       *
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
'*       ì¸â◊êî  ( iB1_NSuryo )                 *
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
    mCurren = CisFun.Val2(iB1_NSuryo(Index).Tag) - CisFun.Val2(iB1_NSuryo(Index))
    If mCurren <> 0 Then
       iB1_Suryo(Index) = CisFun.RSetFld(mCurren, 0, iB1_Suryo(0).cFormat)
    Else
       iB1_Suryo(Index) = ""
    End If
End Sub
'*----------------------------------------------*
'*       ï™äÑî[ì¸êî  ( iB1_Suryo )              *
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
    mCurren = CisFun.Val2(iB1_NSuryo(Index).Tag) - CisFun.Val2(iB1_Suryo(Index))
    If mCurren <> 0 Then
       iB1_NSuryo(Index) = CisFun.RSetFld(mCurren, 0, iB1_NSuryo(0).cFormat)
    Else
       iB1_NSuryo(Index) = ""
    End If
End Sub
'*----------------------------------------------*
'*       äÆî[ó\íËì˙  ( iB1_BYmd )       *
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
'*       äÆî[ó\íËï÷  ( iB1_BBin )               *
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
'*----------------------------------------------*
'*       îıçl  ( iB1_Biko )               *
'*----------------------------------------------*
Private Sub iB1_Biko_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim ID      As Integer
    If KeyCode = vbKeyDown Then
       ID = Index + 1
       Do Until ID > 9
          If iB1_Biko(ID).Visible Then
             iB1_Biko(ID).SetFocus
             Exit Do
          End If
          ID = ID + 1
       Loop
       Exit Sub
    End If
    If KeyCode = vbKeyUp Then
       ID = Index - 1
       Do Until ID < 0
          If iB1_Biko(ID).Visible Then
             iB1_Biko(ID).SetFocus
             Exit Do
          End If
          ID = ID - 1
       Loop
       Exit Sub
    End If
End Sub

'****************************
'*                          *
'*      HEAD/BODY èàóù      *
'*                          *
'****************************
'****************************
'*      ì¸óÕêÿë÷ èàóù        *
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
    
'   ¿ﬁ–∞çÄñ⁄(âÊñ ÇÃÉ`ÉâÉcÉLñhé~)
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
       PB_END.BtnCaption = "ñﬂÇÈ"
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
       Call TX2400OpenAuto
       Call TX2400QRRReadMsg("R")
    End If

' ¿ﬁ–∞çÄñ⁄(ì¸óÕã÷é~)
    Dummy.Enabled = False

End Function
'**************************************
'*      Head/Body èâä˙âªèàóù           *
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
'*      ì¸óÕèàóù(Enter)      *
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
'*      HEAD1¡™Ø∏èàóù
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
'( î[ì¸ì˙ )
    If RTrim$(iH1_NYmd) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  î[ì¸ì˙Çì¸óÕÇµÇƒâ∫Ç≥Ç¢  "
            .MB_Button = Error
            .MBOX
        End With
        iH1_NYmd.SetFocus
        Exit Function
    Else
        If Not iH1_NYmd.cChkResult Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    ê≥ÇµÇ¢ì˙ïtÇì¸óÕÇµÇƒâ∫Ç≥Ç¢         "
                 .MB_Button = Error
                 .MBOX
             End With
             iH1_NYmd.SetFocus
             Exit Function
         End If
    End If
    
'( î[ì¸î‘çÜ )
    If RTrim$(iH1_Nouno) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  î[î‘Çì¸óÕÇµÇƒâ∫Ç≥Ç¢  "
            .MB_Button = Error
            .MBOX
        End With
        iH1_Nouno.SetFocus
        Exit Function
    End If
    
'( î[ïièëï≈ )
    If CisFun.Val2(iH1_Page) = 0 Then iH1_Page = " 1"
    
    If Rv_NCall = "" Or Rv_NCall = "K2" Or CisFun.Val2(iH1_Nouno.Tag) = 0 Then
        gSL_Select = "Select HC.î[ïièëä«óùNO From çﬁóøî≠íçÉeÅ[ÉuÉã HC"
        gSL_Select = gSL_Select & " Where HC.î[ì¸î‘çÜ = '" & RTrim$(iH1_Nouno) & "'"
        gSL_Select = gSL_Select & " And   HC.î[ïièëï≈ =  " & CisFun.Val2(iH1_Page)
        gSL_Select = gSL_Select & " And   HC.î[ì¸ì˙   = '" & RTrim$(iH1_NYmd) & "'"
        gSL_Select = gSL_Select & " Group By HC.î[ïièëä«óùNO"
        If Not ZCTRead(gSL_Select) Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "  äYìñÉfÅ[É^Ç™Ç†ÇËÇ‹ÇπÇÒ  "
                .MB_Button = Error
                .MBOX
            End With
            Exit Function
        End If
        If CisDB.RecordCount = 1 Then
           iH1_Nouno.Tag = ZCT.î[ïièëä«óùNO
           Call ZCTClose
        Else
            gSL_Select = "Select HC.*,HN.ï\é¶ê°ñ@ From çﬁóøî≠íçÉeÅ[ÉuÉã HC "
            gSL_Select = gSL_Select & " Left Outer Join çﬁóøÉ}ÉXÉ^ HN "
            gSL_Select = gSL_Select & "   ON HN.çﬁóøä«óùî‘çÜ = HC.çﬁóøä«óùî‘çÜ "
            gSL_Select = gSL_Select & " Left Outer Join çﬁóøì¸â◊ÉeÅ[ÉuÉã NY "
            gSL_Select = gSL_Select & "   On NY.î≠íçä«óùNO = HC.î≠íçä«óùNO "
            gSL_Select = gSL_Select & " Where HC.î[ì¸î‘çÜ = '" & RTrim$(iH1_Nouno) & "'"
            gSL_Select = gSL_Select & " And   HC.î[ïièëï≈ =  " & CisFun.Val2(iH1_Page)
            gSL_Select = gSL_Select & " And   HC.î[ì¸ì˙   = '" & RTrim$(iH1_NYmd) & "'"
            gSL_Select = gSL_Select & " And   Isnull(NY.ì¸â◊ì˙,'') = ''"
            gSL_Select = gSL_Select & " Order By HC.çﬁéø,HN.ï\é¶ê°ñ@"
            Call ZCTRead(gSL_Select)
            If CisDB.RecordCount = 1 Then
               iH1_Nouno.Tag = ZCT.î[ïièëä«óùNO
               Call ZCTClose
            Else
                CXHZ0215.Show vbModal
                Unload CXHZ0215
                Set CXHZ0215 = Nothing
                If RV_Rtn Then
                   iH1_Nouno.Tag = Rvc_NKNo
                Else
                   With CisFun
                       .MB_Lines = 5
                       .MB_MSG(2) = "  ì`ï[ÇÃëIëÇ™Ç≥ÇÍÇ‹ÇπÇÒÇ≈ÇµÇΩ  "
                       .MB_Button = Error
                       .MBOX
                   End With
                   Exit Function
                End If
            End If
        End If
    End If
'( ÉfÅ[É^ämîF )
    gSL_Select = mSL_Select

    gSL_Select = gSL_Select & " Where HC.î[ïièëä«óùNO = " & CisFun.Val2(iH1_Nouno.Tag)
    gSL_Select = gSL_Select & "   And HC.î[ïièëï≈ = " & CisFun.Val2(iH1_Page)
    
    If Not DBInput("INV") Then Exit Function
    
    Head1Chk = True
End Function
'****************************
'*      HEAD2¡™Ø∏èàóù
'****************************
Private Function Head2Chk() As Boolean
    Head2Chk = False
    
'( ÉoÅ[ÉRÅ[Éh )
    If CisFun.Val2(iH2_Bar) = 0 Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  ÉoÅ[ÉRÅ[ÉhÇì¸óÕÇµÇƒâ∫Ç≥Ç¢  "
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
    
    
'( ÉfÅ[É^ämîF )
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where HC.î[ïièëä«óùNO = " & CisFun.Val2(Mid(iH2_Bar, 1, 10))
    gSL_Select = gSL_Select & "   And HC.î[ïièëï≈ = " & CisFun.Val2(Mid(iH2_Bar, 11, 2))
    
    If Not DBInput("INV") Then Exit Function
    
    Head2Chk = True
End Function
'****************************
'*      Body1¡™Ø∏èàóù
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    If iB1_SYmd = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ì¸â◊ì˙Çì¸óÕÇµÇƒâ∫Ç≥Ç¢         "
            .MB_Button = Error
            .MBOX
        End With
        iB1_SYmd.SetFocus
        Exit Function
    End If
    If Not iB1_SYmd.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ì¸â◊ì˙Çê≥ÇµÇ¢ì˙ïtÇ…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢         "
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
                .MB_MSG(2) = "    ì¸â◊ï÷Çì¸óÕÇµÇƒâ∫Ç≥Ç¢                 "
                .MB_Button = Error
                .MBOX
            End With
            iB1_SBin.SetFocus
            Exit Function
        End If
        If B1lb_NYmd.Tag & B1lb_NBin > iB1_SYmd & iB1_SBin Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ì¸â◊ì˙ÅEï÷ÇÕî[ì¸ì˙ÅEï÷à»ç~Ç…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢         "
                .MB_Button = Error
                .MBOX
            End With
            iB1_SYmd.SetFocus
            Exit Function
        End If
    Else
        If B1lb_NYmd.Tag > iB1_SYmd Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ì¸â◊ì˙ÇÕî[ì¸ì˙à»ç~Ç…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢         "
                .MB_Button = Error
                .MBOX
            End With
            iB1_SYmd.SetFocus
            Exit Function
        End If
    End If
    
    If gSYSOp_CLKb = 1 Then
        gInt = Kdou_Check(iB1_SYmd, 1, ZCT.édì¸êÊ)
        If gInt = 2 Then
            gInt = Kdou_Check(iB1_SYmd)
        End If
    Else
        gInt = Kdou_Check(iB1_SYmd)
    End If
    If gInt = 1 Then
    'ãxì˙
        CisFun.MB_Lines = 4
        CisFun.MB_MSG(1) = "     ì¸â◊ì˙Ç™îÒâ“ì≠ì˙Ç≈Ç∑ÅB     "
        CisFun.MB_MSG(3) = "     Å@ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH      "
        CisFun.MB_Button = OK_CAN
        If Not CisFun.MBOX Then
            iB1_SYmd.SetFocus
            Exit Function
        End If
    End If
    If gInt = 2 Then
    'ÉJÉåÉìÉ_ñ¢ìoò^
        CisFun.MB_Lines = 4
        CisFun.MB_MSG(1) = "     " & RTrim(iB1_SYmd.cYear) & "îN" & RTrim(iB1_SYmd.cMonth) & "åéÇÃÉJÉåÉìÉ_ñ¢ìoò^ÅI     "
        CisFun.MB_MSG(3) = "     Å@Å@(Å@ÉJÉåÉìÉ_É}ÉXÉ^Å@)Å@"
        CisFun.MB_Button = Error
        CisFun.MBOX
        iB1_SYmd.SetFocus
        Exit Function
    End If
    
    Set gObj = Nothing
    
    For gCnt1 = 0 To 9
        If iB1_Suryo(gCnt1).Visible Then
           mCurren = CisFun.Val2(VsGrid1.TextMatrix(gCnt1 + 1, 5))
           
           If Right(B1lb_Nouno, 1) = "9" Then
              If CisFun.Val2(iB1_Suryo(gCnt1)) <> 0 Then
                 CisFun.MB_Lines = 4
                 CisFun.MB_MSG(2) = "     ï™î[âÒêîÇ™ÇXâÒÇí¥Ç¶ÇÈï™î[èàóùÇÕÇ≈Ç´Ç‹ÇπÇÒ            "
                 CisFun.MB_Button = Error
                 CisFun.MBOX
                 iB1_Suryo(gCnt1).SetFocus
                 Exit Function
              End If
           End If
           If CisFun.Val2(iB1_NSuryo(gCnt1)) = 0 And _
              CisFun.Val2(iB1_Suryo(gCnt1)) = 0 Then
              gStr = "ì¸â◊êîñîÇÕï™äÑî[ì¸êîÇì¸óÕÇµÇƒâ∫Ç≥Ç¢"
              Set gObj = iB1_NSuryo(gCnt1)
              Exit For
           End If
           If CisFun.Val2(iB1_NSuryo(gCnt1)) > mCurren And _
              Trim(VsGrid1.TextMatrix(gCnt1 + 1, 6)) = "" Then
              gStr = "ì¸â◊êîÇÕî≠íçêîà»â∫Ç…ÇƒéwíËÇµÇƒâ∫Ç≥Ç¢"
              Set gObj = iB1_NSuryo(gCnt1)
              Exit For
           End If
           If CisFun.Val2(iB1_Suryo(gCnt1)) > mCurren And _
              Trim(VsGrid1.TextMatrix(gCnt1 + 1, 6)) = "" Then
              gStr = "ï™äÑî[ì¸êîÇÕî≠íçêîà»â∫Ç…ÇƒéwíËÇµÇƒâ∫Ç≥Ç¢"
              Set gObj = iB1_Suryo(gCnt1)
              Exit For
           End If
'-----------------------------------------------------------------------------Å•2005.06.17 AddÅ•------
'================================= 2009/01/19 Start
'           If CisFun.Val2(iB1_Suryo(gCnt1)) + CisFun.Val2(iB1_NSuryo(gCnt1)) > mCurren And _
'              Trim(VsGrid1.TextMatrix(gCnt1 + 1, 6)) <> "" Then
'              gStr = "ì¸â◊êîÇ∆ï™äÑî[ì¸êîÇÃçáåvÇÕî≠íçêîà»â∫Ç…ÇƒéwíËÇµÇƒâ∫Ç≥Ç¢"
'              Set gObj = iB1_Suryo(gCnt1)
'              Exit For
'           End If
'================================= 2009/01/19 End
'-----------------------------------------------------------------------------Å£2005.06.17 AddÅ£------
           If Trim(iB1_BYmd(gCnt1)) <> "" And _
              CisFun.Val2(iB1_Suryo(gCnt1)) = 0 Then
               gStr = "äÆî[ó\íËì˙ì¸óÕéûÅAï™äÑî[ì¸êîÇì¸óÕÇµÇƒâ∫Ç≥Ç¢"
               Set gObj = iB1_Suryo(gCnt1)
               Exit For
            End If
              
           If CisFun.Val2(iB1_Suryo(gCnt1)) <> 0 Then
              If Trim(iB1_BYmd(gCnt1)) = "" Then
                 gStr = "ï™äÑî[ì¸êîì¸óÕéûÅAäÆî[ó\íËì˙Çì¸óÕÇµÇƒâ∫Ç≥Ç¢"
                 Set gObj = iB1_BYmd(gCnt1)
                 Exit For
              End If
              If Not iB1_BYmd(gCnt1).cChkResult Then
                 gStr = "äÆî[ó\íËì˙Çê≥ÇµÇ¢ì˙ïtÇ…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢"
                 Set gObj = iB1_BYmd(gCnt1)
                 Exit For
              End If
                If gSYSOp_CLKb = 1 Then
                    gInt = Kdou_Check(iB1_BYmd(gCnt1), 1, ZCT.édì¸êÊ)
                    If gInt = 2 Then
                        gInt = Kdou_Check(iB1_BYmd(gCnt1))
                    End If
                Else
                    gInt = Kdou_Check(iB1_BYmd(gCnt1))
                End If
              If gInt = 1 Then
                'ãxì˙
                 CisFun.MB_Lines = 4
                 CisFun.MB_MSG(1) = "       " & StrConv(gCnt1 + 1, vbWide) & "çsñ⁄ äÆî[ó\íËì˙Ç™îÒâ“ì≠ì˙Ç≈Ç∑ÅB     "
                 CisFun.MB_MSG(3) = "     Å@   ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH          "
                 CisFun.MB_Button = OK_CAN
                 If Not CisFun.MBOX Then
                    iB1_BYmd(gCnt1).SetFocus
                    Exit Function
                 End If
              End If
              If gInt = 2 Then
                'ÉJÉåÉìÉ_ñ¢ìoò^
                 CisFun.MB_Lines = 4
                 CisFun.MB_MSG(1) = "     " & RTrim(iB1_BYmd(gCnt1).cYear) & "îN" & RTrim(iB1_BYmd(gCnt1).cMonth) & "åéÇÃÉJÉåÉìÉ_ñ¢ìoò^ÅI     "
                 CisFun.MB_MSG(3) = "     Å@Å@(Å@ÉJÉåÉìÉ_É}ÉXÉ^Å@)Å@"
                 CisFun.MB_Button = Error
                 CisFun.MBOX
                 iB1_BYmd(gCnt1).SetFocus
                 Exit Function
              End If
              If iB1_BBin(gCnt1).Visible Then
                 If Trim(B1lb_HKbn.Tag) = "1" And CisFun.Val2(iB1_BBin(gCnt1)) = 0 Then
                    gStr = "ï™äÑî[ì¸êîì¸óÕéûÅAäÆî[ó\íËï÷Çì¸óÕÇµÇƒâ∫Ç≥Ç¢"
                    Set gObj = iB1_BBin(gCnt1)
                    Exit For
                 End If
                 If iB1_SYmd & iB1_SBin > iB1_BYmd(gCnt1) & iB1_BBin(gCnt1) Then
                    gStr = "äÆî[ó\íËì˙ÅEï÷ÇÕì¸â◊ì˙ÅEï÷à»ç~Ç…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢"
                    Set gObj = iB1_BYmd(gCnt1)
                    Exit For
                 End If
              Else
                 If iB1_SYmd > iB1_BYmd(gCnt1) Then
                    gStr = "äÆî[ó\íËì˙ÇÕì¸â◊ì˙à»ç~Ç…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢"
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
            .MB_MSG(2) = "          " & StrConv(gCnt1 + 1, vbWide) & "çsñ⁄ " & gStr & "                "
            .MB_Button = Error
            .MBOX
        End With
        gObj.SetFocus
        Exit Function
    End If
    Body1Chk = True
End Function
'+---------------------------------------+
'+  àÍäáÇ©ÇÁÇÃòAågèàóù
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
    gSL_Select = "Select î[ïièëä«óùNO,î[ì¸î‘çÜ,î[ì¸ì˙,î[ïièëï≈ "
    gSL_Select = gSL_Select & " From çﬁóøî≠íçÉeÅ[ÉuÉã "
    gSL_Select = gSL_Select & " Where î[ïièëä«óùNO = " & Rv_HNo(mBunCount)
    gSL_Select = gSL_Select & "   And î[ïièëï≈ = " & Rv_HPage(mBunCount)
    If Not ZCTRead(gSL_Select, 1) Then
       Call PB_END_Click
    End If
    
    iH1_NYmd = ZCT.î[ì¸ì˙
    iH1_Nouno.Tag = ZCT.î[ïièëä«óùNO
    iH1_Nouno = ZCT.î[ì¸î‘çÜ
    iH1_Page = CisFun.RSetFld(ZCT.î[ïièëï≈, 0, iH1_Page.cFormat)
    ProcHB = "H1"
    Call PB_ENT_Click
    
    ContinuChk = True

End Function
'**************************************
'*                                    *
'*       ÉfÅ[É^ÉxÅ[ÉXì¸óÕ Åï ï\é¶       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    
    Call HeadBodyClear("B")
    
    On Error GoTo DBInput_Err:
    If Not ZCTRead(gSL_Select, 0) Then
        If InputKey = "Next" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ç≈èIÉåÉRÅ[ÉhÇ≈Ç∑            "
                .MB_Title = "Åyéüï≈Åz"
                .MB_Button = OK
                .MBOX
            End With
            GoTo DBInput_Ed
        End If
        If InputKey = "Back" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    êÊì™ÉåÉRÅ[ÉhÇ≈Ç∑            "
                .MB_Title = "ÅyëOï≈Åz"
                .MB_Button = OK
                .MBOX
            End With
        End If
        If InputKey = "INV" Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "  äYìñÉfÅ[É^Ç™Ç†ÇËÇ‹ÇπÇÒ  "
                .MB_Button = Error
                .MBOX
            End With
        End If
        If InputKey = "QR" Then
           QRDataArea.ÉGÉâÅ[ì‡óe = "äYìñÉfÅ[É^Ç»Çµ"
        End If
        GoTo DBInput_Ed
    Else
        Call BodySet(InputKey)
        
        If ZCT.ì¸â◊ì˙ <> "" Then
            If InputKey = "QR" Then
               QRDataArea.ÉGÉâÅ[ì‡óe = "ì¸â◊çœÇ›ÉfÅ[É^ÅB"
               Exit Function
            End If
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "      ì¸â◊çœÇ›ÉfÅ[É^Ç≈Ç∑              "
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
'*          BODY ì‡óeæØƒèàóù           *
'**************************************
Private Sub BodySet(SetType As String)
'
    With ZCT
        iB1_SYmd.Tag = ZCT.î[ïièëä«óùNO
        iB1_SBin.Tag = ZCT.î[ïièëï≈
        
        B1lb_Nouno = ZCT.î[ì¸î‘çÜ
        B1lb_Page = ZCT.î[ïièëï≈
    ' ì¸â◊ì˙•èâä˙ï\é¶
        If SetType = "INV" Or SetType = "QR" Then
            iB1_SYmd = .î[ì¸ì˙
            If Rv_NCall = "K" And Trim(Rv_NYmd) <> "" Then
               iB1_SYmd = Rv_NYmd
            End If
            If Trim(ZCT.î[ì¸ï÷) <> "" Then
               iB1_SBin.Visible = True
               iB1_SBin = .î[ì¸ï÷
               If Rv_NCall = "K" And Trim(Rv_NBin) <> "" Then
                  iB1_SBin = Rv_NBin
               End If
            Else
               iB1_SBin.Visible = False
            End If
        End If
        
        If Trim(.édì¸êÊñº) = "" Then
           gStr = .édì¸êÊ
        Else
           gStr = RTrim(.édì¸êÊ) & ":" & .édì¸êÊñº
        End If
        B1lb_Tornm = gStr
        
        B1lb_NYmd.Tag = ZCT.î[ì¸ì˙
        With CisFun
            .DateE_BefVal = ZCT.î[ì¸ì˙
            .DateEdit
            B1lb_NYmd = CisFun.DateE_AfVal
            If Trim(ZCT.î[ì¸ï÷) <> "" Then
                B1lb_NBin = ZCT.î[ì¸ï÷
            End If
        End With
        B1lb_HKbn = .î≠íçãÊï™ñº
        B1lb_HKbn.Tag = .î≠íçãÊï™
        B1lb_TKbn = .éËîzãÊï™ñº
    End With
    
    Call GridSet(SetType)
    
End Sub
'**************************************
'*                                    *
'*          ÉfÅ[É^ÉxÅ[ÉXçXêV           *
'*                                    *
'**************************************
'*-------------------------------------------------*
'*      Éf Å[ É^ çX êV  äm îF ÉÅ Éb ÉZ Å[ ÉW
'*-------------------------------------------------*
Private Function CheckMessage() As Boolean
    CheckMessage = False
    
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     ì¸â◊é¿ê—èàóùÇçsÇ¢Ç‹Ç∑      "
        .MB_MSG(3) = "     ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH      "
        .MB_Title = "é¿ê—çXêVèàóù"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    CheckMessage = True
    
End Function
'+----------------------+
'+    í«â¡•çXêV èàóù     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Static AddFlg               As Boolean
    
'   ( ämîFÉÅÉbÉZÅ[ÉW )
    If Not CheckMessage Then Exit Function
    
    gSL_Select = "Select * From çﬁóøî≠íçÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where î[ïièëä«óùNO = " & CisFun.Val2(iB1_SYmd.Tag)
    gSL_Select = gSL_Select & "   And î[ïièëï≈ = " & CisFun.Val2(iB1_SBin.Tag)
    
    If Not ZCTRead(gSL_Select, , 1) Then
       Call ZCTClose
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     î≠íçÉfÅ[É^àŸèÌ      "
            .MB_Title = "çXêVÉGÉâÅ["
            .MB_Button = OK
            .MBOX
            Exit Function
       End With
    End If
    
    CisDB.ConnectNo = 0
    Call CisDB.DBTran(TransBegin) ' ƒ◊›ªﬁ∏ºÆ›äJén
    
    
    Do Until Not ZCT_RDSTS
        
        AddFlg = False
        
        gSL_Select = "Select * From çﬁóøì¸â◊ÉeÅ[ÉuÉã "
        gSL_Select = gSL_Select & " Where î≠íçä«óùNO = " & ZCT.î≠íçä«óùNO
        If Not ZYTRead(gSL_Select, 1) Then
           Call ItemsClearZYT
           ZYT.î≠íçä«óùNO = ZCT.î≠íçä«óùNO
           ZYT.çﬁóøä«óùî‘çÜ = ZCT.çﬁóøä«óùî‘çÜ
           AddFlg = True
        End If
        
        With ZYT
                
            .î[ì¸î‘çÜ = ZCT.î[ì¸î‘çÜ
            .î[ïièëï≈ = ZCT.î[ïièëï≈
            .î[ïièëçs = ZCT.î[ïièëçs
            .édì¸êÊ = ZCT.édì¸êÊ
            .éÛì¸ = ZCT.éÛì¸
            .çﬁéø = ZCT.çﬁéø
            .î¬å˙ = ZCT.î¬å˙
            .ïù = ZCT.ïù
            .í∑Ç≥ = ZCT.í∑Ç≥
            .îwî‘çÜ = ZCT.îwî‘çÜ
            .é˚óeêî = ZCT.é˚óeêî
            
            .ì¸â◊ì˙ = iB1_SYmd
            If iB1_SBin.Visible Then
                .ì¸â◊ï÷ = Format(iB1_SBin, "00")
            Else
                .ì¸â◊ï÷ = ""
            End If
            gLong = ZCT.î[ïièëçs - 1
            .ì¸â◊í[êî = 0
            .ï™äÑêî = 0
            .î“âÒì˙ = ""
            .î“âÒï÷ = ""
            .ì¸â◊îıçl = RTrim(iB1_Biko(gLong))
            If CisFun.Val2(iB1_Suryo(gLong)) = 0 Then
               .ì¸â◊êî = CisFun.Val2(iB1_NSuryo(gLong))
            Else
               .ì¸â◊êî = CisFun.Val2(iB1_NSuryo(gLong))
               .ï™äÑêî = CisFun.Val2(iB1_Suryo(gLong))
               .î“âÒì˙ = iB1_BYmd(gLong)
               If iB1_BBin(gLong).Visible Then
                  .î“âÒï÷ = Format(iB1_BBin(gLong), "00")
               End If
            End If
            gCnt1 = 0
'            If ZCT.é˚óeêî <> 0 Then
'               .ì¸â◊ñáêî = .ì¸â◊êî \ .é˚óeêî
'               .ì¸â◊í[êî = .ì¸â◊êî Mod .é˚óeêî
'               If .ì¸â◊í[êî <> 0 Then
'                  .ì¸â◊ñáêî = .ì¸â◊ñáêî + 1
'               End If
'            End If
        End With
    
        If AddFlg Then
           ZYT.çÏê¨é“ = gTanto
           Call ZYTInsert
        Else
           ZYT.çXêVé“ = gTanto
           Call ZYTUpdate
        End If
                
        With CisDB
            .SQL = "UPDATE çﬁóøì¸â◊ÉeÅ[ÉuÉã SET "
            .SQL = .SQL & "  ì¸â◊ñáêî = CASE WHEN é˚óeêî = 0 THEN 0"
            .SQL = .SQL & "                  Else"
            .SQL = .SQL & "                Case WHEN ì¸â◊êî - (é˚óeêî * (CAST(ROUND(ì¸â◊êî / é˚óeêî,0,1) as int))) <> 0 Then"
            .SQL = .SQL & "                                       CAST(ROUND(ì¸â◊êî / é˚óeêî,0,1) as int) + 1"
            .SQL = .SQL & "                       ELSE CAST(ROUND(ì¸â◊êî / é˚óeêî,0,1) as int)  END END,"
            .SQL = .SQL & "    ì¸â◊í[êî = CASE WHEN é˚óeêî = 0 THEN 0"
            .SQL = .SQL & "                  Else"
            .SQL = .SQL & "                    Case WHEN ì¸â◊êî < é˚óeêî THEN ì¸â◊êî"
            .SQL = .SQL & "                            WHEN ì¸â◊êî = é˚óeêî THEN 0"
            .SQL = .SQL & "                            WHEN ì¸â◊êî > é˚óeêî THEN"
            .SQL = .SQL & "                                 ì¸â◊êî - (é˚óeêî * (CAST(ROUND(ì¸â◊êî / é˚óeêî,0,1) as int))) "
            .SQL = .SQL & "                    End"
            .SQL = .SQL & "                  End"
            .SQL = .SQL & "    Where î≠íçä«óùNO = " & ZCT.î≠íçä«óùNO
            If Not .DBExec Then GoTo DBPut_Err
        End With
        
        If CisFun.Val2(iB1_Suryo(gLong)) <> 0 Then
           gSL_Select = "Select * From çﬁóøì¸â◊ÉeÅ[ÉuÉã "
           gSL_Select = gSL_Select & " Where î≠íçä«óùNO = " & ZCT.î≠íçä«óùNO
           If ZYTRead(gSL_Select, 1) Then
              If Not HPrtUpdNyuka Then GoTo DBPut_Err
           End If
        End If
        ' ì¸â◊é¿ê—èàóùÉèÅ[ÉN èoóÕ
        Call ItemsClearZJW
        With ZJW
            .èàóùí[ññ = CisFun.WSID
            .î≠íçä«óùNO = ZCT.î≠íçä«óùNO
        End With
        Call ZJWInsert
        
        Call ZCTReadNext(1)
    Loop
    
    Call ZCTClose(1)
    
    CisDB.ConnectNo = 0
    Call CisDB.DBTran(TransCommit)  ' ƒ◊›ªﬁ∏ºÆ› èIóπ
    
    lb_JMai = Format(CisFun.Val2(lb_JMai) + 1, "#,###")
    
    If Rv_NCall = "K" Then
       Rv_NExec(mBunCount) = True
    End If
    
    
    On Error GoTo 0
    
    DBPut = True
    Exit Function
    
DBPut_Err:
    Call ZCTClose(1)
    
    CisDB.ConnectNo = 0
    Call CisDB.DBTran(TransRollback) ' ƒ◊›ªﬁ∏ºÆ›îjä¸
    CisFun.ErrorBox
    End
End Function
'+-----------------------------------+
'+      ì¸â◊èÓïÒÅ@éÛóÃèë ïœçXÅiï™î[ï™)
'+-----------------------------------+
Private Function HPrtUpdNyuka() As Boolean
    HPrtUpdNyuka = False
    
    With ZCT
        gSL_Select = "Update çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã Set "
        gSL_Select = gSL_Select & " î[ì¸ì˙ = '" & RTrim(ZYT.ì¸â◊ì˙) & "',"
        gSL_Select = gSL_Select & " î[ì¸ï÷ = '" & RTrim(ZYT.ì¸â◊ï÷) & "',"
        gSL_Select = gSL_Select & "î†êî" & .î[ïièëçs & " = " & ZYT.ì¸â◊ñáêî & ","
        gSL_Select = gSL_Select & "í[êî" & .î[ïièëçs & " = " & ZYT.ì¸â◊í[êî & ","
        gSL_Select = gSL_Select & "î[ì¸êî" & .î[ïièëçs & " = " & ZYT.ì¸â◊êî & ""
        
        gSL_Select = gSL_Select & " Where î[ïièëä«óùNO = " & .î[ïièëä«óùNO
        gSL_Select = gSL_Select & "   And ï≈NO = " & .î[ïièëï≈
        gSL_Select = gSL_Select & "   And ÉJÅ[Éh = '6'"
        With CisDB
            .SQL = gSL_Select
            If Not .DBExec Then Exit Function
        End With
        
        gSL_Select = "Update çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã Set "
        gSL_Select = gSL_Select & "î†ëçêî = "
        For gInt = 1 To 10
            gSL_Select = gSL_Select & "Isnull(î†êî" & gInt & ",0)"
            If gInt <> 10 Then
               gSL_Select = gSL_Select & " + "
            End If
        Next gInt
        gSL_Select = gSL_Select & " Where î[ïièëä«óùNO = " & .î[ïièëä«óùNO
        gSL_Select = gSL_Select & "   And ï≈NO = " & .î[ïièëï≈
        gSL_Select = gSL_Select & "   And ÉJÅ[Éh = '6'"
        With CisDB
            .SQL = gSL_Select
            If Not .DBExec Then Exit Function
        End With
    End With

    HPrtUpdNyuka = True

End Function
'ÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅû
'ÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅû
'Åû
'Åû                             ÉOÅ@ÉäÅ@ÉbÅ@ÉhÅ@èàÅ@óùÅ@ä÷Å@òA
'Åû
'ÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅû
'ÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅûÅüÅû
'+==========================================================+
'+      ÉOÉäÉbÉhèàóù                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   å©èoÇµæØƒ & èâä˙èÛë‘æØƒ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'Å@∏ﬁÿØƒﬁèâä˙âªèàóù
'        .Clear
'       Call .FixedSet(å©èoÇµà íuê›íË,óÒà íuê›íË,œ∞ºﬁèàóù,
'                      æŸåÖêî,óÒî‘çÜ,å©èoÇµì‡óe,ï™óﬁ¿≤Ãﬂ)
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 0, "áÇ")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 22, 1, "çﬁ éø / ê° ñ@")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
        'Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 5, 2, "îwî‘")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 8, 2, "îwî‘çÜ")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 11, 3, "é˚óeêî")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 4, 4, "î†êî")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 13, 5, "î≠íçêî")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 2, 6, "é¿")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 7, "äÆî[ó\íËì˙")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 8, "ó\íËï÷")
'Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
'        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0,9,"Select")
'  ï\é¶ ﬂ¿∞›æØƒ
'       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
        Call .PatanSet(0, "0,1,2,3,4,5,6")
'  ï“èWì‡óeÉZÉbÉg
'       Call .EditSet(√∑ΩƒµÃﬁºﬁ™∏ƒ,∂◊—NO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'******************************************'
'*****       ÉOÉäÉbÉhì‡óeÉZÉbÉg        *****
'******************************************'
Private Sub GridSet(SetType As String)
    Dim ID          As Long
    Dim Ttl_Kensu   As Long
' íäèoèàóù èâä˙æØƒ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' íäèoåèêî(çáåv)ëﬁî
' ∏ﬁÿØƒﬁçsçƒæØƒ & óÒà íuéÊìæ
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
            If ZCT.î[ïièëçs > 10 Then Exit Do
            'No
            .TextMatrix(ZCT.î[ïièëçs, 0) = ZCT.î[ïièëçs
            'çﬁéø
            gStr = ZCT.çﬁéø
            If Trim(ZCT.ï\é¶ê°ñ@) <> "" Then
               gStr = gStr & vbCr & ZCT.ï\é¶ê°ñ@
            End If
            .TextMatrix(ZCT.î[ïièëçs, 1) = gStr
            ' îwî‘çÜ
            .TextMatrix(ZCT.î[ïièëçs, 2) = " " & ZCT.îwî‘çÜ
            ' é˚óeêî
            .TextMatrix(ZCT.î[ïièëçs, 3) = Format(ZCT.é˚óeêî, "#,##0.000")
            ' ñáêî
            .TextMatrix(ZCT.î[ïièëçs, 4) = Format(ZCT.ñáêî, "#,###")
            ' î≠íçêî
            .TextMatrix(ZCT.î[ïièëçs, 5) = Format(ZCT.î≠íçêî, "#,##0.000")
            ' é¿êîãÊï™
            If ZCT.é¿êîãÊï™ = 1 Then
               .TextMatrix(ZCT.î[ïièëçs, 6) = "Åõ"
            End If
            
            If SetType = "INV" Or SetType = "QR" Then
               iB1_NSuryo(ZCT.î[ïièëçs - 1) = CisFun.RSetFld(ZCT.î≠íçêî, 0, iB1_NSuryo(0).cFormat)
               iB1_NSuryo(ZCT.î[ïièëçs - 1).Tag = ZCT.î≠íçêî
               iB1_Suryo(ZCT.î[ïièëçs - 1).Tag = ZCT.é¿êîãÊï™
            End If
                        
            If Trim(ZCT.ì¸â◊ì˙) <> "" Then
               iB1_NSuryo(ZCT.î[ïièëçs - 1) = CisFun.RSetFld(ZCT.ì¸â◊êî, 0, iB1_NSuryo(0).cFormat)
               ' ì¸â◊ì˙•ï÷
               iB1_SYmd = ZCT.ì¸â◊ì˙
               If Trim(ZCT.ì¸â◊ï÷) <> "" Then
                  iB1_SBin = ZCT.ì¸â◊ï÷
               End If
               ' ï™äÑî[ì¸êî
               iB1_Suryo(ZCT.î[ïièëçs - 1) = CisFun.RSetFld(ZCT.ï™äÑêî, 0, iB1_Suryo(ID - 1).cFormat)
               ' äÆî[ó\íËì˙•ï÷
               iB1_BYmd(ZCT.î[ïièëçs - 1) = ZCT.î“âÒì˙
               If Trim(ZCT.î“âÒï÷) <> "" Then
                   iB1_BBin(ZCT.î[ïièëçs - 1) = ZCT.î“âÒï÷
               End If
            End If
            
            If Trim(ZCT.î[ì¸ï÷) <> "" Then
               iB1_BBin(ZCT.î[ïièëçs - 1).Visible = True
            Else
               iB1_BBin(ZCT.î[ïièëçs - 1).Visible = False
            End If
        End With
        Call ZCTReadNext
    Loop
    Call ZCTClose
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
           iB1_Biko(gObj.Index).Visible = False
        Else
           gObj.Visible = True
           iB1_NSuryo(gObj.Index).Visible = True
           iB1_BYmd(gObj.Index).Visible = True
           iB1_Biko(gObj.Index).Visible = True
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
'*                      Çp  Çq  Éä  Å[  É_  (Å@ÇbÇèÇçÇçÅ@)  èà  óù
'*
'*=======================================================================================================*
'*=======================================================================================================*
'+--------------------------+
'+      ì«  éÊ  èà  óù      +
'+--------------------------+
Private Sub PDQComm1_OnComm()
    
    gStr = QR_CommEvent
    If gStr = "ER" Then GoTo PDQComm1_OnComm_ED
    If gStr <> "OK" Then Exit Sub
    
    With QRDataArea
        If .ÉoÅ[ÉRÅ[ÉhéÌóﬁ <> "B" Then
            If .ÉXÉeÅ[É^ÉX <> 0 Then GoTo PDQComm1_OnComm_ED
            ' ïîïiÇÃÇ›óLå¯
            If .ÉoÅ[ÉRÅ[ÉhéÌóﬁ <> "Z" Then
               .ÉGÉâÅ[ì‡óe = "çﬁóøÇÃî[ïièëÇ≈Ç»Ç¢"
               GoTo PDQComm1_OnComm_ED
            End If
            ' î[ïièëÇÃÇ›óLå¯
            If .ÉJÅ[Éh <> "7" Then
               .ÉGÉâÅ[ì‡óe = "ÉJÅ[ÉhÇ™ëŒè€äOÇ≈Ç∑"
               Call QRReadMsg("E")
               GoTo PDQComm1_OnComm_ED
            End If
        End If
    End With
'( ÉfÅ[É^ämîF )
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where HC.î[ïièëä«óùNO = " & QRDataArea.î[ïièëä«óùNO
    gSL_Select = gSL_Select & "   And HC.î[ïièëï≈ = " & QRDataArea.î[ïièëï≈
    
    If Not DBInput("QR") Then
       Call QRReadMsg("E")
       Exit Sub
    End If
    ProcHB = "B1"
    Call DispChange(ProcHB)
    Exit Sub
PDQComm1_OnComm_ED:
    Call HeadBodyClear("B")
    Call QRReadMsg("E")
End Sub
'+------------------------------------+
'+      ì«  éÊ  èà  óù(QRŒØ ﬂÿ∞¿ﬁ)     +
'+------------------------------------+
Private Sub PDQComm2_OnComm()

    Call TX2400Event
'   * ÉGÉâÅ[ *
    If TX2400_Event = "ERR" Then
         QRDataArea.ÉGÉâÅ[ì‡óe = TX2400_EventMsg
         Call TX2400QRRReadMsg("E")
         TX2400_ReadFlg = True
         Exit Sub            ' í êMÉGÉâÅ[èàóù
    End If

    If TX2400_Event <> "RCV" Then Exit Sub   ' éÛêM√ﬁ∞¿ï“èW
    If Not TX2400SendProc Then Exit Sub    ' CR Ç™ë∂ç›ÇµÇ»Ç¢éûî≤ÇØÇÈ
    
    ' ÉzÉbÉpÉGÉìÉvÉeÉBÅ[(NG)
    If Mid(TX2400_CommData, 2, 1) = "H" Then
       QRDataArea.ÉGÉâÅ[ì‡óe = "ÉzÉbÉpÉGÉìÉvÉeÉBÅ["
       GoTo PDQComm2_OnComm_ED
    End If
    ' èIóπÉLÅ[
    If Mid(TX2400_CommData, 2, 1) = "E" Then
       Call PB_CAN_Click
       Exit Sub
    End If
    
    Call TX2400QRRReadMsg("R")
    
    QRr_ReadData = TX2400_CommData
    If Not QRReadCheck Then
       GoTo PDQComm2_OnComm_ED
    End If
    
    With QRDataArea
        If .ÉoÅ[ÉRÅ[ÉhéÌóﬁ <> "B" Then
            If .ÉXÉeÅ[É^ÉX <> 0 Then
               GoTo PDQComm2_OnComm_ED
            End If
            ' ïîïiÇÃÇ›óLå¯
            If .ÉoÅ[ÉRÅ[ÉhéÌóﬁ <> "Z" Then
               .ÉGÉâÅ[ì‡óe = "çﬁóøÇÃî[ïièëÇ≈Ç»Ç¢"
               GoTo PDQComm2_OnComm_ED
            End If
            ' î[ïièëÇÃÇ›óLå¯
            If .ÉJÅ[Éh <> "7" Then
               .ÉGÉâÅ[ì‡óe = "ÉJÅ[ÉhÇ™ëŒè€äOÇ≈Ç∑"
               GoTo PDQComm2_OnComm_ED
            End If
        End If
    End With
'( ÉfÅ[É^ämîF )
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where HC.î[ïièëä«óùNO = " & QRDataArea.î[ïièëä«óùNO
    gSL_Select = gSL_Select & "   And HC.î[ïièëï≈ = " & QRDataArea.î[ïièëï≈
    
    If Not DBInput("INV") Then Exit Sub
    
    ProcHB = "B1"
    Call DispChange(ProcHB)
    Exit Sub
PDQComm2_OnComm_ED:
    Call TX2400QRRReadMsg("E")
    TX2400_ReadFlg = True
    Call TX2400Stack(False)
    Call TX2400OpenAuto
End Sub



