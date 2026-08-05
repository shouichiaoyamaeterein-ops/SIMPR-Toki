VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Begin VB.Form CXHZ0910 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "çﬁóø î≠íçÅEì¸â◊ÉfÅ[É^èCê≥"
   ClientHeight    =   11130
   ClientLeft      =   0
   ClientTop       =   390
   ClientWidth     =   15360
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
   ScaleHeight     =   11130
   ScaleWidth      =   15360
   WindowState     =   2  'ç≈ëÂâª
   Begin Cis3D_v60.CIS3D ErrMsg 
      Height          =   360
      Left            =   90
      Top             =   1185
      Visible         =   0   'False
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   635
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
      TabIndex        =   38
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
      Left            =   4170
      ScaleHeight     =   1050
      ScaleWidth      =   2025
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   630
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
            TabIndex        =   3
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
         Caption         =   "èàóùñáêî"
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
      Left            =   90
      ScaleHeight     =   1050
      ScaleWidth      =   3735
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1755
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
            TabIndex        =   1
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
            TabIndex        =   2
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
            TabIndex        =   0
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
      Left            =   9105
      ScaleHeight     =   660
      ScaleWidth      =   6060
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   795
      Width           =   6120
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
         Left            =   4335
         Top             =   15
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
         Left            =   5190
         Top             =   15
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
      Height          =   7680
      Left            =   75
      ScaleHeight     =   7620
      ScaleWidth      =   15120
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3180
      Width           =   15180
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   7440
         Left            =   60
         Top             =   45
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   13123
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
         Begin Cis3D_v60.CIS3D CIS3D13 
            Height          =   6030
            Left            =   6660
            Top             =   1290
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   10636
            BackColor       =   12648384
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
            Begin Cis3D_v60.CIS3D CIS3D18 
               Height          =   450
               Left            =   75
               Top             =   15
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   794
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "î†êî"
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
               Index           =   10
               Left            =   75
               Top             =   465
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_Maisu 
                  Height          =   330
                  Index           =   0
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   450
               Index           =   0
               Left            =   630
               Top             =   15
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   794
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "í[êî"
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
               Index           =   11
               Left            =   630
               Top             =   465
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_HHasu 
                  Height          =   330
                  Index           =   0
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D15 
               Height          =   450
               Left            =   2085
               Top             =   15
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   794
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "î≠íçêî"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   555
               Index           =   0
               Left            =   2085
               Top             =   465
               Width           =   1680
               _ExtentX        =   2963
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
                  TabIndex        =   10
                  Top             =   105
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   555
               Index           =   1
               Left            =   2085
               Top             =   1020
               Width           =   1680
               _ExtentX        =   2963
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
                  TabIndex        =   11
                  Top             =   105
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   555
               Index           =   2
               Left            =   2085
               Top             =   1575
               Width           =   1680
               _ExtentX        =   2963
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
                  TabIndex        =   12
                  Top             =   105
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   555
               Index           =   3
               Left            =   2085
               Top             =   2130
               Width           =   1680
               _ExtentX        =   2963
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
                  TabIndex        =   13
                  Top             =   105
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   555
               Index           =   4
               Left            =   2085
               Top             =   2685
               Width           =   1680
               _ExtentX        =   2963
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
                  TabIndex        =   14
                  Top             =   105
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   555
               Index           =   5
               Left            =   2085
               Top             =   3240
               Width           =   1680
               _ExtentX        =   2963
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
                  TabIndex        =   15
                  Top             =   105
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   555
               Index           =   6
               Left            =   2085
               Top             =   3795
               Width           =   1680
               _ExtentX        =   2963
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
                  TabIndex        =   16
                  Top             =   105
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   555
               Index           =   7
               Left            =   2085
               Top             =   4350
               Width           =   1680
               _ExtentX        =   2963
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
                  TabIndex        =   17
                  Top             =   105
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   555
               Index           =   8
               Left            =   2085
               Top             =   4905
               Width           =   1680
               _ExtentX        =   2963
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
                  Left            =   15
                  TabIndex        =   18
                  Top             =   105
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   555
               Index           =   9
               Left            =   2085
               Top             =   5460
               Width           =   1680
               _ExtentX        =   2963
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
                  TabIndex        =   19
                  Top             =   105
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
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   17
               Left            =   75
               Top             =   1020
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_Maisu 
                  Height          =   330
                  Index           =   1
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   19
               Left            =   630
               Top             =   1020
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_HHasu 
                  Height          =   330
                  Index           =   1
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   20
               Left            =   75
               Top             =   1575
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_Maisu 
                  Height          =   330
                  Index           =   2
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   22
               Left            =   630
               Top             =   1575
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_HHasu 
                  Height          =   330
                  Index           =   2
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   23
               Left            =   75
               Top             =   2130
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_Maisu 
                  Height          =   330
                  Index           =   3
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   25
               Left            =   630
               Top             =   2130
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_HHasu 
                  Height          =   330
                  Index           =   3
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   26
               Left            =   75
               Top             =   2685
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_Maisu 
                  Height          =   330
                  Index           =   4
                  Left            =   5
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   28
               Left            =   630
               Top             =   2685
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_HHasu 
                  Height          =   330
                  Index           =   4
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   29
               Left            =   75
               Top             =   3240
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_Maisu 
                  Height          =   330
                  Index           =   5
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   31
               Left            =   630
               Top             =   3240
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_HHasu 
                  Height          =   330
                  Index           =   5
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   32
               Left            =   75
               Top             =   3795
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_Maisu 
                  Height          =   330
                  Index           =   6
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   33
               Left            =   630
               Top             =   3795
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_HHasu 
                  Height          =   330
                  Index           =   6
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   34
               Left            =   75
               Top             =   4350
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_Maisu 
                  Height          =   330
                  Index           =   7
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   35
               Left            =   630
               Top             =   4350
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_HHasu 
                  Height          =   330
                  Index           =   7
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   36
               Left            =   75
               Top             =   4905
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_Maisu 
                  Height          =   330
                  Index           =   8
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   37
               Left            =   630
               Top             =   4905
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_HHasu 
                  Height          =   330
                  Index           =   8
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   38
               Left            =   75
               Top             =   5460
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_Maisu 
                  Height          =   330
                  Index           =   9
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   39
               Left            =   630
               Top             =   5460
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_HHasu 
                  Height          =   330
                  Index           =   9
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
         End
         Begin Cis3D_v60.CIS3D Back_Mntkb 
            Height          =   1095
            Left            =   150
            Top             =   135
            Width           =   6480
            _ExtentX        =   11430
            _ExtentY        =   1931
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
            Begin Cis3D_v60.CIS3D CIS3D22 
               Height          =   1065
               Left            =   15
               Top             =   15
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   1879
               BackColor       =   0
               ForeColor       =   65535
               Caption         =   "èCê≥ãÊï™"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cPositionX      =   60
               cPositionY      =   60
            End
            Begin VB.OptionButton B1Op_ProcKB 
               BackColor       =   &H00C0C0C0&
               Caption         =   "î≠íçéÊè¡"
               Height          =   315
               Index           =   1
               Left            =   420
               TabIndex        =   5
               Top             =   405
               Width           =   2085
            End
            Begin VB.OptionButton B1Op_ProcKB 
               BackColor       =   &H00C0C0C0&
               Caption         =   "î[ì¸ì˙ èCê≥"
               Height          =   315
               Index           =   0
               Left            =   420
               TabIndex        =   4
               Top             =   60
               Width           =   2085
            End
            Begin VB.Label Back_Bunnou 
               Caption         =   "à»ç~ÇÃï™î[ÉfÅ[É^çÌèúÇ‡çsÇ¢Ç‹Ç∑ÅB"
               Height          =   255
               Left            =   660
               TabIndex        =   39
               Top             =   720
               Visible         =   0   'False
               Width           =   4455
            End
         End
         Begin vsFlexLib.vsFlexArray VsGrid1 
            Height          =   6045
            Left            =   120
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1275
            Width           =   6525
            _Version        =   196608
            _ExtentX        =   11509
            _ExtentY        =   10663
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
            FormatString    =   "áÇ|çﬁéø/ê°ñ@                   |îwî‘çÜ |é˚óeêî |î† êî  |î≠íçêî  |ï™äÑî[ì¸êî|äÆî[ó\íËì˙|ó\íËï÷|"
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
            Index           =   1
            Left            =   12555
            Top             =   75
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
               Height          =   360
               Left            =   1890
               TabIndex        =   9
               Top             =   300
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
            Begin CisYMD_With_Btn.CisYMDwB iB1_SYmd 
               Height          =   345
               Left            =   75
               TabIndex        =   8
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
               Object.Height          =   345
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
            Height          =   735
            Left            =   8085
            Top             =   75
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   1296
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "   î[ ì¸ ì˙  - ï÷"
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
            Begin CisText_V60.CisText iB1_NBin 
               Height          =   360
               Left            =   1890
               TabIndex        =   7
               Top             =   300
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
            Begin CisYMD_With_Btn.CisYMDwB iB1_NYmd 
               Height          =   345
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
               Object.Height          =   345
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
            Begin Cis3D_v60.CIS3D B1lb_NYMD 
               Height          =   360
               Index           =   1
               Left            =   75
               Top             =   300
               Width           =   1785
               _ExtentX        =   3149
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
               Index           =   1
               Left            =   1905
               Top             =   300
               Width           =   360
               _ExtentX        =   635
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
         Begin Cis3D_v60.CIS3D CIS3D16 
            Height          =   6030
            Left            =   10470
            Top             =   1290
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   10636
            BackColor       =   16761024
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
               Index           =   30
               Left            =   3540
               Top             =   5460
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
                  TabIndex        =   29
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
               Left            =   3540
               Top             =   4905
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
                  TabIndex        =   28
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
               Left            =   3540
               Top             =   4350
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
               Index           =   21
               Left            =   3540
               Top             =   3795
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
                  TabIndex        =   26
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
               Index           =   18
               Left            =   3540
               Top             =   3240
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
                  TabIndex        =   25
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
               Index           =   16
               Left            =   3540
               Top             =   2685
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
                  TabIndex        =   24
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
               Index           =   15
               Left            =   3540
               Top             =   2130
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
                  TabIndex        =   23
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
               Index           =   14
               Left            =   3540
               Top             =   1575
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
               Index           =   13
               Left            =   3540
               Top             =   1020
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
                  TabIndex        =   21
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
               Index           =   12
               Left            =   3540
               Top             =   465
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
                  TabIndex        =   20
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
            Begin Cis3D_v60.CIS3D CIS3D21 
               Height          =   450
               Left            =   3540
               Top             =   15
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   794
               BackColor       =   14737632
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
            Begin Cis3D_v60.CIS3D CIS3D17 
               Height          =   450
               Left            =   75
               Top             =   15
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   794
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "î†êî"
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
               Left            =   75
               Top             =   465
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_NMaisu 
                  Height          =   330
                  Index           =   0
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   450
               Index           =   3
               Left            =   630
               Top             =   15
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   794
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "í[êî"
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
               Index           =   1
               Left            =   630
               Top             =   465
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NHasu 
                  Height          =   330
                  Index           =   0
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D19 
               Height          =   450
               Left            =   2085
               Top             =   15
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   794
               BackColor       =   14737632
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
               Index           =   2
               Left            =   2085
               Top             =   465
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NSuryo 
                  Height          =   330
                  Index           =   0
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   3
               Left            =   2085
               Top             =   1020
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NSuryo 
                  Height          =   330
                  Index           =   1
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   4
               Left            =   2085
               Top             =   1575
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NSuryo 
                  Height          =   330
                  Index           =   2
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   5
               Left            =   2085
               Top             =   2130
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NSuryo 
                  Height          =   330
                  Index           =   3
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   6
               Left            =   2085
               Top             =   2685
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NSuryo 
                  Height          =   330
                  Index           =   4
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   7
               Left            =   2085
               Top             =   3240
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NSuryo 
                  Height          =   330
                  Index           =   5
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   8
               Left            =   2085
               Top             =   3795
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NSuryo 
                  Height          =   330
                  Index           =   6
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   9
               Left            =   2085
               Top             =   4350
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NSuryo 
                  Height          =   330
                  Index           =   7
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   60
               Left            =   2085
               Top             =   4905
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NSuryo 
                  Height          =   330
                  Index           =   8
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   61
               Left            =   2085
               Top             =   5460
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NSuryo 
                  Height          =   330
                  Index           =   9
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   62
               Left            =   75
               Top             =   1020
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_NMaisu 
                  Height          =   330
                  Index           =   1
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   63
               Left            =   630
               Top             =   1020
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NHasu 
                  Height          =   330
                  Index           =   1
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   64
               Left            =   75
               Top             =   1575
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_NMaisu 
                  Height          =   330
                  Index           =   2
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   65
               Left            =   630
               Top             =   1575
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NHasu 
                  Height          =   330
                  Index           =   2
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   66
               Left            =   75
               Top             =   2130
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_NMaisu 
                  Height          =   330
                  Index           =   3
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   67
               Left            =   630
               Top             =   2130
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NHasu 
                  Height          =   330
                  Index           =   3
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   68
               Left            =   75
               Top             =   2685
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_NMaisu 
                  Height          =   330
                  Index           =   4
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   69
               Left            =   630
               Top             =   2685
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NHasu 
                  Height          =   330
                  Index           =   4
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   70
               Left            =   75
               Top             =   3240
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_NMaisu 
                  Height          =   330
                  Index           =   5
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   71
               Left            =   630
               Top             =   3240
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NHasu 
                  Height          =   330
                  Index           =   5
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   72
               Left            =   75
               Top             =   3795
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_NMaisu 
                  Height          =   330
                  Index           =   6
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   73
               Left            =   630
               Top             =   3795
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NHasu 
                  Height          =   330
                  Index           =   6
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   74
               Left            =   75
               Top             =   4350
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_NMaisu 
                  Height          =   330
                  Index           =   7
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   75
               Left            =   630
               Top             =   4350
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NHasu 
                  Height          =   330
                  Index           =   7
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   76
               Left            =   75
               Top             =   4905
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_NMaisu 
                  Height          =   330
                  Index           =   8
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   77
               Left            =   630
               Top             =   4905
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NHasu 
                  Height          =   330
                  Index           =   8
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   78
               Left            =   75
               Top             =   5460
               Width           =   555
               _ExtentX        =   979
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
               Begin Cis3D_v60.CIS3D B1lb_NMaisu 
                  Height          =   330
                  Index           =   9
                  Left            =   15
                  Top             =   105
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cPositionX      =   -30
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   555
               Index           =   79
               Left            =   630
               Top             =   5460
               Width           =   1455
               _ExtentX        =   2566
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
               Begin Cis3D_v60.CIS3D B1lb_NHasu 
                  Height          =   330
                  Index           =   9
                  Left            =   15
                  Top             =   105
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   582
                  BackColor       =   16777152
                  Caption         =   "######0.000"
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
                  cPositionX      =   -30
               End
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   360
            Left            =   6660
            Top             =   930
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   635
            BackColor       =   12648384
            Caption         =   "<<  î≠íçèÓïÒ  >>"
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
         Begin Cis3D_v60.CIS3D Back_Nyuka 
            Height          =   360
            Left            =   10470
            Top             =   930
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   635
            BackColor       =   16761024
            Caption         =   "<<  ì¸â◊èÓïÒ  >>"
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
      TabIndex        =   31
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
      Caption         =   "Åy çﬁóø î≠íçÅEì¸â◊ÉfÅ[É^èCê≥  Åz"
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
         Width           =   2595
         _ExtentX        =   4577
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
   Begin VB.PictureBox Picture1 
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
      Left            =   6870
      ScaleHeight     =   1050
      ScaleWidth      =   8295
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1740
      Width           =   8355
      Begin Cis3D_v60.CIS3D Back_Nounox 
         Height          =   960
         Left            =   6945
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
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
            Left            =   45
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
      Begin Cis3D_v60.CIS3D Back_Page 
         Height          =   960
         Left            =   7800
         Top             =   60
         Width           =   435
         _ExtentX        =   767
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
            Width           =   270
            _ExtentX        =   476
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
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   960
         Left            =   3615
         Top             =   60
         Width           =   1800
         _ExtentX        =   3175
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
         Begin Cis3D_v60.CIS3D B1lb_NYMD 
            Height          =   360
            Index           =   0
            Left            =   45
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
            Index           =   0
            Left            =   1410
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
         Left            =   5415
         Top             =   60
         Width           =   750
         _ExtentX        =   1323
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
            Width           =   570
            _ExtentX        =   1005
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
         Left            =   6165
         Top             =   60
         Width           =   780
         _ExtentX        =   1376
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
            Left            =   90
            Top             =   390
            Width           =   615
            _ExtentX        =   1085
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
   Begin Cis3D_v60.CIS3D H3_Area1 
      Height          =   1110
      Left            =   270
      Top             =   480
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   1958
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
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   0
         Width           =   240
      End
      Begin Cis3D_v60.CIS3D QRMsg 
         Height          =   945
         Left            =   90
         Top             =   90
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   1667
         BackColor       =   12648384
         ForeColor       =   16711680
         Caption         =   "QRÉäÅ[É_Å[ì«éÊ"
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
End
Attribute VB_Name = "CXHZ0910"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   î≠íçÅEì¸â◊ÉfÅ[É^èCê≥
'**       ÉtÉHÅ[ÉÄID    :   CXHZ0910
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2004/02/12  By CIS
'**       ïœ  çX  ì˙    :   2006/01/26  By CIS  î[ì¸ì˙ïœçXéûÅAÇPçsñ⁄ÇµÇ©ïœçXÇ≥ÇÍÇ»Ç¢( MaisuGetì‡Ç≈åãâ ÉZÉbÉgÇ™ïœÇÌÇ¡ÇƒÇµÇ‹Ç§ )
'**       ïœ  çX  ì˙    :   2009/01/19  By CIS  íºëóópÅAçﬁóøé¿ê—ä«óùÉeÅ[ÉuÉãÇÃçÌèú/çXêVÇí«â¡
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

    Dim mHachuDel               As Boolean
    
    Dim mCurren                 As Currency

    Dim mSuryoStr               As String
    Dim mSyuyoStr               As String
    
    Dim mMaisu                  As Currency
    Dim mHasu                   As Currency
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås    ÇkÇèÇÅÇÑÉCÉxÉìÉg    Åt            +
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
    End If
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    QRDummy.Left = -1000
    
    H2_Area1.Move H1_Area1.Left, H1_Area1.Top
    H3_Area1.Move H1_Area1.Left, H1_Area1.Top
    
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
    B1Op_ProcKB(0).Caption = "î≠íçèÓïÒèCê≥"
    B1Op_ProcKB(1).Caption = "î≠íçèÓïÒçÌèú"
    
'   #------------------#
'   #    èâä˙ì¸óÕê›íË   #
'   #------------------#
    K_Sykbnm.Tag = "3"
    Call PB_CHG_Click
    
'   #-------------------#
'   # âÊñ íÜâõà íuï\é¶   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    If Rv_NCall = "" Then
       Me.Show
    End If
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    ÅsÇ`ÇÉÇîÇâÇñÇÅÇîÇÖ ÉCÉxÉìÉgÅt                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = Not FormAct

    mSL_Select = "SELECT HC.î[ïièëä«óùNO,HC.èââÒî≠íçä«óùNO,HC.î[ïièëçs,HC.î[ì¸ì˙,HC.î[ì¸ï÷,"
    mSL_Select = mSL_Select & "  ISNULL(NK.ì¸â◊ì˙,'') ì¸â◊ì˙,ISNULL(NK.ì¸â◊ï÷,'') ì¸â◊ï÷,"
    mSL_Select = mSL_Select & "  ISNULL(NK.î“âÒì˙,'') î“âÒì˙,ISNULL(NK.î“âÒï÷,'') î“âÒï÷,"
    mSL_Select = mSL_Select & "  ISNULL(NK.ï™äÑêî,0) ï™äÑêî,ISNULL(NK.ì¸â◊îıçl,'') ì¸â◊îıçl,"
    mSL_Select = mSL_Select & "  HC.édì¸êÊ,HC.éÛì¸,HC.î[ì¸î‘çÜ,HC.î[ïièëï≈,"
    mSL_Select = mSL_Select & "  HC.î≠íçãÊï™,HC.éËîzãÊï™,HC.SKãÊï™,"
    mSL_Select = mSL_Select & "  ISNULL(TR.ó™èÃ,'') édì¸êÊñº,"
    mSL_Select = mSL_Select & "  ISNULL(SY1.ílñºèÃ,'') î≠íçãÊï™ñº,"
    mSL_Select = mSL_Select & "  ISNULL(SY2.ílñºèÃ,'') éËîzãÊï™ñº,"
    mSL_Select = mSL_Select & "  HN.ï\é¶ê°ñ@,HC.îwî‘çÜ,"
    mSL_Select = mSL_Select & "  HC.é˚óeêî,HC.í[êî,HC.ñáêî,HC.î≠íçêî,"
    mSL_Select = mSL_Select & "  NK.ì¸â◊ñáêî,NK.ì¸â◊í[êî,NK.ì¸â◊êî,HN.é¿êîãÊï™,HC.çﬁóøä«óùî‘çÜ, "
    mSL_Select = mSL_Select & "  HC.çﬁéø,HC.î¬å˙,HC.ïù,HC.í∑Ç≥,HN.ê°ñ@ãÊï™ "
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
       K_Sykbnm.Tag = "3"
       Call PB_CHG_Click
       mBunCount = 1
       
       Call ContinuChk
       
    End If

End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås QueryUnloadÉCÉxÉìÉg Åt                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
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
                                If PB_ENT.Visible Then                      'Åyì¸óÕÅz
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
    If gLong > 3 Then gLong = 1
    
    mSaveHeadProc = ""
    mSaveHeadProc = "H1"
    If IRN_NKReadBar = 1 Then
       mSaveHeadProc = mSaveHeadProc & "H2"
    Else
       mSaveHeadProc = mSaveHeadProc & "  "
    End If
    If IRN_NKRead = 1 Then
       mSaveHeadProc = mSaveHeadProc & "H3"
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
    
    Select Case mHeadProc
        Case "H2"
            K_Sykbnm.Caption = "Åy  ﬁ∞∫∞ƒﬁ Åz"
            H1_Area1.Visible = False
            H2_Area1.Visible = True
            H3_Area1.Visible = False
        Case "H3"
            K_Sykbnm.Caption = "Åy QRÿ∞¿ﬁ∞ Åz"
            H1_Area1.Visible = False
            H2_Area1.Visible = False
            H3_Area1.Visible = True
            '   ( QR Reader ä¬ã´ )
            Set QRr_Object = PDQComm1
            Set QRr_MsgObject = QRMsg
            Set QRr_ErrMsgObj = ErrMsg
            Call QRInit
        Case Else
            K_Sykbnm.Caption = "Åy î[î‘ Åz"
            H1_Area1.Visible = True
            H2_Area1.Visible = False
            H3_Area1.Visible = False
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
    
    Call DBInput("INV")
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
    
    Call DBInput("INV")
PB_Next_ED:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_NYmd" Then iH1_NYmd.ShowCalender
    If PB_Look.Tag = "iB1_NYmd" Then iB1_NYmd.ShowCalender
    If PB_Look.Tag = "iB1_SYmd" Then iB1_SYmd.ShowCalender
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
'*      î[ ì¸ ì˙  ( B1Op_ProcKB )       *
'*----------------------------------------------*
Private Sub B1Op_ProcKB_Click(Index As Integer)
    If Index = 0 Then
       If iB1_SYmd.Visible Then
          B1Op_ProcKB(0).BackColor = Back_Nyuka.BackColor
          B1Op_ProcKB(1).BackColor = Back_Mntkb.BackColor
       Else
          B1Op_ProcKB(0).BackColor = Back_Hachu.BackColor
          B1Op_ProcKB(1).BackColor = Back_Mntkb.BackColor
       End If
    Else
       If iB1_SYmd.Visible Then
          B1Op_ProcKB(1).BackColor = Back_Nyuka.BackColor
          B1Op_ProcKB(0).BackColor = Back_Mntkb.BackColor
       Else
          B1Op_ProcKB(1).BackColor = Back_Hachu.BackColor
          B1Op_ProcKB(0).BackColor = Back_Mntkb.BackColor
       End If
    End If
End Sub
'*----------------------------------------------*
'*      î[ ì¸ ì˙  ( iB1_NYmd )       *
'*----------------------------------------------*
Private Sub iB1_NYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_NYmd_LostFocus()
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
'*       î≠íçêî  ( iB1_Suryo )                  *
'*----------------------------------------------*
Private Sub iB1_Suryo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim ID  As Integer
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
    
    If CisFun.Val2(VsGrid1.TextMatrix(Index + 1, 3)) <> 0 Then
        mSyuyoStr = VsGrid1.TextMatrix(Index + 1, 3)
        mSuryoStr = Format(CisFun.Val2(iB1_Suryo(Index)), "0.000")
        
        Call MaisuGet
        
        B1lb_Maisu(Index) = CisFun.RSetFld(mMaisu, 0, "###")
        If mHasu <> 0 Then
           B1lb_HHasu(Index) = CisFun.RSetFld(mHasu, 0, iB1_Suryo(Index).cFormat)
        Else
           B1lb_HHasu(Index) = ""
        End If
    End If
End Sub

Private Sub MaisuGet()
    mMaisu = 0
    mHasu = 0
    
    If CisFun.Val2(mSyuyoStr) = 0 Then Exit Sub
    
    With CisDB
        .SQL = "Select Case When " & mSuryoStr & " - " & vbCr
        .SQL = .SQL & "( " & mSyuyoStr & " * ( Cast(ROUND(" & mSuryoStr & "/" & mSyuyoStr & ",0,1) As Int))) <> 0 Then " & vbCr
        .SQL = .SQL & " Cast(Round(" & mSuryoStr & " / " & mSyuyoStr & ",0,1) as Int ) + 1 " & vbCr
        .SQL = .SQL & " Else Cast(Round(" & mSuryoStr & " / " & mSyuyoStr & ",0,1) as Int ) End ñáêî," & vbCr
        .SQL = .SQL & " Case When " & mSuryoStr & " < " & mSyuyoStr & " Then " & mSuryoStr & vbCr
        .SQL = .SQL & "      When " & mSuryoStr & " = " & mSyuyoStr & " Then 0 " & vbCr
        .SQL = .SQL & "      When " & mSuryoStr & " > " & mSyuyoStr & " Then " & vbCr
        .SQL = .SQL & mSuryoStr & " - ( " & mSyuyoStr & " * ( Cast(ROUND(" & mSuryoStr & "/" & mSyuyoStr & ",0,1) As Int))) " & vbCr
        .SQL = .SQL & "End í[êî "
                
'< 2006 / 01 / 26   MNT START >-----------------------------------------------------------------------------------------------------------------------
'
'        If .DBRead(1) Then
        If .DBRead(1, 1) Then
            If Not IsNull(.Recordset("ñáêî")) Then mMaisu = .Recordset("ñáêî")
            If Not IsNull(.Recordset("í[êî")) Then mHasu = .Recordset("í[êî")
            
        End If
'
'< 2006 / 01 / 26   MNT  END  >-----------------------------------------------------------------------------------------------------------------------

    End With
End Sub


'*----------------------------------------------*
'*       îıçl  ( iB1_Biko )                  *
'*----------------------------------------------*
Private Sub iB1_Biko_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim ID  As Integer
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
    Dim B1Mode          As Boolean
    Dim B1Color         As Long
    
    H1Mode = False:     H1Color = gPLostSel
    H2Mode = False:     H2Color = gPLostSel
    H3Mode = False:     H3Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ¿ﬁ–∞çÄñ⁄(âÊñ ÇÃÉ`ÉâÉcÉLñhé~)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "H2":      H2Mode = True:  H2Color = gPGotSel
        Case "H3":      H3Mode = True:  H3Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    H2_Area1.Enabled = H2Mode
    H2_Area1.BackColor = H2Color
    H3_Area1.Enabled = H3Mode
    H3_Area1.BackColor = H3Color
    
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
    If ProcHB = "H3" Then
       Call QR_Comm_Open
    Else
       Call QR_Comm_Close
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
       B1Op_ProcKB(0).BackColor = Back_Mntkb.BackColor
       B1Op_ProcKB(1).BackColor = Back_Mntkb.BackColor
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
    
    gSL_Select = "Select î[ïièëä«óùNO From çﬁóøî≠íçÉeÅ[ÉuÉã "
    gSL_Select = gSL_Select & " Where î[ì¸î‘çÜ = '" & RTrim$(iH1_Nouno) & "'"
    gSL_Select = gSL_Select & " And   î[ïièëï≈ =  " & CisFun.Val2(iH1_Page)
    gSL_Select = gSL_Select & " And   î[ì¸ì˙   = '" & RTrim$(iH1_NYmd) & "'"
    gSL_Select = gSL_Select & " Group By î[ïièëä«óùNO"
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
       gSL_Select = gSL_Select & " Where HC.î[ì¸î‘çÜ = '" & RTrim$(iH1_Nouno) & "'"
       gSL_Select = gSL_Select & " And   HC.î[ïièëï≈ =  " & CisFun.Val2(iH1_Page)
       gSL_Select = gSL_Select & " And   HC.î[ì¸ì˙   = '" & RTrim$(iH1_NYmd) & "'"
       gSL_Select = gSL_Select & " Order By HC.çﬁéø,HN.ï\é¶ê°ñ@"
       Call ZCTRead(gSL_Select)
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
    gSL_Select = gSL_Select & " Order By HC.î[ïièëçs"
    
    If Not DBInput("INV") Then Exit Function
    
    Head2Chk = True
End Function
'****************************
'*      Body1¡™Ø∏èàóù
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    If iB1_SYmd.Visible Then
       If Not ChkNyuka Then Exit Function
    Else
       If Not ChkHachu Then Exit Function
    End If
    Body1Chk = True
End Function
'+----------------------------+
'+      ñ¢ì¸â◊èÓïÒÉ`ÉFÉbÉN
'+----------------------------+
Private Function ChkHachu() As Boolean
    ChkHachu = False
    
    If B1Op_ProcKB(1).Value Then
       ChkHachu = True
       Exit Function
    End If
    
    mHachuDel = False
    For gInt = 0 To 9
        If CisFun.Val2(iB1_Suryo(gInt)) <> 0 Then
           mHachuDel = True
           Exit For
        End If
    Next gInt
    
    If Not mHachuDel Then
       B1Op_ProcKB(1).Value = True
       ChkHachu = True
       Exit Function
    End If
    
    If iB1_NYmd = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    î[ì¸ì˙Çì¸óÕÇµÇƒâ∫Ç≥Ç¢         "
            .MB_Button = Error
            .MBOX
        End With
        iB1_NYmd.SetFocus
        Exit Function
    End If
    If Not iB1_NYmd.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    î[ì¸ì˙Çê≥ÇµÇ¢ì˙ïtÇ…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢         "
            .MB_Button = Error
            .MBOX
        End With
        iB1_NYmd.SetFocus
        Exit Function
    End If
    gInt = Kdou_Check(iB1_NYmd)
    If gInt = 1 Then
    'ãxì˙
        CisFun.MB_Lines = 4
        CisFun.MB_MSG(1) = "     î[ì¸ì˙Ç™îÒâ“ì≠ì˙Ç≈Ç∑ÅB     "
        CisFun.MB_MSG(3) = "     Å@ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH      "
        CisFun.MB_Button = OK_CAN
        If Not CisFun.MBOX Then
            iB1_NYmd.SetFocus
            Exit Function
        End If
    End If
    If gInt = 2 Then
    'ÉJÉåÉìÉ_ñ¢ìoò^
        CisFun.MB_Lines = 4
        CisFun.MB_MSG(1) = "     " & RTrim(iB1_NYmd.cYear) & "îN" & RTrim(iB1_NYmd.cMonth) & "åéÇÃÉJÉåÉìÉ_ñ¢ìoò^ÅI     "
        CisFun.MB_MSG(3) = "     Å@Å@(Å@ÉJÉåÉìÉ_É}ÉXÉ^Å@)Å@"
        CisFun.MB_Button = Error
        CisFun.MBOX
        iB1_NYmd.SetFocus
        Exit Function
    End If
    
    If iB1_NBin.Visible Then
        If Trim(B1lb_HKbn.Tag) = "1" And CisFun.Val2(iB1_NBin) = 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(1) = "    ÅyÅ@í èÌî≠íçéûÅAï÷ ïKê{ÅI Åz     "
                .MB_MSG(3) = "       î[ì¸ï÷Çì¸óÕÇµÇƒâ∫Ç≥Ç¢        "
                .MB_Button = Error
                .MBOX
            End With
            iB1_NBin.SetFocus
            Exit Function
        End If
    End If
    
    ChkHachu = True
End Function
'+----------------------------+
'+      ì¸â◊èÓïÒÉ`ÉFÉbÉN
'+----------------------------+
Private Function ChkNyuka() As Boolean
    ChkNyuka = False
    
    If B1Op_ProcKB(1).Value Then
       ChkNyuka = True
       Exit Function
    End If
    
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
    gInt = Kdou_Check(iB1_SYmd)
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
    
    If iB1_SBin.Visible Then
        If Trim(B1lb_HKbn.Tag) = "1" And CisFun.Val2(iB1_SBin) = 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(1) = "    ÅyÅ@í èÌî≠íçéûÅAï÷ ïKê{ÅI Åz     "
                .MB_MSG(3) = "       ì¸â◊ï÷Çì¸óÕÇµÇƒâ∫Ç≥Ç¢        "
                .MB_Button = Error
                .MBOX
            End With
            iB1_SBin.SetFocus
            Exit Function
        End If
    End If
    
    ChkNyuka = True
End Function
'+---------------------------------------+
'+  àÍäáÇ©ÇÁÇÃòAågèàóù
'+---------------------------------------+
Private Function ContinuChk() As Boolean
'    ContinuChk = False
'    If mBunCount > Rv_NTtlSu Then Exit Function
'    Do Until Rv_HNo(mBunCount) <> 0
'       If mBunCount >= Rv_NTtlSu Then
'          Call PB_END_Click
'          Exit Function
'       End If
'       mBunCount = mBunCount + 1
'    Loop
'
'    mBun_Count = CisFun.RSetFld(mBunCount, 0, "###") & " / " & _
'                CisFun.RSetFld(Rv_NTtlSu, 0, "###")
'    gSL_Select = "Select î[ì¸î‘çÜ,î[ì¸ì˙,î[ïièëï≈ "
'    gSL_Select = gSL_Select & " From î≠íçÉeÅ[ÉuÉã "
'    gSL_Select = gSL_Select & " Where î[ïièëä«óùNO = " & Rv_HNo(mBunCount)
'    gSL_Select = gSL_Select & "   And î[ïièëï≈ = " & Rv_HPage(mBunCount)
'    If Not ZCTRead(gSL_Select, 1) Then
'       Call PB_END_Click
'    End If
'
'    iH1_NYmd = ZCT.î[ì¸ì˙
'    iH1_Nouno = ZCT.î[ì¸î‘çÜ
'    iH1_Page = CisFun.RSetFld(ZCT.î[ïièëï≈, 0, iH1_Page.cFormat)
'    ProcHB = "H1"
'    Call ReturnPress
'
'    ContinuChk = True

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
        GoTo DBInput_Ed
    Else
        Call BodySet("")
        
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
        iB1_NYmd = .î[ì¸ì˙
        If Trim(.ì¸â◊ì˙) = "" Then
           B1lb_NYMD(1).Visible = False
           B1lb_NBin(1).Visible = False
           iB1_NYmd.Visible = True
           iB1_SYmd.Visible = False
           iB1_SBin.Visible = False
           If Trim(ZCT.î[ì¸ï÷) <> "" Then
              iB1_NBin.Visible = True
              iB1_NBin = .î[ì¸ï÷
           Else
              iB1_NBin.Visible = False
           End If
           B1Op_ProcKB(0).Caption = "î≠íçèÓïÒèCê≥"
           B1Op_ProcKB(1).Caption = "î≠íçèÓïÒçÌèú"
           Back_Bunnou.Visible = False
        Else
           B1lb_NYMD(1).Visible = True
           B1lb_NBin(1).Visible = True
           iB1_NYmd.Visible = False
           iB1_NBin.Visible = False
           iB1_SYmd = .ì¸â◊ì˙
           iB1_SYmd.Visible = True
           If Trim(ZCT.ì¸â◊ï÷) <> "" Then
              iB1_SBin.Visible = True
              iB1_SBin = .ì¸â◊ï÷
           Else
              iB1_SBin.Visible = False
           End If
           B1Op_ProcKB(0).Caption = "ì¸â◊ì˙èÓïÒèCê≥"
           B1Op_ProcKB(1).Caption = "ì¸â◊èÓïÒçÌèú"
           Back_Bunnou.Visible = True
        End If
        B1Op_ProcKB(0).Value = True
        
        If Trim(.édì¸êÊñº) = "" Then
           gStr = .édì¸êÊ
        Else
           gStr = RTrim(.édì¸êÊ) & ":" & .édì¸êÊñº
        End If
        B1lb_Tornm = gStr
        B1lb_NYMD(0).Tag = ZCT.î[ì¸ì˙
        With CisFun
            .DateE_BefVal = ZCT.î[ì¸ì˙
            .DateEdit
            B1lb_NYMD(0) = .DateE_AfVal
            .DateEdit J_YMD
            B1lb_NYMD(1) = .DateE_AfVal
            If Trim(ZCT.î[ì¸ï÷) <> "" Then
                B1lb_NBin(0) = ZCT.î[ì¸ï÷
            End If
        End With
        B1lb_NBin(1) = B1lb_NBin(0)
        
        B1lb_HKbn = .î≠íçãÊï™ñº
        B1lb_HKbn.Tag = .î≠íçãÊï™
        B1lb_TKbn = .éËîzãÊï™ñº
    End With
    
    Call GridSet
    
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
        If iB1_SYmd.Visible Then
           If B1Op_ProcKB(1).Value Then
              .MB_MSG(2) = "     ì¸â◊èÓïÒçÌèúèàóùÇçsÇ¢Ç‹Ç∑             "
           Else
              .MB_MSG(2) = "     ì¸â◊èÓïÒèCê≥èàóùÇçsÇ¢Ç‹Ç∑             "
           End If
        Else
           If B1Op_ProcKB(1).Value Then
              .MB_MSG(2) = "     î≠íçèÓïÒçÌèúèàóùÇçsÇ¢Ç‹Ç∑             "
           Else
              .MB_MSG(2) = "     î≠íçèÓïÒèCê≥èàóùÇçsÇ¢Ç‹Ç∑             "
           End If
        End If
        .MB_MSG(4) = "     ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH      "
        .MB_Title = "çXêVèàóù"
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
    
    gSL_Select = "SELECT HC.*,"
    gSL_Select = gSL_Select & "  ISNULL(NK.ì¸â◊ì˙,'') ì¸â◊ì˙,ISNULL(NK.ì¸â◊ï÷,'') ì¸â◊ï÷,"
    gSL_Select = gSL_Select & "  NK.ì¸â◊ñáêî,NK.ì¸â◊í[êî,NK.ì¸â◊êî,Isnull(NK.ì¸â◊îıçl,'') ì¸â◊îıçl "
    gSL_Select = gSL_Select & " ,NK.éxããé¿ê—ä«óùNO "                    '2009/01/19
    gSL_Select = gSL_Select & " From çﬁóøî≠íçÉeÅ[ÉuÉã HC"
    gSL_Select = gSL_Select & "  Left Outer Join çﬁóøì¸â◊ÉeÅ[ÉuÉã NK"
    gSL_Select = gSL_Select & "     On NK.î≠íçä«óùNO = HC.î≠íçä«óùNO"
    gSL_Select = gSL_Select & " Where HC.î[ïièëä«óùNO = " & CisFun.Val2(iB1_SYmd.Tag)
    gSL_Select = gSL_Select & "   And HC.î[ïièëï≈ = " & CisFun.Val2(iB1_SBin.Tag)
    
    
    If Not ZCTRead(gSL_Select) Then
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
    
    If iB1_SYmd.Visible Then
        If Not DBNyukaUPD Then GoTo DBPut_Err
    Else
        If Not DBHachuUPD Then GoTo DBPut_Err
    End If
'
    Call CisDB.DBTran(TransCommit)  ' ƒ◊›ªﬁ∏ºÆ› èIóπ
    
    lb_JMai = Format(CisFun.Val2(lb_JMai) + 1, "#,###")
    
    If Rv_NCall = "K" Then
       Rv_NExec(mBunCount) = True
    End If
    
    
    On Error GoTo 0
    
    DBPut = True
    Exit Function
    
DBPut_Err:
    Call ZCTClose
    
    Call CisDB.DBTran(TransRollback) ' ƒ◊›ªﬁ∏ºÆ›îjä¸
    CisFun.ErrorBox
End Function
'+----------------------------------------------+
'+      î≠íçèÓïÒèCê≥Å@çXêV
'+----------------------------------------------+
Private Function DBHachuUPD() As Boolean
    DBHachuUPD = False
    Dim wHSuryo         As Currency
    
    With ZCT
        If B1Op_ProcKB(1).Value Then        ' î≠íçèÓïÒçÌèú
           '+ î≠íçí†ï[î≠çsÉeÅ[ÉuÉã çÌèú
           If Not HPrtDelHachu(0) Then Exit Function
           
           gSL_Select = "Delete çﬁóøî≠íçÉeÅ[ÉuÉã  "
           gSL_Select = gSL_Select & " Where î[ïièëä«óùNO = " & .î[ïièëä«óùNO
           gSL_Select = gSL_Select & "   And î[ïièëï≈ = " & .î[ïièëï≈
           With CisDB
                .SQL = gSL_Select
                If Not .DBExec Then Exit Function
           End With
           DBHachuUPD = True
           Call ZCTClose
           Exit Function
        End If
    End With
    
    Do Until Not ZCT_RDSTS
        
        With ZCT
            wHSuryo = CisFun.Val2(iB1_Suryo(.î[ïièëçs - 1))
            If .î[ì¸ì˙ <> iB1_NYmd Or .î[ì¸ï÷ <> iB1_NBin Or _
               .î≠íçêî <> wHSuryo Then
                
               If wHSuryo = 0 Then
                  '+ î≠íçí†ï[î≠çsÉeÅ[ÉuÉã çÌèú
                  If Not HPrtDelHachu(1) Then Exit Function
                  
                  gSL_Select = "Delete çﬁóøî≠íçÉeÅ[ÉuÉã  "
                  gSL_Select = gSL_Select & " Where î≠íçä«óùNO = " & .î≠íçä«óùNO & ""
                  With CisDB
                      .SQL = gSL_Select
                      If Not .DBExec Then Exit Function
                  End With
               Else
                  .î[ì¸ì˙ = iB1_NYmd
                  If iB1_NBin.Visible Then
                     If CisFun.Val2(iB1_NBin) <> 0 Then
                        .î[ì¸ï÷ = Format(iB1_NBin, "00")
                     End If
                  End If
                    
                  gCnt1 = 1
                  .í[êî = 0
                  mSyuyoStr = Format(CisFun.Val2(VsGrid1.TextMatrix(.î[ïièëçs, 3)), "0.000")
                  mSuryoStr = Format(wHSuryo, "0.000")
                    
                  Call MaisuGet
                                  
                  .í[êî = mHasu
                  .ñáêî = mMaisu
                  .î≠íçêî = wHSuryo
                    
                  gSL_Select = "Update çﬁóøî≠íçÉeÅ[ÉuÉã Set "
                  gSL_Select = gSL_Select & "î[ì¸ì˙ = '" & RTrim(.î[ì¸ì˙) & "',"
                  gSL_Select = gSL_Select & "î[ì¸ï÷ = '" & RTrim(.î[ì¸ï÷) & "',"
                  gSL_Select = gSL_Select & "ñáêî = " & .ñáêî & ","
                  gSL_Select = gSL_Select & "í[êî = " & .í[êî & ","
                  gSL_Select = gSL_Select & "î≠íçêî = " & .î≠íçêî & ","
                  gSL_Select = gSL_Select & "çXêVì˙ = GETDATE(),"
                  gSL_Select = gSL_Select & "çXêVí[ññ = SubString(Host_Name(),1,20),"
                  gSL_Select = gSL_Select & "çXêVé“ = '" & RTrim(.çXêVé“) & "'"
                  gSL_Select = gSL_Select & " Where î≠íçä«óùNO = " & .î≠íçä«óùNO & ""
               
                  With CisDB
                      .SQL = gSL_Select
                      If Not .DBExec Then Exit Function
                  End With
                  '+ î≠íçí†ï[î≠çsÉeÅ[ÉuÉã çXêV
                  If Not HPrtUpdHachu Then Exit Function
               End If
            End If
        End With
        
        Call ZCTReadNext
    Loop

    Call ZCTClose
    
    DBHachuUPD = True
End Function
'+----------------------------------------------+
'+      ì¸â◊èÓïÒèCê≥Å@çXêV
'+----------------------------------------------+
Private Function DBNyukaUPD() As Boolean
    DBNyukaUPD = False
    Dim wHSuryo         As Long
    Dim wNouno(10)      As String
    Dim wFirstNo(10)    As Long
    Dim wNouKNo(10)     As Long
    Dim wPageNo(10)     As Byte
    
    If B1Op_ProcKB(1).Value Then        ' ì¸â◊èÓïÒçÌèú
        
       Do Until Not ZCT_RDSTS
            
           With ZCT
                wNouno(.î[ïièëçs) = .î[ì¸î‘çÜ
                wFirstNo(.î[ïièëçs) = .èââÒî≠íçä«óùNO
                wNouKNo(.î[ïièëçs) = .î[ïièëä«óùNO
                wPageNo(.î[ïièëçs) = .î[ïièëï≈
            '+ çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã çÌèú
                If Not HPrtDelNyuka Then Exit Function
       
                '==================================== 2009/01/19 Start
                'é©ìÆçÏê¨ÇÃé¿ê—ä«óùÉfÅ[É^ÇçÌèú
                gSL_Select = "Delete çﬁóøé¿ê—ä«óùÉeÅ[ÉuÉã  From çﬁóøé¿ê—ä«óùÉeÅ[ÉuÉã NY "
                gSL_Select = gSL_Select & " Left Outer Join çﬁóøì¸â◊ÉeÅ[ÉuÉã HC "
                gSL_Select = gSL_Select & "   On HC.éxããé¿ê—ä«óùNO = NY.é¿ê—ä«óùNO "
                gSL_Select = gSL_Select & " Where HC.éxããé¿ê—ä«óùNO = " & .éxããé¿ê—ä«óùNO & ""
                gSL_Select = gSL_Select & " And NY.ì¸óÕéÌóﬁ = 'A' "
                gSL_Select = gSL_Select & " And NY.é¿ê—ãÊï™ = 3 "       'íºëó
                With CisDB
                     .SQL = gSL_Select
                     If Not .DBExec Then Exit Function
                End With
                '==================================== 2009/01/19 End
                
                gSL_Select = "Delete çﬁóøì¸â◊ÉeÅ[ÉuÉã  From çﬁóøì¸â◊ÉeÅ[ÉuÉã NY "
                gSL_Select = gSL_Select & " Left Outer Join çﬁóøî≠íçÉeÅ[ÉuÉã HC "
                gSL_Select = gSL_Select & "   On HC.î≠íçä«óùNO = NY.î≠íçä«óùNO "
                gSL_Select = gSL_Select & " Where HC.èââÒî≠íçä«óùNO = " & .èââÒî≠íçä«óùNO & ""
                gSL_Select = gSL_Select & "   And HC.î[ì¸î‘çÜ >= '" & RTrim(.î[ì¸î‘çÜ) & "'"
                With CisDB
                     .SQL = gSL_Select
                     If Not .DBExec Then Exit Function
                End With
             
                gSL_Select = "Delete çﬁóøî≠íçÉeÅ[ÉuÉã  "
                gSL_Select = gSL_Select & " Where èââÒî≠íçä«óùNO = " & .èââÒî≠íçä«óùNO & ""
                gSL_Select = gSL_Select & "   And î[ì¸î‘çÜ > '" & RTrim(.î[ì¸î‘çÜ) & "'"
                With CisDB
                     .SQL = gSL_Select
                     If Not .DBExec Then Exit Function
                End With
             
            End With
            Call ZCTReadNext
       Loop
        
       For gCnt1 = 1 To 10
           If Trim(wNouno(gCnt1)) <> "" Then
              gSL_Select = "Select * From  çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã "
              gSL_Select = gSL_Select & " Where î[ïièëä«óùNO = " & wNouKNo(gCnt1)
              gSL_Select = gSL_Select & "   And î[ì¸î‘çÜ = '" & RTrim(wNouno(gCnt1)) & "'"
              gSL_Select = gSL_Select & "   And ï≈NO = " & wPageNo(gCnt1)
              gSL_Select = gSL_Select & "   And ÉJÅ[Éh = '6'"
              gSL_Select = gSL_Select & "   And î≠çsãÊï™ = 1"
              If ZCHRead(gSL_Select, 1) Then
                 gSL_Select = "Select * From  çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã "
                 gSL_Select = gSL_Select & " Where î[ïièëä«óùNO = " & wNouKNo(gCnt1)
                 gSL_Select = gSL_Select & "   And î[ì¸î‘çÜ = '" & RTrim(wNouno(gCnt1)) & "'"
                 gSL_Select = gSL_Select & "   And ï≈NO = " & wPageNo(gCnt1)
                 gSL_Select = gSL_Select & "   And ÉJÅ[Éh = '7'"
                 If ZCHRead(gSL_Select, 1) Then
                    With CisDB
                         .SQL = "Delete çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã "
                         .SQL = .SQL & " Where î[ïièëä«óùNO = " & wNouKNo(gCnt1)
                         .SQL = .SQL & "   And î[ì¸î‘çÜ = '" & RTrim(wNouno(gCnt1)) & "'"
                         .SQL = .SQL & "   And ï≈NO = " & wPageNo(gCnt1)
                         .SQL = .SQL & "   And ÉJÅ[Éh = '6'"
                         .SQL = .SQL & "   And î≠çsãÊï™ = 1"
                         If Not .DBExec Then Exit Function
                    End With
                    ZCH.ÉJÅ[Éh = "6"
                    If Not ZCHInsert Then Exit Function
                 End If
              End If
           End If
       Next gCnt1

       DBNyukaUPD = True
       Call ZCTClose
       Exit Function
    End If
    
    gInt = 0
    
    Do Until Not ZCT_RDSTS
        
        With ZCT
            If .ì¸â◊ì˙ <> iB1_SYmd Or .ì¸â◊ï÷ <> iB1_SBin Or _
               Trim(.ì¸â◊îıçl) <> Trim(iB1_Biko(.î[ïièëçs - 1)) Then
               .ì¸â◊ì˙ = iB1_SYmd
               If iB1_SBin.Visible Then
                  If CisFun.Val2(iB1_SBin) <> 0 Then
                     .ì¸â◊ï÷ = Format(iB1_SBin, "00")
                  End If
               End If
               gSL_Select = "Update çﬁóøì¸â◊ÉeÅ[ÉuÉã Set "
               gSL_Select = gSL_Select & "ì¸â◊ì˙ = '" & RTrim(.ì¸â◊ì˙) & "',"
               gSL_Select = gSL_Select & "ì¸â◊ï÷ = '" & RTrim(.ì¸â◊ï÷) & "',"
               gSL_Select = gSL_Select & "ì¸â◊îıçl = '" & Trim(iB1_Biko(.î[ïièëçs - 1)) & "',"
               gSL_Select = gSL_Select & "çXêVì˙ = GETDATE(),"
               gSL_Select = gSL_Select & "çXêVí[ññ = SubString(Host_Name(),1,20),"
               gSL_Select = gSL_Select & "çXêVé“ = '" & RTrim(.çXêVé“) & "'"
               gSL_Select = gSL_Select & " Where î≠íçä«óùNO = " & .î≠íçä«óùNO
               
               With CisDB
                   .SQL = gSL_Select
                   If Not .DBExec Then Exit Function
               End With
               '+ çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã çXêV
               If gInt = 0 Then
                  If Not HPrtUpdNyuka Then Exit Function
                  gInt = gInt + 1
               End If
            
               '===================================== 2009/01/19 Start
               gSL_Select = "Update çﬁóøé¿ê—ä«óùÉeÅ[ÉuÉã Set "
               gSL_Select = gSL_Select & "é¿ê—ì˙ = '" & RTrim(.ì¸â◊ì˙) & "',"
               gSL_Select = gSL_Select & "é¿ê—ï÷ = '" & RTrim(.ì¸â◊ï÷) & "',"
               gSL_Select = gSL_Select & "çXêVì˙ = GETDATE(),"
               gSL_Select = gSL_Select & "çXêVí[ññ = SubString(Host_Name(),1,20),"
               gSL_Select = gSL_Select & "çXêVé“ = '" & RTrim(.çXêVé“) & "'"
               gSL_Select = gSL_Select & " Where é¿ê—ä«óùNO = " & .éxããé¿ê—ä«óùNO
               gSL_Select = gSL_Select & " And   ì¸óÕéÌóﬁ = 'A' "
               gSL_Select = gSL_Select & " And   é¿ê—ãÊï™ = 3 "       'íºëó
               
               With CisDB
                   .SQL = gSL_Select
                   If Not .DBExec Then Exit Function
               End With
               '===================================== 2009/01/19 End
               
            End If
        End With
        Call ZCTReadNext
    Loop
        
    Call ZCTClose
    
    DBNyukaUPD = True
End Function
'+////////////////////////////////////////////////////
'+//
'+//      î≠íçí†ï[î≠çsÉeÅ[ÉuÉãÅ@çXêV
'+//
'+////////////////////////////////////////////////////
'+-----------------------------------+
'+      î≠íçèÓïÒÅ@î[ì¸ì˙ïœçX
'+-----------------------------------+
Private Function HPrtUpdHachu() As Boolean
    Dim wPrtType        As Integer
    HPrtUpdHachu = False
    
    With ZCT
        gSL_Select = "Update çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã Set "
        gSL_Select = gSL_Select & " î[ì¸ì˙ = '" & RTrim(.î[ì¸ì˙) & "',"
        gSL_Select = gSL_Select & " î[ì¸ï÷ = '" & RTrim(.î[ì¸ï÷) & "',"
        gSL_Select = gSL_Select & "î†êî" & .î[ïièëçs & " = " & .ñáêî & ","
        gSL_Select = gSL_Select & "í[êî" & .î[ïièëçs & " = " & .í[êî & ","
        gSL_Select = gSL_Select & "î[ì¸êî" & .î[ïièëçs & " = " & .î≠íçêî
        
        gSL_Select = gSL_Select & " Where î[ïièëä«óùNO = " & .î[ïièëä«óùNO
        gSL_Select = gSL_Select & "   And ï≈NO = " & .î[ïièëï≈
        gSL_Select = gSL_Select & "   And ÉJÅ[Éh IN ('6','7','8')"
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
        gSL_Select = gSL_Select & "   And ÉJÅ[Éh IN ('6','7','8')"
        With CisDB
            .SQL = gSL_Select
            If Not .DBExec Then Exit Function
        End With
    End With

    HPrtUpdHachu = True

End Function
'+-----------------------------------+
'+      î≠íçèÓïÒÅ@çÌèú
'+-----------------------------------+
Private Function HPrtDelHachu(DelType As Integer) As Boolean
    Dim wPrtType        As Integer
    HPrtDelHachu = False
    
    With ZCT
        If DelType = 0 Then
           gSL_Select = "Delete çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã "
           gSL_Select = gSL_Select & " Where î[ïièëä«óùNO = " & .î[ïièëä«óùNO
           gSL_Select = gSL_Select & "   And ï≈NO = " & .î[ïièëï≈
           gSL_Select = gSL_Select & "   And ÉJÅ[Éh IN ('6','7','8')"
        Else
           gSL_Select = "Update çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã Set "
           gSL_Select = gSL_Select & "çﬁóøä«óùî‘çÜ" & .î[ïièëçs & " = 0,"
           gSL_Select = gSL_Select & "é˚óeêî" & .î[ïièëçs & " = 0,"
           gSL_Select = gSL_Select & "î†êî" & .î[ïièëçs & " = 0,"
           gSL_Select = gSL_Select & "í[êî" & .î[ïièëçs & " = 0,"
           gSL_Select = gSL_Select & "î[ì¸êî" & .î[ïièëçs & " = 0"
           gSL_Select = gSL_Select & " Where î[ïièëä«óùNO = " & .î[ïièëä«óùNO
           gSL_Select = gSL_Select & "   And ï≈NO = " & .î[ïièëï≈
           gSL_Select = gSL_Select & "   And ÉJÅ[Éh IN ('6','7','8')"
        End If
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
        gSL_Select = gSL_Select & "   And ÉJÅ[Éh IN ('6','7','8')"
        With CisDB
            .SQL = gSL_Select
            If Not .DBExec Then Exit Function
        End With
    End With

    HPrtDelHachu = True

End Function
'+-----------------------------------+
'+      ì¸â◊èÓïÒÅ@ì¸â◊ì˙ïœçX
'+-----------------------------------+
Private Function HPrtUpdNyuka() As Boolean
    HPrtUpdNyuka = False
    
    With ZCT
        gSL_Select = "Update çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã Set "
        gSL_Select = gSL_Select & " î[ì¸ì˙ = '" & iB1_SYmd & "'"
        If iB1_SBin.Visible Then
           If CisFun.Val2(iB1_SBin) <> 0 Then
              gSL_Select = gSL_Select & ", î[ì¸ï÷ = '" & Format(iB1_SBin, "00") & "'"
           End If
        End If
        
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
'+-----------------------------------+
'+      ì¸â◊èÓïÒÅ@çÌèú
'+-----------------------------------+
Private Function HPrtDelNyuka() As Boolean
    HPrtDelNyuka = False
    
    With ZCT
        gSL_Select = "Delete çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã  From çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã NY "
        gSL_Select = gSL_Select & " Left Outer Join çﬁóøî≠íçÉeÅ[ÉuÉã HC "
        gSL_Select = gSL_Select & "   On HC.î[ïièëä«óùNO = NY.î[ïièëä«óùNO "
        gSL_Select = gSL_Select & "   And HC.î[ì¸î‘çÜ = NY.î[ì¸î‘çÜ"
        gSL_Select = gSL_Select & "   And HC.î[ïièëï≈ = NY.ï≈NO"
        gSL_Select = gSL_Select & " Where HC.èââÒî≠íçä«óùNO = " & .èââÒî≠íçä«óùNO & ""
        gSL_Select = gSL_Select & "   And HC.î[ì¸î‘çÜ > '" & RTrim(.î[ì¸î‘çÜ) & "'"
        gSL_Select = gSL_Select & "   And NY.ÉJÅ[Éh IN ('6','7','8')"
        gSL_Select = gSL_Select & "   And SUBSTRING(NY.î[ì¸î‘çÜ,1,4) = '" & Mid(.î[ì¸î‘çÜ, 1, 4) & "'"
        With CisDB
            .SQL = gSL_Select
            If Not .DBExec Then Exit Function
        End With
        
        gSL_Select = "Delete çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã  From çﬁóøî≠íçí†ï[î≠çsÉeÅ[ÉuÉã NY "
        gSL_Select = gSL_Select & " Left Outer Join çﬁóøî≠íçÉeÅ[ÉuÉã HC "
        gSL_Select = gSL_Select & "   On HC.î[ïièëä«óùNO = NY.î[ïièëä«óùNO "
        gSL_Select = gSL_Select & "   And HC.î[ì¸î‘çÜ = NY.î[ì¸î‘çÜ"
        gSL_Select = gSL_Select & "   And HC.î[ïièëï≈ = NY.ï≈NO"
        gSL_Select = gSL_Select & " Where HC.î[ïièëä«óùNO = " & .î[ïièëä«óùNO
        gSL_Select = gSL_Select & "   And HC.î[ì¸î‘çÜ = '" & RTrim(.î[ì¸î‘çÜ) & "'"
        gSL_Select = gSL_Select & "   And HC.î[ïièëï≈ = " & .î[ïièëï≈
        gSL_Select = gSL_Select & "   And ÉJÅ[Éh = '6'"
        gSL_Select = gSL_Select & "   And î≠çsãÊï™ = 2"
        With CisDB
             .SQL = gSL_Select
             If Not .DBExec Then Exit Function
        End With
        
    End With

    HPrtDelNyuka = True

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
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 25, 1, "çﬁÅ@éø / ê°  ñ@")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
        'Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 5, 2, "îwî‘")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 8, 2, "îwî‘çÜ")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 11, 3, "é˚óeêî")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 4, 4, "î†êî")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 13, 5, "î≠íçêî")
        Call .FixedSet(íÜíÜ, íÜíÜ, Ç†ÇË, 4, 6, "é¿êî")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 7, "äÆî[ó\íËì˙")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 8, "ó\íËï÷")
'Å@∏ﬁÿØƒﬁëIëèàóùópÅ@ëﬁî¥ÿ±
'        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç»Çµ, 0,9,"Select")
'  ï\é¶ ﬂ¿∞›æØƒ
'       Call .PatanSet( ﬂ¿∞›NO,ï\é¶ÉpÉ^Å[Éì)
        Call .PatanSet(0, "0,1,2,3,6")
'  ï“èWì‡óeÉZÉbÉg
'       Call .EditSet(√∑ΩƒµÃﬁºﬁ™∏ƒ,∂◊—NO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'******************************************'
'*****       ÉOÉäÉbÉhì‡óeÉZÉbÉg        *****
'******************************************'
Private Sub GridSet()
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
            'çﬁéø/ê°ñ@
            gStr = ZCT.çﬁéø
            If Trim(ZCT.ï\é¶ê°ñ@) <> "" Then
                gStr = gStr & vbCr & ZCT.ï\é¶ê°ñ@
            End If
            .TextMatrix(ZCT.î[ïièëçs, 1) = gStr
            ' îwî‘çÜ
            .TextMatrix(ZCT.î[ïièëçs, 2) = " " & ZCT.îwî‘çÜ
            ' é˚óeêî
            If ZCT.é˚óeêî <> 0 Then
               .TextMatrix(ZCT.î[ïièëçs, 3) = Format(ZCT.é˚óeêî, "###0.000")
            End If
            ' ñáêî
'            .TextMatrix(ZCT.î[ïièëçs, 4) = Format(ZCT.ñáêî, "#,###")
            ' î≠íçêî
            If ZCT.î≠íçêî <> 0 Then
               .TextMatrix(ZCT.î[ïièëçs, 5) = Format(ZCT.î≠íçêî, "###0.000")
            End If
            ' é¿êîãÊï™
            If ZCT.é¿êîãÊï™ = 1 Then
               .TextMatrix(ZCT.î[ïièëçs, CisVsGrid.FixedGet(6)) = "Åõ"
            End If
            
            If Trim(ZCT.ì¸â◊ì˙) <> "" Then
               ' ì¸â◊ì˙•ï÷
               iB1_SYmd = ZCT.ì¸â◊ì˙
               If Trim(ZCT.ì¸â◊ï÷) <> "" Then
                  iB1_SBin = ZCT.ì¸â◊ï÷
               End If
               ' ï™äÑî[ì¸êî
               iB1_Suryo(ZCT.î[ïièëçs - 1) = CisFun.RSetFld(ZCT.ï™äÑêî, 0, iB1_Suryo(0).cFormat)
               ' îıçl
               iB1_Biko(ZCT.î[ïièëçs - 1) = RTrim(ZCT.ì¸â◊îıçl)
            End If
            
            iB1_Suryo(ZCT.î[ïièëçs - 1) = CisFun.RSetFld(ZCT.î≠íçêî, 0, iB1_Suryo(0).cFormat)
            B1lb_Maisu(ZCT.î[ïièëçs - 1) = CisFun.RSetFld(ZCT.ñáêî, 0, "###")
            If ZCT.í[êî <> 0 Then
               B1lb_HHasu(ZCT.î[ïièëçs - 1) = CisFun.RSetFld(ZCT.í[êî, 0, iB1_Suryo(0).cFormat)
            End If
            
            B1lb_NSuryo(ZCT.î[ïièëçs - 1) = CisFun.RSetFld(ZCT.ì¸â◊êî, 0, iB1_Suryo(0).cFormat)
            B1lb_NMaisu(ZCT.î[ïièëçs - 1) = CisFun.RSetFld(ZCT.ì¸â◊ñáêî, 0, "###")
            If ZCT.ì¸â◊í[êî <> 0 Then
               B1lb_NHasu(ZCT.î[ïièëçs - 1) = CisFun.RSetFld(ZCT.ì¸â◊í[êî, 0, iB1_Suryo(0).cFormat)
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
           B1lb_Maisu(gObj.Index).Visible = False
           B1lb_HHasu(gObj.Index).Visible = False
        Else
           B1lb_Maisu(gObj.Index).Visible = True
           B1lb_HHasu(gObj.Index).Visible = True
           gObj.Visible = True
           If Trim(iB1_SYmd) = "" And Mid(B1lb_Nouno, 5, 1) = "0" Then
              gObj.cLostColor = iB1_NYmd.cLostColor
              Back_Suryo(gObj.Index).Enabled = True
           Else
              Back_Suryo(gObj.Index).Enabled = False
              gObj.cLostColor = B1lb_HHasu(0).BackColor
           End If
        End If
    Next gObj
    
    For Each gObj In B1lb_NMaisu
        gStr = Trim(VsGrid1.TextMatrix(gObj.Index + 1, 1))
        If gStr = "" Or Not iB1_SYmd.Visible Then
           gObj.Visible = False
           B1lb_NHasu(gObj.Index).Visible = False
           B1lb_NSuryo(gObj.Index).Visible = False
           iB1_Biko(gObj.Index).Visible = False
        Else
           gObj.Visible = True
           B1lb_NHasu(gObj.Index).Visible = True
           B1lb_NSuryo(gObj.Index).Visible = True
           iB1_Biko(gObj.Index).Visible = True
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
    
    If QR_CommEvent <> "OK" Then Exit Sub
    
    With QRDataArea
        If .ÉoÅ[ÉRÅ[ÉhéÌóﬁ <> "B" Then
            If .ÉXÉeÅ[É^ÉX <> 0 Then
               Call QRReadMsg("E")
               Exit Sub
            End If
            ' ïîïiÇÃÇ›óLå¯
            If .ÉoÅ[ÉRÅ[ÉhéÌóﬁ <> "Z" Then
               .ÉGÉâÅ[ì‡óe = "çﬁóøÇÃì`ï[Ç≈Ç»Ç¢"
               Call QRReadMsg("E")
               Exit Sub
            End If
'            ' î[ïièëÇÃÇ›óLå¯
'            If .ÉJÅ[Éh <> "7" Then
'               .ÉGÉâÅ[ì‡óe = "ÉJÅ[ÉhÇ™ëŒè€äOÇ≈Ç∑"
'               Call QRReadMsg("E")
'               Exit Sub
'            End If
        End If
    End With
'( ÉfÅ[É^ämîF )
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where HC.î[ïièëä«óùNO = " & QRDataArea.î[ïièëä«óùNO
    gSL_Select = gSL_Select & "   And HC.î[ïièëï≈ = " & QRDataArea.î[ïièëï≈
    
    If Not DBInput("INV") Then Exit Sub
    
    ProcHB = "B1"
    Call DispChange(ProcHB)
End Sub

