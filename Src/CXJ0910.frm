VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Begin VB.Form CXJ0910 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "éÛíçèoâ◊ÉfÅ[É^èCê≥"
   ClientHeight    =   10710
   ClientLeft      =   120
   ClientTop       =   1560
   ClientWidth     =   15375
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
   ScaleHeight     =   10710
   ScaleWidth      =   15375
   WindowState     =   2  'ç≈ëÂâª
   Begin Cis3D_v60.CIS3D ErrMsg 
      Height          =   465
      Left            =   1575
      Top             =   1065
      Visible         =   0   'False
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   820
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
      TabIndex        =   12
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
      Height          =   990
      Left            =   1560
      ScaleHeight     =   930
      ScaleWidth      =   2025
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   2085
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   810
         Index           =   1
         Left            =   75
         Top             =   60
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1429
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
         Begin CisText_V60.CisText iH1_Bar 
            Height          =   360
            Left            =   150
            TabIndex        =   0
            Top             =   405
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
      Left            =   8505
      ScaleHeight     =   660
      ScaleWidth      =   6060
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   645
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
      Height          =   7770
      Left            =   1545
      ScaleHeight     =   7710
      ScaleWidth      =   12990
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2655
      Width           =   13050
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   7590
         Left            =   60
         Top             =   45
         Width           =   12885
         _ExtentX        =   22728
         _ExtentY        =   13388
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
         Begin Cis3D_v60.CIS3D BLB_Jyuch 
            Height          =   6090
            Left            =   6450
            Top             =   1290
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   10742
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
               Height          =   375
               Left            =   30
               Top             =   30
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Height          =   375
               Index           =   10
               Left            =   30
               Top             =   405
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   0
               Left            =   990
               Top             =   30
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
            Begin Cis3D_v60.CIS3D CIS3D15 
               Height          =   375
               Left            =   2070
               Top             =   30
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "éÛíçêî"
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
               Height          =   375
               Index           =   0
               Left            =   2070
               Top             =   405
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
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
                  Height          =   345
                  Index           =   0
                  Left            =   45
                  TabIndex        =   13
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   375
               Index           =   11
               Left            =   990
               Top             =   405
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   0
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   12
               Left            =   30
               Top             =   780
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   375
               Index           =   1
               Left            =   2070
               Top             =   780
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
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
                  Height          =   345
                  Index           =   1
                  Left            =   45
                  TabIndex        =   14
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   375
               Index           =   13
               Left            =   990
               Top             =   780
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   1
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   14
               Left            =   30
               Top             =   1155
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   375
               Index           =   2
               Left            =   2070
               Top             =   1155
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
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
                  Height          =   345
                  Index           =   2
                  Left            =   45
                  TabIndex        =   15
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   375
               Index           =   15
               Left            =   990
               Top             =   1155
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   2
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   16
               Left            =   30
               Top             =   1530
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   375
               Index           =   3
               Left            =   2070
               Top             =   1530
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
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
                  Height          =   345
                  Index           =   3
                  Left            =   45
                  TabIndex        =   16
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   375
               Index           =   17
               Left            =   990
               Top             =   1530
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   3
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   18
               Left            =   30
               Top             =   1905
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   375
               Index           =   4
               Left            =   2070
               Top             =   1905
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
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
                  Height          =   345
                  Index           =   4
                  Left            =   45
                  TabIndex        =   17
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   375
               Index           =   19
               Left            =   990
               Top             =   1905
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   4
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   20
               Left            =   30
               Top             =   2280
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   375
               Index           =   5
               Left            =   2070
               Top             =   2280
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
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
                  Height          =   345
                  Index           =   5
                  Left            =   45
                  TabIndex        =   18
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   375
               Index           =   21
               Left            =   990
               Top             =   2280
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   5
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   22
               Left            =   30
               Top             =   2655
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   375
               Index           =   6
               Left            =   2070
               Top             =   2655
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
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
                  Height          =   345
                  Index           =   6
                  Left            =   45
                  TabIndex        =   19
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   375
               Index           =   23
               Left            =   990
               Top             =   2655
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   6
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   24
               Left            =   30
               Top             =   3030
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   375
               Index           =   7
               Left            =   2070
               Top             =   3030
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
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
                  Height          =   345
                  Index           =   7
                  Left            =   45
                  TabIndex        =   20
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   375
               Index           =   25
               Left            =   990
               Top             =   3030
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   7
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   26
               Left            =   30
               Top             =   3405
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   375
               Index           =   8
               Left            =   2070
               Top             =   3405
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
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
                  Height          =   345
                  Index           =   8
                  Left            =   45
                  TabIndex        =   21
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   375
               Index           =   27
               Left            =   990
               Top             =   3405
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   8
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   28
               Left            =   30
               Top             =   3780
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   375
               Index           =   9
               Left            =   2070
               Top             =   3780
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
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
                  Height          =   345
                  Index           =   9
                  Left            =   45
                  TabIndex        =   22
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   375
               Index           =   29
               Left            =   990
               Top             =   3780
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   9
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   30
               Left            =   30
               Top             =   4155
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
                  Index           =   10
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   375
               Index           =   10
               Left            =   2070
               Top             =   4155
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
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
                  Height          =   345
                  Index           =   10
                  Left            =   45
                  TabIndex        =   23
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   375
               Index           =   31
               Left            =   990
               Top             =   4155
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   10
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   32
               Left            =   30
               Top             =   4530
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
                  Index           =   11
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   375
               Index           =   11
               Left            =   2070
               Top             =   4530
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
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
                  Height          =   345
                  Index           =   11
                  Left            =   45
                  TabIndex        =   24
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   375
               Index           =   33
               Left            =   990
               Top             =   4530
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   11
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   34
               Left            =   30
               Top             =   4905
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
                  Index           =   12
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   375
               Index           =   12
               Left            =   2070
               Top             =   4905
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
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
                  Height          =   345
                  Index           =   12
                  Left            =   45
                  TabIndex        =   25
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   375
               Index           =   35
               Left            =   990
               Top             =   4905
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   12
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   36
               Left            =   30
               Top             =   5280
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
                  Index           =   13
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   375
               Index           =   13
               Left            =   2070
               Top             =   5280
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   661
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
                  Height          =   345
                  Index           =   13
                  Left            =   45
                  TabIndex        =   26
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   375
               Index           =   37
               Left            =   990
               Top             =   5280
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   13
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
               Height          =   405
               Index           =   38
               Left            =   30
               Top             =   5655
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   714
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
                  Index           =   14
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "##,###"
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
            Begin Cis3D_v60.CIS3D Back_Suryo 
               Height          =   405
               Index           =   14
               Left            =   2070
               Top             =   5655
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   714
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
                  Height          =   345
                  Index           =   14
                  Left            =   45
                  TabIndex        =   27
                  Top             =   15
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   609
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Height          =   405
               Index           =   39
               Left            =   990
               Top             =   5655
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   714
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
               Begin Cis3D_v60.CIS3D B1lb_Hasu 
                  Height          =   330
                  Index           =   14
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   12648384
                  Caption         =   "###,###"
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
         Begin vsFlexLib.vsFlexArray VsGrid1 
            Height          =   6090
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1275
            Width           =   6315
            _Version        =   196608
            _ExtentX        =   11139
            _ExtentY        =   10742
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
            FormatString    =   "áÇ|ïiÅ@î‘                   |îwî‘çÜ    |é˚óeêî |î† êî  |î≠íçêî  |ï™äÑî[ì¸êî|äÆî[ó\íËì˙|ó\íËï÷|"
            Rows            =   16
            BackColor       =   12648447
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   0
            ScrollBars      =   0
            RowHeightMin    =   375
            GridLines       =   3
            SelectionMode   =   1
            BackColorAlternate=   12648447
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   735
            Index           =   1
            Left            =   10020
            Top             =   105
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   1296
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "   èo â◊ ì˙  - ï÷"
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
               TabIndex        =   4
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
               TabIndex        =   3
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
            Left            =   6795
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
               TabIndex        =   2
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
               TabIndex        =   1
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
            Height          =   6090
            Left            =   9600
            Top             =   1290
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   10742
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
            Begin Cis3D_v60.CIS3D CIS3D17 
               Height          =   375
               Left            =   30
               Top             =   30
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Height          =   375
               Index           =   0
               Left            =   30
               Top             =   405
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   0
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   2
               Left            =   990
               Top             =   30
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Height          =   375
               Index           =   1
               Left            =   990
               Top             =   405
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   0
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   2
               Left            =   30
               Top             =   780
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   1
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   3
               Left            =   990
               Top             =   780
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   1
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   4
               Left            =   30
               Top             =   1155
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   2
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   5
               Left            =   990
               Top             =   1155
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   2
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   6
               Left            =   30
               Top             =   1530
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   3
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   7
               Left            =   990
               Top             =   1530
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   3
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   8
               Left            =   30
               Top             =   1905
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   4
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   9
               Left            =   990
               Top             =   1905
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   4
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   40
               Left            =   30
               Top             =   2280
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   5
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   41
               Left            =   990
               Top             =   2280
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   5
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   42
               Left            =   30
               Top             =   2655
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   6
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   43
               Left            =   990
               Top             =   2655
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   6
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   44
               Left            =   30
               Top             =   3030
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   7
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   45
               Left            =   990
               Top             =   3030
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   7
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   46
               Left            =   30
               Top             =   3405
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   8
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   47
               Left            =   990
               Top             =   3405
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   8
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   48
               Left            =   30
               Top             =   3780
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   9
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   49
               Left            =   990
               Top             =   3780
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   9
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   50
               Left            =   30
               Top             =   4155
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   10
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   51
               Left            =   990
               Top             =   4155
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   10
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   52
               Left            =   30
               Top             =   4530
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   11
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   53
               Left            =   990
               Top             =   4530
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   11
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   54
               Left            =   30
               Top             =   4905
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   12
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   55
               Left            =   990
               Top             =   4905
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   12
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   56
               Left            =   30
               Top             =   5280
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   13
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   375
               Index           =   57
               Left            =   990
               Top             =   5280
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   13
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   405
               Index           =   58
               Left            =   30
               Top             =   5655
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   714
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
               Begin Cis3D_v60.CIS3D B1lb_SMaisu 
                  Height          =   330
                  Index           =   14
                  Left            =   60
                  Top             =   30
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "##,###"
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
               Height          =   405
               Index           =   59
               Left            =   990
               Top             =   5655
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   714
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
               Begin Cis3D_v60.CIS3D B1lb_SHasu 
                  Height          =   330
                  Index           =   14
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   3
               Left            =   2070
               Top             =   30
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "èoâ◊êî"
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
               Height          =   375
               Index           =   60
               Left            =   2070
               Top             =   405
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   0
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   61
               Left            =   2070
               Top             =   780
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   1
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   62
               Left            =   2070
               Top             =   1155
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   2
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   63
               Left            =   2070
               Top             =   1530
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   3
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   64
               Left            =   2070
               Top             =   1905
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   4
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   65
               Left            =   2070
               Top             =   2280
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   5
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   66
               Left            =   2070
               Top             =   2655
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   6
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   67
               Left            =   2070
               Top             =   3030
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   7
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   68
               Left            =   2070
               Top             =   3405
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   8
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   69
               Left            =   2070
               Top             =   3780
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   9
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   70
               Left            =   2070
               Top             =   4155
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   10
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   71
               Left            =   2070
               Top             =   4530
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   11
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   72
               Left            =   2070
               Top             =   4905
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   12
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   375
               Index           =   73
               Left            =   2070
               Top             =   5280
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   661
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   13
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
               Height          =   405
               Index           =   74
               Left            =   2070
               Top             =   5655
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   714
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
               Begin Cis3D_v60.CIS3D B1lb_SSuryo 
                  Height          =   330
                  Index           =   14
                  Left            =   60
                  Top             =   30
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   582
                  BackColor       =   16761024
                  Caption         =   "###,###"
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
            Left            =   6450
            Top             =   930
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   635
            BackColor       =   12648384
            Caption         =   "<<  éÛíçèÓïÒ  >>"
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
            Left            =   9585
            Top             =   930
            Width           =   3180
            _ExtentX        =   5609
            _ExtentY        =   635
            BackColor       =   16761024
            Caption         =   "<<  èoâ◊èÓïÒ  >>"
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
         Begin Cis3D_v60.CIS3D Back_Mntkb 
            Height          =   1095
            Left            =   120
            Top             =   75
            Width           =   5895
            _ExtentX        =   10398
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
            Begin VB.OptionButton B1Op_ProcKB 
               BackColor       =   &H00C0C0C0&
               Caption         =   "éÛíçèÓïÒèCê≥"
               Height          =   315
               Index           =   0
               Left            =   420
               TabIndex        =   29
               Top             =   60
               Width           =   2085
            End
            Begin VB.OptionButton B1Op_ProcKB 
               BackColor       =   &H00C0C0C0&
               Caption         =   "éÛíçèÓïÒçÌèú"
               Height          =   315
               Index           =   1
               Left            =   420
               TabIndex        =   28
               Top             =   405
               Width           =   2085
            End
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
            Begin VB.Label Back_Bunnou 
               Caption         =   "à»ç~ÇÃï™î[ÉfÅ[É^çÌèúÇ‡çsÇ¢Ç‹Ç∑ÅB"
               Height          =   255
               Left            =   660
               TabIndex        =   30
               Top             =   720
               Visible         =   0   'False
               Width           =   4455
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
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "Åy éÛíçèoâ◊ÉfÅ[É^èCê≥  Åz"
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
      Height          =   990
      Left            =   7005
      ScaleHeight     =   930
      ScaleWidth      =   7545
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1605
      Width           =   7605
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   840
         Left            =   3615
         Top             =   60
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   1482
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
         Begin Cis3D_v60.CIS3D B1lb_YMD 
            Height          =   360
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
         Begin Cis3D_v60.CIS3D B1lb_Bin 
            Height          =   360
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
         Height          =   840
         Left            =   60
         Top             =   60
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   1482
         ForeColor       =   0
         Caption         =   "î[ ì¸ êÊ"
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
            Top             =   405
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
         Height          =   840
         Left            =   5415
         Top             =   60
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   1482
         ForeColor       =   0
         Caption         =   "éÛíçãÊ"
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
         Height          =   840
         Left            =   6165
         Top             =   60
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   1482
         ForeColor       =   0
         Caption         =   "SKãÊ"
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
         Begin Cis3D_v60.CIS3D B1lb_SKKbn 
            Height          =   360
            Left            =   90
            Top             =   390
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXX"
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
         End
      End
   End
   Begin Cis3D_v60.CIS3D H2_Area1 
      Height          =   990
      Left            =   270
      Top             =   480
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   1746
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
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   0
         Width           =   240
      End
      Begin Cis3D_v60.CIS3D QRMsg 
         Height          =   825
         Left            =   90
         Top             =   90
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1455
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
Attribute VB_Name = "CXJ0910"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   éÛíçÅEèoâ◊ÉfÅ[É^èCê≥
'**       ÉtÉHÅ[ÉÄID    :   CXJ0910
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2004/04/25  By CIS
'**       ïœ  çX  ì˙    :   2004/06/29  By CIS ÉoÉOèCê≥Åièoâ◊ÉfÅ[É^çÌèúéûÅAéÛíçÉfÅ[É^ÇÃèoâ◊é¿ê—ãÊï™Ç™å≥Ç…ñﬂÇÁÇ»Ç¢Åj
'**       ïœ  çX  ì˙    :   2004/10/06  By CIS ñæç◊ÇÃå©èoÇµÇèCê≥ÅiéÛíçèÓïÒÅEèoâ◊èÓïÒÅj
'**                         2008/04/10  By CIS îwî‘çÜ8åÖ
'**       ïœ  çX  ì˙    :   2009/12/15  By CIS ìyäÚí√Ver•••é¿ê—ä«óù√∞ÃﬁŸ(èoå…√ﬁ∞¿)òAìÆ
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
    
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås    ÇkÇèÇÅÇÑÉCÉxÉìÉg    Åt            +
'+----------------------------------------------------------+
Private Sub form_load()

    If Rv_NCall = "" Then
    '   #-------------------#
    '   # ìÒ èd ãN ìÆ ñh é~  #
    '   #-------------------#
        If Not Dupli_Start Then End
    '   #-------------------#
    '   # èâ ä˙ ì‡ óe éÊ ìæ  #
    '   #-------------------#
        If Not IniGet Then End
    End If
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    QRDummy.Left = -1000
    
    H2_Area1.Move H1_Area1.Left, H1_Area1.Top
'    H3_Area1.Move H1_Area1.Left, H1_Area1.Top
    
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
    B1Op_ProcKB(0).Caption = "éÛíçèÓïÒèCê≥"
    B1Op_ProcKB(1).Caption = "éÛíçèÓïÒçÌèú"
    
    Set QT10QSR_Comm = PDQComm1
    Call QT10QSRInfoGet
    
    If QT10QSR_UseFlg Then
       IRN_NKRead = 1
    Else
       IRN_NKRead = 0
    End If
    
'   #------------------#
'   #    èâä˙ì¸óÕê›íË   #
'   #------------------#
    K_Sykbnm.Tag = "2"
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
    
    PB_Look.Visible = False
    
    mSL_Select = "SELECT"
    mSL_Select = mSL_Select & " JY.éÛíçä«óùNO,JY.èoâ◊éwé¶èëNO,JY.èoâ◊éwé¶èëçsNO,"
    mSL_Select = mSL_Select & " JY.î[ì¸ì˙,JY.î[ì¸ï÷,JY.èoâ◊ó\íËì˙,JY.èoâ◊ó\íËï÷, "
    mSL_Select = mSL_Select & " ISNULL(SK.èoâ◊ì˙,'') èoâ◊ì˙,ISNULL(SK.èoâ◊ï÷,'') èoâ◊ï÷,"
    mSL_Select = mSL_Select & " JY.ï™î[âÒêî,  JY.î[ì¸êÊ,JY.éÛì¸,JY.î[ïièëä«óùNO,"
    mSL_Select = mSL_Select & " JY.éÛíçãÊï™,JY.å≥èoâ◊éwé¶èëNO,JY.å≥èoâ◊éwé¶èëçsNO,"
    mSL_Select = mSL_Select & " ISNULL(TR.ó™èÃ,'') î[ì¸êÊñº,"
    mSL_Select = mSL_Select & " ISNULL(SY1.ílñºèÃ,'') éÛíçãÊï™ñº,"
    mSL_Select = mSL_Select & " CASE WHEN HT.SKãÊï™ = 0 THEN 'Ç©ÇÒÇŒÇÒ' ELSE 'éw é¶' END SKãÊï™ñº,"
    mSL_Select = mSL_Select & " HN.ï\é¶ïiî‘ ï\é¶ïiî‘,JY.îwî‘çÜ,"
    mSL_Select = mSL_Select & " JY.é˚óeêî , JY.í[êî, JY.ñáêî, JY.î[ì¸êî, SK.èoâ◊ñáêî, SK.èoâ◊í[êî, SK.èoâ◊êî"
    mSL_Select = mSL_Select & " From éÛíçÉeÅ[ÉuÉã JY"
    mSL_Select = mSL_Select & " Left Outer Join èoâ◊ÉeÅ[ÉuÉã SK"
    mSL_Select = mSL_Select & "   On SK.èoâ◊éwé¶èëNO = JY.èoâ◊éwé¶èëNO"
    mSL_Select = mSL_Select & "  AND SK.èoâ◊éwé¶èëçsNO = JY.èoâ◊éwé¶èëçsNO"
    mSL_Select = mSL_Select & " Left Outer Join ïiî‘É}ÉXÉ^ HN"
    mSL_Select = mSL_Select & "   ON HN.ïiî‘ = JY.ïiî‘"
    mSL_Select = mSL_Select & " Left Outer Join ïiî‘éÊà¯êÊÉ}ÉXÉ^ HT"
    mSL_Select = mSL_Select & "   ON HT.ïiî‘ = JY.ïiî‘"
    mSL_Select = mSL_Select & "  AND HT.éÊà¯êÊ = JY.î[ì¸êÊ"
    mSL_Select = mSL_Select & "  AND HT.éÛì¸ = JY.éÛì¸"
    mSL_Select = mSL_Select & " Left Outer Join éÊà¯êÊÉ}ÉXÉ^ TR"
    mSL_Select = mSL_Select & "   ON TR.éÊà¯êÊãÊï™ = 0"
    mSL_Select = mSL_Select & "  And TR.éÊà¯êÊCD = JY.î[ì¸êÊ"
    mSL_Select = mSL_Select & " Left Outer Join ñºèÃÉ}ÉXÉ^ SY1"
    mSL_Select = mSL_Select & "   ON SY1.ãÊï™ñºèÃ = 'éÛíçãÊï™'"
    mSL_Select = mSL_Select & "  And SY1.ãÊï™É^ÉCÉv = ''"
    mSL_Select = mSL_Select & " And SY1.íl = JY.éÛíçãÊï™"

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
    If gLong > 2 Then gLong = 1
    
    mSaveHeadProc = ""
    mSaveHeadProc = "H1"
    If IRN_NKRead = 1 Then
       mSaveHeadProc = mSaveHeadProc & "H2"
    Else
       mSaveHeadProc = mSaveHeadProc & "  "
    End If
    If IRN_NKReadBar = 1 Then
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
        Case "H1"
            K_Sykbnm.Caption = "Åy  ﬁ∞∫∞ƒﬁ Åz"
            H1_Area1.Visible = True
            H2_Area1.Visible = False
        Case "H2"
            K_Sykbnm.Caption = "Åy QRÿ∞¿ﬁ∞ Åz"
            H1_Area1.Visible = False
            H2_Area1.Visible = True
    End Select
    ProcHB = mHeadProc
    Call DispChange(ProcHB)
End Sub
'+----------------------------+
'+    å„  ëﬁ  ÉL Å[(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = "SELECT èoâ◊éwé¶èëNO FROM éÛíçÉeÅ[ÉuÉã "
    gSL_Select = gSL_Select & " WHERE  èoâ◊éwé¶èëNO < " & CisFun.Val2(iH1_Bar)
    gSL_Select = gSL_Select & " GROUP BY èoâ◊éwé¶èëNO"
    gSL_Select = gSL_Select & " ORDER BY èoâ◊éwé¶èëNO DESC"
    
    If Not JYTRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    êÊì™ÉåÉRÅ[ÉhÇ≈Ç∑            "
            .MB_Title = "ÅyëOï≈Åz"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Back_ED
    End If
    
    iH1_Bar = JYT.èoâ◊éwé¶èëNO
    
    gSL_Select = mSL_Select
    gSL_Select = gSL_Select & " Where JY.èoâ◊éwé¶èëNO = " & CisFun.Val2(iH1_Bar)
    gSL_Select = gSL_Select & " ORDER BY JY.èoâ◊éwé¶èëNO, JY.èoâ◊éwé¶èëçsNO     "
    
    Call DBInput("INV")
PB_Back_ED:
    
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    ëO  êi  ÉL Å[(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = "SELECT èoâ◊éwé¶èëNO FROM éÛíçÉeÅ[ÉuÉã "
    gSL_Select = gSL_Select & " WHERE  èoâ◊éwé¶èëNO > " & CisFun.Val2(iH1_Bar)
    gSL_Select = gSL_Select & " GROUP BY èoâ◊éwé¶èëNO"
    gSL_Select = gSL_Select & " ORDER BY èoâ◊éwé¶èëNO "
    
    If Not JYTRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ç≈èIÉåÉRÅ[ÉhÇ≈Ç∑            "
            .MB_Title = "Åyéüï≈Åz"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Next_ED
    End If
    
    iH1_Bar = JYT.èoâ◊éwé¶èëNO
    
    gSL_Select = mSL_Select
    gSL_Select = gSL_Select & " Where JY.èoâ◊éwé¶èëNO = " & CisFun.Val2(iH1_Bar)
    gSL_Select = gSL_Select & " ORDER BY JY.èoâ◊éwé¶èëNO, JY.èoâ◊éwé¶èëçsNO     "
    
    Call DBInput("INV")
PB_Next_ED:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iB1_NYmd" Then iB1_NYmd.ShowCalender
    If PB_Look.Tag = "iB1_SYmd" Then iB1_SYmd.ShowCalender
    If PB_Look.Tag = "iH1_Bar" Then Call Look_JYU               '2008/04/10
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
'*       éÛíçêî  ( iB1_Suryo )                  *
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
    gLong = CisFun.Val2(VsGrid1.TextMatrix(Index + 1, 3))
    gCnt3 = 0
    If gLong <> 0 Then
       gCnt2 = CisFun.Val2(iB1_Suryo(Index)) \ gLong
       gCnt3 = CisFun.Val2(iB1_Suryo(Index)) Mod gLong
       If gCnt3 <> 0 Then
          gCnt2 = gCnt2 + 1
       End If
    End If
    B1lb_Maisu(Index) = CisFun.RSetFld(gCnt2, 0, iB1_Suryo(Index).cFormat)
    B1lb_Hasu(Index) = CisFun.RSetFld(gCnt3, 0, iB1_Suryo(Index).cFormat)
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
    If ProcHB = "H2" Then
       Call QT10QSR_Open
       QRMsg.BackColor = &HC0FFC0
       QRMsg.Caption = "QR¿Ø¡ÿ∞¿ﬁì«éÊíÜ"
    Else
       Call QT10QSR_Close
       QRMsg.BackColor = &HFFC0C0
       QRMsg.Caption = "QR¿Ø¡ÿ∞¿ﬁì«éÊ"
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
'+-----------------------------+
'+       éÛíçÉfÅ[É^åüçı        +        2008/04/10
'+-----------------------------+
Private Sub Look_JYU()
    RV_Left = 0
    RV_Top = 0
    
    CXJ0911.Show vbModal
    Unload CXJ0911
    Set CXJ0911 = Nothing

    If RV_Rtn Then
       iH1_Bar = RV_Long
            
       Dummy.Enabled = True: Dummy.SetFocus: DoEvents
       Call ReturnPress
       Dummy.Enabled = False
    
    End If
End Sub
'*----------------------------------------------*
'*      éÛíçÉfÅ[É^åüçıóp  ( iH1_Bar )           *   2008/04/10
'*----------------------------------------------*
Private Sub iH1_Bar_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Bar_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
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
'    If ProcHB = "H2" Then
'        If Not Head2Chk Then GoTo ReturnPress_Ed
'        ProcHB = "B1"
'        Call DispChange(ProcHB)
'        GoTo ReturnPress_Ed
'    End If

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
    
'( ÉoÅ[ÉRÅ[Éh )
    If CisFun.Val2(iH1_Bar) = 0 Then
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
    
    gSL_Select = gSL_Select & " Where JY.èoâ◊éwé¶èëNO = " & CisFun.Val2(iH1_Bar)
    gSL_Select = gSL_Select & " Order By JY.èoâ◊éwé¶èëçsNO"
    
    If Not DBInput("INV") Then Exit Function
    
    Head1Chk = True
End Function
'****************************
'*      Body1¡™Ø∏èàóù
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    If iB1_SYmd.Visible Then
       If Not ChkSyuka Then Exit Function
    Else
       If Not ChkJyuch Then Exit Function
    End If
    Body1Chk = True
End Function
'+----------------------------+
'+      éÛíçèÓïÒÉ`ÉFÉbÉN
'+----------------------------+
Private Function ChkJyuch() As Boolean
    ChkJyuch = False
    
    If B1Op_ProcKB(1).Value Then
       ChkJyuch = True
       Exit Function
    End If
    
    mHachuDel = False
    For gInt = 0 To 14
        If CisFun.Val2(iB1_Suryo(gInt)) <> 0 Then
           mHachuDel = True
           Exit For
        End If
    Next gInt
    
    If Not mHachuDel Then
       B1Op_ProcKB(1).Value = True
       ChkJyuch = True
       Exit Function
    End If
'    If Not mHachuDel Then
'       iB1_NYmd = ""
'       ChkJyuch = True
'       Exit Function
'    End If
    
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
    gInt = Kdou_Check(iB1_NYmd, 0)
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
                .MB_MSG(1) = "    ÅyÅ@í èÌéÛíçéûÅAï÷ ïKê{ÅI Åz     "
                .MB_MSG(3) = "       î[ì¸ï÷Çì¸óÕÇµÇƒâ∫Ç≥Ç¢        "
                .MB_Button = Error
                .MBOX
            End With
            iB1_NBin.SetFocus
            Exit Function
        End If
    End If
    
    ChkJyuch = True
End Function
'+----------------------------+
'+      èoâ◊èÓïÒÉ`ÉFÉbÉN
'+----------------------------+
Private Function ChkSyuka() As Boolean
    ChkSyuka = False
    
    If B1Op_ProcKB(1).Value Then
       ChkSyuka = True
       Exit Function
    End If
'    If iB1_SYmd = "" Then
'        ChkSyuka = True
'        Exit Function
'    End If
    
    
    If iB1_SYmd = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    èoâ◊ì˙Çì¸óÕÇµÇƒâ∫Ç≥Ç¢         "
            .MB_Button = Error
            .MBOX
        End With
        iB1_SYmd.SetFocus
        Exit Function
    End If
    If Not iB1_SYmd.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    èoâ◊ì˙Çê≥ÇµÇ¢ì˙ïtÇ…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢         "
            .MB_Button = Error
            .MBOX
        End With
        iB1_SYmd.SetFocus
        Exit Function
    End If
    gInt = Kdou_Check(iB1_SYmd, 0)
    If gInt = 1 Then
    'ãxì˙
        CisFun.MB_Lines = 4
        CisFun.MB_MSG(1) = "     èoâ◊ì˙Ç™îÒâ“ì≠ì˙Ç≈Ç∑ÅB     "
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
                .MB_MSG(1) = "    ÅyÅ@í èÌéÛíçéûÅAï÷ ïKê{ÅI Åz     "
                .MB_MSG(3) = "       èoâ◊ï÷Çì¸óÕÇµÇƒâ∫Ç≥Ç¢        "
                .MB_Button = Error
                .MBOX
            End With
            iB1_SBin.SetFocus
            Exit Function
        End If
    End If
    
    ChkSyuka = True
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
'    gSL_Select = gSL_Select & " From éÛíçÉeÅ[ÉuÉã "
'    gSL_Select = gSL_Select & " Where î[ïièëä«óùNO = " & Rv_HNo(mBunCount)
'    gSL_Select = gSL_Select & "   And î[ïièëï≈ = " & Rv_HPage(mBunCount)
'    If Not HCTRead(gSL_Select, 1) Then
'       Call PB_END_Click
'    End If
'
'    iH1_NYmd = HCT.î[ì¸ì˙
'    iH1_Nouno = HCT.î[ì¸î‘çÜ
'    iH1_Page = CisFun.RSetFld(HCT.î[ïièëï≈, 0, iH1_Page.cFormat)
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
    If Not JYTRead(gSL_Select, 0) Then
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
    With JYT
    ' èoâ◊ì˙•èâä˙ï\é¶
        iB1_NYmd = .î[ì¸ì˙
        If Trim(.èoâ◊ì˙) = "" Then
           B1lb_NYMD.Visible = False
           B1lb_NBin.Visible = False
           iB1_NYmd.Visible = True
           iB1_SYmd.Visible = False
           iB1_SBin.Visible = False
           If Trim(JYT.î[ì¸ï÷) <> "" Then
              iB1_NBin.Visible = True
              iB1_NBin = .î[ì¸ï÷
           Else
              iB1_NBin.Visible = False
           End If
           B1Op_ProcKB(0).Caption = "éÛíçèÓïÒèCê≥"
           B1Op_ProcKB(1).Caption = "éÛíçèÓïÒçÌèú"
           Back_Bunnou.Visible = False
           BLB_Jyuch.Enabled = True
        Else
           B1lb_NYMD.Visible = True
           B1lb_NBin.Visible = True
           iB1_NYmd.Visible = False
           iB1_NBin.Visible = False
           iB1_SYmd = .èoâ◊ì˙
           iB1_SYmd.Visible = True
           If Trim(JYT.èoâ◊ï÷) <> "" Then
              iB1_SBin.Visible = True
              iB1_SBin = .èoâ◊ï÷
           Else
              iB1_SBin.Visible = False
           End If
           BLB_Jyuch.Enabled = False
           B1Op_ProcKB(0).Caption = "èoâ◊èÓïÒèCê≥"
           B1Op_ProcKB(1).Caption = "èoâ◊èÓïÒçÌèú"
           Back_Bunnou.Visible = True
        End If
        B1Op_ProcKB(0).Value = True
        
        If Trim(.î[ì¸êÊñº) = "" Then
           gStr = .î[ì¸êÊ
        Else
           gStr = RTrim(.î[ì¸êÊ) & ":" & .î[ì¸êÊñº
        End If
        B1lb_Tornm = gStr
        B1lb_YMD.Tag = .î[ì¸ì˙
        With CisFun
            .DateE_BefVal = JYT.î[ì¸ì˙
            .DateEdit
            B1lb_YMD = .DateE_AfVal
            .DateEdit J_YMD
            B1lb_NYMD = .DateE_AfVal
            If Trim(JYT.î[ì¸ï÷) <> "" Then
                B1lb_Bin = JYT.î[ì¸ï÷
            End If
        End With
        B1lb_NBin = B1lb_Bin
        
        B1lb_HKbn = .éÛíçãÊï™ñº
        B1lb_HKbn.Tag = .éÛíçãÊï™
        B1lb_SKKbn = .SKãÊï™ñº
        
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
           If RTrim(iB1_SYmd) = "" Then
              .MB_MSG(2) = "     èoâ◊èÓïÒçÌèúèàóùÇçsÇ¢Ç‹Ç∑             "
           Else
              .MB_MSG(2) = "     èoâ◊èÓïÒèCê≥èàóùÇçsÇ¢Ç‹Ç∑             "
           End If
        Else
           If RTrim(iB1_NYmd) = "" Then
              .MB_MSG(2) = "     éÛíçèÓïÒçÌèúèàóùÇçsÇ¢Ç‹Ç∑             "
           Else
              .MB_MSG(2) = "     éÛíçèÓïÒèCê≥èàóùÇçsÇ¢Ç‹Ç∑             "
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
    
    gSL_Select = "SELECT JY.*,"
    gSL_Select = gSL_Select & "  ISNULL(SK.èoâ◊ì˙,'') èoâ◊ì˙,ISNULL(SK.èoâ◊ï÷,'') èoâ◊ï÷,"
    gSL_Select = gSL_Select & "  SK.èoâ◊ñáêî,SK.èoâ◊í[êî,SK.èoâ◊êî "
    gSL_Select = gSL_Select & " From éÛíçÉeÅ[ÉuÉã JY"
    gSL_Select = gSL_Select & "  Left Outer Join èoâ◊ÉeÅ[ÉuÉã SK"
    gSL_Select = gSL_Select & "     On SK.èoâ◊éwé¶èëNO = JY.èoâ◊éwé¶èëNO"
    gSL_Select = gSL_Select & "    AND SK.èoâ◊éwé¶èëçsNO = JY.èoâ◊éwé¶èëçsNO"
    gSL_Select = gSL_Select & " Where JY.èoâ◊éwé¶èëNO = " & CisFun.Val2(iH1_Bar)
    
    If Not JYTRead(gSL_Select) Then
       Call JYTClose
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     éÛíçÉfÅ[É^àŸèÌ      "
            .MB_Title = "çXêVÉGÉâÅ["
            .MB_Button = OK
            .MBOX
            Exit Function
       End With
    End If
    
    CisDB.ConnectNo = 0
    Call CisDB.DBTran(TransBegin) ' ƒ◊›ªﬁ∏ºÆ›äJén
    
    If iB1_SYmd.Visible Then
        If Not DBSyukaUPD Then GoTo DBPut_Err
    Else
        If Not DBJyuchUPD Then GoTo DBPut_Err
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
    Call JYTClose
    
    Call CisDB.DBTran(TransRollback) ' ƒ◊›ªﬁ∏ºÆ›îjä¸
    CisFun.ErrorBox
End Function
'+----------------------------------------------+
'+      éÛíçèÓïÒèCê≥Å@çXêV
'+----------------------------------------------+
Private Function DBJyuchUPD() As Boolean
    DBJyuchUPD = False
    Dim wHSuryo         As Long
    
    
    With JYT
        If B1Op_ProcKB(1).Value Then        ' éÛíçèÓïÒçÌèú
            gSL_Select = "Delete éÛíçÉeÅ[ÉuÉã  "
            gSL_Select = gSL_Select & " Where èoâ◊éwé¶èëNO = " & .èoâ◊éwé¶èëNO
           
           With CisDB
                .SQL = gSL_Select
                If Not .DBExec Then Exit Function
           End With
           DBJyuchUPD = True
           Call JYTClose
           Exit Function
        End If
    End With
    
    Do Until Not JYT_RDSTS
        
        With JYT
            wHSuryo = CisFun.Val2(iB1_Suryo(.èoâ◊éwé¶èëçsNO - 1))
            If .î[ì¸ì˙ <> iB1_NYmd Or .î[ì¸ï÷ <> iB1_NBin Or _
               .î[ì¸êî <> wHSuryo Then
                
               If wHSuryo = 0 Then
                  
                  gSL_Select = "Delete éÛíçÉeÅ[ÉuÉã  "
                  gSL_Select = gSL_Select & " Where éÛíçä«óùNO = " & .éÛíçä«óùNO & ""
                  gSL_Select = gSL_Select & "   AND èoâ◊éwé¶èëNO = " & .èoâ◊éwé¶èëNO & ""
                  gSL_Select = gSL_Select & "   AND èoâ◊éwé¶èëçsNO = " & .èoâ◊éwé¶èëçsNO & ""
                  
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
                  gLong = CisFun.Val2(VsGrid1.TextMatrix(.èoâ◊éwé¶èëçsNO, 3))
                  If gLong <> 0 Then
                     gCnt1 = wHSuryo \ gLong
                     If wHSuryo Mod gLong <> 0 Then
                        gCnt1 = gCnt1 + 1
                        .í[êî = wHSuryo Mod gLong
                     End If
                  End If
                  .ñáêî = gCnt1
                  .î[ì¸êî = wHSuryo
                    
                  gSL_Select = "Update éÛíçÉeÅ[ÉuÉã Set "
                  gSL_Select = gSL_Select & "î[ì¸ì˙ = '" & RTrim(.î[ì¸ì˙) & "',"
                  gSL_Select = gSL_Select & "î[ì¸ï÷ = '" & RTrim(.î[ì¸ï÷) & "',"
                  gSL_Select = gSL_Select & "ñáêî = " & .ñáêî & ","
                  gSL_Select = gSL_Select & "í[êî = " & .í[êî & ","
                  gSL_Select = gSL_Select & "î[ì¸êî = " & .î[ì¸êî & ","
                  gSL_Select = gSL_Select & "çXêVì˙ = GETDATE(),"
                  gSL_Select = gSL_Select & "çXêVí[ññ = SubString(Host_Name(),1,20),"
                  gSL_Select = gSL_Select & "çXêVé“ = '" & RTrim(.çXêVé“) & "'"
                  gSL_Select = gSL_Select & " Where éÛíçä«óùNO = " & .éÛíçä«óùNO & ""
                  gSL_Select = gSL_Select & "   AND èoâ◊éwé¶èëNO = " & .èoâ◊éwé¶èëNO & ""
                  gSL_Select = gSL_Select & "   AND èoâ◊éwé¶èëçsNO = " & .èoâ◊éwé¶èëçsNO & ""
               
                  With CisDB
                      .SQL = gSL_Select
                      If Not .DBExec Then Exit Function
                  End With
               End If
            End If
        End With
        
        Call JYTReadNext
    Loop

    Call JYTClose
    
    DBJyuchUPD = True
End Function
'+----------------------------------------------+
'+      èoâ◊èÓïÒèCê≥Å@çXêV
'+----------------------------------------------+
Private Function DBSyukaUPD() As Boolean
    DBSyukaUPD = False
    Dim wHSuryo         As Long
    Dim wNouno(10)      As String
    Dim wFirstNo(10)    As Long
    Dim wNouKNo(10)     As Long
    Dim wPageNo(10)     As Byte
    Dim SaveSJNo        As Long
    Dim SaveMSJNo       As Long
    
    If B1Op_ProcKB(1).Value Then        ' èoâ◊èÓïÒçÌèú
       SaveSJNo = 0: SaveMSJNo = 0
       Do Until Not JYT_RDSTS
            
           With JYT
                If SaveSJNo = 0 Then
                   SaveSJNo = .èoâ◊éwé¶èëNO
                End If
                
                gSL_Select = "Delete èoâ◊ÉeÅ[ÉuÉã  "
                gSL_Select = gSL_Select & " Where èoâ◊éwé¶èëNO = " & .èoâ◊éwé¶èëNO
                gSL_Select = gSL_Select & "   And èoâ◊éwé¶èëçsNO = " & .èoâ◊éwé¶èëçsNO
                With CisDB
                     .SQL = gSL_Select
                     If Not .DBExec Then Exit Function
                End With
                
                '===================================================== 2009/12/15 Start
                gSL_Select = "Delete é¿ê—ä«óùÉeÅ[ÉuÉã  "
                gSL_Select = gSL_Select & " Where èoâ◊éwé¶èëNO = " & .èoâ◊éwé¶èëNO
                gSL_Select = gSL_Select & "   And èoâ◊éwé¶èëçsNO = " & .èoâ◊éwé¶èëçsNO
                With CisDB
                     .SQL = gSL_Select
                     If Not .DBExec Then Exit Function
                End With
                '===================================================== 2009/12/15 End
                
                
                gSL_Select = "UPDATE éÛíçÉeÅ[ÉuÉã SET "
                gSL_Select = gSL_Select & " èoâ◊é¿ê—ãÊï™ = 0,"
                gSL_Select = gSL_Select & " çXêVì˙ = GETDATE(),"
                gSL_Select = gSL_Select & " çXêVí[ññ = HOST_NAME(),"
                gSL_Select = gSL_Select & " çXêVé“ = '" & RTrim(gTanto) & "'"
                gSL_Select = gSL_Select & " Where éÛíçä«óùNO = " & .éÛíçä«óùNO & ""
'*--------------------------------------------------<< 2004/06/29 Update Start >>
'               gSL_Select = gSL_Select & "   AND å≥èoâ◊éwé¶èëNO = " & .èoâ◊éwé¶èëNO
'               gSL_Select = gSL_Select & "   AND å≥èoâ◊éwé¶èëçsNO = " & .èoâ◊éwé¶èëçsNO
                gSL_Select = gSL_Select & "   AND èoâ◊éwé¶èëNO = " & .èoâ◊éwé¶èëNO
                gSL_Select = gSL_Select & "   AND èoâ◊éwé¶èëçsNO = " & .èoâ◊éwé¶èëçsNO
'*--------------------------------------------------<< 2004/06/29 Update End >>
                With CisDB
                     .SQL = gSL_Select
                     If Not .DBExec Then Exit Function
                End With
            
            End With
            Call JYTReadNext
       Loop
        
       If SaveSJNo <> 0 Then
          Do
              gSL_Select = "Select * From éÛíçÉeÅ[ÉuÉã  "
              gSL_Select = gSL_Select & " Where å≥èoâ◊éwé¶èëNO = " & SaveSJNo
              gSL_Select = gSL_Select & "   And èoâ◊éwé¶èëNO <> å≥èoâ◊éwé¶èëNO"
              If Not JYTRead(gSL_Select, 1) Then Exit Do
              
              SaveSJNo = JYT.èoâ◊éwé¶èëNO
               
              With CisDB
                   .SQL = "Delete éÛíçÉeÅ[ÉuÉã  "
                   .SQL = .SQL & " Where èoâ◊éwé¶èëNO = " & JYT.èoâ◊éwé¶èëNO
                   If Not .DBExec Then Exit Function
                   
                   .SQL = "Delete èoâ◊ÉeÅ[ÉuÉã  "
                   .SQL = .SQL & " Where èoâ◊éwé¶èëNO = " & JYT.èoâ◊éwé¶èëNO
                   If Not .DBExec Then Exit Function
                                  
                   '=============================================== 2009/12/15 Start
                   .SQL = "Delete é¿ê—ä«óùÉeÅ[ÉuÉã  "
                   .SQL = .SQL & " Where èoâ◊éwé¶èëNO = " & JYT.èoâ◊éwé¶èëNO
                   If Not .DBExec Then Exit Function
                   '=============================================== 2009/12/15 End
              
              End With
          Loop
       End If
        
        
       DBSyukaUPD = True
       Call JYTClose
       Exit Function
    End If
    
    gInt = 0
    
    Do Until Not JYT_RDSTS
        
        With JYT
            If .èoâ◊ì˙ <> iB1_SYmd Or .èoâ◊ï÷ <> iB1_SBin Then
               .èoâ◊ì˙ = iB1_SYmd
               If iB1_SBin.Visible Then
                  If CisFun.Val2(iB1_SBin) <> 0 Then
                     .èoâ◊ï÷ = Format(iB1_SBin, "00")
                  End If
               End If
               gSL_Select = "Update èoâ◊ÉeÅ[ÉuÉã Set "
               gSL_Select = gSL_Select & "èoâ◊ì˙ = '" & RTrim(.èoâ◊ì˙) & "',"
               gSL_Select = gSL_Select & "èoâ◊ï÷ = '" & RTrim(.èoâ◊ï÷) & "',"
               gSL_Select = gSL_Select & "çXêVì˙ = GETDATE(),"
               gSL_Select = gSL_Select & "çXêVí[ññ = SubString(Host_Name(),1,20),"
               gSL_Select = gSL_Select & "çXêVé“ = '" & RTrim(.çXêVé“) & "'"
               gSL_Select = gSL_Select & " Where èoâ◊éwé¶èëNO = " & .èoâ◊éwé¶èëNO
               gSL_Select = gSL_Select & "   AND èoâ◊éwé¶èëçsNO = " & .èoâ◊éwé¶èëçsNO
               
               With CisDB
                   .SQL = gSL_Select
                   If Not .DBExec Then Exit Function
               End With
            
               '=========================================================== 2009/12/15 Satrt
               gSL_Select = "Update é¿ê—ä«óùÉeÅ[ÉuÉã Set "
               gSL_Select = gSL_Select & "é¿ê—ì˙ = '" & RTrim(.èoâ◊ì˙) & "',"
               gSL_Select = gSL_Select & "é¿ê—ï÷ = '" & RTrim(.èoâ◊ï÷) & "',"
               gSL_Select = gSL_Select & "çXêVì˙ = GETDATE(),"
               gSL_Select = gSL_Select & "çXêVí[ññ = SubString(Host_Name(),1,20),"
               gSL_Select = gSL_Select & "çXêVé“ = '" & RTrim(.çXêVé“) & "'"
               gSL_Select = gSL_Select & " Where èoâ◊éwé¶èëNO = " & .èoâ◊éwé¶èëNO
               gSL_Select = gSL_Select & "   AND èoâ◊éwé¶èëçsNO = " & .èoâ◊éwé¶èëçsNO
               
               With CisDB
                   .SQL = gSL_Select
                   If Not .DBExec Then Exit Function
               End With
               '=========================================================== 2009/12/15 End
            
            End If
        End With
        Call JYTReadNext
    Loop
        
    Call JYTClose
    
    DBSyukaUPD = True
End Function
'+////////////////////////////////////////////////////
'+//
'+//      éÛíçí†ï[î≠çsÉeÅ[ÉuÉãÅ@çXêV
'+//
'+////////////////////////////////////////////////////
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
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 1, "ïiÅ@î‘")
        Call .FixedSet(íÜíÜ, ç∂íÜ, Ç†ÇË, 0, 2, "îwî‘çÜ")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 3, "é˚óeêî")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 4, "î† êî")
        Call .FixedSet(íÜíÜ, âEíÜ, Ç†ÇË, 0, 5, "éÛíçêî")
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
    Do Until Not JYT_RDSTS
        With VsGrid1
            ID = ID + 1
            If JYT.î[ïièëçs > 15 Then Exit Do
            'No
            .TextMatrix(JYT.èoâ◊éwé¶èëçsNO, 0) = JYT.èoâ◊éwé¶èëçsNO
            'ïiî‘
            .TextMatrix(JYT.èoâ◊éwé¶èëçsNO, 1) = JYT.ï\é¶ïiî‘
            ' îwî‘çÜ
            .TextMatrix(JYT.èoâ◊éwé¶èëçsNO, 2) = " " & JYT.îwî‘çÜ
            ' é˚óeêî
            .TextMatrix(JYT.èoâ◊éwé¶èëçsNO, 3) = Format(JYT.é˚óeêî, "#,###")
            ' ñáêî
'            .TextMatrix(HCT.î[ïièëçs, 4) = Format(HCT.ñáêî, "#,###")
            ' éÛíçêî
            .TextMatrix(JYT.èoâ◊éwé¶èëçsNO, 5) = Format(JYT.î[ì¸êî, "#,###")
            ' é¿êîãÊï™
'            If JYT.é¿êîãÊï™ = 1 Then
'               .TextMatrix(JYT.èoâ◊éwé¶èëçsNO, CisVsGrid.FixedGet(6)) = "Åõ"
'            End If
            
            If Trim(JYT.èoâ◊ì˙) <> "" Then
               ' èoâ◊ì˙•ï÷
               iB1_SYmd = JYT.èoâ◊ì˙
               If Trim(JYT.èoâ◊ï÷) <> "" Then
                  iB1_SBin = JYT.èoâ◊ï÷
               End If
            End If
            
            iB1_Suryo(JYT.èoâ◊éwé¶èëçsNO - 1) = CisFun.RSetFld(JYT.î[ì¸êî, 0, iB1_Suryo(0).cFormat)
            B1lb_Maisu(JYT.èoâ◊éwé¶èëçsNO - 1) = CisFun.RSetFld(JYT.ñáêî, 0, iB1_Suryo(0).cFormat)
            B1lb_Hasu(JYT.èoâ◊éwé¶èëçsNO - 1) = CisFun.RSetFld(JYT.í[êî, 0, iB1_Suryo(0).cFormat)

            B1lb_SSuryo(JYT.èoâ◊éwé¶èëçsNO - 1) = CisFun.RSetFld(JYT.èoâ◊êî, 0, iB1_Suryo(0).cFormat)
            B1lb_SMaisu(JYT.èoâ◊éwé¶èëçsNO - 1) = CisFun.RSetFld(JYT.èoâ◊ñáêî, 0, iB1_Suryo(0).cFormat)
            B1lb_SHasu(JYT.èoâ◊éwé¶èëçsNO - 1) = CisFun.RSetFld(JYT.èoâ◊í[êî, 0, iB1_Suryo(0).cFormat)

        End With
        Call JYTReadNext
    Loop
    Call JYTClose
    With VsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    
    For Each gObj In iB1_Suryo
        gStr = Trim(VsGrid1.TextMatrix(gObj.Index + 1, 1))
        If gStr = "" Then
           gObj.Visible = False
           B1lb_Maisu(gObj.Index).Visible = False
           B1lb_Hasu(gObj.Index).Visible = False
        Else
           B1lb_Maisu(gObj.Index).Visible = True
           B1lb_Hasu(gObj.Index).Visible = True
           gObj.Visible = True
'           If Trim(iB1_SYmd) = "" And Mid(B1lb_Nouno, 5, 1) = "0" Then
           If Trim(iB1_SYmd) = "" Then
              gObj.cLostColor = iB1_NYmd.cLostColor
              Back_Suryo(gObj.Index).Enabled = True
           Else
              Back_Suryo(gObj.Index).Enabled = False
              gObj.cLostColor = B1lb_Hasu(0).BackColor
           End If
        End If
    Next gObj
    
    For Each gObj In B1lb_SMaisu
        gStr = Trim(VsGrid1.TextMatrix(gObj.Index + 1, 1))
        If gStr = "" Or Not iB1_SYmd.Visible Then
           gObj.Visible = False
           B1lb_SHasu(gObj.Index).Visible = False
           B1lb_SSuryo(gObj.Index).Visible = False
        Else
           gObj.Visible = True
           B1lb_SHasu(gObj.Index).Visible = True
           B1lb_SSuryo(gObj.Index).Visible = True
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
    
    gStr = QT10QSR_Event
    If gStr = "ERR" Then Exit Sub
    If gStr <> "RCV" Then Exit Sub
    
    If Not QT10QSR_RCVProc Then Exit Sub    ' CR Ç™ë∂ç›ÇµÇ»Ç¢éûî≤ÇØÇÈ

'( ÉfÅ[É^ämîF )
    If Trim(QT10QSR_ReadData) = "" Then Exit Sub
    
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where JY.èoâ◊éwé¶èëNO = " & CisFun.Val2(QT10QSR_ReadData)
    gSL_Select = gSL_Select & " Order By JY.èoâ◊éwé¶èëçsNO"
    
    QT10QSR_ReadData = ""
    
    If Not DBInput("INV") Then Exit Sub
    
    iH1_Bar = JYT.èoâ◊éwé¶èëNO
    
    ProcHB = "B1"
    Call DispChange(ProcHB)
End Sub



