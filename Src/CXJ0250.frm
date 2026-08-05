VERSION 5.00
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{BA5D7819-2B67-11D4-806A-00000EA57F9E}#3.0#0"; "CISYMD~1.OCX"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXJ0250 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "èoâ◊É`ÉPÉbÉgì¸óÕ"
   ClientHeight    =   11235
   ClientLeft      =   1800
   ClientTop       =   1515
   ClientWidth     =   15360
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ÇlÇr ÉSÉVÉbÉN"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'Z µ∞¿ﬁ∞
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   WindowState     =   2  'ç≈ëÂâª
   Begin PdqcommLib.PDQComm PDQComm1 
      Height          =   480
      Left            =   5018
      TabIndex        =   100
      Top             =   1830
      Width           =   480
      _Version        =   196610
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9.01
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
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   7808
      ScaleHeight     =   735
      ScaleWidth      =   5265
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   1920
      Width           =   5325
      Begin CisBtn_60.CisBtn PB_RED 
         Height          =   585
         Left            =   60
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "ì«éÊ"
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
         PFCaption       =   "( F5 )"
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
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   930
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   7
         BtnCaption      =   "êVãK"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2655
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   9
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
         Left            =   3510
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   9
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
         Left            =   4365
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   9
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
         Left            =   1800
         Top             =   60
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   9
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
   End
   Begin VB.PictureBox B1_Area3 
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
      Height          =   1065
      Left            =   2228
      ScaleHeight     =   1005
      ScaleWidth      =   2865
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   3750
      Width           =   2925
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   19
         Left            =   2205
         Top             =   60
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "ï÷"
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
         Begin CisText_V60.CisText iB13_Bin 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   390
            Width           =   345
            _ExtentX        =   609
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
            Text            =   "99"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   915
         Left            =   60
         Top             =   60
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "î[  ì¸  ì˙"
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
         Begin CisYMD_With_Btn.CisYMDwB iB13_YMD 
            Height          =   360
            Left            =   90
            TabIndex        =   2
            Top             =   390
            Width           =   1950
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
            Object.Width           =   1950
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
      Height          =   1005
      Left            =   2228
      ScaleHeight     =   945
      ScaleWidth      =   2895
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   1950
      Width           =   2955
      Begin Cis3D_v60.CIS3D lbNDay 
         Height          =   855
         Left            =   60
         Top             =   60
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1508
         ForeColor       =   16711680
         Caption         =   "èoÅ@â◊  ì˙"
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_YMD 
            Height          =   360
            Left            =   90
            TabIndex        =   0
            Top             =   390
            Width           =   1950
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
            Object.Width           =   1950
            cSize           =   1
            cChkResult      =   0   'False
            cAutoEject      =   0   'False
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   855
         Index           =   16
         Left            =   2235
         Top             =   60
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1508
         ForeColor       =   16711680
         Caption         =   "ï÷"
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
         Begin CisText_V60.CisText iH1_BIN 
            Height          =   375
            Left            =   120
            TabIndex        =   1
            Top             =   390
            Width           =   345
            _ExtentX        =   609
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
            Text            =   "99"
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
      Height          =   1065
      Left            =   6338
      ScaleHeight     =   1005
      ScaleWidth      =   6735
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   3750
      Width           =   6795
      Begin Cis3D_v60.CIS3D B1_C_Back 
         Height          =   915
         Left            =   5790
         Top             =   60
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "éxããå≥"
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
         Begin CisText_V60.CisText iB1_Moto 
            Height          =   375
            Left            =   150
            TabIndex        =   13
            Top             =   420
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            cFormat         =   "####"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "9999"
            MaxLength       =   4
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "####"
            cILength        =   4
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   5
         Left            =   5175
         Top             =   60
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "ã@î‘"
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
         Begin CisText_V60.CisText iB1_Kibn 
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   420
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
            cFbComma        =   0
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   4
         Left            =   4560
         Top             =   60
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "éÛì¸"
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
         Begin CisText_V60.CisText iB1_Uke 
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   420
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   3
         Left            =   3675
         Top             =   60
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "éxããêÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_SKTorcd 
            Height          =   375
            Left            =   150
            TabIndex        =   10
            Top             =   420
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            cFormat         =   "####"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "9999"
            MaxLength       =   4
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cGFormat        =   "####"
            cILength        =   4
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   2
         Left            =   2490
         Top             =   60
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "édì¸êÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_SRKojyo 
            Height          =   375
            Left            =   750
            TabIndex        =   9
            Top             =   420
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_SRTorcd 
            Height          =   375
            Left            =   150
            TabIndex        =   8
            Top             =   420
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   661
            cFormat         =   "####"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "9999"
            MaxLength       =   4
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cGFormat        =   "####"
            cILength        =   4
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   1
         Left            =   2025
         Top             =   60
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "#"
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
         Begin CisText_V60.CisText iB1_Sharp 
            Height          =   375
            Left            =   90
            TabIndex        =   7
            Top             =   420
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            cFormat         =   "0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "9"
            MaxLength       =   1
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cGFormat        =   "0"
            cILength        =   1
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   915
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1614
         ForeColor       =   16711680
         Caption         =   "î[ì¸î‘çÜ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_Page 
            Height          =   375
            Left            =   960
            TabIndex        =   5
            Top             =   420
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   661
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
            Text            =   "99"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin CisText_V60.CisText iB1_Noubn 
            Height          =   375
            Left            =   90
            TabIndex        =   4
            Top             =   420
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
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
         End
         Begin CisText_V60.CisText iB1_MPage 
            Height          =   375
            Left            =   1530
            TabIndex        =   6
            Top             =   420
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   661
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
            Text            =   "99"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1350
            TabIndex        =   99
            Top             =   510
            Width           =   135
         End
      End
   End
   Begin VB.PictureBox B1_Area2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4755
      Left            =   2228
      ScaleHeight     =   4695
      ScaleWidth      =   10845
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   5250
      Width           =   10905
      Begin VB.PictureBox B1_Area9 
         Appearance      =   0  'Ã◊Øƒ
         BackColor       =   &H80000004&
         BorderStyle     =   0  'Ç»Çµ
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4545
         Left            =   120
         ScaleHeight     =   4545
         ScaleWidth      =   10635
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   60
         Width           =   10635
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   14
            Left            =   1500
            Top             =   90
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "ïi     î‘"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   10
            Left            =   4995
            Top             =   540
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   375
               Index           =   0
               Left            =   240
               TabIndex        =   19
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D B12_Syuyo_Back 
            Height          =   555
            Index           =   0
            Left            =   3720
            Top             =   540
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   375
               Index           =   0
               Left            =   270
               TabIndex        =   18
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   18
            Left            =   3705
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "é˚óeêî"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   17
            Left            =   4980
            Top             =   90
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
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
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   15
            Left            =   525
            Top             =   90
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "îwî‘çÜ"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   48
            Left            =   540
            Top             =   540
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   375
               Index           =   0
               Left            =   210
               TabIndex        =   14
               Top             =   90
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXX"
               MaxLength       =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   29
            Left            =   1515
            Top             =   540
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hin1 
               Height          =   375
               Index           =   0
               Left            =   240
               TabIndex        =   15
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB12_Hin2 
               Height          =   375
               Index           =   0
               Left            =   960
               TabIndex        =   16
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB12_Hin3 
               Height          =   375
               Index           =   0
               Left            =   1680
               TabIndex        =   17
               Top             =   90
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   20
            Left            =   6015
            Top             =   90
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "î[ì¸êî"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   66
            Left            =   6030
            Top             =   540
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   375
               Index           =   0
               Left            =   270
               TabIndex        =   20
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   6
            Left            =   90
            Top             =   90
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "áÇ"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   525
            Index           =   59
            Left            =   90
            Top             =   540
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   926
            BackColor       =   14737632
            Caption         =   "ÇP"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   435
            Index           =   12
            Left            =   7620
            Top             =   90
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   767
            BackColor       =   14737632
            ForeColor       =   16711680
            Caption         =   "îı     çl"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D Blb_BikoNM 
            Height          =   555
            Index           =   0
            Left            =   7620
            Top             =   540
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
            Caption         =   "ë≈êÿÇËëŒè€"
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
            cPositionX      =   400
            cPositionY      =   120
            Begin CisText_V60.CisText iB12_Biko 
               Height          =   375
               Index           =   0
               Left            =   90
               TabIndex        =   21
               Top             =   90
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               cFormat         =   "#"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9"
               MaxLength       =   1
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#"
               cILength        =   1
            End
            Begin CisText_V60.CisText iB14_Bunsu 
               Height          =   375
               Index           =   0
               Left            =   2010
               TabIndex        =   22
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   7
            Left            =   4995
            Top             =   1110
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   28
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D B12_Syuyo_Back 
            Height          =   555
            Index           =   1
            Left            =   3720
            Top             =   1110
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   375
               Index           =   1
               Left            =   270
               TabIndex        =   27
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   8
            Left            =   540
            Top             =   1110
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   375
               Index           =   1
               Left            =   210
               TabIndex        =   23
               Top             =   90
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXX"
               MaxLength       =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   9
            Left            =   1515
            Top             =   1110
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hin1 
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   24
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB12_Hin2 
               Height          =   375
               Index           =   1
               Left            =   960
               TabIndex        =   25
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB12_Hin3 
               Height          =   375
               Index           =   1
               Left            =   1680
               TabIndex        =   26
               Top             =   90
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   11
            Left            =   6030
            Top             =   1110
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   375
               Index           =   1
               Left            =   270
               TabIndex        =   29
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   525
            Index           =   13
            Left            =   90
            Top             =   1110
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   926
            BackColor       =   14737632
            Caption         =   "ÇQ"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D Blb_BikoNM 
            Height          =   555
            Index           =   1
            Left            =   7620
            Top             =   1110
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
            Caption         =   "ë≈êÿÇËëŒè€"
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
            cPositionX      =   400
            cPositionY      =   120
            Begin CisText_V60.CisText iB12_Biko 
               Height          =   375
               Index           =   1
               Left            =   90
               TabIndex        =   30
               Top             =   90
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               cFormat         =   "#"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9"
               MaxLength       =   1
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#"
               cILength        =   1
            End
            Begin CisText_V60.CisText iB14_Bunsu 
               Height          =   375
               Index           =   1
               Left            =   2010
               TabIndex        =   31
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   21
            Left            =   4995
            Top             =   1680
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   375
               Index           =   2
               Left            =   240
               TabIndex        =   37
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D B12_Syuyo_Back 
            Height          =   555
            Index           =   2
            Left            =   3720
            Top             =   1680
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   375
               Index           =   2
               Left            =   270
               TabIndex        =   36
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   22
            Left            =   540
            Top             =   1680
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   375
               Index           =   2
               Left            =   210
               TabIndex        =   32
               Top             =   90
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXX"
               MaxLength       =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   23
            Left            =   1515
            Top             =   1680
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hin1 
               Height          =   375
               Index           =   2
               Left            =   240
               TabIndex        =   33
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB12_Hin2 
               Height          =   375
               Index           =   2
               Left            =   960
               TabIndex        =   34
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB12_Hin3 
               Height          =   375
               Index           =   2
               Left            =   1680
               TabIndex        =   35
               Top             =   90
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   24
            Left            =   6030
            Top             =   1680
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   375
               Index           =   2
               Left            =   270
               TabIndex        =   38
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   525
            Index           =   25
            Left            =   90
            Top             =   1680
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   926
            BackColor       =   14737632
            Caption         =   "ÇR"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D Blb_BikoNM 
            Height          =   555
            Index           =   2
            Left            =   7620
            Top             =   1680
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
            Caption         =   "ë≈êÿÇËëŒè€"
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
            cPositionX      =   400
            cPositionY      =   120
            Begin CisText_V60.CisText iB12_Biko 
               Height          =   375
               Index           =   2
               Left            =   90
               TabIndex        =   39
               Top             =   90
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               cFormat         =   "#"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9"
               MaxLength       =   1
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#"
               cILength        =   1
            End
            Begin CisText_V60.CisText iB14_Bunsu 
               Height          =   375
               Index           =   2
               Left            =   2010
               TabIndex        =   40
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   26
            Left            =   4995
            Top             =   2250
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   375
               Index           =   3
               Left            =   240
               TabIndex        =   46
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D B12_Syuyo_Back 
            Height          =   555
            Index           =   3
            Left            =   3720
            Top             =   2250
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   375
               Index           =   3
               Left            =   270
               TabIndex        =   45
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   27
            Left            =   540
            Top             =   2250
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   375
               Index           =   3
               Left            =   210
               TabIndex        =   41
               Top             =   90
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXX"
               MaxLength       =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   28
            Left            =   1515
            Top             =   2250
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hin1 
               Height          =   375
               Index           =   3
               Left            =   240
               TabIndex        =   42
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB12_Hin2 
               Height          =   375
               Index           =   3
               Left            =   960
               TabIndex        =   43
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB12_Hin3 
               Height          =   375
               Index           =   3
               Left            =   1680
               TabIndex        =   44
               Top             =   90
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   30
            Left            =   6030
            Top             =   2250
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   375
               Index           =   3
               Left            =   270
               TabIndex        =   47
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   525
            Index           =   31
            Left            =   90
            Top             =   2250
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   926
            BackColor       =   14737632
            Caption         =   "ÇS"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D Blb_BikoNM 
            Height          =   555
            Index           =   3
            Left            =   7620
            Top             =   2250
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
            Caption         =   "ë≈êÿÇËëŒè€"
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
            cPositionX      =   400
            cPositionY      =   120
            Begin CisText_V60.CisText iB12_Biko 
               Height          =   375
               Index           =   3
               Left            =   90
               TabIndex        =   48
               Top             =   90
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               cFormat         =   "#"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9"
               MaxLength       =   1
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#"
               cILength        =   1
            End
            Begin CisText_V60.CisText iB14_Bunsu 
               Height          =   375
               Index           =   3
               Left            =   2010
               TabIndex        =   49
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   32
            Left            =   4995
            Top             =   2820
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   375
               Index           =   4
               Left            =   240
               TabIndex        =   55
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D B12_Syuyo_Back 
            Height          =   555
            Index           =   4
            Left            =   3720
            Top             =   2820
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   375
               Index           =   4
               Left            =   270
               TabIndex        =   54
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   33
            Left            =   540
            Top             =   2820
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   375
               Index           =   4
               Left            =   210
               TabIndex        =   50
               Top             =   90
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXX"
               MaxLength       =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   34
            Left            =   1515
            Top             =   2820
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hin1 
               Height          =   375
               Index           =   4
               Left            =   240
               TabIndex        =   51
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB12_Hin2 
               Height          =   375
               Index           =   4
               Left            =   960
               TabIndex        =   52
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB12_Hin3 
               Height          =   375
               Index           =   4
               Left            =   1680
               TabIndex        =   53
               Top             =   90
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   35
            Left            =   6030
            Top             =   2820
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   375
               Index           =   4
               Left            =   270
               TabIndex        =   56
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   525
            Index           =   36
            Left            =   90
            Top             =   2820
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   926
            BackColor       =   14737632
            Caption         =   "ÇT"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D Blb_BikoNM 
            Height          =   555
            Index           =   4
            Left            =   7620
            Top             =   2820
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
            Caption         =   "ë≈êÿÇËëŒè€"
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
            cPositionX      =   400
            cPositionY      =   120
            Begin CisText_V60.CisText iB12_Biko 
               Height          =   375
               Index           =   4
               Left            =   90
               TabIndex        =   57
               Top             =   90
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               cFormat         =   "#"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9"
               MaxLength       =   1
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#"
               cILength        =   1
            End
            Begin CisText_V60.CisText iB14_Bunsu 
               Height          =   375
               Index           =   4
               Left            =   2010
               TabIndex        =   58
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   37
            Left            =   4995
            Top             =   3390
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   375
               Index           =   5
               Left            =   240
               TabIndex        =   64
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D B12_Syuyo_Back 
            Height          =   555
            Index           =   5
            Left            =   3720
            Top             =   3390
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   375
               Index           =   5
               Left            =   270
               TabIndex        =   63
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   38
            Left            =   540
            Top             =   3390
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   375
               Index           =   5
               Left            =   210
               TabIndex        =   59
               Top             =   90
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXX"
               MaxLength       =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   39
            Left            =   1515
            Top             =   3390
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hin1 
               Height          =   375
               Index           =   5
               Left            =   240
               TabIndex        =   60
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB12_Hin2 
               Height          =   375
               Index           =   5
               Left            =   960
               TabIndex        =   61
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB12_Hin3 
               Height          =   375
               Index           =   5
               Left            =   1680
               TabIndex        =   62
               Top             =   90
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   40
            Left            =   6030
            Top             =   3390
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   375
               Index           =   5
               Left            =   270
               TabIndex        =   65
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   525
            Index           =   41
            Left            =   90
            Top             =   3390
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   926
            BackColor       =   14737632
            Caption         =   "ÇU"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D Blb_BikoNM 
            Height          =   555
            Index           =   5
            Left            =   7620
            Top             =   3390
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
            Caption         =   "ë≈êÿÇËëŒè€"
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
            cPositionX      =   400
            cPositionY      =   120
            Begin CisText_V60.CisText iB12_Biko 
               Height          =   375
               Index           =   5
               Left            =   90
               TabIndex        =   66
               Top             =   90
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               cFormat         =   "#"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9"
               MaxLength       =   1
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#"
               cILength        =   1
            End
            Begin CisText_V60.CisText iB14_Bunsu 
               Height          =   375
               Index           =   5
               Left            =   2010
               TabIndex        =   67
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   42
            Left            =   4995
            Top             =   3960
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hako 
               Height          =   375
               Index           =   6
               Left            =   240
               TabIndex        =   73
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cGFormat        =   "###"
               cILength        =   3
            End
         End
         Begin Cis3D_v60.CIS3D B12_Syuyo_Back 
            Height          =   555
            Index           =   6
            Left            =   3720
            Top             =   3960
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Syuyo 
               Height          =   375
               Index           =   6
               Left            =   270
               TabIndex        =   72
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               cFormat         =   "#####"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "99999"
               MaxLength       =   5
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "#####"
               cILength        =   5
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   43
            Left            =   540
            Top             =   3960
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Seban 
               Height          =   375
               Index           =   6
               Left            =   210
               TabIndex        =   68
               Top             =   90
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXX"
               MaxLength       =   3
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   44
            Left            =   1515
            Top             =   3960
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Hin1 
               Height          =   375
               Index           =   6
               Left            =   240
               TabIndex        =   69
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB12_Hin2 
               Height          =   375
               Index           =   6
               Left            =   960
               TabIndex        =   70
               Top             =   90
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
            End
            Begin CisText_V60.CisText iB12_Hin3 
               Height          =   375
               Index           =   6
               Left            =   1680
               TabIndex        =   71
               Top             =   90
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   555
            Index           =   45
            Left            =   6030
            Top             =   3960
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   979
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
            Begin CisText_V60.CisText iB12_Su 
               Height          =   375
               Index           =   6
               Left            =   270
               TabIndex        =   74
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "######"
               cILength        =   6
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   525
            Index           =   46
            Left            =   90
            Top             =   3960
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   926
            BackColor       =   14737632
            Caption         =   "ÇV"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D Blb_BikoNM 
            Height          =   555
            Index           =   6
            Left            =   7620
            Top             =   3960
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   979
            Caption         =   "ë≈êÿÇËëŒè€"
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
            cPositionX      =   400
            cPositionY      =   120
            Begin CisText_V60.CisText iB12_Biko 
               Height          =   375
               Index           =   6
               Left            =   90
               TabIndex        =   75
               Top             =   90
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               cFormat         =   "#"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "9"
               MaxLength       =   1
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "#"
               cILength        =   1
            End
            Begin CisText_V60.CisText iB14_Bunsu 
               Height          =   375
               Index           =   6
               Left            =   2010
               TabIndex        =   76
               Top             =   90
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               cFormat         =   "######"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "999999"
               MaxLength       =   6
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cGFormat        =   "######"
               cILength        =   6
            End
         End
      End
   End
   Begin VB.Timer Timer2 
      Left            =   2228
      Top             =   1890
   End
   Begin VB.Timer Timer1 
      Left            =   2228
      Top             =   1920
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      Left            =   2228
      TabIndex        =   77
      Text            =   "Text1"
      Top             =   1980
      Width           =   180
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
      Caption         =   "Åy èoâ◊É`ÉPÉbÉgì¸óÕ  Åz"
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
      cPositionX      =   1200
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'Ç»Çµ
         Height          =   345
         Index           =   1
         Left            =   10140
         ScaleHeight     =   345
         ScaleWidth      =   3615
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   30
         Width           =   3615
         Begin Cis3D_v60.CIS3D UC_3D6 
            Height          =   345
            Left            =   30
            Top             =   0
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   609
            BackColor       =   8388736
            ForeColor       =   16777215
            Caption         =   "ì«éÊåèêî"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cFont3D         =   2
            cAlingnment     =   1
            cPositionX      =   90
            Begin Cis3D_v60.CIS3D H1lb_OCRCnt 
               Height          =   255
               Left            =   1110
               Top             =   60
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   8388736
               ForeColor       =   16777215
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
               cFont3D         =   2
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D7 
            Height          =   345
            Left            =   1830
            Top             =   0
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   609
            BackColor       =   8454143
            ForeColor       =   0
            Caption         =   "éËì¸óÕåèêî"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Left            =   1200
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
      End
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Index           =   1
         Left            =   14220
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   7560
         Top             =   60
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   503
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
   End
   Begin CisText_V60.CisText iB14_Yobi1 
      Height          =   345
      Left            =   8858
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   4290
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   609
      cFormat         =   "#"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "9"
      MaxLength       =   1
      cDataType       =   1
      cDataReplace    =   1
      cFbComma        =   0
      cGFormat        =   "#"
      cILength        =   1
   End
   Begin CisText_V60.CisText iB14_Bun 
      Height          =   345
      Left            =   10208
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   4290
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "X"
      MaxLength       =   1
      cDataReplace    =   1
      cFbComma        =   0
   End
   Begin CisText_V60.CisText iB14_HBin 
      Height          =   345
      Left            =   9998
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   609
      cFormat         =   "00"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "99"
      MaxLength       =   2
      cDataType       =   1
      cDataReplace    =   1
      cFaZero         =   0
      cFdAutoFormat   =   1
      cGFormat        =   "##"
      cILength        =   2
   End
   Begin CisText_V60.CisText iB14_MotoYobi 
      Height          =   345
      Left            =   10838
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   4290
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   609
      cFormat         =   "####"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "9999"
      MaxLength       =   4
      cDataType       =   1
      cDataReplace    =   1
      cFbComma        =   0
      cGFormat        =   "####"
      cILength        =   4
   End
   Begin CisText_V60.CisText iB14_Kouku 
      Height          =   345
      Left            =   10478
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   4290
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      cFormat         =   "##"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "99"
      MaxLength       =   2
      cDataType       =   1
      cDataReplace    =   1
      cFbComma        =   0
      cGFormat        =   "##"
      cILength        =   2
   End
   Begin CisText_V60.CisText iB14_HKB 
      Height          =   345
      Left            =   9938
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   4290
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   609
      cFormat         =   "#"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "9"
      MaxLength       =   1
      cDataType       =   1
      cDataReplace    =   1
      cFbComma        =   0
      cGFormat        =   "#"
      cILength        =   1
   End
   Begin CisText_V60.CisText iB14_Gyo 
      Height          =   345
      Left            =   9668
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   4290
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   609
      cFormat         =   "#"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "9"
      MaxLength       =   1
      cDataType       =   1
      cDataReplace    =   1
      cFbComma        =   0
      cGFormat        =   "#"
      cILength        =   1
   End
   Begin CisText_V60.CisText iB14_Yobi3 
      Height          =   345
      Left            =   9398
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   4290
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   609
      cFormat         =   "#"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "9"
      MaxLength       =   1
      cDataType       =   1
      cDataReplace    =   1
      cFbComma        =   0
      cGFormat        =   "#"
      cILength        =   1
   End
   Begin CisText_V60.CisText iB14_Yobi2 
      Height          =   345
      Left            =   9128
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   4290
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   609
      cFormat         =   "#"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "9"
      MaxLength       =   1
      cDataType       =   1
      cDataReplace    =   1
      cFbComma        =   0
      cGFormat        =   "#"
      cILength        =   1
   End
   Begin CisYMD_V60.CisYMD iB14_HYmd 
      Height          =   330
      Left            =   8258
      TabIndex        =   96
      Top             =   3960
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3069
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Object.Width           =   1755
      cChkResult      =   0   'False
   End
   Begin Cis3D_v60.CIS3D K_Sykbnm 
      Height          =   795
      Left            =   2228
      Top             =   900
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   1402
      Caption         =   "ì«éÊ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
   End
   Begin VB.Label ReadColor 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ì«éÊíÜêF"
      Height          =   255
      Left            =   90
      TabIndex        =   101
      Top             =   1020
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "ÇRÅFë≈êÿÇËäÆóπ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   11378
      TabIndex        =   85
      Top             =   10020
      Width           =   1755
   End
   Begin VB.Label Label3 
      Caption         =   "ÇQÅFë≈êÿíÜ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   10133
      TabIndex        =   84
      Top             =   10020
      Width           =   1245
   End
   Begin VB.Label Label3 
      Caption         =   "ÇPÅFë≈êÿÇËëŒè€"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   8348
      TabIndex        =   83
      Top             =   10020
      Width           =   1785
   End
   Begin VB.Label Label2 
      Caption         =   "ÅyîıçlÅz"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7658
      TabIndex        =   82
      Top             =   10020
      Width           =   5475
   End
End
Attribute VB_Name = "CXJ0250"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :  èoâ◊É`ÉPÉbÉgì¸óÕ
'**       ÉtÉHÅ[ÉÄID    :  CXJ0250
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :  2005/06/01   By CIS
'**       ïœ  çX  ì˙    :  2006/03/30   By CIS  ã@î‘-âpêîéöÅAé©é–ì`ï[É`ÉFÉbÉNïœçX
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim ProcHB          As String * 2   ' ì¸óÕàÊêÿë÷
    Dim FormAct         As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim SyoriKB         As String       ' èàóùãÊï™(í«â¡•èCê≥•çÌèú)
    Dim SyoriNM         As String       ' èàóùñºèÃ(í«â¡•èCê≥•çÌèú)
    Dim ErrCnt          As Integer
    Dim Cnt_OCR         As Integer      ' OCRèàóùåèêî
    Dim Cnt_ADD         As Integer      ' í«â¡èàóùåèêî
    
    Dim wObj            As Object
    Dim wNum1           As Integer
    Dim wNum2           As Integer
    Dim wNumL           As Long
    Dim wStr1           As String
    Dim wStr2           As String
    Dim wStr3           As String

    Dim mERMsg          As Boolean
    Dim mFocusObj       As Object
'   * ì`ï[ì‡óeëﬁîÉèÅ[ÉN *
    Dim wcc_Head        As String
    Dim wcc_Body        As String
    Dim wGyo            As Integer
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås    ÇkÇèÇÅÇÑÉCÉxÉìÉg    Åt              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   # ìÒ èd ãN ìÆ ñh é~ #
'   #------------------#
    If Not Dupli_Start Then End
'   #------------------#
'   # èâ ä˙ ì‡ óe éÊ ìæ #
'   #------------------#
    If Not IniGet Then End
'   #------------------#
'   #  çÄ ñ⁄ èâ ä˙ âª   #
'   #------------------#
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld

    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If

    Dummy.Left = -1000
    Dummy.Enabled = False

    Timer1.Interval = 1000   ' éûçèï\é¶ä‘äu
    Timer1.Enabled = True    ' éûçèï\é¶äJén
    Timer2.Interval = 500    ' èàóùãÊï™ì_ñ≈ä‘äu
    Timer2.Enabled = True    ' èàóùãÊï™ì_ñ≈ã÷é~
'
'   #------------------#
'   #    èâä˙ì¸óÕê›íË   #
'   #------------------#
    Call QT10QSRInfoGet
    ProcHB = "H1"
    If Not QT10QSR_UseFlg Then
       SyoriKB = "ADD"
    Else
       SyoriKB = "RED"
    End If
    Call DispChange(ProcHB)
    Call SyoriName(SyoriKB)
'   #-------------------#
'   # âÊñ íÜâõà íuï\é¶   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    Me.Show
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    ÅsÇ`ÇÉÇîÇâÇñÇÅÇîÇÖ ÉCÉxÉìÉgÅt              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    Cnt_OCR = 0: Cnt_ADD = 0
   
'    iH1_YMD.cInitDate = True
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås QueryUnloadÉCÉxÉìÉg Åt                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås ÇjÇÖÇôÇcÇèÇóÇéÉCÉxÉìÉg Åt               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄÅs ÉtÉ@ÉìÉNÉVÉáÉìÉLÅ[ópÉAÉNÉZÉâÉåÅ[É^ Åt       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    If Not FormAct Then Exit Function
    If UCase(IVENT) <> "DOWN" Then Exit Function
    
    Key_Acc = False
    
    Select Case Key_Code
        Case vbKeyF1:       If PB_ADD.Visible Then Call PB_Add_Click    'ÅyêVãKÅz
        Case vbKeyF5:       If PB_RED.Visible Then Call PB_RED_Click    'Åyì«éÊÅz
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  'ÅyåüçıÅz
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    'ÅyèIóπÅz
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_Can_Click    'ÅyéÊè¡Åz
        Case vbKeyReturn
                            If PB_ENT.Visible Then                      'Åyì¸óÕÅz
                                Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                Call PB_ENT_Click
                                Dummy.Enabled = False
                            End If
        Case Else:          Exit Function
    End Select
    
    Key_Acc = True
End Function
'+---------------------+
'+    èI  óπ  ÉL Å[     +
'+---------------------+
Private Sub PB_END_Click()
'   #### √ﬁ∞¿Õﬁ∞Ωêÿíf ####
    Call CisDB.DBDISConnect
'   #### Ã´∞—µÃﬁºﬁ™∏ƒè¡ãé ####
    Call CisFrm.UnLoadDisp
    End
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
Private Sub PB_Can_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call CisFrm.MousePT(11)
    
    If ProcHB = "B1" Then
        If SyoriKB <> "RED" Then
           ProcHB = "H1"
           Call DispChange(ProcHB)
           Call SyoriName(SyoriKB)
        Else
           ProcHB = "B2"
           Call DispChange(ProcHB)
           Call SyoriName(SyoriKB)
        End If
        Call HeadBodyClear("B")
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B2" Then
        ProcHB = "H1"
        Call HeadBodyClear("B")
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    êV  ãK  ÉL Å[(F1)       +
'+----------------------------+
Private Sub PB_Add_Click()
    SyoriKB = "ADD"
    Call SyoriName(SyoriKB)
    If ProcHB = "B2" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
    End If
End Sub
'+----------------------------+
'+      ì«éÊ ÉL Å[(F5)         +
'+----------------------------+
Private Sub PB_RED_Click()
    SyoriKB = "RED"
    Call SyoriName(SyoriKB)
    If ProcHB = "B2" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
    End If
End Sub
'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_YMD" Then iH1_YMD.ShowCalender: Exit Sub
    If PB_Look.Tag = "iB13_YMD" Then iB13_YMD.ShowCalender: Exit Sub
    If PB_Look.Tag Like "iB12_Hin*" Then Call RefHinbn(Right(PB_Look.Tag, 1))
End Sub
'+------------------------------------+'
'+                                    +
'+      éQ               è∆           +
'+                                    +
'+------------------------------------+'
'+-------------------------------+
'+    ïi  î‘  åü  çı              +
'+-------------------------------+
Private Sub RefHinbn(Index As Integer)

    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 0
    RVI_SK = 9
    RVI_Torcd = ""
    RVI_Ukeir = ""
    RVI_Seban = ""
    RVI_Tehai = ""

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iB12_Hin1(Index) = Mid(RV_Hinbn, 1, 5)
        iB12_Hin2(Index) = Mid(RV_Hinbn, 6, 5)
        iB12_Hin3(Index) = Mid(RV_Hinbn, 11, 2)
        If Trim(Left(RV_Seban, 1)) <> "" Then
           iB12_Seban(Index) = Left(RV_Seban, 3)
        Else
           iB12_Seban(Index) = Right(RV_Seban, 3)
        End If
    End If

End Sub
'************************************'
'*        âÊñ ÉtÉBÅ[ÉãÉh
'************************************'
'*---------------------------------------------------*
'*      iH1_YMD   ( GotFocus )        èoâ◊ì˙
'*---------------------------------------------------*
Private Sub iH1_YMD_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
'*---------------------------------------------------*
'*      iH1_YMD   ( LostFocus )       èoâ◊ì˙
'*---------------------------------------------------*
Private Sub iH1_YMD_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Visible = False
        PB_Look.Tag = ""
    End If
End Sub
'*---------------------------------------------------*
'*      iB13_YMD   ( GotFocus )        î[ì¸ì˙
'*---------------------------------------------------*
Private Sub iB13_YMD_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
'*---------------------------------------------------*
'*      iB13_YMD   ( LostFocus )       î[ì¸ì˙
'*---------------------------------------------------*
Private Sub iB13_YMD_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Visible = False
        PB_Look.Tag = ""
    End If
End Sub
'*---------------------------------------------------*
'*      iB12_Hin1   ( GotFocus )        ïiî‘ÇP
'*---------------------------------------------------*
Private Sub iB12_Hin1_GotFocus(Index As Integer)
    PB_Look.Tag = ActiveControl.Name & Index
    PB_Look.Visible = True
End Sub
'*---------------------------------------------------*
'*      iB12_Hin1   ( LostFocus )       ïiî‘ÇP
'*---------------------------------------------------*
Private Sub iB12_Hin1_LostFocus(Index As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Visible = False
        PB_Look.Tag = ""
    End If
End Sub
'*---------------------------------------------------*
'*      iB12_Hin2   ( GotFocus )        ïiî‘ÇQ
'*---------------------------------------------------*
Private Sub iB12_Hin2_GotFocus(Index As Integer)
    PB_Look.Tag = ActiveControl.Name & Index
    PB_Look.Visible = True
End Sub
'*---------------------------------------------------*
'*      iB12_Hin2   ( LostFocus )       ïiî‘ÇQ
'*---------------------------------------------------*
Private Sub iB12_Hin2_LostFocus(Index As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Visible = False
        PB_Look.Tag = ""
    End If
End Sub
'*---------------------------------------------------*
'*      iB12_Hin3   ( GotFocus )        ïiî‘ÇR
'*---------------------------------------------------*
Private Sub iB12_Hin3_GotFocus(Index As Integer)
    PB_Look.Tag = ActiveControl.Name & Index
    PB_Look.Visible = True
End Sub
'*---------------------------------------------------*
'*      iB12_Hin3   ( LostFocus )       ïiî‘ÇR
'*---------------------------------------------------*
Private Sub iB12_Hin3_LostFocus(Index As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Visible = False
        PB_Look.Tag = ""
    End If
End Sub
'*---------------------------------------------------*
'*      iB12_Hako   ( LostFocus )       î†êî
'*---------------------------------------------------*
Private Sub iB12_Hako_LostFocus(Index As Integer)
    If RTrim(iB12_Hin1(Index)) = "" And RTrim(iB12_Hin2(Index)) = "" And RTrim(iB12_Hin3(Index)) = "" Then Exit Sub
'   * è»ó™éû ÇP *
    If CisFun.Val2(iB12_Hako(Index)) = 0 Then
       If CisFun.Val2(iB12_Syuyo(Index)) = 0 Then
          iB12_Hako(Index) = "  1"
       Else
          gInt = CisFun.Val2(iB12_Su(Index)) \ CisFun.Val2(iB12_Syuyo(Index))
          gLong = CisFun.Val2(iB12_Su(Index)) Mod CisFun.Val2(iB12_Syuyo(Index))
          If gLong <> 0 Then
             gInt = gInt + 1
          End If
          iB12_Hako(Index) = CisFun.RSetFld(gInt, 3, "###")
       End If
    End If
'   * î[ì¸êîåvéZ *
    If CisFun.Val2(iB12_Su(Index)) <> 0 Then Exit Sub
    wNumL = CisFun.Val2(iB12_Hako(Index)) * CisFun.Val2(iB12_Syuyo(Index))
    iB12_Su(Index) = Format(Format(wNumL, "#"), String(6, "@"))
End Sub
'*---------------------------------------------------*
'*      iB12_Syuyo  ( LostFocus )       é˚óeêî
'*---------------------------------------------------*
Private Sub iB12_Syuyo_LostFocus(Index As Integer)
    If RTrim(iB12_Hin1(Index)) = "" And RTrim(iB12_Hin2(Index)) = "" And RTrim(iB12_Hin3(Index)) = "" Then Exit Sub
'   * è»ó™éû ÇP *
    If CisFun.Val2(iB12_Syuyo(Index)) = 0 Then iB12_Syuyo(Index) = Format(Format(CisFun.Val2(iB12_Su(Index)), "#"), String(5, "@"))

'   * î[ì¸êîåvéZ *
    If CisFun.Val2(iB12_Su(Index)) <> 0 Then Exit Sub
    wNumL = CisFun.Val2(iB12_Hako(Index)) * CisFun.Val2(iB12_Syuyo(Index))
    iB12_Su(Index) = Format(Format(wNumL, "#"), String(6, "@"))
End Sub
'*---------------------------------------------------*
'*      iB12_Su     ( LostFocus )       î[ì¸êî
'*---------------------------------------------------*
Private Sub iB12_Su_LostFocus(Index As Integer)
    If CisFun.Val2(iB12_Su(Index)) <> 0 Then Exit Sub

    wNumL = CisFun.Val2(iB12_Hako(Index)) * CisFun.Val2(iB12_Syuyo(Index))
    iB12_Su(Index) = Format(Format(wNumL, "#"), String(6, "@"))
End Sub
'*---------------------------------------------------*
'*      iB12_Biko   ( LostFocus )       îıçl
'*---------------------------------------------------*
Private Sub iB12_Biko_LostFocus(Index As Integer)
    Select Case iB12_Biko(Index)
        Case "1": Blb_BikoNM(Index) = "ë≈êÿÇËëŒè€"
        Case "2": Blb_BikoNM(Index) = "ë≈êÿÇËíÜ"
        Case "3": Blb_BikoNM(Index) = "ë≈êÿÇËäÆóπ"
        Case Else
            Blb_BikoNM(Index) = ""
            iB12_Biko(Index) = ""
            iB14_Bunsu(Index) = ""
    End Select
End Sub


'**************************************
'*                                    *
'*      èàóùãÊï™Ç…ÇÊÇËãÊï™ñºï\é¶        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   
    Dim W_Color     As Long
    
    W_Color = &H80000008 'èâä˙êFæØƒ(çï)
    K_Sykbnm.BackColor = &HC0C0C0
    Select Case SyoriKB
     Case "RED"
          SyoriNM = "QRì«éÊ"
          K_Sykbnm.ForeColor = W_Color   '(çï)
          PB_ADD.Visible = True
          PB_RED.Visible = False
     Case "ADD"
          SyoriNM = "éËì¸óÕ"
          K_Sykbnm.ForeColor = &HFFFF&   '(â©êF)
          PB_ADD.Visible = False
          PB_RED.Visible = True
     Case Else
          SyoriNM = ""
    End Select
    
    Timer2.Enabled = False
    If SyoriNM = "" Then
        K_Sykbnm.Visible = False
        K_Sykbnm.Caption = ""
        Exit Sub
    End If
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    Timer2.Enabled = True
    
    If Not QT10QSR_UseFlg Then
       PB_ADD.Visible = False
       PB_RED.Visible = False
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
    Dim H1Mode      As Boolean
    Dim H1Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    Dim B2Mode      As Boolean
    Dim B2Color     As Long

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
    B2Mode = False: B2Color = gPLostSel
' ¿ﬁ–∞çÄñ⁄(âÊñ ÇÃÉ`ÉâÉcÉLñhé~)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1":   H1Mode = True:  H1Color = gPGotSel
       Case "B1":   B1Mode = True:  B1Color = gPGotSel
       Case "B2":   B2Mode = True:  B2Color = gPGotSel
    End Select
        
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B1_Area2.Enabled = B1Mode
    B1_Area2.BackColor = B1Color
    B1_Area3.Enabled = B1Mode
    B1_Area3.BackColor = B1Color

    PB_CAN.Visible = Not H1Mode

    If ProcHB = "B2" Then
        SyoriNM = "Åy ì« éÊ íÜ Åz"
        K_Sykbnm.BackColor = ReadColor.BackColor
        Set QT10QSR_Comm = PDQComm1
        Call QT10QSRInfoGet
        If Not QT10QSR_Open Then
           End
        End If
    End If
    
    PB_ADD.Visible = H1Mode
    PB_RED.Visible = H1Mode
    
    PB_ENT.Visible = Not B2Mode
    
    If Not QT10QSR_UseFlg Then
       PB_ADD.Visible = False
       PB_RED.Visible = False
    End If
' ¿ﬁ–∞çÄñ⁄(ì¸óÕã÷é~)
    Dummy.Enabled = False
'
End Function
'**************************************
'*      Head/Body èâä˙âªèàóù           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String

    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = "" ': .Enabled = True
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Cmb_*" Then .Clear
        End With
    Next gObj
End Sub
'****************************
'*      ì¸óÕèàóù(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
'   * èoâ◊ì˙ *
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        If SyoriKB = "RED" Then
            ProcHB = "B2"
            Call DispChange(ProcHB)
            Call ChiketoProc                         ' QRì«éÊèàóù
            GoTo ReturnPress_Ed
        End If
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If

'   * ì`ï[ì¸óÕ *
    If ProcHB = "B1" Then
        Call Line_Cut
        If Not ContentsChk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        
        If SyoriKB = "ADD" Then ProcHB = "B1" Else ProcHB = "B2"
        
        Call DispChange(ProcHB)
        
        GoTo ReturnPress_Ed
    End If

ReturnPress_Ed:
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1¡™Ø∏èàóù(åüé˚ì˙ÅEé–ì‡ãÊï™)
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
 
    If Trim(iH1_YMD) <> "" Then
        If Not iH1_YMD.cChkResult Then
            With CisFun
                .MB_Lines = 5
                .MB_Button = Error
                .MB_Title = ""
                .MB_MSG(2) = "    èoâ◊ì˙Çê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢        "
                .MBOX
            End With
            Exit Function
        End If
    End If
'
    Head1Chk = True
End Function
'*******************************
'*      É^ÉCÉgÉãÉoÅ[ï“èW        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*      èàóùãÊï™ñºì_ñ≈êßå‰      *
'*******************************
Private Sub Timer2_Timer()
   If K_Sykbnm.Caption <> "" Then
      K_Sykbnm.Caption = ""
   Else
      K_Sykbnm.Caption = SyoriNM
   End If
End Sub
'**************************************
'*                                    *
'*      Ç©ÇÒÇŒÇÒÉ`ÉPÉbÉgì«éÊèàóù      *
'*                                    *
'**************************************
Private Function ChiketoProc()
    
    ChiketoProc = False
' 232C µÃﬁºﬁ™∏ƒèâä˙ê›íË
    Set QT10QSR_Comm = PDQComm1
    Call QT10QSRInfoGet
    If Not QT10QSR_Open Then
        Call PB_Can_Click
    End If
    
    On Error GoTo 0
    Exit Function

OcrProc_Err:
    Call CisFun.ErrorBox("¡πØƒ QRì«éÊ")
    End
    
End Function
'+------------------------------+
'+  ì«éÊèàóù     *
'+------------------------------+
Private Sub PDQComm1_OnComm()
    Static wPNo     As Integer
    Static Rtn      As String
    Dim wComm       As String
    
    gStr = QT10QSR_Event
    If gStr = "ERR" Then Exit Sub
    If gStr <> "RCV" Then Exit Sub
    
    If Not QT10QSR_RCVProc Then Exit Sub    ' CR Ç™ë∂ç›ÇµÇ»Ç¢éûî≤ÇØÇÈ

    wcc_Head = Mid(QT10QSR_ReadData, 1, 56)
    wcc_Body = Mid(QT10QSR_ReadData, 57)
    
'    Debug.Print QT10QSR_ReadData
'
    If ChekData(0, 0, 1) Then
       H1lb_OCRCnt = Format(CisFun.Val2(H1lb_OCRCnt) + 1, "#,###")
       DoEvents
    Else
       If Not mERMsg Then
           Call QT10QSR_Close
           ProcHB = "B1"
           Call SyoriName(SyoriKB)
           Call DispChange(ProcHB)
        
           On Error Resume Next
           mFocusObj.SetFocus
           Set mFocusObj = Nothing
           On Error GoTo 0
       End If
    End If
    
    QT10QSR_ReadData = ""
End Sub
'+------------------------------------------+
'+
'+------------------------------------------+
Private Function ChekData(ReadType As Integer, OutType As Integer, ChekFlg As Integer) As Boolean

    Dim wErNo           As Integer
    Dim wErGyo          As Integer
    
    ChekData = False
    
    mERMsg = False

    With CisDB
        .SQL = "èoâ◊É`ÉPÉbÉgèàóù"
        .StoadoCount = 10
        .ParaNo = 0: .ParaIO = Returnóp
        .ParaNo = 1: .ParaIO = Inputóp      ' èàóùãÊï™ 0:¡™Ø∏&èoóÕ,1:¡™Ø∏
        .ParaNo = 2: .ParaIO = Inputóp      ' èoóÕãÊï™ 0:ÿ∞¿ﬁ∞,1:éËì¸óÕ,2:ÿ∞¿ﬁ∞&¥◊∞ämîF
        .ParaNo = 3: .ParaIO = Inputóp      ' ämîFÉtÉâÉO
        .ParaNo = 4: .ParaIO = Inputóp      ' íSìñé“
        .ParaNo = 5: .ParaIO = Inputóp      ' èoâ◊ì˙
        .ParaNo = 6: .ParaIO = Inputóp      ' èoâ◊ï÷
        .ParaNo = 7: .ParaIO = Inputóp      ' ÉwÉbÉhì‡óe
        .ParaNo = 8: .ParaIO = InOutóp      ' ì`ï[ÇÃñæç◊ïî(<7ñæç◊ï™>)
        .ParaNo = 9: .ParaIO = OutPutóp     ' É`ÉFÉbÉNÉGÉâÅ[NO
        .ParaNo = 10: .ParaIO = OutPutóp     ' É`ÉFÉbÉNÉGÉâÅ[çsNO
        
        .ParaNo = 1: .ParaValue = ReadType  ' 0:¡™Ø∏&èoóÕ,1:¡™Ø∏
        .ParaNo = 2: .ParaValue = OutType   ' 0:ÿ∞¿ﬁ∞,1:éËì¸óÕ,2:ÿ∞¿ﬁ∞&¥◊∞ämîF
        .ParaNo = 3: .ParaValue = ChekFlg   ' 0:ämîFMSGñ¢èàóù,1:ämîFMSGèàóùçœ
        .ParaNo = 4: .ParaValue = gTanto
        .ParaNo = 5: .ParaValue = iH1_YMD
        .ParaNo = 6: .ParaValue = iH1_BIN
        .ParaNo = 7: .ParaValue = wcc_Head
        .ParaNo = 8: .ParaValue = wcc_Body

        If Not .DBStored Then End
        ' ÉGÉâÅ[î‘çÜ
        .ParaNo = 9: wErNo = CisFun.Val2(.ParaValue)
        ' ÉGÉâÅ[çsNo
        .ParaNo = 10: wErGyo = CisFun.Val2(.ParaValue)
        
    End With
'--              1 : ëŒè€äOÉ`ÉPÉbÉg       2 : î≠íçãÊï™ÉGÉâÅ[
'--              3 : î[ì¸êÊïœä∑ÉGÉâÅ[         4 : î[ì¸êÊñ¢ìoò^
'--              5 : é©é–å¸ì`ï[Ç≈Ç»Ç¢         6 :
'--              7 : é˚óeêîà·Ç¢           8 : î[ì¸êîÉGÉâÅ[(î†êîÅ~é˚óeêîÅÇî[ì¸êî)
'--              15: ïiî‘éÊà¯êÊÉ}ÉXÉ^ñ¢ìoò^
'--              16: éÛíçèdï°ÉGÉâÅ[(éÛíçÉeÅ[ÉuÉã)
'
'--              81: éÛíçä«óùNOälìæÉGÉâÅ[
'--              82: èoâ◊éwé¶èëNOälìæÉGÉâÅ[
'--              85: éÛíçÉfÅ[É^èoóÕÉGÉâÅ[
'--              86: èoâ◊ÉfÅ[É^èoóÕÉGÉâÅ[
'--              99: ÉfÅ[É^çXêVÉGÉâÅ[
    
    With CisFun
        .MB_Lines = 5
        .MB_Title = ""
        .MB_Button = Error
            
        If wErNo = 8 And ChekFlg = 0 Then
            .MB_MSG(1) = "Åy  " & StrConv(wErGyo, vbWide) & " çsñ⁄  Åz"
            .MB_MSG(3) = "   î[ì¸êîÇ™àŸÇ»ÇÈ ( î†êîÅ~é˚óeêî )    "
            Set mFocusObj = iB12_Hako(wErGyo - 1)
        End If
        Select Case wErNo
            Case 1
                    .MB_MSG(2) = "     ëŒè€äOÉ`ÉPÉbÉgÇ≈Ç∑   "
                    Set mFocusObj = iB13_YMD
                    If OutType = 0 Then mERMsg = True
            Case 2
                    .MB_MSG(2) = "     èáèòî[ì¸É`ÉPÉbÉgÇ≈ÇÕÇ†ÇËÇ‹ÇπÇÒÅB      "
                    Set mFocusObj = iB13_YMD
                    If OutType = 0 Then mERMsg = True
            Case 3
                    .MB_MSG(2) = "     î[ì¸êÊïœä∑ÉGÉâÅ[      "
                    Set mFocusObj = iB1_SKTorcd
            Case 4
                    .MB_MSG(2) = "     î[ì¸êÊñ¢ìoò^      "
                    Set mFocusObj = iB1_SKTorcd
            Case 5
                    .MB_MSG(2) = "     é©é–å¸ÇØÉ`ÉPÉbÉgÇ≈ÇÕÇ†ÇËÇ‹ÇπÇÒÅB      "
                    Set mFocusObj = iB1_SRTorcd
                    If OutType = 0 Then mERMsg = True
            Case 9
                    .MB_MSG(2) = "Åy  " & StrConv(wErGyo, vbWide) & " çsñ⁄  Åz"
                    .MB_MSG(4) = "       î[ì¸êîñ¢ì¸óÕ     "
                    Set mFocusObj = iB12_Su(wErGyo - 1)
            Case 15
                    .MB_MSG(2) = "Åy  " & StrConv(wErGyo, vbWide) & " çsñ⁄  Åz"
                    .MB_MSG(4) = "       ïiî‘É}ÉXÉ^ñ¢ìoò^     "
                    Set mFocusObj = iB12_Hin1(wErGyo - 1)
            Case 16
                    .MB_MSG(2) = "       É`ÉPÉbÉgì¸óÕçœÇ›Ç≈Ç∑      "
                    Set mFocusObj = iB1_Noubn
            
            Case 81
                    .MB_MSG(2) = "       éÛíçä«óùNOälìæÉGÉâÅ[      "
                    Set mFocusObj = iB13_YMD
            Case 82
                    .MB_MSG(2) = "       èoâ◊éwé¶èëNOälìæÉGÉâÅ[      "
                    Set mFocusObj = iB13_YMD
            Case 85
                    .MB_MSG(2) = "       éÛíçÉfÅ[É^èoóÕÉGÉâÅ[      "
                    Set mFocusObj = iB13_YMD
            Case 86
                    .MB_MSG(2) = "       èoâ◊ÉfÅ[É^èoóÕÉGÉâÅ[      "
                    Set mFocusObj = iB13_YMD
        End Select
    End With


'   * âÊñ çƒï\é¶ *
    Call HeadBodyClear("B")
' ÉGÉâÅ[ï\é¶å„çƒì«Ç›çûÇ›
    If mERMsg Then
       Call CisFun.MBOX
       Exit Function
    End If
'   * âÊñ çƒï\é¶ *
    iB14_Yobi1 = Mid(wcc_Head, 2, 1)
    iB14_Yobi2 = Mid(wcc_Head, 3, 1)
    iB14_Yobi3 = Mid(wcc_Head, 4, 1)
    iB14_Gyo = Mid(wcc_Head, 5, 1)
    iB14_HKB = Mid(wcc_Head, 6, 1)
    iB1_Noubn = Mid(wcc_Head, 7, 5)
    iB1_Page = CisFun.RSetFld(Format(Mid(wcc_Head, 12, 2)), 0, "##")
    iB14_Bun = Format(Mid(wcc_Head, 14, 1), "#")
    iB1_Sharp = Format(Mid(wcc_Head, 15, 1), "0")
    iB1_SRTorcd = Format(Mid(wcc_Head, 16, 4), "0000")
    iB1_SRKojyo = Mid(wcc_Head, 20, 1)
    iB14_Kouku = Format(Mid(wcc_Head, 21, 2), "00")
    iB1_SKTorcd = Format(Mid(wcc_Head, 23, 4), "0000")
    iB1_Uke = Mid(wcc_Head, 27, 2)
    iB1_Kibn = Mid(wcc_Head, 29, 2)
    iB1_Moto = Format(Mid(wcc_Head, 31, 4), "0000")
    iB14_MotoYobi = Format(Mid(wcc_Head, 35, 4), "0000")
    gStr = Mid(wcc_Head, 39, 6)
    If Trim(gStr) <> "" Then
       If Mid(gStr, 1, 2) > "90" Then
          iB13_YMD = "19" & gStr
       Else
          iB13_YMD = "20" & gStr
       End If
       iB13_Bin = Mid(wcc_Head, 45, 2)
    End If
    gStr = Mid(wcc_Head, 47, 6)
    If Trim(gStr) <> "" Then
       If Mid(gStr, 1, 2) > "90" Then
          iB14_HYmd = "19" & gStr
       Else
          iB14_HYmd = "20" & gStr
       End If
       iB14_HBin = Mid(wcc_Head, 53, 2)
    End If
    iB1_MPage = Format(Mid(wcc_Head, 55, 2), "00")
    '
    wStr1 = wcc_Body
'
    For wNum1 = 0 To 6
        wNum2 = wNum1 * 36
        If RTrim(Mid(wStr1, wNum2 + 1, 15)) = "" Then Exit For
'
        iB12_Seban(wNum1) = RTrim(Mid(wStr1, wNum2 + 1, 3))
        iB12_Hin1(wNum1) = RTrim(Mid(wStr1, wNum2 + 4, 5))
        iB12_Hin2(wNum1) = RTrim(Mid(wStr1, wNum2 + 9, 5))
        iB12_Hin3(wNum1) = RTrim(Mid(wStr1, wNum2 + 14, 2))
        iB12_Syuyo(wNum1) = CisFun.RSetFld(Mid(wStr1, wNum2 + 16, 5), 0, "#####")
        iB12_Hako(wNum1) = CisFun.RSetFld(Mid(wStr1, wNum2 + 21, 3), 0, "###")
'
        If CisFun.Val2(Mid(wStr1, wNum2 + 24, 6)) <> 0 Then
            iB12_Su(wNum1) = CisFun.RSetFld(RTrim(Mid(wStr1, wNum2 + 24, 6)), 0, "######")
        End If
        gStr = RTrim(Mid(wStr1, wNum2 + 30, 1))
        iB12_Biko(wNum1) = Format(gStr, "#")
        Select Case gStr
            Case "1": Blb_BikoNM(wNum1) = "ë≈êÿÇËëŒè€"
            Case "2": Blb_BikoNM(wNum1) = "ë≈êÿÇËíÜ"
            Case "3": Blb_BikoNM(wNum1) = "ë≈êÿÇËäÆóπ"
            Case Else
                Blb_BikoNM(wNum1) = ""
        End Select
        If CisFun.Val2(Mid(wStr1, wNum2 + 31, 6)) <> 0 Then
            iB14_Bunsu(wNum1) = CisFun.RSetFld(RTrim(Mid(wStr1, wNum2 + 31, 6)), 0, "######")
        End If
    Next wNum1
    
    If wErNo = 0 Then
        ChekData = True
        Exit Function
    End If
    
    If wErNo = 16 Then
       mERMsg = True
    End If
'
    Call CisFun.MBOX
    
    If wErNo > 50 Then
        Call PB_END_Click
    End If
    
End Function
'**************************************
'*                                    *
'*          ÇªÇÃëºÉtÉ@ÉìÉNÉVÉáÉì      *
'*                                    *
'**************************************
'+----------------------------------------------------------------------+
'+          ñæ ç◊ çs  çs ãl Çﬂ
'+----------------------------------------------------------------------+
Private Sub Line_Cut()
    Static wlc_Flg      As Boolean
    
    For wNum1 = 0 To 6
        If RTrim(iB12_Hin1(wNum1)) = "" And RTrim(iB12_Hin2(wNum1)) = "" And RTrim(iB12_Hin3(wNum1)) = "" And _
           RTrim(iB12_Seban(wNum1)) = "" And CisFun.Val2(iB12_Hako(wNum1)) = 0 And CisFun.Val2(iB12_Syuyo(wNum1)) = 0 And _
           CisFun.Val2(iB12_Su(wNum1)) = 0 Then

            wlc_Flg = False
            For wNum2 = wNum1 To 6
                If RTrim(iB12_Hin1(wNum2)) <> "" Or RTrim(iB12_Hin2(wNum2)) <> "" Or RTrim(iB12_Hin3(wNum2)) <> "" Or _
                   RTrim(iB12_Seban(wNum2)) <> "" Or CisFun.Val2(iB12_Hako(wNum2)) <> 0 Or CisFun.Val2(iB12_Syuyo(wNum2)) Or _
                   CisFun.Val2(iB12_Su(wNum2)) <> 0 Then

                    iB12_Hin1(wNum1) = RTrim(iB12_Hin1(wNum2))
                    iB12_Hin2(wNum1) = RTrim(iB12_Hin2(wNum2))
                    iB12_Hin3(wNum1) = RTrim(iB12_Hin3(wNum2))
                    iB12_Seban(wNum1) = RTrim(iB12_Seban(wNum2))
                    iB12_Hako(wNum1) = iB12_Hako(wNum2)
                    iB12_Syuyo(wNum1) = iB12_Syuyo(wNum2)
                    iB12_Su(wNum1) = iB12_Su(wNum2)
                    iB12_Biko(wNum1) = iB12_Biko(wNum2)
                    Blb_BikoNM(wNum1) = Blb_BikoNM(wNum2)
                    iB14_Bunsu(wNum1) = iB14_Bunsu(wNum2)

                    iB12_Hin1(wNum2) = "":      iB12_Hin2(wNum2) = ""
                    iB12_Hin3(wNum2) = "":      iB12_Seban(wNum2) = ""
                    iB12_Hako(wNum2) = "":      iB12_Syuyo(wNum2) = ""
                    iB12_Su(wNum2) = "":        iB12_Biko(wNum2) = ""
                    Blb_BikoNM(wNum2) = "":     iB14_Bunsu(wNum2) = ""

                    wlc_Flg = True
                    Exit For
                End If
            Next wNum2
            If Not wlc_Flg Then Exit For
        End If
    Next wNum1
End Sub
'*----------------------------------------------------------------*
'*          ì`ï[ì‡óeÉ`ÉFÉbÉN
'*
'*      âÊñ ï\é¶çÄñ⁄ÇÉ`ÉPÉbÉgì«éÊÉfÅ[É^ÇÃólÇ…ÅAï∂éöóÒÇ…ÇµÇƒ
'*                          ÉXÉgÉAÉhÇ…Çƒì‡óeÇÃÉ`ÉFÉbÉNÇçsÇ§ÅB
'*      É`ÉFÉbÉNå„Ç…ÅAÉXÉgÉAÉhÇ©ÇÁÇÃ Output ﬂ◊“∞¿ÇégópÇµÅA
'*                                              çƒï\é¶ÇÇ∑ÇÈÅB
'*----------------------------------------------------------------*
Private Function ContentsChk() As Boolean
    ContentsChk = False
 
    Static wcce_Obj     As Object
    
    If Trim(iB13_YMD) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_Title = ""
            .MB_Button = Error
            .MB_MSG(2) = "   î[ì¸ì˙Çì¸óÕÇµÇƒâ∫Ç≥Ç¢              "
            .MBOX
        End With
        iB13_YMD.SetFocus
        Exit Function
    End If
    If Not iB13_YMD.cChkResult Then
        With CisFun
            .MB_Lines = 5
            .MB_Title = ""
            .MB_Button = Error
            .MB_MSG(2) = "  î[ì¸ì˙Çê≥ÇµÇ¢ì˙ïtÇ…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢                "
            .MBOX
        End With
        iB13_YMD.SetFocus
        Exit Function
    End If

    If CisFun.Val2(iB13_Bin) = 0 Then
        With CisFun
            .MB_Lines = 5
            .MB_Title = ""
            .MB_Button = Error
            .MB_MSG(2) = "      î[ì¸ï÷Çì¸óÕÇµÇƒâ∫Ç≥Ç¢              "
            .MBOX
        End With
        iB13_Bin.SetFocus
        Exit Function
    End If

    If Trim(iB1_Noubn) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_Title = ""
            .MB_Button = Error
            .MB_MSG(2) = "       î[ì¸î‘çÜÇì¸óÕÇµÇƒâ∫Ç≥Ç¢                "
            .MBOX
        End With
        iB1_Noubn.SetFocus
        Exit Function
    End If

    Static HasuChk      As Integer
'
    HasuChk = 9
    For wNum1 = 0 To 6
'
'        If CisFun.Val2(iB12_Syuyo(wNum1)) <> 0 Then
''               * í[êî•î†êîÉ`ÉFÉbÉN*
'            wNumL = CisFun.Val2(iB12_Su(wNum1)) \ CisFun.Val2(iB12_Syuyo(wNum1))
'            If (CisFun.Val2(iB12_Su(wNum1)) Mod CisFun.Val2(iB12_Syuyo(wNum1))) Then
'                wNumL = wNumL + 1
'            End If
''
'            If CisFun.Val2(iB12_Hako(wNum1)) <> wNumL Then
'                With CisFun
'                    .MB_Lines = 5
'                    .MB_Title = ""
'                    .MB_Button = Error
'                    .MB_MSG(1) = "Åy  " & StrConv(wNum1 + 1, vbWide) & " çsñ⁄  Åz"
'                    .MB_MSG(3) = "      î†êîÉGÉâÅ[          "
'                    .MBOX
'                End With
'                iB12_Hako(wNum1).SetFocus
'                Exit Function
'            End If
'        End If
'
'           * í[êîämîFÉÅÉbÉZÅ[ÉW *
        wNumL = CisFun.Val2(iB12_Syuyo(wNum1)) * CisFun.Val2(iB12_Hako(wNum1))
        If CisFun.Val2(iB12_Su(wNum1)) <> wNumL Then
            If HasuChk = 9 Then HasuChk = wNum1
        End If
'
    Next wNum1
'
    If HasuChk <> 9 Then
        With CisFun
            .MB_Lines = 5
            .MB_Title = ""
            .MB_Button = OK_CAN
            .MB_MSG(1) = "Åyé˚óeêî Å~ î†êî ÅÅ î[ì¸êî"
            .MB_MSG(2) = "          Ç…Ç»Ç¡ÇƒÇ¢Ç»Ç¢çsÇ™óLÇËÇ‹Ç∑   "
            .MB_MSG(4) = "         èàóùÇë±çsÇµÇ‹Ç∑Ç©ÅH"
            If Not .MBOX Then iB12_Hako(HasuChk).SetFocus: Exit Function
        End With
    End If

    wGyo = 0
    For wNum1 = 0 To 6
        If Trim(iB12_Seban(wNum1)) <> "" Or _
           Trim(iB12_Hin1(wNum1)) <> "" Or Trim(iB12_Hin2(wNum1)) <> "" Or _
           Trim(iB12_Hin3(wNum1)) <> "" Then
           wGyo = wGyo + 1
        End If
    Next wNum1
    iB14_Gyo = wGyo
' Ωƒ±ƒﬁà¯ìnÇµópÅ@√ﬁ∞¿çÏê¨
    Call CrtStoado("H")

'
    If RTrim(wcc_Head) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_Title = ""
            .MB_Button = Error
            .MB_MSG(2) = "  ÉwÉbÉhïî(î[ì¸ì˙Å`éxããå≥)Çì¸óÕÇµÇƒâ∫Ç≥Ç¢      "
            .MBOX
        End With
        iB13_YMD.SetFocus
        Exit Function
    End If

' Ωƒ±ƒﬁà¯ìnÇµópÅ@√ﬁ∞¿çÏê¨
    Call CrtStoado("B")

    If wGyo = 0 Then
        With CisFun
            .MB_Lines = 5
            .MB_Title = ""
            .MB_Button = Error
            .MB_MSG(2) = "     ñæç◊çsÇì¸óÕÇµÇƒâ∫Ç≥Ç¢          "
            .MBOX
        End With
        iB12_Seban(0).SetFocus
        Exit Function
    End If
'
    If SyoriKB = "RED" Then
       gInt = 2
    Else
       gInt = 1
    End If
    If ChekData(1, gInt, 1) Then
       DoEvents
       ContentsChk = True
       Exit Function
    Else
       On Error Resume Next
       mFocusObj.SetFocus
       Set mFocusObj = Nothing
       On Error GoTo 0
    End If

End Function
'+----------------------------------------------------------------------+
'+      ÉXÉgÉAÉhà¯ìnÇµópÅ@ÉfÅ[É^çÏê¨
'+----------------------------------------------------------------------+
Private Sub CrtStoado(HB As String)
'   * ÉwÉbÉhïîçÏê¨ *
    If HB = "H" Then
        wcc_Head = ""
        wcc_Head = "A"
        wcc_Head = wcc_Head & Format(CisFun.Val2(iB14_Yobi1), "0")
        wcc_Head = wcc_Head & Format(CisFun.Val2(iB14_Yobi2), "0")
        wcc_Head = wcc_Head & Format(CisFun.Val2(iB14_Yobi3), "0")
        wcc_Head = wcc_Head & Format(CisFun.Val2(iB14_Gyo), "0")
        wcc_Head = wcc_Head & Format(CisFun.Val2(iB14_HKB), "6")
        wcc_Head = wcc_Head & Mid(Format(" " & RTrim(iB1_Noubn), "!" & String(5 + 1, "@")), 2)
        wcc_Head = wcc_Head & Format(CisFun.Val2(iB1_Page), "00")
        wcc_Head = wcc_Head & Format(CisFun.Val2(iB14_Bun), "0")
        wcc_Head = wcc_Head & Format(CisFun.Val2(iB1_Sharp), "0")
        wcc_Head = wcc_Head & Format(CisFun.Val2(iB1_SRTorcd), "0000")
        wcc_Head = wcc_Head & Mid(Format(" " & RTrim(iB1_SRKojyo), "!" & String(1 + 1, "@")), 2)
        wcc_Head = wcc_Head & Format(CisFun.Val2(iB14_Kouku), "00")
        wcc_Head = wcc_Head & Format(CisFun.Val2(iB1_SKTorcd), "0000")
        wcc_Head = wcc_Head & Mid(Format(" " & RTrim(iB1_Uke), "!" & String(2 + 1, "@")), 2)
        wcc_Head = wcc_Head & Mid(Format(" " & RTrim(iB1_Kibn), "!" & String(2 + 1, "@")), 2)
        wcc_Head = wcc_Head & Format(CisFun.Val2(iB1_Moto), "0000")
        wcc_Head = wcc_Head & Format(CisFun.Val2(iB14_MotoYobi), "0000")
        wcc_Head = wcc_Head & Mid(Format(" " & Mid(iB13_YMD, 3, 6), "!" & String(6 + 1, "@")), 2)
        wcc_Head = wcc_Head & Mid(Format(" " & RTrim(iB13_Bin), "!" & String(2 + 1, "@")), 2)
        wcc_Head = wcc_Head & Mid(Format(" " & Mid(iB14_HYmd, 3, 6), "!" & String(6 + 1, "@")), 2)
        wcc_Head = wcc_Head & Mid(Format(" " & RTrim(iB14_HBin), "!" & String(2 + 1, "@")), 2)
        wcc_Head = wcc_Head & Format(CisFun.Val2(iB1_MPage), "00")
    End If
    If HB = "B" Then
'   * É{ÉfÉBïîçÏê¨ *
        wcc_Body = ""
        For wNum1 = 0 To 6
    '       * îwî‘çÜ*
            If RTrim(iB12_Seban(wNum1)) = "" Then
                wcc_Body = wcc_Body & Space(3)
            Else
                wcc_Body = wcc_Body & Format(RTrim(iB12_Seban(wNum1)), "@@@")
            End If
            
    '       * ïiî‘ *
            If RTrim(iB12_Hin1(wNum1) & iB12_Hin2(wNum1) & iB12_Hin3(wNum1)) = "" Then
                wcc_Body = wcc_Body & Space(12)
            Else
    '           * ïiî‘ÇP *
                If RTrim(iB12_Hin1(wNum1)) = "" Then
                    wcc_Body = wcc_Body & Space(5)
                Else
                    wcc_Body = wcc_Body & Format(RTrim(iB12_Hin1(wNum1)), "!@@@@@")
                End If
                
    '           * ïiî‘ÇQ *
                If RTrim(iB12_Hin2(wNum1)) = "" Then
                    wcc_Body = wcc_Body & Space(5)
                Else
                    wcc_Body = wcc_Body & Format(RTrim(iB12_Hin2(wNum1)), "!@@@@@")
                End If
                
    '           * ïiî‘ÇR *
                If RTrim(iB12_Hin3(wNum1)) = "" Then
                    wcc_Body = wcc_Body & Space(2)
                Else
                    wcc_Body = wcc_Body & Format(RTrim(iB12_Hin3(wNum1)), "!@@")
                End If
            End If
    '       * é˚óeêî *
            wcc_Body = wcc_Body & Format(CisFun.Val2(iB12_Syuyo(wNum1)), "00000")
    '       * î†êî *
            wcc_Body = wcc_Body & Format(CisFun.Val2(iB12_Hako(wNum1)), "000")
    '       * î[ì¸êî *
            wcc_Body = wcc_Body & Format(CisFun.Val2(iB12_Su(wNum1)), "000000")
    '       * ë≈êÿÇËãÊï™ *
            wcc_Body = wcc_Body & Format(CisFun.Val2(iB12_Biko(wNum1)), "0")
    '       * ë≈êÿÇËécêî *
            wcc_Body = wcc_Body & Format(CisFun.Val2(iB14_Bunsu(wNum1)), "000000")
        Next wNum1
    End If
End Sub
'**************************************
'*                                    *
'*          ÉfÅ[É^ÉxÅ[ÉXçXêV           *
'*                                    *
'**************************************
'+----------------------+
'+    í«â¡ èàóù          +
'+----------------------+
Private Function DBPut()
    DBPut = False
'  ( ämîFÉÅÉbÉZÅ[ÉW )
    If Not DBCheckMsg Then Exit Function
 
    On Error GoTo DBPut_Err

' Ωƒ±ƒﬁà¯ìnÇµópÅ@ÉfÅ[É^çÏê¨
    Call CrtStoado("H")
    Call CrtStoado("B")

    If SyoriKB = "RED" Then
       gInt = 2
    Else
       gInt = 1
    End If
    If Not ChekData(0, gInt, 1) Then
        
    End If
    
    If SyoriKB = "RED" Then
       H1lb_OCRCnt = Format(CisFun.Val2(H1lb_OCRCnt) + 1, "#,###")
    Else
       iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "#,###")
    End If
    
    On Error GoTo 0
    
    DBPut = True
    Exit Function
DBPut_Err:
    Call CisFun.ErrorBox("çXêVÉGÉâÅ[")
    End
End Function
'+------------------------------------+
'+    èàóùämîFÉÅÉbÉZÅ[ÉWÅiDBUpDate)
'+------------------------------------+
Private Function DBCheckMsg()
    DBCheckMsg = False
    
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
         .MB_Lines = 3
         .MB_Title = SyoriNM & "èàóù"
         .MB_Button = OK_CAN
         .MB_MSG(1) = "    Åy " & SyoriNM & "èàóùämîF Åz      "
         
         If Not .MBOX Then Exit Function
    End With
    
    DBCheckMsg = True
End Function
