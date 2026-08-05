VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CNJ0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "ì‡é¶ÉfÅ[É^ÉÅÉìÉeÉiÉìÉX"
   ClientHeight    =   11190
   ClientLeft      =   1365
   ClientTop       =   1035
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
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   Begin Cis3D_v60.CIS3D H1lb_Ku 
      Height          =   375
      Left            =   7110
      Top             =   810
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColor       =   16777152
      Caption         =   "ä™ï‘Çµ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cAlingnment     =   7
      cPositionX      =   30
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   945
      Left            =   930
      ScaleHeight     =   885
      ScaleWidth      =   11205
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   2610
      Width           =   11265
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   795
         Left            =   60
         Top             =   60
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "ëw  ï "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin VB.ComboBox B1Comb_Sobetsu 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   330
            Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   330
            Width           =   1530
         End
         Begin CisText_V60.CisText iB1_Sobetsu 
            Height          =   375
            Left            =   60
            TabIndex        =   5
            Top             =   330
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
      End
      Begin Cis3D_v60.CIS3D NaijiT 
         Height          =   795
         Left            =   1980
         Top             =   60
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "ìñåéì‡é¶êî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_TNSu 
            Height          =   375
            Left            =   60
            TabIndex        =   7
            Top             =   330
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   661
            cFormat         =   "##,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "##,###,###"
            MaxLength       =   10
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "########"
            cILength        =   8
         End
      End
      Begin Cis3D_v60.CIS3D NaijiY 
         Height          =   795
         Left            =   3510
         Top             =   60
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "óÇåéì‡é¶êî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_TYSu 
            Height          =   375
            Left            =   60
            TabIndex        =   8
            Top             =   330
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   661
            cFormat         =   "##,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "##,###,###"
            MaxLength       =   10
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "########"
            cILength        =   8
         End
      End
      Begin Cis3D_v60.CIS3D NaijiYY 
         Height          =   795
         Left            =   5100
         Top             =   60
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "óÇÅXåéì‡é¶êî"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iB1_TYYSu 
            Height          =   375
            Left            =   60
            TabIndex        =   9
            Top             =   330
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   661
            cFormat         =   "##,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "##,###,###"
            MaxLength       =   10
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "########"
            cILength        =   8
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   795
         Left            =   6690
         Top             =   60
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "ìK  óp  ì˙  ït"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   195
            Index           =   19
            Left            =   2010
            Top             =   450
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   344
            ForeColor       =   16711680
            Caption         =   "Å`"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin CisYMD_With_Btn.CisYMDwB iB1_SDate 
            Height          =   360
            Left            =   90
            TabIndex        =   10
            Top             =   360
            Width           =   1845
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
            Object.Width           =   1845
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
         Begin CisYMD_With_Btn.CisYMDwB iB1_EDate 
            Height          =   360
            Left            =   2490
            TabIndex        =   11
            Top             =   360
            Width           =   1845
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
            Object.Width           =   1845
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
      Height          =   945
      Left            =   930
      ScaleHeight     =   885
      ScaleWidth      =   6975
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   1200
      Width           =   7035
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   795
         Left            =   1560
         Top             =   60
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "ïi  î‘"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   30
            TabIndex        =   1
            Top             =   330
            Width           =   3240
            _ExtentX        =   5715
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   795
         Left            =   60
         Top             =   60
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "ëŒè€îNåé"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   375
            Left            =   30
            TabIndex        =   0
            Top             =   330
            Width           =   1410
            _ExtentX        =   2328
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
            cSize           =   -1  'True
            cChkResult      =   0   'False
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   795
         Left            =   4860
         Top             =   60
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "î[ì¸êÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   60
            TabIndex        =   2
            Top             =   330
            Width           =   1005
            _ExtentX        =   1773
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
            Text            =   "XXXXXXX"
            MaxLength       =   7
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   795
         Left            =   6000
         Top             =   60
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "éÛ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   375
            Left            =   60
            TabIndex        =   3
            Top             =   330
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
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   795
         Left            =   6510
         Top             =   60
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "ãÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         Begin CisText_V60.CisText iH1_Ku 
            Height          =   375
            Left            =   90
            TabIndex        =   4
            Top             =   330
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
      Height          =   1335
      Left            =   10215
      ScaleHeight     =   1275
      ScaleWidth      =   4215
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   840
      Width           =   4275
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1695
         Top             =   630
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
         Left            =   2520
         Top             =   630
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
         Left            =   3345
         Top             =   630
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1695
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
         BtnType         =   4
         BtnCaption      =   "çÌèú"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   255
         BtnShadowColor  =   16777215
         PFCaption       =   "( F3 )"
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
         PFBackColor     =   255
      End
      Begin CisBtn_60.CisBtn PB_ADD 
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
         BtnCaption      =   "í«â¡"
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
      Begin CisBtn_60.CisBtn PB_MNT 
         Height          =   585
         Left            =   870
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
         BtnType         =   8
         BtnCaption      =   "èCê≥"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ñæí©"
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   2520
         Top             =   30
         Visible         =   0   'False
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
         Left            =   870
         Top             =   630
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
         Left            =   45
         Top             =   630
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
      Left            =   960
      TabIndex        =   67
      Text            =   "Text1"
      Top             =   1170
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   1140
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Left            =   1170
      Top             =   1110
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
      Caption         =   "Åy ì‡é¶ÉfÅ[É^ÉÅÉìÉeÉiÉìÉX  Åz"
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
         Left            =   10110
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   30
         Width           =   3675
         Begin Cis3D_v60.CIS3D CIS3D15 
            Height          =   345
            Left            =   2460
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   8421631
            ForeColor       =   0
            Caption         =   "çÌèú"
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
            Begin Cis3D_v60.CIS3D iT1_Delsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   8421631
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
         Begin Cis3D_v60.CIS3D UC_3D6 
            Height          =   345
            Left            =   0
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   8454143
            ForeColor       =   0
            Caption         =   "í«â¡"
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
            Begin Cis3D_v60.CIS3D iT1_Addsu 
               Height          =   255
               Left            =   540
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
         Begin Cis3D_v60.CIS3D UC_3D7 
            Height          =   345
            Left            =   1230
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   16777088
            ForeColor       =   0
            Caption         =   "èCê≥"
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
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   16777088
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
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   12
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
   End
   Begin VB.PictureBox B1_Area2 
      BackColor       =   &H00C0C0C0&
      Height          =   6600
      Left            =   915
      ScaleHeight     =   6540
      ScaleWidth      =   13500
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   3840
      Width           =   13560
      Begin VB.PictureBox B1_Area3 
         BackColor       =   &H00C0C0C0&
         Height          =   6465
         Left            =   45
         ScaleHeight     =   6405
         ScaleWidth      =   13365
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   30
         Width           =   13425
         Begin VB.PictureBox Picture2 
            Height          =   6315
            Left            =   9420
            ScaleHeight     =   6255
            ScaleWidth      =   3825
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   45
            Width           =   3885
            Begin Cis3D_v60.CIS3D B1lb_Mnt 
               Height          =   1005
               Left            =   2580
               Top             =   5250
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   1773
               ForeColor       =   16512
               Caption         =   "ïœçX"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   1
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D B1lb_DtataNm 
               Height          =   1005
               Left            =   0
               Top             =   5250
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   1773
               Caption         =   "ÉIÉìÉâÉCÉìÉfÅ[É^(åvïœ)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cPositionY      =   -100
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   855
               Index           =   0
               Left            =   2115
               Top             =   0
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   1508
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin CisText_V60.CisText iB1_Seban 
                  Height          =   375
                  Left            =   255
                  TabIndex        =   14
                  Top             =   360
                  Width           =   1110
                  _ExtentX        =   1958
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
                  Text            =   "XXXXXXXX"
                  MaxLength       =   8
                  IMEMode         =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D10 
               Height          =   855
               Left            =   1410
               Top             =   870
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   1508
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               Begin CisText_V60.CisText iB1_Syuyo 
                  Height          =   375
                  Left            =   30
                  TabIndex        =   18
                  Top             =   360
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   661
                  cFormat         =   "##,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##,###"
                  MaxLength       =   6
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#####"
                  cILength        =   5
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   855
               Index           =   1
               Left            =   0
               Top             =   870
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   1508
               ForeColor       =   16711680
               Caption         =   "ÉTÉCÉNÉã"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_CyCle1 
                  Height          =   375
                  Left            =   45
                  TabIndex        =   15
                  Top             =   360
                  Width           =   1335
                  _ExtentX        =   2355
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
                  Text            =   "XXXXXXXXXX"
                  MaxLength       =   10
                  cFaZero         =   0
                  cFbComma        =   0
               End
               Begin CisText_V60.CisText iB1_CyCle2 
                  Height          =   375
                  Left            =   375
                  TabIndex        =   16
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   405
                  _ExtentX        =   714
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
                  Text            =   "##"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin CisText_V60.CisText iB1_CyCle3 
                  Height          =   375
                  Left            =   795
                  TabIndex        =   17
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   405
                  _ExtentX        =   714
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
                  Text            =   "##"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   855
               Index           =   2
               Left            =   3315
               Top             =   870
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   1508
               ForeColor       =   16711680
               Caption         =   "êÆîı"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Seibi 
                  Height          =   375
                  Left            =   45
                  TabIndex        =   20
                  Top             =   360
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
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   855
               Index           =   3
               Left            =   0
               Top             =   1740
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   1508
               ForeColor       =   16711680
               Caption         =   "íSìñ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Tanto 
                  Height          =   375
                  Left            =   60
                  TabIndex        =   21
                  Top             =   360
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
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   855
               Index           =   4
               Left            =   2340
               Top             =   870
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   1508
               ForeColor       =   16711680
               Caption         =   "ä«óùî‘çÜ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Kanri 
                  Height          =   375
                  Left            =   45
                  TabIndex        =   19
                  Top             =   360
                  Width           =   885
                  _ExtentX        =   1561
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
                  Text            =   "XXXXXX"
                  MaxLength       =   6
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   735
               Index           =   5
               Left            =   4110
               Top             =   0
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   1296
               ForeColor       =   16711680
               Caption         =   "íS"
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
               cAlingnment     =   6
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   855
               Index           =   6
               Left            =   525
               Top             =   1740
               Width           =   3285
               _ExtentX        =   5794
               _ExtentY        =   1508
               ForeColor       =   16711680
               Caption         =   "î[Å@ãÊ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin Cis3D_v60.CIS3D B1lb_Nouku 
                  Height          =   375
                  Left            =   315
                  Top             =   360
                  Width           =   2895
                  _ExtentX        =   5106
                  _ExtentY        =   661
                  BackColor       =   16777152
                  Caption         =   "Daily éwé¶ïîïi"
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
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   1
                  cBoderStyle     =   1
                  cPositionX      =   30
               End
               Begin CisText_V60.CisText iB1_Nouku 
                  Height          =   375
                  Left            =   45
                  TabIndex        =   22
                  Top             =   360
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
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   855
               Index           =   7
               Left            =   0
               Top             =   3480
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   1508
               ForeColor       =   16711680
               Caption         =   "î[ïièëî≠çsñáêî"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cPositionY      =   50
               Begin CisText_V60.CisText iB1_NouSu 
                  Height          =   375
                  Left            =   420
                  TabIndex        =   24
                  Top             =   360
                  Width           =   525
                  _ExtentX        =   926
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
                  Text            =   "###"
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   855
               Index           =   8
               Left            =   1620
               Top             =   3480
               Width           =   2190
               _ExtentX        =   3863
               _ExtentY        =   1508
               ForeColor       =   16711680
               Caption         =   "Ç©ÇÒÇŒÇÒñáêî/ëOâÒç∑"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cPositionY      =   50
               Begin CisText_V60.CisText iB1_KMai 
                  Height          =   375
                  Left            =   270
                  TabIndex        =   25
                  Top             =   360
                  Width           =   525
                  _ExtentX        =   926
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
                  Text            =   "###"
                  MaxLength       =   3
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
               Begin CisText_V60.CisText iB1_KMaiZ 
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   26
                  Top             =   360
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
                  Text            =   "####"
                  MaxLength       =   4
                  cDataType       =   4
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "####"
                  cILength        =   4
               End
               Begin Cis3D_v60.CIS3D BLB_DD 
                  Height          =   195
                  Index           =   45
                  Left            =   840
                  Top             =   480
                  Width           =   420
                  _ExtentX        =   741
                  _ExtentY        =   344
                  ForeColor       =   16711680
                  Caption         =   "Å^"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cBoderColor1    =   12632256
                  cBoderColor2    =   12632256
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   2
                  cAlingnment     =   7
                  cBoderStyle     =   2
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   855
               Index           =   9
               Left            =   0
               Top             =   4350
               Width           =   1860
               _ExtentX        =   3281
               _ExtentY        =   1508
               ForeColor       =   16711680
               Caption         =   "åvâÊïœçXêî"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Hsu 
                  Height          =   375
                  Left            =   210
                  TabIndex        =   27
                  Top             =   360
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   661
                  cFormat         =   "###,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###,###"
                  MaxLength       =   11
                  cDataType       =   4
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#########"
                  cILength        =   9
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   855
               Index           =   10
               Left            =   1860
               Top             =   4350
               Width           =   1950
               _ExtentX        =   3440
               _ExtentY        =   1508
               ForeColor       =   16711680
               Caption         =   "é¿ê—í≤êÆêî"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_CSu 
                  Height          =   375
                  Left            =   240
                  TabIndex        =   28
                  Top             =   360
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   661
                  cFormat         =   "###,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###,###"
                  MaxLength       =   11
                  cDataType       =   4
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#########"
                  cILength        =   9
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   855
               Index           =   11
               Left            =   0
               Top             =   0
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   1508
               ForeColor       =   16711680
               Caption         =   "çH"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Kouku 
                  Height          =   375
                  Left            =   195
                  TabIndex        =   12
                  Top             =   360
                  Width           =   285
                  _ExtentX        =   503
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
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   855
               Index           =   12
               Left            =   675
               Top             =   0
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   1508
               ForeColor       =   16711680
               Caption         =   "å_ñÒêÊ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_KTorcd 
                  Height          =   375
                  Left            =   210
                  TabIndex        =   13
                  Top             =   360
                  Width           =   990
                  _ExtentX        =   1746
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
                  Text            =   "XXXXXXX"
                  MaxLength       =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D9 
               Height          =   855
               Index           =   13
               Left            =   0
               Top             =   2610
               Width           =   3810
               _ExtentX        =   6720
               _ExtentY        =   1508
               ForeColor       =   16711680
               Caption         =   "ÉRÅ@ÉÅÅ@ÉìÅ@Ég"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               Begin CisText_V60.CisText iB1_Comment 
                  Height          =   375
                  Left            =   150
                  TabIndex        =   23
                  Top             =   360
                  Width           =   3615
                  _ExtentX        =   6376
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   30
                  IMEMode         =   4
               End
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D4 
            Height          =   6315
            Left            =   45
            Top             =   45
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   11139
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderStyle     =   1
            Begin Cis3D_v60.CIS3D BLB_Sur 
               Height          =   6225
               Index           =   1
               Left            =   45
               Top             =   45
               Width           =   9285
               _ExtentX        =   16378
               _ExtentY        =   10980
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               Begin Cis3D_v60.CIS3D PB_CLR 
                  Height          =   405
                  Left            =   3330
                  Top             =   5790
                  Width           =   2745
                  _ExtentX        =   4842
                  _ExtentY        =   714
                  BackColor       =   12648384
                  Caption         =   "ì˙ó êîÇëSÇƒè¡ãé(F8)"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   3
                  cAlingnment     =   7
                  cButton         =   -1  'True
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   36
                  Left            =   1350
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   110
                  TabStop         =   0   'False
                  Top             =   5280
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   36
                     Left            =   -15
                     TabIndex        =   65
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   36
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   36
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   35
                  Left            =   30
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   109
                  TabStop         =   0   'False
                  Top             =   5280
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   35
                     Left            =   -15
                     TabIndex        =   64
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   35
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   35
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   34
                  Left            =   7950
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   108
                  TabStop         =   0   'False
                  Top             =   4320
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   34
                     Left            =   -15
                     TabIndex        =   63
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   34
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   34
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   33
                  Left            =   6630
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   107
                  TabStop         =   0   'False
                  Top             =   4320
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   33
                     Left            =   -15
                     TabIndex        =   62
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   33
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   33
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   32
                  Left            =   5310
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   106
                  TabStop         =   0   'False
                  Top             =   4320
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   32
                     Left            =   -15
                     TabIndex        =   61
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   32
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   32
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   31
                  Left            =   3990
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   105
                  TabStop         =   0   'False
                  Top             =   4320
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   31
                     Left            =   -15
                     TabIndex        =   60
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   31
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   31
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   30
                  Left            =   2670
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   104
                  TabStop         =   0   'False
                  Top             =   4320
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   30
                     Left            =   -15
                     TabIndex        =   59
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   30
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   30
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   29
                  Left            =   1350
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   103
                  TabStop         =   0   'False
                  Top             =   4320
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   29
                     Left            =   -15
                     TabIndex        =   58
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   29
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   29
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   28
                  Left            =   30
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   102
                  TabStop         =   0   'False
                  Top             =   4320
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   28
                     Left            =   -15
                     TabIndex        =   57
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   28
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   28
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   27
                  Left            =   7950
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   101
                  TabStop         =   0   'False
                  Top             =   3330
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   27
                     Left            =   -15
                     TabIndex        =   56
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   27
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   27
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   26
                  Left            =   6630
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   100
                  TabStop         =   0   'False
                  Top             =   3330
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   26
                     Left            =   -15
                     TabIndex        =   55
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   26
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   26
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   25
                  Left            =   5310
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   99
                  TabStop         =   0   'False
                  Top             =   3330
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   25
                     Left            =   -15
                     TabIndex        =   54
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   25
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   25
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   24
                  Left            =   3990
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   98
                  TabStop         =   0   'False
                  Top             =   3330
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   24
                     Left            =   -15
                     TabIndex        =   53
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   24
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   24
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   23
                  Left            =   2670
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   97
                  TabStop         =   0   'False
                  Top             =   3330
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   23
                     Left            =   -15
                     TabIndex        =   52
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   23
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   23
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   22
                  Left            =   1350
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   96
                  TabStop         =   0   'False
                  Top             =   3330
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   22
                     Left            =   -15
                     TabIndex        =   51
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   22
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   22
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   21
                  Left            =   30
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   95
                  TabStop         =   0   'False
                  Top             =   3330
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   21
                     Left            =   -15
                     TabIndex        =   50
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   21
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   21
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   20
                  Left            =   7950
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   94
                  TabStop         =   0   'False
                  Top             =   2370
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   20
                     Left            =   -15
                     TabIndex        =   49
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   20
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   20
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   19
                  Left            =   6630
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   93
                  TabStop         =   0   'False
                  Top             =   2370
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   19
                     Left            =   -15
                     TabIndex        =   48
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   19
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   19
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   18
                  Left            =   5310
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   92
                  TabStop         =   0   'False
                  Top             =   2370
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   18
                     Left            =   -15
                     TabIndex        =   47
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   18
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   18
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   17
                  Left            =   3990
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   91
                  TabStop         =   0   'False
                  Top             =   2370
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   17
                     Left            =   -15
                     TabIndex        =   46
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   17
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   17
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   16
                  Left            =   2670
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   90
                  TabStop         =   0   'False
                  Top             =   2370
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   16
                     Left            =   -15
                     TabIndex        =   45
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   16
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   16
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   15
                  Left            =   1350
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   89
                  TabStop         =   0   'False
                  Top             =   2370
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   15
                     Left            =   -15
                     TabIndex        =   44
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   15
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   15
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   14
                  Left            =   30
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   88
                  TabStop         =   0   'False
                  Top             =   2370
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   14
                     Left            =   -15
                     TabIndex        =   43
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   14
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   14
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   13
                  Left            =   7950
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   87
                  TabStop         =   0   'False
                  Top             =   1410
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   13
                     Left            =   -15
                     TabIndex        =   42
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   13
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   13
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   12
                  Left            =   6630
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   86
                  TabStop         =   0   'False
                  Top             =   1410
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   12
                     Left            =   -15
                     TabIndex        =   41
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   12
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   12
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   11
                  Left            =   5310
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   85
                  TabStop         =   0   'False
                  Top             =   1410
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   11
                     Left            =   -15
                     TabIndex        =   40
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   11
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   11
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   10
                  Left            =   3990
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   84
                  TabStop         =   0   'False
                  Top             =   1410
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   10
                     Left            =   -15
                     TabIndex        =   39
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   10
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   10
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   9
                  Left            =   2670
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   83
                  TabStop         =   0   'False
                  Top             =   1410
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   9
                     Left            =   -15
                     TabIndex        =   38
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   9
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   9
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   8
                  Left            =   1350
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   82
                  TabStop         =   0   'False
                  Top             =   1410
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   8
                     Left            =   -15
                     TabIndex        =   37
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   8
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   8
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   7
                  Left            =   30
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   81
                  TabStop         =   0   'False
                  Top             =   1410
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   7
                     Left            =   -15
                     TabIndex        =   36
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   7
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   7
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   6
                  Left            =   7950
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   80
                  TabStop         =   0   'False
                  Top             =   450
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   6
                     Left            =   -15
                     TabIndex        =   35
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   6
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   6
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   5
                  Left            =   6630
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   79
                  TabStop         =   0   'False
                  Top             =   450
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   5
                     Left            =   -15
                     TabIndex        =   34
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   5
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   5
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   4
                  Left            =   5310
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   78
                  TabStop         =   0   'False
                  Top             =   450
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   4
                     Left            =   -15
                     TabIndex        =   33
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   4
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   4
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   3
                  Left            =   3990
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   77
                  TabStop         =   0   'False
                  Top             =   450
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   3
                     Left            =   -15
                     TabIndex        =   32
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   3
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   3
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   2
                  Left            =   2670
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   76
                  TabStop         =   0   'False
                  Top             =   450
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   2
                     Left            =   -15
                     TabIndex        =   31
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   2
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   2
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   1
                  Left            =   1350
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   75
                  TabStop         =   0   'False
                  Top             =   450
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   1
                     Left            =   -15
                     TabIndex        =   30
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   1
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   1
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16777152
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin Cis3D_v60.CIS3D CIS3D18 
                  Height          =   405
                  Left            =   6210
                  Top             =   5310
                  Width           =   1665
                  _ExtentX        =   2937
                  _ExtentY        =   714
                  BackColor       =   14737632
                  Caption         =   "è„íiÅF ç°âÒêî"
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
                  cAlingnment     =   1
                  cBoderStyle     =   1
                  cPositionX      =   60
               End
               Begin VB.PictureBox B_Sur 
                  BackColor       =   &H00808080&
                  BorderStyle     =   0  'Ç»Çµ
                  Height          =   960
                  Index           =   0
                  Left            =   30
                  ScaleHeight     =   960
                  ScaleWidth      =   1305
                  TabIndex        =   71
                  TabStop         =   0   'False
                  Top             =   450
                  Width           =   1305
                  Begin CisText_V60.CisText iB1_Sur 
                     Height          =   345
                     Index           =   0
                     Left            =   -15
                     TabIndex        =   29
                     Top             =   300
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   609
                     cFormat         =   "##,###,###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "##,###,###"
                     MaxLength       =   10
                     cDataType       =   4
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "########"
                     cILength        =   8
                  End
                  Begin Cis3D_v60.CIS3D B1lb_DD 
                     Height          =   330
                     Index           =   0
                     Left            =   0
                     Top             =   -15
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     ForeColor       =   0
                     Caption         =   "ÇPÇU"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ÇlÇr ÉSÉVÉbÉN"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   -2147483643
                     cFont3DColor2   =   8421504
                     cBoderWidth     =   2
                     cAlingnment     =   7
                  End
                  Begin Cis3D_v60.CIS3D B1lb_Kaku 
                     Height          =   315
                     Index           =   0
                     Left            =   0
                     Top             =   630
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     BackColor       =   16761087
                     Caption         =   "##,###,###"
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
                     cPositionX      =   -60
                  End
               End
               Begin Cis3D_v60.CIS3D BLB_Week 
                  Height          =   435
                  Index           =   0
                  Left            =   30
                  Top             =   0
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   767
                  ForeColor       =   16711680
                  Caption         =   "ì˙"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   2
                  cAlingnment     =   7
               End
               Begin Cis3D_v60.CIS3D CIS3D11 
                  Height          =   855
                  Left            =   7920
                  Top             =   5310
                  Width           =   1320
                  _ExtentX        =   2328
                  _ExtentY        =   1508
                  ForeColor       =   16711680
                  Caption         =   "â“ìÆì˙êî"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
                  Begin Cis3D_v60.CIS3D B1lb_Kado 
                     Height          =   345
                     Left            =   285
                     Top             =   390
                     Width           =   720
                     _ExtentX        =   1270
                     _ExtentY        =   609
                     BackColor       =   16777152
                     Caption         =   "##.#"
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
                     cAlingnment     =   4
                     cBoderStyle     =   1
                     cPositionX      =   -30
                  End
               End
               Begin Cis3D_v60.CIS3D BLB_Week 
                  Height          =   435
                  Index           =   1
                  Left            =   1350
                  Top             =   0
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   767
                  ForeColor       =   16711680
                  Caption         =   "åé"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   2
                  cAlingnment     =   7
               End
               Begin Cis3D_v60.CIS3D BLB_Week 
                  Height          =   435
                  Index           =   2
                  Left            =   2670
                  Top             =   0
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   767
                  ForeColor       =   16711680
                  Caption         =   "âŒ"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   2
                  cAlingnment     =   7
               End
               Begin Cis3D_v60.CIS3D BLB_Week 
                  Height          =   435
                  Index           =   3
                  Left            =   3990
                  Top             =   0
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   767
                  ForeColor       =   16711680
                  Caption         =   "êÖ"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   2
                  cAlingnment     =   7
               End
               Begin Cis3D_v60.CIS3D BLB_Week 
                  Height          =   435
                  Index           =   4
                  Left            =   5310
                  Top             =   0
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   767
                  ForeColor       =   16711680
                  Caption         =   "ñÿ"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   2
                  cAlingnment     =   7
               End
               Begin Cis3D_v60.CIS3D BLB_Week 
                  Height          =   435
                  Index           =   5
                  Left            =   6630
                  Top             =   0
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   767
                  ForeColor       =   16711680
                  Caption         =   "ã‡"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   2
                  cAlingnment     =   7
               End
               Begin Cis3D_v60.CIS3D BLB_Week 
                  Height          =   435
                  Index           =   6
                  Left            =   7950
                  Top             =   0
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   767
                  ForeColor       =   16711680
                  Caption         =   "ìy"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   2
                  cAlingnment     =   7
               End
               Begin Cis3D_v60.CIS3D CIS3D19 
                  Height          =   405
                  Left            =   6210
                  Top             =   5730
                  Width           =   1665
                  _ExtentX        =   2937
                  _ExtentY        =   714
                  BackColor       =   14737632
                  Caption         =   "â∫íiÅF ëOâÒêî"
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
                  cAlingnment     =   1
                  cBoderStyle     =   1
                  cPositionX      =   60
               End
               Begin Cis3D_v60.CIS3D PB_ReDisp 
                  Height          =   405
                  Left            =   3330
                  Top             =   5340
                  Width           =   2745
                  _ExtentX        =   4842
                  _ExtentY        =   714
                  BackColor       =   16761087
                  Caption         =   "ïœçXëOÇ…ñﬂÇ∑(F5)"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ÇlÇr ÉSÉVÉbÉN"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   3
                  cAlingnment     =   7
                  cButton         =   -1  'True
               End
            End
         End
      End
   End
End
Attribute VB_Name = "CNJ0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   ì‡é¶ÉfÅ[É^ÉÅÉìÉeÉiÉìÉX
'**       ÉtÉHÅ[ÉÄID    :   CNJ0010
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2003/12/15  By CIS SIMPR-A
'**       ïœ  çX  ì˙    :   2003/05/28  By CIS îwî‘çÜîºäpÉ`ÉFÉbÉNí«â¡
'**       ïœ  çX  ì˙    :   2003/06/01  By CIS î[ì¸êÊÅEéÛì¸Å@åüçıèCê≥
'**       ïœ  çX  ì˙    :   2006/03/09  By CIS ª≤∏ŸåÖêîïœçX
'**       ïœ  çX  ì˙    :   2008/04/11  By CIS îwî‘çÜ8åÖ
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB          As String * 2   ' ì¸óÕàÊêÿë÷
    Dim SyoriKB         As String       ' èàóùãÊï™(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriNM         As String       ' èàóùñºèÃ(êVãK•èCê≥•çÌèú ìô)

    Dim CisComb1        As CisCombo

    Dim mWeekStart      As Integer
    Dim mSaturdayColor  As Long
    Dim mSundayColor    As Long
    
    Dim mNInitColor     As Long
    Dim mNHenColor      As Long
    
    Dim mNSu            As Long         ' ì˙ó êîçáåv
    Dim mNaiji          As String
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
    Timer2.Enabled = False   ' èàóùãÊï™ì_ñ≈ã÷é~
    
    ' ñºèÃÉ}ÉXÉ^ämîF
    Call MeisyoInit
'    '+---------------------+
'    '+ ∫›ŒﬁŒﬁØ∏Ω èâä˙ê›íË  +
'    '+---------------------+
    Set CisComb1 = New CisCombo
    With CisComb1
        Set .Connect = CisDB
        Set .CombObj = B1Comb_Sobetsu
        .Control = "ñºèÃÉ}ÉXÉ^"
        .ComboName = "ëwï ãÊï™"
        .Ess = False
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With

'   #------------------#
'   #    èâä˙ì¸óÕê›íË   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  èàóùãÊï™èâä˙ï\é¶ #
'   #------------------#
    SyoriKB = "ADD"
    Call SyoriName(SyoriKB)
'   #-------------------#
'   # âÊñ íÜâõà íuï\é¶   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    ÅsÇ`ÇÉÇîÇâÇñÇÅÇîÇÖ ÉCÉxÉìÉgÅt                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    FormAct = Not FormAct
    
    mSaturdayColor = &HFFC0C0
    mSundayColor = &HC0C0FF
    
    mNHenColor = B1lb_Kaku(0).BackColor
    mNInitColor = B1lb_Kaku(1).BackColor
    B1lb_Kaku(0).BackColor = B1lb_Kaku(1).BackColor
    ' ëŒè€îNåéèâä˙ï\é¶ílÇãÅÇﬂÇÈ
    gSL_Select = "Select âpêîéö1,âpêîéö2 From ñºèÃÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " Where ãÊï™ñºèÃ = 'ì‡é¶îNåé'"
    gSL_Select = gSL_Select & " And   ãÊï™É^ÉCÉv = ''"
    gSL_Select = gSL_Select & " And   íl = '1'"
    If SYMRead(gSL_Select, 1) Then
        iH1_YM = SYM.âpêîéö1
        mNaiji = Mid(SYM.âpêîéö2, 1, 2)
    Else
        mNaiji = "íËèÌ"
    End If
    ' ójì˙äJénÇãÅÇﬂÇÈ
    gSL_Select = "select êîéö1 From ñºèÃÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " Where ãÊï™ñºèÃ = 'ójì˙äJénä«óù'"
    gSL_Select = gSL_Select & "   And íl = '1'"
    If Not SYMRead(gSL_Select, 1) Then
       SYM.êîéö1 = 0
    End If
    
    mWeekStart = SYM.êîéö1 + 1

    ' ójì˙ÉZÉbÉg
    For gInt = 0 To 6
        BLB_Week(gInt) = WeekdayName(gInt + 1, True, mWeekStart)
        If BLB_Week(gInt) = "ì˙" Then
           BLB_Week(gInt).BackColor = mSundayColor
        End If
        If BLB_Week(gInt) = "ìy" Then
           BLB_Week(gInt).BackColor = mSaturdayColor
        End If
    Next gInt
'
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
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    Select Case Key_Code
           Case vbKeyF1:
                If PB_ADD.Visible Then
                   Call PB_ADD_Click          'ÅyêVãKÅz
                Else
                   If ProcHB = "B1" Then
                   For Each gObj In B_Sur
                       If gObj.Visible Then
                          iB1_Sur(gObj.INDEX).SetFocus
                          Exit For
                       End If
                   Next gObj
                   End If
                End If
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    'ÅyèCê≥Åz
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    'ÅyçÌèúÅz
           Case vbKeyF5:       If PB_ReDisp.Visible Then Call PB_ReDisp_Click  'ÅyïœçXëOÅz
           Case vbKeyF8:       If PB_CLR.Visible Then Call PB_CLR_Click  'Åyè¡ãéÅz
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  'ÅyåüçıÅz
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  'Åyå„ëﬁÅz
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  'ÅyëOêiÅz
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    'ÅyéÊè¡Åz
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    'ÅyèIóπÅz
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
'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås QueryUnloadÉCÉxÉìÉg Åt                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Cancel = 1
    Call PB_END_Click
End Sub
'+---------------------+
'+    èI  óπ  ÉL Å[     +
'+---------------------+
Private Sub PB_END_Click()
'+---------------------+
'+   √ﬁ∞¿Õﬁ∞Ωêÿíf
'+---------------------+
    Call CisDB.DBDISConnect
'+---------------------+
'+   Ã´∞—µÃﬁºﬁ™∏ƒè¡ãé
'+---------------------+
    Call CisFrm.UnLoadDisp
    End
End Sub
'+----------------------------+
'+    ì¸  óÕ  ÉL Å[(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
   Call ReturnPress
End Sub
'+----------------------------+
'+    éÊ  è¡  ÉL Å[(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
   Call CisFrm.MousePT(11)
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      iH1_YM.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    êV  ãK  ÉL Å[(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    If SyoriKB <> "ADD" Then
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
    End If
    iH1_YM.SetFocus
End Sub
'+----------------------------+
'+    èC  ê≥  ÉL Å[(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
    End If
    iH1_YM.SetFocus
End Sub
'+----------------------------+
'+    çÌ  èú  ÉL Å[(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        SyoriKB = "DEL"
    Else
        SyoriKB = ""
    End If
    Call SyoriName(SyoriKB)
    iH1_YM.SetFocus
End Sub
'+----------------------------+
'+    å„  ëﬁ  ÉL Å[(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = "Select * From ì‡é¶ÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where ì‡é¶îNåé + ïiî‘ + î[ì¸êÊ + éÛì¸ + ì‡é¶ãÊï™ < '" & _
                                      RTrim(iH1_YM) & RTrim(iH1_Hinbn) & RTrim(iH1_Torcd) & _
                                      RTrim(iH1_Ukeir) & RTrim(iH1_Ku) & "'"
    gSL_Select = gSL_Select & " ORDER BY ì‡é¶îNåé + ïiî‘ + î[ì¸êÊ + éÛì¸ + ì‡é¶ãÊï™ Desc "
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_YM.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    ëO  êi  ÉL Å[(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = "Select * From ì‡é¶ÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where ì‡é¶îNåé + ïiî‘ + î[ì¸êÊ + éÛì¸ + ì‡é¶ãÊï™ > '" & _
                                      RTrim(iH1_YM) & RTrim(iH1_Hinbn) & RTrim(iH1_Torcd) & _
                                      RTrim(iH1_Ukeir) & RTrim(iH1_Ku) & "'"
    gSL_Select = gSL_Select & " ORDER BY ì‡é¶îNåé + ïiî‘ + î[ì¸êÊ + éÛì¸ + ì‡é¶ãÊï™ "

    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_YM.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    ïœçXëOÇ…ñﬂÇ∑(F5)         +
'+----------------------------+
Private Sub PB_ReDisp_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
            
    For gInt = 0 To 36
        iB1_Sur(gInt) = B1lb_Kaku(gInt).Tag
        Call iB1_Sur_LostFocus(gInt)
    Next gInt
    
    Dummy.Enabled = False
    
    For Each gObj In B_Sur
        If gObj.Visible Then
           iB1_Sur(gObj.INDEX).SetFocus
           Exit For
        End If
    Next gObj
End Sub
'+----------------------------+
'+    ì˙ó êîÇëSÇƒè¡ãé(F8)     +
'+----------------------------+
Private Sub PB_CLR_Click()
    For gInt = 0 To 36
        iB1_Sur(gInt) = ""
        Call iB1_Sur_LostFocus(gInt)
    Next gInt
End Sub
'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd(0): Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd: Exit Sub
    If PB_Look.Tag = "iH1_Ku" Then Call Look_Kubun(0): Exit Sub
    If PB_Look.Tag = "iB1_KTorcd" Then Call Look_Torcd(1): Exit Sub
    If PB_Look.Tag = "iB1_Nouku" Then Call Look_Kubun(1): Exit Sub
    If PB_Look.Tag = "iB1_SDate" Then iB1_SDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iB1_EDate" Then iB1_EDate.ShowCalender: Exit Sub
End Sub
'+-----------------------------------+
'+       ïiî‘åüçı                     +
'+-----------------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 0
    RVI_SK = 9
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Hinbn = RV_Hinbn
        iH1_Torcd = RV_TorcdK
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+-----------------------------------+
'+       å_ñÒêÊåüçı(î[ì¸êÊÅEìæà”êÊ)    +
'+-----------------------------------+
Private Sub Look_Torcd(kb As Integer)
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    
    Select Case kb
        Case 0: RV_TorKb = 3
        Case Else: RV_TorKb = 2
    End Select
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        Select Case kb
            Case 0
                iH1_Torcd = RV_TorcdK
            Case 1
                iB1_KTorcd = RV_TorcdK
        End Select
    End If
End Sub
'+-----------------------------+
'+       éÊà¯êÊéÛì¸åüçı         +
'+-----------------------------+
Private Sub Look_UTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 3
    RV_Torcd = iH1_Torcd
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+----------------------------+
'+       ñºèÃÉ}ÉXÉ^åüçı        +
'+    KB : 0     ì‡é¶ãÊï™      +
'+       : 1     î[ì¸ãÊï™      +
'+----------------------------+
Private Sub Look_Kubun(kb As Integer)

    RV_Left = 0
    RV_Top = 0

    Select Case kb
        Case 0: RVI_Kubun = "ì‡é¶ãÊï™"
        Case Else: RVI_Kubun = "î[ì¸ãÊï™"
    End Select
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        Select Case kb
            Case 0
                iH1_Ku = RV_Code
                Call iH1_Ku_LostFocus
            Case 1
                iB1_Nouku = RV_Code
                Call iB1_Nouku_LostFocus
        End Select
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
'+  iH1_Ku
'+-----------------------------+
Private Sub iH1_Ku_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ku_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_Ku.Caption = ""
    If MeisyoGet("ì‡é¶ãÊï™", iH1_Ku) Then
        H1lb_Ku = SYM_Meisyo
        H1lb_Ku.Visible = True
    Else
        H1lb_Ku.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_SDate
'+-----------------------------+
Private Sub iB1_SDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_SDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_EDate
'+-----------------------------+
Private Sub iB1_EDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_EDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_SobetsuÅ@ëwï 
'+-----------------------------+
Private Sub iB1_Sobetsu_LostFocus()
    With CisComb1
        .Code = iB1_Sobetsu
        .TextGet
    End With
End Sub
'+-----------------------------+
'+  B1Comb_SobetsuÅ@ëwï 
'+-----------------------------+
Private Sub B1Comb_Sobetsu_GotFocus()
    B1Comb_Sobetsu.BackColor = gIGotSel
End Sub
Private Sub B1Comb_Sobetsu_LostFocus()
    B1Comb_Sobetsu.BackColor = gILostSel
End Sub
Private Sub B1Comb_Sobetsu_Click()
    With CisComb1
        .CodeGet
        iB1_Sobetsu = .Code
    End With
End Sub
'+-----------------------------+
'+  å_ñÒêÊ
'+-----------------------------+
Private Sub iB1_KTorcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_KTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Nouku
'+-----------------------------+
Private Sub iB1_Nouku_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Nouku_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_Nouku.Caption = ""
    If MeisyoGet("î[ì¸ãÊï™", iB1_Nouku) Then
        B1lb_Nouku = SYM_Meisyo
    End If
End Sub
'+-----------------------------+
'+  iB1_Sur
'+-----------------------------+
Private Sub iB1_Sur_LostFocus(INDEX As Integer)
    If CisFun.Val2(iB1_Sur(INDEX)) <> CisFun.Val2(B1lb_Kaku(INDEX)) Then
       B1lb_Kaku(INDEX).BackColor = mNHenColor
    Else
       B1lb_Kaku(INDEX).BackColor = mNInitColor
    End If
End Sub

'**************************************
'*                                    *
'*      èàóùãÊï™Ç…ÇÊÇËãÊï™ñºï\é¶        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim W_Color     As Long
   W_Color = &H80000008 'èâä˙êFæØƒ(çï)
   Select Case SyoriKB
     Case "ADD"
          SyoriNM = "í«â¡"
          K_Sykbnm.ForeColor = &HFFFF&   '(â©êF)
          PB_ADD.Visible = False
     Case "MNT"
          SyoriNM = "èCê≥"
          K_Sykbnm.ForeColor = &HFF0000  '(ê¬êF)
     Case "DEL"
          SyoriNM = "çÌèú"
          K_Sykbnm.ForeColor = &HFF&     '(ê‘êF)
          W_Color = &HFF&
     Case "LOK"
          SyoriNM = "è∆âÔ"
          K_Sykbnm.ForeColor = &HFF00FF   '(ÉsÉìÉN)
          PB_ADD.Visible = False
          PB_MNT.Visible = False
          PB_DEL.Visible = False
     Case Else
          SyoriNM = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
'      * PF-Key ï\é¶ê›íË *
       If SyoriKB <> "LOC" Then
          If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
             gObj.Name = "PB_DEL" Then
             gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
          End If
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
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
' ¿ﬁ–∞çÄñ⁄(âÊñ ÇÃÉ`ÉâÉcÉLñhé~)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1"
          H1Mode = True
          H1Color = gPGotSel
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B1_Area2.Enabled = B1Mode
    B1_Area2.BackColor = B1Color
    
    PB_ADD.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_DEL.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    
    PB_ReDisp.Visible = B1Mode
    PB_CLR.Visible = B1Mode
    
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
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
                .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "BX" Then
        CisComb1.ComboSet
    End If
    If HeadBody = "B" Then
       For gLong = 0 To 36
           B1lb_DD(gLong).BackColor = &HC0C0C0
           B1lb_Kaku(gLong).BackColor = mNInitColor
           B1lb_DD(gLong).Tag = ""
           iB1_Sur(gLong).Tag = ""
           B1lb_Kaku(gLong).Tag = ""
       Next gLong
    End If
End Sub
'****************************
'*      ì¸óÕèàóù(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
' << çÌèúèàóù >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        Else
            ProcHB = "B1"
            Call DispChange(ProcHB)
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
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
'*      HEAD1¡™Ø∏èàóù
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
    If Trim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ëŒè€îNåéÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_YM.SetFocus
       Exit Function
    End If
    If Trim(iH1_Hinbn) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ïiî‘Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Hinbn.SetFocus
       Exit Function
    End If
    If Trim(iH1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    î[ì¸êÊÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
    gSL_Select = " SELECT * "
    gSL_Select = gSL_Select & " FROM éÊà¯êÊÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " WHERE éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   And éÊà¯êÊãÊï™ = 0"
    If Not TRMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    î[ì¸êÊ(éÊà¯êÊÉ}ÉXÉ^ñ¢ìoò^)            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
    gSL_Select = " SELECT * "
    gSL_Select = gSL_Select & " FROM éÊà¯êÊéÛì¸É}ÉXÉ^ "
    gSL_Select = gSL_Select & " Where éÊà¯êÊCD = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÊà¯êÊãÊï™ = 0"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    If Not TUMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    éÛì¸(éÊà¯êÊéÛì¸É}ÉXÉ^ñ¢ìoò^)            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Ukeir.SetFocus
       Exit Function
    End If
    
    If Trim(iH1_Ku) <> "" Then
       If Not MeisyoGet("ì‡é¶ãÊï™", iH1_Ku) Then
          With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "    ãÊÅ@ñ¢ìoò^            "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
          End With
          iH1_Ku.SetFocus
        Exit Function
       End If
    End If
    
    '
    gSL_Select = "Select * From ì‡é¶ÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   î[ì¸êÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & " And   ì‡é¶ãÊï™ = '" & RTrim(iH1_Ku) & "'"
    If DBInput("INV") Then
       If SyoriKB = "ADD" Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    ì‡é¶ÉfÅ[É^ìoò^çœ            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iH1_YM.SetFocus
          Exit Function
       End If
       If SyoriKB = "DEL" Then
          If Trim(NJT.ìoò^FCODE) <> "" Then
             With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ÉIÉìÉâÉCÉìÉfÅ[É^ÇÃçÌèúÇÕèoóàÇ‹ÇπÇÒÅB            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
             End With
             iH1_YM.SetFocus
             Exit Function
          End If
       End If
    Else
       If SyoriKB <> "ADD" Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    ì‡é¶ÉfÅ[É^ñ¢ìoò^            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iH1_YM.SetFocus
          Exit Function
       End If
    End If
    If SyoriKB = "DEL" Then
    End If
    
    Head1Chk = True
End Function
'****************************
'*      Body1¡™Ø∏èàóù
'****************************
Private Function Body1Chk() As Boolean
    
    Body1Chk = False
    ' ëwï ãÊï™
    If Trim(iB1_Sobetsu) <> "" And Trim(B1Comb_Sobetsu) = "" Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(1) = "      ëwï ãÊï™ñ¢ìoò^            "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
           iB1_Sobetsu.SetFocus
           Exit Function
       End With
    End If
    ' ìKópäJén-èIóπ
    If Trim(iB1_SDate) <> "" Then
       If Not iB1_SDate.cChkResult Then
          With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "      ìKópä˙ä‘(äJén)Çê≥ÇµÇ¢ì˙ïtÇ…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Button = Error
               .MB_Title = ""
               .MBOX
               iB1_SDate.SetFocus
               Exit Function
           End With
       End If
    End If
    If Trim(iB1_EDate) <> "" Then
       If Not iB1_EDate.cChkResult Then
          With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "      ìKópä˙ä‘(èIóπ)Çê≥ÇµÇ¢ì˙ïtÇ…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
               iB1_EDate.SetFocus
               Exit Function
           End With
       End If
    End If
    If Trim(iB1_SDate) <> "" And Trim(iB1_EDate) <> "" Then
       If Trim(iB1_SDate) > Trim(iB1_EDate) Then
          With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "      ìKópä˙ä‘Çê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢ÅiäJénÅÑèIóπÅj            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
               iB1_SDate.SetFocus
               Exit Function
           End With
       End If
    End If
    ' å_ñÒêÊ
'    If Trim(iB1_KTorcd) = "" Then
'       With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "      å_ñÒêÊÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
'            .MB_Button = Error
'            .MB_Title = ""
'            .MBOX
'            iB1_KTorcd.SetFocus
'            Exit Function
'        End With
'    End If
 ' îwî‘çÜîºäpÉ`ÉFÉbÉN
    If RTrim(iB1_Seban) <> "" And Not HalfSizeChk(iB1_Seban) Then
        With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "    îwî‘çÜÇÕîºäpÇ…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
        End With
        iB1_Seban.SetFocus
        Exit Function
    End If
    If Trim(iB1_KTorcd) <> "" And Not TorNmGet(iB1_KTorcd, 2) Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "      å_ñÒêÊÅ@ñ¢ìoò^            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
            iB1_KTorcd.SetFocus
            Exit Function
        End With
    End If
    ' ÉTÉCÉNÉã
'----------------------------------------- 2006/0309 Delete
'    If Trim(iB1_CyCle1) <> "" Or Trim(iB1_CyCle2) <> "" Or Trim(iB1_CyCle3) <> "" Then
'       If Trim(iB1_CyCle1) = "" Or Trim(iB1_CyCle2) = "" Or Trim(iB1_CyCle3) = "" Then
'           With CisFun
'               .MB_Lines = 4
'               .MB_MSG(2) = "    ÉTÉCÉNÉãÇê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
'               .MB_Title = ""
'               .MB_Button = Error
'               .MBOX
'           End With
'          iB1_CyCle1.SetFocus
'          Exit Function
'       End If
'    End If
'----------------------------------------- 2006/03/09 Delete End
'----------------------------------------- 2006/03/09 Insert
    If Not IsNumeric(iB1_CyCle1) Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ÉTÉCÉNÉãÇÕêîéöÇ…Çƒì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_CyCle1.SetFocus
       Exit Function
    End If
'----------------------------------------- 2006/03/09 Insert End
    ' ì‡é¶êîÉ`ÉFÉbÉN
    mNSu = 0
    For gLong = 0 To 36
        If B_Sur(gLong).Visible Then
           mNSu = mNSu + CisFun.Val2(iB1_Sur(gLong))
        End If
    Next gLong
    ' ì˙ó êîÇ™ñ¢ì¸óÕÇÃéûÅAçƒåvéZ
    If mNSu = 0 Then
        Call DayReCalc
    End If
    
    If CisFun.Val2(iB1_TNSu) <> mNSu Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(1) = "      ìñåéì‡é¶êîÇ™ì˙ó çáåvÇ∆àŸÇ»ÇËÇ‹Ç∑ÅB            "
           .MB_MSG(3) = "      ìñåéì‡é¶êîÇì˙ó çáåvêîÇ∆ÇµÇ‹Ç∑Ç©ÅH            "
           .MB_Title = ""
           .MB_Button = Yes_No
           If Not .MBOX Then
              iB1_TNSu.SetFocus
              Exit Function
           End If
           iB1_TNSu = CisFun.RSetFld(mNSu, 0, iB1_TNSu.cFormat)
       End With
    End If
    
    Body1Chk = True
End Function
'+-----------------------------------------+
'+      ì˙ó êîçƒåvéZ
'+-----------------------------------------+
Private Sub DayReCalc()
    Static wSetDay      As Long
    
    wSetDay = 0
    For gLong = 0 To 36
        If B_Sur(gLong).Visible And B_Sur(gLong).Tag = "" Then
            wSetDay = gLong
            Exit For
        End If
    Next gLong
    
    With CisDB
        .SQL = "ì˙ó êîéZèo"
        .StoadoCount = 35
        .ParaNo = 0:    .ParaIO = Returnóp
        .ParaNo = 1:    .ParaIO = Inputóp
        .ParaNo = 2:    .ParaIO = Inputóp
        .ParaNo = 3:    .ParaIO = Inputóp
        .ParaNo = 4:    .ParaIO = Inputóp

        For gLong = 5 To 35
            .ParaNo = gLong: .ParaIO = OutPutóp
        Next gLong

        .ParaNo = 1:    .ParaValue = iH1_YM
        .ParaNo = 2:    .ParaValue = CisFun.Val2(iB1_TNSu)
        .ParaNo = 3:    .ParaValue = RTrim$(iH1_Torcd)
        .ParaNo = 4:    .ParaValue = RTrim$(iH1_Ukeir)

        .DBStored

        For gInt = 5 To 35
            .ParaNo = gInt
            gLong = wSetDay + (gInt - 5)
            iB1_Sur(gLong) = Format(.ParaValue, "#")
            Call iB1_Sur_LostFocus(CisFun.Val2(gLong))
            mNSu = mNSu + CisFun.Val2(iB1_Sur(gLong))
        Next gInt
    End With
End Sub
'**************************************
'*                                    *
'*       ÉfÅ[É^ÉxÅ[ÉXì¸óÕ Åï ï\é¶       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("B")
    Call HeadBodyClear("BX")
    
    If Not NJTRead(gSL_Select, 1) Then
        If InputKey = "Back" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     êÊì™ÉåÉRÅ[ÉhÇ≈Ç∑            "
                .MB_Title = "Åyå„ëﬁÅz"
                .MB_Button = OK
                .MBOX
            End With
            GoTo DBInput_Ed
        End If
        If InputKey = "Next" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     ç≈èIÉåÉRÅ[ÉhÇ≈Ç∑            "
                .MB_Title = "ÅyëOêiÅz"
                .MB_Button = OK
                .MBOX
            End With
            GoTo DBInput_Ed
        End If
        NJT.ì‡é¶îNåé = ""
        Call DaySet
    Else
        iH1_YM = RTrim(NJT.ì‡é¶îNåé)
        iH1_Hinbn = RTrim(NJT.ïiî‘)
        iH1_Torcd = RTrim(NJT.î[ì¸êÊ)
        iH1_Ukeir = RTrim(NJT.éÛì¸)
        iH1_Ku = RTrim(NJT.ì‡é¶ãÊï™)
        Call iH1_Ku_LostFocus
        Call DaySet
        Call BodySet
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
Private Sub BodySet()
    With NJT
        iB1_Sobetsu = .ëwï ãÊï™
        With CisComb1
            .Code = NJT.ëwï ãÊï™
            .TextGet
        End With
        iB1_TNSu = CisFun.RSetFld(.ì‡é¶êî, 0, iB1_TNSu.cFormat)
        iB1_TYSu = CisFun.RSetFld(.óÇåéì‡é¶êî, 0, iB1_TYSu.cFormat)
        iB1_TYYSu = CisFun.RSetFld(.óÇÅXåéì‡é¶êî, 0, iB1_TYYSu.cFormat)
        iB1_SDate = .ìKópäJénîNåéì˙
        iB1_EDate = .ìKópèIóπîNåéì˙
        
        iB1_Kouku = .çHãÊ
        iB1_KTorcd = .å_ñÒêÊ
        iB1_Seban = .îwî‘çÜ
        iB1_CyCle1 = .ÉTÉCÉNÉã
'--------------------------------------------------- 2006/03/09 Delete
'        iB1_CyCle1 = Mid(.ÉTÉCÉNÉã, 1, 1)
'        iB1_CyCle2 = CisFun.RSetFld(Mid(.ÉTÉCÉNÉã, 2, 2), 0, iB1_CyCle2.cFormat)
'        iB1_CyCle3 = CisFun.RSetFld(Mid(.ÉTÉCÉNÉã, 4, 2), 0, iB1_CyCle3.cFormat)
'--------------------------------------------------- 2006/03/09 Delete End
        iB1_Syuyo = CisFun.RSetFld(.é˚óeêî, 0, iB1_Syuyo.cFormat)
        iB1_Kanri = .ä«óùNO
        iB1_Seibi = .êÆîı
        iB1_Tanto = .íSìñ
        iB1_Nouku = .î[ì¸ãÊï™
        Call iB1_Nouku_LostFocus
        iB1_NouSu = CisFun.RSetFld(.î[ïièëî≠çsñáêî, 0, iB1_NouSu.cFormat)
        iB1_KMai = CisFun.RSetFld(.Ç©ÇÒÇŒÇÒâÒì]ñáêî, 0, iB1_KMai.cFormat)
        If .Ç©ÇÒÇŒÇÒëùå∏ïÑçÜ = "-" Then
           .Ç©ÇÒÇŒÇÒëùå∏ñáêî = .Ç©ÇÒÇŒÇÒëùå∏ñáêî * -1
        End If
        iB1_KMaiZ = CisFun.RSetFld(.Ç©ÇÒÇŒÇÒëùå∏ñáêî, 0, iB1_KMaiZ.cFormat)
        If .ïœçXêîïÑçÜ = "-" Then
           .ïœçXêî = .ïœçXêî * -1
        End If
        iB1_Hsu = CisFun.RSetFld(.ïœçXêî, 0, iB1_Hsu.cFormat)
        If .í≤êÆêîïÑçÜ = "-" Then
           .í≤êÆêî = .í≤êÆêî * -1
        End If
        iB1_CSu = CisFun.RSetFld(.í≤êÆêî, 0, iB1_CSu.cFormat)
        iB1_Comment = RTrim(.ÉRÉÅÉìÉg)
        If Trim(.ìoò^FCODE) <> "" Then
            B1lb_DtataNm = "ÉIÉìÉâÉCÉìÉfÅ[É^"
            If .ïœçXâÒêî <> 0 Then
               B1lb_DtataNm = B1lb_DtataNm & "(åvïœ)"
            Else
               B1lb_DtataNm = B1lb_DtataNm & "(" & mNaiji & ")"
            End If
        Else
            B1lb_DtataNm = ""
        End If
        If .ìoò^ãÊï™ <> 0 Then
           B1lb_Mnt = "ïœçX"
        Else
           B1lb_Mnt = ""
        End If
    End With
                    
    ' ì‡é¶ëOâÒÉeÅ[ÉuÉã ì«éÊ
    gSL_Select = "Select * From ì‡é¶ëOâÒÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   î[ì¸êÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & " And   ì‡é¶ãÊï™ = '" & RTrim(iH1_Ku) & "'"
    If Not NZTRead(gSL_Select, 1) Then
        Call ItemsClearNZT
    End If
    
    ' ì‡é¶êîÅiì˙ó Åj
    gInt = 1
    For gLong = 0 To 36
        If B_Sur(gLong).Visible Then
           iB1_Sur(gLong) = CisFun.RSetFld(NJT.ì‡é¶êîó (gInt), 0, iB1_Sur(0).cFormat)
           B1lb_Kaku(gLong) = Format(NZT.ì‡é¶êîó (gInt), iB1_Sur(0).cFormat)
           B1lb_Kaku(gLong).Tag = iB1_Sur(gLong)
           Call iB1_Sur_LostFocus(CisFun.Val2(gLong))
           gInt = gInt + 1
        End If
    Next gLong
    
End Sub
'+-------------------------+
'+  ì˙ïtèÓïÒÉZÉbÉg
'+-------------------------+
Private Sub DaySet()
    Dim wSetDay     As Long
    Dim Cl_Select   As String
    
    If Trim(iH1_YM) = "" Then Exit Sub
' ì‡é¶îNåéÉZÉbÉg
    If Trim(NJT.ì‡é¶îNåé) = "" Then
       gStr = iH1_YM
    Else
       gStr = NJT.ì‡é¶îNåé
    End If
    NaijiT = StrConv(Mid(gStr, 5, 2), vbWide) & "åéì‡é¶êî"
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add gStr & "01", M, 1, LastDate
        gStr = .Date_Result
        NaijiY = StrConv(CisFun.RSetFld(Mid(gStr, 5, 2), 2, "##"), vbWide) & "åéì‡é¶êî"
        .Date_RsFormat = fYM
        .Date_Add gStr & "01", M, 1, LastDate
        gStr = .Date_Result
        NaijiYY = StrConv(CisFun.RSetFld(Mid(gStr, 5, 2), 2, "##"), vbWide) & "åéì‡é¶êî"
    End With


'       *  ÉJÉåÉìÉ_ïîì˙ït ÉZÉbÉg  *
    Cl_Select = "Select * From ÉJÉåÉìÉ_É}ÉXÉ^"
    Cl_Select = Cl_Select & " Where éÊà¯êÊãÊï™ = 0"
    Cl_Select = Cl_Select & " And   éÊà¯êÊ = '" & RTrim(iH1_Torcd) & "'"
    Cl_Select = Cl_Select & " And   îN = '" & Format(iH1_YM.cYear, "0000") & "'"
    Cl_Select = Cl_Select & " And   åé = '" & Format(iH1_YM.cMonth, "00") & "'"
    Cl_Select = Cl_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    If Not CLMRead(Cl_Select, 1) Then
       Cl_Select = "Select * From ÉJÉåÉìÉ_É}ÉXÉ^"
       Cl_Select = Cl_Select & " Where éÊà¯êÊãÊï™ = 0"
       Cl_Select = Cl_Select & " And   éÊà¯êÊ = '" & RTrim(iH1_Torcd) & "'"
       Cl_Select = Cl_Select & " And   îN = '" & Format(iH1_YM.cYear, "0000") & "'"
       Cl_Select = Cl_Select & " And   åé = '" & Format(iH1_YM.cMonth, "00") & "'"
       Cl_Select = Cl_Select & " And   éÛì¸ = ''"
       If Not CLMRead(Cl_Select, 1) Then
          Cl_Select = "Select * From ÉJÉåÉìÉ_É}ÉXÉ^"
          Cl_Select = Cl_Select & " Where éÊà¯êÊãÊï™ = 0"
          Cl_Select = Cl_Select & " And   éÊà¯êÊ = ''"
          Cl_Select = Cl_Select & " And   îN = '" & Format(iH1_YM.cYear, "0000") & "'"
          Cl_Select = Cl_Select & " And   åé = '" & Format(iH1_YM.cMonth, "00") & "'"
          Cl_Select = Cl_Select & " And   éÛì¸ = ''"
          If Not CLMRead(Cl_Select, 1) Then
'             With CisFun
'                .MB_Lines = 3
'                .MB_MSG(2) = "     ÉJÉåÉìÉ_É}ÉXÉ^ñ¢ìoò^(é©é–)  "
'                .MB_Title = "∂⁄›¿ﬁœΩ¿ämîF"
'                .MB_Button = OK
'                .MBOX
'             End With
             Call ItemsClearCLM
             CLM.â“ìÆãÊï™ = String(31, "0")
             CLM.ñÈãŒãÊï™ = String(31, "0")
             With CisFun
                 .Date_RsFormat = fDD
                 .Date_End iH1_YM
                 CLM.ññì˙ = .Val2(.Date_Result)
                 CLM.â“ìÆì˙êî = CLM.ññì˙
             End With
          End If
        End If
    End If
    
    For gLong = 0 To 36
        B1lb_DD(gLong).BackColor = &HC0C0C0
    Next gLong
    gDate = DateValue(iH1_YM.cYear & "/" & iH1_YM.cMonth & "/01")
    gInt = Weekday(gDate, mWeekStart) - 1
    wSetDay = 0
    For gLong = 0 To 36
        If gLong < gInt Or wSetDay >= CLM.ññì˙ Then
           B_Sur(gLong).Visible = False
           B1lb_DD(gLong).Tag = ""
           iB1_Sur(gLong).Tag = ""
        Else
           B_Sur(gLong).Visible = True
           wSetDay = wSetDay + 1
           B1lb_DD(gLong) = StrConv(wSetDay, vbWide)
           B1lb_DD(gLong).Tag = wSetDay
           iB1_Sur(gLong).Tag = ""
           If Mid(CLM.â“ìÆãÊï™, wSetDay, 1) = "1" And _
              Mid(CLM.ñÈãŒãÊï™, wSetDay, 1) = "1" Then
              B1lb_DD(gLong).BackColor = mSundayColor
              iB1_Sur(gLong).Tag = "1"
           End If
        End If
    Next gLong
    If BLB_Week(0) = "ì˙" Then
        For gLong = 0 To 36 Step 7
            If iB1_Sur(gLong).Tag = "1" Then
                B1lb_DD(gLong).BackColor = mSundayColor
            End If
        Next gLong
        For gLong = 6 To 36 Step 7
            If iB1_Sur(gLong).Tag = "1" Then
                B1lb_DD(gLong).BackColor = mSaturdayColor
            End If
        Next gLong
    Else
        For gLong = 6 To 36 Step 7
            If iB1_Sur(gLong).Tag = "1" Then
               B1lb_DD(gLong).BackColor = mSundayColor
               B1lb_DD(gLong - 1).BackColor = mSaturdayColor
            End If
        Next gLong
    End If
    
    B1lb_Kado = Format(CLM.â“ìÆì˙êî, "#0.0")
    
End Sub

'**************************************
'*                                    *
'*          ÉfÅ[É^ÉxÅ[ÉXçXêV           *
'*                                    *
'**************************************
'+----------------------+
'+      çÌ èú èà óù      +
'+----------------------+
Private Function DBDelete()
    DBDelete = False
'  ( ämîFÉÅÉbÉZÅ[ÉW )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     çÌèúèàóùÇçsÇ¢Ç‹Ç∑      "
        .MB_Title = "çÌèúèàóù"
        .MB_MSG(3) = "     ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ƒ◊›ªﬁ∏ºÆ›äJén
' ( çÌèúèàóù )
        ' éÊà¯êÊÉ}ÉXÉ^
        .SQL = "DELETE FROM ì‡é¶ÉeÅ[ÉuÉã "
        .SQL = .SQL & " Where ì‡é¶îNåé = '" & RTrim(iH1_YM) & "'"
        .SQL = .SQL & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
        .SQL = .SQL & " And   î[ì¸êÊ = '" & RTrim(iH1_Torcd) & "'"
        .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
        .SQL = .SQL & " And   ì‡é¶ãÊï™ = '" & RTrim(iH1_Ku) & "'"
        .DBExec
        .DBTran (TransCommit) ' ƒ◊›ªﬁ∏ºÆ› èIóπ
        
        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")
        
        On Error GoTo 0
    End With
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    í«â¡•çXêV èàóù     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Dim AddFlg  As Boolean

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        If SyoriKB = "ADD" Then
            .MB_MSG(2) = "     í«â¡èàóùÇçsÇ¢Ç‹Ç∑      "
            .MB_Title = "í«â¡èàóù"
        Else
            .MB_MSG(2) = "     èCê≥èàóùÇçsÇ¢Ç‹Ç∑      "
            .MB_Title = "èCê≥èàóù"
        End If
        .MB_MSG(3) = "     ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err
'   çÄñ⁄à⁄ëó
    gSL_Select = "Select * From ì‡é¶ÉeÅ[ÉuÉã"
    gSL_Select = gSL_Select & " Where ì‡é¶îNåé = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & " And   î[ì¸êÊ = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & " And   ì‡é¶ãÊï™ = '" & RTrim(iH1_Ku) & "'"
    
    If NJTRead(gSL_Select, 1) Then
       AddFlg = False
    Else
       Call ItemsClearNJT
       AddFlg = True
       NJT.ì‡é¶îNåé = RTrim(iH1_YM)
       NJT.ïiî‘ = RTrim(iH1_Hinbn)
       NJT.î[ì¸êÊ = RTrim(iH1_Torcd)
       NJT.éÛì¸ = RTrim(iH1_Ukeir)
       NJT.ì‡é¶ãÊï™ = RTrim(iH1_Ku)
    End If
    With NJT
        .å_ñÒêÊ = iB1_KTorcd
        .çHãÊ = iB1_Kouku
        .ëwï ãÊï™ = iB1_Sobetsu
        .çÜï‚ãÊï™ = ""
        .åüçıïiî‘ = ""
        .ìKópäJénîNåéì˙ = iB1_SDate
        .ìKópèIóπîNåéì˙ = iB1_EDate
        .ì‡é¶êî = CisFun.Val2(iB1_TNSu)
        .óÇåéì‡é¶êî = CisFun.Val2(iB1_TYSu)
        .óÇÅXåéì‡é¶êî = CisFun.Val2(iB1_TYYSu)
        .ä«óùNO = iB1_Kanri
        .é˚óeêî = CisFun.Val2(iB1_Syuyo)
        .îwî‘çÜ = iB1_Seban
        .ÉTÉCÉNÉã = iB1_CyCle1
'--------------------------------------------------- 2006/03/09 Delete
'        If CisFun.Val2(iB1_CyCle1) <> 0 Or CisFun.Val2(iB1_CyCle2) <> 0 Or _
'           CisFun.Val2(iB1_CyCle3) <> 0 Then
'           .ÉTÉCÉNÉã = Format(iB1_CyCle1, "0") & _
'                         Format(iB1_CyCle2, "00") & _
'                         Format(iB1_CyCle3, "00")
'        End If
'--------------------------------------------------- 2006/03/09 Delete End
        .íSìñ = iB1_Tanto
        .êÆîı = iB1_Seibi
        .î[ì¸ãÊï™ = iB1_Nouku
        .î[ïièëî≠çsñáêî = CisFun.Val2(iB1_NouSu)
        .Ç©ÇÒÇŒÇÒâÒì]ñáêî = CisFun.Val2(iB1_KMai)
        If CisFun.Val2(iB1_KMaiZ) < 0 Then
            .Ç©ÇÒÇŒÇÒëùå∏ïÑçÜ = "-"
            .Ç©ÇÒÇŒÇÒëùå∏ñáêî = CisFun.Val2(iB1_KMaiZ) * -1
        Else
            .Ç©ÇÒÇŒÇÒëùå∏ïÑçÜ = ""
            .Ç©ÇÒÇŒÇÒëùå∏ñáêî = CisFun.Val2(iB1_KMaiZ)
        End If
        If CisFun.Val2(iB1_Hsu) < 0 Then
            .ïœçXêîïÑçÜ = "-"
            .ïœçXêî = CisFun.Val2(iB1_Hsu) * -1
        Else
            .ïœçXêîïÑçÜ = ""
            .ïœçXêî = CisFun.Val2(iB1_Hsu)
        End If
        If CisFun.Val2(iB1_CSu) < 0 Then
            .í≤êÆêîïÑçÜ = "-"
            .í≤êÆêî = CisFun.Val2(iB1_CSu) * -1
        Else
            .í≤êÆêîïÑçÜ = ""
            .í≤êÆêî = CisFun.Val2(iB1_CSu)
        End If
        gInt = 0
        For gLong = 0 To 36
            If B_Sur(gLong).Visible Then
               gInt = gInt + 1
               .ì‡é¶êîó (gInt) = CisFun.Val2(iB1_Sur(gLong))
            End If
        Next gLong
        .ÉRÉÅÉìÉg = iB1_Comment
    End With
'
    Call CisDB.DBTran(TransBegin)  ' ƒ◊›ªﬁ∏ºÆ›äJén
' ( í«â¡èàóù )
    If AddFlg Then
        NJT.çÏê¨é“ = gTanto
        Call NJTInsert
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
    Else
' ( çXêVèàóù )
        NJT.çXêVé“ = gTanto
        NJT.ìoò^ãÊï™ = 1

        Call NJTUpdate
        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End If
' ÉIÉìÉâÉCÉìÉfÅ[É^à»äOÇÃéûÅAì‡é¶ëOâÒÉeÅ[ÉuÉãÇ÷èoóÕ
    If Trim(NJT.ìoò^FCODE) = "" Then
       With CisDB
            .SQL = "Delete ì‡é¶ëOâÒÉeÅ[ÉuÉã"
            .SQL = .SQL & " Where ì‡é¶îNåé = '" & RTrim(iH1_YM) & "'"
            .SQL = .SQL & " And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
            .SQL = .SQL & " And   î[ì¸êÊ = '" & RTrim(iH1_Torcd) & "'"
            .SQL = .SQL & " And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
            .SQL = .SQL & " And   ì‡é¶ãÊï™ = '" & RTrim(iH1_Ku) & "'"
            .DBExec
            
            .SQL = "Insert Into ì‡é¶ëOâÒÉeÅ[ÉuÉã "
            .SQL = .SQL & " Select * From ì‡é¶ÉeÅ[ÉuÉã "
            .SQL = .SQL & "    Where ì‡é¶îNåé = '" & RTrim(iH1_YM) & "'"
            .SQL = .SQL & "    And   ïiî‘ = '" & RTrim(iH1_Hinbn) & "'"
            .SQL = .SQL & "    And   î[ì¸êÊ = '" & RTrim(iH1_Torcd) & "'"
            .SQL = .SQL & "    And   éÛì¸ = '" & RTrim(iH1_Ukeir) & "'"
            .SQL = .SQL & "    And   ì‡é¶ãÊï™ = '" & RTrim(iH1_Ku) & "'"
            .DBExec
       End With
    End If
    CisDB.DBTran (TransCommit) ' ƒ◊›ªﬁ∏ºÆ› èIóπ
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function

'+---------------------------------------+
'+      ñºèÃÉ}ÉXÉ^èâä˙ê›íË
'+---------------------------------------+
Private Sub MeisyoInit()
' ì‡é¶ãÊï™
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "ì‡é¶ãÊï™"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .êßå‰ãÊï™ = "X"
        .ílñºèÃ = "ä™ï‘Çµ"
        .çÏê¨é“ = gTanto
        .åüçıñº = "ì‡ é¶ ãÊ ï™ åü çı"
        .ÉRÅ[Éhñº = "ì‡é¶ãÊ"
        .ñºèÃñº = "ì‡é¶ãÊï™ñº"
    End With
    If Not SYMInitCreate(True) Then
    End If
' ëwï 
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "ëwï ãÊï™"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .êßå‰ãÊï™ = "X"
        .ílñºèÃ = "çÜå˚ïi"
        .çÏê¨é“ = gTanto
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .íl = "2"
            .ílñºèÃ = "íºëóéxããïi"
        End With
        Call SYMInitCreate
        With SYMCreate
            .íl = "3"
            .ílñºèÃ = "äOíçï ëóïi"
        End With
        Call SYMInitCreate
        With SYMCreate
            .íl = "4"
            .ílñºèÃ = "ï‚ããïi"
        End With
        Call SYMInitCreate
        With SYMCreate
            .íl = "5"
            .ílñºèÃ = "ì‡êªéxããïi"
        End With
        Call SYMInitCreate
    End If
' î[ì¸ãÊï™
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "î[ì¸ãÊï™"
        .ãÊï™åÖêî = 1
        .íl = "A"
        .êßå‰ãÊï™ = "X"
        .ílñºèÃ = "ï ìréwé¶ïîïi"
        .çÏê¨é“ = gTanto
        .åüçıñº = "î[ ì¸ ãÊ ï™ åü çı"
        .ÉRÅ[Éhñº = "î[ì¸ãÊ"
        .ñºèÃñº = "î[ì¸ãÊï™ñº"
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .íl = "B"
            .ílñºèÃ = "Daily éwé¶ïîïi"
        End With
        Call SYMInitCreate
        With SYMCreate
            .íl = "C"
            .ílñºèÃ = "Ç©ÇÒÇŒÇÒïîïi"
        End With
        Call SYMInitCreate
    End If
' ójì˙äJénä«óù
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "ójì˙äJénä«óù"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .êßå‰ãÊï™ = "X"
        .ílñºèÃ = "êîéö1=0:ì˙ójì˙,1:åéójì˙"
        .êîéö1 = 1
        .çÏê¨é“ = gTanto
    End With
    If Not SYMInitCreate(True) Then
    End If
End Sub

