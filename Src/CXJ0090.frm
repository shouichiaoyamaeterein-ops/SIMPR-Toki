VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form CXJ0090 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "éÛíçâÊñ ì¸óÕ(∂⁄›¿ﬁå`éÆ)"
   ClientHeight    =   10980
   ClientLeft      =   1530
   ClientTop       =   1725
   ClientWidth     =   15315
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
   ScaleHeight     =   10980
   ScaleWidth      =   15315
   WindowState     =   2  'ç≈ëÂâª
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
      Height          =   705
      Left            =   7500
      ScaleHeight     =   645
      ScaleWidth      =   7680
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   915
      Width           =   7740
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   5115
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
         Left            =   5970
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
         Left            =   6825
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
         Left            =   2565
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
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   30
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1710
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
      Begin CisBtn_60.CisBtn PB_BACK 
         Height          =   585
         Left            =   3420
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   4260
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
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Ç»Çµ
      Height          =   345
      Index           =   1
      Left            =   10305
      ScaleHeight     =   345
      ScaleWidth      =   3675
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   30
      Width           =   3675
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   345
         Index           =   1
         Left            =   2445
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BackColor       =   14737632
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
            Top             =   45
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   450
            BackColor       =   8421631
            ForeColor       =   0
            Caption         =   "ZZ9"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
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
         Left            =   15
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BackColor       =   14737632
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
               Size            =   12
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
         BackColor       =   14737632
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
            Top             =   45
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   450
            BackColor       =   16777088
            ForeColor       =   0
            Caption         =   "ZZ9"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
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
   Begin VB.PictureBox B1_Area2 
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
      Height          =   2745
      Left            =   9990
      ScaleHeight     =   2685
      ScaleWidth      =   5190
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   2790
      Width           =   5250
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   645
         Left            =   45
         Top             =   2010
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   1138
         Caption         =   "èoâ◊ä«óùïîèê"
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
         cAlingnment     =   1
         cPositionX      =   120
         Begin Cis3D_v60.CIS3D B1lb_SyukaK 
            Height          =   375
            Left            =   1665
            Top             =   135
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   661
            BackColor       =   16777152
            ForeColor       =   0
            Caption         =   "XXXXXXXXX1XXXXX"
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
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1185
         Index           =   6
         Left            =   1935
         Top             =   825
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2090
         ForeColor       =   0
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
         cPositionX      =   30
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Syuyo 
            Height          =   360
            Index           =   1
            Left            =   255
            Top             =   375
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   635
            BackColor       =   16777152
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
            cPositionX      =   -50
         End
         Begin Cis3D_v60.CIS3D B1lb_Syuyo 
            Height          =   360
            Index           =   2
            Left            =   255
            Top             =   735
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   635
            BackColor       =   16777152
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
            cPositionX      =   -50
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1185
         Index           =   7
         Left            =   3390
         Top             =   825
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   2090
         ForeColor       =   0
         Caption         =   "é˚óeäÌ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionX      =   30
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Youki 
            Height          =   360
            Index           =   1
            Left            =   240
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1"
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
            cPositionX      =   50
         End
         Begin Cis3D_v60.CIS3D B1lb_Youki 
            Height          =   360
            Index           =   2
            Left            =   240
            Top             =   735
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1"
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
            cPositionX      =   50
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1185
         Index           =   5
         Left            =   45
         Top             =   825
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   2090
         ForeColor       =   0
         Caption         =   "é˚óeêÿë÷ì˙"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionX      =   30
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_ChgDay 
            Height          =   360
            Left            =   285
            Top             =   510
            Width           =   1335
            _ExtentX        =   2355
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Index           =   1
         Left            =   45
         Top             =   45
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   1376
         ForeColor       =   0
         Caption         =   "ïiÅ@Å@Å@ñº"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionX      =   30
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B1lb_Hinnm 
            Height          =   360
            Left            =   105
            Top             =   315
            Width           =   4890
            _ExtentX        =   8625
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
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
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
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
      Height          =   945
      Left            =   150
      ScaleHeight     =   885
      ScaleWidth      =   9390
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   1785
      Width           =   9450
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Index           =   2
         Left            =   8685
         Top             =   60
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1376
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
         cPositionX      =   30
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D H2lb_Ukeir 
            Height          =   360
            Left            =   120
            Top             =   315
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   635
            BackColor       =   16777152
            ForeColor       =   0
            Caption         =   "XX"
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "î[  ì¸  êÊ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionX      =   30
         cPositionY      =   30
         Begin CisText_V60.CisText iH2_Torcd 
            Height          =   360
            Left            =   90
            TabIndex        =   9
            Top             =   330
            Width           =   1005
            _ExtentX        =   1773
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
            Text            =   "XXXXXXX"
            MaxLength       =   7
         End
         Begin Cis3D_v60.CIS3D H2lb_Tornm 
            Height          =   360
            Left            =   1080
            Top             =   330
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   635
            BackColor       =   16777152
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Index           =   3
         Left            =   3720
         Top             =   60
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "ïiÅ@Å@î‘"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionX      =   30
         cPositionY      =   30
         Begin CisText_V60.CisText iH2_Hinbn 
            Height          =   360
            Left            =   75
            TabIndex        =   10
            Top             =   315
            Width           =   3195
            _ExtentX        =   5636
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
            IMEMode         =   2
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Index           =   4
         Left            =   7065
         Top             =   60
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "îN åé"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
         cPositionX      =   30
         cPositionY      =   30
         Begin CisYM_V60.CisYM iH2_Ym 
            Height          =   360
            Left            =   120
            TabIndex        =   11
            Top             =   315
            Width           =   1395
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
      TabIndex        =   0
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
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   8025
      Left            =   150
      ScaleHeight     =   7965
      ScaleWidth      =   9765
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   2790
      Width           =   9825
      Begin Cis3D_v60.CIS3D BLB_Sur 
         Height          =   7860
         Index           =   1
         Left            =   45
         Top             =   45
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   13864
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
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   300
            Left            =   75
            Top             =   825
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "î[ì¸î‘çÜ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   36
            Left            =   2385
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   6585
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   36
               Left            =   -15
               TabIndex        =   85
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   36
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "36"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   36
               Left            =   0
               Top             =   915
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   36
               Left            =   -15
               TabIndex        =   84
               Top             =   285
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   35
            Left            =   1185
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   6585
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   35
               Left            =   -15
               TabIndex        =   83
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   35
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "35"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   35
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   35
               Left            =   -15
               TabIndex        =   82
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   34
            Left            =   8385
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   126
            TabStop         =   0   'False
            Top             =   5370
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   34
               Left            =   -15
               TabIndex        =   81
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   34
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "34"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   34
               Left            =   0
               Top             =   915
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   34
               Left            =   -15
               TabIndex        =   80
               Top             =   285
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   33
            Left            =   7185
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   125
            TabStop         =   0   'False
            Top             =   5370
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   33
               Left            =   -15
               TabIndex        =   79
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   33
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "33"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   33
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   33
               Left            =   -15
               TabIndex        =   78
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   32
            Left            =   5985
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   124
            TabStop         =   0   'False
            Top             =   5370
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   32
               Left            =   -15
               TabIndex        =   77
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   32
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "32"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   32
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   32
               Left            =   -15
               TabIndex        =   76
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   31
            Left            =   4785
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   5370
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   31
               Left            =   -15
               TabIndex        =   75
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   31
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "31"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   31
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   31
               Left            =   -15
               TabIndex        =   74
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   30
            Left            =   3585
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   122
            TabStop         =   0   'False
            Top             =   5370
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   30
               Left            =   -15
               TabIndex        =   73
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   30
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "30"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   30
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   30
               Left            =   -15
               TabIndex        =   72
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   29
            Left            =   2385
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   5370
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   29
               Left            =   -15
               TabIndex        =   71
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   29
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "29"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   29
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   29
               Left            =   -15
               TabIndex        =   70
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   28
            Left            =   1185
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   5370
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   28
               Left            =   -15
               TabIndex        =   69
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   28
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "28"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   28
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   28
               Left            =   -15
               TabIndex        =   68
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   27
            Left            =   8385
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   4155
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   27
               Left            =   -15
               TabIndex        =   67
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   27
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "27"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   27
               Left            =   0
               Top             =   915
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   27
               Left            =   -15
               TabIndex        =   66
               Top             =   285
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   26
            Left            =   7185
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   4155
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   26
               Left            =   -15
               TabIndex        =   65
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   26
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "26"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   26
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   26
               Left            =   -15
               TabIndex        =   64
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   25
            Left            =   5985
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   4155
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   25
               Left            =   -15
               TabIndex        =   63
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   25
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "25"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   25
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   25
               Left            =   -15
               TabIndex        =   62
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   24
            Left            =   4785
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   4155
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   24
               Left            =   -15
               TabIndex        =   61
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   24
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "24"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   24
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   24
               Left            =   -15
               TabIndex        =   60
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   23
            Left            =   3585
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   4155
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   23
               Left            =   -15
               TabIndex        =   59
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   23
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "23"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   23
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   23
               Left            =   -15
               TabIndex        =   58
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   22
            Left            =   2385
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   4155
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   22
               Left            =   -15
               TabIndex        =   57
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   22
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "22"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   22
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   22
               Left            =   -15
               TabIndex        =   56
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   21
            Left            =   1185
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   4155
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   21
               Left            =   -15
               TabIndex        =   55
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   21
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "21"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   21
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   21
               Left            =   -15
               TabIndex        =   54
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   20
            Left            =   8385
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   2940
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   20
               Left            =   -15
               TabIndex        =   53
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   20
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "20"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   20
               Left            =   0
               Top             =   915
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   20
               Left            =   -15
               TabIndex        =   52
               Top             =   285
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   19
            Left            =   7185
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   2940
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   19
               Left            =   -15
               TabIndex        =   51
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   19
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "19"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   19
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   19
               Left            =   -15
               TabIndex        =   50
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   18
            Left            =   5985
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   2940
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   18
               Left            =   -15
               TabIndex        =   49
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   18
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "18"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   18
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   18
               Left            =   -15
               TabIndex        =   48
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   17
            Left            =   4785
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   2940
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   17
               Left            =   -15
               TabIndex        =   47
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   17
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "17"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   17
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   17
               Left            =   -15
               TabIndex        =   46
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   16
            Left            =   3585
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   2940
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   16
               Left            =   -15
               TabIndex        =   45
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   16
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "16"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   16
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   16
               Left            =   -15
               TabIndex        =   44
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   15
            Left            =   2385
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   2940
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   15
               Left            =   -15
               TabIndex        =   43
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   15
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "15"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   15
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   15
               Left            =   -15
               TabIndex        =   42
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   14
            Left            =   1185
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   2940
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   14
               Left            =   -15
               TabIndex        =   41
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   14
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "14"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   14
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   14
               Left            =   -15
               TabIndex        =   40
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   13
            Left            =   8385
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   1725
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   13
               Left            =   -15
               TabIndex        =   39
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   13
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "13"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   13
               Left            =   0
               Top             =   915
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   13
               Left            =   -15
               TabIndex        =   38
               Top             =   285
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   12
            Left            =   7185
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   1725
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   12
               Left            =   -15
               TabIndex        =   37
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   12
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "12"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   12
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   12
               Left            =   -15
               TabIndex        =   36
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   11
            Left            =   5985
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   1725
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   11
               Left            =   -15
               TabIndex        =   35
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   11
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "11"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   11
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   11
               Left            =   -15
               TabIndex        =   34
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   10
            Left            =   4785
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   1725
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   10
               Left            =   -15
               TabIndex        =   33
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
               cAlignment      =   1
            End
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   10
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "10"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   10
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   10
               Left            =   -15
               TabIndex        =   32
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   9
            Left            =   3585
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   1725
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   9
               Left            =   -15
               TabIndex        =   31
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   9
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "9"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   9
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   9
               Left            =   -15
               TabIndex        =   30
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   8
            Left            =   2385
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   1725
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   8
               Left            =   -15
               TabIndex        =   29
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   8
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "8"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   8
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   8
               Left            =   -15
               TabIndex        =   28
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   7
            Left            =   1185
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   1725
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   7
               Left            =   -15
               TabIndex        =   27
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   7
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "7"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   7
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   7
               Left            =   -15
               TabIndex        =   26
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   900
            Index           =   37
            Left            =   7890
            ScaleHeight     =   900
            ScaleWidth      =   1680
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   6885
            Width           =   1680
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   300
               Index           =   37
               Left            =   -15
               Top             =   15
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   529
               BackColor       =   0
               ForeColor       =   65535
               Caption         =   "çáÅ@åv"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_KakuKei 
               Height          =   300
               Left            =   0
               Top             =   600
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   529
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin Cis3D_v60.CIS3D B1lb_SurKei 
               Height          =   300
               Left            =   0
               Top             =   300
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   529
               BackColor       =   12632319
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
               cPositionX      =   -70
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   6
            Left            =   8385
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   510
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   6
               Left            =   -15
               TabIndex        =   25
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   6
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "6"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   6
               Left            =   0
               Top             =   915
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   6
               Left            =   -15
               TabIndex        =   24
               Top             =   285
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   5
            Left            =   7185
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   510
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   5
               Left            =   -15
               TabIndex        =   23
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   5
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "5"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   5
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   5
               Left            =   -15
               TabIndex        =   22
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   4
            Left            =   5985
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   510
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   4
               Left            =   -15
               TabIndex        =   21
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   4
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "4"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   4
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   4
               Left            =   -15
               TabIndex        =   20
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   3
            Left            =   4785
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   510
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   3
               Left            =   -15
               TabIndex        =   19
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   3
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "3"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   3
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   3
               Left            =   -15
               TabIndex        =   18
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   2
            Left            =   3585
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   510
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   2
               Left            =   -15
               TabIndex        =   17
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   2
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "2"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   2
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   2
               Left            =   -15
               TabIndex        =   16
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   0
            Left            =   1185
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   510
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   0
               Left            =   -15
               TabIndex        =   13
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   0
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "0"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   0
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16761087
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   0
               Left            =   -15
               TabIndex        =   12
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin VB.PictureBox B_Sur 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'Ç»Çµ
            Height          =   1215
            Index           =   1
            Left            =   2385
            ScaleHeight     =   1215
            ScaleWidth      =   1200
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   510
            Width           =   1200
            Begin CisText_V60.CisText iB1_Sur 
               Height          =   315
               Index           =   1
               Left            =   -15
               TabIndex        =   15
               Top             =   600
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
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
            Begin Cis3D_v60.CIS3D B1lb_DD 
               Height          =   300
               Index           =   1
               Left            =   -15
               Top             =   -15
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               ForeColor       =   16711680
               Caption         =   "1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
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
               cPositionY      =   15
            End
            Begin Cis3D_v60.CIS3D B1lb_Kaku 
               Height          =   285
               Index           =   1
               Left            =   0
               Top             =   915
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               BackColor       =   16777152
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
               cPositionX      =   -70
            End
            Begin CisText_V60.CisText iB1_Nobn 
               Height          =   315
               Index           =   1
               Left            =   -15
               TabIndex        =   14
               Top             =   285
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ÇlÇr ÉSÉVÉbÉN"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXX"
               MaxLength       =   9
               cFbComma        =   0
            End
         End
         Begin Cis3D_v60.CIS3D BLB_Week 
            Height          =   435
            Index           =   0
            Left            =   1185
            Top             =   60
            Width           =   1200
            _ExtentX        =   2117
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
            cPositionY      =   15
         End
         Begin Cis3D_v60.CIS3D BLB_Week 
            Height          =   435
            Index           =   1
            Left            =   2385
            Top             =   60
            Width           =   1200
            _ExtentX        =   2117
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
            cPositionY      =   15
         End
         Begin Cis3D_v60.CIS3D BLB_Week 
            Height          =   435
            Index           =   2
            Left            =   3585
            Top             =   60
            Width           =   1200
            _ExtentX        =   2117
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
            cPositionY      =   15
         End
         Begin Cis3D_v60.CIS3D BLB_Week 
            Height          =   435
            Index           =   3
            Left            =   4785
            Top             =   60
            Width           =   1200
            _ExtentX        =   2117
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
            cPositionY      =   15
         End
         Begin Cis3D_v60.CIS3D BLB_Week 
            Height          =   435
            Index           =   4
            Left            =   5985
            Top             =   60
            Width           =   1200
            _ExtentX        =   2117
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
            cPositionY      =   15
         End
         Begin Cis3D_v60.CIS3D BLB_Week 
            Height          =   435
            Index           =   5
            Left            =   7185
            Top             =   60
            Width           =   1200
            _ExtentX        =   2117
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
            cPositionY      =   15
         End
         Begin Cis3D_v60.CIS3D BLB_Week 
            Height          =   435
            Index           =   6
            Left            =   8385
            Top             =   60
            Width           =   1200
            _ExtentX        =   2117
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
            cPositionY      =   15
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   300
            Left            =   75
            Top             =   1125
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "î[ì¸êî"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   300
            Left            =   75
            Top             =   1425
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   16777152
            ForeColor       =   0
            Caption         =   "ì¸óÕçœêî"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   315
            Left            =   75
            Top             =   510
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "ì˙ït"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   435
            Left            =   75
            Top             =   60
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   767
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
            cBoderWidth     =   2
         End
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   300
            Left            =   75
            Top             =   2040
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "î[ì¸î‘çÜ"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   300
            Left            =   75
            Top             =   2340
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "î[ì¸êî"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D11 
            Height          =   300
            Left            =   75
            Top             =   2640
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   16777152
            ForeColor       =   0
            Caption         =   "ì¸óÕçœêî"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D12 
            Height          =   315
            Left            =   75
            Top             =   1725
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "ì˙ït"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D13 
            Height          =   300
            Left            =   75
            Top             =   3255
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "î[ì¸î‘çÜ"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D14 
            Height          =   300
            Left            =   75
            Top             =   3555
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "î[ì¸êî"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D16 
            Height          =   300
            Left            =   75
            Top             =   3855
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   16777152
            ForeColor       =   0
            Caption         =   "ì¸óÕçœêî"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   315
            Left            =   75
            Top             =   2940
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "ì˙ït"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D19 
            Height          =   300
            Left            =   75
            Top             =   4470
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "î[ì¸î‘çÜ"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D20 
            Height          =   300
            Left            =   75
            Top             =   4770
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "î[ì¸êî"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D21 
            Height          =   300
            Left            =   75
            Top             =   5070
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   16777152
            ForeColor       =   0
            Caption         =   "ì¸óÕçœêî"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D22 
            Height          =   315
            Left            =   75
            Top             =   4155
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "ì˙ït"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D23 
            Height          =   300
            Left            =   75
            Top             =   5685
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "î[ì¸î‘çÜ"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D24 
            Height          =   300
            Left            =   75
            Top             =   5985
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "î[ì¸êî"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D25 
            Height          =   300
            Left            =   75
            Top             =   6285
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   16777152
            ForeColor       =   0
            Caption         =   "ì¸óÕçœêî"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D26 
            Height          =   315
            Left            =   75
            Top             =   5370
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "ì˙ït"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D27 
            Height          =   300
            Left            =   75
            Top             =   6900
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "î[ì¸î‘çÜ"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D28 
            Height          =   300
            Left            =   75
            Top             =   7200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "î[ì¸êî"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D29 
            Height          =   300
            Left            =   75
            Top             =   7500
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BackColor       =   16777152
            ForeColor       =   0
            Caption         =   "ì¸óÕçœêî"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D30 
            Height          =   315
            Left            =   75
            Top             =   6585
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            ForeColor       =   16711680
            Caption         =   "ì˙ït"
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
            cAlingnment     =   7
         End
      End
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
      Caption         =   "Åy éÛíçâÊñ ì¸óÕ(∂⁄›¿ﬁ∞å`éÆ)  Åz"
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
            Size            =   14.25
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14355
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
         Left            =   7410
         Top             =   45
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
   End
   Begin Cis3D_v60.CIS3D H1_Area1 
      Height          =   1035
      Left            =   150
      Top             =   690
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1826
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
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   870
         Left            =   90
         Top             =   90
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   1535
         ForeColor       =   16711680
         Caption         =   "éÛíçãÊï™"
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
         cAlingnment     =   1
         cPositionX      =   90
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H008080FF&
            Caption         =   "éwé¶"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1170
            TabIndex        =   1
            Top             =   120
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ì¡îz"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   1170
            TabIndex        =   5
            Top             =   480
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ó’éû"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   3030
            TabIndex        =   3
            Top             =   120
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "í èÌ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3030
            TabIndex        =   7
            Top             =   480
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "çÜéé"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   2100
            TabIndex        =   2
            Top             =   120
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ÇªÇÃëº"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   3960
            TabIndex        =   8
            Top             =   480
            Width           =   1320
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ÉXÉ|ÉbÉg"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   3960
            TabIndex        =   4
            Top             =   120
            Width           =   1320
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ë„ë÷"
            BeginProperty Font 
               Name            =   "ÇlÇr ÉSÉVÉbÉN"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   2100
            TabIndex        =   6
            Top             =   480
            Width           =   840
         End
      End
   End
End
Attribute VB_Name = "CXJ0090"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       ÉtÉHÅ[ÉÄñº    :   éÛíçâÊñ ì¸óÕÅiÉJÉåÉìÉ_Å[å`éÆÅj
'**       ÉtÉHÅ[ÉÄID    :   CXJ0090
'**       èàóùäTóv      :
'**
'**       çÏ  ê¨  ì˙    :   2010/01/20  By CIS
'**       ïœ  çX  ì˙    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl

    Dim FormAct                 As Boolean      ' Ã´∞—±∏√®Ãﬁ Ã◊∏ﬁ
    Dim ProcHB                  As String * 2   ' ì¸óÕàÊêÿë÷
    Dim SyoriKb                 As String       ' èàóùãÊï™(êVãK•èCê≥•çÌèú ìô)
    Dim SyoriNM                 As String       ' èàóùñºèÃ(êVãK•èCê≥•çÌèú ìô)

    Dim mWeekStart              As Integer
    Dim mSaturdayColor          As Long
    Dim mSundayColor            As Long
'   ì¸â◊óLÇÃîwåiêF
    Dim mColorORG               As Long
    Dim mColorNYT               As Long

'   /// Work ///
    Dim wYmd                    As String
    Dim wIndex                  As Long
    Dim wJsuJNW(31)             As Long         ' éÛíçêîÅ^î≠íçâÊñ ì¸óÕÉèÅ[ÉN
    Dim wJsuJYT(31)             As Long         ' éÛíçêîÅ^î≠íçÉeÅ[ÉuÉã
    Dim wNobn(31)               As String
    
    Dim wEndChk         As Boolean
    Dim wI              As Integer
    Dim wJyuSu          As Long
    Dim wInput          As Boolean      ' ì¸óÕÃ◊∏ﬁ

    Dim w_JKbn          As Byte         '>éÛíçãÊï™
    Dim w_DenKnd(10)    As String       '>ì`ï[éÌóﬁ
    Dim w_PageNo        As Byte         '>î[ïièëï≈
    Dim w_PageGyo       As Byte         '>î[ïièëçs
    
    Dim wProcF          As Integer
    Dim Sav_Index       As Integer      ' çsÉRÉsÅ[å≥çsáÇ
    Dim Sto_Index       As Integer      ' çsÉRÉsÅ[êÊçsáÇ
    Dim wSetColor       As Long         '>DragDropæØƒ∂◊∞
    Dim wLosColor       As Long         '>DragDropèIóπ∂◊∞

    Dim mCHG_Flg        As Byte         'î[ì¸î‘çÜì¸óÕå`ë‘êÿë÷ÉtÉâÉO(0:ëΩïi 1:àÍïi)
    Dim mI              As Integer

'+----------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås    ÇkÇèÇÅÇÑÉCÉxÉìÉg    Åt              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   # ìÒ èd ãN ìÆ ñh é~  #
'   #-------------------#
    If Not Dupli_Start Then End
'   #-------------------#
'   # èâ ä˙ ì‡ óe éÊ ìæ  #
'   #-------------------#
    If Not IniGet Then End
'   #-------------------#
'   #  çÄ ñ⁄ èâ ä˙ âª    #
'   #-------------------#
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
    
    H1Op_JKbn(2).Value = True

'   #------------------#
'   #  èàóùãÊï™èâä˙ï\é¶ #
'   #------------------#
    SyoriKb = ""
    Call SyoriName(SyoriKb)
'   #------------------#
'   #    èâä˙ì¸óÕê›íË   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # âÊñ íÜâõà íuï\é¶   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
    
'   >> ñºèÃÉ}ÉXÉ^ämîF
    Call MeisyoInit
'   èâä˙ï\é¶ê›íË
    IRN.bk = "Option": IRN.rno = 48
    If IRNRead("") Then
        mCHG_Flg = CisFun.Val2(IRN.oyian)
    End If

    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    ÅsÇ`ÇÉÇîÇâÇñÇÅÇîÇÖ ÉCÉxÉìÉgÅt              +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

    If InitProc Then        ' èâä˙èàóùÅiÉfÅ[É^çÌèúÅj
        wEndChk = True
    Else
        wEndChk = False
        Call PB_END_Click
    End If

'   îwåiêFÇÃê›íË
    mColorNYT = B1lb_Kaku(0).BackColor
    mColorORG = B1lb_Kaku(1).BackColor
    B1lb_Kaku(0).BackColor = B1lb_Kaku(1).BackColor

    mWeekStart = 1
    mSaturdayColor = &HFFC0C0
    mSundayColor = &HC0C0FF
'   ójì˙äJénÇãÅÇﬂÇÈ
    gSL_Select = ""
    gSL_Select = gSL_Select & "select êîéö1"
    gSL_Select = gSL_Select & "  from ñºèÃÉ}ÉXÉ^"
    gSL_Select = gSL_Select & " where ãÊï™ñºèÃ = 'ójì˙äJénä«óù'"
    gSL_Select = gSL_Select & "   and íl = '1'"
    If SYMRead(gSL_Select, 1) Then
        mWeekStart = SYM.êîéö1 + 1
    End If
'   ójì˙ÉZÉbÉg
    For gInt = 0 To 6
        BLB_Week(gInt) = WeekdayName(gInt + 1, True, mWeekStart)
        Select Case BLB_Week(gInt)
            Case "ì˙": BLB_Week(gInt).BackColor = mSundayColor
            Case "ìy": BLB_Week(gInt).BackColor = mSaturdayColor
        End Select
    Next gInt
'îNåéèâä˙ï\é¶
    'Call GetNJProc_YM(3)
    'iH2_Ym = gn_HNYM
    iH2_Ym = Format(Now, "YYYYMM")

'ÉVÉXÉeÉÄÉIÉvÉVÉáÉìälìæ
    Call GetSYSOption(33)
    
    FormAct = Not FormAct
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄ    Ås ÇjÇÖÇôÇcÇèÇóÇéÉCÉxÉìÉg Åt               +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+-------------------------------------------------------------+
'+      ÉtÉHÅ[ÉÄÅs ÉtÉ@ÉìÉNÉVÉáÉìÉLÅ[ópÉAÉNÉZÉâÉåÅ[É^ Åt       +
'+-------------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
   'Static W_IVENT As String
   'W_IVENT = "Down"
    Key_Acc = False

    Select Case Key_Code
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    'ÅyêVãKÅz
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    'ÅyèCê≥Åz
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    'ÅyçÌèúÅz
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
'+      ÉtÉHÅ[ÉÄ    Ås QueryUnloadÉCÉxÉìÉg Åt               +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub

'+---------------------+
'+    èI  óπ  ÉL Å[    +
'+---------------------+
Private Sub PB_END_Click()
'<< èIóπèàóù >>
    Call EndProc
'   #### √ﬁ∞¿Õﬁ∞Ωêÿíf ####
    Call CisDB.DBDISConnect
'   #### Ã´∞—µÃﬁºﬁ™∏ƒè¡ãé ####
    Call CisFrm.UnLoadDisp
    End
End Sub
'+----------------------------+
'+    ì¸  óÕ  ÉL Å[(Enter)    +
'+----------------------------+
Private Sub PB_ENT_Click()
   Call ReturnPress
End Sub
'+----------------------------+
'+    éÊ  è¡  ÉL Å[(Escape)   +
'+----------------------------+
Private Sub PB_CAN_Click()
   Call CisFrm.MousePT(11)
   If ProcHB = "H2" Then
      Call Head2Clear
      SyoriKb = ""
      Call SyoriName(SyoriKb)
      ProcHB = "H1"
      Call DispChange(ProcHB)
      GoTo PB_CAN_Ed
   End If
   If ProcHB = "B1" Then
      ProcHB = "H2"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKb)
      iH2_Torcd.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    êV  ãK  ÉL Å[(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    If ProcHB = "H2" Then
        If SyoriKb <> "ADD" Then
            SyoriKb = "ADD"
            Call SyoriName(SyoriKb)
        End If
        iH2_Torcd.SetFocus
    End If
End Sub
'+----------------------------+
'+    èC  ê≥  ÉL Å[(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If ProcHB = "H2" Then
        If SyoriKb <> "MNT" Then
            SyoriKb = "MNT"
            Call SyoriName(SyoriKb)
        End If
        iH2_Torcd.SetFocus
    End If
End Sub
'+----------------------------+
'+    çÌ  èú  ÉL Å[(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If ProcHB = "H2" Then
        If SyoriKb <> "DEL" Then
            SyoriKb = "DEL"
        Else
            SyoriKb = ""
        End If
        Call SyoriName(SyoriKb)
        iH2_Torcd.SetFocus
    End If
End Sub
'+----------------------------+
'+    å„  ëﬁ  ÉL Å[(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    If ProcHB = "H2" Then
        Call CisFrm.MousePT(11)

        Call SqlSelectSet("Back")
        Call DBInput("Back")

       'iH2_Torcd.SetFocus

        Call CisFrm.MousePT(1)
    End If
End Sub
'+----------------------------+
'+    ëO  êi  ÉL Å[(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    If ProcHB = "H2" Then
        Call CisFrm.MousePT(11)

        Call SqlSelectSet("Next")
        Call DBInput("Next")

       'iH2_Torcd.SetFocus

        Call CisFrm.MousePT(1)
    End If
End Sub
'+----------------------------+
'+    åü  çı  ÉL Å[(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH2_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH2_Hinbn" Then Call Look_Hinbn: Exit Sub
End Sub
'+----------------------------+
'+       éÊà¯êÊåüçı           +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 3
    RV_Call = "C"

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
       iH2_Torcd = RTrim(RV_TorcdK)
       Call iH2_Torcd_LostFocus
    End If
End Sub
'+----------------------------+
'+       ïiî‘åüçı             +
'+----------------------------+
Private Sub Look_Hinbn()
    wIndex = CisFun.Val2(Right(PB_Look.Tag, 2))

    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 0
    RVI_SK = 9
    RVI_Torcd = RTrim(iH2_Torcd)
    'RVI_Ukeir = RTrim(iH2_Ukeir)
    RVI_Seban = ""
    RVI_Tehai = ""

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       iH2_Hinbn = RTrim(RV_Hinbn)
    End If
End Sub
'+----------------------------------------+
'+  H1Op_Jkbn ( Click )      éÛíçãÊï™
'+----------------------------------------+
Private Sub H1Op_JKbn_Click(Index As Integer)
    For wIndex = 1 To 8
       H1Op_JKbn(wIndex).BackColor = &HC0C0C0
    Next wIndex
    H1Op_JKbn(Index).BackColor = &H8080FF
End Sub
'+-----------------------------+
'+  iH2_Torcd
'+-----------------------------+
Private Sub iH2_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH2_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H2lb_Tornm = ""
'    If TorNmGet(iH2_Torcd, 8) Then
    If TorNmGet(iH2_Torcd, 3) Then
       H2lb_Tornm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iH2_Hinbn
'+-----------------------------+
Private Sub iH2_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH2_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Sur
'+-----------------------------+
Private Sub iB1_Sur_LostFocus(Index As Integer)
    Call SurKeisan
End Sub

'**************************************
'*                                    *
'*      èàóùãÊï™Ç…ÇÊÇËãÊï™ñºï\é¶      *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
    Dim w_Color     As Long
    w_Color = &H80000008 'èâä˙êFæØƒ(çï)
    Select Case SyoriKb
        Case "ADD"
            SyoriNM = "í«â¡"
            K_Sykbnm.ForeColor = &HFFFF&   '(â©êF)
        Case "MNT"
            SyoriNM = "èCê≥"
            K_Sykbnm.ForeColor = &HFF0000  '(ê¬êF)
        Case "DEL"
            SyoriNM = "çÌèú"
            K_Sykbnm.ForeColor = &HFF&     '(ê‘êF)
            w_Color = &HFF&
        Case "LOK"
            SyoriNM = "è∆âÔ"
            K_Sykbnm.ForeColor = &HFF00FF   '(ÉsÉìÉN)
        Case Else
            SyoriNM = ""
    End Select
    For Each gObj In Me.Controls
        If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
        If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
'       * PF-Key ï\é¶ê›íË *
        If SyoriKb <> "LOC" Then
            If gObj.Name = "PB_ADD" Or _
               gObj.Name = "PB_MNT" Or _
               gObj.Name = "PB_DEL" Then
               gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKb)
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
    Dim H2Mode      As Boolean
    Dim H2Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    
    H1Mode = False: H1Color = gPLostSel
    H2Mode = False: H2Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
' ¿ﬁ–∞çÄñ⁄(âÊñ ÇÃÉ`ÉâÉcÉLñhé~)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1":   H1Mode = True:  H1Color = gPGotSel
       Case "H2":   H2Mode = True:  H2Color = gPGotSel
       Case "B1":   B1Mode = True:  B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode:  H1_Area1.BackColor = H1Color
    H2_Area1.Enabled = H2Mode:  H2_Area1.BackColor = H2Color
    B1_Area1.Enabled = B1Mode:  B1_Area1.BackColor = B1Color

    PB_ADD.Visible = H2Mode
    PB_MNT.Visible = H2Mode
    PB_DEL.Visible = H2Mode
    PB_Look.Visible = H2Mode
    PB_NEXT.Visible = H2Mode
    PB_BACK.Visible = H2Mode
    PB_CAN.Visible = Not H1Mode

' ¿ﬁ–∞çÄñ⁄(ì¸óÕã÷é~)
    Dummy.Enabled = False
'
End Function
'**************************************
'*      Head2 èâä˙âªèàóù              *
'**************************************
Private Sub Head2Clear()

    iH2_Torcd = ""
    H2lb_Tornm = ""
    iH2_Hinbn = ""
'    iH2_Ym = ""
    H2lb_Ukeir = ""
    
End Sub
'**************************************
'*      Head/Body èâä˙âªèàóù          *
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
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        For wIndex = 0 To 36
            B1lb_DD(wIndex).BackColor = &HC0C0C0
            B1lb_Kaku(wIndex).BackColor = mColorORG
        Next wIndex
    End If
End Sub
'****************************
'*      ì¸óÕèàóù(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "H2"
        Call DispChange(ProcHB)
        SyoriKb = "ADD"
        Call SyoriName(SyoriKb)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "H2" Then
        If Not Head2Chk Then GoTo ReturnPress_Ed
' << çÌèúèàóù >>
        If SyoriKb = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            ProcHB = "H2"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKb)
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
        ProcHB = "H2"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKb)
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

    For wIndex = 1 To 8
        If H1Op_JKbn(wIndex).Value Then
            w_JKbn = CisFun.Val2(H1Op_JKbn(wIndex).Tag)
            Exit For
        End If
    Next wIndex

    Head1Chk = True
End Function
'****************************
'*      HEAD2¡™Ø∏èàóù
'****************************
Private Function Head2Chk() As Boolean
    Head2Chk = False

    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With

'   << î[ì¸êÊ >>
    If Trim(iH2_Torcd) = "" Then
        CisFun.MB_MSG(2) = "    î[ì¸êÊÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
        CisFun.MBOX
        iH2_Torcd.SetFocus
        Exit Function
    Else
        If Not SqlSelectTRM(iH2_Torcd) Then
            CisFun.MB_MSG(2) = "    î[ì¸êÊÅ@ñ¢ìoò^ÅIÅiéÊà¯êÊÉ}ÉXÉ^Åj            "
            CisFun.MBOX
            iH2_Torcd.SetFocus
            Exit Function
        End If
    End If
'   ïiî‘
    If Trim(iH2_Hinbn) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ïiî‘Çì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Button = Error
            .MBOX
        End With
        iH2_Hinbn.SetFocus
        Exit Function
    End If
'   îNåé
    If Trim(iH2_Ym) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    îNåéÇì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH2_Ym.SetFocus
        Exit Function
    Else
        If Not iH2_Ym.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    îNåéÇê≥ÇµÇ≠ì¸óÕÇµÇƒâ∫Ç≥Ç¢            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH2_Ym.SetFocus
            Exit Function
        End If
'''        If iH2_Ym < Format(Now(), "yyyymm") Then
'''            CisFun.MB_Lines = 4
'''            CisFun.MB_MSG(2) = "    âﬂãéÇÃì˙ïtÇ≈Ç∑ÅIÅ@ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH            "
'''            CisFun.MB_Button = OK_CAN
'''            If Not CisFun.MBOX Then
'''                iH2_Ym.SetFocus
'''                Exit Function
'''            End If
'''        End If
        If Not CalendarGet(iH2_Ym.cYear, iH2_Ym.cMonth, 0, iH2_Torcd, H2lb_Ukeir) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    " & RTrim(iH2_Ym.cYear) & "îN" & RTrim(iH2_Ym.cMonth) & "åéÇÃÉJÉåÉìÉ_ñ¢ìoò^ÅI            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH2_Ym.SetFocus
            Exit Function
        End If
    End If
    
    If Not SqlSelectHTM Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ïiî‘Å@ñ¢ìoò^ÅIÅiïiî‘éÊà¯êÊÉ}ÉXÉ^Åj            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH2_Hinbn.SetFocus
        Exit Function
    End If

    Call SqlSelectSet("INV")
    If DBInput("INV") Then
        If SyoriKb = "ADD" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ç°âÒéÛíçÉfÅ[É^Å@ìoò^çœÅI            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH2_Torcd.SetFocus
            Exit Function
        End If
    Else
        If SyoriKb <> "ADD" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    ç°âÒéÛíçÉfÅ[É^Å@ñ¢ìoò^ÅI            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH2_Torcd.SetFocus
            Exit Function
        End If
    End If

    Head2Chk = True
End Function
'****************************
'*      Body1¡™Ø∏èàóù
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False

    Call SurKeisan
    
'   ì¸óÕÉ`ÉFÉbÉN
    For wIndex = 0 To 36
        If RTrim(iB1_Nobn(wIndex)) <> "" And RTrim(iB1_Sur(wIndex)) = "" Then
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(2) = "   " & B1lb_DD(wIndex) & "ì˙ êîó Ç™ì¸óÕÇ≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅI"
            CisFun.MB_Button = Error
            CisFun.MBOX
            iB1_Sur(wIndex).SetFocus
            Exit Function
        End If
        If RTrim(iB1_Nobn(wIndex)) <> "" Then
            gSL_Select = ""
            gSL_Select = gSL_Select & "select JYT.î[ì¸î‘çÜ"
            gSL_Select = gSL_Select & " from"
            gSL_Select = gSL_Select & "  éÛíçÉeÅ[ÉuÉã JYT"
            gSL_Select = gSL_Select & " where JYT.î[ì¸ì˙ = '" & iH2_Ym & Format(B1lb_DD(wIndex), "00") & "'"
            gSL_Select = gSL_Select & "   and JYT.î[ì¸êÊ = '" & RTrim(iH2_Torcd) & "'"
            gSL_Select = gSL_Select & "   and JYT.éÛì¸ = '" & RTrim(H2lb_Ukeir) & "'"
            gSL_Select = gSL_Select & "   and JYT.ïiî‘ = '" & RTrim(iH2_Hinbn) & "'"
            gSL_Select = gSL_Select & "   and JYT.î[ì¸î‘çÜ = '" & RTrim(iB1_Nobn(wIndex)) & "'"
            gSL_Select = gSL_Select & "   and JYT.î[ì¸êî = " & CisFun.Val2(iB1_Sur(wIndex))
            If JYTRead(gSL_Select) Then
                CisFun.MB_Lines = 4
                CisFun.MB_MSG(2) = "   " & B1lb_DD(wIndex) & "ì˙ ìoò^çœÇ›Ç≈Ç∑ÅI"
                CisFun.MB_Button = Error
                CisFun.MBOX
                iB1_Sur(wIndex).SetFocus
                Exit Function
            End If
        End If
    Next wIndex

    For wIndex = 0 To 36
        If CisFun.Val2(iB1_Sur(wIndex)) <> 0 Then
        '   â“ìÆì˙ÇÃämîF
            If iB1_Sur(wIndex).Tag = "1" Then
                CisFun.MB_Lines = 4
                CisFun.MB_MSG(2) = "   " & B1lb_DD(wIndex) & "ì˙ îÒâ“ì≠ì˙Ç≈Ç∑ÅIÅ@ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH            "
                CisFun.MB_Button = OK_CAN
                If Not CisFun.MBOX Then
                    iB1_Sur(wIndex).SetFocus
                    Exit Function
                End If
            End If
        End If
    Next wIndex

    Body1Chk = True
End Function
'+---------------------------------------+
'+      ïiî‘éÊà¯êÊÉ}ÉXÉ^çıà¯
'+---------------------------------------+
Public Function SqlSelectHTM(Optional DBNo As Integer = 0) As Boolean
    SqlSelectHTM = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select ht.ïiî‘"
    gSL_Select = gSL_Select & "     , ht.îwî‘çÜ"
    gSL_Select = gSL_Select & "     , ht.é˚óeêî1"
    gSL_Select = gSL_Select & "     , ht.é˚óeêî2"
    gSL_Select = gSL_Select & "     , ht.ïœçXì˙"
    gSL_Select = gSL_Select & "     , ht.ïœçXï÷"
    gSL_Select = gSL_Select & "     , ht.ì`ï[éÌóﬁ"
    gSL_Select = gSL_Select & "     , isnull(tr.ê≥éÆñºèÃ,'') éÊà¯êÊñº"
    gSL_Select = gSL_Select & "     , ht.éÊà¯êÊ"
    gSL_Select = gSL_Select & "     , ht.éÛì¸"
    gSL_Select = gSL_Select & "     , ht.èoâ◊ä«óùïîèê"
    gSL_Select = gSL_Select & "     , isnull(tr1.ê≥éÆñºèÃ,'') èoâ◊ä«óùïîèêñº"
    gSL_Select = gSL_Select & "  from ïiî‘éÊà¯êÊÉ}ÉXÉ^ ht"
    gSL_Select = gSL_Select & "  left outer join éÊà¯êÊÉ}ÉXÉ^ tr "
    gSL_Select = gSL_Select & "  on  tr.éÊà¯êÊCD = '" & RTrim(iH2_Torcd) & "'"
    gSL_Select = gSL_Select & "  and tr.éÊà¯êÊãÊï™ = 0 "
    gSL_Select = gSL_Select & "  left outer join éÊà¯êÊÉ}ÉXÉ^ tr1 "
    gSL_Select = gSL_Select & "  on  tr1.éÊà¯êÊCD = ht.èoâ◊ä«óùïîèê "
    gSL_Select = gSL_Select & "  and tr1.éÊà¯êÊãÊï™ = 2 "
    gSL_Select = gSL_Select & "  and tr1.é–ì‡ãÊï™ = 1 "
'   >> èåèæØƒäJén
    gSL_Select = gSL_Select & " where ht.ïiî‘ = '" & RTrim(iH2_Hinbn) & "'"
'   >> èåèæØƒèIóπ
    gSL_Select = gSL_Select & "   and ht.éÊà¯êÊ = '" & RTrim(iH2_Torcd) & "'"
    gSL_Select = gSL_Select & "   and ht.ïiñ⁄ = 0"

    If HTMRead(gSL_Select, , DBNo) Then
        If CisDB.RecordCount = 1 Then
           SqlSelectHTM = True
        Else
           If CisDB.RecordCount >= 2 Then
              CXJ0065.Show vbModal
              Unload CXJ0065
              Set CXJ0065 = Nothing
              If RV_Rtn Then
                 HTM.ïiî‘ = RV_Hinbn
                 HTM.îwî‘çÜ = RV_Seban
                 HTM.éÛì¸ = RV_Ukeir
                 HTM.èoâ◊ä«óùïîèê = RV_STorcd
                 SqlSelectHTM = True
              End If
           End If
        End If
    End If
    
    H2lb_Ukeir = HTM.éÛì¸

    Call HTMClose(DBNo)
End Function
'**************************************
'*                                    *
'*       ÉfÅ[É^ÉxÅ[ÉXì¸óÕ Åï ï\é¶     *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:

    Call HeadBodyClear("B")

    If Not JNWRead(gSL_Select, 1) Then
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
        Call BodySet("")
   Else
        iH2_Torcd = RTrim(JNW.î[ì¸êÊ)
        iH2_Hinbn = RTrim(JNW.ïiî‘)
        iH2_Ym = CisFun.Mid2(JNW.î[ì¸ì˙, 1, 6)
        H2lb_Ukeir = RTrim(JNW.éÛì¸)
        Call BodySet("ALL")
        DBInput = True
    End If

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'****************************
'*      SQLï∂æØƒèàóù
'****************************
Private Sub SqlSelectSet(InputKey As String)
'
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  JNW.î[ì¸êÊ"
    gSL_Select = gSL_Select & ", JNW.éÛì¸"
    gSL_Select = gSL_Select & ", JNW.ïiî‘"
    gSL_Select = gSL_Select & ", JNW.î[ì¸ì˙"
    gSL_Select = gSL_Select & ", JNW.î[ì¸êî"
    gSL_Select = gSL_Select & ", JNW.î[ì¸î‘çÜ"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  éÛíçì¸óÕÉèÅ[ÉN JNW"
    gSL_Select = gSL_Select & " where JNW.í[ññî‘çÜ = substring(host_name(), 1, 20)"
    gSL_Select = gSL_Select & "   and"
    gSL_Select = gSL_Select & " ( CONVERT(VARCHAR(1),éÛíçãÊï™)"
    gSL_Select = gSL_Select & " + JNW.î[ì¸êÊ + space(7 - len(JNW.î[ì¸êÊ))"
    gSL_Select = gSL_Select & " + JNW.éÛì¸ + space(2 - len(JNW.éÛì¸))"
    gSL_Select = gSL_Select & " + JNW.ïiî‘ + space(25 - len(JNW.ïiî‘))"
    gSL_Select = gSL_Select & " + substring(JNW.î[ì¸ì˙, 1, 6)"
    gSL_Select = gSL_Select & " ) "
    Select Case InputKey
        Case "Back": gSL_Select = gSL_Select & "<"
        Case "Next": gSL_Select = gSL_Select & ">"
        Case Else:   gSL_Select = gSL_Select & "="
    End Select
    gSL_Select = gSL_Select & " ( '"
    gSL_Select = gSL_Select & CStr(w_JKbn)
    gSL_Select = gSL_Select & iH2_Torcd & Space(7 - CisFun.Len2(iH2_Torcd))
    gSL_Select = gSL_Select & H2lb_Ukeir & Space(2 - CisFun.Len2(H2lb_Ukeir))
    gSL_Select = gSL_Select & iH2_Hinbn & Space(25 - CisFun.Len2(iH2_Hinbn))
    gSL_Select = gSL_Select & iH2_Ym & Space(6 - CisFun.Len2(iH2_Ym))
    gSL_Select = gSL_Select & "' )"
    Select Case InputKey
        Case "Back"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  JNW.î[ì¸êÊ desc"
            gSL_Select = gSL_Select & ", JNW.éÛì¸ desc"
            gSL_Select = gSL_Select & ", JNW.ïiî‘ desc"
            gSL_Select = gSL_Select & ", JNW.î[ì¸ì˙ desc"
        Case "Next"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  JNW.î[ì¸êÊ asc"
            gSL_Select = gSL_Select & ", JNW.éÛì¸ asc"
            gSL_Select = gSL_Select & ", JNW.ïiî‘ asc"
            gSL_Select = gSL_Select & ", JNW.î[ì¸ì˙ asc"
    End Select
'
End Sub
'**************************************
'*          BODY ì‡óeæØƒèàóù          *
'**************************************
Private Sub BodySet(SetType As String)
    Dim wYm(2)      As String
'
    With CisFun
        wYm(0) = iH2_Ym
        .Date_RsFormat = fYM
        For gInt = 1 To 2
            .Date_Add iH2_Ym, M, gInt
            wYm(gInt) = .Date_Result
        Next gInt
    End With

    For wIndex = 0 To 31
        wJsuJNW(wIndex) = 0
        wJsuJYT(wIndex) = 0
        wNobn(wIndex) = ""
    Next wIndex

'   éÛíçì¸óÕÉèÅ[ÉN
    Call SqlSelectSet("")
    If JNWRead(gSL_Select) Then
        Do Until Not JNW_RDSTS
            wIndex = CisFun.Mid2(JNW.î[ì¸ì˙, 7, 2)
            wJsuJNW(wIndex) = wJsuJNW(wIndex) + JNW.î[ì¸êî
            wJsuJNW(0) = wJsuJNW(0) + JNW.î[ì¸êî
            wNobn(wIndex) = JNW.î[ì¸î‘çÜ
            Call JNWReadNext
        Loop
        B1lb_SurKei = Format(wJsuJNW(0), "#,###")
    End If
    Call JNWClose
'   éÛíçÉeÅ[ÉuÉã
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  JYT.î[ì¸ì˙"
    gSL_Select = gSL_Select & ", JYT.î[ì¸êî"
    gSL_Select = gSL_Select & ", JYT.î[ì¸î‘çÜ"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  éÛíçÉeÅ[ÉuÉã JYT"
    gSL_Select = gSL_Select & " where JYT.î[ì¸êÊ = '" & RTrim(iH2_Torcd) & "'"
    gSL_Select = gSL_Select & "   and JYT.éÛì¸ = '" & RTrim(H2lb_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and JYT.ïiî‘ = '" & RTrim(iH2_Hinbn) & "'"
    gSL_Select = gSL_Select & "   and substring(JYT.î[ì¸ì˙, 1, 6) = '" & iH2_Ym & "'"
    If JYTRead(gSL_Select) Then
        Do Until Not JYT_RDSTS
            wIndex = CisFun.Mid2(JYT.î[ì¸ì˙, 7, 2)
            wJsuJYT(wIndex) = wJsuJYT(wIndex) + JYT.î[ì¸êî
            wJsuJYT(0) = wJsuJYT(0) + JYT.î[ì¸êî

            Call JYTReadNext
        Loop
        B1lb_KakuKei = Format(wJsuJYT(0), "#,###")
    End If
    Call JYTClose

'   <><><><><> ì˙ïtèÓïÒæØƒ <><><><><>
    Call DaySet

'   ïiî‘éÊà¯êÊÉ}ÉXÉ^
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  htm.îwî‘çÜ"
    gSL_Select = gSL_Select & ", htm.é˚óeêî1"
    gSL_Select = gSL_Select & ", htm.é˚óeäÌ1"
    gSL_Select = gSL_Select & ", htm.ïœçXì˙"
    gSL_Select = gSL_Select & ", htm.ïœçXï÷"
    gSL_Select = gSL_Select & ", htm.é˚óeêî2"
    gSL_Select = gSL_Select & ", htm.é˚óeäÌ2"
    gSL_Select = gSL_Select & ", isnull(hnm.ïiñº, '') ïiñº"
    gSL_Select = gSL_Select & ", isnull(trm.ó™èÃ, '') èoâ◊ä«óùïîèêñº"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  ïiî‘éÊà¯êÊÉ}ÉXÉ^ htm"
    gSL_Select = gSL_Select & "  left outer join ïiî‘É}ÉXÉ^ hnm"
    gSL_Select = gSL_Select & "    on hnm.ïiî‘ = htm.ïiî‘"
    gSL_Select = gSL_Select & "  left outer join éÊà¯êÊÉ}ÉXÉ^ trm"
    gSL_Select = gSL_Select & "    on trm.éÊà¯êÊCD = htm.èoâ◊ä«óùïîèê"
    gSL_Select = gSL_Select & "   and trm.éÊà¯êÊãÊï™ = 2"
    gSL_Select = gSL_Select & "   and trm.é–ì‡ãÊï™   = 1"
    gSL_Select = gSL_Select & " where htm.ïiî‘ = '" & RTrim(iH2_Hinbn) & "'"
    gSL_Select = gSL_Select & "   and htm.éÊà¯êÊ = '" & RTrim(iH2_Torcd) & "'"
    gSL_Select = gSL_Select & "   and htm.éÛì¸ = '" & RTrim(H2lb_Ukeir) & "'"
    If HTMRead(gSL_Select, 1) Then
        With HTM
            B1lb_Hinnm = RTrim(.ïiñº)
            B1lb_SyukaK = RTrim(.èoâ◊ä«óùïîèêñº)
            If Trim(.ïœçXì˙) <> "" Then
                B1lb_ChgDay = Format(.ïœçXì˙, "0000/00/00")
            End If
            B1lb_Syuyo(1) = Format(.é˚óeêî1, "#,###")
            B1lb_Syuyo(2) = Format(.é˚óeêî2, "#,###")
            B1lb_Youki(1) = RTrim(.é˚óeäÌ1)
            B1lb_Youki(2) = RTrim(.é˚óeäÌ2)
        End With
    End If
End Sub
'+-------------------------+
'+  ì˙ïtèÓïÒÉZÉbÉg
'+-------------------------+
Private Sub DaySet()
    Dim wSetDay     As Long
    Dim wSetSta     As Long
    Dim wDate       As Date

    If Trim(iH2_Ym) = "" Then Exit Sub

'   *  ÉJÉåÉìÉ_ïîì˙ït ÉZÉbÉg  *
    If Not CalendarGet(iH2_Ym.cYear, iH2_Ym.cMonth, 0, iH2_Torcd, H2lb_Ukeir) Then
        Call ItemsClearCLM
        CLM.â“ìÆãÊï™ = String(31, "0")
        CLM.ñÈãŒãÊï™ = String(31, "0")
        With CisFun
            .Date_RsFormat = fDD
            .Date_End iH2_Ym
            CLM.ññì˙ = .Val2(.Date_Result)
            CLM.â“ìÆì˙êî = CLM.ññì˙
        End With
    End If

    wDate = DateValue(iH2_Ym.cYear & "/" & iH2_Ym.cMonth & "/01")
    wSetSta = Weekday(wDate, mWeekStart) - 1
    wSetDay = 0
    For wIndex = 0 To 36
    If wIndex < wSetSta Or wSetDay >= CLM.ññì˙ Then
            B_Sur(wIndex).Visible = False
            B1lb_DD(wIndex).Tag = ""
            iB1_Sur(wIndex).Tag = ""
        
        Else
            wSetDay = wSetDay + 1
            B1lb_DD(wIndex) = StrConv(wSetDay, vbWide)
            B1lb_DD(wIndex).Tag = wSetDay
            If Mid(CLM.â“ìÆãÊï™, wSetDay, 1) = "1" And _
               Mid(CLM.ñÈãŒãÊï™, wSetDay, 1) = "1" Then
'               ìyójì˙Åïì˙ójì˙ÇÃîwåiêFæØƒ
                wDate = DateValue(iH2_Ym.cYear & "/" & iH2_Ym.cMonth & "/" & wSetDay)
                Select Case Weekday(wDate)
                    Case vbSunday:   B1lb_DD(wIndex).BackColor = mSundayColor
                    Case vbSaturday: B1lb_DD(wIndex).BackColor = mSaturdayColor
                    Case Else:       B1lb_DD(wIndex).BackColor = mSundayColor
                End Select
                iB1_Sur(wIndex).Tag = "1"
            Else
                B1lb_DD(wIndex).BackColor = &HC0C0C0
                iB1_Sur(wIndex).Tag = ""
            End If
            iB1_Sur(wIndex) = wJsuJNW(wSetDay)
            iB1_Nobn(wIndex) = wNobn(wSetDay)
            B1lb_Kaku(wIndex) = Format(wJsuJYT(wSetDay), "#,###")
            
            B_Sur(wIndex).Visible = True
        End If
    Next wIndex

End Sub
'+-------------------------+
'+  ì˙ïtèÓïÒÉZÉbÉg
'+-------------------------+
Private Sub SurKeisan()
    wJsuJNW(0) = 0
    For wIndex = 0 To 36
        wJsuJNW(0) = wJsuJNW(0) + CisFun.Val2(iB1_Sur(wIndex))
    Next wIndex
    B1lb_SurKei = Format(wJsuJNW(0), "#,###")
End Sub

'**************************************
'*                                    *
'*          ÉfÅ[É^ÉxÅ[ÉXçXêV          *
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
         .MB_Lines = 5
         .MB_MSG(2) = "    çÌèúèàóùÇçsÇ¢Ç‹Ç∑ÅB            "
         .MB_MSG(4) = "     ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH            "
         .MB_Title = "çÌèúämîF"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With

    On Error GoTo DBDelete_Err
    Call CisDB.DBTran(TransBegin)   ' ƒ◊›ªﬁ∏ºÆ›äJén

' ( çÌèúèàóù )
    Call DBDelete_JNW
'   åèêîï\é¶
    iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")

    Call CisDB.DBTran(TransCommit)  ' ƒ◊›ªﬁ∏ºÆ› èIóπ
    On Error GoTo 0

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

'  ( ämîFÉÅÉbÉZÅ[ÉW )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 5
        If SyoriKb = "ADD" Then
            .MB_MSG(2) = "    í«â¡èàóùÇçsÇ¢Ç‹Ç∑ÅB            "
            .MB_Title = "í«â¡ämîF"
        Else
            .MB_MSG(2) = "    èCê≥èàóùÇçsÇ¢Ç‹Ç∑ÅB            "
            .MB_Title = "èCê≥ämîF"
        End If
        .MB_MSG(4) = "     ÇÊÇÎÇµÇ¢Ç≈Ç∑Ç©ÅH            "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With

    On Error GoTo DBPut_Err
    Call CisDB.DBTran(TransBegin)   ' ƒ◊›ªﬁ∏ºÆ›äJén

' ( çÌèúèàóù )
    Call DBDelete_JNW
' ( í«â¡èàóù )
    Call DBSet_JNW
    For wIndex = 0 To 36
        If CisFun.Val2(iB1_Sur(wIndex)) <> 0 Then
            Call DBInsert_JNW(wIndex)
        End If
    Next wIndex
'   åèêîï\é¶
    Select Case SyoriKb
        Case "ADD": iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
        Case "MNT": iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End Select

    Call CisDB.DBTran(TransCommit)  ' ƒ◊›ªﬁ∏ºÆ› èIóπ
    On Error GoTo 0

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+------------------------------------+
'+      çÌèúèàóù - î≠íçâÊñ ì¸óÕÉèÅ[ÉN
'+------------------------------------+
Private Function DBDelete_JNW(Optional DBNo As Integer) As Boolean
    DBDelete_JNW = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "delete éÛíçì¸óÕÉèÅ[ÉN"
    gSL_Select = gSL_Select & " where í[ññî‘çÜ = substring(host_name(), 1, 20)"
    gSL_Select = gSL_Select & "   and éÛíçãÊï™ = '" & w_JKbn & "'"
    gSL_Select = gSL_Select & "   and î[ì¸êÊ = '" & RTrim(iH2_Torcd) & "'"
    gSL_Select = gSL_Select & "   and éÛì¸ = '" & RTrim(H2lb_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and substring(î[ì¸ì˙, 1, 6) = '" & iH2_Ym & "'"
    gSL_Select = gSL_Select & "   and ïiî‘ = '" & RTrim(iH2_Hinbn) & "'"
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With

    DBDelete_JNW = True
End Function
'+------------------------------------+
'+      ÉZÉbÉg - éÛíçì¸óÕÉèÅ[ÉN
'+------------------------------------+
Private Function DBSet_JNW()
    DBSet_JNW = False

    Call ItemsClearJNW
    With JNW
        .í[ññî‘çÜ = CisFun.WSID
        .éÛíçä«óùçsNO = 0
        .éÛíçãÊï™ = w_JKbn
        .éÛíçì˙ = Format(Now(), "yyyymmdd")
        .î[ì¸êÊ = RTrim(iH2_Torcd)
        .éÛì¸ = H2lb_Ukeir
        .ïiî‘ = RTrim(iH2_Hinbn)
        .îwî‘çÜ = HTM.îwî‘çÜ
        .çÏê¨éÌï  = "H"
    End With

    DBSet_JNW = True
End Function
'+------------------------------------+
'+      í«â¡èàóù - î≠íçâÊñ ì¸óÕÉèÅ[ÉN
'+------------------------------------+
Private Function DBInsert_JNW(ID As Long, Optional DBNo As Integer) As Boolean
    DBInsert_JNW = False

    With CisDB
        .SQL = "î‘çÜälìæ"
        .StoadoCount = 3
        .ParaNo = 0: .ParaIO = Returnóp
        .ParaNo = 1: .ParaIO = Inputóp
        .ParaNo = 2: .ParaIO = OutPutóp
        .ParaNo = 3: .ParaIO = Returnóp
    '
        .ParaNo = 1: .ParaValue = 11
    '
        If Not .DBStored Then
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(2) = "    éÛíçä«óùNOÇÃälìæÇ…é∏îsÇµÇ‹ÇµÇΩ            "
            CisFun.MB_Title = "çÃî‘èàóù"
            CisFun.MB_Button = Error
            CisFun.MBOX
            Exit Function
        Else
            .ParaNo = 2: JNW.éÛíçä«óùNO = .ParaValue
        End If
    End With

    With JNW
        .î[ì¸ì˙ = iH2_Ym & Format(B1lb_DD(wIndex).Tag, "00")
        .èoâ◊ó\íËì˙ = iH2_Ym & Format(B1lb_DD(wIndex).Tag, "00")
        If Trim(HTM.ïœçXì˙) = "" Or _
           Trim(HTM.ïœçXì˙) > .î[ì¸ì˙ Then
            .é˚óeêî = HTM.é˚óeêî1
        Else
            .é˚óeêî = HTM.é˚óeêî2
        End If
        .î[ì¸êî = CisFun.Val2(iB1_Sur(ID))
        .î[ì¸î‘çÜ = RTrim(iB1_Nobn(ID))
        If .î[ì¸î‘çÜ = "" Then
            .î[ïièëï≈ = 0
            .î[ïièëçs = 0
        Else
            .î[ïièëï≈ = 1
            .î[ïièëçs = 1
        End If
        If .é˚óeêî = 0 Then
            .ñáêî = 0
            .í[êî = 0
        Else
            .ñáêî = .î[ì¸êî \ .é˚óeêî
            .í[êî = .î[ì¸êî Mod .é˚óeêî
            If .í[êî <> 0 Then .ñáêî = .ñáêî + 1
        End If
        Select Case .éÛíçãÊï™
            Case 2, 4, 5, 6, 7, 8
                .ÉGÉtî≠çsñáêî = .ñáêî
                .ÉGÉtî≠çsí[êî = .í[êî
            Case Else
                .ÉGÉtî≠çsñáêî = 0
                .ÉGÉtî≠çsí[êî = 0
        End Select
        
        .î[ïièëî≠çsãÊï™ = 9
    End With
    Call JNWInsert

    DBInsert_JNW = True
End Function

'*---------------------------------------------------------------------*'
'*      ÉJÉåÉìÉ_É}ÉXÉ^ÇÃçıà¯
'*
'*          à¯êî        : DateY = îN
'*          Å@Å@        : DateM = åé
'*          Å@Å@        : Torkb = éÊà¯ãÊï™ ( Optional )
'*          Å@Å@        : Torcd = éÊà¯êÊÅ@ ( Optional )
'*          Å@Å@        : Ukeir = éÛì¸Å@Å@ ( Optional )
'*
'*          ñﬂíl        : True Or False
'*
'*---------------------------------------------------------------------*'
Public Function CalendarGet(DateY As String _
                          , DateM As String, Optional TorKB As Byte = 1 _
                                           , Optional Torcd As String = "" _
                                           , Optional Ukeir As String = "") As Boolean
    CalendarGet = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select clm.*"
    gSL_Select = gSL_Select & "  from ÉJÉåÉìÉ_É}ÉXÉ^ clm"
    gSL_Select = gSL_Select & " where clm.éÊà¯êÊãÊï™ = " & RTrim(TorKB) & ""
    gSL_Select = gSL_Select & "   and clm.éÊà¯êÊ ='" & RTrim(Torcd) & "'"
    gSL_Select = gSL_Select & "   and clm.îN = '" & DateY & "'"
    gSL_Select = gSL_Select & "   and clm.åé = '" & DateM & "'"
    gSL_Select = gSL_Select & "   and clm.éÛì¸ = '" & RTrim(Ukeir) & "'"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  clm.éÊà¯êÊ desc"
    gSL_Select = gSL_Select & ", clm.éÛì¸ desc"
    If CLMRead(gSL_Select, 1) Then
        CalendarGet = True
        Exit Function
    End If

    gSL_Select = ""
    gSL_Select = gSL_Select & "select clm.*"
    gSL_Select = gSL_Select & "  from ÉJÉåÉìÉ_É}ÉXÉ^ clm"
    gSL_Select = gSL_Select & " where clm.éÊà¯êÊãÊï™ = " & RTrim(TorKB) & ""
    gSL_Select = gSL_Select & "   and clm.éÊà¯êÊ ='" & RTrim(Torcd) & "'"
    gSL_Select = gSL_Select & "   and clm.îN = '" & DateY & "'"
    gSL_Select = gSL_Select & "   and clm.åé = '" & DateM & "'"
    gSL_Select = gSL_Select & "   and clm.éÛì¸ = ''"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  clm.éÊà¯êÊ desc"
    gSL_Select = gSL_Select & ", clm.éÛì¸ desc"
    If CLMRead(gSL_Select, 1) Then
        CalendarGet = True
        Exit Function
    End If
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "select clm.*"
    gSL_Select = gSL_Select & "  from ÉJÉåÉìÉ_É}ÉXÉ^ clm"
    gSL_Select = gSL_Select & " where clm.éÊà¯êÊãÊï™ = " & RTrim(TorKB) & ""
    gSL_Select = gSL_Select & "   and clm.éÊà¯êÊ = '' "
    gSL_Select = gSL_Select & "   and clm.îN = '" & DateY & "'"
    gSL_Select = gSL_Select & "   and clm.åé = '" & DateM & "'"
    gSL_Select = gSL_Select & "   and clm.éÛì¸ = ''"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  clm.éÊà¯êÊ desc"
    gSL_Select = gSL_Select & ", clm.éÛì¸ desc"
    If CLMRead(gSL_Select, 1) Then
        CalendarGet = True
        Exit Function
    End If

    CalendarGet = False
End Function
'==========================================================='
'   èIóπèàóùÅiì¸óÕÉfÅ[É^Ç™Ç†ÇÍÇŒå„ï˚èàóùÇçsÇ§Åj
'==========================================================='
Private Sub EndProc()
'   << ì¸óÕåèêîämîF >>
    gSL_Select = ""
    gSL_Select = gSL_Select & "select count(*) åèêî"
    gSL_Select = gSL_Select & "  from éÛíçì¸óÕÉèÅ[ÉN"
    gSL_Select = gSL_Select & " where í[ññî‘çÜ = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and"
    gSL_Select = gSL_Select & " (     èoâ◊éwé¶èëNO = 0"
    gSL_Select = gSL_Select & "    or î[ì¸î‘çÜ = ''"
    gSL_Select = gSL_Select & "    or ó›êœãÊï™ = 0"
    gSL_Select = gSL_Select & " )"
    If JNWRead(gSL_Select, 1) Then
        If JNW.åèêî > 0 Then
        '   èoâ◊éwé¶NOçÃî‘
            gCHT_DelKbn53 = 0
            CXJ0110.Show vbModal
            Unload CXJ0110
            Set CXJ0110 = Nothing
            DoEvents
        '   èoâ◊éwé¶èëî≠çs
            Call CXJ0120LPrint(0)
        '   é©é–ì`ï[î≠çs
            Call CXJ0140LPrint(0)
        '   é©é–ì`ï[Ãß≤ŸèoóÕ
            Call CXJ0150TFile(0)
        '   ÉmÅ[Ç©ÇÒÇŒÇÒÉGÉtî≠çs
            Call CXJ0130LPrint(0)
        End If
    End If
End Sub
'+---------------------------------------+
'+      ñºèÃÉ}ÉXÉ^èâä˙ê›íË
'+---------------------------------------+
Private Sub MeisyoInit()
'   <><><><><> éÛíçãÊï™ <><><><><>
    Call ItemsClearSYMCreate
    With SYMCreate
        .ãÊï™ñºèÃ = "éÛíçãÊï™"
        .ãÊï™åÖêî = 1
        .íl = "1"
        .ílñºèÃ = "í èÌ"
        .çÏê¨é“ = gTanto
        .êßå‰ãÊï™ = "X"
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .íl = "2":  .ílñºèÃ = "ì¡îz":       Call SYMInitCreate
            .íl = "3":  .ílñºèÃ = "ó’éû":       Call SYMInitCreate
            .íl = "4":  .ílñºèÃ = "ÉXÉ|ÉbÉg":   Call SYMInitCreate
            .íl = "5":  .ílñºèÃ = "éwé¶":       Call SYMInitCreate
            .íl = "6":  .ílñºèÃ = "çÜéé":       Call SYMInitCreate
            .íl = "7":  .ílñºèÃ = "ë„ë÷":       Call SYMInitCreate
            .íl = "8":  .ílñºèÃ = "ÇªÇÃëº":     Call SYMInitCreate
        End With
    End If
'   >> éÛíçãÊï™ÇÃï\é¶
    gSL_Select = ""
    gSL_Select = gSL_Select & "select íl"
    gSL_Select = gSL_Select & "     , ílñºèÃ"
    gSL_Select = gSL_Select & "  from ñºèÃÉ}ÉXÉ^"
    gSL_Select = gSL_Select & " where ãÊï™ñºèÃ = 'éÛíçãÊï™'"
    gSL_Select = gSL_Select & "   and ãÊï™É^ÉCÉv = ''"
    gSL_Select = gSL_Select & " order by íl"
    If SYMRead(gSL_Select, 0) Then
        Do Until Not SYM_RDSTS
            Select Case SYM.íl
                Case "1":   H1Op_JKbn(1).Caption = SYM.ílñºèÃ:  H1Op_JKbn(1).Tag = SYM.íl
                Case "5":   H1Op_JKbn(2).Caption = SYM.ílñºèÃ:  H1Op_JKbn(2).Tag = SYM.íl
                Case "6":   H1Op_JKbn(3).Caption = SYM.ílñºèÃ:  H1Op_JKbn(3).Tag = SYM.íl
                Case "4":   H1Op_JKbn(4).Caption = SYM.ílñºèÃ:  H1Op_JKbn(4).Tag = SYM.íl
                Case "3":   H1Op_JKbn(5).Caption = SYM.ílñºèÃ:  H1Op_JKbn(5).Tag = SYM.íl
                Case "2":   H1Op_JKbn(6).Caption = SYM.ílñºèÃ:  H1Op_JKbn(6).Tag = SYM.íl
                Case "7":   H1Op_JKbn(7).Caption = SYM.ílñºèÃ:  H1Op_JKbn(7).Tag = SYM.íl
                Case "8":   H1Op_JKbn(8).Caption = SYM.ílñºèÃ:  H1Op_JKbn(8).Tag = SYM.íl
            End Select
            Call SYMReadNext
        Loop
    End If
    Call SYMClose
End Sub
'******************************************************************
'*      èâä˙èàóù
'******************************************************************
Private Function InitProc() As Boolean
    InitProc = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select *"
    gSL_Select = gSL_Select & "  from éÛíçì¸óÕÉèÅ[ÉN"
    gSL_Select = gSL_Select & " where í[ññî‘çÜ = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and ó›êœãÊï™ = 0"
    If JNWRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 6
            .MB_MSG(1) = "    ìñí[ññèàóùï™ÇÃéÛíçñ¢ó›êœÉfÅ[É^Ç™ë∂ç›ÇµÇ‹Ç∑ÅB            "
            .MB_MSG(4) = "      ÅwéÛíçâÊñ ì¸óÕÅiÉJÉåÉìÉ_Å[å`éÆÅjÅxèàóùÇèIóπÇµÇ‹Ç∑Ç©ÅH"
            .MB_MSG(5) = "      Ç¢Ç¢Ç¶ÇÃèÍçáÇÕÅAñ¢ó›êœÉfÅ[É^ÇçÌèúÇµë±çsÇµÇ‹Ç∑"
            .MB_Title = "ñ¢ó›êœÉfÅ[É^ÇÃämîF"
            .MB_Button = Yes_No
            If .MBOX Then Exit Function
        End With
    End If

    With CisDB
        On Error GoTo InitProc_Err
        .DBTran (TransBegin) ' ƒ◊›ªﬁ∏ºÆ›äJén
'       << éÛíçì¸óÕÉeÅ[ÉuÉã >>
        .SQL = ""
        .SQL = .SQL & "delete"
        .SQL = .SQL & "  from éÛíçì¸óÕÉèÅ[ÉN"
        .SQL = .SQL & " where í[ññî‘çÜ = substring(host_name(),1,20)"
        .DBExec
        .DBTran (TransCommit) ' ƒ◊›ªﬁ∏ºÆ› èIóπ
        On Error GoTo 0
    End With

    InitProc = True
    Exit Function
InitProc_Err:
    CisFun.ErrorBox
    End
End Function
'+---------------------------------------+
'+      éÊà¯êÊÉ}ÉXÉ^çıà¯ÅiñºèÃçıà¯Åj
'+---------------------------------------+
Public Function SqlSelectTRM(Torcd As String, Optional DBNo As Integer = 0) As Boolean
    SqlSelectTRM = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select ó™èÃ"
    gSL_Select = gSL_Select & "     , çHèÍñº"
    gSL_Select = gSL_Select & "     , ì`ï[éÌóﬁ"
    gSL_Select = gSL_Select & "  from éÊà¯êÊÉ}ÉXÉ^ "
    gSL_Select = gSL_Select & " where éÊà¯êÊCD = '" & RTrim(Torcd) & "'"
    gSL_Select = gSL_Select & "   and éÊà¯êÊãÊï™ = 0 "

    If Not TRMRead(gSL_Select, 1, DBNo) Then Exit Function

    SqlSelectTRM = True
End Function
